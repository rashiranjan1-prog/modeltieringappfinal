
# VERSION: 2026-03-11-v5 (correct defaults 0.40/0.40/0.20, SUMPRODUCT formula)
"""Tiering computation — matches Excel SUMPRODUCT formula exactly.

Excel row 22 formula:
  Internal = SUMPRODUCT($E$4:$E$7, H4:H7) * $E$3
           + SUMPRODUCT($E$9:$E$12, H9:H12) * $E$8
           + SUMPRODUCT($E$14:$E$18, H14:H18) * $E$13

  = sum(param_weight_i * level_i)_Materiality  * group_weight_Materiality
  + sum(param_weight_i * level_i)_Criticality  * group_weight_Criticality
  + sum(param_weight_i * level_i)_Complexity   * group_weight_Complexity

  Produces scores 1.0–3.0, identical to the Excel output.
"""
from datetime import datetime
from ..db import get_db


def get_config(db, key, default):
    row = db.execute('SELECT value FROM config_kv WHERE key=?', (key,)).fetchone()
    return float(row['value']) if row else float(default)


def compute_tiering_for_model(model_id):
    db = get_db()
    model = db.execute('SELECT * FROM models WHERE id=?', (model_id,)).fetchone()
    if not model:
        return None

    # Group-level weights — MUST default to 0.40/0.40/0.20 to match Excel
    mat_w  = get_config(db, 'materiality_weight', 0.40)
    crit_w = get_config(db, 'criticality_weight', 0.40)
    comp_w = get_config(db, 'complexity_weight',  0.20)

    scores = db.execute(
        'SELECT ms.*, p.grp, p.weight FROM model_scores ms '
        'JOIN parameters p ON p.id = ms.parameter_id '
        'WHERE ms.model_id=?', (model_id,)
    ).fetchall()

    # No individual scores saved yet
    if not scores:
        existing_score = model['computed_score']
        if existing_score and existing_score > 0:
            tiers = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()
            matched = _match_tier(existing_score, tiers)
            was_overridden = (model['current_tier'] and model['computed_tier'] and
                              model['current_tier'] != model['computed_tier'])
            current = model['current_tier'] if was_overridden else matched
            db.execute('UPDATE models SET computed_tier=?, current_tier=? WHERE id=?',
                       (matched, current, model_id))
            db.commit()
            return db.execute('SELECT * FROM models WHERE id=?', (model_id,)).fetchone()

        db.execute(
            'UPDATE models SET computed_score=0, computed_tier=NULL, current_tier=NULL,'
            ' last_computed_at=? WHERE id=?',
            (datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'), model_id)
        )
        db.commit()
        return db.execute('SELECT * FROM models WHERE id=?', (model_id,)).fetchone()

    # Remove scores for parameters that have been deleted
    db.execute(
        'DELETE FROM model_scores WHERE model_id=? '
        'AND parameter_id NOT IN (SELECT id FROM parameters)', (model_id,)
    )

    # ── SUMPRODUCT formula (matches Excel exactly) ────────────────────────
    # group_score = sum(param_weight_i * level_i)
    # When param weights sum to 1.0 this is a weighted average on the 1–3 scale.
    # We also track weight totals so the formula stays correct if weights are
    # edited and no longer sum to exactly 1.0.
    group_wsum   = {'Materiality': 0.0, 'Criticality': 0.0, 'Complexity': 0.0}
    group_wtotal = {'Materiality': 0.0, 'Criticality': 0.0, 'Complexity': 0.0}

    for s in scores:
        grp = s['grp']
        if grp not in group_wsum:
            continue
        pw    = s['weight']
        level = s['level']
        weighted = pw * level
        group_wsum[grp]   += weighted
        group_wtotal[grp] += pw
        db.execute('UPDATE model_scores SET weighted_score=? WHERE id=?',
                   (weighted, s['id']))

    db.commit()

    def group_score(grp):
        total = group_wtotal[grp]
        if total <= 0:
            return 1.0
        # Equivalent to SUMPRODUCT when weights sum to 1.0;
        # normalises to 1–3 scale if they don't.
        return group_wsum[grp] / total

    # Normalise group-level weights so they always sum to 1.0
    total_gw = mat_w + crit_w + comp_w
    if total_gw <= 0:
        total_gw = 1.0
    mat_w  /= total_gw
    crit_w /= total_gw
    comp_w /= total_gw

    final = (group_score('Materiality') * mat_w +
             group_score('Criticality')  * crit_w +
             group_score('Complexity')   * comp_w)

    tiers   = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()
    matched = _match_tier(final, tiers)

    was_overridden = (model['current_tier'] and model['computed_tier'] and
                      model['current_tier'] != model['computed_tier'])
    current = model['current_tier'] if was_overridden else matched

    db.execute(
        'UPDATE models SET computed_score=?, computed_tier=?, current_tier=?,'
        ' last_computed_at=? WHERE id=?',
        (round(final, 2), matched, current,
         datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'), model_id)
    )
    db.commit()
    return db.execute('SELECT * FROM models WHERE id=?', (model_id,)).fetchone()


def _match_tier(score, tiers):
    for t in tiers:
        if t['lower_bound'] <= score <= t['upper_bound']:
            return t['name']
    if not tiers:
        return None
    return tiers[-1]['name'] if score > tiers[-1]['upper_bound'] else tiers[0]['name']


def compute_tiering_for_all():
    db = get_db()
    models = db.execute('SELECT id FROM models').fetchall()
    for m in models:
        compute_tiering_for_model(m['id'])
    return len(models)
