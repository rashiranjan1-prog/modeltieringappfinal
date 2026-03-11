"""Tiering computation logic using raw sqlite3.

Score formula matches the Excel sheet exactly:
  Internal Score = SUMPRODUCT(param_weights, param_levels) * group_weight
                   summed across all three groups (Materiality, Criticality, Complexity)

  i.e. for each group:
       group_contribution = sum(param_weight_i * level_i) * group_weight

  Final score = Materiality_contribution + Criticality_contribution + Complexity_contribution

  This produces a score in the range 1.0 – 3.0 (matching the Excel output).
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

    # Group-level weights (e.g. Materiality=40%, Criticality=40%, Complexity=20%)
    mat_w  = get_config(db, 'materiality_weight',  0.40)
    crit_w = get_config(db, 'criticality_weight',  0.40)
    comp_w = get_config(db, 'complexity_weight',   0.20)

    scores = db.execute(
        'SELECT ms.*, p.grp, p.weight FROM model_scores ms '
        'JOIN parameters p ON p.id = ms.parameter_id '
        'WHERE ms.model_id=?', (model_id,)
    ).fetchall()

    # If no parameter scores entered, check if Excel already loaded a direct computed_score
    if not scores:
        existing_score = model['computed_score']
        if existing_score and existing_score > 0:
            tiers = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()
            matched = None
            for t in tiers:
                if t['lower_bound'] <= existing_score <= t['upper_bound']:
                    matched = t['name']
                    break
            if matched is None and tiers:
                matched = tiers[-1]['name'] if existing_score > tiers[-1]['upper_bound'] else tiers[0]['name']
            was_overridden = (model['current_tier'] and model['computed_tier'] and
                              model['current_tier'] != model['computed_tier'])
            current = model['current_tier'] if was_overridden else matched
            db.execute('UPDATE models SET computed_tier=?, current_tier=? WHERE id=?',
                       (matched, current, model_id))
            db.commit()
            return db.execute('SELECT * FROM models WHERE id=?', (model_id,)).fetchone()
        # No scores at all — leave as NULL
        db.execute(
            'UPDATE models SET computed_score=0, computed_tier=NULL, current_tier=NULL, last_computed_at=? WHERE id=?',
            (datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'), model_id)
        )
        db.commit()
        return db.execute('SELECT * FROM models WHERE id=?', (model_id,)).fetchone()

    # Delete orphaned score rows for parameters that no longer exist
    db.execute(
        'DELETE FROM model_scores WHERE model_id=? AND parameter_id NOT IN (SELECT id FROM parameters)',
        (model_id,)
    )

    # --- Excel SUMPRODUCT formula replication ---
    # For each group: group_score = sum(param_weight_i * level_i)
    # These param weights are already fractions (e.g. 0.35, 0.35, 0.20, 0.10)
    # that sum to 1.0 within a group, so group_score is already on the 1–3 scale.
    # Final = group_score_mat * mat_w + group_score_crit * crit_w + group_score_comp * comp_w

    group_weighted_sum  = {'Materiality': 0.0, 'Criticality': 0.0, 'Complexity': 0.0}
    group_weight_totals = {'Materiality': 0.0, 'Criticality': 0.0, 'Complexity': 0.0}

    for s in scores:
        grp = s['grp']
        if grp not in group_weighted_sum:
            continue
        param_w = s['weight']   # individual parameter weight within its group (e.g. 0.35)
        level   = s['level']    # 1, 2, or 3
        ws = param_w * level
        group_weighted_sum[grp]  += ws
        group_weight_totals[grp] += param_w
        db.execute('UPDATE model_scores SET weighted_score=? WHERE id=?', (ws, s['id']))

    db.commit()

    def group_score(grp):
        """
        Returns the weighted average score for the group, normalised to the
        1–3 scale even if parameter weights don't sum to exactly 1.0.
        """
        total_w = group_weight_totals[grp]
        if total_w <= 0:
            return 1.0
        raw = group_weighted_sum[grp]
        # If weights already sum to 1 this is just raw.
        # If not (e.g. only some params filled), normalise so the result
        # stays on the 1–3 scale.
        return raw / total_w

    # Normalise group-level weights so they always sum to 1.0
    total_gw = mat_w + crit_w + comp_w
    if total_gw <= 0:
        total_gw = 1.0
    mat_w  /= total_gw
    crit_w /= total_gw
    comp_w /= total_gw

    final = (group_score('Materiality')  * mat_w +
             group_score('Criticality')  * crit_w +
             group_score('Complexity')   * comp_w)

    tiers = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()
    matched = None
    for t in tiers:
        if t['lower_bound'] <= final <= t['upper_bound']:
            matched = t['name']
            break
    if matched is None and tiers:
        matched = tiers[-1]['name'] if final > tiers[-1]['upper_bound'] else tiers[0]['name']

    # Preserve current_tier only if it was deliberately overridden by a user
    was_overridden = (
        model['current_tier'] and
        model['computed_tier'] and
        model['current_tier'] != model['computed_tier']
    )
    current = model['current_tier'] if was_overridden else matched

    db.execute(
        'UPDATE models SET computed_score=?, computed_tier=?, current_tier=?, last_computed_at=? WHERE id=?',
        (round(final, 2), matched, current, datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'), model_id)
    )
    db.commit()
    return db.execute('SELECT * FROM models WHERE id=?', (model_id,)).fetchone()


def compute_tiering_for_all():
    db = get_db()
    models = db.execute('SELECT id FROM models').fetchall()
    for m in models:
        compute_tiering_for_model(m['id'])
    return len(models)
