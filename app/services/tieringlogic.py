"""Tiering computation logic using raw sqlite3."""
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

    mat_w = get_config(db, 'materiality_weight', 0.4)
    crit_w = get_config(db, 'criticality_weight', 0.35)
    comp_w = get_config(db, 'complexity_weight', 0.25)

    scores = db.execute(
        'SELECT ms.*, p.grp, p.weight FROM model_scores ms '
        'JOIN parameters p ON p.id = ms.parameter_id '
        'WHERE ms.model_id=?', (model_id,)
    ).fetchall()

    # If no parameter scores entered, check if Excel already loaded a direct computed_score
    if not scores:
        # If computed_score was already set directly from Excel score row, just assign tier
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

    group_raw = {'Materiality': 0.0, 'Criticality': 0.0, 'Complexity': 0.0}
    group_max = {'Materiality': 0.0, 'Criticality': 0.0, 'Complexity': 0.0}

    for s in scores:
        grp = s['grp']
        if grp in group_raw:
            ws = s['weight'] * s['level']
            group_raw[grp] += ws
            group_max[grp] += s['weight'] * 3
            db.execute('UPDATE model_scores SET weighted_score=? WHERE id=?', (ws, s['id']))

    def normalize(raw, mx):
        # Scale to 1–3: level 1 = 1.0, level 2 = 2.0, level 3 = 3.0
        return (raw / mx) * 3 if mx > 0 else 1.0

    final = (normalize(group_raw['Materiality'], group_max['Materiality']) * mat_w +
             normalize(group_raw['Criticality'], group_max['Criticality']) * crit_w +
             normalize(group_raw['Complexity'], group_max['Complexity']) * comp_w)

    tiers = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()
    matched = None
    for t in tiers:
        if t['lower_bound'] <= final <= t['upper_bound']:
            matched = t['name']
            break
    if matched is None and tiers:
        matched = tiers[-1]['name'] if final > tiers[-1]['upper_bound'] else tiers[0]['name']

    # Preserve current_tier only if it was deliberately overridden
    # (i.e. it differs from the previously computed tier, meaning a user changed it manually).
    # Otherwise always sync current_tier to the new computed value.
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
