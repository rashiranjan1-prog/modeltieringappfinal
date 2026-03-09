"""
Excel loader — supports two formats:

Format A (standard):  Sheets named Models, Parameters, Tiers, Settings, Scores
Format B (matrix):    Sheets named Tiering_Method, Parameters, Model_tier, Models_List
                      (the wide pivot format where models are columns and scores are cells)
"""
from openpyxl import load_workbook
from ..db import get_db
from .tieringlogic import compute_tiering_for_all


def _sheet(wb, *names):
    """Return the first sheet whose name matches any of the given names (case-insensitive)."""
    lookup = {s.lower(): s for s in wb.sheetnames}
    for name in names:
        if name.lower() in lookup:
            return wb[lookup[name.lower()]]
    return None


def _headers(row):
    return [str(c).strip().lower().replace(' ', '_').replace('-', '_') if c else '' for c in row]


def load_excel(filepath):
    wb = load_workbook(filepath, read_only=True)
    results = {'models': 0, 'parameters': 0, 'tiers': 0, 'settings': 0, 'scores': 0, 'computed': 0}

    # Detect format
    sheet_names_lower = [s.lower() for s in wb.sheetnames]
    is_matrix_format = 'tiering_method' in sheet_names_lower or 'model_tier' in sheet_names_lower

    if is_matrix_format:
        _load_matrix_format(wb, results)
    else:
        _load_standard_format(wb, results)

    results['computed'] = compute_tiering_for_all()
    return results


# ─── Format B: Matrix (Tiering_Method / Model_tier / Models_List / Parameters) ──

def _load_matrix_format(wb, results):
    db = get_db()

    # Wipe existing data for a clean reimport — prevents stale scores and missing models
    db.execute('DELETE FROM model_scores')
    db.execute('DELETE FROM overrides')
    db.execute('DELETE FROM models')
    db.execute('DELETE FROM parameters')
    db.commit()

    # 1. Parameters — from 'Parameters' sheet
    ws_params = _sheet(wb, 'Parameters')
    param_map = {}  # sub_parameter_name -> param_id
    if ws_params:
        rows = list(ws_params.iter_rows(values_only=True))
        current_group = None
        for row in rows[1:]:
            if not any(row):
                continue
            grp = str(row[0]).strip() if row[0] else None
            sub = str(row[1]).strip() if row[1] else None
            desc = str(row[2]).strip() if row[2] else None
            if grp:
                current_group = grp
            if not sub or not current_group:
                continue
            c = db.execute(
                'INSERT INTO parameters (grp, sub_parameter, criteria, description, weight) VALUES (?,?,?,?,?)',
                (current_group, sub, '', desc or '', 1.0)
            )
            param_map[sub] = c.lastrowid
            results['parameters'] += 1
        db.commit()

    # 2. Models — from 'Models_List' sheet
    ws_models = _sheet(wb, 'Models_List', 'Models')
    model_map = {}  # model_name -> model_id
    if ws_models:
        rows = list(ws_models.iter_rows(values_only=True))
        hdrs = _headers(rows[0])
        name_col = next((i for i, h in enumerate(hdrs) if 'name' in h or 'model' in h), 1)
        risk_col = next((i for i, h in enumerate(hdrs) if 'risk' in h), 2)
        for row in rows[1:]:
            if not any(row):
                continue
            name = str(row[name_col]).strip() if row[name_col] else None
            risk = str(row[risk_col]).strip() if len(row) > risk_col and row[risk_col] else ''
            if not name or name.lower() in ('none', ''):
                continue
            c = db.execute('INSERT INTO models (name, risk_type) VALUES (?,?)', (name, risk))
            model_map[name] = c.lastrowid
            results['models'] += 1
        db.commit()

    # 3. Scores + weights — from 'Tiering_Method' matrix sheet
    ws_matrix = _sheet(wb, 'Tiering_Method', 'Tiering Method')
    if ws_matrix:
        rows = list(ws_matrix.iter_rows(values_only=True))
        if rows:
            header = rows[0]
            # Cols: 0=Parameter, 1=Sub-Parameter, 2=Description, 3=Value Range, 4=Weight, 5+=models
            MODEL_START_COL = 5
            model_names = [str(h).strip() if h else None for h in header[MODEL_START_COL:]]

            current_group = None
            for row in rows[1:]:
                if not any(row):
                    continue
                grp = str(row[0]).strip() if row[0] else None
                sub = str(row[1]).strip() if row[1] else None
                weight = row[4]
                if grp:
                    current_group = grp
                if not sub:
                    continue

                # Update weight in parameters table
                try:
                    w = float(weight) if weight is not None else 1.0
                except (ValueError, TypeError):
                    w = 1.0

                if sub in param_map:
                    db.execute('UPDATE parameters SET weight=? WHERE id=?', (w, param_map[sub]))

                # Load scores for each model
                for col_idx, model_name in enumerate(model_names):
                    if not model_name or model_name not in model_map:
                        continue
                    score_val = row[MODEL_START_COL + col_idx]
                    try:
                        level = int(float(score_val))
                        level = max(1, min(3, level))
                    except (ValueError, TypeError):
                        continue

                    model_id = model_map[model_name]
                    param_id = param_map.get(sub)
                    if not param_id:
                        continue

                    existing = db.execute(
                        'SELECT id FROM model_scores WHERE model_id=? AND parameter_id=?',
                        (model_id, param_id)
                    ).fetchone()
                    if existing:
                        db.execute('UPDATE model_scores SET level=? WHERE id=?', (level, existing['id']))
                    else:
                        db.execute(
                            'INSERT INTO model_scores (model_id, parameter_id, level) VALUES (?,?,?)',
                            (model_id, param_id, level)
                        )
                    results['scores'] += 1
            db.commit()

    # 4. Pre-assigned tiers — from 'Model_tier' sheet (apply as current_tier override)
    ws_tier = _sheet(wb, 'Model_tier', 'Model tier')
    if ws_tier:
        rows = list(ws_tier.iter_rows(values_only=True))
        hdrs = _headers(rows[0])
        name_col = next((i for i, h in enumerate(hdrs) if 'name' in h or 'model' in h), 1)
        tier_col = next((i for i, h in enumerate(hdrs) if 'tier' in h), 3)
        for row in rows[1:]:
            if not any(row):
                continue
            name = str(row[name_col]).strip() if row[name_col] else None
            tier = str(row[tier_col]).strip() if len(row) > tier_col and row[tier_col] else None
            if name and tier and name in model_map:
                db.execute('UPDATE models SET current_tier=? WHERE id=?', (tier, model_map[name]))
        db.commit()


# ─── Format A: Standard (Models / Parameters / Tiers / Settings / Scores) ───

def _load_standard_format(wb, results):
    db = get_db()

    ws = _sheet(wb, 'Models')
    if ws:
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            hdrs = _headers(rows[0])
            for row in rows[1:]:
                if not any(row):
                    continue
                d = dict(zip(hdrs, row))
                name = d.get('name') or d.get('model_name')
                if not name:
                    continue
                existing = db.execute('SELECT id FROM models WHERE name=?', (str(name),)).fetchone()
                if not existing:
                    db.execute('INSERT INTO models (name, risk_type) VALUES (?,?)',
                               (str(name), str(d.get('risk_type') or '')))
                    results['models'] += 1
            db.commit()

    ws = _sheet(wb, 'Parameters')
    if ws:
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            hdrs = _headers(rows[0])
            for row in rows[1:]:
                if not any(row):
                    continue
                d = dict(zip(hdrs, row))
                grp = d.get('group') or d.get('grp')
                if not grp:
                    continue
                db.execute(
                    'INSERT INTO parameters (grp, sub_parameter, criteria, description, weight) VALUES (?,?,?,?,?)',
                    (str(grp), str(d.get('sub_parameter') or ''), str(d.get('criteria') or ''),
                     str(d.get('description') or ''), float(d.get('weight') or 1.0))
                )
                results['parameters'] += 1
            db.commit()

    ws = _sheet(wb, 'Tiers')
    if ws:
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            hdrs = _headers(rows[0])
            db.execute('DELETE FROM tiers')
            for i, row in enumerate(rows[1:]):
                if not any(row):
                    continue
                d = dict(zip(hdrs, row))
                name = d.get('name') or d.get('tier')
                if not name:
                    continue
                db.execute(
                    'INSERT INTO tiers (name, lower_bound, upper_bound, sort_order) VALUES (?,?,?,?)',
                    (str(name), float(d.get('lower_bound') or 0), float(d.get('upper_bound') or 100), i)
                )
                results['tiers'] += 1
            db.commit()

    ws = _sheet(wb, 'Settings')
    if ws:
        for row in ws.iter_rows(values_only=True):
            if not row or not row[0]:
                continue
            key = str(row[0]).strip()
            value = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ''
            db.execute('INSERT OR REPLACE INTO config_kv (key, value) VALUES (?,?)', (key, value))
            results['settings'] += 1
        db.commit()

    ws = _sheet(wb, 'Scores')
    if ws:
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            hdrs = _headers(rows[0])
            for row in rows[1:]:
                if not any(row):
                    continue
                d = dict(zip(hdrs, row))
                model_name = d.get('model_name') or d.get('model')
                param_name = d.get('sub_parameter') or d.get('parameter')
                level = int(d.get('level') or 1)
                if not model_name or not param_name:
                    continue
                model_row = db.execute('SELECT id FROM models WHERE name=?', (str(model_name),)).fetchone()
                param_row = db.execute('SELECT id FROM parameters WHERE sub_parameter=?', (str(param_name),)).fetchone()
                if not model_row or not param_row:
                    continue
                level = max(1, min(3, level))
                db.execute(
                    'INSERT OR REPLACE INTO model_scores (model_id, parameter_id, level) VALUES (?,?,?)',
                    (model_row['id'], param_row['id'], level)
                )
                results['scores'] += 1
            db.commit()
