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
    wb = load_workbook(filepath, data_only=True)
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
    # Reset autoincrement counters so IDs restart from 1
    db.execute("DELETE FROM sqlite_sequence WHERE name IN ('models', 'parameters', 'model_scores', 'overrides')")
    db.commit()

    # 1. Parameters — from 'Parameters' sheet
    # Columns: A=Group, B=Sub-Parameter, C=Description, D=Low(1), E=Medium(2), F=High(3)
    ws_params = _sheet(wb, 'Parameters')
    param_map = {}  # sub_parameter_name -> param_id
    if ws_params:
        rows = list(ws_params.iter_rows(values_only=True))
        current_group = None
        for row in rows[1:]:
            if not any(row):
                continue
            grp  = str(row[0]).strip() if len(row) > 0 and row[0] else None
            sub  = str(row[1]).strip() if len(row) > 1 and row[1] else None
            desc = str(row[2]).strip() if len(row) > 2 and row[2] else None
            l1   = str(row[3]).strip() if len(row) > 3 and row[3] else 'Low'
            l2   = str(row[4]).strip() if len(row) > 4 and row[4] else 'Medium'
            l3   = str(row[5]).strip() if len(row) > 5 and row[5] else 'High'
            if grp:
                current_group = grp
            if not sub or not current_group:
                continue
            c = db.execute(
                'INSERT INTO parameters (grp, sub_parameter, criteria, description, weight, level1_label, level2_label, level3_label) VALUES (?,?,?,?,?,?,?,?)',
                (current_group, sub, '', desc or '', 1.0, l1, l2, l3)
            )
            param_map[sub] = c.lastrowid
            results['parameters'] += 1
        db.commit()

    # 2. Models — from 'Model_tier' sheet (has name, risk_type AND pre-assigned tier)
    ws_models = _sheet(wb, 'Model_tier', 'Model_List', 'Models_List', 'Models')
    model_map = {}  # model_name -> model_id
    if ws_models:
        rows = list(ws_models.iter_rows(values_only=True))
        hdrs = _headers(rows[0])
        name_col = next((i for i, h in enumerate(hdrs) if 'name' in h or 'model' in h), 1)
        risk_col = next((i for i, h in enumerate(hdrs) if 'risk' in h), 2)
        tier_col = next((i for i, h in enumerate(hdrs) if 'tier' in h), None)
        for row in rows[1:]:
            if not any(row):
                continue
            name = str(row[name_col]).strip() if row[name_col] else None
            risk = str(row[risk_col]).strip() if len(row) > risk_col and row[risk_col] else ''
            tier = str(row[tier_col]).strip() if tier_col and len(row) > tier_col and row[tier_col] else None
            if not name or name.lower() in ('none', ''):
                continue
            c = db.execute('INSERT INTO models (name, risk_type, current_tier) VALUES (?,?,?)', (name, risk, tier))
            model_map[name] = c.lastrowid
            results['models'] += 1
        db.commit()

    # 3. Scores — from 'Tiering_Method' matrix sheet
    ws_matrix = _sheet(wb, 'Tiering_Method', 'Tiering Method')
    if ws_matrix:
        rows = list(ws_matrix.iter_rows(values_only=True))
        if rows:
            header = rows[0]
            # Auto-detect MODEL_START_COL: find first column in header that matches a known model name
            # Model names are already in model_map — first header col matching a model name = start
            MODEL_START_COL = 5  # default fallback
            if model_map:
                for col_idx, cell in enumerate(header):
                    if cell and str(cell).strip() in model_map:
                        MODEL_START_COL = col_idx
                        break
            model_names = [str(h).strip() if h else None for h in header[MODEL_START_COL:]]

            # Find the Internal score row (row 22 in real Excel).
            # Key insight: the Internal row comes AFTER all parameter rows.
            # Parameter rows have an integer score (1/2/3) in model cols.
            # Group subtotal rows have decimals but also have a % weight in col E/F.
            # Internal row: model cols have decimals, AND pre-model cols are all empty/None.
            model_count = len([m for m in model_names if m])

            # First pass: find the last parameter row index (rows with integer 1/2/3 scores)
            last_param_row_idx = 0
            for idx, row in enumerate(rows[1:], 1):
                if not any(row):
                    continue
                # A param row has a sub-param name AND integer scores in model cols
                sub = str(row[1] or '').strip()
                if sub and sub not in ('', 'None'):
                    model_vals = [row[MODEL_START_COL + i] for i in range(min(5, model_count))
                                  if MODEL_START_COL + i < len(row)]
                    int_scores = [v for v in model_vals if v in (1, 2, 3)]
                    if len(int_scores) >= 2:
                        last_param_row_idx = idx

            # Second pass: find score row AFTER last param row
            # It must have decimal values in model cols AND no integer % weight in pre-model cols
            score_row = None
            for idx, row in enumerate(rows[1:], 1):
                if idx <= last_param_row_idx:
                    continue  # skip all rows before/at last param row
                if not any(row):
                    continue
                # Skip rows where pre-model cols contain % weights (group subtotal rows)
                pre_vals = [row[i] for i in range(MODEL_START_COL) if i < len(row)]
                has_pct = any(isinstance(v, str) and '%' in str(v) for v in pre_vals)
                if has_pct:
                    continue
                # Skip tier assignment rows (contain text like 'Tier1', 'Tier2')
                model_vals = [row[MODEL_START_COL + i] for i in range(min(model_count, len(row) - MODEL_START_COL))
                              if MODEL_START_COL + i < len(row)]
                has_tier_text = any(isinstance(v, str) and 'tier' in str(v).lower() for v in model_vals)
                if has_tier_text:
                    continue
                # Count decimal values in model cols
                decimal_vals = []
                for i in range(model_count):
                    actual_col = MODEL_START_COL + i
                    if actual_col >= len(row):
                        continue
                    v = row[actual_col]
                    try:
                        fv = float(v)
                        if 1.0 <= fv <= 3.0 and fv != int(fv):
                            decimal_vals.append(fv)
                    except (TypeError, ValueError):
                        pass
                if len(decimal_vals) >= max(2, model_count * 0.3):
                    score_row = row
                    break  # take the first qualifying row after param rows

            # Log which strategy found the score row
            if score_row is not None:
                results['score_row_found'] = True
            # Apply direct scores to models if score row found
            if score_row is not None:
                for col_idx, model_name in enumerate(model_names):
                    if not model_name or model_name not in model_map:
                        continue
                    actual_col = MODEL_START_COL + col_idx
                    val = score_row[actual_col] if actual_col < len(score_row) else None
                    if val is None:
                        continue
                    try:
                        score = round(float(val), 2)
                        model_id = model_map[model_name]
                        db.execute('UPDATE models SET computed_score=?, last_computed_at=datetime("now") WHERE id=?',
                                   (score, model_id))
                        results['scores'] += 1
                    except (ValueError, TypeError):
                        continue
                db.commit()

            # Pass: read parameter rows — update weights AND save individual scores into model_scores
            current_group = None
            for row in rows[1:]:
                grp = str(row[0]).strip() if row[0] else None
                sub = str(row[1]).strip() if row[1] else None
                weight = row[4]
                if grp:
                    current_group = grp
                if not sub:
                    continue

                # Update parameter weight
                if weight is not None:
                    try:
                        w = float(weight)
                        if sub in param_map:
                            db.execute('UPDATE parameters SET weight=? WHERE id=?', (w, param_map[sub]))
                    except (ValueError, TypeError):
                        pass

                # Save individual model scores (integer 1/2/3) into model_scores
                if sub not in param_map:
                    continue
                param_id = param_map[sub]
                for col_idx, model_name in enumerate(model_names):
                    if not model_name or model_name not in model_map:
                        continue
                    actual_col = MODEL_START_COL + col_idx
                    if actual_col >= len(row):
                        continue
                    cell_val = row[actual_col]
                    try:
                        level = int(cell_val)
                        if level not in (1, 2, 3):
                            continue
                        model_id = model_map[model_name]
                        db.execute(
                            'INSERT OR REPLACE INTO model_scores (model_id, parameter_id, level) VALUES (?,?,?)',
                            (model_id, param_id, level)
                        )
                    except (TypeError, ValueError):
                        continue
            db.commit()

    # Pre-assigned tiers are already loaded in step 2 from Model_tier sheet


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
                    'INSERT INTO parameters (grp, sub_parameter, criteria, description, weight, level1_label, level2_label, level3_label) VALUES (?,?,?,?,?,?,?,?)',
                    (str(grp), str(d.get('sub_parameter') or ''), str(d.get('criteria') or ''),
                     str(d.get('description') or ''), float(d.get('weight') or 1.0),
                     str(d.get('low') or d.get('level1_label') or 'Low'),
                     str(d.get('medium') or d.get('level2_label') or 'Medium'),
                     str(d.get('high') or d.get('level3_label') or 'High'))
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
