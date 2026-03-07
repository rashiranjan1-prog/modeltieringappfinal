from openpyxl import load_workbook
from ..db import get_db
from .tieringlogic import compute_tiering_for_all


def load_excel(filepath):
    wb = load_workbook(filepath, read_only=True)
    db = get_db()
    results = {'models': 0, 'parameters': 0, 'tiers': 0, 'settings': 0, 'computed': 0, 'errors': []}

    # ── Models from Model_tier sheet ─────────────────────────────────────────
    # Headers: Sl. No. | Model Name | Risk Type | Tier
    if 'Model_tier' in wb.sheetnames:
        ws = wb['Model_tier']
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            # Find header row (look for 'Model Name' in any row)
            header_idx = 0
            for i, row in enumerate(rows):
                if row and any(str(c).strip().lower() in ['model name', 'model_name'] for c in row if c):
                    header_idx = i
                    break

            raw_headers = rows[header_idx]
            headers = [str(h).strip().lower().replace(' ', '_') if h else '' for h in raw_headers]

            # Find column indices
            def find_col(keywords):
                for kw in keywords:
                    for i, h in enumerate(headers):
                        if kw in h:
                            return i
                return None

            name_col = find_col(['model_name', 'model name', 'name'])
            risk_col = find_col(['risk_type', 'risk type', 'risk'])
            tier_col = find_col(['tier'])

            for row in rows[header_idx + 1:]:
                if not row or not any(row):
                    continue
                name = row[name_col] if name_col is not None and name_col < len(row) else None
                if not name or str(name).strip() == '' or str(name).strip().lower() in ['model name', 'nan']:
                    continue
                name = str(name).strip()
                risk = str(row[risk_col]).strip() if risk_col is not None and risk_col < len(row) and row[risk_col] else ''
                tier = str(row[tier_col]).strip() if tier_col is not None and tier_col < len(row) and row[tier_col] else ''

                existing = db.execute('SELECT id FROM models WHERE name=?', (name,)).fetchone()
                if not existing:
                    db.execute(
                        'INSERT INTO models (name, risk_type, current_tier, computed_tier) VALUES (?,?,?,?)',
                        (name, risk, tier, tier)
                    )
                    results['models'] += 1
                else:
                    # Update risk type and tier if already exists
                    db.execute(
                        'UPDATE models SET risk_type=?, current_tier=?, computed_tier=? WHERE name=?',
                        (risk, tier, tier, name)
                    )
        db.commit()

    # ── Parameters from Tiering_Method sheet ─────────────────────────────────
    # Headers: Parameter | Sub-Parameter | Description | Value Range | Weight | [model columns...]
    # Groups are determined by the Parameter column value
    if 'Tiering_Method' in wb.sheetnames:
        ws = wb['Tiering_Method']
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            # Find header row
            header_idx = 0
            for i, row in enumerate(rows):
                if row and any(str(c).strip().lower() in ['parameter', 'sub-parameter', 'sub- parameter'] for c in row if c):
                    header_idx = i
                    break

            raw_headers = rows[header_idx]
            headers = [str(h).strip().lower() if h else '' for h in raw_headers]

            def find_col(keywords):
                for kw in keywords:
                    for i, h in enumerate(headers):
                        if kw in h.lower():
                            return i
                return None

            param_col   = find_col(['parameter'])
            sub_col     = find_col(['sub- parameter', 'sub-parameter', 'sub_parameter'])
            desc_col    = find_col(['description'])
            weight_col  = find_col(['weight'])

            # Clear existing parameters before reloading
            db.execute('DELETE FROM parameters')
            db.commit()

            current_group = ''
            for row in rows[header_idx + 1:]:
                if not row or not any(row):
                    continue

                # Get parameter (group)
                param_val = str(row[param_col]).strip() if param_col is not None and param_col < len(row) and row[param_col] else ''
                sub_val   = str(row[sub_col]).strip()   if sub_col   is not None and sub_col   < len(row) and row[sub_col]   else ''
                desc_val  = str(row[desc_col]).strip()  if desc_col  is not None and desc_col  < len(row) and row[desc_col]  else ''

                # Weight
                weight = 1.0
                if weight_col is not None and weight_col < len(row) and row[weight_col]:
                    try:
                        weight = float(row[weight_col])
                    except (ValueError, TypeError):
                        weight = 1.0

                # Track current group
                if param_val and param_val.lower() not in ['parameter', 'nan', '']:
                    current_group = param_val

                # Only insert rows that have a sub-parameter
                if not sub_val or sub_val.lower() in ['sub- parameter', 'sub-parameter', 'nan', '']:
                    continue

                db.execute(
                    'INSERT INTO parameters (grp, sub_parameter, criteria, description, weight) VALUES (?,?,?,?,?)',
                    (current_group, sub_val, '', desc_val, weight)
                )
                results['parameters'] += 1

        db.commit()

    # ── Tiers: derive from unique tier values in Model_tier sheet ────────────
    # Since there's no dedicated Tiers sheet, build tiers from Model_tier data
    existing_tiers = db.execute('SELECT COUNT(*) FROM tiers').fetchone()[0]
    if existing_tiers == 0:
        # Seed default tiers
        default_tiers = [
            ('Tier 1', 67.0, 100.0, 0),
            ('Tier 2', 34.0, 66.99, 1),
            ('Tier 3', 0.0,  33.99, 2),
        ]
        for name, lb, ub, so in default_tiers:
            db.execute('INSERT INTO tiers (name, lower_bound, upper_bound, sort_order) VALUES (?,?,?,?)',
                       (name, lb, ub, so))
        db.commit()
        results['tiers'] = 3

    # Also try to get unique tier names from Model_tier and create matching tiers
    if 'Model_tier' in wb.sheetnames:
        ws = wb['Model_tier']
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            header_idx = 0
            for i, row in enumerate(rows):
                if row and any(str(c).strip().lower() in ['model name', 'model_name'] for c in row if c):
                    header_idx = i
                    break
            raw_headers = rows[header_idx]
            headers = [str(h).strip().lower().replace(' ', '_') if h else '' for h in raw_headers]
            tier_col = next((i for i, h in enumerate(headers) if 'tier' in h), None)

            if tier_col is not None:
                unique_tiers = set()
                for row in rows[header_idx + 1:]:
                    if row and tier_col < len(row) and row[tier_col]:
                        t = str(row[tier_col]).strip()
                        if t and t.lower() not in ['tier', 'nan', '']:
                            unique_tiers.add(t)

                # Replace tiers with actual values from the sheet
                if unique_tiers:
                    db.execute('DELETE FROM tiers')
                    for i, tname in enumerate(sorted(unique_tiers)):
                        db.execute(
                            'INSERT INTO tiers (name, lower_bound, upper_bound, sort_order) VALUES (?,?,?,?)',
                            (tname, 0.0, 100.0, i)
                        )
                    db.commit()
                    results['tiers'] = len(unique_tiers)

    # ── Default settings if not set ──────────────────────────────────────────
    for key, val in [('materiality_weight', '0.4'), ('criticality_weight', '0.35'), ('complexity_weight', '0.25')]:
        existing = db.execute('SELECT value FROM config_kv WHERE key=?', (key,)).fetchone()
        if not existing:
            db.execute('INSERT INTO config_kv (key, value) VALUES (?,?)', (key, val))
            results['settings'] += 1
    db.commit()

    # ── Compute tiering for all models ───────────────────────────────────────
    results['computed'] = compute_tiering_for_all()
    return results
