from openpyxl import load_workbook
from ..db import get_db
from .tieringlogic import compute_tiering_for_all


def load_excel(filepath):
    wb = load_workbook(filepath, read_only=True)
    db = get_db()
    results = {'models': 0, 'parameters': 0, 'tiers': 0, 'settings': 0, 'computed': 0}

    def headers(row):
        return [str(c).strip().lower().replace(' ', '_') if c else '' for c in row]

    if 'Models' in wb.sheetnames:
        ws = wb['Models']
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            hdrs = headers(rows[0])
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

    if 'Parameters' in wb.sheetnames:
        ws = wb['Parameters']
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            hdrs = headers(rows[0])
            for row in rows[1:]:
                if not any(row):
                    continue
                d = dict(zip(hdrs, row))
                grp = d.get('group')
                if not grp:
                    continue
                db.execute(
                    'INSERT INTO parameters (grp, sub_parameter, criteria, description, weight) VALUES (?,?,?,?,?)',
                    (str(grp), str(d.get('sub_parameter') or ''), str(d.get('criteria') or ''),
                     str(d.get('description') or ''), float(d.get('weight') or 1.0))
                )
                results['parameters'] += 1
        db.commit()

    if 'Tiers' in wb.sheetnames:
        ws = wb['Tiers']
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            hdrs = headers(rows[0])
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

    if 'Settings' in wb.sheetnames:
        ws = wb['Settings']
        for row in ws.iter_rows(values_only=True):
            if not row or not row[0]:
                continue
            key = str(row[0]).strip()
            value = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ''
            db.execute('INSERT OR REPLACE INTO config_kv (key, value) VALUES (?,?)', (key, value))
            results['settings'] += 1
        db.commit()

    results['computed'] = compute_tiering_for_all()
    return results
