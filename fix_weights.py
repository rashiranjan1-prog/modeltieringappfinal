import sqlite3
import glob
import os

# Find the database
paths = (
    glob.glob('**/*.db', recursive=True) +
    glob.glob('instance/*.db') +
    glob.glob('instance/**/*.db', recursive=True)
)

if not paths:
    print("ERROR: No .db file found! Make sure you run this from your project folder.")
    input("Press Enter to exit...")
    exit(1)

EXPECTED = {
    'materiality': [0.35, 0.35, 0.20, 0.10],
    'criticality': [0.35, 0.35, 0.20, 0.10],
    'complexity':  [0.20, 0.20, 0.20, 0.20, 0.20],
}

for p in paths:
    print(f"\nDatabase: {p}")
    conn = sqlite3.connect(p)
    conn.row_factory = sqlite3.Row

    rows = conn.execute(
        'SELECT id, sub_parameter, grp, weight FROM parameters ORDER BY grp, id'
    ).fetchall()

    if not rows:
        print("  No parameters found.")
        conn.close()
        continue

    print(f"  {'GRP':<15} {'SUB_PARAMETER':<30} {'CURRENT':>8} {'NEW':>8} {'STATUS'}")
    print(f"  {'-'*75}")

    fixed = 0
    groups = {}
    for row in rows:
        grp = str(row['grp'] or '').strip().lower()
        groups.setdefault(grp, []).append(row)

    for grp_key, params in groups.items():
        expected = EXPECTED.get(grp_key)
        for i, p_row in enumerate(params):
            cur_w = float(p_row['weight'] or 0)
            if expected and i < len(expected):
                new_w = expected[i]
            else:
                new_w = round(1.0 / len(params), 4)

            status = ''
            if abs(cur_w - new_w) > 0.001:
                conn.execute(
                    'UPDATE parameters SET weight=? WHERE id=?',
                    (new_w, p_row['id'])
                )
                fixed += 1
                status = '← FIXED'

            print(f"  {p_row['grp']:<15} {p_row['sub_parameter']:<30} {cur_w:>8.3f} {new_w:>8.3f}  {status}")

    conn.commit()
    conn.close()
    print(f"\n  ✓ Fixed {fixed} weight(s).")

print("\nDone! Refresh your browser — Dependency should now show 0.2")
input("Press Enter to exit...")
