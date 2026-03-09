import sqlite3
import os

db = sqlite3.connect(os.path.join("instance", "modeltiering.db"))

# Fix weights
db.execute("UPDATE config_kv SET value='0.40' WHERE key='materiality_weight'")
db.execute("UPDATE config_kv SET value='0.40' WHERE key='criticality_weight'")
db.execute("UPDATE config_kv SET value='0.20' WHERE key='complexity_weight'")

# Fix tier ranges — works for Tier1, Tier2, Tier3 names
tiers = db.execute("SELECT id, name FROM tiers ORDER BY id").fetchall()

for tid, tname in tiers:
    clean = tname.lower().replace(" ", "").replace("-", "").replace("_", "")
    if "tier1" in clean or "high" in clean:
        db.execute("UPDATE tiers SET lower_bound=2.2, upper_bound=3.0 WHERE id=?", (tid,))
        print(f"  ✓ '{tname}' → 2.2 to 3.0")
    elif "tier2" in clean or "medium" in clean:
        db.execute("UPDATE tiers SET lower_bound=1.8, upper_bound=2.2 WHERE id=?", (tid,))
        print(f"  ✓ '{tname}' → 1.8 to 2.2")
    elif "tier3" in clean or "low" in clean:
        db.execute("UPDATE tiers SET lower_bound=1.0, upper_bound=1.8 WHERE id=?", (tid,))
        print(f"  ✓ '{tname}' → 1.0 to 1.8")

db.commit()

print("\n--- FINAL: Tier ranges ---")
for row in db.execute("SELECT name, lower_bound, upper_bound FROM tiers ORDER BY upper_bound DESC"):
    print(f"  {row[0]}: {row[1]} to {row[2]}")

print("\n--- FINAL: Weights ---")
for row in db.execute("SELECT key, value FROM config_kv"):
    print(f"  {row[0]} = {row[1]}")

db.close()
print("\nDone!")