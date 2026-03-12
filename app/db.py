
"""Database access layer using stdlib sqlite3."""
import sqlite3
import hashlib
import os
from contextlib import contextmanager
from flask import current_app, g


def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(
            current_app.config['DATABASE'],
            detect_types=sqlite3.PARSE_DECLTYPES
        )
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def close_db(e=None):
    db = g.pop('db', None)
    if db is not None:
        db.close()


def init_app(app):
    app.teardown_appcontext(close_db)


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


def check_password(password, hashed):
    return hashlib.sha256(password.encode()).hexdigest() == hashed


# ─── Schema ──────────────────────────────────────────────────────────────────

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'User',
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS models (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    risk_type TEXT,
    computed_score REAL DEFAULT 0,
    computed_tier TEXT,
    current_tier TEXT,
    last_computed_at TEXT
);

CREATE TABLE IF NOT EXISTS parameters (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    grp TEXT NOT NULL,
    sub_parameter TEXT,
    criteria TEXT,
    description TEXT,
    weight REAL DEFAULT 0.2
);

CREATE TABLE IF NOT EXISTS model_scores (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    model_id INTEGER NOT NULL REFERENCES models(id) ON DELETE CASCADE,
    parameter_id INTEGER NOT NULL REFERENCES parameters(id) ON DELETE CASCADE,
    level INTEGER DEFAULT 1,
    weighted_score REAL DEFAULT 0,
    UNIQUE(model_id, parameter_id)
);

CREATE TABLE IF NOT EXISTS tiers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    lower_bound REAL NOT NULL,
    upper_bound REAL NOT NULL,
    sort_order INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS overrides (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    model_id INTEGER NOT NULL REFERENCES models(id) ON DELETE CASCADE,
    user_id INTEGER NOT NULL REFERENCES users(id),
    old_tier TEXT,
    new_tier TEXT,
    reason TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id),
    name TEXT NOT NULL,
    filters_json TEXT DEFAULT '{}',
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS config_kv (
    key TEXT PRIMARY KEY,
    value TEXT
);
"""


def create_tables(db):
    for stmt in SCHEMA.strip().split(';'):
        stmt = stmt.strip()
        if stmt:
            db.execute(stmt)
    # Fix wrong param weights every startup (e.g. Dependency stuck at 1.0)
    # If any group's weights sum > 1.2, at least one param has default 1.0 → reset all
    EXPECTED_WEIGHTS = {
        'Materiality': [0.35, 0.35, 0.20, 0.10],
        'Criticality': [0.35, 0.35, 0.20, 0.10],
        'Complexity':  [0.20, 0.20, 0.20, 0.20, 0.20],
    }
    groups_params = {}
    for row in db.execute('SELECT id, grp, weight FROM parameters').fetchall():
        groups_params.setdefault(row['grp'], []).append((row['id'], float(row['weight'] or 1.0)))

    # Build case-insensitive lookup for expected weights
    EXPECTED_WEIGHTS_NORM = {k.lower().strip(): v for k, v in EXPECTED_WEIGHTS.items()}

    for grp, params in groups_params.items():
        grp_key = grp.lower().strip()
        total_w = sum(w for _, w in params)
        expected = EXPECTED_WEIGHTS_NORM.get(grp_key)
        # Trigger if: any param has weight=1.0 (default), OR total > 1.2
        has_default = any(abs(w - 1.0) < 0.001 for _, w in params)
        if expected and (has_default or total_w > 1.2):
            for i, (pid, _) in enumerate(params):
                w = expected[i] if i < len(expected) else round(1.0 / len(params), 4)
                db.execute('UPDATE parameters SET weight=? WHERE id=?', (w, pid))

    db.commit()


def migrate(db):
    """Add new columns and seed correct defaults — safe to run on every startup."""
    # Add level label columns if missing
    existing = [row[1] for row in db.execute("PRAGMA table_info(parameters)").fetchall()]
    for col in ('level1_label', 'level2_label', 'level3_label'):
        if col not in existing:
            db.execute(f"ALTER TABLE parameters ADD COLUMN {col} TEXT DEFAULT ''")

    # Seed correct group weights (0.40/0.40/0.20) only if not already in DB
    # This fixes fresh installs that had wrong old defaults baked in
    defaults = [
        ('materiality_weight', '0.4'),
        ('criticality_weight', '0.4'),
        ('complexity_weight',  '0.2'),
    ]
    for key, val in defaults:
        exists = db.execute('SELECT 1 FROM config_kv WHERE key=?', (key,)).fetchone()
        if not exists:
            db.execute('INSERT INTO config_kv (key, value) VALUES (?,?)', (key, val))

    # Fix wrong param weights every startup (e.g. Dependency stuck at 1.0)
    EXPECTED_WEIGHTS = {
        'Materiality': [0.35, 0.35, 0.20, 0.10],
        'Criticality': [0.35, 0.35, 0.20, 0.10],
        'Complexity':  [0.20, 0.20, 0.20, 0.20, 0.20],
    }
    groups_params = {}
    for row in db.execute('SELECT id, grp, weight FROM parameters').fetchall():
        groups_params.setdefault(row['grp'], []).append((row['id'], float(row['weight'] or 1.0)))

    # Build case-insensitive lookup for expected weights
    EXPECTED_WEIGHTS_NORM = {k.lower().strip(): v for k, v in EXPECTED_WEIGHTS.items()}

    for grp, params in groups_params.items():
        grp_key = grp.lower().strip()
        total_w = sum(w for _, w in params)
        expected = EXPECTED_WEIGHTS_NORM.get(grp_key)
        # Trigger if: any param has weight=1.0 (default), OR total > 1.2
        has_default = any(abs(w - 1.0) < 0.001 for _, w in params)
        if expected and (has_default or total_w > 1.2):
            for i, (pid, _) in enumerate(params):
                w = expected[i] if i < len(expected) else round(1.0 / len(params), 4)
                db.execute('UPDATE parameters SET weight=? WHERE id=?', (w, pid))

    db.commit()
