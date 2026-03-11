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
    weight REAL DEFAULT 1.0,
    level1_label TEXT DEFAULT 'Low',
    level2_label TEXT DEFAULT 'Medium',
    level3_label TEXT DEFAULT 'High'
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
    db.commit()


def migrate(db):
    """Add new columns to existing databases without losing data."""
    existing = [row[1] for row in db.execute("PRAGMA table_info(parameters)").fetchall()]
    for col, default in [('level1_label', 'Low'), ('level2_label', 'Medium'), ('level3_label', 'High')]:
        if col not in existing:
            db.execute(f"ALTER TABLE parameters ADD COLUMN {col} TEXT DEFAULT '{default}'")
    db.commit()
