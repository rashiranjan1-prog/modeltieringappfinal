import os
import json
import sqlite3
from datetime import datetime
from flask import Flask, session


def create_app(config_object=None):
    app = Flask(__name__, instance_relative_config=True)

    from .config import Config
    app.config.from_object(Config)

    # Ensure dirs exist
    os.makedirs(app.instance_path, exist_ok=True)
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

    # Init DB
    from . import db as database
    database.init_app(app)

    # Run migrations directly (bypasses Flask g, safe on every startup)
    db_path = app.config['DATABASE']
    if os.path.exists(db_path):
        _run_migrations(db_path)

    # Register blueprints
    from .routes import main_bp
    from .authroutes import auth_bp
    app.register_blueprint(main_bp)
    app.register_blueprint(auth_bp)

    # Jinja filters & globals
    @app.template_filter('from_json')
    def from_json_filter(value):
        try:
            return json.loads(value)
        except Exception:
            return {}

    @app.context_processor
    def inject_globals():
        from .auth import get_current_user
        return {
            'now': datetime.utcnow(),
            'current_user': get_current_user(),
        }

    return app


def _run_migrations(db_path):
    """Directly connect and apply schema migrations — no Flask context needed."""
    conn = sqlite3.connect(db_path)
    try:
        existing = [row[1] for row in conn.execute("PRAGMA table_info(parameters)").fetchall()]
        for col, default in [
            ('level1_label', 'Low'),
            ('level2_label', 'Medium'),
            ('level3_label', 'High'),
        ]:
            if col not in existing:
                conn.execute(f"ALTER TABLE parameters ADD COLUMN {col} TEXT DEFAULT '{default}'")
        conn.commit()
    finally:
        conn.close()
