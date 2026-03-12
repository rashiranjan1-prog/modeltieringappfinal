
import os
import json
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

    # Run migrations and seed correct defaults on every startup
    import sqlite3
    from .config import Config
    _db_path = app.config.get('DATABASE', os.path.join(app.instance_path, 'app.db'))
    _conn = sqlite3.connect(_db_path)
    _conn.row_factory = sqlite3.Row
    database.create_tables(_conn)
    database.migrate(_conn)
    _conn.close()

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
