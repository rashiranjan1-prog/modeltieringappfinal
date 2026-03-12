#!/usr/bin/env python3
"""CLI management commands for Model Tiering Web App."""
import sys
import os
import getpass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def get_app():
    from app import create_app
    return create_app()


def cmd_init_db():
    app = get_app()
    with app.app_context():
        from app.db import get_db, create_tables, migrate, hash_password
        db = get_db()
        create_tables(db)
        migrate(db)
        print("✓ Tables created.")

        # Default config
        defaults = [
            ('materiality_weight', '0.4'),
            ('criticality_weight', '0.4'),
            ('complexity_weight',  '0.2'),
        ]
        for k, v in defaults:
            db.execute('INSERT OR IGNORE INTO config_kv (key, value) VALUES (?,?)', (k, v))
        db.commit()
        print("✓ Default settings seeded.")

        # Default tiers
        count = db.execute('SELECT COUNT(*) FROM tiers').fetchone()[0]
        if count == 0:
            tiers = [
                ('Tier 1 - High',   67.0, 100.0, 0),
                ('Tier 2 - Medium', 34.0,  66.99, 1),
                ('Tier 3 - Low',     0.0,  33.99, 2),
            ]
            for name, lb, ub, so in tiers:
                db.execute('INSERT INTO tiers (name, lower_bound, upper_bound, sort_order) VALUES (?,?,?,?)',
                           (name, lb, ub, so))
            db.commit()
            print("✓ Default tier ranges seeded.")
        else:
            print("  Tiers already exist, skipping.")

        # Default users
        default_users = [
            ('admin@system.com', 'admin123', 'Admin'),
            ('dev@system.com',   'dev123',   'Developer'),
            ('user@system.com',  'user123',  'User'),
        ]
        for email, pw, role in default_users:
            existing = db.execute('SELECT id FROM users WHERE email=?', (email,)).fetchone()
            if not existing:
                db.execute('INSERT INTO users (email, password_hash, role) VALUES (?,?,?)',
                           (email, hash_password(pw), role))
                print(f"  Created user: {email} ({role})")
            else:
                print(f"  User already exists: {email}")
        db.commit()
        print("✓ Default users ready.")
        print(f"\n✓ Database: {os.path.abspath(app.config['DATABASE'])}")


def cmd_create_user():
    app = get_app()
    with app.app_context():
        from app.db import get_db, hash_password
        email = input("Email: ").strip()
        if not email:
            print("Email required.")
            return
        password = getpass.getpass("Password: ")
        role = input("Role (Admin/Developer/User) [User]: ").strip() or 'User'
        db = get_db()
        existing = db.execute('SELECT id FROM users WHERE email=?', (email,)).fetchone()
        if existing:
            print(f"User {email} already exists.")
            return
        db.execute('INSERT INTO users (email, password_hash, role) VALUES (?,?,?)',
                   (email, hash_password(password), role))
        db.commit()
        print(f"✓ User {email} ({role}) created.")


def cmd_run():
    app = get_app()
    app.run(host='127.0.0.1', port=5000, debug=True)


def cmd_db_path():
    app = get_app()
    print(os.path.abspath(app.config['DATABASE']))


COMMANDS = {
    'init-db': cmd_init_db,
    'create-user': cmd_create_user,
    'run': cmd_run,
    'db-path': cmd_db_path,
}

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python manage.py <command>")
        print("Commands:", ', '.join(COMMANDS.keys()))
        sys.exit(1)
    cmd = sys.argv[1]
    if cmd not in COMMANDS:
        print(f"Unknown command: {cmd}")
        sys.exit(1)
    COMMANDS[cmd]()
