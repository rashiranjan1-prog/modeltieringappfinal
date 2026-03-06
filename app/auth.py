"""Session-based auth helpers replacing Flask-Login."""
from functools import wraps
from flask import session, redirect, url_for, abort, g
from .db import get_db


class CurrentUser:
    def __init__(self, row=None):
        if row:
            self.id = row['id']
            self.email = row['email']
            self.role = row['role']
            self.is_authenticated = True
        else:
            self.id = None
            self.email = None
            self.role = None
            self.is_authenticated = False

    def __bool__(self):
        return self.is_authenticated


def get_current_user():
    if 'user_id' not in session:
        return CurrentUser()
    if '_current_user' not in g:
        db = get_db()
        row = db.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
        g._current_user = CurrentUser(row) if row else CurrentUser()
    return g._current_user


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        user = get_current_user()
        if not user.is_authenticated:
            return redirect(url_for('auth.login'))
        return f(*args, **kwargs)
    return decorated


def roles_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            user = get_current_user()
            if not user.is_authenticated:
                return redirect(url_for('auth.login'))
            if user.role not in roles:
                abort(403)
            return f(*args, **kwargs)
        return decorated
    return decorator


def has_role(*roles):
    user = get_current_user()
    return user.is_authenticated and user.role in roles
