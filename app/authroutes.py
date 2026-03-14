from flask import Blueprint, render_template, redirect, url_for, flash, request, session, abort
from .db import get_db, hash_password, check_password
from .auth import login_required, roles_required, get_current_user

auth_bp = Blueprint('auth', __name__)


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if get_current_user().is_authenticated:
        return redirect(url_for('main.dashboard'))

    if request.method == 'POST':
        email    = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        db   = get_db()
        user = db.execute('SELECT * FROM users WHERE email=?', (email,)).fetchone()
        if user and check_password(password, user['password_hash']):
            session.clear()
            session['user_id'] = user['id']
            flash(f'Welcome back, {user["email"]}!', 'success')
            next_page = request.args.get('next')
            return redirect(next_page or url_for('main.dashboard'))
        flash('Invalid email or password.', 'danger')

    return render_template('login.html')


@auth_bp.route('/register')
@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    # Public registration is disabled.
    # Accounts are created by Admins via Settings > User Management.
    flash('Account registration is disabled. Contact your administrator to get access.', 'warning')
    return redirect(url_for('auth.login'))


@auth_bp.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out.', 'info')
    return redirect(url_for('auth.login'))


# ── Admin-only user management ────────────────────────────────────────────────

@auth_bp.route('/admin/users/create', methods=['POST'])
@login_required
@roles_required('Admin')
def admin_create_user():
    email    = request.form.get('email', '').strip().lower()
    password = request.form.get('password', '').strip()
    role     = request.form.get('role', 'User').strip()

    if not email or not password:
        flash('Email and password are required.', 'danger')
        return redirect(url_for('main.settings'))
    if len(password) < 6:
        flash('Password must be at least 6 characters.', 'danger')
        return redirect(url_for('main.settings'))
    if role not in ('Admin', 'Developer', 'User'):
        role = 'User'

    db = get_db()
    if db.execute('SELECT id FROM users WHERE email=?', (email,)).fetchone():
        flash('Email already registered.', 'danger')
        return redirect(url_for('main.settings'))

    db.execute('INSERT INTO users (email, password_hash, role) VALUES (?,?,?)',
               (email, hash_password(password), role))
    db.commit()
    flash(f'User {email} created with role {role}.', 'success')
    return redirect(url_for('main.settings'))


@auth_bp.route('/admin/users/<int:user_id>/role', methods=['POST'])
@login_required
@roles_required('Admin')
def admin_change_role(user_id):
    new_role = request.form.get('role', '').strip()
    if new_role not in ('Admin', 'Developer', 'User'):
        flash('Invalid role.', 'danger')
        return redirect(url_for('main.settings'))

    me = get_current_user()
    if user_id == me.id:
        flash('You cannot change your own role.', 'danger')
        return redirect(url_for('main.settings'))

    db = get_db()
    db.execute('UPDATE users SET role=? WHERE id=?', (new_role, user_id))
    db.commit()
    user = db.execute('SELECT email FROM users WHERE id=?', (user_id,)).fetchone()
    flash(f'Role updated to {new_role} for {user["email"]}.', 'success')
    return redirect(url_for('main.settings'))


@auth_bp.route('/admin/users/<int:user_id>/delete', methods=['POST'])
@login_required
@roles_required('Admin')
def admin_delete_user(user_id):
    me = get_current_user()
    if user_id == me.id:
        flash('You cannot delete your own account.', 'danger')
        return redirect(url_for('main.settings'))

    db = get_db()
    user = db.execute('SELECT email FROM users WHERE id=?', (user_id,)).fetchone()
    if not user:
        flash('User not found.', 'danger')
        return redirect(url_for('main.settings'))

    db.execute('DELETE FROM users WHERE id=?', (user_id,))
    db.commit()
    flash(f'User {user["email"]} deleted.', 'success')
    return redirect(url_for('main.settings'))
