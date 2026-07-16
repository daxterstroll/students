from flask import Blueprint, render_template, request, redirect, url_for, session, flash
from werkzeug.security import check_password_hash
from routes.db import get_db
from routes.utils import log_action
import json

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/')
def index():
    return redirect(url_for('auth.login'))

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('students.student_list'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        conn = get_db()
        user = conn.execute(
            "SELECT id, password_hash, role, is_admin, permissions FROM users WHERE username = ?",
            (username,)
        ).fetchone()
        if user and check_password_hash(user['password_hash'], password):
            group_ids = [row['group_id'] for row in conn.execute(
                "SELECT group_id FROM user_groups WHERE user_id = ?",
                (user['id'],)
            ).fetchall()]
            session['user_id'] = user['id']
            session['role'] = user['role']
            session['group_ids'] = group_ids
            session['username'] = username
            session['is_admin'] = bool(user['is_admin'])
            session['permissions'] = json.loads(user['permissions'] or '[]')
            log_action(
                username,
                "ввійшов у систему",
                group_ids=group_ids,
                details=f"роль: {user['role']}, груп: {len(group_ids)}, IP: {request.remote_addr}"
            )
            conn.close()
            return redirect(url_for('students.student_list'))
        else:
            log_action(
                username or "невідомий",
                "невдала спроба входу",
                details=f"логін: '{username}', IP: {request.remote_addr}"
            )
            flash('Невірний логін або пароль', 'error')
            conn.close()
    return render_template('login.html')

@auth_bp.route('/logout')
def logout():
    username = session.get('username', 'невідомо')
    group_ids = session.get('group_ids', [])
    log_action(
        username,
        "вийшов із системи",
        group_ids=group_ids,
        details=f"IP: {request.remote_addr}"
    )
    session.clear()
    return redirect(url_for('auth.login'))