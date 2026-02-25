from flask import Blueprint, render_template, request, redirect, url_for, session, flash
from werkzeug.security import check_password_hash
from db import get_db
from utils import log_action

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/')
def index():
    """Главная страница приложения."""
    return redirect(url_for('auth.login'))

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    """Обработка авторизации пользователя."""
    if 'user_id' in session:
        return redirect(url_for('students.student_list'))

    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        conn = get_db()
        user = conn.execute("SELECT id, password_hash, role FROM users WHERE username = ?", (username,)).fetchone()

        if user and check_password_hash(user['password_hash'], password):
            group_ids = [row['group_id'] for row in conn.execute("SELECT group_id FROM user_groups WHERE user_id = ?", (user['id'],)).fetchall()]
            session['user_id'] = user['id']
            session['role'] = user['role']
            session['group_ids'] = group_ids  # Store list of group IDs
            session['username'] = username
            log_action(username, "ввійшов у систему", group_ids=group_ids)
            conn.close()
            return redirect(url_for('students.student_list'))
        else:
            flash('Невірний логін або пароль', 'error')
            conn.close()
    return render_template('login.html')


@auth_bp.route('/logout')
def logout():
    """Выход пользователя из системы."""
    username = session.get('username', 'невідомо')
    group_ids = session.get('group_ids', [])  # Get list of group IDs
    session.clear()
    log_action(username, "вийшов із системи", group_ids=group_ids)
    return redirect(url_for('auth.login'))