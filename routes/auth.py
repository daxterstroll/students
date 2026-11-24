"""
routes/auth.py
==============
Логін / логаут користувачів. Всі успішні та невдалі спроби входу, а
також виходи із системи записуються через log_action() у app.log.
"""

from flask import Blueprint, render_template, request, redirect, url_for, session, flash
from werkzeug.security import check_password_hash
from routes.db import get_db
from routes.utils import log_action, logger
from routes.helpers import current_username
from routes.security import check_rate_limit, client_ip
from routes.user_scope import get_allowed_group_ids
import json

auth_bp = Blueprint('auth', __name__)


@auth_bp.route('/')
def index():
    """Коренева сторінка сайту - просто перенаправляє на форму логіну."""
    return redirect(url_for('auth.login'))


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    """
    Форма входу в систему.
    """
    if 'user_id' in session:
        return redirect(url_for('students.student_list'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        ip = client_ip()

        ok_ip, retry_ip = check_rate_limit('login_ip', ip, limit=20, window_seconds=300)
        ok_acct, retry_acct = check_rate_limit(
            'login_acct', f"{ip}:{username.casefold()}", limit=8, window_seconds=300
        )
        if not ok_ip or not ok_acct:
            wait_sec = max(retry_ip, retry_acct)
            log_action(
                username or "невідомий",
                "блокування спроб входу (rate limit)",
                level='WARNING',
                category='security',
                action_type='rate_limit',
                section='Авторизація',
                ip_address=ip,
                details=f"логін: '{username}', IP: {ip}, очікування: {wait_sec}с",
                meta={'Введений логін': username or '—', 'Очікування (сек)': wait_sec},
            )
            flash(
                f'Забагато невдалих спроб входу. Спробуйте знову через {wait_sec} сек.',
                'error',
            )
            return render_template('login.html'), 429

        conn = None
        try:
            conn = get_db()
            uname_fold = username.casefold()
            all_users = conn.execute(
                "SELECT id, username, password_hash, role, is_admin, permissions FROM users"
            ).fetchall()

            exact_matches = [
                u for u in all_users
                if (u['username'] or '').strip() == username
            ]
            ci_matches = [
                u for u in all_users
                if (u['username'] or '').strip() != username
                and (u['username'] or '').strip().casefold() == uname_fold
            ]

            user = None
            for cand in exact_matches + ci_matches:
                if check_password_hash(cand['password_hash'], password):
                    user = cand
                    break

            if user:
                canonical_username = user['username']
                group_ids = get_allowed_group_ids(conn, user['id'])
                session.clear()
                session.permanent = True
                session['user_id'] = user['id']
                session['role'] = user['role']
                session['group_ids'] = group_ids
                session['username'] = canonical_username
                session['is_admin'] = bool(user['is_admin']) or (user['role'] == 'admin')
                session['permissions'] = json.loads(user['permissions'] or '[]')
                log_action(
                    canonical_username,
                    "ввійшов у систему",
                    group_ids=group_ids,
                    category='auth',
                    action_type='login',
                    section='Авторизація',
                    entity_type='user',
                    entity_id=user['id'],
                    entity_name=canonical_username,
                    ip_address=ip,
                    details=f"роль: {user['role']}, груп: {len(group_ids)}, IP: {ip}",
                    meta={
                        'Роль': 'Адміністратор' if session['is_admin'] else 'Працівник',
                        'Доступних груп': len(group_ids),
                        'Прав': len(session['permissions']),
                    },
                )
                return redirect(url_for('students.student_list'))
            else:
                log_action(
                    username or "невідомий",
                    "невдала спроба входу",
                    level='WARNING',
                    category='security',
                    action_type='login_failed',
                    section='Авторизація',
                    ip_address=ip,
                    details=f"логін: '{username}', IP: {ip}",
                    meta={'Введений логін': username or '—', 'Причина': 'Невірний логін або пароль'},
                )
                flash('Невірний логін або пароль', 'error')

        except Exception as e:
            logger.error(f"Помилка бази даних під час спроби входу (логін: '{username}'): {e}", exc_info=True)
            flash('Тимчасова помилка сервера. Спробуйте увійти пізніше.', 'error')
        finally:
            if conn is not None:
                conn.close()

    return render_template('login.html')


@auth_bp.route('/logout')
def logout():
    """Завершує сесію користувача, логує подію "вийшов із системи" та повертає на форму логіну."""
    username = current_username()
    group_ids = session.get('group_ids', [])
    log_action(
        username,
        "вийшов із системи",
        group_ids=group_ids,
        category='auth',
        action_type='logout',
        section='Авторизація',
        entity_type='user',
        entity_id=session.get('user_id'),
        entity_name=username,
        ip_address=client_ip(),
        details=f"IP: {client_ip()}"
    )
    session.clear()
    return redirect(url_for('auth.login'))