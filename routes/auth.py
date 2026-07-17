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

    GET  - показує форму логіну (якщо користувач вже залогінений -
           одразу редіректить на список студентів).
    POST - перевіряє логін/пароль проти таблиці users. При успіху
           заповнює сесію (user_id, role, group_ids, username, is_admin,
           permissions) та логує подію "ввійшов у систему". При невдачі
           логує подію "невдала спроба входу" (це важливо для
           відстеження спроб підбору пароля) і показує повідомлення
           про помилку без деталізації причини (щоб не підказувати
           зловмиснику, чи існує такий логін).
    """
    if 'user_id' in session:
        return redirect(url_for('students.student_list'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        conn = None
        try:
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
                return redirect(url_for('students.student_list'))
            else:
                log_action(
                    username or "невідомий",
                    "невдала спроба входу",
                    details=f"логін: '{username}', IP: {request.remote_addr}"
                )
                flash('Невірний логін або пароль', 'error')

        except Exception as e:
            # Раніше помилка БД під час логіну (наприклад, файл students.db
            # заблокований, пошкоджений, чи недоступний за правами доступу)
            # призвела б до необробленого винятку і "голої" сторінки 500
            # замість зрозумілого користувачу повідомлення. З'єднання conn
            # також могло залишитись незакритим (витік з'єднань), бо
            # close() викликався лише на "щасливому" шляху виконання.
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
        details=f"IP: {request.remote_addr}"
    )
    session.clear()
    return redirect(url_for('auth.login'))
