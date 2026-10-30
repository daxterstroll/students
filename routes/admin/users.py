"""
routes/admin/users.py
======================
Користувачі та їхні права доступу, журнал дій (app.log).
"""
from flask import render_template, request, redirect, url_for, flash, session, send_file
from routes.db import get_db
from routes.utils import log_action, permission_required, logger
from routes.helpers import current_username, sort_ukrainian
from routes.admin import admin_bp
import sqlite3
import os
from werkzeug.security import generate_password_hash
from datetime import datetime
import json
import re

PERMISSIONS = [
    'manage_users',
    'view_logs',
    'group_export',
    'import_from_excel',
    'manage_education_documents',
    'manage_passport_documents',
    'import_passport_documents',
    'study_periods',
    'manage_groups',
    'manage_subjects',
    'manage_activities',
    'import_subjects',
    'archive',
    'manage_students',
    'manage_accreditations',
    'manage_diplomas',
    'import_education_docs',
    'manage_templates',
    'import_grades',
    'analytics',
    'manage_specialties',
    'manage_degree_levels',
    'manage_educational_programs',
    'manage_qualification_names',
    'manage_courses',
    'manage_frozen_students',
    'manage_expulsion',
    'manage_licenses',
    'manage_license_transfer',
    'license_report',
    'manage_pending_students',
    'manage_military_reports'
]


_LOG_ENTRY_RE = re.compile(r'(\d{4}-\d{2}-\d{2})\s*\|\s*(\d{2}:\d{2}:\d{2})\s*\|\s*(\w+)\s*\|\s*(.*)')
_LOG_USER_RE = re.compile(r'👤\s*([^\s-][^-]*?)\s+-\s+(.*)')


def _parse_log_file(path):
    """Читає один місячний лог-файл і розбирає кожен рядок на дату/час/рівень/користувача/дію."""
    entries = []
    if not os.path.exists(path):
        return entries
    try:
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                entry = {'raw': line, 'date': '', 'time': '', 'level': 'INFO', 'username': '', 'action': line}
                m = _LOG_ENTRY_RE.match(line)
                if m:
                    entry['date'] = m.group(1)
                    entry['time'] = m.group(2)
                    entry['level'] = m.group(3)
                    rest = m.group(4).strip()
                    entry['action'] = rest
                    u = _LOG_USER_RE.match(rest)
                    if u:
                        entry['username'] = u.group(1).strip()
                        entry['action'] = u.group(2).strip()
                entries.append(entry)
    except Exception as e:
        logger.error(f"Помилка при читанні логів {path}: {e}")
    return entries


@admin_bp.route('/admin/view_logs')
@permission_required('view_logs')
def view_logs():
    """
    Показує журнал дій за обраний місяць (logs/app-YYYY-MM.log) у зручному
    розібраному вигляді: дата/час/рівень/користувач/дія, з фільтрацією на
    фронтенді. Раніше читало і парсило ОДИН суцільний app.log, що з часом
    ріс без обмежень - тепер на кожен запит вичитується лише файл одного
    місяця (обраного через ?month=YYYY-MM, за замовчуванням - поточний).
    """
    from routes.utils import LOGS_DIR

    month_pattern = re.compile(r'^app-(\d{4}-\d{2})\.log$')
    available_months = []
    if os.path.isdir(LOGS_DIR):
        for fname in os.listdir(LOGS_DIR):
            m = month_pattern.match(fname)
            if m:
                fpath = os.path.join(LOGS_DIR, fname)
                available_months.append({'value': m.group(1), 'size_kb': round(os.path.getsize(fpath) / 1024, 1)})
    available_months.sort(key=lambda x: x['value'], reverse=True)

    current_month = datetime.now().strftime('%Y-%m')
    selected_month = request.args.get('month', current_month)
    if not re.match(r'^\d{4}-\d{2}$', selected_month):
        selected_month = current_month

    parsed_logs = _parse_log_file(os.path.join(LOGS_DIR, f"app-{selected_month}.log"))
    parsed_logs.reverse()
    usernames = sorted({e['username'] for e in parsed_logs if e['username']})

    # Останні помилки за кілька останніх місяців (а не лише з обраного) -
    # щоб критичне не губилось, якщо зараз дивишся старий місяць. Межа в
    # 6 місяців - щоб не вичитувати всю історію на кожне відкриття сторінки.
    recent_errors = []
    for month_info in available_months[:6]:
        for entry in _parse_log_file(os.path.join(LOGS_DIR, f"app-{month_info['value']}.log")):
            if entry['level'] == 'ERROR':
                recent_errors.append({**entry, 'month': month_info['value']})
    recent_errors.sort(key=lambda e: (e['date'], e['time']), reverse=True)
    recent_errors = recent_errors[:20]

    log_action(current_username(), "переглянув журнал дій")

    from datetime import date
    return render_template('view_logs.html', logs=parsed_logs, usernames=usernames,
                           now=date.today().strftime('%Y-%m-%d'),
                           available_months=available_months, selected_month=selected_month,
                           current_month=current_month, recent_errors=recent_errors)


@admin_bp.route('/admin/download_log')
@permission_required('view_logs')
def download_log():
    """Завантажити .log-файл обраного місяця (напр. щоб переслати розробнику)."""
    from routes.utils import LOGS_DIR
    month = request.args.get('month', '')
    if not re.match(r'^\d{4}-\d{2}$', month):
        flash("Невірний формат місяця.", "error")
        return redirect(url_for('admin.view_logs'))
    fpath = os.path.join(LOGS_DIR, f"app-{month}.log")
    if not os.path.exists(fpath):
        flash("Файл логів за цей місяць не знайдено.", "error")
        return redirect(url_for('admin.view_logs'))
    log_action(current_username(), f"завантажив лог за {month}")
    return send_file(fpath, as_attachment=True, download_name=f"app-{month}.log")


@admin_bp.route('/admin/users', methods=['GET', 'POST'])
@permission_required('manage_users')
def manage_users():
    """Список користувачів системи та редагування їхніх прав доступу (is_admin + список дозволів permissions)."""
    conn = get_db()
    try:
        users = conn.execute("""
            SELECT u.id, u.username, u.role, u.is_admin, u.permissions,
                   GROUP_CONCAT(g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)', ', ') AS group_names
            FROM users u
            LEFT JOIN user_groups ug ON u.id = ug.user_id
            LEFT JOIN groups g ON ug.group_id = g.id
            GROUP BY u.id ORDER BY u.username
        """).fetchall()

        perm_names_ua = {
            'manage_users': 'Список користувачів та управління правами',
            'view_logs': 'Журнал дій',
            'group_export': 'Масова генерація документів',
            'import_from_excel': 'Інпорт студентів',
            'manage_education_documents': 'Управління документами про освіту',
            'manage_passport_documents': 'Управління паспортними даними',
            'import_passport_documents': 'Імпорт паспортних даних',
            'study_periods': 'Періоди навчання',
            'manage_groups': 'Управління групами',
            'manage_subjects': 'Предмети',
            'manage_activities': 'Управління діяльностями',
            'import_subjects': 'Імпорт предметів з Excel',
            'archive': 'Управління архівом',
            'manage_students': 'Управління студентами (Видалення студента та його війс. док.)',
            'manage_accreditations': 'Управління акредетаціями',
            'manage_diplomas': 'Управління номерами диплому і додатку',
            'import_education_docs': 'Управління імпортом документів',
            'manage_templates': 'Управління шаблонами документів',
            'import_grades': 'Імпорт оцінок з Excel',
            'analytics': 'Аналітика',
            'manage_specialties': 'Спеціальності',
            'manage_degree_levels': 'Ступені',
            'manage_educational_programs': 'Освітня програма',
            'manage_qualification_names': 'Назва кваліфікації',
            'manage_courses': 'Курси (та переведення на курс / випуск)',
            'manage_frozen_students': 'Заморожені студенти',
            'manage_expulsion': 'Наказ про відрахування',
            'manage_licenses': 'Ліцензії',
            'manage_license_transfer': 'Перевести між ліцензіями',
            'license_report': 'Звіт по ліцензіях',
            'manage_pending_students': 'Заявки на реєстрацію (публічна анкета)',
            'manage_military_reports': 'Військовий облік (Додаток 5)',
        }

        if request.method == 'POST':
            user_id = request.form.get('user_id')
            if not user_id:
                flash('Не вказано користувача', 'danger')
                return redirect(url_for('admin.manage_users'))

            is_admin = 1 if 'is_admin' in request.form else 0
            selected_perms = [p for p in PERMISSIONS if p in request.form]

            target_user = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()

            conn.execute("UPDATE users SET is_admin=?, permissions=? WHERE id=?",
                         (is_admin, json.dumps(selected_perms), user_id))
            conn.commit()

            log_action(
                current_username(),
                f"змінив права: {target_user['username']} (ID {user_id})",
                details=f"is_admin: {bool(is_admin)}, дозволи: {', '.join(selected_perms) or 'жодного'}"
            )
            flash('Права успішно оновлено', 'success')
            return redirect(url_for('admin.manage_users'))

        return render_template('manage_users.html', users=users, permissions=PERMISSIONS, perm_names_ua=perm_names_ua)

    except sqlite3.Error as e:
        logger.error(f"Помилка бази даних у manage_users: {e}", exc_info=True)
        flash(f'Помилка бази даних: {e}', 'danger')
        return redirect(url_for('admin.manage_users'))
    finally:
        conn.close()


@admin_bp.route('/admin/users/add', methods=['GET', 'POST'])
@permission_required('manage_users')
def add_user():
    """Форма створення нового користувача (логін/пароль/роль/групи)."""
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        role = request.form.get('role')
        group_ids = request.form.getlist('group_id')

        if not all([username, password, role]):
            flash('Заповніть усі обовязкові поля', 'danger')
            return redirect(url_for('admin.add_user'))

        is_admin = 1 if role == 'admin' else 0
        permissions = json.dumps([])

        conn = get_db()
        try:
            exists = conn.execute("SELECT 1 FROM users WHERE username=?", (username,)).fetchone()
            if exists:
                flash(f'Користувач з іменем "{username}" вже існує', 'danger')
                return redirect(url_for('admin.add_user'))

            cursor = conn.execute(
                "INSERT INTO users (username, password_hash, role, is_admin, permissions) VALUES (?, ?, ?, ?, ?)",
                (username, generate_password_hash(password), role, is_admin, permissions)
            )

            user_id = cursor.lastrowid

            for gid in group_ids:
                if gid:
                    conn.execute("INSERT INTO user_groups (user_id, group_id) VALUES (?, ?)", (user_id, gid))

            conn.commit()
            log_action(
                current_username(),
                f"додав користувача: {username} (ID {user_id})",
                details=f"роль: {role}, груп: {len(group_ids)}"
            )
            flash('Користувача успішно додано', 'success')
            return redirect(url_for('admin.manage_users'))

        except sqlite3.IntegrityError as e:
            logger.error(f"Помилка БД при додаванні користувача '{username}': {e}", exc_info=True)
            flash(f'Помилка бази даних: {e}', 'danger')
        finally:
            conn.close()

    conn = get_db()
    try:
        groups = conn.execute("""
            SELECT id, name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
            FROM groups ORDER BY name, start_year
        """).fetchall()
    finally:
        conn.close()

    return render_template('add_user.html', groups=groups)


@admin_bp.route('/admin/users/<int:user_id>/edit', methods=['GET', 'POST'])
@permission_required('manage_users')
def edit_user(user_id):
    """Форма редагування ролі та прив'язаних груп існуючого користувача."""
    conn = get_db()
    try:
        if request.method == 'POST':
            role = request.form.get('role')
            group_ids = request.form.getlist('group_id')

            if not role:
                flash('Роль обовязкова', 'danger')
                return redirect(url_for('admin.edit_user', user_id=user_id))

            is_admin = 1 if role == 'admin' else 0
            user_row = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()

            conn.execute("UPDATE users SET role=?, is_admin=? WHERE id=?", (role, is_admin, user_id))
            conn.execute("DELETE FROM user_groups WHERE user_id=?", (user_id,))
            for gid in group_ids:
                if gid:
                    conn.execute("INSERT INTO user_groups (user_id, group_id) VALUES (?, ?)", (user_id, gid))

            conn.commit()
            log_action(
                current_username(),
                f"змінив роль/групи: {user_row['username']} (ID {user_id})",
                details=f"роль: {role}, груп: {len(group_ids)}"
            )
            flash('Дані користувача оновлено', 'success')
            return redirect(url_for('admin.manage_users'))

        user = conn.execute("SELECT id, username, role FROM users WHERE id=?", (user_id,)).fetchone()
        if not user:
            flash('Користувача не знайдено', 'danger')
            return redirect(url_for('admin.manage_users'))

        current_groups = conn.execute("SELECT group_id FROM user_groups WHERE user_id=?", (user_id,)).fetchall()
        current_group_ids = [row['group_id'] for row in current_groups]

        groups = conn.execute("""
            SELECT id, name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
            FROM groups ORDER BY name, start_year
        """).fetchall()

        return render_template('edit_user.html', user=user, groups=groups, current_group_ids=current_group_ids)
    finally:
        conn.close()


@admin_bp.route('/admin/users/<int:user_id>/change-password', methods=['GET', 'POST'])
@permission_required('manage_users')
def change_password(user_id):
    """Форма зміни пароля вказаного користувача адміністратором."""
    if request.method == 'POST':
        password = request.form.get('password')
        if not password or len(password) < 6:
            flash('Пароль повинен бути не коротшим 6 символів', 'danger')
            return redirect(url_for('admin.change_password', user_id=user_id))

        conn = get_db()
        try:
            target = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()
            conn.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(password), user_id))
            conn.commit()
            log_action(
                current_username(),
                f"змінив пароль: {target['username']} (ID {user_id})"
            )
            flash('Пароль успішно змінено', 'success')
            return redirect(url_for('admin.manage_users'))
        finally:
            conn.close()

    conn = get_db()
    try:
        user = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()
        if not user:
            flash('Користувача не знайдено', 'danger')
            return redirect(url_for('admin.manage_users'))
    finally:
        conn.close()

    return render_template('change_password.html', user_id=user_id, username=user['username'])


@admin_bp.route('/admin/users/<int:user_id>/delete', methods=['POST'])
@permission_required('manage_users')
def delete_user(user_id):
    """Видалення користувача та його зв'язків з групами (user_groups)."""
    conn = get_db()
    try:
        username_row = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()
        if not username_row:
            flash('Користувача не знайдено', 'danger')
            return redirect(url_for('admin.manage_users'))

        conn.execute("DELETE FROM users WHERE id=?", (user_id,))
        conn.execute("DELETE FROM user_groups WHERE user_id=?", (user_id,))
        conn.commit()

        log_action(
            current_username(),
            f"ВИДАЛИВ користувача: {username_row['username']} (ID {user_id})"
        )
        flash('Користувача успішно видалено', 'success')
    except sqlite3.Error as e:
        logger.error(f"Помилка БД при видаленні користувача (ID {user_id}): {e}", exc_info=True)
        flash(f'Помилка при видаленні: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('admin.manage_users'))