"""
routes/admin/users.py
======================
Користувачі: ролі, права (permissions), доступ до груп / спеціальностей / ступенів.
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

# ── Усі права + підписи + групи для UI ──────────────────────────
PERMISSION_DEFINITIONS = [
    # (code, label, group_key)
    ('manage_users', 'Користувачі та права', 'system'),
    ('view_logs', 'Журнал дій', 'system'),

    ('manage_students', 'Студенти (картки, список)', 'students'),
    ('manage_pending_students', 'Заявки на вступ', 'students'),
    ('study_periods', 'Періоди навчання', 'students'),
    ('manage_passport_documents', 'Паспортні документи', 'students'),
    ('manage_education_documents', 'Документи про освіту', 'students'),
    

    ('manage_groups', 'Групи', 'education'),
    ('manage_specialties', 'Спеціальності', 'education'),
    ('manage_degree_levels', 'Ступені', 'education'),
    ('manage_educational_programs', 'Освітні програми', 'education'),
    ('manage_qualification_names', 'Назви кваліфікацій', 'education'),
    ('manage_accreditations', 'Акредитації', 'education'),
    ('manage_subjects', 'Предмети', 'education'),
    ('manage_activities', 'Практики / курсові / атестації', 'education'),

    ('manage_courses', 'Курси (канбан, переведення)', 'movement'),
    ('manage_expulsion', 'Відрахування', 'movement'),
    ('manage_frozen_students', 'Заморожені студенти', 'movement'),
    ('archive', 'Архів груп', 'movement'),

    ('import_from_excel', 'Імпорт студентів з Excel', 'import'),
    ('import_passport_documents', 'Імпорт паспортів', 'import'),
    ('import_education_docs', 'Імпорт документів про освіту', 'import'),
    ('import_subjects', 'Імпорт предметів', 'import'),
    ('import_grades', 'Імпорт оцінок', 'import'),
    ('import_military_pdf', 'Масовий імпорт PDF «Резерв+»', 'import'),
    ('bulk_photos', 'Масове завантаження фото', 'import'),

    ('group_export', 'Експорт групи / фото', 'docs'),
    ('manage_diplomas', 'Дипломи та додатки', 'docs'),
    ('manage_templates', 'Шаблони документів', 'docs'),

    ('manage_licenses', 'Ліцензії закладу', 'licenses'),
    ('manage_license_transfer', 'Переведення між ліцензіями', 'licenses'),
    ('license_report', 'Звіт по ліцензіях', 'licenses'),

    ('analytics', 'Аналітика', 'reports'),
    ('custom_export', 'Кастомний експорт даних', 'reports'),
    ('manage_military_reports', 'Військовий облік (дашборд, додатки)', 'reports'),
]

PERMISSIONS = [p[0] for p in PERMISSION_DEFINITIONS]

PERMISSION_LABELS = {code: label for code, label, _ in PERMISSION_DEFINITIONS}

PERMISSION_GROUPS = [
    ('system', 'Система'),
    ('students', 'Студенти та документи'),
    ('education', 'Освітня програма'),
    ('movement', 'Рух контингенту'),
    ('import', 'Імпорт і масові операції'),
    ('docs', 'Документи та експорт'),
    ('licenses', 'Ліцензії'),
    ('reports', 'Звіти та військовий облік'),
]


def _ensure_user_scope_tables(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_specialties (
            user_id INTEGER NOT NULL,
            specialty_code TEXT NOT NULL,
            PRIMARY KEY (user_id, specialty_code),
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_degree_levels (
            user_id INTEGER NOT NULL,
            degree_level TEXT NOT NULL,
            PRIMARY KEY (user_id, degree_level),
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        )
    """)
    conn.commit()


def _permissions_by_group():
    grouped = {k: [] for k, _ in PERMISSION_GROUPS}
    for code, label, gkey in PERMISSION_DEFINITIONS:
        grouped.setdefault(gkey, []).append({'code': code, 'label': label})
    return grouped


_LOG_ENTRY_RE = re.compile(
    r'(\d{4}-\d{2}-\d{2})\s*\|\s*(\d{2}:\d{2}:\d{2})\s*\|\s*(\w+)\s*\|\s*(.*)'
)
_LOG_USER_RE = re.compile(r'👤\s+(.+?)\s+-\s+(.*)$')


def _parse_log_file(path):
    entries = []
    if not os.path.exists(path):
        return entries
    try:
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                entry = {
                    'raw': line, 'date': '', 'time': '', 'level': 'INFO',
                    'username': '', 'action': line,
                }
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
    except OSError as e:
        logger.error(f'log read error: {e}')
    return entries


@admin_bp.route('/admin/view_logs')
@permission_required('view_logs')
def view_logs():
    """
    Журнал дій за обраний місяць (logs/app-YYYY-MM.log).
    """
    from routes.utils import LOGS_DIR
    from datetime import date

    month_pattern = re.compile(r'^app-(\d{4}-\d{2})\.log$')
    available_months = []
    if os.path.isdir(LOGS_DIR):
        for fname in os.listdir(LOGS_DIR):
            m = month_pattern.match(fname)
            if m:
                fpath = os.path.join(LOGS_DIR, fname)
                available_months.append({
                    'value': m.group(1),
                    'size_kb': round(os.path.getsize(fpath) / 1024, 1),
                })
    available_months.sort(key=lambda x: x['value'], reverse=True)

    current_month = datetime.now().strftime('%Y-%m')
    selected_month = request.args.get('month', current_month)
    if not re.match(r'^\d{4}-\d{2}$', selected_month):
        selected_month = current_month

    parsed_logs = _parse_log_file(
        os.path.join(LOGS_DIR, f'app-{selected_month}.log')
    )
    parsed_logs.reverse()
    usernames = sorted({e['username'] for e in parsed_logs if e.get('username')})

    recent_errors = []
    for month_info in available_months[:6]:
        for entry in _parse_log_file(
            os.path.join(LOGS_DIR, f"app-{month_info['value']}.log")
        ):
            if entry.get('level') == 'ERROR':
                recent_errors.append({**entry, 'month': month_info['value']})
    recent_errors.sort(key=lambda e: (e.get('date', ''), e.get('time', '')), reverse=True)
    recent_errors = recent_errors[:20]

    log_action(current_username(), 'переглянув журнал дій')

    return render_template(
        'view_logs.html',
        logs=parsed_logs,
        usernames=usernames,
        now=date.today().strftime('%Y-%m-%d'),
        available_months=available_months,
        selected_month=selected_month,
        current_month=current_month,
        recent_errors=recent_errors,
    )


@admin_bp.route('/admin/download_log')
@permission_required('view_logs')
def download_log():
    """Завантажити .log за обраний місяць."""
    from routes.utils import LOGS_DIR
    month = request.args.get('month', '')
    if not re.match(r'^\d{4}-\d{2}$', month):
        flash('Невірний формат місяця.', 'error')
        return redirect(url_for('admin.view_logs'))
    fpath = os.path.join(LOGS_DIR, f'app-{month}.log')
    if not os.path.exists(fpath):
        flash('Файл логів за цей місяць не знайдено.', 'error')
        return redirect(url_for('admin.view_logs'))
    log_action(current_username(), f'завантажив лог за {month}')
    return send_file(fpath, as_attachment=True, download_name=f'app-{month}.log')


@admin_bp.route('/admin/users')
@permission_required('manage_users')
def manage_users():
    conn = get_db()
    conn.row_factory = sqlite3.Row
    users = conn.execute(
        "SELECT id, username, role, is_admin, permissions FROM users ORDER BY username"
    ).fetchall()
    rows = []
    for u in users:
        d = dict(u)
        try:
            d['perm_list'] = json.loads(d['permissions'] or '[]')
        except Exception:
            d['perm_list'] = []
        d['perm_labels'] = [
            PERMISSION_LABELS.get(p, p) for p in d['perm_list']
        ]
        rows.append(d)
    conn.close()
    return render_template('manage_users.html', users=rows)


def _load_scope_options(conn):
    groups = conn.execute("""
        SELECT id,
               name || ' (' || start_year || ', ' || study_form || ', ' ||
               program_credits || ' кр.)' AS display_name,
               specialty_code, degree_level
        FROM groups
        WHERE COALESCE(archived, 0) = 0
        ORDER BY name COLLATE UKRAINIAN, start_year
    """).fetchall()
    specialties = conn.execute("""
        SELECT code, name_ua FROM specialties
        WHERE COALESCE(is_active, 1) = 1
        ORDER BY code
    """).fetchall()
    degrees = conn.execute("""
        SELECT name_ua FROM degree_levels
        WHERE COALESCE(is_active, 1) = 1
        ORDER BY id
    """).fetchall()
    return groups, specialties, degrees


@admin_bp.route('/admin/users/add', methods=['GET', 'POST'])
@permission_required('manage_users')
def add_user():
    conn = get_db()
    conn.row_factory = sqlite3.Row
    _ensure_user_scope_tables(conn)

    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        role = request.form.get('role') or 'user'
        group_ids = [g for g in request.form.getlist('group_id') if g]
        specialty_codes = [s for s in request.form.getlist('specialty_code') if s]
        degree_levels = [d for d in request.form.getlist('degree_level') if d]
        perms = [p for p in request.form.getlist('permissions') if p in PERMISSIONS]

        if not username or not password:
            flash("Логін і пароль обов'язкові", 'error')
            conn.close()
            return redirect(url_for('admin.add_user'))
        if len(password) < 6:
            flash('Пароль не коротший за 6 символів', 'error')
            conn.close()
            return redirect(url_for('admin.add_user'))

        is_admin = 1 if role == 'admin' else 0
        if is_admin:
            perms = []  # адмін і так має все

        try:
            cur = conn.execute(
                """
                INSERT INTO users (username, password_hash, role, is_admin, permissions)
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    username,
                    generate_password_hash(password),
                    role,
                    is_admin,
                    json.dumps(perms, ensure_ascii=False),
                ),
            )
            uid = cur.lastrowid
            for gid in group_ids:
                conn.execute(
                    "INSERT INTO user_groups (user_id, group_id) VALUES (?, ?)",
                    (uid, int(gid)),
                )
            for code in specialty_codes:
                conn.execute(
                    "INSERT INTO user_specialties (user_id, specialty_code) VALUES (?, ?)",
                    (uid, code),
                )
            for deg in degree_levels:
                conn.execute(
                    "INSERT INTO user_degree_levels (user_id, degree_level) VALUES (?, ?)",
                    (uid, deg),
                )
            conn.commit()
            log_action(
                current_username(),
                f'створив користувача {username}',
                details=f'роль={role}, прав={len(perms)}, груп={len(group_ids)}, '
                        f'спец={len(specialty_codes)}, ступенів={len(degree_levels)}',
            )
            flash(f'Користувача «{username}» створено', 'success')
            conn.close()
            return redirect(url_for('admin.manage_users'))
        except sqlite3.IntegrityError:
            conn.close()
            flash('Такий логін уже існує', 'error')
            return redirect(url_for('admin.add_user'))

    groups, specialties, degrees = _load_scope_options(conn)
    conn.close()
    return render_template(
        'add_user.html',
        groups=groups,
        specialties=specialties,
        degrees=degrees,
        permission_groups=PERMISSION_GROUPS,
        permissions_by_group=_permissions_by_group(),
        selected_permissions=[],
        selected_group_ids=[],
        selected_specialties=[],
        selected_degrees=[],
    )


@admin_bp.route('/admin/users/<int:user_id>/edit', methods=['GET', 'POST'])
@permission_required('manage_users')
def edit_user(user_id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    _ensure_user_scope_tables(conn)

    user = conn.execute(
        "SELECT id, username, role, is_admin, permissions FROM users WHERE id=?",
        (user_id,),
    ).fetchone()
    if not user:
        conn.close()
        flash('Користувача не знайдено', 'error')
        return redirect(url_for('admin.manage_users'))

    if request.method == 'POST':
        role = request.form.get('role') or 'user'
        group_ids = [g for g in request.form.getlist('group_id') if g]
        specialty_codes = [s for s in request.form.getlist('specialty_code') if s]
        degree_levels = [d for d in request.form.getlist('degree_level') if d]
        perms = [p for p in request.form.getlist('permissions') if p in PERMISSIONS]

        is_admin = 1 if role == 'admin' else 0
        if is_admin:
            perms = []

        conn.execute(
            "UPDATE users SET role=?, is_admin=?, permissions=? WHERE id=?",
            (role, is_admin, json.dumps(perms, ensure_ascii=False), user_id),
        )
        conn.execute("DELETE FROM user_groups WHERE user_id=?", (user_id,))
        for gid in group_ids:
            conn.execute(
                "INSERT INTO user_groups (user_id, group_id) VALUES (?, ?)",
                (user_id, int(gid)),
            )
        conn.execute("DELETE FROM user_specialties WHERE user_id=?", (user_id,))
        for code in specialty_codes:
            conn.execute(
                "INSERT INTO user_specialties (user_id, specialty_code) VALUES (?, ?)",
                (user_id, code),
            )
        conn.execute("DELETE FROM user_degree_levels WHERE user_id=?", (user_id,))
        for deg in degree_levels:
            conn.execute(
                "INSERT INTO user_degree_levels (user_id, degree_level) VALUES (?, ?)",
                (user_id, deg),
            )
        conn.commit()
        log_action(
            current_username(),
            f"редагував користувача {user['username']} (ID {user_id})",
            details=f'роль={role}, прав={len(perms)}, груп={len(group_ids)}, '
                    f'спец={len(specialty_codes)}, ступенів={len(degree_levels)}',
        )
        flash('Дані користувача оновлено', 'success')
        conn.close()
        return redirect(url_for('admin.manage_users'))

    try:
        current_perms = json.loads(user['permissions'] or '[]')
    except Exception:
        current_perms = []

    current_group_ids = [
        r['group_id']
        for r in conn.execute(
            "SELECT group_id FROM user_groups WHERE user_id=?", (user_id,)
        ).fetchall()
    ]
    current_specialties = [
        r['specialty_code']
        for r in conn.execute(
            "SELECT specialty_code FROM user_specialties WHERE user_id=?", (user_id,)
        ).fetchall()
    ]
    current_degrees = [
        r['degree_level']
        for r in conn.execute(
            "SELECT degree_level FROM user_degree_levels WHERE user_id=?", (user_id,)
        ).fetchall()
    ]

    groups, specialties, degrees = _load_scope_options(conn)
    conn.close()
    return render_template(
        'edit_user.html',
        user=user,
        groups=groups,
        specialties=specialties,
        degrees=degrees,
        permission_groups=PERMISSION_GROUPS,
        permissions_by_group=_permissions_by_group(),
        selected_permissions=current_perms,
        selected_group_ids=current_group_ids,
        selected_specialties=current_specialties,
        selected_degrees=current_degrees,
    )


@admin_bp.route('/admin/users/<int:user_id>/change-password', methods=['GET', 'POST'])
@permission_required('manage_users')
def change_password(user_id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    user = conn.execute(
        "SELECT id, username FROM users WHERE id=?", (user_id,)
    ).fetchone()
    if not user:
        conn.close()
        flash('Користувача не знайдено', 'error')
        return redirect(url_for('admin.manage_users'))

    if request.method == 'POST':
        password = request.form.get('password') or ''
        if len(password) < 6:
            flash('Пароль не коротший за 6 символів', 'error')
            conn.close()
            return redirect(url_for('admin.change_password', user_id=user_id))
        conn.execute(
            "UPDATE users SET password_hash=? WHERE id=?",
            (generate_password_hash(password), user_id),
        )
        conn.commit()
        log_action(current_username(), f"змінив пароль користувача {user['username']}")
        flash('Пароль оновлено', 'success')
        conn.close()
        return redirect(url_for('admin.manage_users'))

    conn.close()
    return render_template('change_password.html', user=user)


@admin_bp.route('/admin/users/<int:user_id>/delete', methods=['POST'])
@permission_required('manage_users')
def delete_user(user_id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    user = conn.execute(
        "SELECT username FROM users WHERE id=?", (user_id,)
    ).fetchone()
    if not user:
        conn.close()
        flash('Користувача не знайдено', 'error')
        return redirect(url_for('admin.manage_users'))
    if user['username'] == current_username():
        conn.close()
        flash('Не можна видалити власний обліковий запис', 'error')
        return redirect(url_for('admin.manage_users'))

    _ensure_user_scope_tables(conn)
    conn.execute("DELETE FROM user_groups WHERE user_id=?", (user_id,))
    conn.execute("DELETE FROM user_specialties WHERE user_id=?", (user_id,))
    conn.execute("DELETE FROM user_degree_levels WHERE user_id=?", (user_id,))
    conn.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit()
    conn.close()
    log_action(current_username(), f"видалив користувача {user['username']}")
    flash('Користувача видалено', 'success')
    return redirect(url_for('admin.manage_users'))
