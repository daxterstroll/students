"""
routes/admin/users.py
======================
Користувачі: ролі, права (permissions), доступ до груп / спеціальностей / ступенів.
"""
from flask import render_template, request, redirect, url_for, flash, session, send_file, Response
from routes.db import get_db
from routes.utils import (
    log_action,
    permission_required,
    logger,
    ensure_audit_logs_table,
    infer_audit_metadata,
    parse_inline_changes_from_text,
    AUDIT_CATEGORY_UA,
    LOGS_DIR,
)
from routes.helpers import current_username, sort_ukrainian
from routes.admin import admin_bp
from routes.security import ALLOWED_USER_ROLES, validate_password_strength, sanitize_spreadsheet_cell
import sqlite3
import os
import csv
import io
from werkzeug.security import generate_password_hash
from datetime import datetime, date, timedelta
import json
import re

# ── Усі права + підписи + групи для UI ──────────────────────────
PERMISSION_DEFINITIONS = [
    # ─── Система ────────────────────────────────────────────────
    ('manage_users',                'Користувачі та права',                     'system'),
    ('view_logs',                   'Журнал дій',                                'system'),
    ('analytics',                   'Аналітика',                                 'system'),
    ('custom_export',               'Кастомний експорт даних',                    'system'),
    ('manage_templates',            'Шаблони документів',                        'system'),

    # ─── Студенти (базове) ──────────────────────────────────────
    ('manage_students',             'Студенти — редагування, додавання, видалення', 'students'),
    ('view_student_card',           'Картка студента (перегляд)',                  'students'),
    ('view_student_completeness',   'Заповненість профілів (колонка, фільтр і прогалини)', 'students'),
    ('manage_pending_students',     'Заявки на реєстрацію',                      'students'),
    ('manage_update_requests',      'Заявки на оновлення даних',                 'students'),
    ('manage_passport_documents',   'Паспортні дані',                            'students'),
    ('manage_education_documents',  'Документи про освіту',                      'students'),
    ('study_periods',               'Періоди навчання',                          'students'),

    # ─── Освітня програма ───────────────────────────────────────
    ('manage_groups',               'Групи',                                     'education'),
    ('manage_faculties',            'Факультети',                                'education'),
    ('manage_specialties',          'Спеціальності',                             'education'),
    ('manage_degree_levels',        'Ступені',                                   'education'),
    ('manage_educational_programs', 'Освітні програми',                          'education'),
    ('manage_qualification_names',  'Назви кваліфікацій',                        'education'),
    ('manage_accreditations',       'Акредитації',                               'education'),
    ('manage_subjects',             'Предмети',                                  'education'),
    ('manage_activities',           'Практики / курсові / атестації',            'education'),
    ('manage_diplomas',             'Номери диплому / додатку',                  'education'),

    # ─── Рух студентів ──────────────────────────────────────────
    ('manage_courses',              'Курси (дошка)',                             'movement'),
    ('manage_course_transfer',      'Переведення на курс',                       'movement'),
    ('manage_graduation',           'Оформлення випуску',                        'movement'),
    ('manage_expulsion',            'Відрахування',                              'movement'),
    ('manage_frozen_students',      'Заморожені студенти',                       'movement'),
    ('archive',                     'Архів груп',                                'movement'),

    # ─── Накази ─────────────────────────────────────────────────
    ('manage_orders',               'Реєстр наказів',                            'orders'),
    ('create_enrollment_order',     'Наказ про зарахування',                     'orders'),

    # ─── Сесії ──────────────────────────────────────────────────
    ('manage_exam_sessions',        'Сесії та відомості — створення та редагування', 'sessions'),
    ('reopen_exam_sheets',          'Сесії та відомості — повторне відкриття закритих відомостей', 'sessions'),
    ('delete_exam_sessions',        'Сесії та відомості — ВИДАЛЕННЯ',                'sessions'),
    ('view_exam_session_history',   'Історія сесій та відомостей (перегляд)',        'sessions'),
    ('transfer_exam_grades',        'Перенос оцінок із сесій у план',                'sessions'),

    # ─── Імпорт ─────────────────────────────────────────────────
    ('import_from_excel',           'Імпорт студентів з Excel',                  'import'),
    ('import_passport_documents',   'Імпорт паспортів',                          'import'),
    ('import_education_docs',       'Імпорт документів про освіту',              'import'),
    ('import_subjects',             'Імпорт предметів',                          'import'),
    ('import_grades',               'Імпорт оцінок',                             'import'),
    ('import_military_pdf',         'Масовий імпорт PDF «Резерв+»',              'import'),
    ('bulk_photos',                 'Масове завантаження фото',                  'import'),

    # ─── Експорт ────────────────────────────────────────────────
    ('group_export',                'Масова генерація документів',               'export'),

    # ─── Ліцензії ───────────────────────────────────────────────
    ('manage_licenses',             'Ліцензії закладу',                          'licenses'),
    ('manage_license_transfer',     'ПОНОВЛЕННЯ',                                'licenses'),
    ('license_report',              'Звіт по ліцензіях',                         'licenses'),

    # ─── Військовий облік ───────────────────────────────────────
    ('manage_military_reports',     'Військовий облік (дашборд, додатки)',       'military'),
    ('manage_military_telegram',    'Telegram-бот ВОД «Резерв+» (розсилки, прив’язка, перевірка)', 'military'),
]

PERMISSIONS = [p[0] for p in PERMISSION_DEFINITIONS]

PERMISSION_LABELS = {code: label for code, label, _ in PERMISSION_DEFINITIONS}

PERMISSION_GROUPS = [
    ('system',    'Система'),
    ('students',  'Студенти та документи'),
    ('education', 'Освітня програма'),
    ('movement',  'Рух контингенту'),
    ('orders',    'Накази'),
    ('sessions',  'Сесії та відомості'),
    ('import',    'Імпорт і масові операції'),
    ('export',    'Експорт'),
    ('licenses',  'Ліцензії'),
    ('military',  'Військовий облік'),
]


def _ensure_user_scope_tables(conn):
    from routes.user_scope import ensure_user_scope_tables
    ensure_user_scope_tables(conn)
    try:
        conn.commit()
    except Exception:
        pass


def _permissions_by_group():
    grouped = {k: [] for k, _ in PERMISSION_GROUPS}
    for code, label, gkey in PERMISSION_DEFINITIONS:
        grouped.setdefault(gkey, []).append({'code': code, 'label': label})
    return grouped


_LOG_ENTRY_RE = re.compile(
    r'(\d{4}-\d{2}-\d{2})\s*\|\s*(\d{2}:\d{2}:\d{2})\s*\|\s*(\w+)\s*\|\s*(.*)'
)
_LOG_USER_RE = re.compile(r'👤\s+(.+?)\s+-\s+(.*)$')
_LOG_REF_RE = re.compile(r'\s*\[ref:([0-9a-fA-F]{6,32})\]\s*$')
_LOG_TAG_RE = re.compile(r'\s+\((режим|де|групи|IP):\s*([^)]*)\)')

_CATEGORY_TONES = {
    'create': 'ok',
    'import': 'ok',
    'update': 'info',
    'export': 'info',
    'sessions': 'info',
    'permissions': 'warn',
    'orders': 'warn',
    'delete': 'miss',
    'security': 'miss',
    'auth': 'muted',
    'view': 'muted',
    'system': 'muted',
}

_ENTITY_TYPE_UA = {
    'student': 'Студент',
    'order': 'Наказ',
    'exam_session': 'Сесія',
    'exam_sheet': 'Відомість',
    'user': 'Користувач',
    'passport_document': 'Паспортні дані',
    'education_document': 'Документ про освіту',
    'template': 'Шаблон',
    'license': 'Ліцензія',
    'specialty': 'Спеціальність',
    'degree_level': 'Ступінь',
    'educational_program': 'Освітня програма',
    'qualification_name': 'Кваліфікація',
    'tck': 'ТЦК',
    'teacher': 'Викладач',
    'group': 'Група',
    'pending_student': 'Заявка на реєстрацію',
    'update_request': 'Заявка на оновлення',
    'study_period': 'Період навчання',
    'military': 'Військовий облік',
    'activity_grade': 'Оцінка активності',
    'grade': 'Оцінка',
}

_ROLE_UA = {
    'admin': 'Адміністратор',
    'user': 'Користувач',
    'teacher': 'Викладач',
    'system': 'Система',
}

_META_KEY_UA = {
    'student_id': 'ID студента',
    'group_id': 'ID групи',
    'order_id': 'ID наказу',
    'sheet_id': 'ID відомості',
    'session_id': 'ID сесії',
    'user_id': 'ID користувача',
    'username': 'Логін',
    'role': 'Роль',
    'mode': 'Режим',
    'reason': 'Причина',
    'ip': 'IP-адреса',
    'count': 'Кількість',
    'updated': 'Оновлено',
    'created': 'Створено',
    'skipped': 'Пропущено',
    'failed': 'З помилкою',
    'errors': 'Помилок',
    'added': 'Додано',
    'removed': 'Видалено',
    'filename': 'Файл',
    'file': 'Файл',
    'year': 'Рік',
    'semester': 'Семестр',
    'subject': 'Предмет',
    'teacher': 'Викладач',
    'records': 'Записів',
    'status': 'Статус',
    'imported': 'Імпортовано',
    'matched': 'Співставлено',
    'total': 'Всього',
    'perms_count': 'Прав доступу',
    'groups_count': 'Груп у доступі',
}


def _format_date_ua(date_str):
    if not date_str or len(date_str) < 10:
        return date_str or ''
    parts = date_str[:10].split('-')
    if len(parts) == 3:
        return f"{parts[2]}.{parts[1]}.{parts[0]}"
    return date_str


def _build_search_blob(entry):
    parts = [
        entry.get('date', ''),
        entry.get('date_ua', ''),
        entry.get('time', ''),
        entry.get('level', ''),
        entry.get('username', ''),
        entry.get('user_role_label', ''),
        entry.get('category_label', ''),
        entry.get('section', ''),
        entry.get('action', ''),
        entry.get('entity_label', ''),
        entry.get('entity_name', ''),
        entry.get('entity_id', ''),
        entry.get('group_names', ''),
        entry.get('mode', ''),
        entry.get('details_text', ''),
        entry.get('route_path', ''),
        entry.get('ip_address', ''),
        entry.get('user_agent', ''),
    ]
    for ch in entry.get('changes') or []:
        parts.extend([ch.get('label', ''), ch.get('old', ''), ch.get('new', '')])
    for mi in entry.get('meta_items') or []:
        parts.extend([mi.get('label', ''), mi.get('value', '')])
    return ' '.join(str(p) for p in parts if p).lower()


def _format_meta_items(meta_obj):
    if not isinstance(meta_obj, dict):
        return []
    items = []
    for k, v in meta_obj.items():
        if v is None or v == '' or v == []:
            continue
        lbl = _META_KEY_UA.get(k, str(k))
        if isinstance(v, bool):
            val_str = 'Так' if v else 'Ні'
        elif isinstance(v, (list, tuple)):
            val_str = ', '.join(str(x) for x in v)
        elif isinstance(v, dict):
            val_str = ', '.join(f"{dk}: {dv}" for dk, dv in v.items())
        else:
            val_str = str(v)
        items.append({'label': lbl, 'value': val_str})
    return items


def _audit_row_to_entry(row):
    d = dict(row)
    created_at = (d.get('created_at') or '').strip()
    date_part = created_at[:10] if len(created_at) >= 10 else ''
    time_part = created_at[11:19] if len(created_at) >= 19 else ''

    changes = []
    if d.get('changes_json'):
        try:
            parsed = json.loads(d['changes_json'])
            if isinstance(parsed, list):
                for item in parsed:
                    if isinstance(item, dict):
                        changes.append({
                            'label': str(item.get('label') or item.get('field') or 'Поле'),
                            'old': str(item.get('old') if item.get('old') is not None else '—'),
                            'new': str(item.get('new') if item.get('new') is not None else '—'),
                        })
        except Exception:
            pass

    meta_items = []
    if d.get('details_json'):
        try:
            parsed_meta = json.loads(d['details_json'])
            meta_items = _format_meta_items(parsed_meta)
        except Exception:
            pass

    category = (d.get('category') or 'system').strip()
    section = (d.get('section') or 'Система').strip()
    entity_type = d.get('entity_type') or ''
    user_role = d.get('user_role') or ''

    entry = {
        'id': d.get('id'),
        'log_uid': d.get('log_uid') or '',
        'date': date_part,
        'date_ua': _format_date_ua(date_part),
        'time': time_part,
        'level': (d.get('level') or 'INFO').upper(),
        'username': d.get('username') or '',
        'user_role': user_role,
        'user_role_label': _ROLE_UA.get(user_role, user_role),
        'category': category,
        'category_label': AUDIT_CATEGORY_UA.get(category, category),
        'category_tone': _CATEGORY_TONES.get(category, 'muted'),
        'action_type': d.get('action_type') or 'other',
        'section': section,
        'action': d.get('action_text') or '',
        'entity_type': entity_type,
        'entity_label': _ENTITY_TYPE_UA.get(entity_type, entity_type),
        'entity_id': d.get('entity_id') or '',
        'entity_name': d.get('entity_name') or '',
        'entity_url': d.get('entity_url') or '',
        'group_names': d.get('group_names') or '',
        'mode': d.get('mode') or '',
        'changes': changes,
        'meta_items': meta_items,
        'details_text': d.get('details_text') or '',
        'route_path': d.get('route_path') or '',
        'http_method': d.get('http_method') or '',
        'ip_address': d.get('ip_address') or '',
        'user_agent': d.get('user_agent') or '',
        'source': 'db',
        'raw': f"{created_at} | {d.get('level') or 'INFO'} | {d.get('username') or ''} - {d.get('action_text') or ''}",
    }
    entry['search_blob'] = _build_search_blob(entry)
    return entry


def _parse_log_file(path, *, skip_uids=None):
    entries = []
    if not os.path.exists(path):
        return entries
    skip_uids = skip_uids or set()
    try:
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue

                log_uid = ''
                ref_m = _LOG_REF_RE.search(line)
                if ref_m:
                    log_uid = ref_m.group(1).lower()
                    if log_uid in skip_uids:
                        continue
                    line_clean = line[:ref_m.start()].strip()
                else:
                    line_clean = line

                date_part = ''
                time_part = ''
                level = 'INFO'
                username = ''
                action_raw = line_clean

                m = _LOG_ENTRY_RE.match(line_clean)
                if m:
                    date_part = m.group(1)
                    time_part = m.group(2)
                    level = m.group(3).upper()
                    rest = m.group(4).strip()
                    action_raw = rest
                    u = _LOG_USER_RE.match(rest)
                    if u:
                        username = u.group(1).strip()
                        action_raw = u.group(2).strip()

                # Розбиваємо хвостові секції " | Зміни: ..." та " | деталі"
                segments = [seg.strip() for seg in action_raw.split(' | ') if seg.strip()]
                head = segments[0] if segments else action_raw
                changes_str = ''
                details_parts = []
                for seg in segments[1:]:
                    if seg.startswith('Зміни:'):
                        changes_str = seg[len('Зміни:'):].strip()
                    else:
                        details_parts.append(seg)
                details_text = ' | '.join(details_parts)

                # Витягуємо теги (режим: ...) (де: ...) (групи: ...) (IP: ...)
                mode = ''
                where_section = ''
                route_path = ''
                group_names = ''
                ip_address = ''

                def _consume_tag(match):
                    nonlocal mode, where_section, route_path, group_names, ip_address
                    key, val = match.group(1), match.group(2).strip()
                    if key == 'режим':
                        mode = val
                    elif key == 'де':
                        if ' · ' in val:
                            where_section, route_path = [p.strip() for p in val.split(' · ', 1)]
                        else:
                            where_section = val
                    elif key == 'групи':
                        group_names = val
                    elif key == 'IP':
                        ip_address = val
                    return ''

                action_clean = _LOG_TAG_RE.sub(_consume_tag, head).strip()

                if not ip_address and details_text:
                    ip_m = re.search(r'(?:^|[,;\s])IP[:=]\s*([0-9a-fA-F:.]+)', details_text, re.IGNORECASE)
                    if ip_m:
                        ip_address = ip_m.group(1)

                changes = parse_inline_changes_from_text(changes_str or details_text)
                inferred = infer_audit_metadata(action_clean, details_text, route_path=route_path)
                section = where_section or inferred['section']
                category = inferred['category']
                entity_type = inferred.get('entity_type') or ''

                entry = {
                    'id': None,
                    'log_uid': log_uid,
                    'date': date_part,
                    'date_ua': _format_date_ua(date_part),
                    'time': time_part,
                    'level': level,
                    'username': username,
                    'user_role': '',
                    'user_role_label': '',
                    'category': category,
                    'category_label': AUDIT_CATEGORY_UA.get(category, category),
                    'category_tone': _CATEGORY_TONES.get(category, 'muted'),
                    'action_type': inferred.get('action_type') or 'other',
                    'section': section,
                    'action': action_clean,
                    'entity_type': entity_type,
                    'entity_label': _ENTITY_TYPE_UA.get(entity_type, entity_type),
                    'entity_id': inferred.get('entity_id') or '',
                    'entity_name': '',
                    'entity_url': inferred.get('entity_url') or '',
                    'group_names': group_names,
                    'mode': mode,
                    'changes': changes,
                    'meta_items': [],
                    'details_text': details_text,
                    'route_path': route_path,
                    'http_method': '',
                    'ip_address': ip_address,
                    'user_agent': '',
                    'source': 'file',
                    'raw': line_clean,
                }
                entry['search_blob'] = _build_search_blob(entry)
                entries.append(entry)
    except OSError as e:
        logger.error(f'log read error: {e}')
    return entries


def _list_available_log_months():
    month_pattern = re.compile(r'^app-(\d{4}-\d{2})\.log$')
    months_map = {}
    if os.path.isdir(LOGS_DIR):
        for fname in os.listdir(LOGS_DIR):
            m = month_pattern.match(fname)
            if m:
                fpath = os.path.join(LOGS_DIR, fname)
                months_map[m.group(1)] = {
                    'value': m.group(1),
                    'size_kb': round(os.path.getsize(fpath) / 1024, 1),
                }
    try:
        conn = get_db()
        try:
            ensure_audit_logs_table(conn)
            for r in conn.execute(
                "SELECT DISTINCT substr(created_at, 1, 7) AS ym FROM audit_logs WHERE length(created_at) >= 7"
            ).fetchall():
                ym = r['ym']
                if ym and re.match(r'^\d{4}-\d{2}$', ym) and ym not in months_map:
                    months_map[ym] = {'value': ym, 'size_kb': 0.0}
        finally:
            conn.close()
    except Exception:
        pass
    available = list(months_map.values())
    available.sort(key=lambda x: x['value'], reverse=True)
    return available


def _collect_log_entries(selected_month, available_months):
    db_entries = []
    known_uids = set()
    try:
        conn = get_db()
        conn.row_factory = sqlite3.Row
        try:
            ensure_audit_logs_table(conn)
            if selected_month == 'all':
                rows = conn.execute(
                    "SELECT * FROM audit_logs ORDER BY created_at DESC, id DESC LIMIT 15000"
                ).fetchall()
                uid_rows = conn.execute(
                    "SELECT log_uid FROM audit_logs WHERE log_uid IS NOT NULL"
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM audit_logs WHERE created_at LIKE ? ORDER BY created_at DESC, id DESC",
                    (f"{selected_month}-%",),
                ).fetchall()
                uid_rows = conn.execute(
                    "SELECT log_uid FROM audit_logs WHERE created_at LIKE ? AND log_uid IS NOT NULL",
                    (f"{selected_month}-%",),
                ).fetchall()
            for r in rows:
                db_entries.append(_audit_row_to_entry(r))
            for ur in uid_rows:
                if ur['log_uid']:
                    known_uids.add(str(ur['log_uid']).lower())
        finally:
            conn.close()
    except Exception as e:
        logger.error(f"audit_logs read error: {e}")

    file_entries = []
    if selected_month == 'all':
        for m_info in available_months[:12]:
            fpath = os.path.join(LOGS_DIR, f"app-{m_info['value']}.log")
            file_entries.extend(_parse_log_file(fpath, skip_uids=known_uids))
    else:
        fpath = os.path.join(LOGS_DIR, f'app-{selected_month}.log')
        file_entries = _parse_log_file(fpath, skip_uids=known_uids)

    combined = db_entries + file_entries
    combined.sort(
        key=lambda e: (e.get('date', ''), e.get('time', ''), e.get('id') or 0),
        reverse=True,
    )
    return combined


def _collect_recent_errors(available_months):
    errors = []
    seen_uids = set()
    try:
        conn = get_db()
        conn.row_factory = sqlite3.Row
        try:
            ensure_audit_logs_table(conn)
            for r in conn.execute(
                "SELECT * FROM audit_logs WHERE level = 'ERROR' ORDER BY created_at DESC, id DESC LIMIT 25"
            ).fetchall():
                entry = _audit_row_to_entry(r)
                entry['month'] = (entry.get('date') or '')[:7]
                errors.append(entry)
                if entry.get('log_uid'):
                    seen_uids.add(entry['log_uid'].lower())
        finally:
            conn.close()
    except Exception:
        pass

    for month_info in available_months[:6]:
        m_val = month_info['value']
        fpath = os.path.join(LOGS_DIR, f"app-{m_val}.log")
        for entry in _parse_log_file(fpath, skip_uids=seen_uids):
            if entry.get('level') == 'ERROR':
                errors.append({**entry, 'month': m_val})

    errors.sort(key=lambda e: (e.get('date', ''), e.get('time', ''), e.get('id') or 0), reverse=True)
    return errors[:25]


def _filter_log_entries(entries, args):
    q = (args.get('q') or '').strip().lower()
    user_f = (args.get('user') or '').strip()
    cat_f = (args.get('category') or '').strip()
    sec_f = (args.get('section') or '').strip()
    level_f = (args.get('level') or '').strip().upper()
    date_from = (args.get('date_from') or '').strip()
    date_to = (args.get('date_to') or '').strip()
    only_changes = args.get('only_changes') in ('1', 'true', 'on')
    hide_views = args.get('hide_views') in ('1', 'true', 'on')

    if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_from):
        date_from = ''
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_to):
        date_to = ''

    q_tokens = [t for t in q.split() if t] if q else []

    filtered = []
    for e in entries:
        if user_f and e.get('username') != user_f:
            continue
        if cat_f and e.get('category') != cat_f:
            continue
        if sec_f and e.get('section') != sec_f:
            continue
        if level_f and e.get('level') != level_f:
            continue
        if date_from and (not e.get('date') or e['date'] < date_from):
            continue
        if date_to and (not e.get('date') or e['date'] > date_to):
            continue
        if only_changes and not e.get('changes'):
            continue
        if hide_views and e.get('category') == 'view':
            continue
        if q_tokens:
            blob = e.get('search_blob') or ''
            if not all(tok in blob for tok in q_tokens):
                continue
        filtered.append(e)

    filters = {
        'q': args.get('q', '').strip(),
        'user': user_f,
        'category': cat_f,
        'section': sec_f,
        'level': level_f,
        'date_from': date_from,
        'date_to': date_to,
        'only_changes': only_changes,
        'hide_views': hide_views,
    }
    return filtered, filters


@admin_bp.route('/admin/view_logs')
@permission_required('view_logs')
def view_logs():
    """
    Детальний журнал дій користувачів: аудит змін «Було → Стало», фільтри, експорт CSV та перегляд .log.
    """
    available_months = _list_available_log_months()
    current_month = datetime.now().strftime('%Y-%m')
    raw_month = (request.args.get('month') or current_month).strip()
    if raw_month == 'all':
        selected_month = 'all'
    elif re.match(r'^\d{4}-\d{2}$', raw_month):
        selected_month = raw_month
    else:
        selected_month = current_month

    all_entries = _collect_log_entries(selected_month, available_months)
    filtered_entries, filters = _filter_log_entries(all_entries, request.args)

    # Опції фільтрів будуємо з усіх подій періоду
    usernames = sort_ukrainian(list({e['username'] for e in all_entries if e.get('username')}), lambda x: x)
    sections = sort_ukrainian(list({e['section'] for e in all_entries if e.get('section')}), lambda x: x)
    cat_codes_present = {e['category'] for e in all_entries if e.get('category')}
    categories = [
        {'code': code, 'label': label}
        for code, label in AUDIT_CATEGORY_UA.items()
        if code in cat_codes_present
    ]

    today_str = date.today().strftime('%Y-%m-%d')
    week_ago_str = (date.today() - timedelta(days=6)).strftime('%Y-%m-%d')

    summary = {
        'total_all': len(all_entries),
        'total_filtered': len(filtered_entries),
        'today': sum(1 for e in filtered_entries if e.get('date') == today_str),
        'unique_users': len({e['username'] for e in filtered_entries if e.get('username')}),
        'with_changes': sum(1 for e in filtered_entries if e.get('changes')),
        'security_and_deletes': sum(
            1 for e in filtered_entries
            if e.get('category') in ('delete', 'security') or e.get('level') == 'ERROR'
        ),
        'errors_warnings': sum(1 for e in filtered_entries if e.get('level') in ('ERROR', 'WARNING')),
    }

    try:
        per_page = int(request.args.get('per_page', 50))
    except (TypeError, ValueError):
        per_page = 50
    if per_page not in (50, 100, 200):
        per_page = 50

    total_filtered = len(filtered_entries)
    total_pages = max(1, (total_filtered + per_page - 1) // per_page)
    try:
        page = int(request.args.get('page', 1))
    except (TypeError, ValueError):
        page = 1
    page = max(1, min(page, total_pages))

    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page
    paged_logs = filtered_entries[start_idx:end_idx]

    recent_errors = _collect_recent_errors(available_months)
    has_log_file = (
        selected_month != 'all'
        and os.path.exists(os.path.join(LOGS_DIR, f'app-{selected_month}.log'))
    )

    # Логуємо лише первинний перегляд журналу (без фільтрації/пагінації), щоб не засмічувати видачу
    is_pure_visit = (
        page == 1
        and not any(
            filters[k]
            for k in ('q', 'user', 'category', 'section', 'level', 'date_from', 'date_to', 'only_changes', 'hide_views')
        )
    )
    if is_pure_visit:
        log_action(
            current_username(),
            f'переглянув журнал дій ({selected_month})',
            category='view',
            action_type='view',
            section='Журнал дій',
            meta={'month': selected_month},
        )

    return render_template(
        'view_logs.html',
        logs=paged_logs,
        total_filtered=total_filtered,
        total_all=len(all_entries),
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        start_num=(start_idx + 1) if total_filtered else 0,
        end_num=min(end_idx, total_filtered),
        usernames=usernames,
        sections=sections,
        categories=categories,
        filters=filters,
        summary=summary,
        now=today_str,
        week_ago=week_ago_str,
        available_months=available_months,
        selected_month=selected_month,
        current_month=current_month,
        has_log_file=has_log_file,
        recent_errors=recent_errors,
    )


@admin_bp.route('/admin/export_logs')
@permission_required('view_logs')
def export_logs():
    """Експорт відфільтрованого журналу дій у CSV (UTF-8 BOM)."""
    available_months = _list_available_log_months()
    current_month = datetime.now().strftime('%Y-%m')
    raw_month = (request.args.get('month') or current_month).strip()
    if raw_month == 'all':
        selected_month = 'all'
    elif re.match(r'^\d{4}-\d{2}$', raw_month):
        selected_month = raw_month
    else:
        selected_month = current_month

    all_entries = _collect_log_entries(selected_month, available_months)
    filtered_entries, _ = _filter_log_entries(all_entries, request.args)

    buf = io.StringIO()
    buf.write('\ufeff')
    writer = csv.writer(buf, delimiter=';')
    writer.writerow([
        'Дата',
        'Час',
        'Рівень',
        'Користувач',
        'Роль',
        'Категорія',
        'Розділ',
        'Дія',
        "Об'єкт",
        'Групи',
        'Що на що замінено',
        'Деталі',
        'Маршрут',
        'Метод',
        'IP-адреса',
        'Браузер / ОС',
        'Джерело',
    ])

    for e in filtered_entries:
        entity_str = ''
        if e.get('entity_label') or e.get('entity_name') or e.get('entity_id'):
            parts = [p for p in (e.get('entity_label'), e.get('entity_name')) if p]
            entity_str = ': '.join(parts)
            if e.get('entity_id'):
                entity_str = f"{entity_str} (#{e['entity_id']})" if entity_str else f"#{e['entity_id']}"

        changes_str = '; '.join(
            f"{c['label']}: «{c['old']}» -> «{c['new']}»" for c in (e.get('changes') or [])
        )
        details_parts = []
        if e.get('mode'):
            details_parts.append(f"Режим: {e['mode']}")
        for mi in (e.get('meta_items') or []):
            details_parts.append(f"{mi['label']}: {mi['value']}")
        if e.get('details_text'):
            details_parts.append(e['details_text'])
        details_joined = '; '.join(details_parts)

        row_vals = [
            e.get('date_ua') or e.get('date') or '',
            e.get('time') or '',
            e.get('level') or '',
            e.get('username') or '',
            e.get('user_role_label') or '',
            e.get('category_label') or '',
            e.get('section') or '',
            e.get('action') or '',
            entity_str,
            e.get('group_names') or '',
            changes_str,
            details_joined,
            e.get('route_path') or '',
            e.get('http_method') or '',
            e.get('ip_address') or '',
            e.get('user_agent') or '',
            'БД' if e.get('source') == 'db' else 'Файл .log',
        ]
        writer.writerow([sanitize_spreadsheet_cell(v, for_csv=True) for v in row_vals])

    log_action(
        current_username(),
        f'експортував журнал дій у CSV ({selected_month})',
        category='export',
        action_type='export',
        section='Журнал дій',
        meta={'records': len(filtered_entries)},
    )

    fname = f"audit-logs-{selected_month}-{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    return Response(
        buf.getvalue(),
        mimetype='text/csv; charset=utf-8',
        headers={'Content-Disposition': f'attachment; filename="{fname}"'},
    )


@admin_bp.route('/admin/download_log')
@permission_required('view_logs')
def download_log():
    """Завантажити .log за обраний місяць."""
    month = request.args.get('month', '')
    if not re.match(r'^\d{4}-\d{2}$', month):
        flash('Невірний формат місяця.', 'error')
        return redirect(url_for('admin.view_logs'))
    fpath = os.path.join(LOGS_DIR, f'app-{month}.log')
    if not os.path.exists(fpath):
        flash('Файл логів за цей місяць не знайдено.', 'error')
        return redirect(url_for('admin.view_logs'))
    log_action(
        current_username(),
        f'завантажив файл логів за {month}',
        category='export',
        action_type='export',
        section='Журнал дій',
        entity_type='log_file',
        entity_name=f'app-{month}.log',
    )
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
    knowledge_areas = conn.execute("""
        SELECT code, name_ua FROM knowledge_fields
        ORDER BY code
    """).fetchall()
    return groups, specialties, degrees, knowledge_areas


@admin_bp.route('/admin/users/add', methods=['GET', 'POST'])
@permission_required('manage_users')
def add_user():
    conn = get_db()
    conn.row_factory = sqlite3.Row
    _ensure_user_scope_tables(conn)

    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        role = (request.form.get('role') or 'user').strip()
        if role not in ALLOWED_USER_ROLES:
            flash('Недопустима роль користувача', 'error')
            conn.close()
            return redirect(url_for('admin.add_user'))
        group_ids = [g for g in request.form.getlist('group_id') if g]
        specialty_codes = [s for s in request.form.getlist('specialty_code') if s]
        degree_levels = [d for d in request.form.getlist('degree_level') if d]
        knowledge_codes = [k for k in request.form.getlist('knowledge_field_code') if k]
        perms = [p for p in request.form.getlist('permissions') if p in PERMISSIONS]

        if not username or not password:
            flash("Логін і пароль обов'язкові", 'error')
            conn.close()
            return redirect(url_for('admin.add_user'))
        ok_pwd, pwd_err = validate_password_strength(password)
        if not ok_pwd:
            flash(pwd_err, 'error')
            conn.close()
            return redirect(url_for('admin.add_user'))

        uname_fold = username.casefold()
        existing_usernames = conn.execute("SELECT username FROM users").fetchall()
        if any((r['username'] or '').strip().casefold() == uname_fold for r in existing_usernames):
            flash('Такий логін уже існує', 'error')
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
            for kc in knowledge_codes:
                conn.execute(
                    "INSERT INTO user_knowledge_areas (user_id, knowledge_field_code) VALUES (?, ?)",
                    (uid, kc),
                )
            conn.commit()
            add_changes = [
                {'label': 'Роль', 'old': '—', 'new': _ROLE_UA.get(role, role)},
            ]
            if is_admin:
                add_changes.append({'label': 'Права', 'old': '—', 'new': 'Усі права (адміністратор)'})
            elif perms:
                add_changes.append({
                    'label': f'Права ({len(perms)})',
                    'old': '—',
                    'new': ', '.join(PERMISSION_LABELS.get(p, p) for p in perms),
                })
            if group_ids:
                g_names = [
                    r['name'] for r in conn.execute(
                        f"SELECT name FROM groups WHERE id IN ({','.join('?' for _ in group_ids)}) ORDER BY name",
                        [int(g) for g in group_ids],
                    ).fetchall()
                ]
                add_changes.append({'label': f'Групи ({len(g_names)})', 'old': '—', 'new': ', '.join(g_names)})
            if specialty_codes:
                add_changes.append({'label': 'Спеціальності', 'old': '—', 'new': ', '.join(specialty_codes)})
            if degree_levels:
                add_changes.append({'label': 'Ступені', 'old': '—', 'new': ', '.join(degree_levels)})
            if knowledge_codes:
                add_changes.append({'label': 'Галузі знань', 'old': '—', 'new': ', '.join(knowledge_codes)})

            log_action(
                current_username(),
                f'створив користувача {username}',
                category='permissions',
                action_type='create',
                section='Користувачі та права',
                entity_type='user',
                entity_id=uid,
                entity_name=username,
                entity_url=url_for('admin.edit_user', user_id=uid),
                changes=add_changes,
                meta={'role': role, 'perms_count': len(perms), 'groups_count': len(group_ids)},
                conn=conn,
            )
            conn.commit()
            flash(f'Користувача «{username}» створено', 'success')
            conn.close()
            return redirect(url_for('admin.manage_users'))
        except sqlite3.IntegrityError:
            conn.close()
            flash('Такий логін уже існує', 'error')
            return redirect(url_for('admin.add_user'))

    groups, specialties, degrees, knowledge_areas = _load_scope_options(conn)
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
        try:
            old_perms = json.loads(user['permissions'] or '[]')
        except Exception:
            old_perms = []
        old_gids = [
            int(r['group_id'])
            for r in conn.execute("SELECT group_id FROM user_groups WHERE user_id=?", (user_id,)).fetchall()
        ]
        old_specs = [
            r['specialty_code']
            for r in conn.execute("SELECT specialty_code FROM user_specialties WHERE user_id=?", (user_id,)).fetchall()
        ]
        old_degs = [
            r['degree_level']
            for r in conn.execute("SELECT degree_level FROM user_degree_levels WHERE user_id=?", (user_id,)).fetchall()
        ]
        old_knows = [
            r['knowledge_field_code']
            for r in conn.execute("SELECT knowledge_field_code FROM user_knowledge_areas WHERE user_id=?", (user_id,)).fetchall()
        ]

        role = (request.form.get('role') or 'user').strip()
        if role not in ALLOWED_USER_ROLES:
            flash('Недопустима роль користувача', 'error')
            conn.close()
            return redirect(url_for('admin.edit_user', user_id=user_id))
        group_ids = [g for g in request.form.getlist('group_id') if g]
        specialty_codes = [s for s in request.form.getlist('specialty_code') if s]
        degree_levels = [d for d in request.form.getlist('degree_level') if d]
        knowledge_codes = [k for k in request.form.getlist('knowledge_field_code') if k]
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
        conn.execute("DELETE FROM user_knowledge_areas WHERE user_id=?", (user_id,))
        for kc in knowledge_codes:
            conn.execute(
                "INSERT INTO user_knowledge_areas (user_id, knowledge_field_code) VALUES (?, ?)",
                (user_id, kc),
            )

        edit_changes = []
        if user['role'] != role:
            edit_changes.append({
                'label': 'Роль',
                'old': _ROLE_UA.get(user['role'], user['role']),
                'new': _ROLE_UA.get(role, role),
            })

        old_p_set, new_p_set = set(old_perms), set(perms)
        added_p = [PERMISSION_LABELS.get(p, p) for p in PERMISSIONS if p in (new_p_set - old_p_set)]
        removed_p = [PERMISSION_LABELS.get(p, p) for p in PERMISSIONS if p in (old_p_set - new_p_set)]
        if added_p or removed_p:
            edit_changes.append({
                'label': 'Права доступу',
                'old': ('Видалено: ' + ', '.join(removed_p)) if removed_p else '—',
                'new': ('Додано: ' + ', '.join(added_p)) if added_p else '—',
            })

        new_gids = [int(g) for g in group_ids]
        old_g_set, new_g_set = set(old_gids), set(new_gids)
        if old_g_set != new_g_set:
            all_gids = sorted(old_g_set | new_g_set)
            gmap = {
                r['id']: r['name']
                for r in conn.execute(
                    f"SELECT id, name FROM groups WHERE id IN ({','.join('?' for _ in all_gids)})",
                    all_gids,
                ).fetchall()
            } if all_gids else {}
            added_g = [gmap.get(g, f"#{g}") for g in sorted(new_g_set - old_g_set)]
            removed_g = [gmap.get(g, f"#{g}") for g in sorted(old_g_set - new_g_set)]
            edit_changes.append({
                'label': 'Доступ до груп',
                'old': ('Забрано: ' + ', '.join(removed_g)) if removed_g else '—',
                'new': ('Надано: ' + ', '.join(added_g)) if added_g else '—',
            })

        for label_txt, old_seq, new_seq in (
            ('Спеціальності', old_specs, specialty_codes),
            ('Ступені', old_degs, degree_levels),
            ('Галузі знань', old_knows, knowledge_codes),
        ):
            if set(old_seq) != set(new_seq):
                edit_changes.append({
                    'label': label_txt,
                    'old': ', '.join(sorted(old_seq)) if old_seq else '—',
                    'new': ', '.join(sorted(new_seq)) if new_seq else '—',
                })

        log_action(
            current_username(),
            f"редагував користувача {user['username']} (ID {user_id})",
            category='permissions',
            action_type='update',
            section='Користувачі та права',
            entity_type='user',
            entity_id=user_id,
            entity_name=user['username'],
            entity_url=url_for('admin.edit_user', user_id=user_id),
            changes=edit_changes,
            meta={'role': role, 'perms_count': len(perms), 'groups_count': len(new_gids)},
            conn=conn,
        )
        conn.commit()
        flash('Дані користувача оновлено', 'success')
        conn.close()
        return redirect(url_for('admin.manage_users'))

    try:
        current_perms = json.loads(user['permissions'] or '[]')
    except Exception:
        current_perms = []

    current_group_ids = [
        int(r['group_id'])
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
    current_knowledge = [
        r['knowledge_field_code']
        for r in conn.execute(
            "SELECT knowledge_field_code FROM user_knowledge_areas WHERE user_id=?", (user_id,)
        ).fetchall()
    ]

    groups, specialties, degrees, knowledge_areas = _load_scope_options(conn)
    conn.close()
    return render_template(
        'edit_user.html',
        user=user,
        groups=groups,
        specialties=specialties,
        degrees=degrees,
        knowledge_areas=knowledge_areas,
        permission_groups=PERMISSION_GROUPS,
        permissions_by_group=_permissions_by_group(),
        selected_permissions=current_perms,
        selected_group_ids=current_group_ids,
        selected_specialties=current_specialties,
        selected_degrees=current_degrees,
        selected_knowledge=current_knowledge,
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
        ok_pwd, pwd_err = validate_password_strength(password)
        if not ok_pwd:
            flash(pwd_err, 'error')
            conn.close()
            return redirect(url_for('admin.change_password', user_id=user_id))
        conn.execute(
            "UPDATE users SET password_hash=? WHERE id=?",
            (generate_password_hash(password), user_id),
        )
        log_action(
            current_username(),
            f"змінив пароль користувача {user['username']}",
            category='permissions',
            action_type='update',
            section='Користувачі та права',
            entity_type='user',
            entity_id=user_id,
            entity_name=user['username'],
            entity_url=url_for('admin.edit_user', user_id=user_id),
            changes=[{'label': 'Пароль', 'old': 'поточний пароль', 'new': 'встановлено новий пароль'}],
            conn=conn,
        )
        conn.commit()
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
        "SELECT username, role FROM users WHERE id=?", (user_id,)
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
    conn.execute("DELETE FROM user_knowledge_areas WHERE user_id=?", (user_id,))
    conn.execute("DELETE FROM users WHERE id=?", (user_id,))
    log_action(
        current_username(),
        f"видалив користувача {user['username']}",
        category='delete',
        action_type='delete',
        section='Користувачі та права',
        entity_type='user',
        entity_id=user_id,
        entity_name=user['username'],
        changes=[{'label': 'Обліковий запис', 'old': f"{user['username']} ({_ROLE_UA.get(user['role'], user['role'])})", 'new': 'Видалено'}],
        conn=conn,
    )
    conn.commit()
    conn.close()
    flash('Користувача видалено', 'success')
    return redirect(url_for('admin.manage_users'))
