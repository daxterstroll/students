"""
routes/exam_sessions/history.py
================================
История изменений по сессиям и ведомостям.

Содержит:
  - log_history()                          — запись события (вызывается из admin/public/transfer)
  - session_history(session_id)            — история ОДНОЙ живой сессии
  - sheet_history(sheet_id)                — история ОДНОЙ живой ведомости
  - deleted_session_history(session_id)    — история УДАЛЁННОЙ сессии (из snapshot в details)
  - history_index()                        — ОБЩИЙ журнал всех изменений с фильтрами

Право на просмотр — view_exam_session_history.
Админ (is_admin=True) — может всё.
"""
from __future__ import annotations

import json
import sqlite3

from flask import render_template, request, redirect, url_for, flash, session

from routes.db import get_db
from routes.utils import permission_required
from routes.helpers import current_username
from routes.exam_sessions import exam_sessions_bp
from routes.exam_sessions.schema import ensure_exam_session_tables
from routes.exam_sessions.service import (
    user_can_access_group, allowed_groups_for_user,
)


ACTION_LABELS = {
    'create':       'Створено',
    'update':       'Змінено',
    'delete':       'Видалено',
    'token_issued': 'Видано токен',
    'approve':      'Затверджено',
    'unlink':       'Відв\'язано від плану',
    'lock':         'Заблоковано',
    'grade_set':    'Оцінка змінена',
    'transfer':     'Перенесено в план',
    'refresh':      'Оновлено список студентів',
}

ENTITY_LABELS = {
    'session':       'Сесія',
    'sheet':         'Відомість',
    'sheet_student': 'Студент',
}


def log_history(
    conn,
    *,
    entity_type: str,
    entity_id: int,
    action: str,
    session_id=None,
    sheet_id=None,
    field_name=None,
    old_value=None,
    new_value=None,
    details=None,
):
    """Записывает событие. НЕ делает commit — вызывается в той же транзакции."""
    try:
        conn.execute("""
            INSERT INTO exam_session_history
                (entity_type, entity_id, session_id, sheet_id,
                 action, field_name, old_value, new_value, details, changed_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            entity_type, entity_id, session_id, sheet_id,
            action, field_name,
            None if old_value is None else str(old_value),
            None if new_value is None else str(new_value),
            json.dumps(details, ensure_ascii=False) if details else None,
            current_username(),
        ))
    except Exception as e:
        from routes.utils import logger
        logger.error(f'[history] не удалось записать: {e}', exc_info=True)


# ═══════════════════════════════════════════════════════════════════════
# История одной живой сессии
# ═══════════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/admin/exam-sessions/<int:session_id>/history')
@permission_required('view_exam_session_history')
def session_history(session_id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)

    es = conn.execute("""
        SELECT es.*, g.name AS group_name, g.course, g.study_form
        FROM exam_sessions es
        JOIN groups g ON g.id = es.group_id
        WHERE es.id = ?
    """, (session_id,)).fetchone()

    if not es:
        # Сессии нет — а может, она удалена? Проверим историю
        has_history = conn.execute(
            "SELECT 1 FROM exam_session_history WHERE session_id = ? LIMIT 1",
            (session_id,),
        ).fetchone()
        conn.close()
        if has_history:
            # Перебрасываем на специальную страницу для удалённых
            return redirect(url_for('exam_sessions.deleted_session_history',
                                    session_id=session_id))
        flash('Сесію не знайдено', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))

    if not user_can_access_group(conn, session, es['group_id']):
        conn.close()
        flash('Немає доступу до цієї сесії', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))

    rows = conn.execute("""
        SELECT * FROM exam_session_history
        WHERE session_id = ?
        ORDER BY id DESC
        LIMIT 1000
    """, (session_id,)).fetchall()

    conn.close()
    return render_template(
        'exam_sessions/history.html',
        mode='session',
        entity=dict(es),
        rows=[dict(r) for r in rows],
        action_labels=ACTION_LABELS,
        entity_labels=ENTITY_LABELS,
    )


# ═══════════════════════════════════════════════════════════════════════
# История одной живой ведомости
# ═══════════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/admin/exam-sessions/sheets/<int:sheet_id>/history')
@permission_required('view_exam_session_history')
def sheet_history(sheet_id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)

    sheet = conn.execute("""
        SELECT ss.*, es.group_id, es.name AS session_name,
               g.name AS group_name
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        WHERE ss.id = ?
    """, (sheet_id,)).fetchone()

    if not sheet or not user_can_access_group(conn, session, sheet['group_id']):
        conn.close()
        flash('Відомість не знайдено або немає доступу', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))

    rows = conn.execute("""
        SELECT * FROM exam_session_history
        WHERE sheet_id = ?
        ORDER BY id DESC
        LIMIT 1000
    """, (sheet_id,)).fetchall()

    conn.close()
    return render_template(
        'exam_sessions/history.html',
        mode='sheet',
        entity=dict(sheet),
        rows=[dict(r) for r in rows],
        action_labels=ACTION_LABELS,
        entity_labels=ENTITY_LABELS,
    )


# ═══════════════════════════════════════════════════════════════════════
# История УДАЛЁННОЙ сессии (по snapshot из details записи action='delete')
# ═══════════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/admin/exam-sessions/history/deleted/<int:session_id>')
@permission_required('view_exam_session_history')
def deleted_session_history(session_id):
    """
    Показывает историю удалённой сессии.

    Сессии в exam_sessions уже нет (удалена), но exam_session_history
    её помнит — потому что мы НЕ удаляем записи истории (см. delete_session
    в admin.py — там ON DELETE CASCADE намеренно не ставится).

    Данные о сессии берём из details записи action='delete':
      { name, group_id, group_name, academic_year, semester,
        sheets_count, sheets:[{id, discipline_name}, ...] }
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)

    # Берём последнюю запись delete (или первую delete — она одна)
    delete_row = conn.execute("""
        SELECT * FROM exam_session_history
        WHERE session_id = ? AND entity_type = 'session' AND action = 'delete'
        ORDER BY id DESC LIMIT 1
    """, (session_id,)).fetchone()

    if not delete_row:
        conn.close()
        flash('Запис про видалення не знайдено', 'warning')
        return redirect(url_for('exam_sessions.history_index'))

    # Snapshot из details
    try:
        snap = json.loads(delete_row['details'] or '{}')
    except Exception:
        snap = {}

    group_id = snap.get('group_id')

    # Проверка доступа: если группа ещё существует — по её allowed_groups_for_user;
    # если группы уже нет — разрешаем смотреть только админам
    if group_id:
        if not user_can_access_group(conn, session, group_id):
            conn.close()
            flash('Немає доступу до цієї сесії', 'error')
            return redirect(url_for('exam_sessions.history_index'))
    else:
        if not session.get('is_admin'):
            conn.close()
            flash('Немає доступу до цієї сесії', 'error')
            return redirect(url_for('exam_sessions.history_index'))

    rows = conn.execute("""
        SELECT * FROM exam_session_history
        WHERE session_id = ?
        ORDER BY id DESC
        LIMIT 1000
    """, (session_id,)).fetchall()

    conn.close()

    entity = {
        'id': session_id,
        'name': snap.get('name') or f'Сесія #{session_id} (видалена)',
        'group_id': group_id,
        'group_name': snap.get('group_name') or '—',
        'academic_year': snap.get('academic_year') or '',
        'semester': snap.get('semester') or '',
        'is_deleted': True,
        'deleted_at': delete_row['changed_at'],
        'deleted_by': delete_row['changed_by'],
        'sheets_count': snap.get('sheets_count') or 0,
    }

    return render_template(
        'exam_sessions/history.html',
        mode='deleted_session',
        entity=entity,
        rows=[dict(r) for r in rows],
        action_labels=ACTION_LABELS,
        entity_labels=ENTITY_LABELS,
    )


# ═══════════════════════════════════════════════════════════════════════
# ОБЩАЯ история — журнал всех изменений с фильтрами
# ═══════════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/admin/exam-sessions/history')
@permission_required('view_exam_session_history')
def history_index():
    """
    Общий журнал: всё, что когда-либо происходило с сессиями и ведомостями.

    Фильтры (все опциональны):
      - entity_type: session | sheet | sheet_student
      - action:      create | update | delete | ...
      - user:        username
      - date_from, date_to: YYYY-MM-DD
      - search:      подстрока в field_name / old_value / new_value / details / changed_by
      - session_id:  конкретная сессия (в т.ч. удалённая)
      - sheet_id:    конкретная ведомость

    Показываем только те записи, к которым у пользователя есть доступ
    (не-админ видит только свои группы).
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)

    is_admin = session.get('is_admin')
    allowed_group_ids = {g['id'] for g in allowed_groups_for_user(conn, session)}

    # Параметры
    entity_type = (request.args.get('entity_type') or '').strip()
    action_filter = (request.args.get('action') or '').strip()
    user_filter = (request.args.get('user') or '').strip()
    date_from = (request.args.get('date_from') or '').strip()
    date_to = (request.args.get('date_to') or '').strip()
    search = (request.args.get('search') or '').strip()
    session_id_filter = request.args.get('session_id', type=int)
    sheet_id_filter = request.args.get('sheet_id', type=int)
    page = max(1, request.args.get('page', 1, type=int))
    per_page = request.args.get('per_page', 100, type=int)
    if per_page not in (50, 100, 200, 500):
        per_page = 100

    where = ['1=1']
    params = []

    if entity_type in ('session', 'sheet', 'sheet_student'):
        where.append('entity_type = ?')
        params.append(entity_type)
    if action_filter in ACTION_LABELS:
        where.append('action = ?')
        params.append(action_filter)
    if user_filter:
        where.append('changed_by = ?')
        params.append(user_filter)
    if date_from:
        where.append("date(changed_at) >= date(?)")
        params.append(date_from)
    if date_to:
        where.append("date(changed_at) <= date(?)")
        params.append(date_to)
    if session_id_filter:
        where.append('session_id = ?')
        params.append(session_id_filter)
    if sheet_id_filter:
        where.append('sheet_id = ?')
        params.append(sheet_id_filter)
    if search:
        q = f'%{search}%'
        where.append("""
            (field_name LIKE ?
             OR old_value LIKE ?
             OR new_value LIKE ?
             OR details LIKE ?
             OR changed_by LIKE ?)
        """)
        params.extend([q, q, q, q, q])

    where_sql = ' AND '.join(where)

    # Считаем общее количество
    total = conn.execute(
        f"SELECT COUNT(*) FROM exam_session_history WHERE {where_sql}",
        params,
    ).fetchone()[0]

    total_pages = max(1, (total + per_page - 1) // per_page)
    if page > total_pages:
        page = total_pages
    offset = (page - 1) * per_page

    raw_rows = conn.execute(f"""
        SELECT * FROM exam_session_history
        WHERE {where_sql}
        ORDER BY id DESC
        LIMIT ? OFFSET ?
    """, params + [per_page, offset]).fetchall()

    # Достаём живые сессии одним запросом — чтобы отличить живую от удалённой
    session_ids = {r['session_id'] for r in raw_rows if r['session_id']}
    live_sessions = {}
    if session_ids:
        ph = ','.join('?' * len(session_ids))
        for row in conn.execute(f"""
            SELECT es.id, es.name, es.group_id,
                   g.name AS group_name, g.course, g.study_form
            FROM exam_sessions es
            LEFT JOIN groups g ON g.id = es.group_id
            WHERE es.id IN ({ph})
        """, list(session_ids)).fetchall():
            live_sessions[row['id']] = dict(row)

    # Фильтрация по доступу (не-админ видит только свои группы)
    rows = []
    for r in raw_rows:
        d = dict(r)
        sid = d.get('session_id')

        if not is_admin:
            if sid and sid in live_sessions:
                gid = live_sessions[sid].get('group_id')
                if gid and gid not in allowed_group_ids:
                    continue
            elif sid:
                # Удалённая сессия — проверяем по snapshot group_id
                try:
                    snap = json.loads(d.get('details') or '{}')
                    gid = snap.get('group_id')
                except Exception:
                    gid = None
                if gid and gid not in allowed_group_ids:
                    continue
                elif not gid:
                    # Совсем без группы — не показываем не-админам
                    continue
            else:
                # Нет session_id — совсем служебная запись
                continue

        # Информация для отображения в таблице
        d['session_is_deleted'] = sid and sid not in live_sessions
        d['session_name'] = (
            live_sessions[sid]['name'] if sid in live_sessions
            else _extract_name_from_snapshot(d.get('details'))
        )
        d['session_group'] = (
            live_sessions[sid].get('group_name') if sid in live_sessions
            else _extract_group_from_snapshot(d.get('details'))
        )
        rows.append(d)

    # Список пользователей для фильтра (все, кто оставлял записи)
    users = [r['changed_by'] for r in conn.execute(
        "SELECT DISTINCT changed_by FROM exam_session_history "
        "WHERE changed_by IS NOT NULL ORDER BY changed_by"
    ).fetchall()]

    conn.close()

    return render_template(
        'exam_sessions/history_index.html',
        rows=rows,
        total=total,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        users=users,
        action_labels=ACTION_LABELS,
        entity_labels=ENTITY_LABELS,
        # значения фильтров
        filter_entity_type=entity_type,
        filter_action=action_filter,
        filter_user=user_filter,
        filter_date_from=date_from,
        filter_date_to=date_to,
        filter_search=search,
        filter_session_id=session_id_filter,
        filter_sheet_id=sheet_id_filter,
    )


def _extract_name_from_snapshot(details_json):
    """Достаёт имя сессии из details snapshot (для удалённых)."""
    try:
        snap = json.loads(details_json or '{}')
        return snap.get('name') or snap.get('session_name') or '—'
    except Exception:
        return '—'


def _extract_group_from_snapshot(details_json):
    """Достаёт название группы из details snapshot (для удалённых)."""
    try:
        snap = json.loads(details_json or '{}')
        return snap.get('group_name') or '—'
    except Exception:
        return '—'