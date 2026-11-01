"""
routes/admin/enrollment_orders.py
=================================
Крок 3: накази про зарахування (enrollment).

Маршрути на спільному admin_bp:
  GET, POST /admin/orders/enrollment          — bulk: вибір групи / студентів
  POST      /admin/orders/enrollment/confirm  — створення наказу

Право: manage_orders (або fallback manage_students / admin).

Імпорт у routes/admin/__init__.py:
    from . import enrollment_orders  # noqa: F401
"""

from __future__ import annotations

import sqlite3

from flask import render_template, request, redirect, url_for, flash, session

from routes.admin import admin_bp
from routes.db import get_db
from routes.helpers import current_username, sort_ukrainian
from routes.utils import log_action, permission_required, logger
from routes.orders_service import (
    create_order,
    get_enrollment_student_ids,
    orders_tables_ready,
    ORDER_TYPE_LABELS,
)


def _can_manage_orders() -> bool:
    """Адмін або явний дозвіл manage_orders / manage_students."""
    if session.get('is_admin'):
        return True
    perms = set(session.get('permissions') or [])
    return bool(perms & {'manage_orders', 'manage_students', 'manage_expulsion'})


@admin_bp.route('/admin/orders/enrollment', methods=['GET', 'POST'])
@permission_required('manage_students')  # тимчасово; пізніше manage_orders
def enrollment_order():
    """
    Крок 1–2 майстра: вибір групи (або всіх без наказу) і відмітка студентів.
    POST з action=preview — показує форму з номером/датою наказу.
    """
    if not _can_manage_orders():
        flash('Немає доступу до наказів.', 'error')
        return redirect(url_for('students.student_list'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    if not orders_tables_ready(conn=conn):
        conn.close()
        flash(
            'Таблиці наказів ще не створені. Запустіть python migrate_orders.py',
            'error',
        )
        return redirect(url_for('students.student_list'))

    groups = conn.execute(
        """
        SELECT id, name, start_year, study_form, program_credits
        FROM groups
        WHERE COALESCE(archived, 0) = 0
        ORDER BY start_year DESC, name COLLATE UKRAINIAN
        """
    ).fetchall()
    groups = sort_ukrainian([dict(g) for g in groups], lambda g: g['name'])

    enrolled_ids = get_enrollment_student_ids(conn=conn)

    # ── POST: зібрати обраних студентів і показати confirm ──
    if request.method == 'POST':
        action = request.form.get('action') or 'preview'
        group_id = request.form.get('group_id', type=int)
        only_missing = request.form.get('only_missing') == '1'
        selected = [int(x) for x in request.form.getlist('student_ids') if x.isdigit()]

        if action == 'confirm':
            return _do_confirm(conn, selected)

        # preview
        students = _load_candidates(conn, group_id, only_missing, enrolled_ids)
        # якщо прийшли з галочками — залишаємо лише вибраних; інакше всі кандидати
        if selected:
            sel_set = set(selected)
            students = [s for s in students if s['id'] in sel_set]

        if not students:
            flash('Не обрано жодного студента для наказу про зарахування.', 'warning')
            conn.close()
            return redirect(url_for('admin.enrollment_order'))

        conn.close()
        return render_template(
            'enrollment_order_confirm.html',
            students=students,
            group_id=group_id,
            only_missing=only_missing,
            order_type_label=ORDER_TYPE_LABELS['enrollment'],
        )

    # ── GET: форма вибору ──
    group_id = request.args.get('group_id', type=int)
    only_missing = request.args.get('only_missing', '1') == '1'
    students = _load_candidates(conn, group_id, only_missing, enrolled_ids)

    missing_total = conn.execute(
        """
        SELECT COUNT(*) FROM students s
        WHERE COALESCE(s.archived, 0) = 0
        """
    ).fetchone()[0] - len(enrolled_ids)
    # точніше: лише активні без enrollment
    missing_total = max(0, sum(
        1 for r in conn.execute(
            "SELECT id FROM students WHERE COALESCE(archived, 0)=0"
        ).fetchall()
        if r[0] not in enrolled_ids
    ))

    conn.close()
    return render_template(
        'enrollment_order.html',
        groups=groups,
        students=students,
        group_id=group_id,
        only_missing=only_missing,
        enrolled_count=len(enrolled_ids),
        missing_total=missing_total,
        order_type_label=ORDER_TYPE_LABELS['enrollment'],
    )


def _load_candidates(conn, group_id, only_missing, enrolled_ids):
    sql = """
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, g.id AS group_id, g.name AS group_name, g.start_year
        FROM students s
        LEFT JOIN groups g ON s.group_id = g.id
        WHERE COALESCE(s.archived, 0) = 0
    """
    params = []
    if group_id:
        sql += " AND s.group_id = ?"
        params.append(group_id)
    sql += " ORDER BY s.last_name_UA COLLATE UKRAINIAN, s.first_name_UA COLLATE UKRAINIAN"
    rows = [dict(r) for r in conn.execute(sql, params).fetchall()]
    for r in rows:
        r['has_enrollment'] = r['id'] in enrolled_ids
        r['full_name'] = ' '.join(
            x for x in (r['last_name_UA'], r['first_name_UA'], r['middle_name_UA'] or '') if x
        )
    if only_missing:
        rows = [r for r in rows if not r['has_enrollment']]
    return rows


def _do_confirm(conn, selected_ids):
    order_number = (request.form.get('order_number') or '').strip()
    order_date = (request.form.get('order_date') or '').strip()
    notes = (request.form.get('notes') or '').strip() or None

    if not order_number or not order_date:
        flash('Заповніть номер і дату наказу.', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    if not selected_ids:
        flash('Немає студентів у наказі.', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    # підтягнемо group_id для payload
    placeholders = ','.join('?' for _ in selected_ids)
    rows = conn.execute(
        f"""
        SELECT id, group_id FROM students
        WHERE id IN ({placeholders}) AND COALESCE(archived, 0)=0
        """,
        selected_ids,
    ).fetchall()
    found = {r['id']: r['group_id'] for r in rows}
    missing = [i for i in selected_ids if i not in found]
    if missing:
        flash(f'Частина студентів не знайдена або в архіві: {missing[:5]}…', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    # не дублюємо, якщо вже є enrollment
    already = get_enrollment_student_ids(conn=conn)
    to_add = [sid for sid in selected_ids if sid not in already]
    skipped = len(selected_ids) - len(to_add)

    if not to_add:
        flash('У всіх обраних студентів уже є наказ про зарахування.', 'warning')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    items = [
        {
            'entity_type': 'student',
            'entity_id': sid,
            'payload': {'group_id': found[sid]},
        }
        for sid in to_add
    ]

    scan_files = request.files.getlist('scan_files')

    try:
        order_id = create_order(
            'enrollment',
            order_number,
            order_date,
            items,
            created_by=current_username(),
            notes=notes,
            files=scan_files or None,
            upload_subdir='orders',
            conn=conn,
        )
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        logger.error(f'enrollment order failed: {e}', exc_info=True)
        flash(f'Помилка створення наказу: {e}', 'error')
        return redirect(url_for('admin.enrollment_order'))

    conn.close()

    log_action(
        current_username(),
        f'наказ про зарахування №{order_number} від {order_date} '
        f'(студентів: {len(to_add)}, order_id={order_id})',
    )
    msg = f'Наказ про зарахування №{order_number} оформлено. Студентів: {len(to_add)}.'
    if skipped:
        msg += f' Пропущено (уже мали наказ): {skipped}.'
    flash(msg, 'success')
    return redirect(url_for('admin.enrollment_order'))
