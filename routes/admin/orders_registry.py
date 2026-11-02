"""
routes/admin/orders_registry.py
===============================
Реєстр наказів: список, картка, редагування реквізитів і складу.

Маршрути:
  GET  /admin/orders
  GET  /admin/orders/<id>
  GET, POST /admin/orders/<id>/edit
  POST /admin/orders/<id>/items/<item_id>/remove
"""

from __future__ import annotations

import json
import sqlite3

from flask import render_template, request, redirect, url_for, flash, session

from routes.admin import admin_bp
from routes.db import get_db
from routes.helpers import current_username, sort_ukrainian
from routes.utils import permission_required, log_action, save_multiple_attachments
from routes.orders_service import (
    ORDER_TYPE_LABELS,
    ORDER_TYPES,
    get_order,
    orders_tables_ready,
    delete_order,
    move_order_item,
)


def _can_manage() -> bool:
    if session.get('is_admin'):
        return True
    perms = set(session.get('permissions') or [])
    return bool(perms & {'manage_orders', 'manage_students', 'manage_expulsion'})


CREATE_ORDER_LINKS = [
    {
        'endpoint': 'admin.enrollment_order',
        'label': 'Наказ про зарахування',
        'icon': 'bi-file-earmark-plus',
        'type': 'enrollment',
    },
    {
        'endpoint': 'admin.course_transfer',
        'label': 'Переведення на курс',
        'icon': 'bi-arrow-up-circle',
        'type': 'course_transfer',
    },
    {
        'endpoint': 'admin.expulsion',
        'label': 'Відрахування',
        'icon': 'bi-person-x',
        'type': 'expulsion',
    },
    {
        'endpoint': 'admin.graduation',
        'label': 'Випуск',
        'icon': 'bi-mortarboard',
        'type': 'graduation',
    },
]


@admin_bp.route('/admin/orders')
@permission_required('manage_students')
def orders_list():
    if not _can_manage():
        flash('Немає доступу до реєстру наказів.', 'error')
        return redirect(url_for('students.student_list'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    if not orders_tables_ready(conn=conn):
        conn.close()
        flash('Таблиці наказів не створені. Запустіть migrate_orders.py', 'error')
        return redirect(url_for('students.student_list'))

    order_type = (request.args.get('type') or '').strip() or None
    search = (request.args.get('search') or '').strip()
    page = max(1, request.args.get('page', 1, type=int))
    per_page = request.args.get('per_page', 50, type=int)
    if per_page not in (20, 50, 100, 200):
        per_page = 50

    where = ["o.status = 'active'"]
    params = []
    if order_type and order_type in ORDER_TYPES:
        where.append('o.order_type = ?')
        params.append(order_type)
    if search:
        where.append(
            "(o.order_number LIKE ? OR IFNULL(o.title, '') LIKE ? OR IFNULL(o.notes, '') LIKE ?)"
        )
        q = f'%{search}%'
        params.extend([q, q, q])

    where_sql = ' AND '.join(where)

    total = conn.execute(
        f'SELECT COUNT(*) FROM orders o WHERE {where_sql}', params
    ).fetchone()[0]

    total_pages = max(1, (total + per_page - 1) // per_page)
    if page > total_pages:
        page = total_pages
    offset = (page - 1) * per_page

    rows = conn.execute(
        f"""
        SELECT o.*,
               (SELECT COUNT(*) FROM order_items oi WHERE oi.order_id = o.id) AS items_count
        FROM orders o
        WHERE {where_sql}
        ORDER BY o.id DESC
        LIMIT ? OFFSET ?
        """,
        params + [per_page, offset],
    ).fetchall()

    type_counts = {t: 0 for t in ORDER_TYPES}
    for r in conn.execute(
        """
        SELECT order_type, COUNT(*) AS c
        FROM orders WHERE status = 'active'
        GROUP BY order_type
        """
    ).fetchall():
        type_counts[r['order_type']] = r['c']
    total_all = sum(type_counts.values())

    conn.close()

    orders = []
    for r in rows:
        d = dict(r)
        d['type_label'] = ORDER_TYPE_LABELS.get(d['order_type'], d['order_type'])
        orders.append(d)

    return render_template(
        'orders_list.html',
        orders=orders,
        order_type=order_type or '',
        search=search,
        page=page,
        per_page=per_page,
        total=total,
        total_pages=total_pages,
        type_counts=type_counts,
        total_all=total_all,
        type_labels=ORDER_TYPE_LABELS,
        order_types=ORDER_TYPES,
        create_links=CREATE_ORDER_LINKS,
    )


def _enrich_order(conn, order: dict) -> dict:
    enriched = []
    for it in order.get('items') or []:
        item = dict(it)
        if item['entity_type'] == 'student':
            row = conn.execute(
                """
                SELECT last_name_UA, first_name_UA, middle_name_UA, group_id
                FROM students WHERE id = ?
                """,
                (item['entity_id'],),
            ).fetchone()
            if row:
                item['label'] = ' '.join(
                    x for x in (
                        row['last_name_UA'],
                        row['first_name_UA'],
                        row['middle_name_UA'] or '',
                    ) if x
                )
                item['student_id'] = item['entity_id']
            else:
                item['label'] = f"студент #{item['entity_id']} (немає в БД)"
        elif item['entity_type'] == 'group':
            row = conn.execute(
                "SELECT name, start_year, study_form FROM groups WHERE id = ?",
                (item['entity_id'],),
            ).fetchone()
            if row:
                item['label'] = f"{row['name']} ({row['start_year']}, {row['study_form']})"
            else:
                item['label'] = f"група #{item['entity_id']}"
        else:
            item['label'] = f"{item['entity_type']} #{item['entity_id']}"

        payload = item.get('payload') or {}
        if isinstance(payload, str):
            try:
                payload = json.loads(payload)
            except Exception:
                payload = {}
        item['payload'] = payload
        item['payload_text'] = _format_payload(order['order_type'], payload)
        enriched.append(item)
    order['items'] = enriched
    return order


@admin_bp.route('/admin/orders/<int:order_id>')
@permission_required('manage_students')
def order_detail(order_id):
    if not _can_manage():
        flash('Немає доступу.', 'error')
        return redirect(url_for('students.student_list'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    if not orders_tables_ready(conn=conn):
        conn.close()
        flash('Таблиці наказів не створені.', 'error')
        return redirect(url_for('admin.orders_list'))

    order = get_order(order_id, conn=conn)
    if not order:
        conn.close()
        flash('Наказ не знайдено.', 'error')
        return redirect(url_for('admin.orders_list'))

    order = _enrich_order(conn, order)

    # Інші накази того ж типу — для перенесення студентів
    peer_orders = conn.execute(
        """
        SELECT id, order_number, order_date, title
        FROM orders
        WHERE order_type = ? AND id != ?
        ORDER BY order_date DESC, id DESC
        LIMIT 100
        """,
        (order.get('order_type'), order_id),
    ).fetchall()
    peer_orders = [dict(r) for r in peer_orders]
    conn.close()

    return render_template(
        'order_detail.html',
        order=order,
        type_labels=ORDER_TYPE_LABELS,
        peer_orders=peer_orders,
    )


@admin_bp.route('/admin/orders/<int:order_id>/edit', methods=['GET', 'POST'])
@permission_required('manage_students')
def order_edit(order_id):
    if not _can_manage():
        flash('Немає доступу.', 'error')
        return redirect(url_for('students.student_list'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    if not orders_tables_ready(conn=conn):
        conn.close()
        flash('Таблиці наказів не створені.', 'error')
        return redirect(url_for('admin.orders_list'))

    order = get_order(order_id, conn=conn)
    if not order:
        conn.close()
        flash('Наказ не знайдено.', 'error')
        return redirect(url_for('admin.orders_list'))

    if request.method == 'POST':
        action = request.form.get('action') or 'save_meta'

        if action == 'save_meta':
            order_number = (request.form.get('order_number') or '').strip()
            order_date = (request.form.get('order_date') or '').strip()
            title = (request.form.get('title') or '').strip() or None
            notes = (request.form.get('notes') or '').strip() or None

            if not order_number or not order_date:
                flash('Номер і дата наказу обов\'язкові.', 'error')
                conn.close()
                return redirect(url_for('admin.order_edit', order_id=order_id))

            conn.execute(
                """
                UPDATE orders
                SET order_number = ?, order_date = ?, title = ?, notes = ?
                WHERE id = ?
                """,
                (order_number, order_date, title, notes, order_id),
            )
            scan_files = request.files.getlist('scan_files')
            if scan_files and any(f and f.filename for f in scan_files):
                save_multiple_attachments(
                    conn, 'order', order_id, scan_files, 'orders', current_username()
                )
            conn.commit()
            log_action(
                current_username(),
                f'редагував наказ #{order_id} №{order_number} від {order_date}',
            )
            flash('Реквізити наказу збережено.', 'success')
            conn.close()
            return redirect(url_for('admin.order_detail', order_id=order_id))

        if action == 'add_student':
            student_id = request.form.get('student_id', type=int)
            reason = (request.form.get('reason') or '').strip() or None
            if not student_id:
                flash('Оберіть студента.', 'error')
                conn.close()
                return redirect(url_for('admin.order_edit', order_id=order_id))

            exists = conn.execute(
                """
                SELECT 1 FROM order_items
                WHERE order_id = ? AND entity_type = 'student' AND entity_id = ?
                """,
                (order_id, student_id),
            ).fetchone()
            if exists:
                flash('Цей студент уже є в наказі.', 'warning')
                conn.close()
                return redirect(url_for('admin.order_edit', order_id=order_id))

            st = conn.execute(
                "SELECT id, group_id FROM students WHERE id = ?", (student_id,)
            ).fetchone()
            if not st:
                flash('Студента не знайдено.', 'error')
                conn.close()
                return redirect(url_for('admin.order_edit', order_id=order_id))

            payload = {}
            if order['order_type'] == 'enrollment':
                payload = {'group_id': st['group_id']}
            elif order['order_type'] == 'expulsion':
                payload = {
                    'previous_group_id': st['group_id'],
                    'reason': reason or '',
                }
            elif reason:
                payload = {'reason': reason}

            conn.execute(
                """
                INSERT INTO order_items (order_id, entity_type, entity_id, payload)
                VALUES (?, 'student', ?, ?)
                """,
                (
                    order_id,
                    student_id,
                    json.dumps(payload, ensure_ascii=False) if payload else None,
                ),
            )
            conn.commit()
            log_action(
                current_username(),
                f'додав студента #{student_id} до наказу #{order_id}',
            )
            flash('Студента додано до наказу.', 'success')
            conn.close()
            return redirect(url_for('admin.order_edit', order_id=order_id))

        flash('Невідома дія.', 'error')
        conn.close()
        return redirect(url_for('admin.order_edit', order_id=order_id))

    order = _enrich_order(conn, order)

    existing_ids = {
        it['entity_id']
        for it in (order.get('items') or [])
        if it.get('entity_type') == 'student'
    }
    students = conn.execute(
        """
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               g.name AS group_name
        FROM students s
        LEFT JOIN groups g ON g.id = s.group_id
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA COLLATE UKRAINIAN, s.first_name_UA COLLATE UKRAINIAN
        LIMIT 2000
        """
    ).fetchall()
    candidates = []
    for s in students:
        if s['id'] in existing_ids:
            continue
        full = ' '.join(
            x for x in (s['last_name_UA'], s['first_name_UA'], s['middle_name_UA'] or '') if x
        )
        candidates.append({
            'id': s['id'],
            'label': f"{full}" + (f" — {s['group_name']}" if s['group_name'] else ''),
        })
    candidates = sort_ukrainian(candidates, lambda c: c['label'])

    conn.close()
    return render_template(
        'order_edit.html',
        order=order,
        type_labels=ORDER_TYPE_LABELS,
        candidates=candidates,
    )


@admin_bp.route('/admin/orders/<int:order_id>/items/<int:item_id>/remove', methods=['POST'])
@permission_required('manage_students')
def order_item_remove(order_id, item_id):
    if not _can_manage():
        flash('Немає доступу.', 'error')
        return redirect(url_for('students.student_list'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    row = conn.execute(
        "SELECT id, entity_type, entity_id FROM order_items WHERE id = ? AND order_id = ?",
        (item_id, order_id),
    ).fetchone()
    if not row:
        conn.close()
        flash('Позицію не знайдено.', 'error')
        return redirect(url_for('admin.order_detail', order_id=order_id))

    cnt = conn.execute(
        "SELECT COUNT(*) FROM order_items WHERE order_id = ?", (order_id,)
    ).fetchone()[0]
    if cnt <= 1:
        conn.close()
        flash(
            'Не можна прибрати останню позицію. Спочатку додайте іншу.',
            'error',
        )
        return redirect(url_for('admin.order_edit', order_id=order_id))

    conn.execute("DELETE FROM order_items WHERE id = ?", (item_id,))
    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f'прибрав {row["entity_type"]} #{row["entity_id"]} з наказу #{order_id}',
    )
    flash('Позицію прибрано з наказу.', 'success')
    return redirect(url_for('admin.order_edit', order_id=order_id))


def _format_payload(order_type: str, payload: dict) -> str:
    if not payload:
        return ''
    if order_type == 'course_transfer':
        a, b = payload.get('course_from'), payload.get('course_to')
        if a is not None and b is not None:
            return f'курс {a} → {b}'
    if order_type == 'expulsion' and payload.get('reason'):
        return str(payload['reason'])
    if order_type == 'license_transfer':
        return (
            f"ліцензія {payload.get('previous_license_id')} → "
            f"{payload.get('new_license_id')}"
        )
    if order_type == 'enrollment' and payload.get('group_id'):
        return f"група id={payload['group_id']}"
    if order_type == 'graduation' and payload.get('action'):
        return str(payload['action'])
    parts = [f'{k}={v}' for k, v in payload.items() if v is not None and v != '']
    return ', '.join(parts)


@admin_bp.route('/admin/orders/<int:order_id>/delete', methods=['POST'])
@permission_required('manage_students')
def order_delete(order_id):
    """Видалити наказ разом із пунктами (order_items)."""
    if not _can_manage():
        flash('Немає доступу до реєстру наказів.', 'error')
        return redirect(url_for('students.student_list'))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    if not orders_tables_ready(conn=conn):
        conn.close()
        flash('Таблиці наказів не готові.', 'error')
        return redirect(url_for('admin.orders_list'))

    order = get_order(order_id, conn=conn)
    if not order:
        conn.close()
        flash('Наказ не знайдено.', 'error')
        return redirect(url_for('admin.orders_list'))

    number = order.get('order_number') or order_id
    otype = order.get('order_type') or ''
    ok = delete_order(order_id, conn=conn)
    if ok:
        conn.commit()
        log_action(
            current_username(),
            f'видалив наказ № {number} (id={order_id}, type={otype})',
        )
        flash(f'Наказ № {number} видалено.', 'success')
    else:
        flash('Не вдалося видалити наказ.', 'error')
    conn.close()
    return redirect(url_for('admin.orders_list'))


@admin_bp.route(
    '/admin/orders/<int:order_id>/items/<int:item_id>/move',
    methods=['POST'],
)
@permission_required('manage_students')
def order_item_move(order_id, item_id):
    """Перенести студента (order_item) в інший наказ того ж типу."""
    if not _can_manage():
        flash('Немає доступу.', 'error')
        return redirect(url_for('students.student_list'))

    target_id = request.form.get('target_order_id', type=int)
    if not target_id:
        flash('Оберіть наказ для перенесення.', 'error')
        return redirect(url_for('admin.order_detail', order_id=order_id))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    ok, msg = move_order_item(item_id, target_id, conn=conn)
    if ok:
        conn.commit()
        log_action(
            current_username(),
            f'переніс пункт наказу item={item_id} з order={order_id} → {target_id}',
            details=msg,
        )
        flash(msg, 'success')
    else:
        flash(msg, 'error')
    conn.close()
    return redirect(url_for('admin.order_detail', order_id=order_id))
