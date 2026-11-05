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
    format_payload,
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

def _get_group_field_for_order_type(order_type):
    """У яке поле payload записувати group_id для цього типу наказу."""
    return {
        'enrollment':          'group_id',
        'expulsion':           'previous_group_id',
        'graduation':          'group_id',
        'license_transfer':    'group_id',
        'freeze':              'previous_group_id',
        'group_transfer':      'from_group_id',
        'study_form_transfer': 'from_group_id',
    }.get(order_type)


def _extract_group_id_from_payload(order_type, payload):
    """Дістає group_id з payload залежно від типу наказу."""
    if not payload:
        return None
    field = _get_group_field_for_order_type(order_type)
    if field and payload.get(field):
        try:
            return int(payload[field])
        except (TypeError, ValueError):
            return None
    return None

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
    number_filter = (request.args.get('number') or '').strip()
    only_duplicates = request.args.get('only_duplicates') == '1'
    page = max(1, request.args.get('page', 1, type=int))
    per_page = request.args.get('per_page', 50, type=int)
    if per_page not in (20, 50, 100, 200):
        per_page = 50

    # ── Сортування ──
    sort_by = (request.args.get('sort_by') or 'id').strip()
    sort_order = (request.args.get('sort_order') or 'desc').strip().lower()
    if sort_order not in ('asc', 'desc'):
        sort_order = 'desc'

    # ── Дублікати номерів (серед активних наказів) ──
    dup_where = ["status = 'active'"]
    dup_params = []
    if order_type and order_type in ORDER_TYPES:
        dup_where.append('order_type = ?')
        dup_params.append(order_type)

    dup_rows = conn.execute(
        f"""
        SELECT order_number, COUNT(*) AS cnt
        FROM orders
        WHERE {' AND '.join(dup_where)}
        GROUP BY order_number
        HAVING COUNT(*) > 1
        """,
        dup_params,
    ).fetchall()
    duplicate_numbers = {r['order_number']: r['cnt'] for r in dup_rows}

    # ── WHERE ──
    where = ["o.status = 'active'"]
    params = []
    if order_type and order_type in ORDER_TYPES:
        where.append('o.order_type = ?')
        params.append(order_type)
    if number_filter:
        where.append("o.order_number LIKE ?")
        params.append(f'%{number_filter}%')
    if search:
        where.append(
            "(o.order_number LIKE ? OR IFNULL(o.title, '') LIKE ? OR IFNULL(o.notes, '') LIKE ?)"
        )
        q = f'%{search}%'
        params.extend([q, q, q])
    if only_duplicates:
        if duplicate_numbers:
            ph = ','.join('?' * len(duplicate_numbers))
            where.append(f"o.order_number IN ({ph})")
            params.extend(list(duplicate_numbers.keys()))
        else:
            where.append('1=0')

    where_sql = ' AND '.join(where)

    # ── Сортування: whitelist колонок ──
    sortable_columns = {
        'id': 'o.id',
        'order_date': (
            "substr(o.order_date, 7, 4) || substr(o.order_date, 4, 2) "
            "|| substr(o.order_date, 1, 2)"
        ),
        'order_number': "o.order_number COLLATE UKRAINIAN",
        'order_type': 'o.order_type',
        'items_count': 'items_count',
        'created_by': "IFNULL(o.created_by, '') COLLATE UKRAINIAN",
        'title': "IFNULL(o.title, '') COLLATE UKRAINIAN",
    }
    if sort_by not in sortable_columns:
        sort_by = 'id'
    sort_clause = sortable_columns[sort_by]
    # Другорядне сортування — стабільне при рівних значеннях
    order_sql = f"ORDER BY {sort_clause} {sort_order.upper()}, o.id DESC"

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
        {order_sql}
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
        d['duplicates_count'] = duplicate_numbers.get(d['order_number'], 0)
        orders.append(d)

    return render_template(
        'orders_list.html',
        orders=orders,
        order_type=order_type or '',
        search=search,
        number_filter=number_filter,
        only_duplicates=only_duplicates,
        duplicate_numbers=duplicate_numbers,
        page=page,
        per_page=per_page,
        total=total,
        total_pages=total_pages,
        type_counts=type_counts,
        total_all=total_all,
        type_labels=ORDER_TYPE_LABELS,
        order_types=ORDER_TYPES,
        create_links=CREATE_ORDER_LINKS,
        sort_by=sort_by,
        sort_order=sort_order,
    )

def _enrich_order(conn, order: dict) -> dict:
    """
    Збагачує items наказу назвами/деталями і ДЕДУПЛІКУЄ їх.

    Дедуплікація потрібна, бо:
      - UNIQUE-обмеження на (order_id, entity_type, entity_id) у старих БД
        може бути відсутнім;
      - після merge наказ може містити 2 однакові рядки одного студента.

    Ключ: (entity_type, entity_id, group_id_from_payload). Студент з
    паралельним навчанням має РІЗНІ group_id — обидві позиції
    показуються. Студент у тій самій групі — показується один раз.
    """
    enriched = []
    seen_keys = set()
    dupes = 0

    for it in order.get('items') or []:
        item = dict(it)

        # Парсимо payload ДО перевірки дедуплікації
        payload = item.get('payload') or {}
        if isinstance(payload, str):
            try:
                payload = json.loads(payload)
            except Exception:
                payload = {}
        item['payload'] = payload

        # Дедуплікація
        if item['entity_type'] == 'student':
            gid = _extract_group_id_from_payload(order['order_type'], payload)
            key = (item['entity_type'], item['entity_id'], gid)
        else:
            key = (item['entity_type'], item['entity_id'], None)

        if key in seen_keys:
            dupes += 1
            continue
        seen_keys.add(key)

        # ── Збагачення (як раніше) ──
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

        item['payload_text'] = format_payload(conn, order['order_type'], payload)

        if item['entity_type'] == 'student':
            gid = _extract_group_id_from_payload(order['order_type'], payload)
            if gid:
                g_row = conn.execute(
                    "SELECT name FROM groups WHERE id=?", (gid,)
                ).fetchone()
                item['group_name_for_item'] = g_row['name'] if g_row else f'#{gid}'
            else:
                item['group_name_for_item'] = None
        else:
            item['group_name_for_item'] = None

        enriched.append(item)

    order['items'] = enriched
    order['duplicates_hidden'] = dupes
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
            pair_raw = (request.form.get('pair') or '').strip()
            reason = (request.form.get('reason') or '').strip() or None

            if not pair_raw:
                flash('Оберіть студента і групу.', 'error')
                conn.close()
                return redirect(url_for('admin.order_edit', order_id=order_id))

            try:
                sid_str, gid_str = pair_raw.split('_', 1)
                student_id = int(sid_str)
                group_id = int(gid_str)
            except (ValueError, AttributeError):
                flash('Некоректний формат пари.', 'error')
                conn.close()
                return redirect(url_for('admin.order_edit', order_id=order_id))

            # ── Перевірка дубліката по парі ──
            group_field = _get_group_field_for_order_type(order['order_type'])
            existing_pairs, existing_student_only = set(), set()
            for it in order.get('items') or []:
                if it.get('entity_type') == 'student':
                    p = it.get('payload') or {}
                    gid_old = _extract_group_id_from_payload(order['order_type'], p)
                    if gid_old:
                        existing_pairs.add((it['entity_id'], gid_old))
                    else:
                        existing_student_only.add(it['entity_id'])

            if group_field:
                if (student_id, group_id) in existing_pairs:
                    flash('Цей студент уже є в наказі для цієї групи.', 'warning')
                    conn.close()
                    return redirect(url_for('admin.order_edit', order_id=order_id))
            else:
                if student_id in existing_student_only:
                    flash('Цей студент уже є в наказі.', 'warning')
                    conn.close()
                    return redirect(url_for('admin.order_edit', order_id=order_id))

            # ── Студент існує і не архівний ──
            st = conn.execute(
                "SELECT id FROM students WHERE id=? AND COALESCE(archived,0)=0",
                (student_id,)
            ).fetchone()
            if not st:
                flash('Студента не знайдено або він архівний.', 'error')
                conn.close()
                return redirect(url_for('admin.order_edit', order_id=order_id))

            # ── Пара реально активна (enrollment або legacy group_id) ──
            is_active = conn.execute(
                """
                SELECT 1 FROM student_enrollments
                WHERE student_id=? AND group_id=? AND status='active'
                """,
                (student_id, group_id)
            ).fetchone()
            if not is_active:
                legacy = conn.execute(
                    "SELECT 1 FROM students WHERE id=? AND group_id=?",
                    (student_id, group_id)
                ).fetchone()
                if not legacy:
                    flash('Цей студент не має активного навчання в цій групі.', 'error')
                    conn.close()
                    return redirect(url_for('admin.order_edit', order_id=order_id))

            # ── payload ──
            payload = {}
            if group_field:
                payload[group_field] = group_id
            if order['order_type'] == 'expulsion' and reason:
                payload['reason'] = reason
            elif reason and not group_field:
                payload['reason'] = reason

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
                f'додав студента #{student_id} (група #{group_id}) до наказу #{order_id}',
            )
            flash('Студента додано до наказу.', 'success')
            conn.close()
            return redirect(url_for('admin.order_edit', order_id=order_id))

        flash('Невідома дія.', 'error')
        conn.close()
        return redirect(url_for('admin.order_edit', order_id=order_id))

    # ── GET ──
    order = _enrich_order(conn, order)

    # Пари, які вже в наказі
    existing_pairs, existing_student_only = set(), set()
    for it in (order.get('items') or []):
        if it.get('entity_type') == 'student':
            p = it.get('payload') or {}
            gid = _extract_group_id_from_payload(order['order_type'], p)
            if gid:
                existing_pairs.add((it['entity_id'], gid))
            else:
                existing_student_only.add(it['entity_id'])

    # ── Активні пари (студент × група) ──
    candidates_rows = conn.execute("""
        SELECT s.id AS student_id,
               s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               g.id AS group_id, g.name AS group_name,
               e.is_primary
        FROM students s
        JOIN student_enrollments e
             ON e.student_id = s.id AND e.status = 'active'
        JOIN groups g ON g.id = e.group_id
        WHERE COALESCE(s.archived, 0) = 0
          AND COALESCE(g.archived, 0) = 0
          AND NOT EXISTS (
              SELECT 1 FROM frozen_students fz
              WHERE fz.student_id = s.id AND fz.resolved_at IS NULL
          )
          AND NOT EXISTS (
              SELECT 1 FROM order_items oi
              JOIN orders o ON o.id = oi.order_id
              WHERE oi.entity_type = 'student'
                AND oi.entity_id = s.id
                AND o.order_type = 'expulsion'
                AND o.status = 'active'
          )
        UNION
        SELECT s.id AS student_id,
               s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               g.id AS group_id, g.name AS group_name,
               1 AS is_primary
        FROM students s
        JOIN groups g ON g.id = s.group_id
        WHERE COALESCE(s.archived, 0) = 0
          AND COALESCE(g.archived, 0) = 0
          AND s.group_id IS NOT NULL
          AND NOT EXISTS (
              SELECT 1 FROM student_enrollments e2
              WHERE e2.student_id = s.id AND e2.status = 'active'
          )
          AND NOT EXISTS (
              SELECT 1 FROM frozen_students fz
              WHERE fz.student_id = s.id AND fz.resolved_at IS NULL
          )
          AND NOT EXISTS (
              SELECT 1 FROM order_items oi
              JOIN orders o ON o.id = oi.order_id
              WHERE oi.entity_type = 'student'
                AND oi.entity_id = s.id
                AND o.order_type = 'expulsion'
                AND o.status = 'active'
          )
        ORDER BY is_primary DESC,
                 last_name_UA COLLATE UKRAINIAN,
                 first_name_UA COLLATE UKRAINIAN
        LIMIT 5000
    """).fetchall()

    group_field = _get_group_field_for_order_type(order['order_type'])

    candidates = []
    for s in candidates_rows:
        key = (s['student_id'], s['group_id'])
        if group_field:
            if key in existing_pairs:
                continue
        else:
            if s['student_id'] in existing_student_only:
                continue

        full = ' '.join(
            x for x in (s['last_name_UA'], s['first_name_UA'],
                        s['middle_name_UA'] or '') if x
        )
        label = full + ' — ' + s['group_name']
        label += ' (осн.)' if s['is_primary'] else ' (парал.)'

        candidates.append({
            'student_id': s['student_id'],
            'group_id': s['group_id'],
            'label': label,
        })

    candidates = sort_ukrainian(candidates, lambda c: c['label'])

    conn.close()
    return render_template(
        'order_edit.html',
        order=order,
        type_labels=ORDER_TYPE_LABELS,
        candidates=candidates,
        group_field=group_field,
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

@admin_bp.route('/admin/orders/merge', methods=['POST'])
@permission_required('manage_students')
def orders_merge():
    if not _can_manage():
        flash('Немає доступу.', 'error')
        return redirect(url_for('students.student_list'))

    from routes.orders_service import merge_orders

    order_ids = []
    for v in request.form.getlist('order_ids'):
        try:
            order_ids.append(int(v))
        except (TypeError, ValueError):
            continue

    main_id = request.form.get('main_order_id', type=int)
    new_number = (request.form.get('new_number') or '').strip() or None
    new_date = (request.form.get('new_date') or '').strip() or None
    new_notes_raw = request.form.get('new_notes')
    new_notes = (new_notes_raw or '').strip() if new_notes_raw is not None else None

    if len(order_ids) < 2:
        flash('Оберіть хоча б два накази для об\'єднання.', 'error')
        return redirect(url_for('admin.orders_list'))

    if main_id is None or main_id not in order_ids:
        flash('Оберіть, будь ласка, "головний" наказ серед обраних.', 'error')
        return redirect(url_for('admin.orders_list'))

    source_ids = [oid for oid in order_ids if oid != main_id]

    conn = get_db()
    conn.row_factory = sqlite3.Row

    if not orders_tables_ready(conn=conn):
        conn.close()
        flash('Таблиці наказів не створені.', 'error')
        return redirect(url_for('admin.orders_list'))

    result = merge_orders(
        main_id, source_ids, conn=conn,
        new_number=new_number, new_date=new_date, new_notes=new_notes,
    )

    if not result['ok']:
        conn.close()
        flash(f"Помилка об'єднання: {result['error']}", 'error')
        return redirect(url_for('admin.orders_list'))

    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"об'єднав накази: "
        f"{', '.join(f'#{i}' for i in source_ids)} → #{main_id}",
        details=(
            f"тип: {result['order_type']}, "
            f"перенесено позицій: {result['moved_items']}, "
            f"дублікатів пропущено: {result['skipped_duplicates']}, "
            f"сканів: {result['moved_attachments']}"
        ),
    )

    msg = (f"Об'єднано {len(order_ids)} наказів у "
           f"№ {result['main_number']} від {result['main_date']}.")
    if result['moved_items']:
        msg += f" Перенесено позицій: {result['moved_items']}."
    if result['skipped_duplicates']:
        msg += f" Дублікатів пропущено: {result['skipped_duplicates']}."
    if result['moved_attachments']:
        msg += f" Сканів перенесено: {result['moved_attachments']}."
    flash(msg, 'success')

    return redirect(url_for('admin.order_detail', order_id=main_id))

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
