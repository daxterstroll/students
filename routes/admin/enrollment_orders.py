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
    orders_tables_ready,
    ORDER_TYPE_LABELS,
)


def _can_manage_orders() -> bool:
    """Адмін або явний дозвіл manage_orders / manage_students."""
    if session.get('is_admin'):
        return True
    perms = set(session.get('permissions') or [])
    return 'create_enrollment_order' in perms

def _parse_bulk_line(line: str) -> dict | None:
    """
    Розпізнає рядок:
        ПІБ [TAB] [Група: X [TAB]] Код наказу про зарахування: Y; Номер наказу: Z; Дата наказу: W
    Поле "Група:" необов'язкове.
    """
    line = line.strip()
    if not line or line.startswith('#'):
        return None

    if '\t' in line:
        parts = line.split('\t')
        fio = parts[0].strip()
        rest = '\t'.join(parts[1:])
    elif 'Код наказу' in line:
        idx = line.index('Код наказу')
        fio = line[:idx].strip().rstrip(';').strip()
        rest = line[idx:]
    elif 'Номер наказу' in line:
        idx = line.index('Номер наказу')
        fio = line[:idx].strip().rstrip(';').strip()
        rest = line[idx:]
    else:
        return None

    group_match  = re.search(r'Група:\s*([^\t;\n]+)', rest)
    code_match   = re.search(r'Код наказу про зарахування:\s*([^;\n]+)', rest)
    number_match = re.search(r'Номер наказу:\s*([^;\n]+)', rest)
    date_match   = re.search(r'Дата наказу:\s*([\d\.]+)', rest)

    return {
        'fio':        fio,
        'group_hint': group_match.group(1).strip() if group_match else None,
        'code':       code_match.group(1).strip() if code_match else None,
        'number':     number_match.group(1).strip() if number_match else None,
        'date':       date_match.group(1).strip() if date_match else None,
        'raw':        line,
    }


@admin_bp.route('/admin/orders/enrollment', methods=['GET', 'POST'])
@permission_required('create_enrollment_order')
def enrollment_order():
    """
    Крок 1–2 майстра: вибір групи (або всіх без наказу) і відмітка
    пар (студент, група). Студент з паралельним навчанням з'являється
    окремим рядком для кожної групи, де він ще без наказу.
    """
    if not _can_manage_orders():
        flash('Немає доступу до наказів.', 'error')
        return redirect(url_for('students.student_list'))

    from routes.orders_service import get_enrollment_pairs
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

    existing_pairs = get_enrollment_pairs(conn=conn)

    # ── POST: зібрати обрані пари і показати confirm ──
    if request.method == 'POST':
        action = request.form.get('action') or 'preview'
        group_id = request.form.get('group_id', type=int)
        only_missing = request.form.get('only_missing') == '1'
        selected_pairs = request.form.getlist('pairs')

        # ── Додати до існуючих наказів (інлайн-вибір біля кожного рядка) ──
        if action == 'assign_existing':
            assignment = {}
            for pair_raw in selected_pairs:
                oid_raw = request.form.get(f'assign_{pair_raw}')
                if not oid_raw:
                    continue  # цей рядок піде в "створити новий"
                try:
                    sid_str, gid_str = pair_raw.split('_', 1)
                    sid, gid = int(sid_str), int(gid_str)
                    oid = int(oid_raw)
                except (ValueError, AttributeError):
                    continue
                assignment[(sid, gid)] = oid

            return _do_confirm_assign_existing(
                conn, assignment,
                redirect_group_id=group_id,
                redirect_only_missing=only_missing,
            )

        if action == 'confirm':
            return _do_confirm(conn, selected_pairs)

        # ── preview: беремо ЛИШЕ тих, у кого не обрано існуючий наказ ──
        filtered_pairs = []
        for pair_raw in selected_pairs:
            oid_raw = request.form.get(f'assign_{pair_raw}')
            if oid_raw:
                continue  # цей рядок призначений для існуючого наказу
            filtered_pairs.append(pair_raw)
        selected_pairs = filtered_pairs

        # preview
        students = _load_candidates(conn, group_id, only_missing, existing_pairs)
        if selected_pairs:
            sel_set = set(selected_pairs)
            students = [s for s in students
                        if f"{s['student_id']}_{s['group_id']}" in sel_set]

        if not students:
            flash(
                'Не обрано жодного запису для нового наказу. '
                'Або оберіть студентів, або призначте їм існуючий наказ.',
                'warning',
            )
            conn.close()
            return redirect(url_for('admin.enrollment_order',
                                    group_id=group_id,
                                    only_missing='1' if only_missing else None))

        existing_orders = _load_existing_enrollment_orders(conn)

        conn.close()
        return render_template(
            'enrollment_order_confirm.html',
            students=students,
            existing_orders=existing_orders,
            group_id=group_id,
            only_missing=only_missing,
            order_type_label=ORDER_TYPE_LABELS['enrollment'],
        )

    # ── GET: форма вибору ──
    group_id = request.args.get('group_id', type=int)
    only_missing = request.args.get('only_missing', '1') == '1'
    students = _load_candidates(conn, group_id, only_missing, existing_pairs)

    # Статистика
    total_pairs_in_db = conn.execute(
        """
        SELECT COUNT(*) AS c FROM (
            SELECT s.id, e.group_id FROM students s
            JOIN student_enrollments e ON e.student_id = s.id AND e.status = 'active'
            WHERE COALESCE(s.archived, 0) = 0
            UNION
            SELECT s.id, s.group_id FROM students s
            WHERE COALESCE(s.archived, 0) = 0
              AND s.group_id IS NOT NULL
              AND NOT EXISTS (SELECT 1 FROM student_enrollments e2
                              WHERE e2.student_id = s.id AND e2.status = 'active')
        )
        """
    ).fetchone()['c']

    enrolled_count = len(existing_pairs)
    missing_total = max(0, total_pairs_in_db - enrolled_count)

    # ── Існуючі накази для "Додати до існуючого" ──
    existing_orders = _load_existing_enrollment_orders(conn)

    conn.close()
    return render_template(
        'enrollment_order.html',
        groups=groups,
        students=students,
        group_id=group_id,
        only_missing=only_missing,
        enrolled_count=enrolled_count,
        missing_total=missing_total,
        existing_orders=existing_orders,      
        order_type_label=ORDER_TYPE_LABELS['enrollment'],
    )

def _load_candidates(conn, group_id, only_missing, existing_pairs):
    """
    Повертає СПИСОК ПАР (студент, група) — по одному рядку на кожен
    активний enrollment студента. Студент з паралельним навчанням
    з'явиться двічі (для КН-21 та для МГ-1 окремо).

    ВИКЛЮЧАЄ:
      - архівних студентів (s.archived = TRUE)
      - архівні групи (g.archived = TRUE)
      - заморожених студентів (активний запис у frozen_students)
      - відрахованих студентів (активний наказ про відрахування)

    existing_pairs — множина (student_id, group_id), для яких наказ про
    зарахування вже створено.
    """
    common_filters = """
        AND COALESCE(s.archived, 0) = 0
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
    """

    sql = f"""
        SELECT s.id AS student_id,
               s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date,
               g.id AS group_id, g.name AS group_name,
               g.start_year, g.study_form, g.program_credits,
               e.is_primary
        FROM students s
        JOIN student_enrollments e
             ON e.student_id = s.id AND e.status = 'active'
        JOIN groups g ON g.id = e.group_id
        WHERE COALESCE(g.archived, 0) = 0
        {common_filters}
    """
    params = []
    if group_id:
        sql += " AND g.id = ?"
        params.append(group_id)
    sql += (" ORDER BY s.last_name_UA COLLATE UKRAINIAN, "
            "s.first_name_UA COLLATE UKRAINIAN, e.is_primary DESC")

    rows = [dict(r) for r in conn.execute(sql, params).fetchall()]

    # Legacy-fallback: студенти без жодного активного enrollment'а,
    # але з students.group_id. Теж з тими самими винятками.
    legacy_sql = f"""
        SELECT s.id AS student_id,
               s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date,
               g.id AS group_id, g.name AS group_name,
               g.start_year, g.study_form, g.program_credits,
               1 AS is_primary
        FROM students s
        JOIN groups g ON g.id = s.group_id
        WHERE COALESCE(g.archived, 0) = 0
          AND s.group_id IS NOT NULL
          AND NOT EXISTS (
              SELECT 1 FROM student_enrollments e2
              WHERE e2.student_id = s.id AND e2.status = 'active'
          )
        {common_filters}
    """
    legacy_params = []
    if group_id:
        legacy_sql += " AND g.id = ?"
        legacy_params.append(group_id)
    legacy_sql += (" ORDER BY s.last_name_UA COLLATE UKRAINIAN, "
                   "s.first_name_UA COLLATE UKRAINIAN")
    rows += [dict(r) for r in conn.execute(legacy_sql, legacy_params).fetchall()]

    for r in rows:
        r['has_enrollment'] = (r['student_id'], r['group_id']) in existing_pairs
        r['full_name'] = ' '.join(
            x for x in (r['last_name_UA'], r['first_name_UA'], r['middle_name_UA'] or '')
            if x
        )

    if only_missing:
        rows = [r for r in rows if not r['has_enrollment']]
    return rows
    
def _load_existing_enrollment_orders(conn, limit=100):
    """Список активних наказів про зарахування для dropdown."""
    return conn.execute(
        """
        SELECT o.id, o.order_number, o.order_date,
               (SELECT COUNT(*) FROM order_items oi
                WHERE oi.order_id = o.id AND oi.entity_type = 'student') AS items_count
        FROM orders o
        WHERE o.order_type = 'enrollment' AND o.status = 'active'
        ORDER BY o.order_date DESC, o.id DESC
        LIMIT ?
        """,
        (limit,),
    ).fetchall()

def _do_confirm_assign_existing(conn, assignment, redirect_group_id=None,
                                redirect_only_missing=False):
    """
    Розподіляє пари (студент, група) по існуючих наказах.

    assignment — dict {(student_id, group_id): order_id}

    Групує по order_id і для кожного додає items одним пакетом.
    Повертає redirect на сторінку enrollment (щоб можна було продовжити
    обробку решти студентів, якщо користувач обрав лише частину).
    """
    if not assignment:
        flash(
            'Не обрано жодного наказу. Оберіть наказ у колонці «Наказ» '
            'біля потрібних студентів.',
            'warning',
        )
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    # Групуємо по order_id
    by_order = {}
    for (sid, gid), oid in assignment.items():
        by_order.setdefault(oid, []).append((sid, gid))

    # Отримуємо інфо про накази одним запитом
    order_ids = list(by_order.keys())
    ph = ','.join('?' * len(order_ids))
    orders_info = conn.execute(
        f"SELECT id, order_number, order_date, order_type, status "
        f"FROM orders WHERE id IN ({ph})",
        order_ids,
    ).fetchall()
    orders_map = {r['id']: dict(r) for r in orders_info}

    from routes.orders_service import get_enrollment_pairs
    existing_pairs = get_enrollment_pairs(conn=conn)

    total_added = 0
    total_skipped_existing = 0
    total_skipped_archived = 0
    total_skipped_not_active = 0
    skipped_orders = []
    per_order_summary = []

    try:
        for oid, pairs in by_order.items():
            order = orders_map.get(oid)
            if not order:
                skipped_orders.append(f'#{oid} (не знайдено)')
                continue
            if order['status'] != 'active':
                skipped_orders.append(f"№{order['order_number']} (не активний)")
                continue
            if order['order_type'] != 'enrollment':
                skipped_orders.append(
                    f"№{order['order_number']} (не наказ про зарахування)"
                )
                continue

            added_here = 0
            for sid, gid in pairs:
                if (sid, gid) in existing_pairs:
                    total_skipped_existing += 1
                    continue

                st = conn.execute(
                    "SELECT archived FROM students WHERE id=?", (sid,)
                ).fetchone()
                if not st or st['archived']:
                    total_skipped_archived += 1
                    continue

                is_active = conn.execute(
                    "SELECT 1 FROM student_enrollments "
                    "WHERE student_id=? AND group_id=? AND status='active'",
                    (sid, gid),
                ).fetchone()
                if not is_active:
                    legacy = conn.execute(
                        "SELECT 1 FROM students WHERE id=? AND group_id=?",
                        (sid, gid),
                    ).fetchone()
                    if not legacy:
                        total_skipped_not_active += 1
                        continue

                payload = json.dumps({'group_id': gid}, ensure_ascii=False)
                conn.execute(
                    "INSERT INTO order_items "
                    "(order_id, entity_type, entity_id, payload) "
                    "VALUES (?, 'student', ?, ?)",
                    (oid, sid, payload),
                )
                existing_pairs.add((sid, gid))
                added_here += 1
                total_added += 1

            if added_here:
                per_order_summary.append({
                    'id': oid,
                    'number': order['order_number'],
                    'date': order['order_date'],
                    'added': added_here,
                })

        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        logger.error(f'assign_existing: {e}', exc_info=True)
        flash(f'Помилка додавання: {e}', 'error')
        return redirect(url_for('admin.enrollment_order'))

    conn.close()

    redirect_args = {}
    if redirect_group_id:
        redirect_args['group_id'] = redirect_group_id
    if redirect_only_missing:
        redirect_args['only_missing'] = '1'

    if total_added == 0:
        parts = []
        if total_skipped_existing:
            parts.append(f'уже мали наказ: {total_skipped_existing}')
        if total_skipped_archived:
            parts.append(f'архівні: {total_skipped_archived}')
        if total_skipped_not_active:
            parts.append(f'не активні: {total_skipped_not_active}')
        if skipped_orders:
            parts.append(f'пропущено наказів: {", ".join(skipped_orders)}')
        msg = 'Нічого не додано'
        if parts:
            msg += ' (' + '; '.join(parts) + ')'
        flash(msg + '.', 'warning')
        return redirect(url_for('admin.enrollment_order', **redirect_args))

    details_str = ', '.join(
        f"№{o['number']} (+{o['added']})" for o in per_order_summary
    )
    log_action(
        current_username(),
        f'додав до існуючих наказів про зарахування: '
        f'{total_added} запис(ів) у {len(per_order_summary)} наказ(ів)',
        details=details_str,
    )

    msg = (
        f'Додано {total_added} запис(ів) до {len(per_order_summary)} '
        f'наказ(ів): {details_str}.'
    )
    parts = []
    if total_skipped_existing:
        parts.append(f'уже мали наказ: {total_skipped_existing}')
    if total_skipped_archived:
        parts.append(f'архівні: {total_skipped_archived}')
    if total_skipped_not_active:
        parts.append(f'не активні: {total_skipped_not_active}')
    if skipped_orders:
        parts.append(f'пропущено наказів: {", ".join(skipped_orders)}')
    if parts:
        msg += ' Пропущено: ' + '; '.join(parts) + '.'
    flash(msg, 'success')
    return redirect(url_for('admin.enrollment_order', **redirect_args))

def _do_confirm_add_to_existing(conn, parsed_pairs):
    """
    Додає пари (студент, група) до вже існуючого наказу про зарахування.
    Повертає redirect на картку наказу.
    """
    target_order_id = request.form.get('target_order_id', type=int)

    if not target_order_id:
        flash('Оберіть наказ, до якого додати студентів.', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    target = conn.execute(
        "SELECT id, order_number, order_date, order_type, status "
        "FROM orders WHERE id=?",
        (target_order_id,),
    ).fetchone()

    if not target:
        flash('Наказ не знайдено.', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))
    if target['status'] != 'active':
        flash('Цей наказ не активний — додавання неможливе.', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))
    if target['order_type'] != 'enrollment':
        flash('Додавати можна лише до наказів про зарахування.', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    # Дедуплікація внутрішньої форми
    parsed_pairs = list(dict.fromkeys(parsed_pairs))

    # Усі пари, які вже мають наказ про зарахування (включно з цим наказом)
    from routes.orders_service import get_enrollment_pairs
    all_existing_pairs = get_enrollment_pairs(conn=conn)

    added = 0
    skipped_existing = 0
    skipped_not_active = 0
    skipped_archived = 0

    for sid, gid in parsed_pairs:
        # Уже є наказ для цієї пари (в цьому або будь-якому іншому)
        if (sid, gid) in all_existing_pairs:
            skipped_existing += 1
            continue

        # Студент існує і не в архіві
        st = conn.execute(
            "SELECT archived FROM students WHERE id=?",
            (sid,),
        ).fetchone()
        if not st or st['archived']:
            skipped_archived += 1
            continue

        # Пара реально активна (enrollment або legacy group_id)
        is_active = conn.execute(
            "SELECT 1 FROM student_enrollments "
            "WHERE student_id=? AND group_id=? AND status='active'",
            (sid, gid),
        ).fetchone()
        if not is_active:
            legacy = conn.execute(
                "SELECT 1 FROM students WHERE id=? AND group_id=?",
                (sid, gid),
            ).fetchone()
            if not legacy:
                skipped_not_active += 1
                continue

        # Додаємо до наказу
        payload = json.dumps({'group_id': gid}, ensure_ascii=False)
        try:
            conn.execute(
                """
                INSERT INTO order_items (order_id, entity_type, entity_id, payload)
                VALUES (?, 'student', ?, ?)
                """,
                (target_order_id, sid, payload),
            )
            all_existing_pairs.add((sid, gid))
            added += 1
        except Exception as e:
            logger.error(
                f'[enrollment] failed to add pair ({sid}, {gid}) '
                f'to order #{target_order_id}: {e}'
            )

    if added == 0:
        conn.rollback()
        conn.close()
        parts = []
        if skipped_existing:
            parts.append(f'уже мають наказ: {skipped_existing}')
        if skipped_archived:
            parts.append(f'архівні: {skipped_archived}')
        if skipped_not_active:
            parts.append(f'не активні enrollments: {skipped_not_active}')
        msg = 'Нічого не додано'
        if parts:
            msg += ' (' + '; '.join(parts) + ')'
        flash(msg + '.', 'warning')
        return redirect(url_for('admin.enrollment_order'))

    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f'додав до існуючого наказу №{target["order_number"]} '
        f'від {target["order_date"]} (order_id={target_order_id})',
        details=(
            f'студентів додано: {added}'
            + (f'; уже були в наказі: {skipped_existing}' if skipped_existing else '')
            + (f'; архівні: {skipped_archived}' if skipped_archived else '')
            + (f'; не активні: {skipped_not_active}' if skipped_not_active else '')
        ),
    )

    msg = (
        f'До наказу №{target["order_number"]} від {target["order_date"]} '
        f'додано {added} запис(ів).'
    )
    parts = []
    if skipped_existing:
        parts.append(f'уже були в наказі: {skipped_existing}')
    if skipped_archived:
        parts.append(f'архівні: {skipped_archived}')
    if skipped_not_active:
        parts.append(f'не активні: {skipped_not_active}')
    if parts:
        msg += ' Пропущено: ' + '; '.join(parts) + '.'
    flash(msg, 'success')
    return redirect(url_for('admin.order_detail', order_id=target_order_id))    
    
def _do_confirm(conn, selected_pairs):
    """
    Створює наказ про зарахування для вибраних пар (студент, група)
    АБО додає їх до вже існуючого наказу (mode='existing').
    """
    mode = (request.form.get('mode') or 'new').strip()

    # ── Розпаковуємо пари (потрібно для обох режимів) ──
    parsed_pairs = []
    for raw in selected_pairs:
        try:
            sid_str, gid_str = raw.split('_', 1)
            parsed_pairs.append((int(sid_str), int(gid_str)))
        except (ValueError, AttributeError):
            continue

    if not parsed_pairs:
        flash('Некоректний формат вибраних записів.', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    # ── Режим "додати до існуючого наказу" ──
    if mode == 'existing':
        return _do_confirm_add_to_existing(conn, parsed_pairs)

    # ── Режим "створити новий наказ" ──
    order_number = (request.form.get('order_number') or '').strip()
    order_date = (request.form.get('order_date') or '').strip()
    notes = (request.form.get('notes') or '').strip() or None

    if not order_number or not order_date:
        flash('Заповніть номер і дату наказу.', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    if not selected_pairs:
        flash('Немає записів у наказі.', 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    # Дедуплікація (на випадок дублювання з JS)
    parsed_pairs = list(dict.fromkeys(parsed_pairs))

    # Перевірка, що кожна пара — реально активний enrollment студента
    existing_pairs = get_enrollment_pairs(conn=conn)
    valid_pairs = []
    missing_students = 0
    skipped_already = 0
    skipped_not_active = 0
    skipped_frozen = 0
    skipped_expelled = 0

    for sid, gid in parsed_pairs:
        st = conn.execute(
            "SELECT archived FROM students WHERE id=? AND COALESCE(archived, 0)=0",
            (sid,)
        ).fetchone()
        if not st:
            missing_students += 1
            continue

        is_active = conn.execute(
            "SELECT 1 FROM student_enrollments "
            "WHERE student_id=? AND group_id=? AND status='active'",
            (sid, gid)
        ).fetchone()
        if not is_active:
            # legacy-fallback
            legacy = conn.execute(
                "SELECT 1 FROM students WHERE id=? AND group_id=?", (sid, gid)
            ).fetchone()
            if not legacy:
                skipped_not_active += 1
                continue
                # ── Заморожений? ──
        frozen = conn.execute(
            "SELECT 1 FROM frozen_students "
            "WHERE student_id=? AND resolved_at IS NULL",
            (sid,)
        ).fetchone()
        if frozen:
            skipped_frozen += 1
            continue

        # ── Відрахований? (активний наказ про відрахування) ──
        expelled = conn.execute(
            """
            SELECT 1 FROM order_items oi
            JOIN orders o ON o.id = oi.order_id
            WHERE oi.entity_type='student' AND oi.entity_id=?
              AND o.order_type='expulsion' AND o.status='active'
            LIMIT 1
            """,
            (sid,)
        ).fetchone()
        if expelled:
            skipped_expelled += 1
            continue

        if (sid, gid) in existing_pairs:
            skipped_already += 1
            continue

        valid_pairs.append((sid, gid))

    if not valid_pairs:
        msg = 'Немає валідних записів для наказу.'
        parts = []
        if skipped_already: parts.append(f'уже мали наказ: {skipped_already}')
        if skipped_not_active: parts.append(f'не активні enrollments: {skipped_not_active}')
        if missing_students: parts.append(f'не знайдені/архів: {missing_students}')
        if parts:
            msg += ' (' + '; '.join(parts) + ')'
        flash(msg, 'error')
        conn.close()
        return redirect(url_for('admin.enrollment_order'))

    items = [
        {
            'entity_type': 'student',
            'entity_id': sid,
            'payload': {'group_id': gid},
        }
        for sid, gid in valid_pairs
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
        f'(записів: {len(valid_pairs)}, order_id={order_id})',
    )
    msg = (f'Наказ про зарахування №{order_number} оформлено. '
           f'Записів: {len(valid_pairs)}.')
    if skipped_already:
        msg += f' Пропущено (уже мали наказ): {skipped_already}.'
    if skipped_not_active:
        msg += f' Пропущено (не активні enrollments): {skipped_not_active}.'
    flash(msg, 'success')
    if skipped_frozen:
        msg += f' Пропущено (заморожені): {skipped_frozen}.'
    if skipped_expelled:
        msg += f' Пропущено (відраховані): {skipped_expelled}.'
    return redirect(url_for('admin.enrollment_order'))
    
# ═══════════════════════════════════════════════════════════════════════
# Масове створення наказів про зарахування (v2 — з чанкінгом і редіректом)
# ═══════════════════════════════════════════════════════════════════════
import re
import json
import uuid
import time
from datetime import datetime
from pathlib import Path
from flask import jsonify

BULK_SESSIONS_DIR = Path('static') / 'uploads' / 'import_sessions' / 'enrollment_bulk'
BULK_SESSIONS_DIR.mkdir(parents=True, exist_ok=True)
BULK_SESSION_TTL = 3600  # 1 година


def _cleanup_bulk_sessions():
    now = time.time()
    for f in BULK_SESSIONS_DIR.iterdir():
        try:
            if f.is_file() and now - f.stat().st_mtime > BULK_SESSION_TTL:
                f.unlink()
        except OSError:
            pass


def _process_bulk_text(raw_text: str):
    from routes.utils import logger
    from routes.orders_service import (
        get_enrollment_pairs, get_active_groups_for_student,
    )

    raw_text = (raw_text or '').strip()
    logger.info(f"[enrollment_bulk] processing {len(raw_text)} chars, "
                f"{raw_text.count(chr(10)) + 1} lines")

    if not raw_text:
        flash('Вставте, будь ласка, список у текстове поле', 'error')
        return redirect(url_for('admin.enrollment_bulk'))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    existing_pairs = get_enrollment_pairs(conn=conn)   # (student, group)
    seen_pairs = set()                                  # створені в цьому batch

    parsed, errors = [], []
    parsed_lines = []

    for line_num, line in enumerate(raw_text.splitlines(), 1):
        if not line.strip():
            continue
        data = _parse_bulk_line(line)
        if not data:
            errors.append({'line': line_num, 'raw': line,
                           'reason': 'не вдалося розпізнати формат'})
            continue
        if not data['fio']:
            errors.append({'line': line_num, 'raw': line, 'reason': 'порожнє ПІБ'})
            continue
        if not data['number'] or not data['date']:
            errors.append({'line': line_num, 'raw': line, 'fio': data['fio'],
                           'reason': 'не вказано номер або дату наказу'})
            continue
        try:
            datetime.strptime(data['date'], '%d.%m.%Y')
        except ValueError:
            errors.append({'line': line_num, 'raw': line, 'fio': data['fio'],
                           'reason': f"некоректна дата '{data['date']}'"})
            continue
        parsed_lines.append((line_num, data))

    all_fios = [d['fio'] for _ln, d in parsed_lines]
    lookups = _batch_lookup_students(cursor, all_fios)

    for line_num, data in parsed_lines:
        student_id, matched_name, score = lookups[data['fio']]

        if not student_id:
            errors.append({'line': line_num, 'raw': data['raw'], 'fio': data['fio'],
                           'reason': 'студента не знайдено (або неоднозначно)'})
            continue

        st = conn.execute("SELECT archived FROM students WHERE id = ?",
                          (student_id,)).fetchone()
        if st and st['archived']:
            errors.append({'line': line_num, 'raw': data['raw'], 'fio': matched_name,
                           'reason': 'студент в архіві'})
            continue

        active_groups = get_active_groups_for_student(cursor, student_id)
        if not active_groups:
            errors.append({'line': line_num, 'raw': data['raw'], 'fio': matched_name,
                           'reason': 'у студента немає активних груп'})
            continue

        target_group = target_group_name = None

        # ── Варіант A: явно вказана група в рядку ──
        if data['group_hint']:
            hint = data['group_hint'].strip()
            for gid, gname, _ip in active_groups:
                if str(gid) == hint or gname == hint or hint.lower() in gname.lower():
                    target_group = gid
                    target_group_name = gname
                    break

            if target_group is None:
                errors.append({'line': line_num, 'raw': data['raw'], 'fio': matched_name,
                               'reason': f"вказана група «{hint}» не знайдена серед активних"})
                continue

            if (student_id, target_group) in existing_pairs:
                errors.append({'line': line_num, 'raw': data['raw'], 'fio': matched_name,
                               'reason': f"у групі «{target_group_name}» наказ про "
                                         f"зарахування вже оформлено"})
                continue
            if (student_id, target_group) in seen_pairs:
                errors.append({'line': line_num, 'raw': data['raw'], 'fio': matched_name,
                               'reason': f"у групі «{target_group_name}» наказ уже "
                                         f"додано вище в цьому списку"})
                continue

        # ── Варіант B: автовибір — перша група без наказу ──
        else:
            for gid, gname, _ip in active_groups:
                if (student_id, gid) in existing_pairs:
                    continue
                if (student_id, gid) in seen_pairs:
                    continue
                target_group = gid
                target_group_name = gname
                break

            if target_group is None:
                n = len(active_groups)
                suffix = f' ({n} груп)' if n > 1 else ''
                errors.append({'line': line_num, 'raw': data['raw'], 'fio': matched_name,
                               'reason': f'усі групи вже мають наказ про зарахування{suffix}'})
                continue

        seen_pairs.add((student_id, target_group))

        # Скільки ще активних груп у студента без наказу (для попередження в превью)
        remaining = sum(
            1 for gid, _n, _p in active_groups
            if (student_id, gid) not in existing_pairs
            and (student_id, gid) not in seen_pairs
        )

        parsed.append({
            'line': line_num,
            'student_id': student_id,
            'matched_name': matched_name,
            'fio_input': data['fio'],
            'score': round(score, 1),
            'code': data['code'],
            'number': data['number'],
            'date': data['date'],
            'group_id': target_group,
            'group_name': target_group_name,
            'other_groups_count': remaining,
            'is_primary_group': any(
                gid == target_group and ip
                for gid, _n, ip in active_groups
            ),
        })

    # Групування за (номер, дата)
    groups = {}
    for p in parsed:
        key = (p['number'], p['date'])
        g = groups.setdefault(key, {'number': p['number'], 'date': p['date'],
                                     'codes': set(), 'students': []})
        if p['code']:
            g['codes'].add(p['code'])
        g['students'].append(p)

    def _key(k):
        try:
            return (datetime.strptime(k[1], '%d.%m.%Y'), k[0])
        except (ValueError, TypeError):
            return (datetime.max, k[0] or '')

    group_list = []
    for k in sorted(groups.keys(), key=_key):
        g = groups[k]
        g['codes'] = sorted(g['codes'])
        group_list.append(g)

    conn.close()

    token = uuid.uuid4().hex[:16]
    state = {'groups': group_list, 'errors': errors}
    with open(BULK_SESSIONS_DIR / f'{token}.json', 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False)

    logger.info(f"[enrollment_bulk] done: {len(parsed)} students, "
                f"{len(group_list)} orders, {len(errors)} errors, token={token}")

    return redirect(url_for('admin.enrollment_bulk_preview', token=token))

def _batch_lookup_students(cursor, fios: list[str]) -> dict:
    """
    ОДИН запит до БД для ВСІХ ПІБ — замість 1 запиту на кожен рядок
    (було 1400 запитів, стало 1). Це і є головне прискорення.
    Повертає {fio: (student_id|None, matched_name|None, score|None)}.
    """
    from rapidfuzz import process, fuzz
    from routes.db import normalize_name_for_match

    cursor.execute(
        "SELECT id, last_name_UA, first_name_UA, middle_name_UA "
        "FROM students WHERE archived=FALSE"
    )
    students = cursor.fetchall()

    if not students:
        return {fio: (None, None, 0.0) for fio in fios}

    norm_names = []
    for s in students:
        parts = (s["last_name_UA"], s["first_name_UA"], s["middle_name_UA"])
        joined = ' '.join(str(p) for p in parts if p and str(p).strip())
        norm_names.append(normalize_name_for_match(joined))

    results = {}
    for fio in fios:
        target = normalize_name_for_match(fio)
        if not target:
            results[fio] = (None, None, 0.0)
            continue

        # 1) точний збіг
        exact_idx = next((i for i, n in enumerate(norm_names) if n == target), None)
        if exact_idx is not None:
            results[fio] = (students[exact_idx]["id"], norm_names[exact_idx], 100.0)
            continue

        # 2) нечіткий
        matches = process.extract(target, norm_names, scorer=fuzz.token_set_ratio, limit=3)
        if not matches or matches[0][1] < 95:
            results[fio] = (None, None, matches[0][1] if matches else 0.0)
            continue

        best_name, best_score, best_idx = matches[0]
        # 3) неоднозначність
        if len(matches) > 1 and matches[1][1] >= 95 and (best_score - matches[1][1]) < 5:
            results[fio] = (None, None, best_score)
            continue

        results[fio] = (students[best_idx]["id"], best_name, best_score)

    return results

# ── Форма ──────────────────────────────────────────────────────────
@admin_bp.route('/admin/orders/enrollment/bulk', methods=['GET', 'POST'])
@permission_required('create_enrollment_order')
def enrollment_bulk():
    if not _can_manage_orders():
        flash('Немає доступу до наказів.', 'error')
        return redirect(url_for('students.student_list'))

    _cleanup_bulk_sessions()

    if request.method == 'GET':
        return render_template('enrollment_bulk_paste.html')

    return _process_bulk_text(request.form.get('raw_text', ''))


# ── Чанкове завантаження ──────────────────────────────────────────
@admin_bp.route('/admin/orders/enrollment/bulk/upload_chunk', methods=['POST'])
@permission_required('create_enrollment_order') 
def enrollment_bulk_upload_chunk():
    if not _can_manage_orders():
        return jsonify(ok=False, error='forbidden'), 403

    token = (request.form.get('session_token') or '').strip()
    chunk = request.form.get('chunk') or ''
    chunk_num = request.form.get('chunk_num', '?')
    total = request.form.get('total', '?')

    if not re.fullmatch(r'[0-9a-f]{16}', token):
        return jsonify(ok=False, error='invalid_token'), 400

    path = BULK_SESSIONS_DIR / f'{token}.txt'
    with open(path, 'a', encoding='utf-8') as f:
        f.write(chunk)

    from routes.utils import logger
    logger.info(f"[enrollment_bulk] chunk {chunk_num}/{total} "
                f"({len(chunk.encode('utf-8'))} bytes) → {token}.txt")

    return jsonify(ok=True)


@admin_bp.route('/admin/orders/enrollment/bulk/finalize', methods=['POST'])
@permission_required('create_enrollment_order')
def enrollment_bulk_finalize():
    if not _can_manage_orders():
        flash('Немає доступу.', 'error')
        return redirect(url_for('students.student_list'))

    token = (request.form.get('session_token') or '').strip()
    if not re.fullmatch(r'[0-9a-f]{16}', token):
        flash('Невірний токен сесії', 'error')
        return redirect(url_for('admin.enrollment_bulk'))

    path = BULK_SESSIONS_DIR / f'{token}.txt'
    if not path.exists():
        flash('Сесія не знайдена — почніть заново', 'error')
        return redirect(url_for('admin.enrollment_bulk'))

    raw_text = path.read_text(encoding='utf-8')
    try:
        path.unlink()
    except OSError:
        pass

    return _process_bulk_text(raw_text)


# ── Превью (окремий GET за токеном) ───────────────────────────────
@admin_bp.route('/admin/orders/enrollment/bulk/preview/<token>')
@permission_required('create_enrollment_order')
def enrollment_bulk_preview(token):
    if not re.fullmatch(r'[0-9a-f]{16}', token or ''):
        flash('Невірний токен', 'error')
        return redirect(url_for('admin.enrollment_bulk'))

    state_file = BULK_SESSIONS_DIR / f'{token}.json'
    if not state_file.exists():
        flash('Сесія застаріла — почніть заново', 'error')
        return redirect(url_for('admin.enrollment_bulk'))

    with open(state_file, 'r', encoding='utf-8') as f:
        state = json.load(f)

    return render_template(
        'enrollment_bulk_preview.html',
        token=token,
        groups=state['groups'],
        errors=state['errors'],
        total_students=sum(len(g['students']) for g in state['groups']),
    )


@admin_bp.route('/admin/orders/enrollment/bulk/confirm', methods=['POST'])
@permission_required('create_enrollment_order')
def enrollment_bulk_confirm():
    if not _can_manage_orders():
        flash('Немає доступу.', 'error')
        return redirect(url_for('students.student_list'))

    token = (request.form.get('token') or '').strip()
    if not re.match(r'^[0-9a-f]{16}$', token):
        flash('Невірний токен сесії', 'error')
        return redirect(url_for('admin.enrollment_bulk'))

    state_file = BULK_SESSIONS_DIR / f'{token}.json'
    if not state_file.exists():
        flash('Сесія застаріла — почніть заново', 'error')
        return redirect(url_for('admin.enrollment_bulk'))

    with open(state_file, 'r', encoding='utf-8') as f:
        state = json.load(f)

    # Чекбокси мають value "<student_id>_<group_id>"
    include_present = request.form.get('include_present') == '1'
    included_keys = set()
    if include_present:
        for v in request.form.getlist('include'):
            included_keys.add(v)

    def is_included(sid, gid) -> bool:
        if not include_present:
            return True  # backward-compat

        # Точне співпадіння пари (студент, група)
        if f'{sid}_{gid}' in included_keys:
            return True

        # Backward-compat для старих сесій, де group_id був None
        # (вони створювались старою версією _process_bulk_text без group_id,
        #  тоді в шаблоні value = "123_None"). Якщо в БД є наказ саме
        # для цієї групи — вважаємо рядок включеним.
        if f'{sid}_None' in included_keys:
            return True

        return False

    from routes.orders_service import get_enrollment_pairs
    conn = get_db()
    conn.row_factory = sqlite3.Row

    current_pairs = get_enrollment_pairs(conn=conn)
    created_orders = []
    skipped_already = 0
    skipped_by_user = 0
    # Множина унікальних студентів, яким створено наказ (для підсумку)
    created_students = set()

    try:
        for group in state['groups']:
            to_create = []
            for s in group['students']:
                sid = s['student_id']
                gid = s.get('group_id')

                # Fallback для старих сесій без group_id — беремо primary
                if gid is None:
                    row = conn.execute(
                        "SELECT group_id FROM students WHERE id=?", (sid,)
                    ).fetchone()
                    gid = row['group_id'] if row else None
                    if gid is None:
                        skipped_already += 1
                        continue
                    s['group_id'] = gid

                if (sid, gid) in current_pairs:
                    skipped_already += 1
                elif not is_included(sid, gid):
                    skipped_by_user += 1
                else:
                    to_create.append(s)

            if not to_create:
                continue

            items = [
                {
                    'entity_type': 'student',
                    'entity_id': s['student_id'],
                    'payload': {'group_id': s['group_id']},
                }
                for s in to_create
            ]

            notes = None
            if group.get('codes'):
                notes = 'Код наказу про зарахування: ' + ', '.join(group['codes'])

            order_id = create_order(
                'enrollment', group['number'], group['date'], items,
                created_by=current_username(), notes=notes, conn=conn,
            )

            for s in to_create:
                current_pairs.add((s['student_id'], s['group_id']))
                created_students.add(s['student_id'])

            created_orders.append({
                'order_id': order_id, 'number': group['number'],
                'date': group['date'], 'students_count': len(to_create),
            })

        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        logger.error(f'bulk enrollment confirm: {e}', exc_info=True)
        flash(f'Помилка створення наказів: {e}', 'error')
        return redirect(url_for('admin.enrollment_bulk'))

    conn.close()

    try:
        state_file.unlink()
    except OSError:
        pass

    total_records = sum(o['students_count'] for o in created_orders)
    log_action(
        current_username(),
        f'масове створення наказів: {len(created_orders)} наказ(ів), '
        f'{total_records} записів, {len(created_students)} унік. студент(ів)',
        details=', '.join(f"№{o['number']} від {o['date']} ({o['students_count']} ос.)"
                          for o in created_orders)
        + (f'; виключено вручну: {skipped_by_user}' if skipped_by_user else ''),
    )

    msg = (f'Створено {len(created_orders)} наказ(ів) для {total_records} записів '
           f'({len(created_students)} унікальних студентів)')
    parts = []
    if skipped_already:
        parts.append(f'пропущено (уже мали наказ): {skipped_already}')
    if skipped_by_user:
        parts.append(f'виключено вручну: {skipped_by_user}')
    if parts:
        msg += '. ' + '; '.join(parts) + '.'
    flash(msg, 'success')
    return redirect(url_for('admin.orders_list', type='enrollment'))