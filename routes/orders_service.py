"""
routes/orders_service.py
========================
Єдина точка роботи з наказами (orders / order_items).

Крок 2 плану: сервісний шар без UI.
Підключення:
  from routes.orders_service import (
      create_order, get_order, list_orders,
      get_student_orders, student_has_enrollment,
      count_students_without_enrollment,
      ORDER_TYPE_LABELS,
  )

Перехідний період:
  старі confirm-ендпоінти (course_transfer / expulsion / license_transfer)
  після запису в legacy-таблиці додатково викликають create_order(...,
  legacy_table=..., legacy_id=...).
"""

from __future__ import annotations

import json
import os
from typing import Any, Iterable, Optional

from routes.db import get_db
from routes.utils import save_multiple_attachments

# ── Довідник типів (UI / FAQ / фільтри) ───────────────────────────────
ORDER_TYPE_LABELS = {
    'enrollment':           'Зарахування',
    'course_transfer':      'Переведення на курс',
    'license_transfer':     'Переведення між ліцензіями',
    'group_transfer':       'Переведення з групи в групу',
    'study_form_transfer':  'Зміна форми навчання',
    'name_change':          'Зміна ПІБ',
    'freeze':               'Заморозка (академовідпустка)',
    'unfreeze':             'Поновлення після заморозки',
    'expulsion':            'Відрахування',
    'graduation':           'Випуск',
}

ORDER_TYPES = tuple(ORDER_TYPE_LABELS.keys())


def _payload_dumps(payload: Optional[dict]) -> Optional[str]:
    if not payload:
        return None
    return json.dumps(payload, ensure_ascii=False)


def _payload_loads(raw: Optional[str]) -> dict:
    if not raw:
        return {}
    try:
        data = json.loads(raw)
        return data if isinstance(data, dict) else {}
    except (TypeError, ValueError, json.JSONDecodeError):
        return {}

def format_payload(conn, order_type: str, payload: dict) -> str:
    """Людиночитабельний опис payload наказу — те, що показується
    в картці наказу (`/admin/orders/<id>`), у списку наказів студента
    (`/students/<id>`, блок «Накази») та в реєстрі.

    Для expulsion: «з групи КН-21 (2022, Денна); причина».
    Для course_transfer: «курс 1 → 2».
    Для license_transfer: «ліцензія Львівська → Польська».
    Для edebo_code_change: «ЄДЕБО 123 → 456».
    Для enrollment: «у групу КН-21 (2022, Денна)».
    Інакше — `key=value, …`.
    """
    if not payload:
        return ''

    if order_type == 'course_transfer':
        a, b = payload.get('course_from'), payload.get('course_to')
        if a is not None and b is not None:
            return f'курс {a} → {b}'

    if order_type == 'expulsion':
        parts = []
        prev_gid = payload.get('previous_group_id')
        if prev_gid:
            row = conn.execute(
                "SELECT name, start_year, study_form FROM groups WHERE id = ?",
                (prev_gid,),
            ).fetchone()
            if row:
                parts.append(
                    f'з групи {row["name"]} ({row["start_year"]}, {row["study_form"]})'
                )
            else:
                parts.append(f'з групи #{prev_gid}')
        reason = payload.get('reason')
        if reason:
            parts.append(str(reason))
        return '; '.join(parts) if parts else ''

    if order_type == 'license_transfer':
        def _lic_name(lid):
            if not lid:
                return '—'
            row = conn.execute(
                "SELECT short_name_ua, name_ua FROM institution_licenses WHERE id = ?",
                (lid,),
            ).fetchone()
            if not row:
                return f'#{lid}'
            return row['short_name_ua'] or row['name_ua']

        return (
            f"ліцензія {_lic_name(payload.get('previous_license_id'))} → "
            f"{_lic_name(payload.get('new_license_id'))}"
        )

    if order_type == 'enrollment' and payload.get('group_id'):
        row = conn.execute(
            "SELECT name, start_year, study_form FROM groups WHERE id = ?",
            (payload['group_id'],),
        ).fetchone()
        if row:
            return f'у групу {row["name"]} ({row["start_year"]}, {row["study_form"]})'
        return f"група id={payload['group_id']}"

    if order_type in ('group_transfer', 'study_form_transfer'):
        from_name = payload.get('from_group_name') or f"#{payload.get('from_group_id') or '—'}"
        to_name = payload.get('to_group_name') or f"#{payload.get('to_group_id') or '—'}"
        text = f'{from_name} → {to_name}'
        if order_type == 'study_form_transfer':
            ff = payload.get('from_form')
            tf = payload.get('to_form')
            if ff and tf:
                text += f' ({ff} → {tf})'
        return text

    if order_type == 'name_change':
        old = payload.get('old_name') or {}
        new = payload.get('new_name') or {}
        old_pib = ' '.join(x for x in [old.get('last_name_UA'), old.get('first_name_UA'),
                                        old.get('middle_name_UA')] if x) or '—'
        new_pib = ' '.join(x for x in [new.get('last_name_UA'), new.get('first_name_UA'),
                                        new.get('middle_name_UA')] if x) or '—'
        return f'{old_pib} → {new_pib}'

    if order_type == 'graduation' and payload.get('action'):
        return str(payload['action'])

    parts = [f'{k}={v}' for k, v in payload.items() if v is not None and v != '']
    return ', '.join(parts)
    
def create_order(
    order_type: str,
    order_number: str,
    order_date: str,
    items: Iterable[dict],
    *,
    created_by: str = '',
    title: Optional[str] = None,
    notes: Optional[str] = None,
    files=None,
    upload_subdir: str = 'orders',
    legacy_table: Optional[str] = None,
    legacy_id: Optional[int] = None,
    conn=None,
) -> int:
    """
    Створює наказ і його позиції.

    items: ітератор словників
      {
        'entity_type': 'student' | 'group',
        'entity_id': int,
        'payload': dict | None,   # довільний контекст
      }

    files: request.files.getlist(...) або None.
    Якщо передано conn — коміт НЕ робиться (викликач комітить сам),
    інакше відкривається власне з'єднання і комітиться тут.

    Повертає id нового наказу.
    """
    if order_type not in ORDER_TYPES:
        raise ValueError(f"Невідомий order_type: {order_type}")
    order_number = (order_number or '').strip()
    order_date = (order_date or '').strip()
    if not order_number or not order_date:
        raise ValueError("order_number і order_date обов'язкові")

    items = list(items or [])
    if not items:
        raise ValueError("Наказ без позицій (items) не створюється")

    own_conn = conn is None
    if own_conn:
        conn = get_db()

    try:
        cur = conn.execute(
            """
            INSERT INTO orders
                (order_type, order_number, order_date, title, notes,
                 created_by, status, legacy_table, legacy_id)
            VALUES (?, ?, ?, ?, ?, ?, 'active', ?, ?)
            """,
            (
                order_type,
                order_number,
                order_date,
                title,
                notes,
                created_by or '',
                legacy_table,
                legacy_id,
            ),
        )
        order_id = cur.lastrowid

        for it in items:
            et = it.get('entity_type')
            eid = it.get('entity_id')
            if et not in ('student', 'group') or not eid:
                raise ValueError(f"Некоректний item: {it}")
            conn.execute(
                """
                INSERT INTO order_items
                    (order_id, entity_type, entity_id, payload)
                VALUES (?, ?, ?, ?)
                """,
                (order_id, et, int(eid), _payload_dumps(it.get('payload'))),
            )

        if files:
            save_multiple_attachments(
                conn, 'order', order_id, files, upload_subdir, created_by or None
            )

        if own_conn:
            conn.commit()
        return order_id
    except Exception:
        if own_conn:
            conn.rollback()
        raise
    finally:
        if own_conn:
            conn.close()


def get_order(order_id: int, conn=None) -> Optional[dict]:
    """Повний наказ: реквізити + items + attachments."""
    own_conn = conn is None
    if own_conn:
        conn = get_db()
    try:
        row = conn.execute(
            "SELECT * FROM orders WHERE id=?", (order_id,)
        ).fetchone()
        if not row:
            return None
        order = dict(row)
        order['type_label'] = ORDER_TYPE_LABELS.get(
            order['order_type'], order['order_type']
        )

        items = conn.execute(
            """
            SELECT id, entity_type, entity_id, payload
            FROM order_items
            WHERE order_id=?
            ORDER BY id
            """,
            (order_id,),
        ).fetchall()
        order['items'] = [
            {
                'id': r['id'],
                'entity_type': r['entity_type'],
                'entity_id': r['entity_id'],
                'payload': _payload_loads(r['payload']),
            }
            for r in items
        ]

        atts = conn.execute(
            """
            SELECT id, file_path, original_name, uploaded_at, uploaded_by
            FROM attachments
            WHERE entity_type='order' AND entity_id=?
            ORDER BY id
            """,
            (order_id,),
        ).fetchall()
        order['attachments'] = [dict(a) for a in atts]
        return order
    finally:
        if own_conn:
            conn.close()


def list_orders(
    *,
    order_type: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    search: Optional[str] = None,
    limit: int = 100,
    offset: int = 0,
    conn=None,
) -> list[dict]:
    """
    Список наказів для реєстру.
    search — підрядок у order_number або title.
    date_from / date_to — рядки DD.MM.YYYY (порівняння лексикографічне
    для дат у цьому форматі некоректне; для фільтра краще нормалізувати
    на стороні викликача або тримати ISO в БД пізніше).
    """
    own_conn = conn is None
    if own_conn:
        conn = get_db()
    try:
        sql = ["SELECT * FROM orders WHERE status='active'"]
        params: list[Any] = []

        if order_type:
            sql.append("AND order_type=?")
            params.append(order_type)
        if date_from:
            sql.append("AND order_date >= ?")
            params.append(date_from)
        if date_to:
            sql.append("AND order_date <= ?")
            params.append(date_to)
        if search:
            sql.append("AND (order_number LIKE ? OR IFNULL(title,'') LIKE ?)")
            q = f"%{search.strip()}%"
            params.extend([q, q])

        sql.append("ORDER BY id DESC LIMIT ? OFFSET ?")
        params.extend([limit, offset])

        rows = conn.execute(" ".join(sql), params).fetchall()
        result = []
        for r in rows:
            d = dict(r)
            d['type_label'] = ORDER_TYPE_LABELS.get(
                d['order_type'], d['order_type']
            )
            cnt = conn.execute(
                "SELECT COUNT(*) FROM order_items WHERE order_id=?",
                (d['id'],),
            ).fetchone()[0]
            d['items_count'] = cnt
            result.append(d)

        # ── Обогащение: license_summary для license_transfer-наказов ──
        # Показываем "Львівська → Київська" прямо в списке наказов,
        # чтобы админу не нужно было открывать каждую карточку.
        license_summary_by_order = {}
        lt_order_ids = [
            d['id'] for d in result if d['order_type'] == 'license_transfer'
        ]
        if lt_order_ids:
            placeholders = ','.join('?' for _ in lt_order_ids)
            items_rows = conn.execute(
                f"SELECT order_id, payload FROM order_items "
                f"WHERE order_id IN ({placeholders})",
                lt_order_ids,
            ).fetchall()

            per_order = {}
            for ir in items_rows:
                payload = _payload_loads(ir['payload'])
                prev_id = payload.get('previous_license_id')
                new_id = payload.get('new_license_id')
                if new_id is None and prev_id is None:
                    continue
                per_order.setdefault(ir['order_id'], set()).add((prev_id, new_id))

            license_ids_needed = set()
            for pairs in per_order.values():
                for p, n in pairs:
                    if p:
                        license_ids_needed.add(int(p))
                    if n:
                        license_ids_needed.add(int(n))

            lic_names = {}
            if license_ids_needed:
                l_ph = ','.join('?' for _ in license_ids_needed)
                for lr in conn.execute(
                    f"SELECT id, short_name_ua, name_ua FROM institution_licenses "
                    f"WHERE id IN ({l_ph})",
                    list(license_ids_needed),
                ).fetchall():
                    lic_names[lr['id']] = lr['short_name_ua'] or lr['name_ua']

            for order_id, pairs in per_order.items():
                prevs = {p for p, _ in pairs if p}
                news = {n for _, n in pairs if n}
                new_str = (
                    '/'.join(lic_names.get(n, f'#{n}') for n in sorted(news))
                    if news else '—'
                )
                if len(prevs) == 1 and len(news) == 1:
                    prev = next(iter(prevs))
                    prev_str = lic_names.get(prev, f'#{prev}')
                    license_summary_by_order[order_id] = f'{prev_str} → {new_str}'
                elif news:
                    license_summary_by_order[order_id] = f'→ {new_str}'
                else:
                    license_summary_by_order[order_id] = ''

        for d in result:
            d['license_summary'] = license_summary_by_order.get(d['id'], '')

        return result
    finally:
        if own_conn:
            conn.close()


def get_student_orders(student_id: int, conn=None) -> list[dict]:
    """
    Таймлайн наказів студента.

    Включає:
      1) прямі order_items entity_type='student'
      2) накази на поточну групу студента
      3) накази на групи з previous_group_id (відрахування / заморозка)
      4) накази на group_id з payload enrollment цього студента
      5) накази на ВСІ групи з enrollments студента (активні + історичні)

    Кожен наказ отримує поле `group_refs` — список груп, яких він
    стосується для цього студента (див. _order_group_refs).
    """
    own_conn = conn is None
    if own_conn:
        conn = get_db()
    try:
        group_ids = set()

        row = conn.execute(
            "SELECT group_id FROM students WHERE id=?", (student_id,)
        ).fetchone()
        if row and row[0]:
            group_ids.add(int(row[0]))

        # ВСІ групи з enrollments студента (активні + graduated/expelled/frozen)
        for r in conn.execute(
            "SELECT DISTINCT group_id FROM student_enrollments "
            "WHERE student_id=? AND group_id IS NOT NULL",
            (student_id,),
        ).fetchall():
            group_ids.add(int(r[0]))

        # Групи з заморозки
        for r in conn.execute(
            "SELECT previous_group_id FROM frozen_students "
            "WHERE student_id=? AND previous_group_id IS NOT NULL",
            (student_id,),
        ).fetchall():
            if r[0]:
                group_ids.add(int(r[0]))

        # Group_id з payload enrollment / expulsion
        for r in conn.execute(
            """
            SELECT oi.payload
            FROM order_items oi
            JOIN orders o ON o.id = oi.order_id
            WHERE oi.entity_type='student' AND oi.entity_id=?
              AND o.status='active' AND oi.payload IS NOT NULL
            """,
            (student_id,),
        ).fetchall():
            pl = _payload_loads(r[0])
            for k in ('group_id', 'previous_group_id', 'from_group_id', 'to_group_id'):
                gid = pl.get(k)
                if gid:
                    try:
                        group_ids.add(int(gid))
                    except (TypeError, ValueError):
                        pass

        # 1) прямі
        direct = conn.execute(
            """
            SELECT o.*, oi.payload AS item_payload, oi.entity_type, oi.entity_id
            FROM orders o
            JOIN order_items oi ON oi.order_id = o.id
            WHERE o.status='active'
              AND oi.entity_type='student'
              AND oi.entity_id=?
            """,
            (student_id,),
        ).fetchall()

        # 2) через групи
        group_orders = []
        if group_ids:
            placeholders = ','.join('?' for _ in group_ids)
            group_orders = conn.execute(
                f"""
                SELECT o.*, oi.payload AS item_payload, oi.entity_type, oi.entity_id
                FROM orders o
                JOIN order_items oi ON oi.order_id = o.id
                WHERE o.status='active'
                  AND oi.entity_type='group'
                  AND oi.entity_id IN ({placeholders})
                """,
                list(group_ids),
            ).fetchall()

        seen = set()
        merged = []
        for r in list(direct) + list(group_orders):
            if r['id'] in seen:
                continue
            seen.add(r['id'])
            d = dict(r)
            d['type_label'] = ORDER_TYPE_LABELS.get(
                d['order_type'], d['order_type']
            )
            d['payload'] = _payload_loads(d.pop('item_payload', None))
            d['payload_text'] = format_payload(conn, d['order_type'], d['payload'])

            # ── Усі групи, яких стосується наказ для цього студента ──
            # SQL-фільтр: лише items, які належать САМЕ цьому студенту,
            # або group items (для масових наказів). Це також швидше —
            # з БД повертається менший результат.
            full_items_rows = conn.execute(
                "SELECT entity_type, entity_id, payload FROM order_items "
                "WHERE order_id=? "
                "  AND (entity_type = 'group' "
                "       OR (entity_type = 'student' AND entity_id = ?))",
                (d['id'], student_id),
            ).fetchall()
            full_items = [
                {'entity_type': ir['entity_type'],
                 'entity_id': ir['entity_id'],
                 'payload': _payload_loads(ir['payload'])}
                for ir in full_items_rows
            ]
            d['group_refs'] = _order_group_refs(
                conn, d['order_type'], full_items, group_ids, student_id
            )

            atts = conn.execute(
                """
                SELECT id, file_path, original_name
                FROM attachments
                WHERE entity_type='order' AND entity_id=?
                ORDER BY id
                """,
                (d['id'],),
            ).fetchall()
            d['attachments'] = [dict(a) for a in atts]
            merged.append(d)

        def sort_key(o):
            od = o.get('order_date') or ''
            parts = od.replace('/', '.').split('.')
            if len(parts) == 3 and len(parts[2]) == 4:
                ymd = parts[2] + parts[1].zfill(2) + parts[0].zfill(2)
            else:
                ymd = od
            return (ymd, o.get('id') or 0)

        merged.sort(key=sort_key)
        return merged
    finally:
        if own_conn:
            conn.close()


def student_has_enrollment(student_id: int, conn=None) -> bool:
    """Чи є активний наказ типу enrollment на цього студента."""
    own_conn = conn is None
    if own_conn:
        conn = get_db()
    try:
        row = conn.execute(
            """
            SELECT 1
            FROM orders o
            JOIN order_items oi ON oi.order_id = o.id
            WHERE o.status='active'
              AND o.order_type='enrollment'
              AND oi.entity_type='student'
              AND oi.entity_id=?
            LIMIT 1
            """,
            (student_id,),
        ).fetchone()
        return row is not None
    finally:
        if own_conn:
            conn.close()


def get_enrollment_student_ids(conn=None) -> set[int]:
    """Множина student_id, у яких уже є enrollment-наказ (для дашборду)."""
    own_conn = conn is None
    if own_conn:
        conn = get_db()
    try:
        rows = conn.execute(
            """
            SELECT DISTINCT oi.entity_id
            FROM orders o
            JOIN order_items oi ON oi.order_id = o.id
            WHERE o.status='active'
              AND o.order_type='enrollment'
              AND oi.entity_type='student'
            """
        ).fetchall()
        return {int(r[0]) for r in rows}
    finally:
        if own_conn:
            conn.close()


def count_students_without_enrollment(
    *,
    student_ids: Optional[Iterable[int]] = None,
    conn=None,
) -> int:
    """
    Скільки студентів з переданого набору (або всіх активних) не мають
    enrollment-наказу. Зручно для «Потребують уваги».
    """
    own_conn = conn is None
    if own_conn:
        conn = get_db()
    try:
        enrolled = get_enrollment_student_ids(conn=conn)
        if student_ids is not None:
            return sum(1 for sid in student_ids if sid not in enrolled)

        rows = conn.execute(
            """
            SELECT id FROM students
            WHERE COALESCE(archived, 0) = 0
            """
        ).fetchall()
        return sum(1 for r in rows if r[0] not in enrolled)
    finally:
        if own_conn:
            conn.close()


def orders_tables_ready(conn=None) -> bool:
    """Чи існують таблиці orders (після migrate_orders.py / init_db)."""
    own_conn = conn is None
    if own_conn:
        conn = get_db()
    try:
        row = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='orders'"
        ).fetchone()
        return row is not None
    finally:
        if own_conn:
            conn.close()


# ── Хелпери для підключення до старих confirm-ендпоінтів ──────────────


def move_order_item(item_id: int, target_order_id: int, conn=None) -> tuple:
    """
    Переносить order_item в інший наказ.
    Повертає (ok: bool, message: str).
    """
    own = conn is None
    if own:
        from routes.db import get_db
        conn = get_db()
    try:
        item = conn.execute(
            """
            SELECT oi.id, oi.order_id, oi.entity_type, oi.entity_id,
                   o.order_type, o.order_number
            FROM order_items oi
            JOIN orders o ON o.id = oi.order_id
            WHERE oi.id = ?
            """,
            (item_id,),
        ).fetchone()
        if not item:
            return False, 'Пункт наказу не знайдено'

        target = conn.execute(
            "SELECT id, order_type, order_number FROM orders WHERE id = ?",
            (target_order_id,),
        ).fetchone()
        if not target:
            return False, 'Цільовий наказ не знайдено'

        # Row access
        try:
            src_type = item['order_type']
            tgt_type = target['order_type']
            src_order_id = item['order_id']
            tgt_num = target['order_number']
        except Exception:
            src_type = item[4]
            tgt_type = target[1]
            src_order_id = item[1]
            tgt_num = target[2]

        if src_order_id == target_order_id:
            return False, 'Студент уже в цьому наказі'

        if src_type != tgt_type:
            return False, (
                f'Типи наказів різні ({src_type} → {tgt_type}). '
                'Переносити можна лише між наказами одного типу.'
            )

        # Уникаємо дубля студента в цільовому наказі
        try:
            entity_type, entity_id = item['entity_type'], item['entity_id']
        except Exception:
            entity_type, entity_id = item[2], item[3]

        dup = conn.execute(
            """
            SELECT id FROM order_items
            WHERE order_id = ? AND entity_type = ? AND entity_id = ?
            """,
            (target_order_id, entity_type, entity_id),
        ).fetchone()
        if dup:
            return False, f'Цей студент уже є в наказі № {tgt_num}'

        conn.execute(
            "UPDATE order_items SET order_id = ? WHERE id = ?",
            (target_order_id, item_id),
        )
        if own:
            conn.commit()
        return True, f'Перенесено в наказ № {tgt_num}'
    finally:
        if own:
            conn.close()


def delete_order(order_id: int, conn=None) -> bool:

    """
    Повністю видаляє наказ і всі order_items.
    Повертає True, якщо наказ існував і видалений.
    """
    own = conn is None
    if own:
        from routes.db import get_db
        conn = get_db()
    try:
        row = conn.execute("SELECT id, order_number, order_type FROM orders WHERE id = ?", (order_id,)).fetchone()
        if not row:
            return False
        conn.execute("DELETE FROM order_items WHERE order_id = ?", (order_id,))
        conn.execute("DELETE FROM orders WHERE id = ?", (order_id,))
        if own:
            conn.commit()
        return True
    finally:
        if own:
            conn.close()


def mirror_course_transfer_order(
    conn,
    *,
    legacy_id: int,
    order_number: str,
    order_date: str,
    group_items: list[dict],
    created_by: str,
    files=None,
) -> Optional[int]:
    """
    Викликати в кінці course_transfer_confirm після INSERT у
    course_transfer_orders. group_items: [{group_id, course_from, course_to}, ...]
    """
    if not orders_tables_ready(conn=conn):
        return None
    items = [
        {
            'entity_type': 'group',
            'entity_id': g['group_id'],
            'payload': {
                'course_from': g['course_from'],
                'course_to': g['course_to'],
            },
        }
        for g in group_items
    ]
    return create_order(
        'course_transfer',
        order_number,
        order_date,
        items,
        created_by=created_by,
        files=files,
        upload_subdir='orders',
        legacy_table='course_transfer_orders',
        legacy_id=legacy_id,
        conn=conn,
    )


def mirror_expulsion_order(
    conn,
    *,
    legacy_id: int,
    order_number: str,
    order_date: str,
    student_items: list[dict],
    created_by: str,
    files=None,
) -> Optional[int]:
    """
    student_items: [{student_id, previous_group_id, reason}, ...]
    """
    if not orders_tables_ready(conn=conn):
        return None
    items = [
        {
            'entity_type': 'student',
            'entity_id': s['student_id'],
            'payload': {
                'previous_group_id': s.get('previous_group_id'),
                'reason': s.get('reason') or '',
            },
        }
        for s in student_items
    ]
    return create_order(
        'expulsion',
        order_number,
        order_date,
        items,
        created_by=created_by,
        files=files,
        upload_subdir='orders',
        legacy_table='expulsion_orders',
        legacy_id=legacy_id,
        conn=conn,
    )


def mirror_license_transfer_order(
    conn,
    *,
    legacy_id: int,
    order_number: str,
    order_date: str,
    student_items: list[dict],
    created_by: str,
    files=None,
) -> Optional[int]:
    """
    student_items: [{student_id, previous_license_id, new_license_id}, ...]
    """
    if not orders_tables_ready(conn=conn):
        return None
    items = [
        {
            'entity_type': 'student',
            'entity_id': s['student_id'],
            'payload': {
                'previous_license_id': s.get('previous_license_id'),
                'new_license_id': s.get('new_license_id'),
            },
        }
        for s in student_items
    ]
    return create_order(
        'license_transfer',
        order_number,
        order_date,
        items,
        created_by=created_by,
        files=files,
        upload_subdir='orders',
        legacy_table='license_transfer_orders',
        legacy_id=legacy_id,
        conn=conn,
    )

def get_enrollment_pairs(conn=None) -> set:
    """
    Множина пар (student_id, group_id), для яких уже є активний наказ
    типу 'enrollment'. Використовується для перевірки, чи в конкретній
    ГРУПІ студента вже оформлено зарахування (а не «студент має якийсь
    наказ десь», як у get_enrollment_student_ids).

    Потрібно для паралельного навчання: якщо студент вчиться у групах
    A і B, наказ для A не повинен блокувати створення наказу для B.
    """
    own = conn is None
    if own:
        conn = get_db()
    try:
        rows = conn.execute("""
            SELECT oi.entity_id AS student_id, oi.payload
            FROM orders o
            JOIN order_items oi ON oi.order_id = o.id
            WHERE o.status = 'active'
              AND o.order_type = 'enrollment'
              AND oi.entity_type = 'student'
        """).fetchall()
        result = set()
        for r in rows:
            try:
                payload = json.loads(r['payload'] or '{}')
                gid = payload.get('group_id')
                if gid:
                    result.add((int(r['student_id']), int(gid)))
            except (ValueError, TypeError):
                continue
        return result
    finally:
        if own:
            conn.close()


def get_active_groups_for_student(cursor, student_id: int) -> list:
    """
    Повертає список активних груп студента:
        [(group_id, group_name, is_primary), ...]
    Primary йде першим, далі паралельні за порядком створення enrollment'а.

    Legacy-fallback: студент без enrollment'ів, але з group_id —
    повертає одну групу.
    """
    rows = cursor.execute("""
        SELECT e.group_id, g.name, e.is_primary
        FROM student_enrollments e
        JOIN groups g ON g.id = e.group_id
        WHERE e.student_id = ? AND e.status = 'active'
        ORDER BY e.is_primary DESC, e.id
    """, (student_id,)).fetchall()

    if rows:
        return [(r['group_id'], r['name'], bool(r['is_primary'])) for r in rows]

    row = cursor.execute("""
        SELECT s.group_id, g.name
        FROM students s
        LEFT JOIN groups g ON g.id = s.group_id
        WHERE s.id = ?
    """, (student_id,)).fetchone()
    if row and row['group_id']:
        return [(row['group_id'], row['name'], True)]
    return []
    
def _order_group_refs(conn, order_type: str, items: list,
                      known_group_ids: set, student_id: int) -> list:
    """
    Повертає список груп, яких стосується наказ САМЕ ДЛЯ ЦЬОГО студента.

    ВАЖЛИВО: беремо лише ті `items`, де entity_type='student' і
    entity_id == student_id (для паралельного студента в наказі є
    items й інших студентів — їхні payload.group_id ігноруємо), та
    entity_type='group' (для масових наказів на кшталт course_transfer).
    """
    refs, seen = [], set()

    def add(gid, role=None, check_known=True):
        if not gid or gid in seen:
            return
        try:
            gid = int(gid)
        except (TypeError, ValueError):
            return
        if check_known and gid not in known_group_ids:
            return
        row = conn.execute(
            "SELECT name, start_year, study_form FROM groups WHERE id=?", (gid,)
        ).fetchone()
        if not row:
            return
        seen.add(gid)
        refs.append({
            'id': gid,
            'name': row['name'],
            'start_year': row['start_year'],
            'study_form': row['study_form'],
            'role': role,
        })

    for it in items:
        et = it.get('entity_type')
        eid = it.get('entity_id')
        payload = it.get('payload') or {}

        if et == 'group':
            add(eid, order_type)

        # ── КЛЮЧОВЕ ВИПРАВЛЕННЯ ──
        # Беремо лише items, які належать САМЕ ЦЬОМУ студентові.
        if et == 'student' and eid == student_id:
            if order_type == 'enrollment':
                add(payload.get('group_id'), 'enrollment')
            elif order_type == 'expulsion':
                add(payload.get('previous_group_id'), 'expulsion')
            elif order_type == 'license_transfer':
                add(payload.get('group_id'), 'license')
            elif order_type == 'graduation':
                add(payload.get('group_id'), 'graduation')
            elif order_type in ('group_transfer', 'study_form_transfer'):
                add(payload.get('from_group_id'), 'from', check_known=False)
                add(payload.get('to_group_id'),   'to',   check_known=False)

    return refs
    
def _order_group_id_for_dedup(order_type: str, payload: dict):
    """
    Group_id, який використовується для дедуплікації items у наказі.

    Для enrollment/expulsion/graduation/license_transfer/freeze — це
    поле з payload. Для інших типів (course_transfer) — None.

    Використовується, щоб:
      - студент з паралельним навчанням міг мати 2 окремі позиції
        (для КН-21 і для МГ-1) — бо це РІЗНІ групи;
      - справжній дублікат (той самий студент у тій самій групі)
        згортався в одну позицію.
    """
    if not payload:
        return None
    field_map = {
        'enrollment':          'group_id',
        'expulsion':           'previous_group_id',
        'graduation':          'group_id',
        'license_transfer':    'group_id',
        'freeze':              'previous_group_id',
        'group_transfer':      'from_group_id',
        'study_form_transfer': 'from_group_id',
    }
    field = field_map.get(order_type)
    if not field:
        return None
    gid = payload.get(field)
    if not gid:
        return None
    try:
        return int(gid)
    except (TypeError, ValueError):
        return None    
    
def merge_orders(
    target_id: int,
    source_ids: list[int],
    *,
    new_number: str | None = None,
    new_date: str | None = None,
    new_notes: str | None = None,
    conn=None,
) -> dict:
    """
    Об'єднує накази: переносить order_items та attachments з source_ids
    у target_id, оновлює реквізити target (опційно) та видаляє source.

    Обмеження:
      - усі накази мають бути 'active';
      - усі мають бути одного order_type;
      - target має бути серед обраних.

    Дублікати (та сама пара entity_type + entity_id, яка вже є в
    target) — пропускаються (не переносяться), збільшують лічильник
    skipped_duplicates.

    Повертає dict:
        {ok, error?, moved_items, skipped_duplicates,
         moved_attachments, order_type, main_number, main_date}
    """
    own = conn is None
    if own:
        conn = get_db()
    try:
        source_ids = [int(x) for x in source_ids if int(x) != target_id]
        if not source_ids:
            return {'ok': False, 'error': 'Немає наказів для об\'єднання'}

        all_ids = [target_id] + source_ids
        ph = ','.join('?' * len(all_ids))
        rows = conn.execute(
            f"SELECT id, order_type, status, order_number, order_date "
            f"FROM orders WHERE id IN ({ph})",
            all_ids,
        ).fetchall()

        if len(rows) != len(all_ids):
            return {'ok': False, 'error': 'Деякі накази не знайдено'}

        by_id = {r['id']: r for r in rows}
        target_row = by_id[target_id]
        target_type = target_row['order_type']

        for sid in all_ids:
            r = by_id[sid]
            if r['status'] != 'active':
                return {
                    'ok': False,
                    'error': f"Наказ #{sid} не активний",
                }
            if r['order_type'] != target_type:
                return {
                    'ok': False,
                    'error': (
                        f"Наказ #{sid} іншого типу "
                        f"({r['order_type']} ≠ {target_type})"
                    ),
                }

        # ── Позиції: дедуплікація за (entity_type, entity_id, group_id) ──
        existing_pairs = set()
        for r in conn.execute(
            "SELECT entity_type, entity_id, payload FROM order_items WHERE order_id=?",
            (target_id,),
        ).fetchall():
            payload = _payload_loads(r['payload'])
            gid = (
                _order_group_id_for_dedup(target_type, payload)
                if r['entity_type'] == 'student' else None
            )
            existing_pairs.add((r['entity_type'], r['entity_id'], gid))

        moved_items = 0
        skipped_duplicates = 0

        for src_id in source_ids:
            items = conn.execute(
                "SELECT id, entity_type, entity_id, payload FROM order_items WHERE order_id=?",
                (src_id,),
            ).fetchall()
            for item in items:
                payload = _payload_loads(item['payload'])
                gid = (
                    _order_group_id_for_dedup(target_type, payload)
                    if item['entity_type'] == 'student' else None
                )
                pair = (item['entity_type'], item['entity_id'], gid)
                if pair in existing_pairs:
                    skipped_duplicates += 1
                    continue
                conn.execute(
                    "UPDATE order_items SET order_id=? WHERE id=?",
                    (target_id, item['id']),
                )
                existing_pairs.add(pair)
                moved_items += 1

        # ── Скани наказу ──
        moved_atts = 0
        for src_id in source_ids:
            cur = conn.execute(
                "UPDATE attachments SET entity_id=? "
                "WHERE entity_type='order' AND entity_id=?",
                (target_id, src_id),
            )
            moved_atts += cur.rowcount or 0

        # ── Переприв'язка посилань на order_id ──
        for src_id in source_ids:
            for tbl in ('frozen_students',
                        'student_enrollments',
                        'student_name_history'):
                try:
                    conn.execute(
                        f"UPDATE {tbl} SET order_id=? WHERE order_id=?",
                        (target_id, src_id),
                    )
                except Exception:
                    # Таблиця може не існувати у старій БД — не критично
                    pass

        # ── Оновлення реквізитів target ──
        updates = []
        params = []
        if new_number:
            updates.append("order_number=?")
            params.append(new_number)
        if new_date:
            updates.append("order_date=?")
            params.append(new_date)
        if new_notes is not None:
            updates.append("notes=?")
            params.append(new_notes)
        if updates:
            params.append(target_id)
            conn.execute(
                f"UPDATE orders SET {', '.join(updates)} WHERE id=?",
                params,
            )

        # ── Видалення source ──
        for src_id in source_ids:
            conn.execute("DELETE FROM order_items WHERE order_id=?", (src_id,))
            conn.execute("DELETE FROM orders WHERE id=?", (src_id,))

        if own:
            conn.commit()
        return {
            'ok': True,
            'moved_items': moved_items,
            'skipped_duplicates': skipped_duplicates,
            'moved_attachments': moved_atts,
            'order_type': target_type,
            'main_number': new_number or target_row['order_number'],
            'main_date': new_date or target_row['order_date'],
        }
    except Exception as e:
        if own:
            conn.rollback()
        return {'ok': False, 'error': str(e)}
    finally:
        if own:
            conn.close()