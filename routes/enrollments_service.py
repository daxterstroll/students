"""
Паралельне навчання: один студент — кілька груп (enrollment).
students.group_id = primary (для сумісності зі старим кодом).
student_enrollments.license_id = ліцензія, під якою студент вчиться
в конкретній групі (може відрізнятися від інших груп того ж студента).
student_enrollments.edebo_code = код ЄДЕБО, під яким студент зарахований
у конкретній групі (так само може відрізнятися між групами).
"""
from __future__ import annotations

from typing import Optional

from routes.db import get_db


def ensure_enrollments_table(conn=None) -> None:
    """CREATE TABLE IF NOT EXISTS + безпечне додавання колонок license_id / edebo_code."""
    own = conn is None
    if own:
        conn = get_db()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS student_enrollments (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id  INTEGER NOT NULL,
                group_id    INTEGER NOT NULL,
                is_primary  INTEGER NOT NULL DEFAULT 0,
                status      TEXT NOT NULL DEFAULT 'active',
                license_id  INTEGER,
                edebo_code  TEXT,
                order_id    INTEGER,
                is_correction INTEGER NOT NULL DEFAULT 0,
                correction_note TEXT,
                correction_by TEXT,
                correction_at TEXT,
                notes       TEXT,
                created_at  TEXT DEFAULT (datetime('now', 'localtime')),
                UNIQUE (student_id, group_id),
                FOREIGN KEY (student_id) REFERENCES students(id),
                FOREIGN KEY (group_id) REFERENCES groups(id),
                FOREIGN KEY (license_id) REFERENCES institution_licenses(id)
            )
        """)
        # На випадок, якщо таблиця вже існує без цих колонок
        try:
            conn.execute("ALTER TABLE student_enrollments ADD COLUMN license_id INTEGER")
        except Exception:
            pass
        try:
            conn.execute("ALTER TABLE student_enrollments ADD COLUMN edebo_code TEXT")
        except Exception:
            pass
        for _col, _type in (
            ('is_correction', 'INTEGER NOT NULL DEFAULT 0'),
            ('correction_note', 'TEXT'),
            ('correction_by', 'TEXT'),
            ('correction_at', 'TEXT'),
        ):
            try:
                conn.execute(f"ALTER TABLE student_enrollments ADD COLUMN {_col} {_type}")
            except Exception:
                pass
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_enrollments_student "
            "ON student_enrollments(student_id)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_enrollments_group "
            "ON student_enrollments(group_id)"
        )
        if own:
            conn.commit()
    finally:
        if own:
            conn.close()


def backfill_primary_from_students(conn=None) -> int:
    """Для студентів з group_id, яких ще немає в student_enrollments,
    додає primary-рядок (з копією students.license_id та students.edebo_code)."""
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        cur = conn.execute("""
            INSERT OR IGNORE INTO student_enrollments
                (student_id, group_id, is_primary, status, license_id, edebo_code)
            SELECT id, group_id, 1, 'active', license_id, edebo_code
            FROM students
            WHERE group_id IS NOT NULL
        """)
        if own:
            conn.commit()
        return cur.rowcount or 0
    finally:
        if own:
            conn.close()


def list_student_enrollments(student_id: int, conn=None) -> list:
    """Повертає ВСІ enrollments студента з назвами групи та ліцензії.

    license_id / edebo_code / назви беруться з enrollment, а якщо там NULL
    (студент створений до впровадження per-enrollment значень і міграція
    ще не пройшла) — fallback на students.license_id / students.edebo_code.
    Так карточка ніколи не показує «не вказано», якщо на рівні студента є.
    """
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        rows = conn.execute("""
            SELECT e.id, e.student_id, e.group_id, e.is_primary, e.status,
                   COALESCE(e.license_id, s.license_id) AS license_id,
                   COALESCE(e.edebo_code, s.edebo_code) AS edebo_code,
                   e.order_id, e.notes, e.created_at,
                   e.is_correction, e.correction_note, e.correction_by, e.correction_at,
                   g.name AS group_name, g.start_year, g.study_form,
                   g.degree_level, g.course, g.specialty_code,
                   g.program_credits,
                   il.short_name_ua AS license_short_name_ua,
                   il.name_ua       AS license_name_ua,
                   il.short_name_en AS license_short_name_en,
                   il.name_en       AS license_name_en,
                   e.license_id IS NULL AS license_is_inherited,
                   e.edebo_code IS NULL AS edebo_is_inherited
            FROM student_enrollments e
            JOIN students s ON s.id = e.student_id
            LEFT JOIN groups g ON g.id = e.group_id
            LEFT JOIN institution_licenses il
                   ON il.id = COALESCE(e.license_id, s.license_id)
            WHERE e.student_id = ?
            ORDER BY e.is_primary DESC,
                     CASE e.status WHEN 'active' THEN 0 ELSE 1 END,
                     e.id
        """, (student_id,)).fetchall()

        _seen = set()
        _deduped = []
        for r in rows:
            key = (r['student_id'], r['group_id'])
            if key in _seen:
                continue
            _seen.add(key)
            _deduped.append(dict(r))
        return _deduped
    finally:
        if own:
            conn.close()


def get_active_group_ids_for_student(student_id: int, conn=None) -> list:
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        rows = conn.execute("""
            SELECT group_id FROM student_enrollments
            WHERE student_id = ? AND status = 'active'
        """, (student_id,)).fetchall()
        ids = [r['group_id'] if hasattr(r, 'keys') else r[0] for r in rows]
        if not ids:
            row = conn.execute(
                "SELECT group_id FROM students WHERE id = ?", (student_id,)
            ).fetchone()
            gid = None
            if row is not None:
                gid = row['group_id'] if hasattr(row, 'keys') else row[0]
            if gid:
                ids = [gid]
        return ids
    finally:
        if own:
            conn.close()


def student_ids_in_group(group_id: int, include_inactive: bool = False, conn=None) -> set:
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        sql = "SELECT student_id FROM student_enrollments WHERE group_id = ?"
        if not include_inactive:
            sql += " AND status = 'active'"
        ids = {
            (r['student_id'] if hasattr(r, 'keys') else r[0])
            for r in conn.execute(sql, (group_id,)).fetchall()
        }
        extra = conn.execute("""
            SELECT id FROM students
            WHERE group_id = ?
              AND id NOT IN (
                  SELECT student_id FROM student_enrollments WHERE group_id = ?
              )
        """, (group_id, group_id)).fetchall()
        for r in extra:
            ids.add(r['id'] if hasattr(r, 'keys') else r[0])
        return ids
    finally:
        if own:
            conn.close()


def add_enrollment(
    student_id: int,
    group_id: int,
    *,
    is_primary: bool = False,
    license_id: Optional[int] = None,
    edebo_code: Optional[str] = None,
    order_id=None,
    notes=None,
    status: str = 'active',
    conn=None,
):
    """Додає навчання в групу. Повертає (ok: bool, message: str)."""
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)

        st = conn.execute(
            "SELECT id, license_id, edebo_code FROM students WHERE id = ?", (student_id,)
        ).fetchone()
        if not st:
            return False, 'Студента не знайдено'

        gr = conn.execute(
            "SELECT id FROM groups WHERE id = ?", (group_id,)
        ).fetchone()
        if not gr:
            return False, 'Групу не знайдено'

        # Якщо license_id не передано — успадковуємо від студента
        if license_id is None:
            license_id = st['license_id'] if hasattr(st, 'keys') else st[1]
            if license_id is None:
                base_row = conn.execute(
                    "SELECT id FROM institution_licenses WHERE is_base = 1 LIMIT 1"
                ).fetchone()
                if base_row:
                    license_id = base_row[0] if not hasattr(base_row, 'keys') else base_row['id']

        # Якщо edebo_code не передано — успадковуємо від студента
        if edebo_code is None:
            edebo_code = st['edebo_code'] if hasattr(st, 'keys') else st[2]

        existing = conn.execute(
            """
            SELECT id, status FROM student_enrollments
            WHERE student_id = ? AND group_id = ?
            """,
            (student_id, group_id),
        ).fetchone()

        if existing:
            eid = existing['id'] if hasattr(existing, 'keys') else existing[0]
            est = existing['status'] if hasattr(existing, 'keys') else existing[1]
            if est == 'active':
                return False, 'Студент уже на навчанні в цій групі'
            conn.execute(
                """
                UPDATE student_enrollments
                SET status = ?, order_id = COALESCE(?, order_id),
                    notes = COALESCE(?, notes), is_primary = ?,
                    license_id = COALESCE(?, license_id),
                    edebo_code = COALESCE(?, edebo_code)
                WHERE id = ?
                """,
                (status, order_id, notes, 1 if is_primary else 0,
                 license_id, edebo_code, eid),
            )
        else:
            conn.execute(
                """
                INSERT INTO student_enrollments
                    (student_id, group_id, is_primary, status,
                     license_id, edebo_code, order_id, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (student_id, group_id, 1 if is_primary else 0, status,
                 license_id, edebo_code, order_id, notes),
            )

        if is_primary:
            conn.execute(
                """
                UPDATE student_enrollments SET is_primary = 0
                WHERE student_id = ? AND group_id != ?
                """,
                (student_id, group_id),
            )
            conn.execute(
                "UPDATE students SET group_id = ?, "
                "license_id = COALESCE(?, license_id), "
                "edebo_code = COALESCE(?, edebo_code) "
                "WHERE id = ?",
                (group_id, license_id, edebo_code, student_id),
            )

        conn.execute("""
            UPDATE students SET archived = FALSE
            WHERE id = ? AND archived = TRUE
              AND EXISTS (SELECT 1 FROM student_enrollments
                          WHERE student_id = ? AND status = 'active')
        """, (student_id, student_id))

        if own:
            conn.commit()
        return True, 'Навчання додано'
    finally:
        if own:
            conn.close()


def update_enrollment_license(student_id: int, group_id: int,
                               license_id: Optional[int], conn=None):
    """Змінює ліцензію конкретного enrollment. Якщо це primary —
    синхронізує students.license_id."""
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        row = conn.execute(
            "SELECT id, is_primary FROM student_enrollments "
            "WHERE student_id = ? AND group_id = ?",
            (student_id, group_id),
        ).fetchone()
        if not row:
            return False, 'Запис про навчання в цій групі не знайдено'
        eid = row['id'] if hasattr(row, 'keys') else row[0]
        is_primary = bool(row['is_primary'] if hasattr(row, 'keys') else row[1])

        conn.execute(
            "UPDATE student_enrollments SET license_id = ? WHERE id = ?",
            (license_id, eid),
        )
        if is_primary:
            conn.execute(
                "UPDATE students SET license_id = ? WHERE id = ?",
                (license_id, student_id),
            )
        if own:
            conn.commit()
        return True, 'Ліцензію оновлено'
    finally:
        if own:
            conn.close()


def update_enrollment_edebo_code(student_id: int, group_id: int,
                                  edebo_code: Optional[str], conn=None):
    """Змінює код ЄДЕБО конкретного enrollment. Якщо це primary —
    синхронізує students.edebo_code."""
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        row = conn.execute(
            "SELECT id, is_primary FROM student_enrollments "
            "WHERE student_id = ? AND group_id = ?",
            (student_id, group_id),
        ).fetchone()
        if not row:
            return False, 'Запис про навчання в цій групі не знайдено'
        eid = row['id'] if hasattr(row, 'keys') else row[0]
        is_primary = bool(row['is_primary'] if hasattr(row, 'keys') else row[1])

        conn.execute(
            "UPDATE student_enrollments SET edebo_code = ? WHERE id = ?",
            (edebo_code, eid),
        )
        if is_primary:
            conn.execute(
                "UPDATE students SET edebo_code = ? WHERE id = ?",
                (edebo_code, student_id),
            )
        if own:
            conn.commit()
        return True, 'Код ЄДЕБО оновлено'
    finally:
        if own:
            conn.close()


def set_primary_enrollment(student_id: int, group_id: int, conn=None):
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        row = conn.execute(
            """
            SELECT id, license_id, edebo_code FROM student_enrollments
            WHERE student_id = ? AND group_id = ? AND status = 'active'
            """,
            (student_id, group_id),
        ).fetchone()
        if not row:
            return False, 'Немає активного навчання в цій групі'

        lic = row['license_id'] if hasattr(row, 'keys') else row[1]
        edc = row['edebo_code'] if hasattr(row, 'keys') else row[2]

        conn.execute(
            "UPDATE student_enrollments SET is_primary = 0 WHERE student_id = ?",
            (student_id,),
        )
        conn.execute(
            """
            UPDATE student_enrollments SET is_primary = 1
            WHERE student_id = ? AND group_id = ?
            """,
            (student_id, group_id),
        )
        conn.execute(
            "UPDATE students SET group_id = ?, "
            "license_id = COALESCE(?, license_id), "
            "edebo_code = COALESCE(?, edebo_code) WHERE id = ?",
            (group_id, lic, edc, student_id),
        )
        if own:
            conn.commit()
        return True, 'Основне навчання оновлено'
    finally:
        if own:
            conn.close()


def end_enrollment(student_id: int, group_id: int, *, status: str = 'completed', conn=None):
    """status: completed | expelled"""
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        row = conn.execute(
            """
            SELECT id, is_primary FROM student_enrollments
            WHERE student_id = ? AND group_id = ? AND status = 'active'
            """,
            (student_id, group_id),
        ).fetchone()
        if not row:
            return False, 'Активного навчання в цій групі немає'

        eid = row['id'] if hasattr(row, 'keys') else row[0]
        was_primary = bool(row['is_primary'] if hasattr(row, 'keys') else row[1])

        conn.execute(
            "UPDATE student_enrollments SET status = ?, is_primary = 0 WHERE id = ?",
            (status, eid),
        )

        if was_primary:
            nxt = conn.execute(
                """
                SELECT group_id, license_id, edebo_code FROM student_enrollments
                WHERE student_id = ? AND status = 'active'
                ORDER BY id LIMIT 1
                """,
                (student_id,),
            ).fetchone()
            if nxt:
                new_gid = nxt['group_id'] if hasattr(nxt, 'keys') else nxt[0]
                new_lic = nxt['license_id'] if hasattr(nxt, 'keys') else nxt[1]
                new_edc = nxt['edebo_code'] if hasattr(nxt, 'keys') else nxt[2]
                conn.execute(
                    """
                    UPDATE student_enrollments SET is_primary = 1
                    WHERE student_id = ? AND group_id = ?
                    """,
                    (student_id, new_gid),
                )
                conn.execute(
                    "UPDATE students SET group_id = ?, "
                    "license_id = COALESCE(?, license_id), "
                    "edebo_code = COALESCE(?, edebo_code) WHERE id = ?",
                    (new_gid, new_lic, new_edc, student_id),
                )
            else:
                conn.execute(
                    "UPDATE students SET group_id = NULL WHERE id = ?",
                    (student_id,),
                )

        conn.execute("""
            UPDATE students SET archived = TRUE
            WHERE id = ? AND archived = FALSE
              AND NOT EXISTS (SELECT 1 FROM student_enrollments
                              WHERE student_id = ? AND status = 'active')
        """, (student_id, student_id))

        if own:
            conn.commit()
        return True, 'Навчання завершено'
    finally:
        if own:
            conn.close()


def get_student_licenses(student_id: int, conn=None) -> set:
    """Множина license_id усіх АКТИВНИХ enrollment'ів студента."""
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        rows = conn.execute("""
            SELECT DISTINCT license_id FROM student_enrollments
            WHERE student_id = ? AND status = 'active' AND license_id IS NOT NULL
        """, (student_id,)).fetchall()
        return {(r['license_id'] if hasattr(r, 'keys') else r[0]) for r in rows}
    finally:
        if own:
            conn.close()


def student_has_license(student_id: int, license_id: int, conn=None) -> bool:
    if not license_id:
        return False
    return license_id in get_student_licenses(student_id, conn=conn)


def student_has_base_license(student_id: int, base_license_id: int,
                              conn=None) -> bool:
    lic_ids = get_student_licenses(student_id, conn=conn)
    if not lic_ids:
        return True
    return base_license_id in lic_ids


def list_student_enrollments_without_license(conn=None) -> list:
    """Активні enrollment'и без license_id — для ручного виправлення
    адміністратором. Використовується на сторінках звітності та на
    списку студентів як окремий фільтр «без ліцензії».
    """
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        rows = conn.execute("""
            SELECT e.id, e.student_id, e.group_id, e.status,
                   s.last_name_UA, s.first_name_UA, s.middle_name_UA,
                   g.name AS group_name
            FROM student_enrollments e
            JOIN students s ON s.id = e.student_id
            JOIN groups g ON g.id = e.group_id
            WHERE e.status = 'active' AND e.license_id IS NULL
              AND COALESCE(s.archived, 0) = 0
            ORDER BY s.last_name_UA COLLATE UKRAINIAN, s.first_name_UA
        """).fetchall()
        return [dict(r) for r in rows]
    finally:
        if own:
            conn.close()