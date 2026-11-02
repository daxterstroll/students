"""
Паралельне навчання: один студент — кілька груп (enrollment).
students.group_id = primary (для сумісності зі старим кодом).
"""
from __future__ import annotations

from typing import Optional

from routes.db import get_db


def ensure_enrollments_table(conn=None) -> None:
    """CREATE TABLE IF NOT EXISTS — безпечно, якщо таблиця вже є."""
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
                order_id    INTEGER,
                notes       TEXT,
                created_at  TEXT DEFAULT (datetime('now', 'localtime')),
                UNIQUE (student_id, group_id),
                FOREIGN KEY (student_id) REFERENCES students(id),
                FOREIGN KEY (group_id) REFERENCES groups(id)
            )
        """)
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
    """
    Для студентів з group_id, яких ще немає в student_enrollments —
    додає primary-рядок. Якщо вже заповнювали вручну — просто 0.
    """
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        cur = conn.execute("""
            INSERT OR IGNORE INTO student_enrollments
                (student_id, group_id, is_primary, status)
            SELECT id, group_id, 1, 'active'
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
    """Повертає ВСІ enrollments студента (включно з завершеними),
    ordered: primary first, потім active, потім решта."""
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        rows = conn.execute("""
            SELECT e.id, e.student_id, e.group_id, e.is_primary, e.status,
                   e.order_id, e.notes, e.created_at,
                   g.name AS group_name, g.start_year, g.study_form,
                   g.degree_level, g.course, g.specialty_code,
                   g.program_credits
            FROM student_enrollments e
            LEFT JOIN groups g ON g.id = e.group_id
            WHERE e.student_id = ?
            ORDER BY e.is_primary DESC,
                     CASE e.status WHEN 'active' THEN 0 ELSE 1 END,
                     e.id
        """, (student_id,)).fetchall()
        return [dict(r) for r in rows]
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
            "SELECT id FROM students WHERE id = ?", (student_id,)
        ).fetchone()
        if not st:
            return False, 'Студента не знайдено'

        gr = conn.execute(
            "SELECT id FROM groups WHERE id = ?", (group_id,)
        ).fetchone()
        if not gr:
            return False, 'Групу не знайдено'

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
                    notes = COALESCE(?, notes), is_primary = ?
                WHERE id = ?
                """,
                (status, order_id, notes, 1 if is_primary else 0, eid),
            )
        else:
            conn.execute(
                """
                INSERT INTO student_enrollments
                    (student_id, group_id, is_primary, status, order_id, notes)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (student_id, group_id, 1 if is_primary else 0, status, order_id, notes),
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
                "UPDATE students SET group_id = ? WHERE id = ?",
                (group_id, student_id),
            )

        # ── Якщо в студента є хоч один active enrollment — знімаємо з архіву ──
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


def set_primary_enrollment(student_id: int, group_id: int, conn=None):
    own = conn is None
    if own:
        conn = get_db()
    try:
        ensure_enrollments_table(conn)
        row = conn.execute(
            """
            SELECT id FROM student_enrollments
            WHERE student_id = ? AND group_id = ? AND status = 'active'
            """,
            (student_id, group_id),
        ).fetchone()
        if not row:
            return False, 'Немає активного навчання в цій групі'

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
            "UPDATE students SET group_id = ? WHERE id = ?",
            (group_id, student_id),
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
                SELECT group_id FROM student_enrollments
                WHERE student_id = ? AND status = 'active'
                ORDER BY id LIMIT 1
                """,
                (student_id,),
            ).fetchone()
            if nxt:
                new_gid = nxt['group_id'] if hasattr(nxt, 'keys') else nxt[0]
                conn.execute(
                    """
                    UPDATE student_enrollments SET is_primary = 1
                    WHERE student_id = ? AND group_id = ?
                    """,
                    (student_id, new_gid),
                )
                conn.execute(
                    "UPDATE students SET group_id = ? WHERE id = ?",
                    (new_gid, student_id),
                )
            else:
                conn.execute(
                    "UPDATE students SET group_id = NULL WHERE id = ?",
                    (student_id,),
                )

        # ── Якщо більше немає active enrollment'ів — архівуємо студента ──
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