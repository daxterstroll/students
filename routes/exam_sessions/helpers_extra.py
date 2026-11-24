"""
Допоміжні утиліти для модуля сесій:
- ідемпотентне розширення схеми БД (перездачі, закриття викладачем, повторне відкриття)
- автоматичне створення відомості перездачі для студентів з незадовільними оцінками / неявками
- спільна логіка затвердження та визначення проблемних оцінок
"""
from __future__ import annotations

import sqlite3
from typing import Any


_EXTRA_SHEET_COLUMNS: tuple[tuple[str, str], ...] = (
    ('sheet_kind', "TEXT DEFAULT 'main'"),
    ('parent_sheet_id', 'INTEGER'),
    ('teacher_closed', 'INTEGER DEFAULT 0'),
    ('teacher_closed_at', 'TEXT'),
    ('reopened_at', 'TEXT'),
    ('reopened_by', 'TEXT'),
)


def ensure_exam_extra_columns(conn: sqlite3.Connection) -> set[str]:
    """Гарантує наявність додаткових колонок у таблиці session_sheets."""
    exists = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='session_sheets'"
    ).fetchone()
    if not exists:
        return set()

    cols = {
        r[1] for r in conn.execute("PRAGMA table_info(session_sheets)").fetchall()
    }
    changed = False
    for col_name, col_def in _EXTRA_SHEET_COLUMNS:
        if col_name not in cols:
            conn.execute(f"ALTER TABLE session_sheets ADD COLUMN {col_name} {col_def}")
            cols.add(col_name)
            changed = True

    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_ss_parent_sheet ON session_sheets(parent_sheet_id)"
    )
    if changed:
        try:
            conn.commit()
        except Exception:
            pass
    return cols


def is_failing_or_absent_student(st_row: Any) -> bool:
    """Повертає True, якщо студент не склав дисципліну (< 60 балів / FX / F) або не з'явився (points IS NULL)."""
    def _get(k: str, default: Any = None) -> Any:
        if isinstance(st_row, dict):
            return st_row.get(k, default)
        try:
            return st_row[k]
        except Exception:
            return default

    pts = _get('points')
    if pts is None or str(pts).strip() == '':
        return True
    try:
        if float(pts) < 60:
            return True
    except (TypeError, ValueError):
        return True

    ects = str(_get('ects') or '').strip().upper()
    if ects in ('FX', 'F'):
        return True

    nat = str(_get('national_grade') or '').strip().casefold()
    if any(tok in nat for tok in (
        'незадовіл',
        'незарахован',
        "не з'яв",
        'не зʼяв',
        'неяв',
        'fail',
    )):
        return True

    return False


def create_retake_sheet_if_needed(
    conn: sqlite3.Connection,
    sheet_id: int,
    *,
    created_by: str | None = None,
    allow_from_retake: bool = False,
) -> dict[str, Any] | None:
    """
    Якщо у затвердженій відомості є студенти з незадовільними оцінками (<60)
    або без оцінки (не з'явилися), створює окрему відомість перездачі
    тільки для цих студентів (без дублювання).
    """
    ensure_exam_extra_columns(conn)

    orig = conn.execute(
        "SELECT * FROM session_sheets WHERE id = ?", (sheet_id,)
    ).fetchone()
    if not orig:
        return None
    orig_d = dict(orig)

    orig_kind = (orig_d.get('sheet_kind') or 'main').strip().lower()
    if orig_kind == 'retake' and not allow_from_retake:
        return None

    # Перевіряємо, чи вже є дочірня відомість перездачі для цієї відомості
    existing = conn.execute(
        "SELECT id FROM session_sheets WHERE parent_sheet_id = ? ORDER BY id DESC LIMIT 1",
        (sheet_id,),
    ).fetchone()
    if existing and not allow_from_retake:
        return None

    students = conn.execute(
        """
        SELECT student_id, full_name, edebo_code, work_topic,
               points, national_grade, ects, position
        FROM session_sheet_students
        WHERE sheet_id = ?
        ORDER BY position, full_name
        """,
        (sheet_id,),
    ).fetchall()

    problem_students = [dict(r) for r in students if is_failing_or_absent_student(r)]
    if not problem_students:
        return None

    teacher_name = orig_d.get('teacher_name')
    if orig_d.get('teacher_ref_id'):
        t = conn.execute(
            "SELECT full_name, short_name FROM teachers WHERE id = ?",
            (orig_d['teacher_ref_id'],),
        ).fetchone()
        if t:
            teacher_name = (t[1] or t[0] or '').strip() or teacher_name

    cur = conn.execute(
        """
        INSERT INTO session_sheets (
            session_id, discipline_name, work_topic, entity_kind, control_form,
            hours_credits, teacher_name, teacher_ref_id, exam_date, statement_number,
            license_id, status, sheet_kind, parent_sheet_id, teacher_closed, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, 'draft', 'retake', ?, 0, ?)
        """,
        (
            orig_d['session_id'],
            orig_d['discipline_name'],
            orig_d.get('work_topic'),
            orig_d.get('entity_kind') or 'subject',
            orig_d.get('control_form') or 'Залік',
            orig_d.get('hours_credits'),
            teacher_name,
            orig_d.get('teacher_ref_id'),
            orig_d.get('license_id'),
            sheet_id,
            created_by,
        ),
    )
    new_sheet_id = cur.lastrowid

    for idx, st in enumerate(problem_students, start=1):
        conn.execute(
            """
            INSERT INTO session_sheet_students (
                sheet_id, student_id, full_name, edebo_code, work_topic, position,
                points_mid, points_current, points_exam, points, national_grade, ects
            ) VALUES (?, ?, ?, ?, ?, ?, NULL, NULL, NULL, NULL, NULL, NULL)
            """,
            (
                new_sheet_id,
                st['student_id'],
                st['full_name'],
                st.get('edebo_code'),
                st.get('work_topic'),
                idx,
            ),
        )

    return {
        'sheet_id': new_sheet_id,
        'parent_sheet_id': sheet_id,
        'session_id': orig_d['session_id'],
        'discipline_name': orig_d['discipline_name'],
        'students_count': len(problem_students),
        'license_id': orig_d.get('license_id'),
        'teacher_ref_id': orig_d.get('teacher_ref_id'),
    }
