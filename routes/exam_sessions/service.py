"""
Бізнес-логіка сесій: шкала оцінок, snapshot студентів, картотека викладачів.
"""
from __future__ import annotations

from datetime import datetime, timedelta
from typing import Optional

import secrets
import sqlite3

from routes.user_scope import get_allowed_group_ids
from routes.exam_sessions.schema import ensure_exam_session_tables


# ── Шкала оцінок ───────────────────────────────────────────────────

def parse_points(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    if s == '':
        return None
    try:
        v = int(float(s))
    except (TypeError, ValueError):
        return None
    if v < 0:
        v = 0
    if v > 100:
        v = 100
    return v


def compute_final_points(mid, current, exam, explicit_total=None):
    parts = [parse_points(x) for x in (mid, current, exam)]
    present = [p for p in parts if p is not None]
    if present:
        return min(100, sum(present))
    return parse_points(explicit_total)


def points_to_national_ects(points: Optional[int], control_form: str = 'Залік'):
    if points is None or points == '':
        return None, None
    try:
        p = int(points)
    except (TypeError, ValueError):
        return None, None
    if p < 0:
        p = 0
    if p > 100:
        p = 100

    if p >= 90:
        ects, exam_nat = 'A', 'Відмінно / Excellent'
    elif p >= 82:
        ects, exam_nat = 'B', 'Добре/ Good'
    elif p >= 74:
        ects, exam_nat = 'C', 'Добре/ Good'
    elif p >= 64:
        ects, exam_nat = 'D', 'Задовільно/ Satisfactory'
    elif p >= 60:
        ects, exam_nat = 'E', 'Задовільно/ Satisfactory'
    elif p >= 35:
        ects, exam_nat = 'FX', 'Незадовільно/ Fail'
    else:
        ects, exam_nat = 'F', 'Незадовільно/ Fail'

    if control_form == 'Залік':
        if p >= 60:
            return 'Зараховано', ects
        return 'Незараховано', ects if ects in ('FX', 'F') else 'F'

    return exam_nat, ects


def grade_stats(rows):
    buckets = {k: 0 for k in ('A', 'B', 'C', 'D', 'E', 'FX', 'F', 'absent')}
    for r in rows:
        pts = r.get('points') if isinstance(r, dict) else r['points'] if 'points' in r.keys() else None
        ects = r.get('ects') if isinstance(r, dict) else (r['ects'] if 'ects' in r.keys() else None)
        if pts is None and not ects:
            buckets['absent'] += 1
            continue
        key = (ects or 'F').upper()
        if key not in buckets:
            key = 'F'
        buckets[key] += 1
    return buckets


# ── Перевірка статусу відомості ────────────────────────────────────

def recompute_sheet_status(conn, sheet_id):
    """
    Перераховує статус відомості за реальною заповненістю.

    Статуси, які встановлюються тут:
      - 'draft'       — оцінок немає, токен не виданий
      - 'sent'        — токен видано, оцінок ще немає
      - 'in_progress' — частина оцінок є, але не всі
      - 'submitted'   — усі студенти заповнені

    Не чіпає 'approved' / 'locked' (ставляться вручну).
    Повертає новий статус (або поточний, якщо не змінювався).
    """
    row = conn.execute("""
        SELECT
            COUNT(*) AS total,
            SUM(CASE WHEN points IS NOT NULL THEN 1 ELSE 0 END) AS filled
        FROM session_sheet_students
        WHERE sheet_id = ?
    """, (sheet_id,)).fetchone()

    total = (row['total'] if row else 0) or 0
    filled = (row['filled'] if row else 0) or 0

    cur = conn.execute(
        "SELECT status, token FROM session_sheets WHERE id = ?", (sheet_id,)
    ).fetchone()
    if not cur:
        return None

    if cur['status'] in ('approved', 'locked'):
        return cur['status']

    if total == 0:
        new_status = 'draft'
    elif filled == 0:
        new_status = 'sent' if cur['token'] else 'draft'
    elif filled < total:
        new_status = 'in_progress'
    else:
        new_status = 'submitted'

    if new_status != cur['status']:
        conn.execute(
            "UPDATE session_sheets SET status = ? WHERE id = ?",
            (new_status, sheet_id),
        )
    return new_status


# ── Скоуп ──────────────────────────────────────────────────────────

def user_can_access_group(conn, session_obj, group_id: int) -> bool:
    if session_obj.get('is_admin') or session_obj.get('role') == 'admin':
        return True
    uid = session_obj.get('user_id')
    if not uid:
        return False
    return int(group_id) in set(get_allowed_group_ids(conn, uid))


def allowed_groups_for_user(conn, session_obj):
    ensure_exam_session_tables(conn)
    if session_obj.get('is_admin') or session_obj.get('role') == 'admin':
        rows = conn.execute("""
            SELECT id, name, start_year, study_form, program_credits, course,
                   degree_level, specialty, specialty_code, knowledge_area
            FROM groups WHERE COALESCE(archived, 0) = 0
            ORDER BY name, start_year
        """).fetchall()
        return [dict(r) for r in rows]

    uid = session_obj.get('user_id')
    ids = get_allowed_group_ids(conn, uid) if uid else []
    if not ids:
        return []
    ph = ','.join('?' * len(ids))
    rows = conn.execute(f"""
        SELECT id, name, start_year, study_form, program_credits, course,
               degree_level, specialty, specialty_code, knowledge_area
        FROM groups WHERE id IN ({ph}) AND COALESCE(archived, 0) = 0
        ORDER BY name, start_year
    """, ids).fetchall()
    return [dict(r) for r in rows]


# ── Студенти відомості ─────────────────────────────────────────────

def students_for_sheet(conn, group_id: int, license_id) -> list:
    ensure_exam_session_tables(conn)
    conn.row_factory = sqlite3.Row

    not_frozen = """
        AND NOT EXISTS (
            SELECT 1 FROM frozen_students fz
            WHERE fz.student_id = s.id AND fz.resolved_at IS NULL
        )
    """
    not_archived = "AND COALESCE(s.archived, 0) = 0"

    if license_id is not None:
        rows = conn.execute(f"""
            SELECT s.id AS student_id,
                   TRIM(s.last_name_UA || ' ' || s.first_name_UA || ' ' || COALESCE(s.middle_name_UA, '')) AS full_name,
                   e.edebo_code
            FROM student_enrollments e
            JOIN students s ON s.id = e.student_id
            WHERE e.group_id = ?
              AND e.status = 'active'
              AND e.license_id = ?
              {not_archived}
              {not_frozen}
            ORDER BY s.last_name_UA, s.first_name_UA
        """, (group_id, license_id)).fetchall()
        return [dict(r) for r in rows]

    rows = conn.execute(f"""
        SELECT s.id AS student_id,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA || ' ' || COALESCE(s.middle_name_UA, '')) AS full_name,
               e.edebo_code
        FROM student_enrollments e
        JOIN students s ON s.id = e.student_id
        WHERE e.group_id = ?
          AND e.status = 'active'
          AND e.license_id IS NULL
          {not_archived}
          {not_frozen}
        ORDER BY s.last_name_UA, s.first_name_UA
    """, (group_id,)).fetchall()
    result = [dict(r) for r in rows]
    seen = {r['student_id'] for r in result}

    legacy = conn.execute(f"""
        SELECT s.id AS student_id,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA || ' ' || COALESCE(s.middle_name_UA, '')) AS full_name,
               NULL AS edebo_code
        FROM students s
        WHERE s.group_id = ?
          {not_archived}
          {not_frozen}
          AND NOT EXISTS (
              SELECT 1 FROM student_enrollments e
              WHERE e.student_id = s.id AND e.group_id = ? AND e.status = 'active'
          )
        ORDER BY s.last_name_UA, s.first_name_UA
    """, (group_id, group_id)).fetchall()
    for r in legacy:
        d = dict(r)
        if d['student_id'] not in seen:
            result.append(d)
    return result


def licenses_present_in_group(conn, group_id: int) -> list:
    rows = conn.execute("""
        SELECT DISTINCT e.license_id AS id,
               COALESCE(l.short_name_ua, l.name_ua, '—') AS name
        FROM student_enrollments e
        LEFT JOIN institution_licenses l ON l.id = e.license_id
        WHERE e.group_id = ? AND e.status = 'active' AND e.license_id IS NOT NULL
        ORDER BY name
    """, (group_id,)).fetchall()
    return [dict(r) for r in rows]


def populate_sheet_students(conn, sheet_id: int, group_id: int, license_id) -> int:
    conn.execute("DELETE FROM session_sheet_students WHERE sheet_id = ?", (sheet_id,))
    students = students_for_sheet(conn, group_id, license_id)
    for i, st in enumerate(students, start=1):
        conn.execute("""
            INSERT INTO session_sheet_students
                (sheet_id, student_id, full_name, edebo_code, position)
            VALUES (?, ?, ?, ?, ?)
        """, (sheet_id, st['student_id'], st['full_name'], st.get('edebo_code'), i))
    # Після перезапису — перераховуємо статус
    recompute_sheet_status(conn, sheet_id)
    return len(students)


def issue_sheet_token(conn, sheet_id: int, days: int = 21) -> str:
    token = secrets.token_urlsafe(24)
    expires = (datetime.now() + timedelta(days=days)).strftime('%Y-%m-%d %H:%M:%S')
    conn.execute("""
        UPDATE session_sheets
        SET token = ?, token_expires_at = ?, status = CASE
            WHEN status = 'draft' THEN 'sent'
            ELSE status
        END
        WHERE id = ?
    """, (token, expires, sheet_id))
    return token


def sheet_token_valid(row) -> bool:
    if not row or not row['token']:
        return False
    if row['status'] in ('locked',):
        return False
    exp = row['token_expires_at']
    if exp:
        try:
            if datetime.now() > datetime.strptime(exp, '%Y-%m-%d %H:%M:%S'):
                return False
        except ValueError:
            pass
    return True


# ═══════════════════════════════════════════════════════════════════
# Картотека викладачів
# ═══════════════════════════════════════════════════════════════════

def generate_teacher_token() -> str:
    return secrets.token_urlsafe(24)


def teacher_stats(conn, teacher_id: int):
    """Статистика по викладачу (для сторінки деталей)."""
    rows = conn.execute("""
        SELECT ss.id, ss.status, ss.token,
               ss.discipline_name, ss.control_form,
               es.academic_year, es.semester,
               g.name AS group_name,
               (SELECT COUNT(*) FROM session_sheet_students x
                WHERE x.sheet_id = ss.id) AS students_count,
               (SELECT COUNT(*) FROM session_sheet_students x
                WHERE x.sheet_id = ss.id AND x.points IS NOT NULL) AS graded_count
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        WHERE ss.teacher_ref_id = ?
        ORDER BY
            CASE ss.status
                WHEN 'draft' THEN 0
                WHEN 'sent' THEN 1
                WHEN 'in_progress' THEN 2
                WHEN 'submitted' THEN 3
                WHEN 'approved' THEN 4
                WHEN 'locked' THEN 5
                ELSE 6
            END,
            es.academic_year DESC, es.semester DESC,
            ss.discipline_name
    """, (teacher_id,)).fetchall()

    rows = [dict(r) for r in rows]
    total = len(rows)
    submitted = sum(1 for r in rows if r['status'] in ('submitted', 'approved', 'locked'))
    in_progress = sum(1 for r in rows if r['status'] == 'in_progress')
    pending = total - submitted - in_progress

    return {
        'total': total,
        'submitted': submitted,
        'in_progress': in_progress,
        'pending': pending,
        'percent': round(submitted / total * 100) if total else 0,
        'sheets': rows,
    }