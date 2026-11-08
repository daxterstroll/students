"""
Бізнес-логіка сесій: шкала оцінок, snapshot студентів, перевірка скоупу.
"""
from __future__ import annotations

from datetime import datetime, timedelta
from typing import Optional

import secrets
import sqlite3

from routes.user_scope import get_allowed_group_ids
from routes.exam_sessions.schema import ensure_exam_session_tables


# ── Стандартна шкала (ЄКТС / національна) ──────────────────────────

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
    """
    Два режими:
      1) Є хоч одне з mid/current/exam → підсумок = їх сума (макс. 100)
      2) Інакше → береться explicit_total (пряме введення «Підсумкова»)
    Порожні компоненти в сумі рахуються як 0 лише якщо є хоч одне заповнене.
    """
    parts = [parse_points(x) for x in (mid, current, exam)]
    present = [p for p in parts if p is not None]
    if present:
        return min(100, sum(present))
    return parse_points(explicit_total)


def points_to_national_ects(points: Optional[int], control_form: str = 'Залік'):
    """
    Повертає (national_grade, ects) за балами 0..100.
    Для заліку: >=60 → Зараховано / відповідний ECTS; <60 → Незараховано / F.
    Для екзамену: повна шкала Відмінно…Незадовільно.
    """
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
    """Підсумки A..F + не з'явився для Excel."""
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


# ── Скоуп ──────────────────────────────────────────────────────────

def user_can_access_group(conn, session_obj, group_id: int) -> bool:
    if session_obj.get('is_admin') or session_obj.get('role') == 'admin':
        return True
    uid = session_obj.get('user_id')
    if not uid:
        return False
    return int(group_id) in set(get_allowed_group_ids(conn, uid))


def allowed_groups_for_user(conn, session_obj):
    """Список груп (id, name, …) доступних користувачу."""
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


# ── Студенти відомості за ліцензією ────────────────────────────────

def students_for_sheet(conn, group_id: int, license_id) -> list:
    """
    Активні студенти групи для відомості.

    Виключаються:
      - archived студенти (s.archived = 1)
      - заморожені (активний запис у frozen_students, resolved_at IS NULL)
      - enrollment не active

    license_id is not None → лише enrollment з цією ліцензією.
    license_id is None → enrollment без license_id або legacy students.group_id.
    """
    ensure_exam_session_tables(conn)
    conn.row_factory = sqlite3.Row

    # Як у analytics / students list: активна «заморозка»
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

    # Без ліцензії: active enrollment з license_id IS NULL
    # + студенти з primary group_id без жодного enrollment (legacy)
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
    """Ліцензії, які реально є серед активних enrollment групи (+ опція NULL)."""
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
    """Перезаписує snapshot студентів відомості. Повертає кількість."""
    conn.execute("DELETE FROM session_sheet_students WHERE sheet_id = ?", (sheet_id,))
    students = students_for_sheet(conn, group_id, license_id)
    for i, st in enumerate(students, start=1):
        conn.execute("""
            INSERT INTO session_sheet_students
                (sheet_id, student_id, full_name, edebo_code, position)
            VALUES (?, ?, ?, ?, ?)
        """, (sheet_id, st['student_id'], st['full_name'], st.get('edebo_code'), i))
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
