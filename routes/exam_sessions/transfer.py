"""
Перенос затверджених відомостей у subjects/practices/courseworks/attestations + grades.
Право: transfer_exam_grades (або admin).
"""
from __future__ import annotations

import sqlite3

from flask import render_template, request, redirect, url_for, flash, session

from routes.db import get_db
from routes.utils import permission_required, log_action, logger
from routes.helpers import current_username
from routes.exam_sessions import exam_sessions_bp
from routes.exam_sessions.schema import ensure_exam_session_tables
from routes.exam_sessions.service import user_can_access_group


def _has_transfer_perm():
    if session.get('is_admin') or session.get('role') == 'admin':
        return True
    return 'transfer_exam_grades' in (session.get('permissions') or [])


def _require_transfer():
    """decorator-like check used inside views that also allow manage for list."""
    return _has_transfer_perm()


@exam_sessions_bp.route('/admin/exam-sessions/transfer')
@permission_required('transfer_exam_grades')
def transfer_queue():
    """Черга затверджених відомостей, ще не прив’язаних (або вже прив’язаних)."""
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)

    from routes.exam_sessions.service import allowed_groups_for_user
    groups = allowed_groups_for_user(conn, session)
    group_ids = [g['id'] for g in groups]
    rows = []
    if group_ids:
        ph = ','.join('?' * len(group_ids))
        rows = [dict(r) for r in conn.execute(f"""
            SELECT ss.id, ss.discipline_name, ss.entity_kind, ss.control_form,
                   ss.status, ss.linked_entity_type, ss.linked_entity_id, ss.linked_at,
                   es.academic_year, es.semester, es.name AS session_name,
                   g.name AS group_name, g.id AS group_id,
                   l.short_name_ua AS license_name,
                   (SELECT COUNT(*) FROM session_sheet_students x WHERE x.sheet_id = ss.id) AS students_count
            FROM session_sheets ss
            JOIN exam_sessions es ON es.id = ss.session_id
            JOIN groups g ON g.id = es.group_id
            LEFT JOIN institution_licenses l ON l.id = ss.license_id
            WHERE es.group_id IN ({ph})
              AND ss.status IN ('approved', 'locked')
            ORDER BY ss.linked_entity_id IS NOT NULL, ss.approved_at DESC
        """, group_ids).fetchall()]
    conn.close()
    return render_template('exam_sessions/transfer_queue.html', rows=rows)


@exam_sessions_bp.route('/admin/exam-sessions/sheets/<int:sheet_id>/transfer', methods=['GET', 'POST'])
@permission_required('transfer_exam_grades')
def transfer_sheet(sheet_id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)

    sheet = conn.execute("""
        SELECT ss.*, es.group_id, es.name AS session_name, g.name AS group_name
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        WHERE ss.id = ?
    """, (sheet_id,)).fetchone()

    if not sheet or not user_can_access_group(conn, session, sheet['group_id']):
        conn.close()
        flash('Відомість не знайдено', 'error')
        return redirect(url_for('exam_sessions.transfer_queue'))

    if sheet['status'] not in ('approved', 'locked'):
        conn.close()
        flash('Спочатку затвердіть відомість', 'error')
        return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

    sheet = dict(sheet)
    group_id = sheet['group_id']
    kind = sheet['entity_kind']

    table_map = {
        'subject': ('subjects', 'grades', 'subject_id'),
        'practice': ('practices', 'activity_grades', 'entity_id'),
        'coursework': ('courseworks', 'activity_grades', 'entity_id'),
        'attestation': ('attestations', 'activity_grades', 'entity_id'),
    }
    entity_table, grade_table, grade_fk = table_map.get(kind, table_map['subject'])

    candidates = [dict(r) for r in conn.execute(
        f"SELECT id, code, name, credits, type FROM {entity_table} WHERE group_id = ? ORDER BY position, name",
        (group_id,),
    ).fetchall()]

    if request.method == 'POST':
        target_id = request.form.get('target_id', type=int)
        if not target_id or not any(c['id'] == target_id for c in candidates):
            flash('Оберіть елемент навчального плану', 'error')
        else:
            students = conn.execute(
                "SELECT student_id, points, national_grade, ects, work_topic "
                "FROM session_sheet_students WHERE sheet_id=?",
                (sheet_id,),
            ).fetchall()
            written = 0
            for st in students:
                # Зберігаємо в grades як текст балів або національну — як прийнято в системі (TEXT)
                grade_value = None
                if st['points'] is not None:
                    grade_value = str(st['points'])
                elif st['national_grade']:
                    grade_value = st['national_grade']
                if grade_value is None:
                    continue

                if grade_table == 'grades':
                    existing = conn.execute(
                        "SELECT id FROM grades WHERE student_id=? AND subject_id=?",
                        (st['student_id'], target_id),
                    ).fetchone()
                    if existing:
                        conn.execute(
                            "UPDATE grades SET grade=? WHERE id=?",
                            (grade_value, existing['id']),
                        )
                    else:
                        conn.execute(
                            "INSERT INTO grades (student_id, subject_id, grade) VALUES (?, ?, ?)",
                            (st['student_id'], target_id, grade_value),
                                        )
                else:
                    # activity_grades: grade INTEGER, name TEXT (тема роботи)
                    pts = st['points']
                    work_topic = st['work_topic']

                    existing = conn.execute(
                        "SELECT id FROM activity_grades WHERE student_id=? AND entity_id=? AND entity_type=?",
                        (st['student_id'], target_id, kind),
                    ).fetchone()

                    if existing:
                        conn.execute(
                            "UPDATE activity_grades "
                            "SET grade=?, name=COALESCE(NULLIF(?, ''), name) "
                            "WHERE id=?",
                            (pts, work_topic, existing['id']),
                        )
                    else:
                        conn.execute(
                            "INSERT INTO activity_grades "
                            "(student_id, entity_id, entity_type, grade, name) "
                            "VALUES (?, ?, ?, ?, ?)",
                            (st['student_id'], target_id, kind, pts, work_topic),
                        )
                written += 1

            conn.execute("""
                UPDATE session_sheets
                SET linked_entity_type=?, linked_entity_id=?,
                    linked_at=datetime('now','localtime'), linked_by=?
                WHERE id=?
            """, (kind, target_id, current_username(), sheet_id))
            conn.commit()
            log_action(
                current_username(),
                f"переніс відомість #{sheet_id} → {kind}#{target_id} ({written} оцінок)",
            )
            flash(f'Перенесено оцінок: {written}', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.transfer_queue'))

    # fuzzy hint by name
    import difflib
    disc = (sheet['discipline_name'] or '').lower()
    for c in candidates:
        c['score'] = difflib.SequenceMatcher(None, disc, (c['name'] or '').lower()).ratio()
    candidates.sort(key=lambda x: -x['score'])

    conn.close()
    return render_template(
        'exam_sessions/transfer_sheet.html',
        sheet=sheet, candidates=candidates, entity_table=entity_table,
    )
