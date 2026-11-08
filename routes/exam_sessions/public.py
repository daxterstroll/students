"""
Публічна форма викладача: /session-grade/<token> (без логіну).
"""
from __future__ import annotations

import sqlite3

from flask import render_template, request, redirect, url_for, flash

from routes.db import get_db
from routes.utils import logger
from routes.exam_sessions import exam_sessions_bp
from routes.exam_sessions.schema import ensure_exam_session_tables
from routes.exam_sessions.service import parse_points, compute_final_points, sheet_token_valid, points_to_national_ects


@exam_sessions_bp.route('/session-grade/<token>', methods=['GET', 'POST'])
def teacher_grade_form(token):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)

    sheet = conn.execute("""
        SELECT ss.*, es.academic_year, es.semester, es.name AS session_name,
               g.name AS group_name, g.course, g.study_form, g.specialty,
               g.degree_level, g.knowledge_area,
               l.short_name_ua AS license_name
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        WHERE ss.token = ?
    """, (token,)).fetchone()

    if not sheet or not sheet_token_valid(sheet):
        conn.close()
        return render_template('exam_sessions/teacher_invalid.html'), 404

    sheet = dict(sheet)
    # Після затвердження викладач більше не редагує
    readonly = sheet['status'] in ('approved', 'locked')

    if request.method == 'POST' and not readonly:
        by_sid = {}
        for key, val in request.form.items():
            for prefix in ('points_mid_', 'points_current_', 'points_exam_', 'points_', 'work_topic_'):
                if key.startswith(prefix):
                    sid = key[len(prefix):]
                    by_sid.setdefault(sid, {})[prefix.rstrip('_')] = val
                    break
        for sid, fields in by_sid.items():
            if 'work_topic' in fields:
                topic = (fields.get('work_topic') or '').strip() or None
                conn.execute("""
                    UPDATE session_sheet_students SET work_topic=?
                    WHERE sheet_id=? AND student_id=?
                """, (topic, sheet['id'], sid))
            mid = parse_points(fields.get('points_mid'))
            cur = parse_points(fields.get('points_current'))
            exam = parse_points(fields.get('points_exam'))
            total = compute_final_points(mid, cur, exam, fields.get('points'))
            nat, ects = points_to_national_ects(total, sheet['control_form'])
            conn.execute("""
                UPDATE session_sheet_students
                SET points_mid=?, points_current=?, points_exam=?,
                    points=?, national_grade=?, ects=?
                WHERE sheet_id=? AND student_id=?
            """, (mid, cur, exam, total, nat, ects, sheet['id'], sid))

        new_status = 'submitted' if sheet['status'] in ('draft', 'sent') else sheet['status']
        conn.execute("""
            UPDATE session_sheets
            SET status=?, submitted_at=datetime('now','localtime')
            WHERE id=?
        """, (new_status, sheet['id']))
        conn.commit()
        logger.info(f"[session-grade] teacher saved sheet_id={sheet['id']}")
        flash('Оцінки збережено. Дякуємо!', 'success')
        conn.close()
        return redirect(url_for('exam_sessions.teacher_grade_form', token=token))

    students = [dict(r) for r in conn.execute("""
        SELECT * FROM session_sheet_students
        WHERE sheet_id = ?
        ORDER BY position, full_name
    """, (sheet['id'],)).fetchall()]
    conn.close()
    return render_template(
        'exam_sessions/teacher_form.html',
        sheet=sheet, students=students, readonly=readonly, token=token,
    )
