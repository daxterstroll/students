# routes/exam_sessions/public.py

from __future__ import annotations

import sqlite3

from flask import render_template, request, redirect, url_for, flash

from routes.db import get_db
from routes.utils import logger
from routes.exam_sessions import exam_sessions_bp
from routes.exam_sessions.schema import ensure_exam_session_tables
from routes.exam_sessions.service import (
    parse_points, compute_final_points, sheet_token_valid,
    points_to_national_ects, recompute_sheet_status,
)


def _save_teacher_grades(conn, sheet_id, sheet_control_form, form):
    """Спільна логіка збереження оцінок (для кабінету і для старої ссылки)."""
    by_sid = {}
    for key, val in form.items():
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
            """, (topic, sheet_id, sid))
        mid = parse_points(fields.get('points_mid'))
        cur = parse_points(fields.get('points_current'))
        exam = parse_points(fields.get('points_exam'))
        total = compute_final_points(mid, cur, exam, fields.get('points'))
        nat, ects = points_to_national_ects(total, sheet_control_form)
        conn.execute("""
            UPDATE session_sheet_students
            SET points_mid=?, points_current=?, points_exam=?,
                points=?, national_grade=?, ects=?
            WHERE sheet_id=? AND student_id=?
        """, (mid, cur, exam, total, nat, ects, sheet_id, sid))


# ═══════════════════════════════════════════════════════════════════
# Кабінет викладача (постійна ссылка)
# ═══════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/session-grade/teacher/<teacher_token>')
def teacher_cabinet(teacher_token):
    """Кабінет викладача: усі його відомості (постійна ссылка)."""
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)

    teacher = conn.execute(
        "SELECT * FROM teachers WHERE token = ? AND is_active = 1",
        (teacher_token,),
    ).fetchone()

    if not teacher:
        conn.close()
        return render_template('exam_sessions/teacher_invalid.html'), 404

    sheets = conn.execute("""
        SELECT ss.*,
               es.academic_year, es.semester, es.name AS session_name,
               g.name AS group_name, g.course, g.study_form, g.specialty,
               g.degree_level,
               l.short_name_ua AS license_name,
               (SELECT COUNT(*) FROM session_sheet_students x
                WHERE x.sheet_id = ss.id) AS students_count,
               (SELECT COUNT(*) FROM session_sheet_students x
                WHERE x.sheet_id = ss.id AND x.points IS NOT NULL) AS graded_count
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
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
            ss.discipline_name, ss.license_id
    """, (teacher['id'],)).fetchall()

    # Розширена статистика з окремими категоріями статусів.
    # Раніше в шаблоні була спроба вивести `stats.needs_work - stats.draft - stats.sent`,
    # але `draft` і `sent` не передавались → виводився 0.
    stats = {
        'total': len(sheets),
        'draft': sum(1 for s in sheets if s['status'] == 'draft'),
        'sent': sum(1 for s in sheets if s['status'] == 'sent'),
        'in_progress': sum(1 for s in sheets if s['status'] == 'in_progress'),
        'submitted': sum(1 for s in sheets if s['status'] in ('submitted', 'approved', 'locked')),
    }
    stats['not_started'] = stats['draft'] + stats['sent']
    stats['needs_work'] = stats['draft'] + stats['sent'] + stats['in_progress']
    stats['percent'] = round(stats['submitted'] / stats['total'] * 100) if stats['total'] else 0

    conn.close()
    return render_template(
        'exam_sessions/teacher_cabinet.html',
        teacher=teacher,
        teacher_token=teacher_token,
        sheets=sheets,
        stats=stats,
    )


@exam_sessions_bp.route('/session-grade/teacher/<teacher_token>/sheet/<int:sheet_id>',
                       methods=['GET', 'POST'])
def teacher_sheet(teacher_token, sheet_id):
    """Форма однієї відомості в кабінеті викладача."""
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)

    teacher = conn.execute(
        "SELECT * FROM teachers WHERE token = ? AND is_active = 1",
        (teacher_token,),
    ).fetchone()
    if not teacher:
        conn.close()
        return render_template('exam_sessions/teacher_invalid.html'), 404

    sheet = conn.execute("""
        SELECT ss.*, es.academic_year, es.semester, es.name AS session_name,
               g.name AS group_name, g.course, g.study_form, g.specialty,
               g.degree_level, g.knowledge_area,
               l.short_name_ua AS license_name
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        WHERE ss.id = ?
    """, (sheet_id,)).fetchone()

    if not sheet or sheet['teacher_ref_id'] != teacher['id']:
        conn.close()
        return render_template('exam_sessions/teacher_invalid.html'), 404

    sheet = dict(sheet)
    readonly = sheet['status'] in ('approved', 'locked')

    if request.method == 'POST' and not readonly:
        _save_teacher_grades(conn, sheet_id, sheet['control_form'], request.form)
        new_status = recompute_sheet_status(conn, sheet_id)

        from routes.exam_sessions.history import log_history
        log_history(
            conn,
            entity_type='sheet', entity_id=sheet_id,
            session_id=sheet['session_id'], sheet_id=sheet_id,
            action='update', field_name='status',
            old_value=sheet['status'], new_value=new_status,
            details={'saved_by': 'teacher (cabinet)'},
        )

        conn.commit()

        next_sheet = conn.execute("""
            SELECT id FROM session_sheets
            WHERE teacher_ref_id = ?
              AND status IN ('draft', 'sent', 'in_progress')
              AND id != ?
            ORDER BY
                CASE status
                    WHEN 'in_progress' THEN 0
                    WHEN 'sent' THEN 1
                    WHEN 'draft' THEN 2
                    ELSE 3
                END,
                discipline_name
            LIMIT 1
        """, (teacher['id'], sheet_id)).fetchone()

        if next_sheet:
            conn.close()
            flash('Оцінки збережено. Переходимо до наступної відомості.', 'success')
            return redirect(url_for(
                'exam_sessions.teacher_sheet',
                teacher_token=teacher_token, sheet_id=next_sheet['id'],
            ))

        conn.close()
        flash('Оцінки збережено. Усі відомості заповнено!', 'success')
        return redirect(url_for(
            'exam_sessions.teacher_cabinet', teacher_token=teacher_token,
        ))

    students = [dict(r) for r in conn.execute("""
        SELECT * FROM session_sheet_students
        WHERE sheet_id = ?
        ORDER BY position, full_name
    """, (sheet_id,)).fetchall()]

    all_ids = [
        r['id'] for r in conn.execute("""
            SELECT id FROM session_sheets
            WHERE teacher_ref_id = ?
            ORDER BY
                CASE status
                    WHEN 'draft' THEN 0
                    WHEN 'sent' THEN 1
                    WHEN 'in_progress' THEN 2
                    WHEN 'submitted' THEN 3
                    WHEN 'approved' THEN 4
                    WHEN 'locked' THEN 5
                    ELSE 6
                END,
                discipline_name
        """, (teacher['id'],)).fetchall()
    ]
    try:
        idx = all_ids.index(sheet_id)
    except ValueError:
        idx = 0

    prev_id = all_ids[idx - 1] if idx > 0 else None
    next_id = all_ids[idx + 1] if idx < len(all_ids) - 1 else None

    conn.close()
    return render_template(
        'exam_sessions/teacher_sheet.html',
        teacher=teacher,
        teacher_token=teacher_token,
        sheet=sheet,
        students=students,
        readonly=readonly,
        prev_sheet_id=prev_id,
        next_sheet_id=next_id,
        position={'current': idx + 1, 'total': len(all_ids)},
        token=None,   # стара одноразова ссылка не використовується
    )


# ═══════════════════════════════════════════════════════════════════
# Стара одноразова ссылка /session-grade/<token> (залишаємо для сумісності)
# ═══════════════════════════════════════════════════════════════════
# ВАЖЛИВО: раніше ця функція рендерила `teacher_form.html`, який насправді
# є АДМІНСЬКИМ шаблоном (потребує teacher, users, stats) — тому сторінка
# ламалася з помилкою рендерингу. Тепер рендеримо `teacher_sheet.html` і
# передаємо teacher_token=None — шаблон адаптується (без навігації).

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
    readonly = sheet['status'] in ('approved', 'locked')

    if request.method == 'POST' and not readonly:
        _save_teacher_grades(conn, sheet['id'], sheet['control_form'], request.form)
        new_status = recompute_sheet_status(conn, sheet['id'])

        from routes.exam_sessions.history import log_history
        log_history(
            conn,
            entity_type='sheet', entity_id=sheet['id'],
            session_id=sheet['session_id'], sheet_id=sheet['id'],
            action='update', field_name='status',
            old_value=sheet['status'], new_value=new_status,
            details={'saved_by': 'teacher (public form)'},
        )

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

    # Рендеримо той самий teacher_sheet.html, але без навігації кабінету
    # (teacher_token=None, prev/next=None) — це "стара" публічна форма.
    return render_template(
        'exam_sessions/teacher_sheet.html',
        sheet=sheet,
        students=students,
        readonly=readonly,
        teacher=None,
        teacher_token=None,
        prev_sheet_id=None,
        next_sheet_id=None,
        position=None,
        token=token,
    )