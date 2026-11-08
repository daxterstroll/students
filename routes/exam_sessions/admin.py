"""
Адмін/відповідальний: список сесій, CRUD, відомості, токен, затвердження, Excel.
"""
from __future__ import annotations

import sqlite3
from datetime import datetime

from flask import (
    render_template, request, redirect, url_for, flash, session, send_file
)

from routes.db import get_db
from routes.utils import permission_required, log_action, logger
from routes.helpers import current_username
from routes.config import PUBLIC_APPLY_BASE_URL
from routes.exam_sessions import exam_sessions_bp
from routes.exam_sessions.schema import ensure_exam_session_tables
from routes.exam_sessions.service import (
    allowed_groups_for_user, user_can_access_group,
    licenses_present_in_group, populate_sheet_students,
    issue_sheet_token, points_to_national_ects, parse_points, compute_final_points,
)
from routes.exam_sessions.excel_export import build_statement_xlsx, build_group_course_summary_xlsx
from routes.exam_sessions.word_export import build_sheet_docx, build_session_summary_docx


def _conn():
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)
    return conn


# ── Список сесій ───────────────────────────────────────────────────

@exam_sessions_bp.route('/admin/exam-sessions')
@permission_required('manage_exam_sessions')
def sessions_list():
    conn = _conn()
    groups = allowed_groups_for_user(conn, session)
    group_ids = [g['id'] for g in groups]

    filter_group = request.args.get('group_id', type=int)
    filter_year = (request.args.get('academic_year') or '').strip()
    filter_status = (request.args.get('status') or '').strip()

    rows = []
    if group_ids:
        ph = ','.join('?' * len(group_ids))
        sql = f"""
            SELECT es.*, g.name AS group_name, g.course, g.study_form,
                   g.start_year, g.specialty,
                   (SELECT COUNT(*) FROM session_sheets ss WHERE ss.session_id = es.id) AS sheets_count
            FROM exam_sessions es
            JOIN groups g ON g.id = es.group_id
            WHERE es.group_id IN ({ph})
        """
        params = list(group_ids)
        if filter_group and filter_group in group_ids:
            sql += " AND es.group_id = ?"
            params.append(filter_group)
        if filter_year:
            sql += " AND es.academic_year = ?"
            params.append(filter_year)
        if filter_status in ('draft', 'open', 'closed'):
            sql += " AND es.status = ?"
            params.append(filter_status)
        sql += " ORDER BY es.created_at DESC"
        rows = [dict(r) for r in conn.execute(sql, params).fetchall()]

    years = sorted({r['academic_year'] for r in rows}, reverse=True)
    conn.close()
    return render_template(
        'exam_sessions/sessions_list.html',
        rows=rows, groups=groups, years=years,
        filter_group=filter_group, filter_year=filter_year,
        filter_status=filter_status,
    )


# ── Створення / редагування сесії ─────────────────────────────────

@exam_sessions_bp.route('/admin/exam-sessions/new', methods=['GET', 'POST'])
@permission_required('manage_exam_sessions')
def session_new():
    conn = _conn()
    groups = allowed_groups_for_user(conn, session)
    if request.method == 'POST':
        group_id = request.form.get('group_id', type=int)
        academic_year = (request.form.get('academic_year') or '').strip()
        semester = request.form.get('semester', type=int)
        name = (request.form.get('name') or '').strip()
        notes = (request.form.get('notes') or '').strip() or None

        if not group_id or not user_can_access_group(conn, session, group_id):
            flash('Оберіть доступну групу', 'error')
        elif not academic_year or semester not in (1, 2, 3):
            flash('Вкажіть навчальний рік і семестр (1–3)', 'error')
        else:
            if not name:
                name = f"Сесія {semester} семестру {academic_year}"
            cur = conn.execute("""
                INSERT INTO exam_sessions
                    (group_id, academic_year, semester, name, status, notes, created_by)
                VALUES (?, ?, ?, ?, 'open', ?, ?)
            """, (group_id, academic_year, semester, name, notes, current_username()))
            conn.commit()
            sid = cur.lastrowid
            log_action(current_username(), f"створив сесію #{sid}: {name}")
            conn.close()
            flash('Сесію створено', 'success')
            return redirect(url_for('exam_sessions.session_detail', session_id=sid))
    conn.close()
    return render_template('exam_sessions/session_form.html', groups=groups, row=None)


@exam_sessions_bp.route('/admin/exam-sessions/<int:session_id>', methods=['GET', 'POST'])
@permission_required('manage_exam_sessions')
def session_detail(session_id):
    conn = _conn()
    es = conn.execute("""
        SELECT es.*, g.name AS group_name, g.course, g.study_form,
               g.start_year, g.specialty, g.specialty_code, g.degree_level,
               g.knowledge_area, g.program_credits
        FROM exam_sessions es
        JOIN groups g ON g.id = es.group_id
        WHERE es.id = ?
    """, (session_id,)).fetchone()
    if not es or not user_can_access_group(conn, session, es['group_id']):
        conn.close()
        flash('Сесію не знайдено або немає доступу', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))
    es = dict(es)

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'update_meta':
            name = (request.form.get('name') or '').strip() or es['name']
            status = request.form.get('status') or es['status']
            if status not in ('draft', 'open', 'closed'):
                status = es['status']
            notes = (request.form.get('notes') or '').strip() or None
            conn.execute("""
                UPDATE exam_sessions
                SET name=?, status=?, notes=?, updated_at=datetime('now','localtime')
                WHERE id=?
            """, (name, status, notes, session_id))
            conn.commit()
            flash('Сесію оновлено', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

    sheets = [dict(r) for r in conn.execute("""
        SELECT ss.*,
               l.short_name_ua AS license_name,
               (SELECT COUNT(*) FROM session_sheet_students x WHERE x.sheet_id = ss.id) AS students_count,
               (SELECT COUNT(*) FROM session_sheet_students x
                WHERE x.sheet_id = ss.id AND x.points IS NOT NULL) AS graded_count
        FROM session_sheets ss
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        WHERE ss.session_id = ?
        ORDER BY ss.discipline_name, ss.license_id
    """, (session_id,)).fetchall()]

    licenses = licenses_present_in_group(conn, es['group_id'])
    can_transfer = session.get('is_admin') or 'transfer_exam_grades' in (session.get('permissions') or [])
    conn.close()
    return render_template(
        'exam_sessions/session_detail.html',
        es=es, sheets=sheets, licenses=licenses, can_transfer=can_transfer,
    )


# ── Відомість: створення ──────────────────────────────────────────

@exam_sessions_bp.route('/admin/exam-sessions/<int:session_id>/sheets/new', methods=['GET', 'POST'])
@permission_required('manage_exam_sessions')
def sheet_new(session_id):
    conn = _conn()
    es = conn.execute(
        "SELECT * FROM exam_sessions WHERE id = ?", (session_id,)
    ).fetchone()
    if not es or not user_can_access_group(conn, session, es['group_id']):
        conn.close()
        flash('Сесію не знайдено', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))
    es = dict(es)
    licenses = licenses_present_in_group(conn, es['group_id'])

    if request.method == 'POST':
        discipline_name = (request.form.get('discipline_name') or '').strip()
        work_topic = (request.form.get('work_topic') or '').strip() or None
        entity_kind = request.form.get('entity_kind') or 'subject'
        control_form = request.form.get('control_form') or 'Залік'
        hours_credits = (request.form.get('hours_credits') or '').strip() or None
        teacher_name = (request.form.get('teacher_name') or '').strip() or None
        exam_date = (request.form.get('exam_date') or '').strip() or None
        statement_number = (request.form.get('statement_number') or '').strip() or None
        lic_raw = request.form.get('license_id')
        license_id = int(lic_raw) if lic_raw and lic_raw.isdigit() else None

        if entity_kind not in ('subject', 'practice', 'coursework', 'attestation'):
            entity_kind = 'subject'
        if control_form not in ('Залік', 'Екзамен'):
            control_form = 'Залік'

        if not discipline_name:
            flash('Вкажіть назву дисципліни', 'error')
        else:
            cur = conn.execute("""
                INSERT INTO session_sheets (
                    session_id, discipline_name, work_topic, entity_kind, control_form,
                    hours_credits, teacher_name, exam_date, statement_number,
                    license_id, status, created_by
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'draft', ?)
            """, (
                session_id, discipline_name, work_topic, entity_kind, control_form,
                hours_credits, teacher_name, exam_date, statement_number,
                license_id, current_username(),
            ))
            sheet_id = cur.lastrowid
            n = populate_sheet_students(conn, sheet_id, es['group_id'], license_id)
            conn.commit()
            log_action(
                current_username(),
                f"створив відомість #{sheet_id} «{discipline_name}» у сесії #{session_id} ({n} студ.)",
            )
            conn.close()
            flash(f'Відомість створено, студентів у списку: {n}', 'success')
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

    conn.close()
    return render_template(
        'exam_sessions/sheet_form.html',
        es=es, licenses=licenses, sheet=None,
    )


# ── Відомість: деталі / токен / затвердження ───────────────────────

@exam_sessions_bp.route('/admin/exam-sessions/sheets/<int:sheet_id>', methods=['GET', 'POST'])
@permission_required('manage_exam_sessions')
def sheet_detail(sheet_id):
    conn = _conn()
    sheet = conn.execute("""
        SELECT ss.*, es.group_id, es.name AS session_name, es.academic_year,
               es.semester, g.name AS group_name, g.course, g.study_form,
               g.specialty, g.degree_level, g.knowledge_area,
               l.short_name_ua AS license_name, l.name_ua AS license_name_full
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        WHERE ss.id = ?
    """, (sheet_id,)).fetchone()
    if not sheet or not user_can_access_group(conn, session, sheet['group_id']):
        conn.close()
        flash('Відомість не знайдено', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))
    sheet = dict(sheet)

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'save_meta' and sheet['status'] not in ('locked',):
            fields = {
                'discipline_name': (request.form.get('discipline_name') or '').strip(),
                'work_topic': (request.form.get('work_topic') or '').strip() or None,
                'entity_kind': request.form.get('entity_kind') or sheet['entity_kind'],
                'control_form': request.form.get('control_form') or sheet['control_form'],
                'hours_credits': (request.form.get('hours_credits') or '').strip() or None,
                'teacher_name': (request.form.get('teacher_name') or '').strip() or None,
                'exam_date': (request.form.get('exam_date') or '').strip() or None,
                'statement_number': (request.form.get('statement_number') or '').strip() or None,
            }
            if fields['discipline_name']:
                conn.execute("""
                    UPDATE session_sheets SET
                        discipline_name=?, work_topic=?, entity_kind=?, control_form=?,
                        hours_credits=?, teacher_name=?, exam_date=?, statement_number=?
                    WHERE id=?
                """, (
                    fields['discipline_name'], fields['work_topic'], fields['entity_kind'], fields['control_form'],
                    fields['hours_credits'], fields['teacher_name'], fields['exam_date'],
                    fields['statement_number'], sheet_id,
                ))
                conn.commit()
                flash('Метадані збережено', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'refresh_students' and sheet['status'] in ('draft', 'sent'):
            n = populate_sheet_students(conn, sheet_id, sheet['group_id'], sheet['license_id'])
            conn.commit()
            flash(f'Список студентів оновлено ({n})', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'issue_token':
            days = request.form.get('token_days', type=int) or 21
            token = issue_sheet_token(conn, sheet_id, days=days)
            conn.commit()
            log_action(current_username(), f"видано токен відомості #{sheet_id}")
            link = f"{PUBLIC_APPLY_BASE_URL.rstrip('/')}/session-grade/{token}"
            flash('Посилання для викладача створено', 'success')
            students = _load_students(conn, sheet_id)
            sheet = dict(conn.execute("SELECT ss.*, es.group_id FROM session_sheets ss JOIN exam_sessions es ON es.id=ss.session_id WHERE ss.id=?", (sheet_id,)).fetchone())
            # re-fetch full
            sheet = dict(conn.execute("""
                SELECT ss.*, es.group_id, es.name AS session_name, es.academic_year,
                       es.semester, g.name AS group_name, g.course, g.study_form,
                       g.specialty, g.degree_level, g.knowledge_area,
                       l.short_name_ua AS license_name
                FROM session_sheets ss
                JOIN exam_sessions es ON es.id = ss.session_id
                JOIN groups g ON g.id = es.group_id
                LEFT JOIN institution_licenses l ON l.id = ss.license_id
                WHERE ss.id = ?
            """, (sheet_id,)).fetchone())
            can_transfer = session.get('is_admin') or 'transfer_exam_grades' in (session.get('permissions') or [])
            conn.close()
            return render_template(
                'exam_sessions/sheet_detail.html',
                sheet=sheet, students=students, teacher_link=link, can_transfer=can_transfer,
            )

        if action == 'save_grades' and sheet['status'] not in ('locked',):
            # збираємо поля по student_id
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
                    """, (topic, sheet_id, sid))
                mid = parse_points(fields.get('points_mid'))
                cur = parse_points(fields.get('points_current'))
                exam = parse_points(fields.get('points_exam'))
                explicit = fields.get('points')
                # компоненти → сума; інакше пряма підсумкова
                total = compute_final_points(mid, cur, exam, explicit)
                nat, ects = points_to_national_ects(total, sheet['control_form'])
                conn.execute("""
                    UPDATE session_sheet_students
                    SET points_mid=?, points_current=?, points_exam=?,
                        points=?, national_grade=?, ects=?
                    WHERE sheet_id=? AND student_id=?
                """, (mid, cur, exam, total, nat, ects, sheet_id, sid))
            if sheet['status'] == 'draft':
                conn.execute(
                    "UPDATE session_sheets SET status='submitted', submitted_at=datetime('now','localtime') WHERE id=?",
                    (sheet_id,),
                )
            conn.commit()
            flash('Оцінки збережено', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'approve' and sheet['status'] in ('submitted', 'sent', 'approved'):
            conn.execute("""
                UPDATE session_sheets
                SET status='approved', approved_at=datetime('now','localtime'), approved_by=?
                WHERE id=?
            """, (current_username(), sheet_id))
            conn.commit()
            log_action(current_username(), f"затвердив відомість #{sheet_id}")
            flash('Відомість затверджено', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'lock' and sheet['status'] == 'approved':
            conn.execute("UPDATE session_sheets SET status='locked' WHERE id=?", (sheet_id,))
            conn.commit()
            flash('Відомість заблоковано', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        conn.close()
        return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

    students = _load_students(conn, sheet_id)
    teacher_link = None
    if sheet.get('token'):
        teacher_link = f"{PUBLIC_APPLY_BASE_URL.rstrip('/')}/session-grade/{sheet['token']}"
    can_transfer = session.get('is_admin') or 'transfer_exam_grades' in (session.get('permissions') or [])
    conn.close()
    return render_template(
        'exam_sessions/sheet_detail.html',
        sheet=sheet, students=students, teacher_link=teacher_link, can_transfer=can_transfer,
    )


def _load_students(conn, sheet_id):
    return [dict(r) for r in conn.execute("""
        SELECT * FROM session_sheet_students
        WHERE sheet_id = ?
        ORDER BY position, full_name
    """, (sheet_id,)).fetchall()]


# ── Excel ──────────────────────────────────────────────────────────



@exam_sessions_bp.route('/admin/exam-sessions/group-summary')
@permission_required('manage_exam_sessions')
def group_course_summary():
    """Зведена відомість по групі за всі сесії (Excel)."""
    group_id = request.args.get('group_id', type=int)
    year = (request.args.get('academic_year') or '').strip() or None
    if not group_id:
        flash('Оберіть групу', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))

    conn = _conn()
    if not user_can_access_group(conn, session, group_id):
        conn.close()
        flash('Немає доступу', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))

    g = conn.execute("""
        SELECT id, name, course, study_form, degree_level, specialty, specialty_code,
               knowledge_area, institution_name_and_status
        FROM groups WHERE id=?
    """, (group_id,)).fetchone()
    if not g:
        conn.close()
        flash('Групу не знайдено', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))

    sql = "SELECT es.* FROM exam_sessions es WHERE es.group_id=?"
    params = [group_id]
    if year:
        sql += " AND es.academic_year=?"
        params.append(year)
    sql += " ORDER BY es.academic_year, es.semester"
    sessions = [dict(r) for r in conn.execute(sql, params).fetchall()]

    sessions_blocks = []
    student_map = {}

    for es in sessions:
        sheets = [dict(r) for r in conn.execute("""
            SELECT ss.id, ss.discipline_name, ss.control_form, ss.hours_credits, ss.license_id,
                   l.short_name_ua AS license_name
            FROM session_sheets ss
            LEFT JOIN institution_licenses l ON l.id = ss.license_id
            WHERE ss.session_id=?
            ORDER BY ss.discipline_name, ss.license_id
        """, (es['id'],)).fetchall()]

        grades = {}
        for sh in sheets:
            rows = conn.execute("""
                SELECT student_id, full_name, points, national_grade, ects
                FROM session_sheet_students
                WHERE sheet_id=?
                ORDER BY position, full_name
            """, (sh['id'],)).fetchall()
            for r in rows:
                sid = r['student_id']
                if sid not in student_map:
                    student_map[sid] = r['full_name']
                grades.setdefault(sid, {})[sh['id']] = {
                    'national': r['national_grade'],
                    'points': r['points'],
                    'ects': r['ects'],
                }

        for sh in sheets:
            if sh.get('license_name'):
                sh['discipline_name'] = f"{sh['discipline_name']} ({sh['license_name']})"

        sessions_blocks.append({
            'academic_year': es.get('academic_year'),
            'semester': es.get('semester'),
            'sheets': sheets,
            'grades': grades,
        })

    conn.close()

    students = [{'id': sid, 'full_name': name} for sid, name in student_map.items()]
    students.sort(key=lambda x: (x['full_name'] or '').lower())

    meta = dict(g)
    meta['group_name'] = g['name']
    meta['students'] = students
    meta['title'] = f"Зведена відомість · {g['name']}" + (f" · {year}" if year else '')

    if not sessions_blocks or not any(b['sheets'] for b in sessions_blocks):
        flash('Немає відомостей для зведення', 'warning')
        q = {'group_id': group_id}
        if year:
            q['academic_year'] = year
        return redirect(url_for('exam_sessions.sessions_list', **q))

    bio = build_group_course_summary_xlsx(meta, sessions_blocks)
    fname = f"zvedena_{g['name']}"
    if year:
        fname += f"_{year}"
    fname += ".xlsx"
    fname = ''.join(c if c.isalnum() or c in '._- ' else '_' for c in fname)
    return send_file(
        bio,
        as_attachment=True,
        download_name=fname,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@exam_sessions_bp.route('/admin/exam-sessions/sheets/<int:sheet_id>/excel')
@permission_required('manage_exam_sessions')
def sheet_excel(sheet_id):
    conn = _conn()
    sheet = conn.execute("""
        SELECT ss.*, es.group_id, es.academic_year, es.semester,
               g.name AS group_name, g.course, g.study_form, g.specialty,
               g.degree_level, g.knowledge_area, g.specialty_code,
               g.institution_name_and_status,
               l.short_name_ua AS license_name
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        WHERE ss.id = ?
    """, (sheet_id,)).fetchone()
    if not sheet or not user_can_access_group(conn, session, sheet['group_id']):
        conn.close()
        flash('Немає доступу', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))
    students = _load_students(conn, sheet_id)
    conn.close()

    bio = build_statement_xlsx(dict(sheet), students)
    fname = f"vidomist_{sheet_id}_{sheet['discipline_name'][:40]}.xlsx"
    fname = ''.join(c if c.isalnum() or c in '._- ' else '_' for c in fname)
    return send_file(
        bio,
        as_attachment=True,
        download_name=fname,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@exam_sessions_bp.route('/admin/exam-sessions/sheets/<int:sheet_id>/word')
@permission_required('manage_exam_sessions')
def sheet_word(sheet_id):
    """Конкретна відомість (одна дисципліна) у Word."""
    conn = _conn()
    sheet = conn.execute("""
        SELECT ss.*, es.group_id, es.academic_year, es.semester,
               g.name AS group_name, g.course, g.study_form, g.specialty,
               g.degree_level, g.knowledge_area, g.specialty_code,
               g.institution_name_and_status,
               l.short_name_ua AS license_name
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        WHERE ss.id = ?
    """, (sheet_id,)).fetchone()
    if not sheet or not user_can_access_group(conn, session, sheet['group_id']):
        conn.close()
        flash('Немає доступу', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))
    students = _load_students(conn, sheet_id)
    conn.close()

    bio = build_sheet_docx(dict(sheet), students)
    fname = f"vidomist_{sheet_id}_{(sheet['discipline_name'] or '')[:40]}.docx"
    fname = ''.join(c if c.isalnum() or c in '._- ' else '_' for c in fname)
    return send_file(
        bio,
        as_attachment=True,
        download_name=fname,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


@exam_sessions_bp.route('/admin/exam-sessions/<int:session_id>/summary-word')
@permission_required('manage_exam_sessions')
def session_summary_word(session_id):
    """Загальна відомість сесії (усі дисципліни × студенти) у Word."""
    conn = _conn()
    es = conn.execute("""
        SELECT es.*, g.name AS group_name, g.course, g.study_form, g.specialty,
               g.degree_level, g.knowledge_area, g.specialty_code,
               g.institution_name_and_status
        FROM exam_sessions es
        JOIN groups g ON g.id = es.group_id
        WHERE es.id = ?
    """, (session_id,)).fetchone()
    if not es or not user_can_access_group(conn, session, es['group_id']):
        conn.close()
        flash('Немає доступу', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))

    sheets = conn.execute("""
        SELECT ss.*, l.short_name_ua AS license_name
        FROM session_sheets ss
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        WHERE ss.session_id = ?
        ORDER BY ss.discipline_name, ss.license_id
    """, (session_id,)).fetchall()
    sheets = [dict(s) for s in sheets]
    by_students = {}
    for sh in sheets:
        by_students[sh['id']] = _load_students(conn, sh['id'])
    meta = dict(es)
    conn.close()

    bio = build_session_summary_docx(meta, sheets, by_students)
    fname = f"zagalna_sesia_{session_id}_{meta.get('group_name') or ''}.docx"
    fname = ''.join(c if c.isalnum() or c in '._- ' else '_' for c in fname)
    return send_file(
        bio,
        as_attachment=True,
        download_name=fname,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )

