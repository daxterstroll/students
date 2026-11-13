"""
Адмін/відповідальний: сесії, CRUD відомостей, картотека викладачів, статистика, Excel.
"""
from __future__ import annotations

import secrets
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
    recompute_sheet_status, generate_teacher_token, teacher_stats,
)
from routes.exam_sessions.history import log_history
from routes.exam_sessions.excel_export import build_statement_xlsx, build_group_course_summary_xlsx
from routes.exam_sessions.word_export import build_sheet_docx, build_session_summary_docx


def _conn():
    conn = get_db()
    conn.row_factory = sqlite3.Row
    ensure_exam_session_tables(conn)
    return conn


def _format_exam_date_ua(iso_date_str):
    if not iso_date_str:
        return None
    s = str(iso_date_str).strip()
    import re as _re
    m = _re.match(r'^(\d{4})-(\d{2})-(\d{2})$', s)
    if not m:
        return s
    year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
    months_ua = {
        1: 'січня', 2: 'лютого', 3: 'березня', 4: 'квітня',
        5: 'травня', 6: 'червня', 7: 'липня', 8: 'серпня',
        9: 'вересня', 10: 'жовтня', 11: 'листопада', 12: 'грудня',
    }
    name = months_ua.get(month)
    if not name:
        return s
    return f"{day} {name} {year}"


def _to_iso_date(date_str):
    if not date_str:
        return ''
    s = str(date_str).strip()
    import re as _re
    if _re.match(r'^\d{4}-\d{2}-\d{2}$', s):
        return s

    months_ua = {
        'січня': 1, 'лютого': 2, 'березня': 3, 'квітня': 4,
        'травня': 5, 'червня': 6, 'липня': 7, 'серпня': 8,
        'вересня': 9, 'жовтня': 10, 'листопада': 11, 'грудня': 12,
    }
    m = _re.match(r'^(\d{1,2})\s+([а-яіїєґ]+)\s+(\d{4})$', s.lower())
    if m:
        day = int(m.group(1))
        month = months_ua.get(m.group(2))
        year = int(m.group(3))
        if month:
            return f"{year:04d}-{month:02d}-{day:02d}"

    m = _re.match(r'^(\d{1,2})[./](\d{1,2})[./](\d{4})$', s)
    if m:
        day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{year:04d}-{month:02d}-{day:02d}"
    return ''


def _load_students(conn, sheet_id):
    return [dict(r) for r in conn.execute("""
        SELECT * FROM session_sheet_students
        WHERE sheet_id = ?
        ORDER BY position, full_name
    """, (sheet_id,)).fetchall()]


# ═══════════════════════════════════════════════════════════════════
# КАРТОТЕКА ВИКЛАДАЧІВ
# ═══════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/admin/teachers')
@permission_required('manage_exam_sessions')
def teachers_list():
    """Картотека викладачів: список з посиланнями і статистикою."""
    conn = _conn()

    rows = conn.execute("""
        SELECT t.*,
               u.username AS linked_username,
               (SELECT COUNT(*) FROM session_sheets ss
                WHERE ss.teacher_ref_id = t.id) AS sheets_total,
               (SELECT COUNT(*) FROM session_sheets ss
                WHERE ss.teacher_ref_id = t.id
                  AND ss.status IN ('submitted', 'approved', 'locked')) AS sheets_submitted
        FROM teachers t
        LEFT JOIN users u ON u.id = t.user_id
        ORDER BY t.is_active DESC, t.full_name COLLATE UKRAINIAN
    """).fetchall()

    teachers = []
    for r in rows:
        d = dict(r)
        total = d['sheets_total'] or 0
        submitted = d['sheets_submitted'] or 0
        d['sheets_pending'] = total - submitted
        d['percent'] = round(submitted / total * 100) if total else 0
        teachers.append(d)

    conn.close()
    return render_template('exam_sessions/teachers_list.html', teachers=teachers)


@exam_sessions_bp.route('/admin/teachers/new', methods=['GET', 'POST'])
@permission_required('manage_exam_sessions')
def teacher_new():
    conn = _conn()

    if request.method == 'POST':
        full_name = (request.form.get('full_name') or '').strip()
        short_name = (request.form.get('short_name') or '').strip() or None
        email = (request.form.get('email') or '').strip() or None
        phone = (request.form.get('phone') or '').strip() or None
        department = (request.form.get('department') or '').strip() or None
        user_id = request.form.get('user_id', type=int) or None
        notes = (request.form.get('notes') or '').strip() or None

        if not full_name:
            flash("Вкажіть ПІБ викладача", "error")
        else:
            token = generate_teacher_token()
            cur = conn.execute("""
                INSERT INTO teachers
                    (full_name, short_name, email, phone, department,
                     user_id, token, notes, is_active, created_by)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
            """, (full_name, short_name, email, phone, department,
                  user_id, token, notes, current_username()))
            tid = cur.lastrowid
            conn.commit()
            log_action(
                current_username(),
                f"створив викладача в картотеці: {full_name} (ID {tid})",
            )
            conn.close()
            flash(f'Викладача «{full_name}» додано. Постійне посилання створено.', 'success')
            return redirect(url_for('exam_sessions.teacher_edit', teacher_id=tid))

    users = conn.execute(
        "SELECT id, username FROM users ORDER BY username COLLATE UKRAINIAN"
    ).fetchall()
    conn.close()
    return render_template(
        'exam_sessions/teacher_form.html',
        teacher=None, users=users,
    )


@exam_sessions_bp.route('/admin/teachers/<int:teacher_id>/edit', methods=['GET', 'POST'])
@permission_required('manage_exam_sessions')
def teacher_edit(teacher_id):
    conn = _conn()
    teacher = conn.execute(
        "SELECT * FROM teachers WHERE id = ?", (teacher_id,)
    ).fetchone()
    if not teacher:
        conn.close()
        flash("Викладача не знайдено", "error")
        return redirect(url_for('exam_sessions.teachers_list'))
    teacher = dict(teacher)

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'save':
            full_name = (request.form.get('full_name') or '').strip()
            short_name = (request.form.get('short_name') or '').strip() or None
            email = (request.form.get('email') or '').strip() or None
            phone = (request.form.get('phone') or '').strip() or None
            department = (request.form.get('department') or '').strip() or None
            user_id = request.form.get('user_id', type=int) or None
            notes = (request.form.get('notes') or '').strip() or None
            is_active = 1 if request.form.get('is_active') else 0

            if not full_name:
                flash("Вкажіть ПІБ викладача", "error")
            else:
                conn.execute("""
                    UPDATE teachers SET
                        full_name=?, short_name=?, email=?, phone=?,
                        department=?, user_id=?, notes=?, is_active=?
                    WHERE id=?
                """, (full_name, short_name, email, phone, department,
                      user_id, notes, is_active, teacher_id))
                conn.commit()
                log_action(
                    current_username(),
                    f"редагував викладача: {full_name} (ID {teacher_id})",
                )
                flash("Дані викладача оновлено", "success")
            conn.close()
            return redirect(url_for('exam_sessions.teacher_edit', teacher_id=teacher_id))

        if action == 'reset_token':
            new_token = generate_teacher_token()
            conn.execute("UPDATE teachers SET token = ? WHERE id = ?", (new_token, teacher_id))
            conn.commit()
            log_action(
                current_username(),
                f"скинув токен викладача {teacher['full_name']} (ID {teacher_id})",
            )
            conn.close()
            flash(
                f"Створено нове посилання для {teacher['full_name']}. "
                f"Стара ссылка більше не працює.",
                "success",
            )
            return redirect(url_for('exam_sessions.teacher_edit', teacher_id=teacher_id))

        if action == 'delete':
            if teacher['user_id']:
                flash(
                    f"Неможливо видалити — викладач прив'язаний до користувача. "
                    f"Спершу зніміть прив'язку або деактивуйте.",
                    "error",
                )
            else:
                sheets_count = conn.execute(
                    "SELECT COUNT(*) FROM session_sheets WHERE teacher_ref_id = ?",
                    (teacher_id,),
                ).fetchone()[0]
                if sheets_count > 0:
                    flash(
                        f"Неможливо видалити — у викладача {sheets_count} відомостей. "
                        f"Деактивуйте замість видалення.",
                        "error",
                    )
                else:
                    conn.execute("DELETE FROM teachers WHERE id = ?", (teacher_id,))
                    conn.commit()
                    log_action(
                        current_username(),
                        f"видалив викладача: {teacher['full_name']} (ID {teacher_id})",
                    )
                    conn.close()
                    flash(f"Викладача «{teacher['full_name']}» видалено", "success")
                    return redirect(url_for('exam_sessions.teachers_list'))
            conn.close()
            return redirect(url_for('exam_sessions.teacher_edit', teacher_id=teacher_id))

    users = conn.execute(
        "SELECT id, username FROM users ORDER BY username COLLATE UKRAINIAN"
    ).fetchall()

    # Отомости цього викладача (для блоку знизу)
    stats = teacher_stats(conn, teacher_id)
    conn.close()

    return render_template(
        'exam_sessions/teacher_form.html',
        teacher=teacher, users=users, stats=stats,
    )


# ═══════════════════════════════════════════════════════════════════
# СЕСІЇ
# ═══════════════════════════════════════════════════════════════════

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
        elif not academic_year or semester not in (1, 2):
            flash('Вкажіть навчальний рік і семестр (1 або 2)', 'error')
        else:
            if not name:
                name = f"Сесія {semester} семестру {academic_year}"
            cur = conn.execute("""
                INSERT INTO exam_sessions
                    (group_id, academic_year, semester, name, status, notes, created_by)
                VALUES (?, ?, ?, ?, 'open', ?, ?)
            """, (group_id, academic_year, semester, name, notes, current_username()))
            sid = cur.lastrowid
            log_history(
                conn,
                entity_type='session', entity_id=sid,
                session_id=sid,
                action='create',
                details={
                    'name': name,
                    'group_id': group_id,
                    'academic_year': academic_year,
                    'semester': semester,
                },
            )
            conn.commit()
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
            academic_year = (request.form.get('academic_year') or '').strip() or es['academic_year']
            semester = request.form.get('semester', type=int) or es['semester']

            sheets_count = conn.execute(
                "SELECT COUNT(*) FROM session_sheets WHERE session_id = ?",
                (session_id,),
            ).fetchone()[0]

            if sheets_count > 0:
                if academic_year != es['academic_year']:
                    flash('Не можна змінити навчальний рік — у сесії вже є відомості. '
                          'Спочатку видаліть їх.', 'error')
                    academic_year = es['academic_year']
                if semester != es['semester']:
                    flash('Не можна змінити семестр — у сесії вже є відомості. '
                          'Спочатку видаліть їх.', 'error')
                    semester = es['semester']

            changes = {}
            if name != es['name']:
                changes['name'] = (es['name'], name)
            if status != es['status']:
                changes['status'] = (es['status'], status)
            if notes != (es['notes'] or None):
                changes['notes'] = (es['notes'] or None, notes)
            if academic_year != es['academic_year']:
                changes['academic_year'] = (es['academic_year'], academic_year)
            if semester != es['semester']:
                changes['semester'] = (es['semester'], semester)

            if changes:
                conn.execute("""
                    UPDATE exam_sessions
                    SET name=?, status=?, notes=?, academic_year=?, semester=?,
                        updated_at=datetime('now','localtime')
                    WHERE id=?
                """, (name, status, notes, academic_year, semester, session_id))
                for field, (old, new) in changes.items():
                    log_history(
                        conn,
                        entity_type='session', entity_id=session_id,
                        session_id=session_id,
                        action='update',
                        field_name=field, old_value=old, new_value=new,
                    )
                conn.commit()
                flash('Сесію оновлено', 'success')
            else:
                flash('Немає змін', 'info')
            conn.close()
            return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

        if action in ('delete_session', 'delete_session_with_links'):
            if not (session.get('is_admin')
                    or 'delete_exam_sessions' in (session.get('permissions') or [])):
                conn.close()
                flash('Немає права видаляти сесії', 'error')
                return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

            remove_grades = request.form.get('remove_grades') == '1'

            linked_sheets = conn.execute("""
                SELECT id, linked_entity_type, linked_entity_id
                FROM session_sheets
                WHERE session_id = ? AND linked_entity_id IS NOT NULL
            """, (session_id,)).fetchall()

            if action == 'delete_session' and linked_sheets:
                conn.close()
                flash(f'Неможливо видалити: {len(linked_sheets)} відомостей уже '
                      f'перенесено в навчальний план. Використайте «Видалити сесію» '
                      f'з галочкою «Видалити оцінки з плану».', 'error')
                return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

            sheets_snapshot = conn.execute("""
                SELECT id, discipline_name FROM session_sheets WHERE session_id = ?
            """, (session_id,)).fetchall()
            session_snapshot = {
                'name': es['name'],
                'group_id': es['group_id'],
                'group_name': es['group_name'],
                'academic_year': es['academic_year'],
                'semester': es['semester'],
                'sheets_count': len(sheets_snapshot),
                'sheets': [{'id': s['id'], 'discipline': s['discipline_name']}
                           for s in sheets_snapshot],
                'linked_sheets_count': len(linked_sheets),
                'grades_removed_from_plan': 0,
            }

            try:
                if linked_sheets and remove_grades:
                    total_removed = 0
                    for sh in linked_sheets:
                        total_removed += _unlink_grades_from_plan(
                            conn, sh['id'],
                            sh['linked_entity_type'], sh['linked_entity_id'],
                            session_id, sh['linked_entity_type'],
                        )
                    session_snapshot['grades_removed_from_plan'] = total_removed

                conn.execute("DELETE FROM exam_sessions WHERE id = ?", (session_id,))
                log_history(
                    conn,
                    entity_type='session', entity_id=session_id,
                    session_id=session_id,
                    action='delete',
                    details=session_snapshot,
                )
                conn.commit()
            except Exception as e:
                conn.rollback()
                conn.close()
                logger.error(f'delete session {session_id}: {e}', exc_info=True)
                flash(f'Помилка видалення: {e}', 'error')
                return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

            conn.close()
            log_action(
                current_username(),
                f"ВИДАЛИВ сесію #{session_id} «{session_snapshot['name']}» "
                f"({session_snapshot['sheets_count']} відомостей, "
                f"з них прив'язаних: {session_snapshot['linked_sheets_count']}, "
                f"оцінок видалено з плану: {session_snapshot['grades_removed_from_plan']})",
            )
            flash(
                f'Сесію «{session_snapshot["name"]}» видалено '
                f'разом з {session_snapshot["sheets_count"]} відомостями.'
                + (f' З плану видалено {session_snapshot["grades_removed_from_plan"]} оцінок.'
                   if session_snapshot['grades_removed_from_plan'] else ''),
                'success',
            )
            return redirect(url_for('exam_sessions.sessions_list'))

    sheets = [dict(r) for r in conn.execute("""
        SELECT ss.*,
               t.short_name AS teacher_short,
               t.full_name AS teacher_full,
               l.short_name_ua AS license_name,
               (SELECT COUNT(*) FROM session_sheet_students x WHERE x.sheet_id = ss.id) AS students_count,
               (SELECT COUNT(*) FROM session_sheet_students x
                WHERE x.sheet_id = ss.id AND x.points IS NOT NULL) AS graded_count
        FROM session_sheets ss
        LEFT JOIN teachers t ON t.id = ss.teacher_ref_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        WHERE ss.session_id = ?
        ORDER BY ss.discipline_name, ss.license_id
    """, (session_id,)).fetchall()]

    licenses = licenses_present_in_group(conn, es['group_id'])
    can_transfer = session.get('is_admin') or 'transfer_exam_grades' in (session.get('permissions') or [])
    teachers = conn.execute(
        "SELECT id, full_name, short_name FROM teachers "
        "WHERE is_active = 1 ORDER BY full_name COLLATE UKRAINIAN"
    ).fetchall()
    conn.close()
    return render_template(
        'exam_sessions/session_detail.html',
        es=es, sheets=sheets, licenses=licenses, can_transfer=can_transfer,
        teachers=teachers,
    )


# ═══════════════════════════════════════════════════════════════════
# ВІДОМІСТЬ: створення / деталі
# ═══════════════════════════════════════════════════════════════════

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
        teacher_ref_id = request.form.get('teacher_ref_id', type=int) or None
        exam_date = _format_exam_date_ua(request.form.get('exam_date') or '')
        statement_number = (request.form.get('statement_number') or '').strip() or None
        lic_raw = request.form.get('license_id')
        license_id = int(lic_raw) if lic_raw and lic_raw.isdigit() else None

        # Snapshot teacher_name из картотеки
        teacher_name = None
        if teacher_ref_id:
            t = conn.execute(
                "SELECT full_name, short_name FROM teachers WHERE id = ?",
                (teacher_ref_id,),
            ).fetchone()
            if t:
                teacher_name = t['short_name'] or t['full_name']

        if entity_kind not in ('subject', 'practice', 'coursework', 'attestation'):
            entity_kind = 'subject'
        if control_form not in ('Залік', 'Екзамен'):
            control_form = 'Залік'

        if not discipline_name:
            flash('Вкажіть назву дисципліни', 'error')
        elif not teacher_ref_id:
            flash('Оберіть викладача з картотеки', 'error')
        else:
            cur = conn.execute("""
                INSERT INTO session_sheets (
                    session_id, discipline_name, work_topic, entity_kind, control_form,
                    hours_credits, teacher_name, teacher_ref_id, exam_date, statement_number,
                    license_id, status, created_by
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'draft', ?)
            """, (
                session_id, discipline_name, work_topic, entity_kind, control_form,
                hours_credits, teacher_name, teacher_ref_id, exam_date, statement_number,
                license_id, current_username(),
            ))
            sheet_id = cur.lastrowid
            n = populate_sheet_students(conn, sheet_id, es['group_id'], license_id)
            log_history(
                conn,
                entity_type='sheet', entity_id=sheet_id,
                session_id=session_id, sheet_id=sheet_id,
                action='create',
                details={
                    'discipline_name': discipline_name,
                    'entity_kind': entity_kind,
                    'control_form': control_form,
                    'license_id': license_id,
                    'teacher_ref_id': teacher_ref_id,
                    'students_count': n,
                },
            )
            conn.commit()
            log_action(
                current_username(),
                f"створив відомість #{sheet_id} «{discipline_name}» у сесії #{session_id} "
                f"({n} студ., викладач: {teacher_name})",
            )
            conn.close()
            flash(
                f'Відомість створено. Студентів у списку: {n}. '
                f'Відомість автоматично з\'явилась у викладача в кабінеті.',
                'success',
            )
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

    teachers = conn.execute(
        "SELECT id, full_name, short_name, department FROM teachers "
        "WHERE is_active = 1 ORDER BY full_name COLLATE UKRAINIAN"
    ).fetchall()
    conn.close()
    return render_template(
        'exam_sessions/sheet_form.html',
        es=es, licenses=licenses, sheet=None, teachers=teachers,
    )


@exam_sessions_bp.route('/admin/exam-sessions/sheets/<int:sheet_id>', methods=['GET', 'POST'])
@permission_required('manage_exam_sessions')
def sheet_detail(sheet_id):
    conn = _conn()
    sheet = conn.execute("""
        SELECT ss.*, es.group_id, es.name AS session_name, es.academic_year,
               es.semester, g.name AS group_name, g.course, g.study_form,
               g.specialty, g.degree_level, g.knowledge_area,
               l.short_name_ua AS license_name, l.name_ua AS license_name_full,
               t.full_name AS teacher_full, t.short_name AS teacher_short,
               t.token AS teacher_token
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        LEFT JOIN teachers t ON t.id = ss.teacher_ref_id
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
            teacher_ref_id = request.form.get('teacher_ref_id', type=int) or sheet['teacher_ref_id']
            teacher_name = sheet['teacher_name']
            if teacher_ref_id != sheet['teacher_ref_id']:
                t = conn.execute(
                    "SELECT full_name, short_name FROM teachers WHERE id = ?",
                    (teacher_ref_id,),
                ).fetchone()
                if t:
                    teacher_name = t['short_name'] or t['full_name']

            fields = {
                'discipline_name': (request.form.get('discipline_name') or '').strip(),
                'work_topic': (request.form.get('work_topic') or '').strip() or None,
                'entity_kind': request.form.get('entity_kind') or sheet['entity_kind'],
                'control_form': request.form.get('control_form') or sheet['control_form'],
                'hours_credits': (request.form.get('hours_credits') or '').strip() or None,
                'teacher_name': teacher_name,
                'teacher_ref_id': teacher_ref_id,
                'exam_date': _format_exam_date_ua(request.form.get('exam_date') or ''),
                'statement_number': (request.form.get('statement_number') or '').strip() or None,
            }

            if fields['discipline_name']:
                changes = {
                    f: (sheet.get(f), v)
                    for f, v in fields.items()
                    if sheet.get(f) != v
                }

                conn.execute("""
                    UPDATE session_sheets SET
                        discipline_name=?, work_topic=?, entity_kind=?, control_form=?,
                        hours_credits=?, teacher_name=?, teacher_ref_id=?,
                        exam_date=?, statement_number=?
                    WHERE id=?
                """, (
                    fields['discipline_name'], fields['work_topic'],
                    fields['entity_kind'], fields['control_form'],
                    fields['hours_credits'], fields['teacher_name'], fields['teacher_ref_id'],
                    fields['exam_date'], fields['statement_number'], sheet_id,
                ))

                for field, (old, new) in changes.items():
                    log_history(
                        conn,
                        entity_type='sheet', entity_id=sheet_id,
                        session_id=sheet['session_id'], sheet_id=sheet_id,
                        action='update', field_name=field,
                        old_value=old, new_value=new,
                    )
                conn.commit()
                flash('Метадані збережено', 'success')
            else:
                flash("Назва дисципліни обов'язкова", 'error')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'refresh_students' and sheet['status'] in ('draft', 'sent', 'in_progress'):
            n = populate_sheet_students(conn, sheet_id, sheet['group_id'], sheet['license_id'])
            log_history(
                conn,
                entity_type='sheet', entity_id=sheet_id,
                session_id=sheet['session_id'], sheet_id=sheet_id,
                action='refresh',
                details={'students_count': n},
            )
            conn.commit()
            flash(f'Список студентів оновлено ({n})', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'save_grades' and sheet['status'] not in ('locked',):
            current_rows = conn.execute("""
                SELECT student_id, full_name,
                       points_mid, points_current, points_exam,
                       points, national_grade, ects, work_topic
                FROM session_sheet_students
                WHERE sheet_id = ?
            """, (sheet_id,)).fetchall()
            current_by_sid = {r['student_id']: dict(r) for r in current_rows}

            by_sid = {}
            for key, val in request.form.items():
                for prefix in ('points_mid_', 'points_current_', 'points_exam_', 'points_', 'work_topic_'):
                    if key.startswith(prefix):
                        sid = key[len(prefix):]
                        by_sid.setdefault(sid, {})[prefix.rstrip('_')] = val
                        break

            for sid, fields in by_sid.items():
                old = current_by_sid.get(int(sid)) or {}
                student_name = old.get('full_name') or f'student#{sid}'

                if 'work_topic' in fields:
                    topic = (fields.get('work_topic') or '').strip() or None
                    if topic != (old.get('work_topic') or None):
                        log_history(
                            conn,
                            entity_type='sheet_student', entity_id=int(sid),
                            session_id=sheet['session_id'], sheet_id=sheet_id,
                            action='update', field_name='work_topic',
                            old_value=old.get('work_topic'), new_value=topic,
                            details={'student_name': student_name},
                        )
                        conn.execute(
                            "UPDATE session_sheet_students SET work_topic=? "
                            "WHERE sheet_id=? AND student_id=?",
                            (topic, sheet_id, sid),
                        )

                mid = parse_points(fields.get('points_mid'))
                cur = parse_points(fields.get('points_current'))
                exam = parse_points(fields.get('points_exam'))
                explicit = fields.get('points')
                total = compute_final_points(mid, cur, exam, explicit)
                nat, ects = points_to_national_ects(total, sheet['control_form'])

                new_vals = {
                    'points_mid': mid, 'points_current': cur, 'points_exam': exam,
                    'points': total, 'national_grade': nat, 'ects': ects,
                }
                field_changes = {
                    f: (old.get(f), v) for f, v in new_vals.items()
                    if old.get(f) != v
                }
                if field_changes:
                    for field, (old_v, new_v) in field_changes.items():
                        log_history(
                            conn,
                            entity_type='sheet_student', entity_id=int(sid),
                            session_id=sheet['session_id'], sheet_id=sheet_id,
                            action='grade_set', field_name=field,
                            old_value=old_v, new_value=new_v,
                            details={'student_name': student_name},
                        )
                    conn.execute("""
                        UPDATE session_sheet_students
                        SET points_mid=?, points_current=?, points_exam=?,
                            points=?, national_grade=?, ects=?
                        WHERE sheet_id=? AND student_id=?
                    """, (mid, cur, exam, total, nat, ects, sheet_id, sid))

            # Перераховуємо статус за фактичною заповненістю
            recompute_sheet_status(conn, sheet_id)
            conn.commit()
            flash('Оцінки збережено', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'approve' and sheet['status'] in ('submitted', 'in_progress', 'sent', 'approved'):
            # Дозволяємо затвердити навіть частково заповнену — адмін бачить і вирішує
            conn.execute("""
                UPDATE session_sheets
                SET status='approved', approved_at=datetime('now','localtime'), approved_by=?
                WHERE id=?
            """, (current_username(), sheet_id))
            log_history(
                conn,
                entity_type='sheet', entity_id=sheet_id,
                session_id=sheet['session_id'], sheet_id=sheet_id,
                action='approve',
                old_value=sheet['status'], new_value='approved',
            )
            conn.commit()
            log_action(current_username(), f"затвердив відомість #{sheet_id}")
            flash('Відомість затверджено', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'lock' and sheet['status'] == 'approved':
            conn.execute("UPDATE session_sheets SET status='locked' WHERE id=?", (sheet_id,))
            log_history(
                conn,
                entity_type='sheet', entity_id=sheet_id,
                session_id=sheet['session_id'], sheet_id=sheet_id,
                action='lock',
                old_value='approved', new_value='locked',
            )
            conn.commit()
            flash('Відомість заблоковано', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'unlink_from_plan':
            if not (session.get('is_admin')
                    or 'manage_exam_sessions' in (session.get('permissions') or [])):
                conn.close()
                flash('Немає права змінювати відомість', 'error')
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

            if not sheet.get('linked_entity_id'):
                conn.close()
                flash("Відомість не прив'язана до плану", 'info')
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

            old_kind = sheet.get('linked_entity_type')
            old_target_id = sheet.get('linked_entity_id')

            removed = _unlink_grades_from_plan(
                conn, sheet_id, old_kind, old_target_id,
                sheet['session_id'], old_kind,
            )

            conn.execute("""
                UPDATE session_sheets
                SET linked_entity_type = NULL,
                    linked_entity_id = NULL,
                    linked_at = NULL,
                    linked_by = NULL
                WHERE id = ?
            """, (sheet_id,))

            log_history(
                conn,
                entity_type='sheet', entity_id=sheet_id,
                session_id=sheet['session_id'], sheet_id=sheet_id,
                action='unlink',
                details={
                    'old_target_kind': old_kind,
                    'old_target_id': old_target_id,
                    'removed_grades': removed,
                    'discipline_name': sheet['discipline_name'],
                },
            )
            conn.commit()
            conn.close()

            log_action(
                current_username(),
                f"відв'язав відомість #{sheet_id} від плану "
                f"({old_kind}#{old_target_id}, видалено оцінок: {removed})",
            )
            flash(
                f"Відомість відв'язано від плану. "
                f"Видалено оцінок: {removed}. "
                f"Тепер можна видалити відомість або перенести в інше місце.",
                'success',
            )
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'force_delete_sheet':
            if not (session.get('is_admin')
                    or 'delete_exam_sessions' in (session.get('permissions') or [])):
                conn.close()
                flash('Немає права видаляти відомості', 'error')
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

            linked_kind = sheet.get('linked_entity_type')
            linked_target = sheet.get('linked_entity_id')
            remove_grades = request.form.get('remove_grades') == '1'

            students_count = conn.execute(
                "SELECT COUNT(*) FROM session_sheet_students WHERE sheet_id = ?",
                (sheet_id,),
            ).fetchone()[0]

            snapshot = {
                'discipline_name': sheet['discipline_name'],
                'session_id': sheet['session_id'],
                'session_name': sheet['session_name'],
                'group_name': sheet['group_name'],
                'students_count': students_count,
                'was_linked': bool(linked_target),
                'linked_kind': linked_kind,
                'linked_target_id': linked_target,
                'grades_removed_from_plan': 0,
            }

            try:
                if linked_target and remove_grades:
                    removed = _unlink_grades_from_plan(
                        conn, sheet_id, linked_kind, linked_target,
                        sheet['session_id'], linked_kind,
                    )
                    snapshot['grades_removed_from_plan'] = removed

                conn.execute("DELETE FROM session_sheets WHERE id = ?", (sheet_id,))
                log_history(
                    conn,
                    entity_type='sheet', entity_id=sheet_id,
                    session_id=sheet['session_id'], sheet_id=sheet_id,
                    action='delete',
                    details=snapshot,
                )
                conn.commit()
            except Exception as e:
                conn.rollback()
                conn.close()
                logger.error(f'force_delete sheet {sheet_id}: {e}', exc_info=True)
                flash(f'Помилка видалення: {e}', 'error')
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

            conn.close()
            log_action(
                current_username(),
                f"ВИДАЛИВ відомість #{sheet_id} «{snapshot['discipline_name']}» "
                f"(примусово, оцінок видалено з плану: "
                f"{snapshot['grades_removed_from_plan']})",
            )
            flash(
                f'Відомість «{snapshot["discipline_name"]}» видалено'
                + (f' (з плану видалено {snapshot["grades_removed_from_plan"]} оцінок).'
                   if snapshot['grades_removed_from_plan'] else '.'),
                'success',
            )
            return redirect(url_for('exam_sessions.session_detail',
                                    session_id=sheet['session_id']))

        if action == 'delete_sheet':
            if not (session.get('is_admin')
                    or 'delete_exam_sessions' in (session.get('permissions') or [])):
                conn.close()
                flash('Немає права видаляти відомості', 'error')
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

            if sheet['linked_entity_id']:
                conn.close()
                flash('Неможливо видалити: відомість уже перенесена в навчальний план.',
                      'error')
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

            snapshot = {
                'discipline_name': sheet['discipline_name'],
                'session_id': sheet['session_id'],
                'session_name': sheet['session_name'],
                'group_name': sheet['group_name'],
                'students_count': conn.execute(
                    "SELECT COUNT(*) FROM session_sheet_students WHERE sheet_id=?",
                    (sheet_id,)
                ).fetchone()[0],
            }

            try:
                conn.execute("DELETE FROM session_sheets WHERE id = ?", (sheet_id,))
                log_history(
                    conn,
                    entity_type='sheet', entity_id=sheet_id,
                    session_id=sheet['session_id'], sheet_id=sheet_id,
                    action='delete',
                    details=snapshot,
                )
                conn.commit()
            except Exception as e:
                conn.rollback()
                conn.close()
                logger.error(f'delete sheet {sheet_id}: {e}', exc_info=True)
                flash(f'Помилка видалення: {e}', 'error')
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

            conn.close()
            log_action(
                current_username(),
                f"ВИДАЛИВ відомість #{sheet_id} «{snapshot['discipline_name']}» "
                f"({snapshot['students_count']} студентів)",
            )
            flash(f'Відомість «{snapshot["discipline_name"]}» видалено.', 'success')
            return redirect(url_for('exam_sessions.session_detail',
                                    session_id=sheet['session_id']))

        conn.close()
        return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

    students = _load_students(conn, sheet_id)
    teacher_link = None
    if sheet.get('teacher_token'):
        teacher_link = f"{PUBLIC_APPLY_BASE_URL.rstrip('/')}/session-grade/teacher/{sheet['teacher_token']}"
    can_transfer = session.get('is_admin') or 'transfer_exam_grades' in (session.get('permissions') or [])
    teachers = conn.execute(
        "SELECT id, full_name, short_name FROM teachers "
        "WHERE is_active = 1 OR id = ? ORDER BY full_name COLLATE UKRAINIAN",
        (sheet['teacher_ref_id'],),
    ).fetchall()
    conn.close()

    sheet['exam_date_iso'] = _to_iso_date(sheet.get('exam_date'))

    return render_template(
        'exam_sessions/sheet_detail.html',
        sheet=sheet, students=students, teacher_link=teacher_link,
        can_transfer=can_transfer, teachers=teachers,
    )


# ═══════════════════════════════════════════════════════════════════
# СТАТИСТИКА ВИКЛАДАЧІВ
# ═══════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/admin/exam-sessions/teacher-stats')
@permission_required('manage_exam_sessions')
def teacher_stats_dashboard():
    """Статистика по всіх викладачах з картотеки."""
    conn = _conn()

    rows = conn.execute("""
        SELECT
            t.id AS teacher_id,
            t.full_name,
            t.short_name,
            t.department,
            t.is_active,
            COUNT(ss.id) AS sheets_total,
            SUM(CASE WHEN ss.status IN ('submitted', 'approved', 'locked') THEN 1 ELSE 0 END) AS sheets_submitted,
            SUM(CASE WHEN ss.status = 'in_progress' THEN 1 ELSE 0 END) AS sheets_in_progress,
            SUM(CASE WHEN ss.status IN ('draft', 'sent') THEN 1 ELSE 0 END) AS sheets_pending
        FROM teachers t
        LEFT JOIN session_sheets ss ON ss.teacher_ref_id = t.id
        GROUP BY t.id
        ORDER BY t.is_active DESC, t.full_name COLLATE UKRAINIAN
    """).fetchall()

    stats = []
    for r in rows:
        total = r['sheets_total'] or 0
        submitted = r['sheets_submitted'] or 0
        stats.append({
            'teacher_id': r['teacher_id'],
            'full_name': r['full_name'],
            'short_name': r['short_name'],
            'department': r['department'],
            'is_active': r['is_active'],
            'total': total,
            'submitted': submitted,
            'in_progress': r['sheets_in_progress'] or 0,
            'pending': r['sheets_pending'] or 0,
            'percent': round(submitted / total * 100) if total else 0,
        })

    conn.close()
    return render_template('exam_sessions/teacher_stats.html', stats=stats)


@exam_sessions_bp.route('/admin/exam-sessions/teacher-stats/<int:teacher_id>')
@permission_required('manage_exam_sessions')
def teacher_stats_detail(teacher_id):
    """Детальна сторінка одного викладача — редірект на картку викладача."""
    return redirect(url_for('exam_sessions.teacher_edit', teacher_id=teacher_id))


# ═══════════════════════════════════════════════════════════════════
# EXCEL / WORD
# ═══════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/admin/exam-sessions/group-summary')
@permission_required('manage_exam_sessions')
def group_course_summary():
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


@exam_sessions_bp.route('/admin/exam-sessions/<int:session_id>/summary-xlsx')
@permission_required('manage_exam_sessions')
def session_summary_xlsx(session_id):
    from routes.exam_sessions.excel_export import build_session_summary_xlsx

    conn = _conn()
    es = conn.execute("""
        SELECT es.*, g.name AS group_name, g.course, g.study_form, g.specialty,
               g.specialty_code, g.degree_level, g.knowledge_area,
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

    bio = build_session_summary_xlsx(meta, sheets, by_students)
    fname = f"zagalna_sesia_{session_id}_{meta.get('group_name') or ''}.xlsx"
    fname = ''.join(c if c.isalnum() or c in '._- ' else '_' for c in fname)

    return send_file(
        bio,
        as_attachment=True,
        download_name=fname,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


def _unlink_grades_from_plan(conn, sheet_id, linked_kind, linked_target_id,
                              session_id, entity_kind):
    student_ids = [
        r['student_id'] for r in conn.execute(
            "SELECT student_id FROM session_sheet_students WHERE sheet_id = ?",
            (sheet_id,),
        ).fetchall()
    ]
    if not student_ids:
        return 0

    placeholders = ','.join('?' * len(student_ids))
    removed = 0

    if linked_kind == 'subject':
        cur = conn.execute(
            f"DELETE FROM grades WHERE subject_id = ? "
            f"AND student_id IN ({placeholders})",
            [linked_target_id] + student_ids,
        )
        removed = cur.rowcount or 0
    elif linked_kind in ('practice', 'coursework', 'attestation'):
        cur = conn.execute(
            f"DELETE FROM activity_grades WHERE entity_type = ? "
            f"AND entity_id = ? AND student_id IN ({placeholders})",
            [linked_kind, linked_target_id] + student_ids,
        )
        removed = cur.rowcount or 0

    return removed