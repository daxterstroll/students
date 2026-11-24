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
from routes.exam_sessions.helpers_extra import (
    ensure_exam_extra_columns,
    create_retake_sheet_if_needed,
    is_failing_or_absent_student,
)
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
    ensure_exam_extra_columns(conn)
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


def _sheet_row_has_grade(st):
    pts = st.get('points')
    nat = (st.get('national_grade') or '').strip()
    ects = (st.get('ects') or '').strip()
    return (pts is not None and str(pts).strip() != '') or bool(nat) or bool(ects)


def _is_retake_sheet(sh):
    return ((sh.get('sheet_kind') or 'main').strip().lower() == 'retake') or bool(sh.get('parent_sheet_id'))


def _merge_summary_sheets_by_discipline(sheets, by_students):
    """Об'єднує відомості однієї дисципліни (різні ліцензії та відомості перездач)
    в одну спільну колонку для загальної відомості сесії.
    Оцінки з відомості перездачі перекривають незадовільні/відсутні оцінки основної
    відомості і позначаються прапорцем is_retake=True.
    """
    from routes.helpers import sort_ukrainian

    ordered_sheets = sorted(
        [dict(sh) if not isinstance(sh, dict) else dict(sh) for sh in (sheets or [])],
        key=lambda s: (1 if _is_retake_sheet(s) else 0, int(s.get('id') or 0)),
    )

    merged_sheets = []
    merged_by_students = {}
    groups_by_key = {}

    for sh_dict in ordered_sheets:
        is_retake = _is_retake_sheet(sh_dict)
        name_norm = ' '.join((sh_dict.get('discipline_name') or '').split()).lower()
        ctrl_norm = ' '.join((sh_dict.get('control_form') or '').split()).lower()
        kind_norm = (sh_dict.get('entity_kind') or 'subject').strip().lower()
        key = (name_norm, ctrl_norm, kind_norm)

        if key not in groups_by_key:
            canon = sh_dict
            groups_by_key[key] = canon
            merged_sheets.append(canon)
            merged_by_students[canon['id']] = {}
        else:
            canon = groups_by_key[key]
            if not canon.get('hours_credits') and sh_dict.get('hours_credits'):
                canon['hours_credits'] = sh_dict.get('hours_credits')

        stu_map = merged_by_students[canon['id']]
        for st in (by_students.get(sh_dict['id']) or []):
            st_dict = dict(st) if not isinstance(st, dict) else dict(st)
            raw_sid = st_dict.get('student_id')
            if raw_sid is not None and str(raw_sid).strip() != '':
                try:
                    sid_key = ('id', int(raw_sid))
                except (TypeError, ValueError):
                    sid_key = ('id', str(raw_sid).strip())
            else:
                sid_key = ('name', ' '.join((st_dict.get('full_name') or '').split()).lower())

            if is_retake:
                if _sheet_row_has_grade(st_dict):
                    prev = stu_map.get(sid_key)
                    st_dict['is_retake'] = True
                    st_dict['retake_sheet_id'] = sh_dict.get('id')
                    st_dict['retake_date'] = sh_dict.get('exam_date')
                    if prev:
                        st_dict['orig_points'] = prev.get('points')
                    stu_map[sid_key] = st_dict
                elif sid_key not in stu_map:
                    stu_map[sid_key] = st_dict
            else:
                if sid_key not in stu_map:
                    stu_map[sid_key] = st_dict
                elif not _sheet_row_has_grade(stu_map[sid_key]) and _sheet_row_has_grade(st_dict):
                    stu_map[sid_key] = st_dict

    out_by_students = {}
    for sh in merged_sheets:
        cid = sh['id']
        rows = list(merged_by_students[cid].values())
        out_by_students[cid] = sort_ukrainian(rows, lambda r: r.get('full_name') or '')
    return merged_sheets, out_by_students


def _approve_sheet_and_spawn_retake(conn, sheet_id: int, session_id: int, username: str):
    """Затверджує відомість і автоматично створює відомість перездачі для незданих/неявок."""
    sh = conn.execute(
        "SELECT id, status, discipline_name, sheet_kind FROM session_sheets WHERE id = ? AND session_id = ?",
        (sheet_id, session_id),
    ).fetchone()
    if not sh:
        return None, None

    conn.execute("""
        UPDATE session_sheets
        SET status='approved', approved_at=datetime('now','localtime'), approved_by=?
        WHERE id=?
    """, (username, sheet_id))
    log_history(
        conn,
        entity_type='sheet', entity_id=sheet_id,
        session_id=session_id, sheet_id=sheet_id,
        action='approve', old_value=sh['status'], new_value='approved',
    )

    retake_info = create_retake_sheet_if_needed(conn, sheet_id, created_by=username)
    if retake_info:
        log_history(
            conn,
            entity_type='sheet', entity_id=retake_info['sheet_id'],
            session_id=session_id, sheet_id=retake_info['sheet_id'],
            action='create_retake',
            details={
                'discipline_name': retake_info['discipline_name'],
                'parent_sheet_id': sheet_id,
                'retake_students_count': retake_info['students_count'],
                'license_id': retake_info.get('license_id'),
                'teacher_ref_id': retake_info.get('teacher_ref_id'),
            },
        )
    return dict(sh), retake_info


def _reopen_sheet_for_teacher(conn, sheet_row, username: str):
    """Повторно відкриває закриту викладачем або затверджену (не заблоковану) відомість."""
    sheet_id = sheet_row['id']
    old_status = sheet_row['status']
    conn.execute("""
        UPDATE session_sheets
        SET teacher_closed = 0,
            teacher_closed_at = NULL,
            approved_at = NULL,
            approved_by = NULL,
            reopened_at = datetime('now', 'localtime'),
            reopened_by = ?
        WHERE id = ?
    """, (username, sheet_id))

    counts = conn.execute("""
        SELECT COUNT(*) AS total_cnt,
               SUM(CASE WHEN points IS NOT NULL THEN 1 ELSE 0 END) AS graded_cnt
        FROM session_sheet_students
        WHERE sheet_id = ?
    """, (sheet_id,)).fetchone()
    graded_cnt = (counts['graded_cnt'] or 0) if counts else 0

    new_status = 'in_progress' if graded_cnt > 0 else 'sent'
    conn.execute("UPDATE session_sheets SET status = ? WHERE id = ?", (new_status, sheet_id))

    log_history(
        conn,
        entity_type='sheet', entity_id=sheet_id,
        session_id=sheet_row['session_id'], sheet_id=sheet_id,
        action='reopen', field_name='status',
        old_value=old_status, new_value=new_status,
        details={
            'reopened_by': username,
            'prev_status': old_status,
            'new_status': new_status,
            'teacher_closed': False,
        },
    )
    return new_status


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
                  AND ss.status IN ('submitted', 'approved', 'locked')) AS sheets_submitted,
               (SELECT COUNT(*) FROM session_sheets ss
                WHERE ss.teacher_ref_id = t.id
                  AND ss.status IN ('approved', 'locked')) AS sheets_approved,
               (SELECT COUNT(*) FROM session_sheets ss
                WHERE ss.teacher_ref_id = t.id
                  AND ss.status = 'in_progress') AS sheets_in_progress,
               (SELECT COUNT(*) FROM session_sheets ss
                WHERE ss.teacher_ref_id = t.id
                  AND ss.status IN ('draft', 'sent')) AS sheets_pending
        FROM teachers t
        LEFT JOIN users u ON u.id = t.user_id
        ORDER BY t.is_active DESC, t.full_name COLLATE UKRAINIAN
    """).fetchall()

    base_url = PUBLIC_APPLY_BASE_URL.rstrip('/')
    teachers = []
    departments_set = set()
    total_sheets = 0
    total_submitted = 0
    total_in_progress = 0
    total_pending = 0
    active_count = 0

    for r in rows:
        d = dict(r)
        total = d['sheets_total'] or 0
        submitted = d['sheets_submitted'] or 0
        in_prog = d['sheets_in_progress'] or 0
        pending = d['sheets_pending'] or 0
        d['sheets_in_progress'] = in_prog
        d['sheets_pending'] = pending
        d['percent'] = round(submitted / total * 100) if total else 0
        d['cabinet_url'] = f"{base_url}/session-grade/teacher/{d['token']}" if d.get('token') else ''
        if d.get('department'):
            departments_set.add(d['department'].strip())
        if d.get('is_active'):
            active_count += 1
        total_sheets += total
        total_submitted += submitted
        total_in_progress += in_prog
        total_pending += pending
        teachers.append(d)

    from routes.helpers import sort_ukrainian
    departments = sort_ukrainian(list(departments_set), lambda x: x)

    summary = {
        'total_teachers': len(teachers),
        'active_teachers': active_count,
        'inactive_teachers': len(teachers) - active_count,
        'total_sheets': total_sheets,
        'submitted_sheets': total_submitted,
        'in_progress_sheets': total_in_progress,
        'pending_sheets': total_pending,
        'overall_percent': round(total_submitted / total_sheets * 100) if total_sheets else 0,
    }

    conn.close()
    return render_template(
        'exam_sessions/teachers_list.html',
        teachers=teachers, departments=departments, summary=summary,
    )


@exam_sessions_bp.route('/admin/teachers/<int:teacher_id>/toggle', methods=['POST'])
@permission_required('manage_exam_sessions')
def teacher_toggle(teacher_id):
    conn = _conn()
    teacher = conn.execute("SELECT id, full_name, is_active FROM teachers WHERE id = ?", (teacher_id,)).fetchone()
    if not teacher:
        conn.close()
        flash("Викладача не знайдено", "error")
        return redirect(url_for('exam_sessions.teachers_list'))
    new_active = 0 if teacher['is_active'] else 1
    conn.execute("UPDATE teachers SET is_active = ? WHERE id = ?", (new_active, teacher_id))
    conn.commit()
    log_action(
        current_username(),
        f"{'активував' if new_active else 'деактивував'} викладача: {teacher['full_name']} (ID {teacher_id})",
    )
    conn.close()
    flash(
        f"Викладача «{teacher['full_name']}» {'активовано' if new_active else 'деактивовано'}.",
        "success",
    )
    return redirect(url_for('exam_sessions.teachers_list'))


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
    departments = [
        r[0] for r in conn.execute(
            "SELECT DISTINCT TRIM(department) FROM teachers WHERE TRIM(COALESCE(department, '')) != '' ORDER BY 1 COLLATE UKRAINIAN"
        ).fetchall()
    ]
    conn.close()
    return render_template(
        'exam_sessions/teacher_form.html',
        teacher=None, users=users, departments=departments, cabinet_url='',
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
                f"Старе посилання більше не працює.",
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
    departments = [
        r[0] for r in conn.execute(
            "SELECT DISTINCT TRIM(department) FROM teachers WHERE TRIM(COALESCE(department, '')) != '' ORDER BY 1 COLLATE UKRAINIAN"
        ).fetchall()
    ]

    # Відомості цього викладача (для блоку знизу)
    stats = teacher_stats(conn, teacher_id)
    if stats and stats.get('sheets'):
        sids = [s['id'] for s in stats['sheets']]
        ph = ','.join('?' * len(sids))
        extra_map = {
            r['id']: dict(r)
            for r in conn.execute(f"""
                SELECT ss.id, ss.session_id, ss.entity_kind, ss.exam_date,
                       ss.linked_entity_id, es.name AS session_name,
                       l.short_name_ua AS license_name
                FROM session_sheets ss
                JOIN exam_sessions es ON es.id = ss.session_id
                LEFT JOIN institution_licenses l ON l.id = ss.license_id
                WHERE ss.id IN ({ph})
            """, sids).fetchall()
        }
        for sh in stats['sheets']:
            ex = extra_map.get(sh['id']) or {}
            sh.update(ex)

    cabinet_url = f"{PUBLIC_APPLY_BASE_URL.rstrip('/')}/session-grade/teacher/{teacher['token']}"
    conn.close()

    return render_template(
        'exam_sessions/teacher_form.html',
        teacher=teacher, users=users, stats=stats,
        departments=departments, cabinet_url=cabinet_url,
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
    years = []
    if group_ids:
        ph = ','.join('?' * len(group_ids))
        years = [
            r[0] for r in conn.execute(
                f"SELECT DISTINCT academic_year FROM exam_sessions WHERE group_id IN ({ph}) ORDER BY academic_year DESC",
                group_ids,
            ).fetchall() if r[0]
        ]
        sql = f"""
            SELECT es.*, g.name AS group_name, g.course, g.study_form,
                   g.start_year, g.specialty,
                   (SELECT COUNT(*) FROM session_sheets ss WHERE ss.session_id = es.id) AS sheets_count,
                   (SELECT COUNT(*) FROM session_sheets ss WHERE ss.session_id = es.id AND ss.status IN ('submitted', 'approved', 'locked')) AS sheets_submitted,
                   (SELECT COUNT(*) FROM session_sheets ss WHERE ss.session_id = es.id AND ss.status IN ('approved', 'locked')) AS sheets_approved,
                   (SELECT COUNT(*) FROM session_sheets ss WHERE ss.session_id = es.id AND ss.status = 'in_progress') AS sheets_in_progress,
                   (SELECT COUNT(*) FROM session_sheets ss WHERE ss.session_id = es.id AND ss.status IN ('draft', 'sent')) AS sheets_pending,
                   (SELECT COUNT(*) FROM session_sheets ss WHERE ss.session_id = es.id AND ss.linked_entity_id IS NOT NULL) AS sheets_linked
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

    for r in rows:
        tot = r['sheets_count'] or 0
        sub = r['sheets_submitted'] or 0
        r['progress_percent'] = round(sub / tot * 100) if tot else 0

    total_sheets = sum(r['sheets_count'] or 0 for r in rows)
    done_sheets = sum(r['sheets_submitted'] or 0 for r in rows)
    pending_sheets = sum(r['sheets_pending'] or 0 for r in rows)
    in_progress_sheets = sum(r['sheets_in_progress'] or 0 for r in rows)
    summary = {
        'total_sessions': len(rows),
        'open_sessions': sum(1 for r in rows if r.get('status') == 'open'),
        'closed_sessions': sum(1 for r in rows if r.get('status') == 'closed'),
        'total_sheets': total_sheets,
        'done_sheets': done_sheets,
        'in_progress_sheets': in_progress_sheets,
        'pending_sheets': pending_sheets,
        'overall_percent': round(done_sheets / total_sheets * 100) if total_sheets else 0,
    }
    can_transfer = bool(session.get('is_admin') or 'transfer_exam_grades' in (session.get('permissions') or []))
    conn.close()
    return render_template(
        'exam_sessions/sessions_list.html',
        rows=rows, groups=groups, years=years,
        filter_group=filter_group, filter_year=filter_year,
        filter_status=filter_status, summary=summary, can_transfer=can_transfer,
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
        elif not academic_year or semester not in (1, 2, 3):
            flash('Вкажіть навчальний рік і семестр (1, 2 або 3)', 'error')
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
    now = datetime.now()
    start_y = now.year if now.month >= 8 else now.year - 1
    default_year = f"{start_y}–{start_y + 1}"
    conn.close()
    return render_template('exam_sessions/session_form.html', groups=groups, row=None, default_year=default_year)


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

        if action == 'quick_approve':
            target_sheet_id = request.form.get('sheet_id', type=int)
            sh = conn.execute(
                "SELECT id, status, discipline_name FROM session_sheets WHERE id = ? AND session_id = ?",
                (target_sheet_id, session_id),
            ).fetchone()
            if sh and sh['status'] in ('submitted', 'in_progress', 'sent'):
                _, retake_info = _approve_sheet_and_spawn_retake(
                    conn, target_sheet_id, session_id, current_username()
                )
                conn.commit()
                log_action(current_username(), f"затвердив відомість #{target_sheet_id} у сесії #{session_id}")
                msg = f"Відомість «{sh['discipline_name']}» затверджено."
                if retake_info:
                    log_action(
                        current_username(),
                        f"автоматично створив відомість перездачі #{retake_info['sheet_id']} "
                        f"з відомості #{target_sheet_id} ({retake_info['students_count']} студ.)",
                    )
                    msg += (
                        f" Автоматично створено відомість перездачі #{retake_info['sheet_id']} "
                        f"для {retake_info['students_count']} студ. (нездані / неявки)."
                    )
                flash(msg, 'success')
            conn.close()
            return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

        if action == 'reopen_sheet':
            can_reopen = bool(session.get('is_admin') or 'reopen_exam_sheets' in (session.get('permissions') or []))
            if not can_reopen:
                conn.close()
                flash('Немає права повторно відкривати відомості', 'error')
                return redirect(url_for('exam_sessions.session_detail', session_id=session_id))
            target_sheet_id = request.form.get('sheet_id', type=int)
            sh = conn.execute(
                "SELECT * FROM session_sheets WHERE id = ? AND session_id = ?",
                (target_sheet_id, session_id),
            ).fetchone()
            if sh:
                sh_d = dict(sh)
                if sh_d['status'] == 'locked':
                    flash('Відомість остаточно заблокована — повторне відкриття неможливе.', 'error')
                elif sh_d.get('teacher_closed') or sh_d['status'] in ('submitted', 'approved'):
                    _reopen_sheet_for_teacher(conn, sh_d, current_username())
                    conn.commit()
                    log_action(current_username(), f"повторно відкрив відомість #{target_sheet_id} у сесії #{session_id}")
                    flash(f"Відомість «{sh_d['discipline_name']}» повторно відкрито для викладача.", 'success')
            conn.close()
            return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

        if action == 'bulk_approve':
            raw_ids = request.form.getlist('sheet_ids')
            sids = [int(x) for x in raw_ids if str(x).isdigit()]
            approved_count = 0
            retakes_created = 0
            retake_students_total = 0
            if sids:
                ph = ','.join('?' * len(sids))
                candidates = conn.execute(
                    f"SELECT id, status FROM session_sheets WHERE session_id = ? AND id IN ({ph}) AND status IN ('submitted', 'in_progress', 'sent')",
                    [session_id] + sids,
                ).fetchall()
                for c in candidates:
                    _, retake_info = _approve_sheet_and_spawn_retake(
                        conn, c['id'], session_id, current_username()
                    )
                    approved_count += 1
                    if retake_info:
                        retakes_created += 1
                        retake_students_total += retake_info['students_count']
                if approved_count:
                    conn.commit()
                    log_action(
                        current_username(),
                        f"масово затвердив відомості ({approved_count}) у сесії #{session_id}"
                        + (f"; створено перездач: {retakes_created}" if retakes_created else ""),
                    )
            conn.close()
            if approved_count:
                msg = f'Затверджено відомостей: {approved_count}.'
                if retakes_created:
                    msg += f' Автоматично створено відомостей перездачі: {retakes_created} (студентів: {retake_students_total}).'
                flash(msg, 'success')
            else:
                flash('Не знайдено відомостей для затвердження серед обраних.', 'info')
            return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

        if action == 'bulk_refresh_students':
            raw_ids = request.form.getlist('sheet_ids')
            sids = [int(x) for x in raw_ids if str(x).isdigit()]
            refreshed = 0
            if sids:
                ph = ','.join('?' * len(sids))
                candidates = conn.execute(
                    f"""SELECT id, license_id FROM session_sheets
                        WHERE session_id = ? AND id IN ({ph})
                          AND status IN ('draft', 'sent', 'in_progress')
                          AND COALESCE(sheet_kind, 'main') != 'retake'
                          AND COALESCE(teacher_closed, 0) = 0""",
                    [session_id] + sids,
                ).fetchall()
                for c in candidates:
                    n = populate_sheet_students(conn, c['id'], es['group_id'], c['license_id'])
                    log_history(
                        conn,
                        entity_type='sheet', entity_id=c['id'],
                        session_id=session_id, sheet_id=c['id'],
                        action='refresh', details={'students_count': n},
                    )
                    refreshed += 1
                if refreshed:
                    conn.commit()
            conn.close()
            if refreshed:
                flash(f'Оновлено списки студентів у відомостях: {refreshed}.', 'success')
            else:
                flash('Список студентів можна оновити лише у відкритих основних відомостях (чернетки, надіслані або в процесі).', 'info')
            return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

    sheets = [dict(r) for r in conn.execute("""
        SELECT ss.*,
               t.short_name AS teacher_short,
               t.full_name AS teacher_full,
               t.token AS teacher_token,
               l.short_name_ua AS license_name,
               (SELECT COUNT(*) FROM session_sheet_students x WHERE x.sheet_id = ss.id) AS students_count,
               (SELECT COUNT(*) FROM session_sheet_students x
                WHERE x.sheet_id = ss.id AND x.points IS NOT NULL) AS graded_count,
               (SELECT COUNT(*) FROM session_sheet_students x
                WHERE x.sheet_id = ss.id AND (x.points IS NULL OR x.points < 60)) AS problem_count,
               (SELECT c.id FROM session_sheets c
                WHERE c.parent_sheet_id = ss.id ORDER BY c.id DESC LIMIT 1) AS child_retake_id
        FROM session_sheets ss
        LEFT JOIN teachers t ON t.id = ss.teacher_ref_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        WHERE ss.session_id = ?
        ORDER BY ss.discipline_name,
                 CASE WHEN COALESCE(ss.sheet_kind, 'main') = 'retake' THEN 1 ELSE 0 END,
                 ss.license_id, ss.id
    """, (session_id,)).fetchall()]

    base_url = PUBLIC_APPLY_BASE_URL.rstrip('/')
    session_licenses_map = {}
    session_teachers_map = {}
    for s in sheets:
        s['teacher_cabinet_url'] = f"{base_url}/session-grade/teacher/{s['teacher_token']}" if s.get('teacher_token') else ''
        if s.get('license_id'):
            session_licenses_map[s['license_id']] = s.get('license_name') or f"Ліцензія #{s['license_id']}"
        if s.get('teacher_ref_id'):
            session_teachers_map[s['teacher_ref_id']] = s.get('teacher_short') or s.get('teacher_full') or ''

    total_sheets = len(sheets)
    submitted_sheets = sum(1 for s in sheets if s['status'] in ('submitted', 'approved', 'locked'))
    approved_sheets = sum(1 for s in sheets if s['status'] in ('approved', 'locked'))
    in_progress_sheets = sum(1 for s in sheets if s['status'] == 'in_progress')
    pending_sheets = sum(1 for s in sheets if s['status'] in ('draft', 'sent'))
    linked_sheets = sum(1 for s in sheets if s.get('linked_entity_id'))
    retake_sheets_count = sum(1 for s in sheets if (s.get('sheet_kind') or 'main') == 'retake')
    total_students_slots = sum(s['students_count'] or 0 for s in sheets)
    total_graded_slots = sum(s['graded_count'] or 0 for s in sheets)

    session_stats = {
        'total': total_sheets,
        'submitted': submitted_sheets,
        'approved': approved_sheets,
        'in_progress': in_progress_sheets,
        'pending': pending_sheets,
        'linked': linked_sheets,
        'retakes': retake_sheets_count,
        'percent': round(submitted_sheets / total_sheets * 100) if total_sheets else 0,
        'students_total': total_students_slots,
        'students_graded': total_graded_slots,
        'students_percent': round(total_graded_slots / total_students_slots * 100) if total_students_slots else 0,
    }

    licenses = licenses_present_in_group(conn, es['group_id'])
    can_transfer = bool(session.get('is_admin') or 'transfer_exam_grades' in (session.get('permissions') or []))
    can_reopen = bool(session.get('is_admin') or 'reopen_exam_sheets' in (session.get('permissions') or []))
    teachers = conn.execute(
        "SELECT id, full_name, short_name FROM teachers "
        "WHERE is_active = 1 ORDER BY full_name COLLATE UKRAINIAN"
    ).fetchall()
    conn.close()
    return render_template(
        'exam_sessions/session_detail.html',
        es=es, sheets=sheets, licenses=licenses, can_transfer=can_transfer,
        can_reopen=can_reopen,
        teachers=teachers, session_stats=session_stats,
        session_licenses=list(session_licenses_map.items()),
        session_teachers=list(session_teachers_map.items()),
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
        lic_raw = (request.form.get('license_id') or '').strip()
        create_for_all = (lic_raw == 'all')
        license_id = int(lic_raw) if lic_raw and lic_raw.isdigit() else None

        # Snapshot teacher_name из картотеки (скорочене ім'я)
        teacher_name = None
        if teacher_ref_id:
            t = conn.execute(
                "SELECT full_name, short_name FROM teachers WHERE id = ?",
                (teacher_ref_id,),
            ).fetchone()
            if t:
                teacher_name = (t['short_name'] or t['full_name'] or '').strip()

        if entity_kind not in ('subject', 'practice', 'coursework', 'attestation'):
            entity_kind = 'subject'
        if control_form not in ('Залік', 'Екзамен'):
            control_form = 'Залік'

        if not discipline_name:
            flash('Вкажіть назву дисципліни', 'error')
        elif not teacher_ref_id:
            flash('Оберіть викладача з картотеки', 'error')
        elif not create_for_all and not license_id:
            flash('Оберіть ліцензію — вона обовʼязкова', 'error')
        elif create_for_all and not licenses:
            flash('У цій групі немає студентів із призначеною ліцензією', 'error')
        else:
            target_license_ids = [l['id'] for l in licenses] if create_for_all else [license_id]
            created_sheet_ids = []
            total_students = 0
            for lid in target_license_ids:
                cur = conn.execute("""
                    INSERT INTO session_sheets (
                        session_id, discipline_name, work_topic, entity_kind, control_form,
                        hours_credits, teacher_name, teacher_ref_id, exam_date, statement_number,
                        license_id, status, created_by
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'draft', ?)
                """, (
                    session_id, discipline_name, work_topic, entity_kind, control_form,
                    hours_credits, teacher_name, teacher_ref_id, exam_date, statement_number,
                    lid, current_username(),
                ))
                sheet_id = cur.lastrowid
                n = populate_sheet_students(conn, sheet_id, es['group_id'], lid)
                total_students += n
                created_sheet_ids.append(sheet_id)
                log_history(
                    conn,
                    entity_type='sheet', entity_id=sheet_id,
                    session_id=session_id, sheet_id=sheet_id,
                    action='create',
                    details={
                        'discipline_name': discipline_name,
                        'entity_kind': entity_kind,
                        'control_form': control_form,
                        'license_id': lid,
                        'teacher_ref_id': teacher_ref_id,
                        'students_count': n,
                        'created_for_all_licenses': create_for_all,
                    },
                )
            conn.commit()
            if len(created_sheet_ids) == 1:
                sid = created_sheet_ids[0]
                log_action(
                    current_username(),
                    f"створив відомість #{sid} «{discipline_name}» у сесії #{session_id} "
                    f"({total_students} студ., викладач: {teacher_name})",
                )
                conn.close()
                flash(
                    f'Відомість створено. Студентів у списку: {total_students}. '
                    f'Відомість автоматично з\'явилась у викладача в кабінеті.',
                    'success',
                )
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sid))
            else:
                log_action(
                    current_username(),
                    f"створив {len(created_sheet_ids)} відомостей для всіх ліцензій «{discipline_name}» "
                    f"у сесії #{session_id} ({total_students} студ., викладач: {teacher_name})",
                )
                conn.close()
                flash(
                    f'Створено відомостей для всіх ліцензій: {len(created_sheet_ids)} '
                    f'(студентів загалом: {total_students}). '
                    f'Відомості автоматично з\'явились у викладача в кабінеті.',
                    'success',
                )
                return redirect(url_for('exam_sessions.session_detail', session_id=session_id))

    teachers = conn.execute(
        "SELECT id, full_name, short_name, department FROM teachers "
        "WHERE is_active = 1 ORDER BY full_name COLLATE UKRAINIAN"
    ).fetchall()

    # Підказки з навчального плану групи (щоб не набирати вручну і точно збігалось при переносі)
    existing_disciplines = {
        (
            ' '.join((r['discipline_name'] or '').split()).lower(),
            (r['entity_kind'] or 'subject').strip().lower(),
        )
        for r in conn.execute(
            "SELECT discipline_name, entity_kind FROM session_sheets WHERE session_id = ?",
            (session_id,),
        ).fetchall()
    }
    plan_items = []
    for tbl, kind_key in (
        ('subjects', 'subject'),
        ('practices', 'practice'),
        ('courseworks', 'coursework'),
        ('attestations', 'attestation'),
    ):
        try:
            q_rows = conn.execute(
                f"SELECT code, name, credits, type FROM {tbl} WHERE group_id = ? ORDER BY position, name",
                (es['group_id'],),
            ).fetchall()
            for r in q_rows:
                nm = (r['name'] or '').strip()
                if not nm:
                    continue
                cr = r['credits']
                hc = ''
                if cr is not None and str(cr).strip() != '':
                    try:
                        cr_f = float(cr)
                        hrs = int(round(cr_f * 30))
                        cr_txt = f"{cr_f:.1f}" if not cr_f.is_integer() else f"{int(cr_f)}.0"
                        hc = f"{hrs}/{cr_txt}"
                    except (TypeError, ValueError):
                        pass
                plan_items.append({
                    'kind': kind_key,
                    'code': r['code'] or '',
                    'name': nm,
                    'control_form': r['type'] if r['type'] in ('Залік', 'Екзамен') else 'Залік',
                    'hours_credits': hc,
                    'exists': (' '.join(nm.split()).lower(), kind_key) in existing_disciplines,
                })
        except Exception:
            pass

    conn.close()
    return render_template(
        'exam_sessions/sheet_form.html',
        es=es, licenses=licenses, sheet=None, teachers=teachers,
        plan_items=plan_items,
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

        if action == 'copy_sheet':
            teacher_name = sheet.get('teacher_name')
            if sheet.get('teacher_ref_id'):
                t = conn.execute(
                    "SELECT full_name, short_name FROM teachers WHERE id = ?",
                    (sheet['teacher_ref_id'],),
                ).fetchone()
                if t:
                    teacher_name = (t['short_name'] or t['full_name'] or '').strip()

            target_license_id = request.form.get('license_id', type=int) or sheet.get('license_id')

            cur = conn.execute("""
                INSERT INTO session_sheets (
                    session_id, discipline_name, work_topic, entity_kind, control_form,
                    hours_credits, teacher_name, teacher_ref_id, exam_date, statement_number,
                    license_id, status, created_by
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'draft', ?)
            """, (
                sheet['session_id'],
                sheet['discipline_name'],
                sheet.get('work_topic'),
                sheet.get('entity_kind') or 'subject',
                sheet.get('control_form') or 'Залік',
                sheet.get('hours_credits'),
                teacher_name,
                sheet.get('teacher_ref_id'),
                sheet.get('exam_date'),
                sheet.get('statement_number'),
                target_license_id,
                current_username(),
            ))
            new_sheet_id = cur.lastrowid
            n_copied = populate_sheet_students(conn, new_sheet_id, sheet['group_id'], target_license_id)

            log_history(
                conn,
                entity_type='sheet', entity_id=new_sheet_id,
                session_id=sheet['session_id'], sheet_id=new_sheet_id,
                action='create', field_name='copy_from', new_value=str(sheet_id),
                details={
                    'source_sheet_id': sheet_id,
                    'discipline_name': sheet['discipline_name'],
                    'license_id': target_license_id,
                    'students_count': n_copied,
                },
            )
            conn.commit()
            log_action(
                current_username(),
                f"скопіював відомість #{sheet_id} → #{new_sheet_id} "
                f"«{sheet['discipline_name']}» у сесії #{sheet['session_id']} "
                f"(студентів: {n_copied})",
            )
            conn.close()
            if target_license_id != sheet.get('license_id'):
                flash(
                    f'Відомість скопійовано (#{new_sheet_id}) з новою ліцензією '
                    f'(студентів: {n_copied}).',
                    'success',
                )
            else:
                flash(
                    f'Відомість скопійовано (#{new_sheet_id}). Оберіть потрібну ліцензію '
                    f'у «Метадані відомості» та натисніть «Зберегти метадані».',
                    'success',
                )
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=new_sheet_id))

        if action == 'save_meta' and sheet['status'] not in ('locked',):
            teacher_ref_id = request.form.get('teacher_ref_id', type=int) or sheet['teacher_ref_id']
            teacher_name = sheet['teacher_name']
            if teacher_ref_id != sheet['teacher_ref_id']:
                t = conn.execute(
                    "SELECT full_name, short_name FROM teachers WHERE id = ?",
                    (teacher_ref_id,),
                ).fetchone()
                if t:
                    teacher_name = (t['short_name'] or t['full_name'] or '').strip()

            new_license_id = request.form.get('license_id', type=int)
            if new_license_id is None:
                new_license_id = sheet.get('license_id')

            if new_license_id != sheet.get('license_id') and sheet.get('linked_entity_id'):
                flash('Ліцензію не змінено: відомість уже перенесена в навчальний план.', 'warning')
                new_license_id = sheet.get('license_id')

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
                'license_id': new_license_id,
            }

            if fields['discipline_name']:
                changes = {
                    f: (sheet.get(f), v)
                    for f, v in fields.items()
                    if sheet.get(f) != v
                }
                license_changed = 'license_id' in changes

                conn.execute("""
                    UPDATE session_sheets SET
                        discipline_name=?, work_topic=?, entity_kind=?, control_form=?,
                        hours_credits=?, teacher_name=?, teacher_ref_id=?,
                        exam_date=?, statement_number=?, license_id=?
                    WHERE id=?
                """, (
                    fields['discipline_name'], fields['work_topic'],
                    fields['entity_kind'], fields['control_form'],
                    fields['hours_credits'], fields['teacher_name'], fields['teacher_ref_id'],
                    fields['exam_date'], fields['statement_number'], fields['license_id'],
                    sheet_id,
                ))

                for field, (old, new) in changes.items():
                    log_history(
                        conn,
                        entity_type='sheet', entity_id=sheet_id,
                        session_id=sheet['session_id'], sheet_id=sheet_id,
                        action='update', field_name=field,
                        old_value=old, new_value=new,
                    )

                if license_changed:
                    n_students = populate_sheet_students(
                        conn, sheet_id, sheet['group_id'], fields['license_id']
                    )
                    conn.commit()
                    log_action(
                        current_username(),
                        f"змінив ліцензію відомості #{sheet_id} "
                        f"({sheet.get('license_id')} → {fields['license_id']}, студентів: {n_students})",
                    )
                    flash(
                        f'Метадані збережено. Ліцензію змінено — список студентів оновлено ({n_students}).',
                        'success',
                    )
                else:
                    conn.commit()
                    flash('Метадані збережено', 'success')
            else:
                flash("Назва дисципліни обов'язкова", 'error')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'refresh_students' and sheet['status'] in ('draft', 'sent', 'in_progress'):
            if (sheet.get('sheet_kind') or 'main').strip().lower() == 'retake':
                flash(
                    'Список студентів відомості перездачі сформовано з незданих/неявок і не оновлюється загальним списком групи.',
                    'warning',
                )
                conn.close()
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))
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
            if sheet.get('teacher_closed') and sheet['status'] not in ('approved', 'locked'):
                conn.execute("UPDATE session_sheets SET status='submitted' WHERE id=?", (sheet_id,))
            conn.commit()
            flash('Оцінки збережено', 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'approve' and sheet['status'] in ('submitted', 'in_progress', 'sent', 'approved'):
            _, retake_info = _approve_sheet_and_spawn_retake(
                conn, sheet_id, sheet['session_id'], current_username()
            )
            conn.commit()
            log_action(current_username(), f"затвердив відомість #{sheet_id}")
            msg = 'Відомість затверджено.'
            if retake_info:
                log_action(
                    current_username(),
                    f"автоматично створив відомість перездачі #{retake_info['sheet_id']} "
                    f"з відомості #{sheet_id} ({retake_info['students_count']} студ.)",
                )
                msg += (
                    f" Автоматично створено відомість перездачі #{retake_info['sheet_id']} "
                    f"для {retake_info['students_count']} студ. (нездані / неявки)."
                )
            flash(msg, 'success')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'create_retake' and sheet['status'] in ('approved', 'locked'):
            retake_info = create_retake_sheet_if_needed(
                conn, sheet_id, created_by=current_username(), allow_from_retake=True
            )
            if retake_info:
                log_history(
                    conn,
                    entity_type='sheet', entity_id=retake_info['sheet_id'],
                    session_id=sheet['session_id'], sheet_id=retake_info['sheet_id'],
                    action='create_retake',
                    details={
                        'discipline_name': retake_info['discipline_name'],
                        'parent_sheet_id': sheet_id,
                        'retake_students_count': retake_info['students_count'],
                        'license_id': retake_info.get('license_id'),
                        'teacher_ref_id': retake_info.get('teacher_ref_id'),
                    },
                )
                conn.commit()
                log_action(
                    current_username(),
                    f"створив відомість перездачі #{retake_info['sheet_id']} з відомості #{sheet_id} "
                    f"({retake_info['students_count']} студ.)",
                )
                conn.close()
                flash(
                    f"Створено відомість перездачі #{retake_info['sheet_id']} "
                    f"для {retake_info['students_count']} студ.",
                    'success',
                )
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=retake_info['sheet_id']))
            flash('Немає студентів із незадовільними оцінками або неявкою для перездачі.', 'info')
            conn.close()
            return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))

        if action == 'reopen_sheet':
            can_reopen = bool(session.get('is_admin') or 'reopen_exam_sheets' in (session.get('permissions') or []))
            if not can_reopen:
                conn.close()
                flash('Немає права повторно відкривати відомості', 'error')
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))
            if sheet['status'] == 'locked':
                conn.close()
                flash('Відомість остаточно заблокована — повторне відкриття неможливе.', 'error')
                return redirect(url_for('exam_sessions.sheet_detail', sheet_id=sheet_id))
            if sheet.get('teacher_closed') or sheet['status'] in ('submitted', 'approved'):
                _reopen_sheet_for_teacher(conn, sheet, current_username())
                conn.commit()
                log_action(current_username(), f"повторно відкрив відомість #{sheet_id} для викладача")
                flash('Відомість повторно відкрито — викладач знову може редагувати оцінки.', 'success')
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

                conn.execute("UPDATE session_sheets SET parent_sheet_id = NULL WHERE parent_sheet_id = ?", (sheet_id,))
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
                conn.execute("UPDATE session_sheets SET parent_sheet_id = NULL WHERE parent_sheet_id = ?", (sheet_id,))
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
    problem_students_count = sum(1 for st in students if is_failing_or_absent_student(st))
    parent_sheet = None
    if sheet.get('parent_sheet_id'):
        p_row = conn.execute(
            "SELECT id, discipline_name, status, exam_date FROM session_sheets WHERE id = ?",
            (sheet['parent_sheet_id'],),
        ).fetchone()
        if p_row:
            parent_sheet = dict(p_row)
    child_retakes = [dict(r) for r in conn.execute("""
        SELECT id, status, exam_date, teacher_closed, created_at,
               (SELECT COUNT(*) FROM session_sheet_students x WHERE x.sheet_id = session_sheets.id) AS students_count,
               (SELECT COUNT(*) FROM session_sheet_students x WHERE x.sheet_id = session_sheets.id AND x.points IS NOT NULL) AS graded_count
        FROM session_sheets
        WHERE parent_sheet_id = ?
        ORDER BY id DESC
    """, (sheet_id,)).fetchall()]

    base_url = PUBLIC_APPLY_BASE_URL.rstrip('/')
    teacher_link = None
    if sheet.get('teacher_token'):
        teacher_link = f"{base_url}/session-grade/teacher/{sheet['teacher_token']}"
    sheet_direct_link = None
    if sheet.get('token'):
        sheet_direct_link = f"{base_url}/session-grade/{sheet['token']}"
    can_transfer = bool(session.get('is_admin') or 'transfer_exam_grades' in (session.get('permissions') or []))
    can_reopen = bool(session.get('is_admin') or 'reopen_exam_sheets' in (session.get('permissions') or []))
    teachers = conn.execute(
        "SELECT id, full_name, short_name FROM teachers "
        "WHERE is_active = 1 OR id = ? ORDER BY full_name COLLATE UKRAINIAN",
        (sheet['teacher_ref_id'],),
    ).fetchall()
    group_licenses = licenses_present_in_group(conn, sheet['group_id'])
    group_license_ids = {l['id'] for l in group_licenses}
    all_active_licenses = [
        dict(r) for r in conn.execute(
            "SELECT id, COALESCE(short_name_ua, name_ua, '—') AS name "
            "FROM institution_licenses WHERE is_active = 1 OR id = ? "
            "ORDER BY name COLLATE UKRAINIAN",
            (sheet.get('license_id'),),
        ).fetchall()
    ]
    other_licenses = [l for l in all_active_licenses if l['id'] not in group_license_ids]
    conn.close()

    sheet['exam_date_iso'] = _to_iso_date(sheet.get('exam_date'))

    return render_template(
        'exam_sessions/sheet_detail.html',
        sheet=sheet, students=students, teacher_link=teacher_link,
        sheet_direct_link=sheet_direct_link,
        can_transfer=can_transfer, can_reopen=can_reopen, teachers=teachers,
        group_licenses=group_licenses, other_licenses=other_licenses,
        parent_sheet=parent_sheet, child_retakes=child_retakes,
        problem_students_count=problem_students_count,
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
            t.token,
            COUNT(ss.id) AS sheets_total,
            SUM(CASE WHEN ss.status IN ('submitted', 'approved', 'locked') THEN 1 ELSE 0 END) AS sheets_submitted,
            SUM(CASE WHEN ss.status IN ('approved', 'locked') THEN 1 ELSE 0 END) AS sheets_approved,
            SUM(CASE WHEN ss.status = 'in_progress' THEN 1 ELSE 0 END) AS sheets_in_progress,
            SUM(CASE WHEN ss.status IN ('draft', 'sent') THEN 1 ELSE 0 END) AS sheets_pending,
            COALESCE((
                SELECT COUNT(*) FROM session_sheet_students x
                JOIN session_sheets s2 ON s2.id = x.sheet_id
                WHERE s2.teacher_ref_id = t.id
            ), 0) AS students_total,
            COALESCE((
                SELECT COUNT(*) FROM session_sheet_students x
                JOIN session_sheets s2 ON s2.id = x.sheet_id
                WHERE s2.teacher_ref_id = t.id AND x.points IS NOT NULL
            ), 0) AS students_graded
        FROM teachers t
        LEFT JOIN session_sheets ss ON ss.teacher_ref_id = t.id
        GROUP BY t.id
        ORDER BY t.is_active DESC, sheets_pending DESC, t.full_name COLLATE UKRAINIAN
    """).fetchall()

    base_url = PUBLIC_APPLY_BASE_URL.rstrip('/')
    stats = []
    departments_set = set()
    tot_sheets = 0
    tot_submitted = 0
    tot_in_prog = 0
    tot_pending = 0
    with_pending_teachers = 0

    for r in rows:
        total = r['sheets_total'] or 0
        submitted = r['sheets_submitted'] or 0
        in_prog = r['sheets_in_progress'] or 0
        pending = r['sheets_pending'] or 0
        st_tot = r['students_total'] or 0
        st_grd = r['students_graded'] or 0
        dept = (r['department'] or '').strip()
        if dept:
            departments_set.add(dept)
        if pending > 0 or in_prog > 0:
            with_pending_teachers += 1
        tot_sheets += total
        tot_submitted += submitted
        tot_in_prog += in_prog
        tot_pending += pending
        stats.append({
            'teacher_id': r['teacher_id'],
            'full_name': r['full_name'],
            'short_name': r['short_name'],
            'department': r['department'],
            'is_active': r['is_active'],
            'cabinet_url': f"{base_url}/session-grade/teacher/{r['token']}" if r['token'] else '',
            'total': total,
            'submitted': submitted,
            'approved': r['sheets_approved'] or 0,
            'in_progress': in_prog,
            'pending': pending,
            'students_total': st_tot,
            'students_graded': st_grd,
            'percent': round(submitted / total * 100) if total else 0,
        })

    from routes.helpers import sort_ukrainian
    departments = sort_ukrainian(list(departments_set), lambda x: x)
    summary = {
        'teachers_count': len(stats),
        'with_pending_teachers': with_pending_teachers,
        'total_sheets': tot_sheets,
        'submitted_sheets': tot_submitted,
        'in_progress_sheets': tot_in_prog,
        'pending_sheets': tot_pending,
        'overall_percent': round(tot_submitted / tot_sheets * 100) if tot_sheets else 0,
    }

    conn.close()
    return render_template(
        'exam_sessions/teacher_stats.html',
        stats=stats, departments=departments, summary=summary,
    )


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
        raw_sheets = [dict(r) for r in conn.execute("""
            SELECT ss.id, ss.discipline_name, ss.control_form, ss.hours_credits,
                   ss.entity_kind, ss.license_id, ss.sheet_kind, ss.parent_sheet_id,
                   ss.exam_date, l.short_name_ua AS license_name
            FROM session_sheets ss
            LEFT JOIN institution_licenses l ON l.id = ss.license_id
            WHERE ss.session_id=?
            ORDER BY ss.discipline_name, ss.license_id
        """, (es['id'],)).fetchall()]

        raw_by_students = {}
        for sh in raw_sheets:
            rows = [dict(r) for r in conn.execute("""
                SELECT student_id, full_name, points, national_grade, ects
                FROM session_sheet_students
                WHERE sheet_id=?
                ORDER BY position, full_name
            """, (sh['id'],)).fetchall()]
            raw_by_students[sh['id']] = rows

        sheets, merged_by_students = _merge_summary_sheets_by_discipline(raw_sheets, raw_by_students)

        grades = {}
        for sh in sheets:
            for r in merged_by_students.get(sh['id'], []):
                sid = r['student_id']
                if sid not in student_map:
                    student_map[sid] = r['full_name']
                grades.setdefault(sid, {})[sh['id']] = {
                    'national': r['national_grade'],
                    'points': r['points'],
                    'ects': r['ects'],
                    'is_retake': bool(r.get('is_retake')),
                }

        sessions_blocks.append({
            'academic_year': es.get('academic_year'),
            'semester': es.get('semester'),
            'sheets': sheets,
            'grades': grades,
        })

    conn.close()

    from routes.helpers import sort_ukrainian
    students = [{'id': sid, 'full_name': name} for sid, name in student_map.items()]
    students = sort_ukrainian(students, lambda x: x['full_name'] or '')

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
               l.short_name_ua AS license_name,
               t.short_name AS teacher_short, t.full_name AS teacher_full
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        LEFT JOIN teachers t ON t.id = ss.teacher_ref_id
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
               l.short_name_ua AS license_name,
               t.short_name AS teacher_short, t.full_name AS teacher_full
        FROM session_sheets ss
        JOIN exam_sessions es ON es.id = ss.session_id
        JOIN groups g ON g.id = es.group_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        LEFT JOIN teachers t ON t.id = ss.teacher_ref_id
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
    sheets, by_students = _merge_summary_sheets_by_discipline(sheets, by_students)
    for sh_id, st_rows in by_students.items():
        for st in st_rows:
            if st.get('is_retake'):
                nat = (st.get('national_grade') or '').strip()
                if nat and 'перездача' not in nat.lower():
                    st['national_grade'] = f"{nat} (перездача)"
                ects = (st.get('ects') or '').strip()
                if ects and not ects.endswith('*'):
                    st['ects'] = f"{ects}*"
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
    sheets, by_students = _merge_summary_sheets_by_discipline(sheets, by_students)

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


# ═══════════════════════════════════════════════════════════════════
# КОПІЮВАННЯ СЕСІЇ ДЛЯ ІНШОЇ ГРУПИ
# ═══════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/admin/exam-sessions/<int:session_id>/copy', methods=['GET', 'POST'])
@permission_required('manage_exam_sessions')
def session_copy(session_id):
    """
    Копія сесії + відомостей для іншої групи.
    Можна копіювати список студентів і (за бажанням) оцінки.
    Після копіювання — редірект на нову сесію для редагування.
    """
    conn = _conn()
    es = conn.execute("""
        SELECT es.*, g.name AS group_name, g.course, g.study_form
        FROM exam_sessions es
        JOIN groups g ON g.id = es.group_id
        WHERE es.id = ?
    """, (session_id,)).fetchone()
    if not es or not user_can_access_group(conn, session, es['group_id']):
        conn.close()
        flash('Сесію не знайдено', 'error')
        return redirect(url_for('exam_sessions.sessions_list'))
    es = dict(es)

    if request.method == 'POST':
        new_group_id = request.form.get('group_id', type=int)
        copy_students = request.form.get('copy_students') == '1'
        copy_grades = request.form.get('copy_grades') == '1'
        new_name = (request.form.get('name') or '').strip() or f"{es.get('name') or 'Сесія'} (копія)"

        if not new_group_id:
            flash('Оберіть групу', 'error')
            conn.close()
            return redirect(url_for('exam_sessions.session_copy', session_id=session_id))

        if not user_can_access_group(conn, session, new_group_id):
            flash('Немає доступу до цієї групи', 'error')
            conn.close()
            return redirect(url_for('exam_sessions.session_copy', session_id=session_id))

        cur = conn.execute("""
            INSERT INTO exam_sessions (name, group_id, academic_year, semester, status, created_by)
            VALUES (?, ?, ?, ?, 'draft', ?)
        """, (new_name, new_group_id, es.get('academic_year'), es.get('semester'), current_username()))
        new_session_id = cur.lastrowid

        sheets = [dict(r) for r in conn.execute(
            "SELECT * FROM session_sheets WHERE session_id = ? AND COALESCE(sheet_kind, 'main') != 'retake'",
            (session_id,),
        ).fetchall()]

        from routes.gen_docx import _format_teacher_short

        for sh in sheets:
            license_id = sh.get('license_id')
            if not license_id:
                lic = licenses_present_in_group(conn, new_group_id)
                license_id = lic[0]['id'] if lic else None
            if not license_id:
                # без ліцензії не створюємо відомість
                continue

            teacher_name = sh.get('teacher_name')
            if sh.get('teacher_ref_id'):
                t = conn.execute(
                    "SELECT full_name, short_name FROM teachers WHERE id = ?",
                    (sh['teacher_ref_id'],),
                ).fetchone()
                if t:
                    teacher_name = (t['short_name'] or t['full_name'] or '').strip()

            cur = conn.execute("""
                INSERT INTO session_sheets (
                    session_id, discipline_name, work_topic, entity_kind,
                    control_form, hours_credits, teacher_name, teacher_ref_id,
                    exam_date, statement_number, license_id, status, created_by
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'draft', ?)
            """, (
                new_session_id, sh.get('discipline_name'), sh.get('work_topic'),
                sh.get('entity_kind') or 'subject', sh.get('control_form') or 'Залік',
                sh.get('hours_credits'), teacher_name, sh.get('teacher_ref_id'),
                sh.get('exam_date'), sh.get('statement_number'), license_id,
                current_username(),
            ))
            new_sheet_id = cur.lastrowid

            if copy_students:
                n = populate_sheet_students(conn, new_sheet_id, new_group_id, license_id)
                if copy_grades and n:
                    old_students = conn.execute(
                        "SELECT * FROM session_sheet_students WHERE sheet_id = ?",
                        (sh['id'],)
                    ).fetchall()
                    old_by_name = {r['full_name']: dict(r) for r in old_students}
                    new_students = conn.execute(
                        "SELECT id, full_name FROM session_sheet_students WHERE sheet_id = ?",
                        (new_sheet_id,)
                    ).fetchall()
                    for ns in new_students:
                        src = old_by_name.get(ns['full_name'])
                        if not src:
                            continue
                        conn.execute("""
                            UPDATE session_sheet_students SET
                                points_mid=?, points_current=?, points_exam=?,
                                points=?, national_grade=?, ects=?, work_topic=?
                            WHERE id=?
                        """, (
                            src.get('points_mid'), src.get('points_current'),
                            src.get('points_exam'), src.get('points'),
                            src.get('national_grade'), src.get('ects'),
                            src.get('work_topic'), ns['id'],
                        ))

        log_history(
            conn,
            entity_type='session', entity_id=new_session_id,
            session_id=new_session_id, action='create',
            field_name='copy_from', new_value=str(session_id),
            details={'source_session_id': session_id, 'sheets_copied': len(sheets)},
        )
        conn.commit()
        log_action(
            current_username(),
            f"скопіював сесію #{session_id} → #{new_session_id} "
            f"(група {new_group_id}, відомостей: {len(sheets)})",
        )
        conn.close()
        flash('Сесію скопійовано. Можна редагувати відомості.', 'success')
        return redirect(url_for('exam_sessions.session_detail', session_id=new_session_id))

    groups = allowed_groups_for_user(conn, session)
    # прибрати поточну групу зі списку (опційно залишити)
    groups = [g for g in groups if g['id'] != es['group_id']]
    conn.close()
    return render_template(
        'exam_sessions/session_copy.html',
        es=es, groups=groups,
    )
