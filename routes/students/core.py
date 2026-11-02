"""
routes/students/core.py
========================
Список студентів (пошук/фільтри/пагінація), картка студента, додати/
редагувати/видалити студента, фото 3х4, заморозка, періоди навчання,
посилання на оновлення даних (/update-info), генерація документів для
одного студента, FAQ.
"""
from flask import render_template, request, redirect, url_for, session, flash
from routes.db import get_db
from routes.utils import log_action, login_required, permission_required, logger
from routes.helpers import current_username
from routes.students import students_bp
import sqlite3
from routes.utils import (
    generate_english_name, is_student_on_reduced_program, save_multiple_attachments,
    program_track_condition, apply_track_overrides, get_attachments, get_templates_with_metadata,
)
from routes.gen_docx import gen_doc
from routes import office_editor
from routes.utils import get_templates_with_metadata
from routes.students.grades import _groups_for_student, _user_can_access_student
from routes.config import PUBLIC_APPLY_BASE_URL
from routes.gender import detect_gender_from_middle_name
from datetime import datetime, timedelta
import os
import uuid
import secrets
import json
from routes.passport_types import PASSPORT_TYPES_ID_STYLE, passport_template_context
from routes.orders_service import get_student_orders, orders_tables_ready
from routes.user_scope import get_allowed_group_ids
from routes.enrollments_service import (
    ensure_enrollments_table,
    list_student_enrollments,
    add_enrollment,
    set_primary_enrollment,
    end_enrollment,
)

UPDATE_REQUEST_FIELD_LABELS = {
    'last_name_UA': "Прізвище (українською)",
    'first_name_UA': "Ім'я (українською)",
    'middle_name_UA': "По батькові (українською)",
    'phone': "Телефон",
    'phone_backup': "Резервний телефон",
    'email': "Email",
    'tax_id': "Ідентифікаційний код",
    'edebo_code': "Код ЄДЕБО",
    'photo': "Фото 3х4",
    'passport': "Паспортні дані (окремий розділ)",
    'education_document': "Документ про освіту (окремий розділ)",
    'education_document_reference': "— Академічна довідка",
    'education_document_recognition': "— Свідоцтво про визнання в Україні іноземного документа про освіту",
    'military': "Військові дані (окремий розділ)",
}

# --------------------------------------------------------------------
# Військові поля — розрахунок заповненості ЗАЛЕЖНО ВІД СТАТУСУ
# --------------------------------------------------------------------
# «Ядро» — є в БУДЬ-ЯКОМУ PDF «Резерв+», незалежно від категорії
# (призовник, військовозобов'язаний, знято з обліку, резервіст).
MILITARY_COMMON_FIELDS = [
    'status_badge',
    'valid_until',
    'registration_category',
    'tck_name',
    'registry_number',
    'data_update_date',
    # data_verified_on_time рахується окремо (0/1)
]

# Специфічні для статусу поля. Якщо статус не в списку — вважаємо,
# що обов'язкові лише «ядро» + ВЛК (як у призовника).
MILITARY_STATUS_FIELDS = {
    "Військовозобов'язаний": [
        'military_rank', 'vos_code', 'vos_description',
        'deferral_type', 'deferral_until',
        'vlk_decision', 'vlk_date',
    ],
    "Військовозобов’язаний": [
        'military_rank', 'vos_code', 'vos_description',
        'deferral_type', 'deferral_until',
        'vlk_decision', 'vlk_date',
    ],
    'Призовник': [
        'vlk_decision', 'vlk_date',
    ],
    'Резервіст': [
        'military_rank', 'vos_code', 'vos_description',
        'deferral_type', 'deferral_until',
        'vlk_decision', 'vlk_date',
    ],
    "Невійськовозобов'язаний": [
        'vlk_decision', 'vlk_date',
    ],
    "Невійськовозобов’язаний": [
        'vlk_decision', 'vlk_date',
    ],
    'Знято з обліку': [
        'removal_basis',
    ],
}

# Поля, які можуть бути відсутні у PDF («—») — їх не додаємо в
# обов'язкові, якщо в конкретного студента там порожньо.
_OPTIONAL_IF_EMPTY = {
    'deferral_type', 'deferral_until',      # немає відстрочки — не вимагаємо
    'vlk_decision', 'vlk_date',             # немає ВЛК — не вимагаємо
    'removal_basis',                        # не знято — не вимагаємо
    'police_reason', 'police_date',         # не звертався до поліції
    'disability_group', 'disability_valid_until', 'disability_reason',
}


def _is_field_applicable(field, military_row):
    """
    Повертає True, якщо поле має бути заповнене для цього студента.
    Для полів зі списку _OPTIONAL_IF_EMPTY — перевіряємо, чи в БД
    вже щось є (якщо порожньо/«—» — не вимагаємо).
    """
    if military_row is None:
        return False
    value = military_row.get(field)
    if field in _OPTIONAL_IF_EMPTY:
        return bool(value) and str(value).strip() not in ('', '—', '-')
    return True


def _get_applicable_military_fields(status_badge, military_row):
    """Повертає список полів military, які мають бути заповнені для цього студента."""
    fields = list(MILITARY_COMMON_FIELDS)
    status = (status_badge or '').strip()
    status_fields = MILITARY_STATUS_FIELDS.get(status, ['vlk_decision', 'vlk_date'])
    for f in status_fields:
        if _is_field_applicable(f, military_row):
            fields.append(f)
    return fields


def _compute_military_completeness(military_row):
    """
    Повертає (filled, total) для одного запису military з урахуванням
    статусу. Якщо military_row is None — (0, len(common)+1).
    """
    if not military_row:
        return 0, len(MILITARY_COMMON_FIELDS) + 1
    applicable = _get_applicable_military_fields(
        military_row.get('status_badge'), military_row
    )
    filled = sum(
        1 for f in applicable
        if military_row.get(f) and str(military_row.get(f)).strip()
    )
    if military_row.get('data_verified_on_time'):
        filled += 1
    total = len(applicable) + 1
    return filled, total


FAQ_SECTION_PERMISSIONS = {
    'students': [
        'manage_students',
        'manage_passport_documents',
        'manage_education_documents',
        'manage_pending_students',
        'study_periods',
    ],
    'education': [
        'manage_specialties',
        'manage_degree_levels',
        'manage_educational_programs',
        'manage_qualification_names',
        'manage_accreditations',
    ],
    'courses': [
        'manage_courses',
        'manage_groups',
        'manage_subjects',
        'manage_activities',
        'manage_expulsion',
        'manage_frozen_students',
        'archive',
    ],
    'import': [
        'import_from_excel',
        'import_passport_documents',
        'import_subjects',
        'import_education_docs',
        'import_grades',
        'import_military_pdf',
        'bulk_photos',
    ],
    'bulk': [
        'group_export',
        'manage_diplomas',
        'manage_templates',
    ],
    'system': [
        'manage_users',
        'view_logs',
    ],
    'analytics': [
        'analytics',
    ],
    'military': [
        'manage_military_reports',
    ],
    'licenses': [
        'manage_licenses',
        'manage_license_transfer',
        'license_report',
    ],
}


def _refresh_session_permissions():
    """
    Підтягує is_admin / permissions з БД у сесію.
    Потрібно після зміни прав адміном — інакше FAQ і меню
    показують старий список до повторного входу.
    """
    import json
    user_id = session.get('user_id')
    if not user_id:
        return
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT is_admin, role, permissions FROM users WHERE id = ?",
            (user_id,),
        ).fetchone()
    finally:
        conn.close()
    if not row:
        return
    # sqlite Row or tuple
    try:
        is_admin = bool(row['is_admin']) or (row['role'] == 'admin')
        raw = row['permissions']
    except (TypeError, KeyError, IndexError):
        is_admin = bool(row[0]) or (row[1] == 'admin')
        raw = row[2]
    try:
        perms = json.loads(raw or '[]')
    except Exception:
        perms = []
    if not isinstance(perms, list):
        perms = []
    session['is_admin'] = is_admin
    session['permissions'] = perms
    session.modified = True


def _allowed_faq_sections():
    """Які розділи FAQ доступні поточному користувачу (з актуальними правами)."""
    _refresh_session_permissions()
    if session.get('is_admin'):
        return set(FAQ_SECTION_PERMISSIONS.keys())
    perms = set(session.get('permissions') or [])
    allowed = set()
    for section, need in FAQ_SECTION_PERMISSIONS.items():
        if perms.intersection(need):
            allowed.add(section)
    return allowed


@students_bp.route('/faq')
@login_required('')
def faq():
    """Довідка (FAQ) — лише розділи за актуальними правами."""
    sections = _allowed_faq_sections()
    return render_template('faq.html', faq_sections=sections)

@students_bp.route('/students', methods=['GET'])
@login_required('')
def student_list():
    """Головна сторінка списку студентів."""
    search = request.args.get('search', '')
    group_id = request.args.get('group_id', type=int)
    degree_level = request.args.get('degree_level', '')
    study_form = request.args.get('study_form', '')
    course = request.args.get('course', type=int)
    license_id = request.args.get('license_id', type=int)
    reduced = request.args.get('reduced', '')
    include_frozen = '1' if request.args.get('include_frozen') == '1' else ''
    include_archived = '1' if request.args.get('include_archived') == '1' else ''
    page = max(1, request.args.get('page', 1, type=int))
    per_page = request.args.get('per_page', 10, type=int)
    sort_by = request.args.get('sort_by', 'id')
    sort_order = request.args.get('sort_order', 'desc')

    if per_page not in [10, 20, 50, 100]:
        per_page = 10

    allowed_sort_fields = ['id', 'last_name_UA', 'first_name_UA', 'middle_name_UA', 'birth_date', 'group_id']
    if sort_by not in allowed_sort_fields:
        sort_by = 'id'
    if sort_order not in ['asc', 'desc']:
        sort_order = 'desc'

    offset = (page - 1) * per_page

    conn = get_db()
    role = session.get('role')
    user_id = session.get('user_id')

    group_ids = get_allowed_group_ids(conn, user_id)

    # ── SELECT з усіма новими військовими полями ──
    base_query = """
        SELECT s.*,
               m.id AS has_military,
               m.status_badge,
               m.valid_until,
               m.registration_category,
               m.removal_basis,
               m.tck_name,
               m.military_rank,
               m.registry_number,
               m.vos_code,
               m.vos_description,
               m.deferral_type,
               m.deferral_until,
               m.police_reason,
               m.police_date,
               m.vlk_decision,
               m.vlk_date,
               m.disability_group,
               m.disability_valid_until,
               m.disability_reason,
               m.data_verified_on_time,
               m.data_update_date,
               m.military_registration_document,
               m.issued_VOD,
               m.military_accounting_specialty_number,
               m.change_credentials,
               m.reason_for_changing_credentials,
               m.being_on_military_registration,
               g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS group_name,
               g.study_form,
               g.program_credits,
               g.qualification_name,
               g.degree_level,
               g.specialty,
               g.educational_program,
               g.knowledge_area,
               g.qualification_name_en,
               g.degree_level_en,
               g.specialty_en,
               g.isced_f_ua,
               g.isced_f_en,
               g.educational_program_en,
               g.knowledge_area_en,
               -- Список активних enrollments (паралельне навчання). Використовуємо
               -- той самий підхід, що й в analytics.py::student_options() —
               -- GROUP_CONCAT через ' | ' (потім у Python розбиваємо).
               (SELECT GROUP_CONCAT(g2.name, ' | ')
                FROM student_enrollments e2
                JOIN groups g2 ON g2.id = e2.group_id
                WHERE e2.student_id = s.id AND e2.status = 'active') AS all_groups,
               (SELECT COUNT(*) FROM student_enrollments e2
                WHERE e2.student_id = s.id AND e2.status = 'active') AS enrollment_count,
               fs.id AS frozen_record_id,
               fs.reason AS frozen_reason,
               pg.name AS frozen_previous_group_name
        FROM students s
        LEFT JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON s.group_id = g.id
        LEFT JOIN frozen_students fs ON fs.student_id = s.id AND fs.resolved_at IS NULL
        LEFT JOIN groups pg ON pg.id = fs.previous_group_id
    """
    count_query = "SELECT COUNT(*) FROM students s LEFT JOIN groups g ON s.group_id = g.id"
    where_clauses = []
    if not include_archived:
        where_clauses.append("s.archived = FALSE")
    if not include_frozen:
        where_clauses.append("s.id NOT IN (SELECT student_id FROM frozen_students WHERE resolved_at IS NULL)")
    params = []

    if group_id:
        where_clauses.append("""
            (
                s.group_id = ?
                OR s.id IN (
                    SELECT student_id FROM student_enrollments
                    WHERE group_id = ? AND status = 'active'
                )
            )
        """)
        params.extend([group_id, group_id])

    if degree_level:
        where_clauses.append("g.degree_level = ?")
        params.append(degree_level)

    if study_form:
        where_clauses.append("g.study_form = ?")
        params.append(study_form)

    if course:
        where_clauses.append("g.course = ?")
        params.append(course)

    if license_id:
        where_clauses.append("s.license_id = ?")
        params.append(license_id)

    if reduced == '1':
        where_clauses.append("s.program_credits_override IS NOT NULL AND s.program_credits_override < g.program_credits")
    elif reduced == '0':
        where_clauses.append("(s.program_credits_override IS NULL OR s.program_credits_override >= g.program_credits)")

    if role != 'admin' and not group_id:
        if group_ids:
            placeholders = ','.join('?' for _ in group_ids)
            where_clauses.append(f"""
                (
                    s.group_id IN ({placeholders})
                    OR s.id IN (
                        SELECT student_id FROM student_enrollments
                        WHERE status = 'active' AND group_id IN ({placeholders})
                    )
                )
            """)
            params.extend(list(group_ids) + list(group_ids))
        else:
            conn.close()
            return render_template(
                'students.html',
                students=[],
                search=search, group_id=group_id, groups=[], groups_by_degree={},
                degree_level=degree_level, degree_levels_available=[],
                study_form=study_form, study_forms_available=[],
                course=course, courses_available=[],
                license_id=license_id, licenses_available=[],
                reduced=reduced, include_frozen=include_frozen, include_archived=include_archived,
                page=page, per_page=per_page, total_pages=0, total_students=0,
                sort_by=sort_by, sort_order=sort_order
            )

    if role != 'admin' and group_id and group_id not in group_ids:
        conn.close()
        flash("У вас немає доступу до цієї групи.", "error")
        return redirect(url_for('students.student_list'))

    if search:
        for word in search.split():
            where_clauses.append(
                "(LOWER_UA(s.last_name_UA) LIKE LOWER_UA(?) OR LOWER_UA(s.first_name_UA) LIKE LOWER_UA(?) OR LOWER_UA(s.middle_name_UA) LIKE LOWER_UA(?))"
            )
            params.extend([f'%{word}%', f'%{word}%', f'%{word}%'])

    if sort_by == 'birth_date':
        where_clauses.append("LENGTH(s.birth_date) = 10 AND INSTR(s.birth_date, '.') = 3 AND INSTR(SUBSTR(s.birth_date, 4), '.') = 3")

    if where_clauses:
        where_sql = " WHERE " + " AND ".join(where_clauses)
        base_query += where_sql
        count_query += where_sql

    if sort_by == 'birth_date':
        base_query += f" ORDER BY SUBSTR(s.birth_date, 7, 4) || SUBSTR(s.birth_date, 4, 2) || SUBSTR(s.birth_date, 1, 2) {sort_order.upper()}"
    elif sort_by in ['last_name_UA', 'first_name_UA', 'middle_name_UA']:
        base_query += f" ORDER BY s.{sort_by} COLLATE UKRAINIAN {sort_order.upper()}"
    else:
        base_query += f" ORDER BY s.{sort_by} {sort_order.upper()}"

    base_query += " LIMIT ? OFFSET ?"
    params_with_limit = params + [per_page, offset]

    students = conn.execute(base_query, params_with_limit).fetchall()
    total_students = conn.execute(count_query, params).fetchone()[0]

    students_with_filled_fields = []
    student_fields = ['last_name_UA', 'first_name_UA', 'middle_name_UA', 'last_name_ENG', 'first_name_ENG', 'birth_date', 'group_id', 'edebo_code', 'tax_id', 'phone', 'email']
    student_total_fields = len(student_fields)

    for student in students:
        student_dict = dict(student)

        if student_dict.get('frozen_record_id'):
            student_dict['status'] = 'frozen'
        elif student_dict.get('archived'):
            student_dict['status'] = 'archived'
        else:
            student_dict['status'] = 'active'

        student_filled_fields = sum(1 for field in student_fields if student_dict.get(field) and str(student_dict.get(field)).strip())
        student_dict['filled_fields'] = student_filled_fields
        student_dict['total_fields'] = student_total_fields
        student_dict['has_photo'] = bool(student_dict.get('photo'))

        # ── Паралельне навчання: розбираємо GROUP_CONCAT у список ──
        raw_groups = (student_dict.get('all_groups') or '').strip()
        student_dict['all_groups_list'] = [
            name.strip() for name in raw_groups.split('|') if name.strip()
        ]
        # Дублюємо primary у all_groups, якщо його там з якоїсь причини
        # немає (наприклад, у старих даних enrollment ще не забекфілений
        # або s.group_id змінили вручну). Це гарантує, що primary завжди
        # присутній в all_groups_list.
        if student_dict.get('group_name'):
            primary_short = (student_dict['group_name'].split(' (')[0]).strip()
            if primary_short and primary_short not in student_dict['all_groups_list']:
                student_dict['all_groups_list'].insert(0, primary_short)

        student_dict['is_parallel'] = len(student_dict['all_groups_list']) > 1
        student_dict['extra_groups_count'] = max(0, len(student_dict['all_groups_list']) - 1)

        # ── Заповненість військових даних (залежно від статусу) ──
        military_row_for_calc = student_dict if student_dict.get('has_military') else None
        mil_filled, mil_total = _compute_military_completeness(military_row_for_calc)
        student_dict['military_filled_fields'] = mil_filled
        student_dict['military_total_fields'] = mil_total
        # Чи є PDF «Резерв+» серед сканів?
        if student_dict.get('has_military'):
            pdf_row = conn.execute(
                "SELECT 1 FROM attachments "
                "WHERE entity_type='military' AND entity_id=? "
                "AND (original_name LIKE '%.pdf' OR file_path LIKE '%.pdf') "
                "LIMIT 1",
                (student_dict['has_military'],)
            ).fetchone()
            student_dict['military_has_pdf'] = pdf_row is not None
        else:
            student_dict['military_has_pdf'] = False

        is_reduced = bool(student_dict['group_id']) and is_student_on_reduced_program(conn, student_dict['id'], student_dict['group_id'])
        subjects_query = "SELECT id FROM subjects WHERE group_id = ?" + program_track_condition(is_reduced)
        subjects = conn.execute(subjects_query, (student_dict['group_id'],)).fetchall()
        subject_ids = [s['id'] for s in subjects]
        student_dict['has_grades'] = len(subjects) > 0
        if subject_ids:
            placeholders = ','.join('?' for _ in subject_ids)
            grades = conn.execute(
                f"SELECT grade FROM grades WHERE student_id = ? AND subject_id IN ({placeholders})",
                (student_dict['id'], *subject_ids)
            ).fetchall()
        else:
            grades = []
        grades_filled = sum(1 for grade in grades if grade['grade'] is not None and str(grade['grade']).strip())
        grades_total = len(subjects)
        student_dict['grades_filled'] = grades_filled
        student_dict['grades_total'] = grades_total
        student_dict['grades_fill_percentage'] = (grades_filled / grades_total * 100) if grades_total > 0 else 0

        practices = conn.execute("SELECT id FROM practices WHERE group_id = ?", (student_dict['group_id'],)).fetchall()
        courseworks = conn.execute("SELECT id FROM courseworks WHERE group_id = ?", (student_dict['group_id'],)).fetchall()
        attestations = conn.execute("SELECT id FROM attestations WHERE group_id = ?", (student_dict['group_id'],)).fetchall()
        student_dict['has_activities'] = len(practices) > 0 or len(courseworks) > 0 or len(attestations) > 0
        activities_grades = conn.execute(
            "SELECT grade FROM activity_grades WHERE student_id = ? AND entity_type IN ('practice', 'coursework', 'attestation')",
            (student_dict['id'],)
        ).fetchall()
        activities_total = len(practices) + len(courseworks) + len(attestations)
        activities_filled = sum(1 for grade in activities_grades if grade['grade'] is not None and str(grade['grade']).strip())
        student_dict['activities_filled_fields'] = activities_filled
        student_dict['activities_total_fields'] = activities_total
        student_dict['activities_fill_percentage'] = (activities_filled / activities_total * 100) if activities_total > 0 else 0

        passport_doc = conn.execute("""
            SELECT document_type, series, number, issued_by, issue_date, valid_until, unique_number
            FROM passport_documents WHERE student_id = ? ORDER BY id DESC LIMIT 1
        """, (student_dict['id'],)).fetchone()
        student_dict['has_passport'] = passport_doc is not None
        if passport_doc:
            pd = dict(passport_doc)
            if pd.get('document_type') in PASSPORT_TYPES_ID_STYLE:
                passport_fields = ['document_type', 'number', 'issued_by', 'issue_date', 'valid_until', 'unique_number']
            else:
                passport_fields = ['document_type', 'series', 'number', 'issued_by', 'issue_date']
            passport_filled = sum(1 for f in passport_fields if pd.get(f) and str(pd.get(f)).strip())
            student_dict['passport_filled_fields'] = passport_filled
            student_dict['passport_total_fields'] = len(passport_fields)
        else:
            student_dict['passport_filled_fields'] = 0
            student_dict['passport_total_fields'] = 0

        edu_doc = conn.execute("""
            SELECT document_type, document_type_en, document_number, institution_name,
                   institution_name_en, country, country_en, completion_date
            FROM education_documents WHERE student_id = ? ORDER BY id DESC LIMIT 1
        """, (student_dict['id'],)).fetchone()
        student_dict['has_education_doc'] = edu_doc is not None
        if edu_doc:
            ed = dict(edu_doc)
            edu_fields = ['document_type', 'document_type_en', 'document_number', 'institution_name',
                          'institution_name_en', 'country', 'country_en', 'completion_date']
            edu_filled = sum(1 for f in edu_fields if ed.get(f) and str(ed.get(f)).strip())
            student_dict['edu_doc_filled_fields'] = edu_filled
            student_dict['edu_doc_total_fields'] = len(edu_fields)
        else:
            student_dict['edu_doc_filled_fields'] = 0
            student_dict['edu_doc_total_fields'] = 0

        students_with_filled_fields.append(student_dict)

    groups = conn.execute("""
        SELECT id, name, start_year, study_form, degree_level, course, program_credits, specialty
        FROM groups
        WHERE archived = FALSE
        ORDER BY start_year DESC, name COLLATE UKRAINIAN
    """).fetchall()

    degree_order = {
        row['id_name']: row['ord']
        for row in conn.execute("SELECT name_ua AS id_name, id AS ord FROM degree_levels").fetchall()
    }
    groups_by_degree_raw = {}
    for g in groups:
        groups_by_degree_raw.setdefault(g['degree_level'] or 'Без ступеня', []).append(g)
    groups_by_degree = {
        d: groups_by_degree_raw[d]
        for d in sorted(groups_by_degree_raw.keys(), key=lambda d: (degree_order.get(d, 999), d))
    }

    degree_levels_available = sorted(
        {g['degree_level'] for g in groups if g['degree_level']},
        key=lambda d: (degree_order.get(d, 999), d)
    )
    study_forms_available = sorted({g['study_form'] for g in groups if g['study_form']})
    courses_available = sorted({g['course'] for g in groups if g['course']})
    licenses_available = conn.execute(
        "SELECT id, short_name_ua FROM institution_licenses WHERE is_active = 1 ORDER BY id"
    ).fetchall()

    conn.close()

    total_pages = (total_students + per_page - 1) // per_page

    return render_template(
        'students.html',
        students=students_with_filled_fields,
        search=search, group_id=group_id, groups=groups,
        groups_by_degree=groups_by_degree,
        degree_level=degree_level, degree_levels_available=degree_levels_available,
        study_form=study_form, study_forms_available=study_forms_available,
        course=course, courses_available=courses_available,
        license_id=license_id, licenses_available=licenses_available,
        reduced=reduced, include_frozen=include_frozen, include_archived=include_archived,
        page=page, per_page=per_page,
        total_pages=total_pages, total_students=total_students,
        sort_by=sort_by, sort_order=sort_order
    )


@students_bp.route('/students/<int:student_id>')
@login_required('')
def student_details(student_id):
    """Картка студента."""
    conn = get_db()
    student = conn.execute("""
        SELECT s.*, g.program_credits, g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS group_name,
               il.name_ua AS license_name_ua, il.short_name_ua AS license_short_name_ua
        FROM students s
        LEFT JOIN groups g ON s.group_id = g.id
        LEFT JOIN institution_licenses il ON s.license_id = il.id
        WHERE s.id = ?
    """, (student_id,)).fetchone()

    if not student:
        log_action(current_username(), f"спроба перегляду неіснуючого студента (ID {student_id})")
        flash("Студента не знайдено")
        conn.close()
        return redirect(url_for('students.student_list'))

    log_action(
        current_username(),
        f"переглянув картку студента: {student['last_name_UA']} {student['first_name_UA']} (ID {student_id})",
        group_ids=[student['group_id']],
        details=f"група: {student['group_name']}"
    )

    # ── Enrollment'и завантажуємо ДО побудови groups_data ──
    ensure_enrollments_table(conn)
    enrollments = list_student_enrollments(student_id, conn=conn)

    military = conn.execute("SELECT * FROM military WHERE student_id = ?", (student_id,)).fetchone()
    military_attachments = get_attachments(conn, 'military', military['id']) if military else []
    military_history = conn.execute(
        """SELECT changed_at, changed_by, source, scope, field_name, old_value, new_value
           FROM military_history
           WHERE student_id = ?
           ORDER BY changed_at DESC, id DESC
           LIMIT 5""",
        (student_id,),
    ).fetchall()

    def _field_label(name, scope='military'):
        from routes.students.military import MILITARY_FIELDS, LEGACY_FIELDS, STUDENT_FIELDS
        src = (MILITARY_FIELDS + LEGACY_FIELDS) if scope == 'military' else STUDENT_FIELDS
        for key, label in src:
            if key == name:
                return label
        return name

    # ── Дані по КОЖНІЙ групі, в якій навчається студент ──
    active_enrollments = [e for e in enrollments if e['status'] == 'active']
    if student['group_id'] and not any(
        e['group_id'] == student['group_id'] for e in active_enrollments
    ):
        active_enrollments.insert(0, {
            'group_id': student['group_id'],
            'group_name': student['group_name'],
            'is_primary': 1,
            'program_credits': student['program_credits'],
        })

    groups_data = []
    for enr in active_enrollments:
        gid = enr['group_id']
        group_program_credits = enr.get('program_credits')
        is_reduced_g = is_student_on_reduced_program(conn, student_id, gid)

        subjects_query = """
            SELECT s.id, s.code, s.name, s.type, s.reduced_type,
                   s.credits, s.reduced_credits, g.grade
            FROM subjects s
            LEFT JOIN grades g ON g.subject_id = s.id AND g.student_id = ?
            WHERE s.group_id = ?
        """ + program_track_condition(is_reduced_g, 's') + " ORDER BY s.position"
        subj_rows = conn.execute(subjects_query, [student_id, gid]).fetchall()
        subjects_list = [
            {**apply_track_overrides(dict(s), is_reduced_g), 'grade': s['grade']}
            for s in subj_rows
        ]

        practices_query = """
            SELECT p.id, p.code, p.name, p.type, p.reduced_type,
                   p.credits, p.reduced_credits, ag.grade
            FROM practices p
            LEFT JOIN activity_grades ag
                   ON ag.entity_id = p.id AND ag.entity_type = 'practice'
                  AND ag.student_id = ?
            WHERE p.group_id = ?
        """ + program_track_condition(is_reduced_g, 'p') + " ORDER BY p.position"
        prac_rows = conn.execute(practices_query, [student_id, gid]).fetchall()
        practices_list = [
            {**apply_track_overrides(dict(p), is_reduced_g), 'grade': p['grade']}
            for p in prac_rows
        ]

        courseworks_query = """
            SELECT c.id, c.code, c.name, c.type, c.reduced_type,
                   c.credits, c.reduced_credits, ag.grade
            FROM courseworks c
            LEFT JOIN activity_grades ag
                   ON ag.entity_id = c.id AND ag.entity_type = 'coursework'
                  AND ag.student_id = ?
            WHERE c.group_id = ?
        """ + program_track_condition(is_reduced_g, 'c') + " ORDER BY c.position"
        cw_rows = conn.execute(courseworks_query, [student_id, gid]).fetchall()
        courseworks_list = [
            {**apply_track_overrides(dict(c), is_reduced_g), 'grade': c['grade']}
            for c in cw_rows
        ]

        attestations_query = """
            SELECT a.id, a.code, a.name, a.type, a.reduced_type,
                   a.credits, a.reduced_credits, ag.grade, ag.name AS student_name
            FROM attestations a
            LEFT JOIN activity_grades ag
                   ON ag.entity_id = a.id AND ag.entity_type = 'attestation'
                  AND ag.student_id = ?
            WHERE a.group_id = ?
        """ + program_track_condition(is_reduced_g, 'a') + " ORDER BY a.position"
        att_rows = conn.execute(attestations_query, [student_id, gid]).fetchall()
        attestations_list = [
            {**apply_track_overrides(dict(a), is_reduced_g),
             'grade': a['grade'], 'student_name': a['student_name']}
            for a in att_rows
        ]

        # Коротка назва групи (без " (2022, Денна, 240 кредитів)")
        short_name = (enr.get('group_name') or '').split(' (')[0].strip()

        groups_data.append({
            'group_id': gid,
            'group_name': enr.get('group_name') or '',
            'short_group_name': short_name or f"Група #{gid}",
            'is_primary': bool(enr['is_primary']),
            'is_reduced': is_reduced_g,
            'group_program_credits': group_program_credits,
            'student_credits_override': student['program_credits_override'],
            'subjects': subjects_list,
            'practices': practices_list,
            'courseworks': courseworks_list,
            'attestations': attestations_list,
        })

    # ── Backward compatibility: старі змінні → на ОСНОВНУ групу ──
    _primary = next(
        (g for g in groups_data if g['is_primary']),
        groups_data[0] if groups_data else None
    )
    subject_grades = _primary['subjects'] if _primary else []
    practice_data = _primary['practices'] if _primary else []
    coursework_data = _primary['courseworks'] if _primary else []
    attestation_data = _primary['attestations'] if _primary else []

    education_docs = conn.execute("""
        SELECT ed.id, ed.document_type, ed.document_type_en, ed.document_number,
               ed.institution_name, ed.institution_name_en, ed.country, ed.country_en, ed.completion_date,
               fed.reference_number, fed.reference_institution, fed.reference_institution_en,
               fed.reference_country, fed.reference_country_en, fed.reference_issue_date,
               fed.recognition_certificate_number, fed.recognition_issuer,
               fed.recognition_issuer_en, fed.recognition_date
        FROM education_documents ed
        LEFT JOIN foreign_education_docs fed ON ed.id = fed.education_doc_id
        WHERE ed.student_id = ?
        ORDER BY ed.completion_date DESC
    """, (student_id,)).fetchall()
    education_docs = [dict(d) for d in education_docs]
    for d in education_docs:
        d['attachments'] = get_attachments(conn, 'education_document', d['id'])

    student_attachments = get_attachments(conn, 'student', student_id)

    passport_docs = conn.execute(
        "SELECT id, document_type, series, number, issued_by, issue_date, valid_until, unique_number "
        "FROM passport_documents WHERE student_id = ? ORDER BY id DESC", (student_id,)
    ).fetchall()
    passport_docs = [dict(d) for d in passport_docs]
    for d in passport_docs:
        d['attachments'] = get_attachments(conn, 'passport_document', d['id'])

    study_periods = conn.execute("""
        SELECT id, filiya, filiya_en, group_name, start_date, end_date, period_order, note
        FROM student_study_periods
        WHERE student_id = ?
        ORDER BY period_order ASC, start_date ASC
    """, (student_id,)).fetchall()

    student_dict = dict(student)
    military_dict = dict(military) if military else None

    student_orders = []
    if orders_tables_ready(conn=conn):
        student_orders = get_student_orders(student_id, conn=conn)

    all_groups = conn.execute("""
        SELECT id,
               name || ' (' || start_year || ', ' || study_form || ', ' ||
               COALESCE(degree_level, '') || ')' AS display_name
        FROM groups
        WHERE COALESCE(archived, 0) = 0
        ORDER BY start_year DESC, name COLLATE UKRAINIAN
    """).fetchall()

    conn.close()

    return render_template(
        'student_details.html',
        student=student_dict,
        military=military_dict,
        military_attachments=military_attachments,
        enrollments=enrollments,
        all_groups=all_groups,
        groups_data=groups_data,
        military_history=military_history,
        field_label=_field_label,
        subject_grades=subject_grades,
        practice_data=practice_data,
        coursework_data=coursework_data,
        attestation_data=attestation_data,
        education_docs=education_docs,
        student_attachments=student_attachments,
        passport_docs=passport_docs,
        study_periods=study_periods,
        student_orders=student_orders,
        **passport_template_context(),
    )

@students_bp.route('/students/<int:student_id>/study_periods', methods=['GET', 'POST'])
@permission_required('study_periods')
def manage_study_periods(student_id):
    """CRUD-сторінка періодів навчання конкретного студента."""
    db = get_db()
    cursor = db.cursor()

    cursor.execute("""
        SELECT id, last_name_UA, first_name_UA, middle_name_UA
        FROM students WHERE id = ?
    """, (student_id,))
    student = cursor.fetchone()
    if not student:
        flash('Студента не знайдено', 'danger')
        return redirect(url_for('admin.manage_education_documents'))

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            filiya = (request.form.get('filiya') or '').strip()
            filiya_en = (request.form.get('filiya_en') or '').strip() or None
            group_name = (request.form.get('group_name') or '').strip() or None
            start_date = (request.form.get('start_date') or '').strip() or None
            end_date = (request.form.get('end_date') or '').strip() or None
            note = (request.form.get('note') or '').strip() or None

            if not filiya:
                flash('Потрібно вказати філію', 'danger')
            else:
                next_order = cursor.execute(
                    "SELECT COALESCE(MAX(period_order), 0) + 1 FROM student_study_periods WHERE student_id = ?",
                    (student_id,)
                ).fetchone()[0]
                cursor.execute("""
                    INSERT INTO student_study_periods
                        (student_id, filiya, filiya_en, group_name, start_date, end_date, period_order, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (student_id, filiya, filiya_en, group_name, start_date, end_date, next_order, note))
                db.commit()
                log_action(current_username(),
                           f"додав період навчання студенту ID {student_id}: {filiya}")
                flash('Період навчання додано', 'success')

        elif action == 'edit':
            period_id = request.form.get('period_id')
            filiya = (request.form.get('filiya') or '').strip()
            filiya_en = (request.form.get('filiya_en') or '').strip() or None
            group_name = (request.form.get('group_name') or '').strip() or None
            start_date = (request.form.get('start_date') or '').strip() or None
            end_date = (request.form.get('end_date') or '').strip() or None
            period_order = request.form.get('period_order', type=int) or 0
            note = (request.form.get('note') or '').strip() or None

            if not filiya:
                flash('Потрібно вказати філію', 'danger')
            else:
                cursor.execute("""
                    UPDATE student_study_periods SET
                        filiya=?, filiya_en=?, group_name=?, start_date=?, end_date=?, period_order=?, note=?
                    WHERE id=? AND student_id=?
                """, (filiya, filiya_en, group_name, start_date, end_date, period_order, note, period_id, student_id))
                db.commit()
                log_action(current_username(),
                           f"оновив період навчання ID {period_id} студенту ID {student_id}")
                flash('Період навчання оновлено', 'success')

        elif action == 'delete':
            period_id = request.form.get('period_id')
            cursor.execute("""
                DELETE FROM student_study_periods WHERE id=? AND student_id=?
            """, (period_id, student_id))
            db.commit()
            log_action(current_username(),
                       f"видалив період навчання ID {period_id} студенту ID {student_id}")
            flash('Період навчання видалено', 'success')

        return redirect(url_for('students.manage_study_periods', student_id=student_id))

    cursor.execute("""
        SELECT id, filiya, filiya_en, group_name, start_date, end_date, period_order, note
        FROM student_study_periods
        WHERE student_id = ?
        ORDER BY period_order ASC, start_date ASC
    """, (student_id,))
    periods = cursor.fetchall()

    return render_template(
        'manage_study_periods.html',
        student=student,
        periods=periods
    )


@students_bp.route('/students/add', methods=['GET', 'POST'])
@login_required('')
def add_student():
    """Форма додавання нового студента."""
    conn = get_db()
    role = session.get('role')
    group_ids = session.get('group_ids', [])

    if role == 'admin':
        groups = conn.execute("""
            SELECT id, name, start_year, study_form, program_credits,
                   name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
            FROM groups WHERE archived = FALSE ORDER BY name, start_year
        """).fetchall()
    else:
        placeholders = ','.join('?' for _ in group_ids)
        groups = conn.execute(f"""
            SELECT id, name, start_year, study_form, program_credits,
                   name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
            FROM groups WHERE id IN ({placeholders}) AND archived = FALSE ORDER BY name, start_year
        """, group_ids).fetchall()

    licenses = conn.execute(
        "SELECT id, name_ua, short_name_ua FROM institution_licenses WHERE is_active = 1 ORDER BY id"
    ).fetchall()

    if request.method == 'POST':
        group = request.form.get('group_id')
        try:
            group_int = int(group)
        except (ValueError, TypeError):
            flash("Некоректна група", "error")
            conn.close()
            return render_template('add_student.html', groups=groups, licenses=licenses)

        if role != 'admin' and group_int not in group_ids:
            flash("Доступ заборонено: група не належить до ваших груп", "error")
            conn.close()
            return render_template('add_student.html', groups=groups, licenses=licenses)

        birth_date_raw = request.form['birth_date'].strip()
        birth_date_clean = birth_date_raw.replace("-", ".")
        try:
            datetime.strptime(birth_date_clean, "%d.%m.%Y")
            birth_date = birth_date_clean
        except ValueError:
            flash("Невірний формат дати. Введіть у форматі ДД.ММ.РРРР")
            conn.close()
            return render_template('add_student.html', groups=groups, licenses=licenses)

        last_name_ua = request.form['last_name_UA']
        first_name_ua = request.form['first_name_UA']
        middle_name_ua = request.form.get('middle_name_UA', '')
        last_name_eng, first_name_eng = generate_english_name(last_name_ua, first_name_ua)

        # ── Усі поля з форми читаємо ОДИН раз, ДО будь-яких UPDATE. ──
        # Раніше gender/license_id/phone/... визначались ПІСЛЯ if existing:,
        # тому в цій гілці UPDATE падав з NameError → 500 при повторному
        # зарахуванні архівного студента.
        gender_form = (request.form.get('gender') or 'auto').strip()
        if gender_form in ('M', 'F'):
            gender = gender_form
        else:
            gender = detect_gender_from_middle_name(middle_name_ua)

        license_id = request.form.get('license_id') or None
        phone = (request.form.get('phone') or '').strip() or None
        phone_backup = (request.form.get('phone_backup') or '').strip() or None
        email = (request.form.get('email') or '').strip() or None
        address_of_residence = (request.form.get('address_of_residence') or '').strip() or None
        edebo_code_form = (request.form.get('edebo_code') or '').strip() or None

        # ── Перевірка: чи ця людина вже є в системі (активна чи архівна)? ──
        tax_id_form = (request.form.get('tax_id') or '').strip() or None
        existing = None
        if tax_id_form:
            existing = conn.execute(
                "SELECT id, archived FROM students WHERE tax_id = ?",
                (tax_id_form,),
            ).fetchone()

        # Fallback по ПІБ+дата ЛИШЕ якщо tax_id НЕ введено — щоб не
        # зіставити двох різних людей з однаковим ПІБ+датою, але різними
        # ІПН (раніше цей fallback спрацьовував навіть коли ІПН був
        # введений, але не знаходився).
        if not existing and not tax_id_form:
            # birth_date у базі може бути '2000-05-15' / '15.05.2000' /
            # '15/05/2000' (імпорт з Excel, старі дані) — порівнюємо
            # нормалізовано, а не посимвольно.
            existing = conn.execute("""
                SELECT id, archived FROM students
                WHERE LOWER_UA(last_name_UA) = LOWER_UA(?)
                  AND LOWER_UA(first_name_UA) = LOWER_UA(?)
                  AND LOWER_UA(COALESCE(middle_name_UA, '')) = LOWER_UA(?)
                  AND REPLACE(REPLACE(REPLACE(COALESCE(birth_date,''), '-', '.'), '/', '.'), ' ', '')
                      = REPLACE(REPLACE(REPLACE(?, '-', '.'), '/', '.'), ' ', '')
            """, (last_name_ua, first_name_ua, middle_name_ua, birth_date)).fetchone()

        if existing:
            existing_id = existing['id']
            was_archived = bool(existing['archived'])

            already_in_group = conn.execute("""
                SELECT 1 FROM student_enrollments
                WHERE student_id = ? AND group_id = ? AND status = 'active'
            """, (existing_id, group_int)).fetchone()

            # Студент архівний, але enrollment у цій групі досі 'active'
            # (типовий наслідок archive_group без end_enrollment) —
            # просто знімаємо з архіву, БЕЗ попередження «вже навчається»
            # (це була б неправда: він архівний).
            if already_in_group and was_archived:
                conn.execute("UPDATE students SET archived = FALSE WHERE id = ?", (existing_id,))
                conn.commit()
                log_action(
                    current_username(),
                    f"розархівовано студента при повторному додаванні: "
                    f"{last_name_ua} {first_name_ua} (ID {existing_id})",
                    group_ids=[group_int],
                )
                conn.close()
                flash(
                    f"Студент {last_name_ua} {first_name_ua} уже був у цій групі в архіві — "
                    f"повернуто в активні.",
                    "success",
                )
                return redirect(url_for('students.student_details', student_id=existing_id))

            if already_in_group:
                flash(
                    f"Студент {last_name_ua} {first_name_ua} уже навчається в цій групі.",
                    "warning",
                )
                conn.close()
                return redirect(url_for('students.student_details', student_id=existing_id))

            # Активуємо enrollment (якщо в цій групі був expelled/graduated —
            # переактивується; якщо запису не було — створюється новий).
            ok_enroll, msg_enroll = add_enrollment(
                existing_id, group_int, is_primary=True, conn=conn,
            )
            if not ok_enroll:
                conn.rollback()
                conn.close()
                flash(f"Не вдалося додати навчання в групу: {msg_enroll}", "error")
                return redirect(url_for('students.student_details', student_id=existing_id))

            # add_enrollment уже знімає archived=FALSE при наявності
            # active enrollment — дублюємо для певності (idempotent).
            conn.execute("UPDATE students SET archived = FALSE WHERE id = ?", (existing_id,))

            # COALESCE: порожні поля у формі НЕ затирають наявні дані
            # (раніше UPDATE сліпо перезаписував усе — телефон, email,
            # ЄДЕБО, ліцензія, адреса, ПІБ англ., по батькові ставали
            # NULL, якщо у формі їх не заповнили).
            conn.execute("""
                UPDATE students SET
                    last_name_UA         = COALESCE(NULLIF(?, ''), last_name_UA),
                    first_name_UA        = COALESCE(NULLIF(?, ''), first_name_UA),
                    middle_name_UA       = COALESCE(NULLIF(?, ''), middle_name_UA),
                    last_name_ENG        = COALESCE(NULLIF(?, ''), last_name_ENG),
                    first_name_ENG       = COALESCE(NULLIF(?, ''), first_name_ENG),
                    birth_date           = COALESCE(NULLIF(?, ''), birth_date),
                    edebo_code           = COALESCE(?, edebo_code),
                    tax_id               = COALESCE(?, tax_id),
                    license_id           = COALESCE(?, license_id),
                    phone                = COALESCE(?, phone),
                    phone_backup         = COALESCE(?, phone_backup),
                    email                = COALESCE(?, email),
                    address_of_residence = COALESCE(?, address_of_residence),
                    gender               = COALESCE(?, gender)
                WHERE id=?
            """, (last_name_ua, first_name_ua, middle_name_ua or '',
                  last_name_eng, first_name_eng, birth_date,
                  edebo_code_form, tax_id_form, license_id,
                  phone, phone_backup, email, address_of_residence, gender,
                  existing_id))
            conn.commit()

            log_action(
                current_username(),
                f"повторно зарахував раніше архівного студента: "
                f"{last_name_ua} {first_name_ua} (ID {existing_id}) → група {group_int}",
                group_ids=[group_int],
            )
            conn.close()
            flash(
                f"Студент {last_name_ua} {first_name_ua} вже був у системі — "
                f"додано нове навчання в обраній групі.",
                "success",
            )
            return redirect(url_for('students.student_details', student_id=existing_id))

        # ── Новий студент ──
        # gender / license_id / phone / phone_backup / email /
        # address_of_residence уже визначені ВИЩЕ, тому тут НЕ
        # повторюємо їхні визначення.
        issued_VOD_raw = request.form.get('issued_VOD', '').strip()
        issued_VOD = None
        if issued_VOD_raw:
            issued_VOD_clean = issued_VOD_raw.replace("-", ".")
            try:
                datetime.strptime(issued_VOD_clean, "%d.%m.%Y")
                issued_VOD = issued_VOD_clean
            except ValueError:
                flash("Невірний формат дати видачі ВОД. Введіть у форматі ДД.ММ.РРРР")
                conn.close()
                return render_template('add_student.html', groups=groups, licenses=licenses)
        conn.execute("""
            INSERT INTO students (
                last_name_UA, first_name_UA, middle_name_UA,
                last_name_ENG, first_name_ENG, birth_date, group_id, edebo_code, tax_id, license_id,
                phone, phone_backup, email, address_of_residence, gender
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (last_name_ua, first_name_ua, middle_name_ua, last_name_eng, first_name_eng,
              birth_date, group_int, request.form.get('edebo_code'), request.form.get('tax_id') or None,
              license_id, phone, phone_backup, email, address_of_residence, gender))
        student_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        add_enrollment(student_id, group_int, is_primary=True, conn=conn)
        conn.commit()

        military_fields_list = [
            'status_badge', 'valid_until', 'registration_category', 'removal_basis',
            'tck_name', 'military_rank', 'registry_number', 'vos_code', 'vos_description',
            'deferral_type', 'deferral_until', 'police_reason', 'police_date',
            'vlk_decision', 'vlk_date', 'disability_group', 'disability_valid_until',
            'disability_reason',
            'military_registration_document', 'issued_VOD',
            'military_accounting_specialty_number', 'change_credentials',
            'reason_for_changing_credentials', 'being_on_military_registration',
        ]
        has_military = any((request.form.get(f) or '').strip() for f in military_fields_list)
        if has_military:
            data = {f: (request.form.get(f) or '').strip() or None for f in military_fields_list}
            data['issued_VOD'] = issued_VOD
            columns = ['student_id'] + list(data.keys())
            placeholders = ', '.join('?' for _ in columns)
            conn.execute(
                f"INSERT INTO military ({', '.join(columns)}) VALUES ({placeholders})",
                [student_id] + list(data.values())
            )
            conn.commit()

        log_action(
            current_username(),
            f"додав студента: {last_name_ua} {first_name_ua} {middle_name_ua} (ID {student_id}, дата нар.: {birth_date})",
            group_ids=[group_int],
            details=f"ЄДЕБО: {request.form.get('edebo_code') or 'не вказано'}, військові дані: {'так' if has_military else 'ні'}"
        )
        conn.close()
        return redirect(url_for('students.student_list'))

    conn.close()
    return render_template('add_student.html', groups=groups, licenses=licenses)


@students_bp.route('/students/<int:student_id>/photo', methods=['GET', 'POST'])
@login_required('')
def upload_photo(student_id):
    """Завантаження/заміна фото студента (POST)."""
    if request.method == 'GET':
        return redirect(url_for('students.student_details', student_id=student_id))

    conn = get_db()
    role = session.get('role')
    group_ids = session.get('group_ids', [])

    student = conn.execute("SELECT group_id FROM students WHERE id = ?", (student_id,)).fetchone()
    if not student:
        conn.close()
        flash('Студента не знайдено', 'error')
        return redirect(url_for('students.student_list'))

    if role != 'admin' and student['group_id'] not in group_ids:
        conn.close()
        flash('Ви не маєте доступу до цього студента', 'error')
        return redirect(url_for('students.student_list'))

    file = request.files.get('photo_file')
    if not file or file.filename == '':
        conn.close()
        flash('Оберіть файл фотографії', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    try:
        crop_box = (
            float(request.form['crop_x']),
            float(request.form['crop_y']),
            float(request.form['crop_w']),
            float(request.form['crop_h']),
        )
    except (KeyError, ValueError):
        conn.close()
        flash('Некоректні дані обрізки фото - спробуйте ще раз', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    from routes.photo import process_and_save_photo
    try:
        photo_rel_path = process_and_save_photo(file.read(), crop_box, student_id)
    except ValueError as e:
        conn.close()
        flash(str(e), 'error')
        return redirect(url_for('students.student_details', student_id=student_id))
    except Exception as e:
        logger.error(f"Помилка обробки фото студента (ID {student_id}): {e}", exc_info=True)
        conn.close()
        flash('Не вдалося обробити фотографію. Спробуйте інший файл.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    conn.execute("UPDATE students SET photo = ? WHERE id = ?", (photo_rel_path, student_id))
    conn.commit()
    log_action(
        current_username(),
        f"завантажив фото студента (ID {student_id})",
        group_ids=[student['group_id']]
    )
    conn.close()
    flash('Фото збережено', 'success')
    return redirect(url_for('students.student_details', student_id=student_id))


@students_bp.route('/students/<int:student_id>/photo/delete', methods=['POST'])
@login_required('')
def delete_photo(student_id):
    """Видаляє фото студента."""
    conn = get_db()
    role = session.get('role')
    group_ids = session.get('group_ids', [])

    student = conn.execute("SELECT group_id FROM students WHERE id = ?", (student_id,)).fetchone()
    if not student:
        conn.close()
        flash('Студента не знайдено', 'error')
        return redirect(url_for('students.student_list'))

    if role != 'admin' and student['group_id'] not in group_ids:
        conn.close()
        flash('Ви не маєте доступу до цього студента', 'error')
        return redirect(url_for('students.student_list'))

    from routes.photo import delete_photo as delete_photo_file
    try:
        delete_photo_file(student_id)
    except Exception as e:
        logger.error(f"Помилка видалення файлу фото студента (ID {student_id}): {e}", exc_info=True)

    conn.execute("UPDATE students SET photo = NULL WHERE id = ?", (student_id,))
    conn.commit()
    log_action(
        current_username(),
        f"видалив фото студента (ID {student_id})",
        group_ids=[student['group_id']]
    )
    conn.close()
    flash('Фото видалено', 'success')
    return redirect(url_for('students.student_details', student_id=student_id))


@students_bp.route('/students/<int:student_id>/freeze', methods=['POST'])
@permission_required('manage_students')
def freeze_student(student_id):
    """Самостійне заморожування одного студента."""
    conn = get_db()
    conn.row_factory = sqlite3.Row
    reason = (request.form.get('reason') or '').strip()

    student = conn.execute(
        "SELECT group_id, COALESCE(archived, 0) AS archived FROM students WHERE id = ?", (student_id,)
    ).fetchone()
    if not student:
        conn.close()
        flash('Студента не знайдено', 'error')
        return redirect(url_for('students.student_list'))

    if not reason:
        conn.close()
        flash('Вкажіть причину заморожування.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    if student['archived']:
        conn.close()
        flash('Студент архівний - заморожування недоступне.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    if not student['group_id']:
        conn.close()
        flash('Студент і так не прикріплений до жодної групи.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    document_files = request.files.getlist('document_files')

    cur = conn.execute(
        "INSERT INTO frozen_students (student_id, previous_group_id, order_id, reason, document_file, frozen_by) VALUES (?, ?, NULL, ?, NULL, ?)",
        (student_id, student['group_id'], reason, current_username())
    )
    frozen_id = cur.lastrowid
    saved_paths = save_multiple_attachments(conn, 'frozen_student', frozen_id, document_files, 'frozen_students', current_username())
    if saved_paths:
        conn.execute("UPDATE frozen_students SET document_file = ? WHERE id = ?", (saved_paths[0], frozen_id))
    conn.execute("UPDATE students SET group_id = NULL WHERE id = ?", (student_id,))
    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"заморозив студента (ID {student_id}) поза межами масового переведення",
        details=f"причина: {reason}"
    )
    flash('Студента заморожено. Керувати вирішенням - на сторінці «Заморожені студенти».', 'success')
    return redirect(url_for('students.student_details', student_id=student_id))


@students_bp.route('/students/<int:student_id>/edit', methods=['GET', 'POST'])
@login_required('')
def edit_student(student_id):
    """Форма редагування особистих даних студента."""
    conn = get_db()
    role = session.get('role')
    group_ids = session.get('group_ids', [])

    student = conn.execute("SELECT * FROM students WHERE id = ?", (student_id,)).fetchone()
    if not student:
        flash('Студента не знайдено', 'error')
        conn.close()
        return redirect(url_for('students.student_list'))

    if role != 'admin' and student['group_id'] not in group_ids:
        flash('Ви не маєте доступу до цього студента', 'error')
        conn.close()
        return redirect(url_for('students.student_list'))

    if role == 'admin':
        groups = conn.execute("""
            SELECT id, name, start_year, study_form, program_credits,
                   name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
            FROM groups WHERE archived = FALSE ORDER BY name, start_year
        """).fetchall()
    else:
        placeholders = ','.join('?' for _ in group_ids)
        groups = conn.execute(f"""
            SELECT id, name, start_year, study_form, program_credits,
                   name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
            FROM groups WHERE id IN ({placeholders}) ORDER BY name, start_year
        """, group_ids).fetchall()

    licenses = conn.execute(
        "SELECT id, name_ua, short_name_ua FROM institution_licenses WHERE is_active = 1 OR id = ? ORDER BY id",
        (student['license_id'],)
    ).fetchall()

    if request.method == 'POST':
        group = request.form.get('group_id')
        try:
            group_int = int(group)
        except (ValueError, TypeError):
            flash("Некоректна група", "error")
            conn.close()
            return render_template('edit_student.html', student=student, groups=groups, licenses=licenses)

        if role != 'admin' and group_int not in group_ids:
            flash("Доступ заборонено: група не належить до ваших груп", "error")
            conn.close()
            return render_template('edit_student.html', student=student, groups=groups, licenses=licenses)

        if 'update_english_names' in request.form:
            last_name_ua = request.form['last_name_UA']
            first_name_ua = request.form['first_name_UA']
            last_name_eng, first_name_eng = generate_english_name(last_name_ua, first_name_ua)
            conn.execute("""
                UPDATE students SET last_name_ENG = ?, first_name_ENG = ? WHERE id = ?
            """, (last_name_eng, first_name_eng, student_id))
            conn.commit()
            log_action(
                current_username(),
                f"оновив англійські імена: {last_name_ua} {first_name_ua} (ID {student_id})",
                group_ids=[group_int],
                details=f"→ {last_name_eng} {first_name_eng}"
            )
            flash('Англійські імена оновлено.', 'success')
            conn.close()
            return redirect(url_for('students.edit_student', student_id=student_id))
        else:
            birth_date_raw = request.form['birth_date'].strip()
            birth_date_clean = birth_date_raw.replace("-", ".")
            try:
                datetime.strptime(birth_date_clean, "%d.%m.%Y")
                birth_date = birth_date_clean
            except ValueError:
                flash("Невірний формат дати. Введіть у форматі ДД.ММ.РРРР")
                conn.close()
                return render_template('edit_student.html', student=student, groups=groups, licenses=licenses)

            old_group = student['group_id']
            license_id = request.form.get('license_id') or None
            phone = (request.form.get('phone') or '').strip() or None
            phone_backup = (request.form.get('phone_backup') or '').strip() or None
            email = (request.form.get('email') or '').strip() or None
            address_of_residence = (request.form.get('address_of_residence') or '').strip() or None
            program_credits_override_raw = request.form.get('program_credits_override', '').strip()
            if program_credits_override_raw:
                try:
                    program_credits_override = int(program_credits_override_raw)
                except ValueError:
                    flash("Кредити (скорочена програма) мають бути числом", "error")
                    conn.close()
                    return render_template('edit_student.html', student=student, groups=groups, licenses=licenses)
            else:
                program_credits_override = None
                        # Стать: або явно обрана в формі ('M'/'F'), або авто за по батькові.
            gender_form = (request.form.get('gender') or 'auto').strip()
            middle_name_form = request.form.get('middle_name_UA') or ''
            if gender_form in ('M', 'F'):
                gender = gender_form
            else:
                gender = detect_gender_from_middle_name(middle_name_form)

            conn.execute("""
                UPDATE students SET
                    last_name_UA=?, first_name_UA=?, middle_name_UA=?,
                    last_name_ENG=?, first_name_ENG=?, birth_date=?,
                    group_id=?, edebo_code=?, tax_id=?, license_id=?, program_credits_override=?,
                    phone=?, phone_backup=?, email=?, address_of_residence=?, gender=?
                WHERE id=?
            """, (
                request.form['last_name_UA'], request.form['first_name_UA'],
                request.form.get('middle_name_UA'), request.form.get('last_name_ENG'),
                request.form.get('first_name_ENG'), birth_date,
                group_int, request.form.get('edebo_code'), request.form.get('tax_id') or None,
                license_id, program_credits_override,
                phone, phone_backup, email, address_of_residence, gender, student_id
            ))
            conn.commit()

            details = f"стара група: {old_group} → нова група: {group_int}" if old_group != group_int else f"група: {group_int}"
            log_action(
                current_username(),
                f"редагував студента: {request.form['last_name_UA']} {request.form['first_name_UA']} (ID {student_id})",
                group_ids=[group_int],
                details=details
            )
        conn.close()
        return redirect(url_for('students.student_list'))

    conn.close()
    return render_template('edit_student.html', student=student, groups=groups, licenses=licenses)


@students_bp.route('/students/<int:student_id>/delete')
@permission_required('manage_students')
def delete_student(student_id):
    """Видаляє студента разом з усіма пов'язаними записами."""
    conn = get_db()
    student = conn.execute(
        "SELECT last_name_UA, first_name_UA, middle_name_UA, group_id, photo FROM students WHERE id = ?", (student_id,)
    ).fetchone()

    if not student:
        flash("Студента не знайдено")
        conn.close()
        return redirect(url_for('students.student_list'))

    try:
        conn.execute("UPDATE pending_students SET resulting_student_id = NULL WHERE resulting_student_id = ?", (student_id,))

        doc_ids = [r[0] for r in conn.execute("SELECT id FROM education_documents WHERE student_id = ?", (student_id,)).fetchall()]
        passport_ids = [r[0] for r in conn.execute("SELECT id FROM passport_documents WHERE student_id = ?", (student_id,)).fetchall()]
        all_attachments = list(get_attachments(conn, 'student', student_id))
        for doc_id in doc_ids:
            all_attachments.extend(get_attachments(conn, 'education_document', doc_id))
        for passport_id in passport_ids:
            all_attachments.extend(get_attachments(conn, 'passport_document', passport_id))
        for att in all_attachments:
            att_path = os.path.join('static', att['file_path'])
            if os.path.exists(att_path):
                os.remove(att_path)
        conn.execute("DELETE FROM attachments WHERE entity_type='student' AND entity_id=?", (student_id,))
        if doc_ids:
            placeholders = ','.join('?' for _ in doc_ids)
            conn.execute(f"DELETE FROM attachments WHERE entity_type='education_document' AND entity_id IN ({placeholders})", doc_ids)
        if passport_ids:
            placeholders = ','.join('?' for _ in passport_ids)
            conn.execute(f"DELETE FROM attachments WHERE entity_type='passport_document' AND entity_id IN ({placeholders})", passport_ids)

        for row in conn.execute("SELECT document_file FROM frozen_students WHERE student_id = ? AND document_file IS NOT NULL", (student_id,)).fetchall():
            doc_path = os.path.join('static', row[0])
            if os.path.exists(doc_path):
                os.remove(doc_path)

        if student['photo']:
            photo_path = os.path.join('static', student['photo'])
            if os.path.exists(photo_path):
                os.remove(photo_path)

        conn.execute("""
            DELETE FROM foreign_education_docs
            WHERE education_doc_id IN (SELECT id FROM education_documents WHERE student_id = ?)
        """, (student_id,))
        conn.execute("DELETE FROM education_documents WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM passport_documents WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM frozen_students WHERE student_id = ?", (student_id,))
        conn.execute(
            """
            DELETE FROM order_items
            WHERE entity_type = 'student' AND entity_id = ?
            """,
            (student_id,),
        )
        conn.execute("DELETE FROM diplomas WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM student_enrollments WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM military WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM grades WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM student_enrollments WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM activity_grades WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM students WHERE id = ?", (student_id,))
        conn.commit()

        log_action(
            current_username(),
            f"ВИДАЛИВ студента: {student['last_name_UA']} {student['first_name_UA']} {student['middle_name_UA']} (ID {student_id})",
            group_ids=[student['group_id']],
            details="каскадне видалення: military, military_history, grades, activity_grades, education_documents, passport_documents, "
                    "frozen_students, expulsion_order_students, diplomas, вкладення, фото"
        )
    except Exception as e:
        conn.rollback()
        logger.error(f"Помилка при видаленні студента ID {student_id}: {e}", exc_info=True)
        flash(f"Помилка при видаленні студента: {e}", "error")
    finally:
        conn.close()

    return redirect(url_for('students.student_list'))


@students_bp.route('/students/<int:student_id>/generate_update_link', methods=['GET', 'POST'])
@permission_required('manage_students')
def generate_update_link(student_id):
    """Створює токен-посилання для оновлення даних студентом."""
    conn = get_db()
    student = conn.execute(
        "SELECT id, last_name_UA, first_name_UA, middle_name_UA FROM students WHERE id = ?", (student_id,)
    ).fetchone()
    if not student:
        conn.close()
        flash("Студента не знайдено", "error")
        return redirect(url_for('students.student_list'))

    generated_link = None
    if request.method == 'POST':
        selected_fields = [f for f in UPDATE_REQUEST_FIELD_LABELS if request.form.get(f)]
        if not selected_fields:
            flash("Оберіть хоча б одне поле", "error")
        else:
            token = secrets.token_urlsafe(24)
            expires_at = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
            conn.execute(
                "INSERT INTO update_requests (token, student_id, allowed_fields, created_by, expires_at) VALUES (?, ?, ?, ?, ?)",
                (token, student_id, json.dumps(selected_fields, ensure_ascii=False), current_username(), expires_at)
            )
            conn.commit()
            log_action(
                current_username(),
                f"створив посилання для оновлення даних: {student['last_name_UA']} {student['first_name_UA']} (ID {student_id})",
                details=", ".join(selected_fields)
            )
            generated_link = f"{PUBLIC_APPLY_BASE_URL.rstrip('/')}/update-info/{token}"

    conn.close()
    return render_template(
        'generate_update_link.html',
        student=student, field_labels=UPDATE_REQUEST_FIELD_LABELS, generated_link=generated_link,
    )


@students_bp.route('/students/<int:student_id>/generate', methods=['GET', 'POST'])
@login_required('')
def generate(student_id):
    """Генерація .docx-документа для одного студента за обраним шаблоном.

    2-кроковий сценарій для паралельного навчання:
      1. Якщо груп декілька — вибір групи (group_id не вказано).
      2. Після вибору — форма з шаблоном; генерація з даними обраної
         групи (предмети, практики, кредити, освітня програма тощо).
    Якщо група одна — форма з шаблоном одразу.

    ВАЖЛИВО: student_dict['group_id'] перезаписується на вибрану
    групу перед викликом gen_doc() — саме звідти gen_doc бере
    subjects/practices/courseworks/attestations.
    """
    conn = get_db()

    student_basic = conn.execute(
        "SELECT id, last_name_UA, first_name_UA, middle_name_UA, group_id "
        "FROM students WHERE id = ?",
        (student_id,)
    ).fetchone()

    if not student_basic:
        conn.close()
        flash("Студента не знайдено", "error")
        return redirect(url_for('students.student_list'))

    student_groups = _groups_for_student(conn, student_id)

    if not _user_can_access_student(session, student_groups, student_basic['group_id']):
        conn.close()
        flash("Доступ заборонено: студент не належить до ваших груп", "error")
        return redirect(url_for('students.student_list'))

    # ── Визначаємо обрану групу ────────────────────────────────
    if request.method == 'POST':
        selected_group_id = request.form.get('group_id', type=int)
    else:
        selected_group_id = request.args.get('group_id', type=int)

    if not selected_group_id and len(student_groups) == 1:
        selected_group_id = student_groups[0]['id']

    if len(student_groups) > 1 and not selected_group_id:
        conn.close()
        return render_template(
            'generate_word.html',
            student=student_basic,
            student_groups=student_groups,
            selected_group=None,
            is_choice_view=True,
            available_templates=[],
        )

    selected_group = next(
        (g for g in student_groups if g['id'] == selected_group_id), None,
    )
    if not selected_group:
        conn.close()
        flash("Обрана група не належить студенту", "error")
        return redirect(url_for('students.generate', student_id=student_id))

    gid = selected_group['id']

    # ── Повні дані студента + обрана група + ліцензія ──────────
    student = conn.execute("""
        SELECT s.*,
               g.name || ' (' || g.start_year || ', ' || g.study_form || ', '
                    || g.program_credits || ' кредитів)' AS group_name,
               g.study_form, g.start_year, g.program_credits,
               g.qualification_name, g.degree_level, g.specialty,
               g.educational_program, g.knowledge_area,
               g.qualification_name_en, g.degree_level_en, g.specialty_en,
               g.isced_f_ua, g.isced_f_en,
               g.educational_program_en, g.knowledge_area_en,
               il.name_ua  AS institution_name_and_status,
               il.name_en  AS institution_name_and_status_en,
               il.short_name_ua AS license_short_name_ua,
               il.short_name_en AS license_short_name_en,
               g.entry_requirements,     g.entry_requirements_en,
               g.learning_outcomes,      g.learning_outcomes_en,
               g.program_includes,       g.program_includes_en,
               g.entry_requirements_reduced,     g.entry_requirements_reduced_en,
               g.learning_outcomes_reduced,      g.learning_outcomes_reduced_en,
               g.program_includes_reduced,       g.program_includes_reduced_en
        FROM students s
        LEFT JOIN groups g ON g.id = ?
        LEFT JOIN institution_licenses il ON s.license_id = il.id
        WHERE s.id = ?
    """, (gid, student_id)).fetchone()

    if not student:
        conn.close()
        flash("Не вдалося завантажити дані студента", "error")
        return redirect(url_for('students.student_list'))

    available_templates = get_templates_with_metadata(
        is_admin=session.get('is_admin', False),
    )

    # ── POST: генерація ────────────────────────────────────────
    if request.method == 'POST':
        template = (request.form.get('template') or '').strip()
        allowed_paths = {t['path'] for t in available_templates}
        if template not in allowed_paths:
            conn.close()
            flash("У вас немає прав для генерації документів цим шаблоном", "danger")
            return redirect(url_for(
                'students.generate', student_id=student_id, group_id=gid,
            ))

        military = conn.execute(
            "SELECT * FROM military WHERE student_id = ?", (student_id,)
        ).fetchone()

        student_dict = dict(student)
        # ── Ключовий рядок: підміняємо primary-групу на обрану ──
        student_dict['group_id'] = gid

        military_dict = dict(military) if military else {}

        display_name = (
            f"{student['last_name_UA']}_{student['first_name_UA']}.docx"
            .replace(' ', '_')
        )
        full_path = os.path.join(
            office_editor.SESSIONS_DIR, f"{uuid.uuid4().hex}.docx",
        )

        try:
            gen_doc(
                student_dict, military_dict,
                template=template, out=full_path,
                user_name=current_username(),
            )
            doc_id = office_editor.create_editing_session(
                full_path, display_name, session['user_id'],
            )
            conn.close()
            return redirect(url_for('office.edit', doc_id=doc_id))
        except Exception as e:
            logger.error(
                f"Помилка генерації документа для студента {student_id}: {e}",
                exc_info=True,
            )
            conn.close()
            flash(f"Помилка генерації: {e}", "danger")
            return redirect(url_for(
                'students.generate', student_id=student_id, group_id=gid,
            ))

    conn.close()
    return render_template(
        'generate_word.html',
        student=student,
        student_groups=student_groups,
        selected_group=selected_group,
        is_choice_view=False,
        available_templates=available_templates,
    )
    
@students_bp.route('/students/<int:student_id>/enrollments/add', methods=['POST'])
@permission_required('manage_students')
def enrollment_add(student_id):
    group_id = request.form.get('group_id', type=int)
    is_primary = request.form.get('is_primary') == '1'
    if not group_id:
        flash('Оберіть групу.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    ok, msg = add_enrollment(student_id, group_id, is_primary=is_primary)
    flash(msg, 'success' if ok else 'error')
    if ok:
        log_action(
            current_username(),
            f"додав навчання student={student_id} → group={group_id}",
            details=f"primary={is_primary}",
        )
    return redirect(url_for('students.student_details', student_id=student_id))


@students_bp.route('/students/<int:student_id>/enrollments/set_primary', methods=['POST'])
@permission_required('manage_students')
def enrollment_set_primary(student_id):
    group_id = request.form.get('group_id', type=int)
    if not group_id:
        flash('Не вказано групу.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    ok, msg = set_primary_enrollment(student_id, group_id)
    flash(msg, 'success' if ok else 'error')
    if ok:
        log_action(
            current_username(),
            f"змінив основне навчання student={student_id} → group={group_id}",
        )
    return redirect(url_for('students.student_details', student_id=student_id))


@students_bp.route('/students/<int:student_id>/enrollments/end', methods=['POST'])
@permission_required('manage_students')
def enrollment_end(student_id):
    group_id = request.form.get('group_id', type=int)
    status = request.form.get('status') or 'completed'
    if status not in ('completed', 'expelled'):
        status = 'completed'
    if not group_id:
        flash('Не вказано групу.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    ok, msg = end_enrollment(student_id, group_id, status=status)
    flash(msg, 'success' if ok else 'error')
    if ok:
        log_action(
            current_username(),
            f"завершив навчання student={student_id} group={group_id} status={status}",
        )
    return redirect(url_for('students.student_details', student_id=student_id))
