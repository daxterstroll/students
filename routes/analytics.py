# -*- coding: utf-8 -*-
"""
Сторінка аналітики (/admin/analytics): загальна статистика по студентах,
групах, заповненості даних та якості оцінок, плюс "перевірки цілісності"
(дублікати номерів диплома, розбіжність суми кредитів тощо) та
детальний перегляд по конкретному студенту.

Усі метрики рахуються прямими запитами до бази - без парсингу лог-файлу
(app.log), який для точної статистики ненадійний (ротація файлів, зміна
формату повідомлень з часом).

Доступ: лише адміністратор або користувач з явно виданим правом
'analytics' (permission_required нижче - те саме, що й у решти сторінок
адміністрування). Не-адміністратор бачить статистику лише по своїх
групах (session['group_ids']); загальносистемні розрізи (кількість
користувачів, роки вступу, перевірки цілісності, шаблони) для нього
приховані на рівні шаблону.

Визначення "заповненості" студента навмисно ті самі поля, що вже
використовуються на головній сторінці списку студентів (students.py,
student_list()) - щоб цифри тут і там завжди узгоджувались. Формула
"права на відзнаку" - та сама, що й у gen_docx.py. Формула end_year -
та сама, що й у gen_docx.py (щоб прогноз випуску збігався з тим, що
реально піде в документи).
"""
from flask import Blueprint, render_template, session, request
from routes.db import get_db
from routes.utils import permission_required

analytics_bp = Blueprint('analytics', __name__)

STUDENT_FIELDS = ['last_name_UA', 'first_name_UA', 'middle_name_UA',
                   'last_name_ENG', 'first_name_ENG', 'birth_date', 'group_id', 'edebo_code']
MILITARY_FIELDS = [
    'registration_number_of_the_DRPVR', 'military_registration_document', 'issued_VOD',
    'military_accounting_specialty_number', 'military_rank', 'change_credentials',
    'reason_for_changing_credentials', 'being_on_military_registration', 'address_of_residence'
]


def _grade_letter(grade):
    """Літера за шкалою ECTS з числової оцінки 0-100 (та сама шкала, що й gen_docx.format_grade)."""
    try:
        g = int(grade)
    except (TypeError, ValueError):
        return None
    if g < 60:
        return 'Незадовільно'
    if g <= 63:
        return 'E'
    if g <= 73:
        return 'D'
    if g <= 81:
        return 'C'
    if g <= 89:
        return 'B'
    return 'A'


def _is_passing(grade):
    letter = _grade_letter(grade)
    return letter is not None and letter != 'Незадовільно'


def _compute_end_year(start_year, degree_level, credits):
    """Рік випуску - та сама формула, що й у gen_docx.py (щоб прогноз збігався з документами)."""
    try:
        if not (credits and start_year):
            return None
        credits = int(credits)
        year = int(start_year)
        if degree_level == 'Бакалавр':
            if credits == 240:
                return year + 4
            if credits == 180:
                return year + 3
        elif degree_level == 'Магістр':
            if credits in (90, 120):
                return year + 2
        return year + (credits // 60)
    except (ValueError, TypeError):
        return None


@analytics_bp.route('/admin/analytics')
@permission_required('analytics')
def dashboard():
    conn = get_db()

    include_archived = request.args.get('include_archived') == '1'
    selected_student_id = request.args.get('student_id', type=int)

    role = session.get('role')
    is_admin = (role == 'admin')
    user_group_ids = session.get('group_ids', []) or []
    scope_clause = ""
    scope_params = []
    if not is_admin and user_group_ids:
        placeholders = ','.join('?' for _ in user_group_ids)
        scope_clause = f" AND s.group_id IN ({placeholders})"
        scope_params = list(user_group_ids)

    student_archived_cond = "" if include_archived else "AND COALESCE(s.archived,0)=0"
    group_archived_cond = "" if include_archived else "AND COALESCE(g.archived,0)=0"

    if not is_admin and user_group_ids:
        placeholders = ','.join('?' for _ in user_group_ids)
        group_where = f"WHERE g.id IN ({placeholders}) {group_archived_cond}"
        group_where_params = list(user_group_ids)
    else:
        group_where = f"WHERE 1=1 {group_archived_cond}"
        group_where_params = []

    # ================= KPI-картки =================
    total_students = conn.execute(
        f"SELECT COUNT(*) FROM students s WHERE 1=1 {student_archived_cond}{scope_clause}", scope_params
    ).fetchone()[0]

    if is_admin:
        active_groups = conn.execute(
            "SELECT COUNT(*) FROM groups" + ("" if include_archived else " WHERE COALESCE(archived,0)=0")
        ).fetchone()[0]
        archived_groups = conn.execute("SELECT COUNT(*) FROM groups WHERE COALESCE(archived,0)=1").fetchone()[0]
        total_users = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    else:
        active_groups = len(user_group_ids)
        archived_groups = 0
        total_users = None

    avg_group_size = round(total_students / active_groups, 1) if active_groups else 0

    # ================= Групи (повна інформація - потрібна для кількох перевірок нижче) =================
    groups_full = [dict(r) for r in conn.execute(f"""
        SELECT g.* FROM groups g {group_where}
    """, group_where_params).fetchall()]
    groups_by_id = {g['id']: g for g in groups_full}

    # ================= Студенти по групах / формах / ступенях =================
    students_per_group = conn.execute(f"""
        SELECT g.id, g.name, g.start_year, g.study_form, COALESCE(g.archived,0) AS archived,
               COUNT(s.id) FILTER (WHERE 1=1 {student_archived_cond}) AS cnt
        FROM groups g LEFT JOIN students s ON s.group_id = g.id
        {group_where}
        GROUP BY g.id ORDER BY cnt DESC
    """, group_where_params).fetchall()

    groups_by_year = conn.execute(f"""
        SELECT g.start_year, COUNT(*) AS cnt FROM groups g
        {group_where}
        GROUP BY g.start_year ORDER BY g.start_year
    """, group_where_params).fetchall()

    students_by_form = conn.execute(f"""
        SELECT g.study_form, COUNT(s.id) AS cnt
        FROM students s JOIN groups g ON g.id = s.group_id
        WHERE 1=1 {student_archived_cond} {group_archived_cond} {scope_clause}
        GROUP BY g.study_form
    """, scope_params).fetchall()

    students_by_degree = conn.execute(f"""
        SELECT g.degree_level, COUNT(s.id) AS cnt
        FROM students s JOIN groups g ON g.id = s.group_id
        WHERE 1=1 {student_archived_cond} {group_archived_cond} {scope_clause}
        GROUP BY g.degree_level
    """, scope_params).fetchall()

    # ================= Прогноз випуску по роках =================
    graduation_forecast = {}
    for row in students_per_group:
        g = groups_by_id.get(row['id'])
        if not g or not row['cnt']:
            continue
        end_year = _compute_end_year(g['start_year'], g['degree_level'], g['program_credits'])
        if end_year:
            graduation_forecast[end_year] = graduation_forecast.get(end_year, 0) + row['cnt']
    graduation_forecast = sorted(graduation_forecast.items())

    # ================= Дані для заповненості/оцінок (у Python, простіше й надійніше за велику SQL) =================
    students_raw = conn.execute(f"""
        SELECT s.* FROM students s
        WHERE 1=1 {student_archived_cond}{scope_clause}
    """, scope_params).fetchall()
    students_raw = [dict(r) for r in students_raw]
    student_ids = [s['id'] for s in students_raw]

    # Список студентів для випадаючого списку "Якість оцінок конкретного студента"
    student_options = conn.execute(f"""
        SELECT s.id, s.last_name_UA || ' ' || s.first_name_UA || ' ' || COALESCE(s.middle_name_UA,'') AS full_name,
               g.name AS group_name
        FROM students s JOIN groups g ON g.id = s.group_id
        WHERE 1=1 {student_archived_cond} {group_archived_cond} {scope_clause}
        ORDER BY g.name, s.last_name_UA
    """, scope_params).fetchall()

    military_by_student = {}
    if student_ids:
        placeholders = ','.join('?' for _ in student_ids)
        for row in conn.execute(f"SELECT * FROM military WHERE student_id IN ({placeholders})", student_ids):
            military_by_student.setdefault(row['student_id'], dict(row))

    grades_by_student = {}
    all_grade_rows = []
    if student_ids:
        placeholders = ','.join('?' for _ in student_ids)
        for row in conn.execute(f"""
            SELECT gr.student_id, gr.grade, sub.id AS subject_id, sub.type AS subject_type,
                   sub.name AS subject_name, sub.code AS subject_code, sub.credits AS subject_credits,
                   sub.group_id AS group_id
            FROM grades gr JOIN subjects sub ON sub.id = gr.subject_id
            WHERE gr.student_id IN ({placeholders})
        """, student_ids):
            row = dict(row)
            grades_by_student.setdefault(row['student_id'], []).append(row)
            all_grade_rows.append(row)

    activities_by_student = {}
    attestation_by_student = {}
    activity_rows_by_student = {}
    if student_ids:
        placeholders = ','.join('?' for _ in student_ids)
        # activity_grades.name НЕ є назвою практики/курсової/атестації - це
        # окреме поле для теми кваліфікаційної роботи (лише для атестації).
        # Реальна назва зберігається в самій сутності (practices/courseworks/
        # attestations) і приєднується через entity_id - тому окремий JOIN
        # для кожного типу, як і в students.py.
        queries = {
            'practice': f"""
                SELECT ag.student_id, 'practice' AS entity_type, ag.grade, p.code, p.name
                FROM activity_grades ag JOIN practices p ON p.id = ag.entity_id
                WHERE ag.entity_type='practice' AND ag.student_id IN ({placeholders})
            """,
            'coursework': f"""
                SELECT ag.student_id, 'coursework' AS entity_type, ag.grade, c.code, c.name
                FROM activity_grades ag JOIN courseworks c ON c.id = ag.entity_id
                WHERE ag.entity_type='coursework' AND ag.student_id IN ({placeholders})
            """,
            'attestation': f"""
                SELECT ag.student_id, 'attestation' AS entity_type, ag.grade, a.code, a.name, ag.name AS thesis_title
                FROM activity_grades ag JOIN attestations a ON a.id = ag.entity_id
                WHERE ag.entity_type='attestation' AND ag.student_id IN ({placeholders})
            """,
        }
        for sql in queries.values():
            for row in conn.execute(sql, student_ids):
                row = dict(row)
                activity_rows_by_student.setdefault(row['student_id'], []).append(row)
                if row['grade'] is not None and str(row['grade']).strip():
                    activities_by_student.setdefault(row['student_id'], []).append(row['grade'])
                if row['entity_type'] == 'attestation':
                    attestation_by_student.setdefault(row['student_id'], []).append(row['grade'])

    subjects_total_by_group = {}
    subjects_credits_by_group = {}
    subjects_by_group = {}
    for row in conn.execute("SELECT id, name, code, credits, group_id FROM subjects"):
        row = dict(row)
        subjects_total_by_group[row['group_id']] = subjects_total_by_group.get(row['group_id'], 0) + 1
        subjects_credits_by_group[row['group_id']] = subjects_credits_by_group.get(row['group_id'], 0) + (row['credits'] or 0)
        subjects_by_group.setdefault(row['group_id'], []).append(row)

    # Немає єдиної таблиці "активностей" - практики/курсові/атестації
    # зберігаються в окремих таблицях, тому рахуємо суму по кожній групі
    # (і кількість, і кредити - кредити потрібні для перевірки цілісності програми).
    activities_total_by_group = {}
    activities_credits_by_group = {}
    for tbl in ('practices', 'courseworks', 'attestations'):
        for row in conn.execute(f"SELECT group_id, credits FROM {tbl}"):
            activities_total_by_group[row['group_id']] = activities_total_by_group.get(row['group_id'], 0) + 1
            activities_credits_by_group[row['group_id']] = activities_credits_by_group.get(row['group_id'], 0) + (row['credits'] or 0)

    edu_doc_students = set()
    if student_ids:
        placeholders = ','.join('?' for _ in student_ids)
        for row in conn.execute(f"SELECT DISTINCT student_id FROM education_documents WHERE student_id IN ({placeholders})", student_ids):
            edu_doc_students.add(row['student_id'])

    diploma_by_student = {}
    if student_ids:
        placeholders = ','.join('?' for _ in student_ids)
        for row in conn.execute(f"SELECT student_id, diploma_number, appendix_number FROM diplomas WHERE student_id IN ({placeholders})", student_ids):
            diploma_by_student[row['student_id']] = dict(row)

    # ================= Обчислення заповненості по кожному студенту (+ по групах окремо) =================
    personal_pcts, military_pcts, grades_pcts, activities_pcts = [], [], [], []
    has_military_cnt = has_edu_doc_cnt = has_diploma_num_cnt = has_appendix_num_cnt = 0
    readiness_scores = []
    all_letters = []
    honor_eligible = 0
    at_risk_students = []  # один запис на студента: {'student','group','count','worst_letter','details'}

    per_group_stats = {}  # group_id -> {'personal': [...], 'grades': [...], ...}

    for st in students_raw:
        gid = st['group_id']
        per_group_stats.setdefault(gid, {'personal': [], 'grades': [], 'name': groups_by_id.get(gid, {}).get('name', '?')})

        personal_filled = sum(1 for f in STUDENT_FIELDS if st.get(f) and str(st.get(f)).strip())
        personal_pct = personal_filled / len(STUDENT_FIELDS) * 100
        personal_pcts.append(personal_pct)
        per_group_stats[gid]['personal'].append(personal_pct)

        mil = military_by_student.get(st['id'])
        if mil:
            has_military_cnt += 1
            mil_filled = sum(1 for f in MILITARY_FIELDS if mil.get(f) and str(mil.get(f)).strip())
            military_pcts.append(mil_filled / len(MILITARY_FIELDS) * 100)

        group_subj_total = subjects_total_by_group.get(gid, 0)
        st_grades = grades_by_student.get(st['id'], [])
        grades_pct = (len(st_grades) / group_subj_total * 100) if group_subj_total else None
        if grades_pct is not None:
            grades_pcts.append(grades_pct)
            per_group_stats[gid]['grades'].append(grades_pct)

        group_act_total = activities_total_by_group.get(gid, 0)
        st_acts = activities_by_student.get(st['id'], [])
        act_pct = (len(st_acts) / group_act_total * 100) if group_act_total else None
        if act_pct is not None:
            activities_pcts.append(act_pct)

        has_edu = st['id'] in edu_doc_students
        if has_edu:
            has_edu_doc_cnt += 1
        dip = diploma_by_student.get(st['id'], {})
        if dip.get('diploma_number'):
            has_diploma_num_cnt += 1
        if dip.get('appendix_number'):
            has_appendix_num_cnt += 1

        dims = [personal_pct]
        if grades_pct is not None:
            dims.append(grades_pct)
        dims.append(100 if has_edu else 0)
        dims.append(100 if dip.get('diploma_number') else 0)
        readiness_scores.append(sum(dims) / len(dims))

        letters = [(_grade_letter(g['grade']), g) for g in st_grades]
        risky_items = []
        for letter, g in letters:
            if letter:
                all_letters.append(letter)
            if letter == 'Незадовільно':
                risky_items.append((f"{g['subject_code']} {g['subject_name']}", g['grade']))
        for a in activity_rows_by_student.get(st['id'], []):
            letter = _grade_letter(a['grade'])
            if letter == 'Незадовільно':
                risky_items.append((f"{a['code']} {a['name']}" if a['name'] else a['entity_type'], a['grade']))

        if risky_items:
            at_risk_students.append({
                'student': f"{st['last_name_UA']} {st['first_name_UA']}",
                'group': groups_by_id.get(gid, {}).get('name', '?'),
                'count': len(risky_items),
                'worst_letter': 'Незадовільно',
                'details': risky_items,
            })

        letters_only = [l for l, _ in letters if l]
        attestation_letters = [_grade_letter(g) for g in attestation_by_student.get(st['id'], [])]
        attestation_letters = [l for l in attestation_letters if l]
        if letters_only:
            a_share = letters_only.count('A') / len(letters_only)
            no_fail = not any(l in ('D', 'E', 'Незадовільно') for l in letters_only)
            attestation_ok = all(l == 'A' for l in attestation_letters) if attestation_letters else True
            if a_share >= 0.75 and no_fail and attestation_ok:
                honor_eligible += 1

    at_risk_students.sort(key=lambda r: -r['count'])

    def _avg(lst):
        return round(sum(lst) / len(lst), 1) if lst else None

    n = len(students_raw) or 1
    completeness = {
        'personal_avg': _avg(personal_pcts),
        'military_avg': _avg(military_pcts),
        'military_have_pct': round(has_military_cnt / n * 100, 1),
        'grades_avg': _avg(grades_pcts),
        'activities_avg': _avg(activities_pcts),
        'edu_doc_pct': round(has_edu_doc_cnt / n * 100, 1),
        'diploma_num_pct': round(has_diploma_num_cnt / n * 100, 1),
        'appendix_num_pct': round(has_appendix_num_cnt / n * 100, 1),
    }

    # Заповненість по кожній групі окремо (а не лише загальний середній)
    per_group_completeness = []
    for gid, data in per_group_stats.items():
        per_group_completeness.append({
            'group_name': data['name'],
            'personal_avg': _avg(data['personal']) or 0,
            'grades_avg': _avg(data['grades']) or 0,
            'students_count': len(data['personal']),
        })
    per_group_completeness.sort(key=lambda x: x['personal_avg'])

    readiness_buckets = {'Повністю готові (90-100%)': 0, 'Майже готові (70-89%)': 0,
                          'Частково (40-69%)': 0, 'Тільки почали (<40%)': 0}
    for score in readiness_scores:
        if score >= 90:
            readiness_buckets['Повністю готові (90-100%)'] += 1
        elif score >= 70:
            readiness_buckets['Майже готові (70-89%)'] += 1
        elif score >= 40:
            readiness_buckets['Частково (40-69%)'] += 1
        else:
            readiness_buckets['Тільки почали (<40%)'] += 1

    grade_distribution = {}
    for letter in ['A', 'B', 'C', 'D', 'E', 'Незадовільно']:
        grade_distribution[letter] = all_letters.count(letter)
    total_grades = len(all_letters) or 1

    # ================= Рейтинг предметів за середнім балом =================
    subject_grades_map = {}
    for row in all_grade_rows:
        try:
            val = int(row['grade'])
        except (TypeError, ValueError):
            continue
        key = row['subject_id']
        subject_grades_map.setdefault(key, {'name': f"{row['subject_code']} {row['subject_name']}",
                                             'group': groups_by_id.get(row['group_id'], {}).get('name', '?'),
                                             'values': []})
        subject_grades_map[key]['values'].append(val)

    subject_ranking = []
    for data in subject_grades_map.values():
        if data['values']:
            subject_ranking.append({
                'name': data['name'], 'group': data['group'],
                'avg': round(sum(data['values']) / len(data['values']), 1),
                'count': len(data['values']),
            })
    subject_ranking.sort(key=lambda x: x['avg'])
    worst_subjects = subject_ranking[:5]
    best_subjects = list(reversed(subject_ranking[-5:])) if len(subject_ranking) > 5 else list(reversed(subject_ranking))

    # ================= Перевірки цілісності (тільки адмін - стосуються всієї системи) =================
    credit_mismatches = []
    missing_accreditation = []
    if is_admin:
        for g in groups_full:
            gid = g['id']
            total_credits = (subjects_credits_by_group.get(gid, 0) + activities_credits_by_group.get(gid, 0))
            if total_credits and g['program_credits'] and total_credits != g['program_credits']:
                credit_mismatches.append({
                    'group': g['name'], 'expected': g['program_credits'], 'actual': total_credits,
                })

        accreditation_pairs = {(row['degree'], row['specialty']) for row in
                                conn.execute("SELECT degree, specialty FROM accreditations")}
        for g in groups_full:
            if (g['degree_level'], g['specialty']) not in accreditation_pairs:
                missing_accreditation.append(g['name'])

    duplicate_diploma_numbers, duplicate_appendix_numbers = [], []
    if is_admin:
        for row in conn.execute("""
            SELECT diploma_number, COUNT(*) AS cnt FROM diplomas
            WHERE diploma_number IS NOT NULL AND diploma_number != ''
            GROUP BY diploma_number HAVING cnt > 1
        """):
            duplicate_diploma_numbers.append({'number': row['diploma_number'], 'count': row['cnt']})
        for row in conn.execute("""
            SELECT appendix_number, COUNT(*) AS cnt FROM diplomas
            WHERE appendix_number IS NOT NULL AND appendix_number != ''
            GROUP BY appendix_number HAVING cnt > 1
        """):
            duplicate_appendix_numbers.append({'number': row['appendix_number'], 'count': row['cnt']})

    # ================= Якість оцінок конкретного студента =================
    student_detail = None
    accessible_ids = {row['id'] for row in student_options}
    if selected_student_id and selected_student_id in accessible_ids:
        s_row = next((s for s in students_raw if s['id'] == selected_student_id), None)
        gid = s_row['group_id'] if s_row else None

        s_grades = grades_by_student.get(selected_student_id, [])
        s_letters = [(g['subject_code'], g['subject_name'], g['grade'], _grade_letter(g['grade']), g['subject_credits'])
                     for g in s_grades]

        s_activities = activity_rows_by_student.get(selected_student_id, [])
        s_activity_rows = [(a['entity_type'], f"{a['code']} {a['name']}" if a['name'] else '', a['grade'],
                             _grade_letter(a['grade']), a.get('thesis_title')) for a in s_activities]

        letters_only = [t[3] for t in s_letters if t[3]]
        attestation_letters = [_grade_letter(g) for g in attestation_by_student.get(selected_student_id, [])]
        attestation_letters = [l for l in attestation_letters if l]

        a_share = (letters_only.count('A') / len(letters_only)) if letters_only else 0
        no_fail = not any(l in ('D', 'E', 'Незадовільно') for l in letters_only)
        attestation_ok = all(l == 'A' for l in attestation_letters) if attestation_letters else True
        is_honor_eligible = bool(letters_only) and a_share >= 0.75 and no_fail and attestation_ok

        # Список конкретних НЕзаповнених предметів (не лише кількість)
        taken_subject_ids = {g['subject_id'] for g in s_grades}
        missing_subjects = [f"{s['code']} {s['name']}" for s in subjects_by_group.get(gid, [])
                             if s['id'] not in taken_subject_ids]

        # Порівняння із середнім по групі
        numeric_grades = [int(g['grade']) for g in s_grades if str(g['grade']).isdigit()]
        student_avg = round(sum(numeric_grades) / len(numeric_grades), 1) if numeric_grades else None
        group_numeric_grades = [int(r['grade']) for r in all_grade_rows
                                 if r['group_id'] == gid and str(r['grade']).isdigit()]
        group_avg = round(sum(group_numeric_grades) / len(group_numeric_grades), 1) if group_numeric_grades else None

        # Зароблені кредити (тільки предмети зі позитивною оцінкою) vs кредити програми групи
        earned_credits = sum(g['subject_credits'] or 0 for g in s_grades if _is_passing(g['grade']))
        required_credits = groups_by_id.get(gid, {}).get('program_credits')

        student_row = next((s for s in student_options if s['id'] == selected_student_id), None)
        student_detail = {
            'id': selected_student_id,
            'full_name': student_row['full_name'] if student_row else '',
            'group_name': student_row['group_name'] if student_row else '',
            'subjects': s_letters,
            'activities': s_activity_rows,
            'a_share': round(a_share * 100, 1),
            'no_fail': no_fail,
            'attestation_ok': attestation_ok,
            'is_honor_eligible': is_honor_eligible,
            'grades_filled': len(s_grades),
            'grades_total': subjects_total_by_group.get(gid, 0),
            'missing_subjects': missing_subjects,
            'student_avg': student_avg,
            'group_avg': group_avg,
            'earned_credits': earned_credits,
            'required_credits': required_credits,
        }

    template_selected_student_id = selected_student_id if selected_student_id in accessible_ids else None

    # ================= Шаблони документів =================
    templates_total = conn.execute("SELECT COUNT(*) FROM document_templates").fetchone()[0]
    templates_hidden = conn.execute("SELECT COUNT(*) FROM document_templates WHERE hidden=1").fetchone()[0]
    templates_admin_only = conn.execute("SELECT COUNT(*) FROM document_templates WHERE admin_only=1").fetchone()[0]

    conn.close()

    return render_template(
        'analytics.html',
        total_students=total_students,
        active_groups=active_groups,
        archived_groups=archived_groups,
        total_users=total_users,
        avg_group_size=avg_group_size,
        students_per_group=students_per_group,
        groups_by_year=groups_by_year,
        students_by_form=students_by_form,
        students_by_degree=students_by_degree,
        graduation_forecast=graduation_forecast,
        completeness=completeness,
        per_group_completeness=per_group_completeness,
        readiness_buckets=readiness_buckets,
        total_ready_students=len(students_raw),
        grade_distribution=grade_distribution,
        total_grades=total_grades,
        honor_eligible=honor_eligible,
        worst_subjects=worst_subjects,
        best_subjects=best_subjects,
        at_risk_students=at_risk_students,
        credit_mismatches=credit_mismatches,
        missing_accreditation=missing_accreditation,
        duplicate_diploma_numbers=duplicate_diploma_numbers,
        duplicate_appendix_numbers=duplicate_appendix_numbers,
        templates_total=templates_total,
        templates_hidden=templates_hidden,
        templates_admin_only=templates_admin_only,
        is_admin=is_admin,
        include_archived=include_archived,
        student_options=student_options,
        selected_student_id=template_selected_student_id,
        student_detail=student_detail,
    )