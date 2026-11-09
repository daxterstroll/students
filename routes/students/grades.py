"""
routes/students/grades.py
==========================
Виставлення оцінок студенту: звичайні предмети (edit_grades) і
активності - практики/курсові/атестації (edit_activities_grades), з
урахуванням скороченої програми.

Паралельне навчання (student_enrollments): edit_activities_grades
працює з ОДНІЄЮ обраною групою за раз (2-кроковий UI - спочатку
вибір групи, потім редагування). edit_grades для сумісності
залишається на student.group_id (primary), бо для предметів
зазвичай використовується окремий контекст через сторінку групи.
"""
from flask import render_template, request, redirect, url_for, session, flash
from routes.db import get_db
from routes.utils import log_action, login_required, permission_required, logger
from routes.helpers import current_username
from routes.students import students_bp
import sqlite3
from routes.utils import is_student_on_reduced_program, program_track_condition, apply_track_overrides
from routes.enrollments_service import get_active_group_ids_for_student, ensure_enrollments_table
from datetime import datetime


def _groups_for_student(conn, student_id, include_inactive=False):
    """Повертає список груп студента (dict).

    include_inactive=False (за замовчуванням) — лише активні enrollments.
        Використовується на сторінках редагування оцінок / активностей
        (редагувати оцінки в закритій групі не має сенсу).

    include_inactive=True — активні + історичні (graduated / expelled /
        frozen) + legacy fallback на students.group_id.
        Використовується для генерації документів: випускник або
        відрахований студент повинен мати можливість отримати додаток
        до диплому / інший документ для СВОЄЇ колишньої групи.

    Кожен елемент результату:
        {'id', 'name', 'start_year', 'study_form', 'program_credits',
         'status', 'is_primary'}
    """
    if not include_inactive:
        group_ids = get_active_group_ids_for_student(student_id, conn=conn)
        result = []
        for gid in group_ids:
            row = conn.execute(
                "SELECT id, name, start_year, study_form, program_credits "
                "FROM groups WHERE id = ?",
                (gid,),
            ).fetchone()
            if row:
                d = dict(row)
                d['status'] = 'active'
                d['is_primary'] = True
                result.append(d)
        return result

    # ── include_inactive=True: активні + історичні ──
    ensure_enrollments_table(conn)
    rows = conn.execute("""
        SELECT e.group_id, e.status, e.is_primary,
               g.id, g.name, g.start_year, g.study_form, g.program_credits
        FROM student_enrollments e
        JOIN groups g ON g.id = e.group_id
        WHERE e.student_id = ?
        ORDER BY
            CASE e.status WHEN 'active' THEN 0 ELSE 1 END,
            e.is_primary DESC,
            e.id DESC
    """, (student_id,)).fetchall()

    result = []
    seen = set()
    for r in rows:
        gid = r['group_id']
        if gid in seen:
            continue
        seen.add(gid)
        result.append({
            'id': r['id'],
            'name': r['name'],
            'start_year': r['start_year'],
            'study_form': r['study_form'],
            'program_credits': r['program_credits'],
            'status': r['status'],
            'is_primary': bool(r['is_primary']),
        })

    # Legacy-fallback: студент без жодного enrollment'а, але з group_id
    if not result:
        row = conn.execute("""
            SELECT s.group_id AS gid,
                   g.id, g.name, g.start_year, g.study_form, g.program_credits
            FROM students s
            JOIN groups g ON g.id = s.group_id
            WHERE s.id = ?
        """, (student_id,)).fetchone()
        if row:
            result.append({
                'id': row['id'],
                'name': row['name'],
                'start_year': row['start_year'],
                'study_form': row['study_form'],
                'program_credits': row['program_credits'],
                'status': 'archived',
                'is_primary': True,
            })

    return result


def _user_can_access_student(session_obj, student_groups, primary_group_id):
    """True, якщо користувач має доступ хоча б до однієї з груп студента
    (або є адміном)."""
    if session_obj.get('role') == 'admin' or session_obj.get('is_admin'):
        return True
    user_group_ids = set(session_obj.get('group_ids', []) or [])
    if primary_group_id and primary_group_id in user_group_ids:
        return True
    for g in student_groups:
        if g['id'] in user_group_ids:
            return True
    return False


# ============================================================
# Редагування оцінок з активностей (практики / курсові / атестації)
# ============================================================

@students_bp.route('/activities_grades/<int:student_id>', methods=['GET', 'POST'])
@permission_required('manage_students')
def edit_activities_grades(student_id):
    """Виставлення оцінок студента з практик/курсових/атестацій.

    2-кроковий сценарій для паралельного навчання:
      1. Якщо груп декілька - показати вибір групи (group_id не вказано).
      2. Після вибору - редагування оцінок обраної групи.
    Якщо група одна - одразу відкривається редагування.
    """
    conn = get_db()

    student = conn.execute("""
        SELECT s.*,
               g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS group_name,
               g.study_form, g.program_credits, g.qualification_name, g.degree_level, g.specialty,
               g.educational_program, g.knowledge_area, g.qualification_name_en, g.degree_level_en,
               g.specialty_en, g.educational_program_en, g.knowledge_area_en
        FROM students s
        LEFT JOIN groups g ON s.group_id = g.id
        WHERE s.id = ?
    """, (student_id,)).fetchone()

    if not student:
        conn.close()
        flash("Студента не знайдено", "error")
        return redirect(url_for('students.student_list'))

    student_groups = _groups_for_student(conn, student_id)

    if not _user_can_access_student(session, student_groups, student['group_id']):
        conn.close()
        flash("Доступ заборонено: студент не належить до ваших груп", "error")
        return redirect(url_for('students.student_list'))

    # ── Визначаємо обрану групу ────────────────────────────────
    if request.method == 'POST':
        selected_group_id = request.form.get('group_id', type=int)
    else:
        selected_group_id = request.args.get('group_id', type=int)

    # Якщо група одна - обираємо автоматично
    if not selected_group_id and len(student_groups) == 1:
        selected_group_id = student_groups[0]['id']

    # Якщо груп декілька і жодну не обрано - крок вибору групи
    if len(student_groups) > 1 and not selected_group_id:
        conn.close()
        return render_template(
            'edit_activities_grades.html',
            student=student,
            student_groups=student_groups,
            selected_group=None,
            is_choice_view=True,
            practices=[],
            courseworks=[],
            attestations=[],
            grade_map={},
            is_reduced=False,
        )

    # Валідація: обрана група має належати студенту
    selected_group = next(
        (g for g in student_groups if g['id'] == selected_group_id), None,
    )
    if not selected_group:
        conn.close()
        flash("Обрана група не належить студенту", "error")
        return redirect(url_for('students.edit_activities_grades', student_id=student_id))

    gid = selected_group['id']
    is_reduced = is_student_on_reduced_program(conn, student_id, gid)
    track_cond = program_track_condition(is_reduced)

    # ── Завантаження активностей для обраної групи ────────────
    practices = [dict(r) for r in conn.execute(f"""
        SELECT id, code, name, credits, reduced_credits, type, reduced_type, position
        FROM practices WHERE group_id = ?{track_cond} ORDER BY position
    """, (gid,)).fetchall()]
    for p in practices:
        apply_track_overrides(p, is_reduced)

    courseworks = [dict(r) for r in conn.execute(f"""
        SELECT id, code, name, credits, reduced_credits, type, reduced_type, position
        FROM courseworks WHERE group_id = ?{track_cond} ORDER BY position
    """, (gid,)).fetchall()]
    for c in courseworks:
        apply_track_overrides(c, is_reduced)

    attestations = [dict(r) for r in conn.execute(f"""
        SELECT a.id, a.code, a.name, a.credits, a.reduced_credits, a.type, a.reduced_type,
               a.position, ag.name AS student_name
        FROM attestations a
        LEFT JOIN activity_grades ag
          ON ag.entity_id = a.id AND ag.entity_type = 'attestation' AND ag.student_id = ?
        WHERE a.group_id = ?{track_cond} ORDER BY position
    """, (student_id, gid)).fetchall()]
    for a in attestations:
        apply_track_overrides(a, is_reduced)

    existing_grades = conn.execute("""
        SELECT id, entity_id, entity_type, grade, name
        FROM activity_grades WHERE student_id = ?
    """, (student_id,)).fetchall()
    grade_map = {
        (g['entity_id'], g['entity_type']):
            {'id': g['id'], 'grade': g['grade'], 'name': g['name']}
        for g in existing_grades
    }

    # ── Збереження ────────────────────────────────────────────
    if request.method == 'POST':
        try:
            for entity_type, entities in [
                ('practice', practices),
                ('coursework', courseworks),
                ('attestation', attestations),
            ]:
                for entity in entities:
                    grade_key = f'grade_{entity_type}_{entity["id"]}'
                    name_key = f'name_{entity_type}_{entity["id"]}' if entity_type == 'attestation' else None
                    grade_value = request.form.get(grade_key)
                    student_name = request.form.get(name_key) if entity_type == 'attestation' else ''
                    key = (entity['id'], entity_type)

                    if grade_value:
                        try:
                            grade_value = int(grade_value)
                            if not 0 <= grade_value <= 100:
                                flash(f"Некоректна оцінка для {entity['name']}: має бути від 0 до 100", "error")
                                continue
                            if key in grade_map:
                                conn.execute("""
                                    UPDATE activity_grades SET grade = ?, name = ?
                                    WHERE id = ? AND student_id = ? AND entity_id = ? AND entity_type = ?
                                """, (grade_value, student_name, grade_map[key]['id'],
                                      student_id, entity['id'], entity_type))
                            else:
                                conn.execute("""
                                    INSERT INTO activity_grades (student_id, entity_id, entity_type, grade, name)
                                    VALUES (?, ?, ?, ?, ?)
                                """, (student_id, entity['id'], entity_type, grade_value, student_name))
                        except ValueError:
                            conn.execute("""
                                DELETE FROM activity_grades
                                WHERE student_id = ? AND entity_id = ? AND entity_type = ?
                            """, (student_id, entity['id'], entity_type))
                            flash(f"Некоректна оцінка для {entity['name']}: має бути числом", "error")
                    else:
                        if key in grade_map:
                            conn.execute("""
                                DELETE FROM activity_grades
                                WHERE id = ? AND student_id = ? AND entity_id = ? AND entity_type = ?
                            """, (grade_map[key]['id'], student_id, entity['id'], entity_type))
                        else:
                            conn.execute("""
                                DELETE FROM activity_grades
                                WHERE student_id = ? AND entity_id = ? AND entity_type = ?
                            """, (student_id, entity['id'], entity_type))

            conn.commit()
            flash("Оцінки успішно збережено", "success")
            log_action(
                current_username(),
                f"змінив активності: {student['last_name_UA']} {student['first_name_UA']} (ID {student_id})",
                group_ids=[gid],
                details=f"практики: {len(practices)}, курсові: {len(courseworks)}, "
                        f"атестації: {len(attestations)}, група: {selected_group['name']}",
            )
            conn.close()
            # Повертаємось на ту ж групу (щоб можна було продовжити),
            # або на список, якщо група одна
            if len(student_groups) > 1:
                return redirect(url_for('students.edit_activities_grades',
                                        student_id=student_id, group_id=gid))
            return redirect(url_for('students.student_list'))
        except Exception as e:
            conn.rollback()
            logger.error(f"Помилка при збереженні оцінок з активностей (student_id={student_id}): {e}",
                         exc_info=True)
            flash(f"Помилка при збереженні оцінок: {str(e)}", "error")

    conn.close()
    return render_template(
        'edit_activities_grades.html',
        student=student,
        student_groups=student_groups,
        selected_group=selected_group,
        is_choice_view=False,
        practices=practices,
        courseworks=courseworks,
        attestations=attestations,
        grade_map=grade_map,
        is_reduced=is_reduced,
    )


# ============================================================
# Редагування оцінок з предметів (2-кроковий вибір групи)
# ============================================================

@students_bp.route('/grades/<int:student_id>', methods=['GET', 'POST'])
@permission_required('manage_students')
def edit_grades(student_id):
    """Виставлення оцінок студента з предметів.

    2-кроковий сценарій для паралельного навчання:
      1. Якщо груп декілька - показати вибір групи (group_id не вказано).
      2. Після вибору - редагування оцінок обраної групи.
    Якщо група одна - одразу відкривається редагування.
    """
    conn = get_db()

    student = conn.execute("""
        SELECT s.*,
               g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS group_name
        FROM students s
        LEFT JOIN groups g ON s.group_id = g.id
        WHERE s.id = ?
    """, (student_id,)).fetchone()

    if not student:
        conn.close()
        flash("Студента не знайдено", "error")
        return redirect(url_for('students.student_list'))

    student_groups = _groups_for_student(conn, student_id)

    if not _user_can_access_student(session, student_groups, student['group_id']):
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
            'edit_grades.html',
            student=student,
            student_groups=student_groups,
            selected_group=None,
            is_choice_view=True,
            subjects=[],
            grade_map={},
            is_reduced=False,
        )

    selected_group = next(
        (g for g in student_groups if g['id'] == selected_group_id), None,
    )
    if not selected_group:
        conn.close()
        flash("Обрана група не належить студенту", "error")
        return redirect(url_for('students.edit_grades', student_id=student_id))

    gid = selected_group['id']
    is_reduced = is_student_on_reduced_program(conn, student_id, gid)
    subjects_query = "SELECT * FROM subjects WHERE group_id = ?" + program_track_condition(is_reduced)
    subjects = [dict(s) for s in conn.execute(subjects_query, (gid,)).fetchall()]
    for subject in subjects:
        apply_track_overrides(subject, is_reduced)

    existing_grades = conn.execute(
        "SELECT subject_id, grade FROM grades WHERE student_id = ?", (student_id,)
    ).fetchall()
    grade_map = {g['subject_id']: g['grade'] for g in existing_grades}

    if request.method == 'POST':
        filled = 0
        for subject in subjects:
            grade_value = request.form.get(f'grade_{subject["id"]}')
            if grade_value:
                filled += 1
                if subject["id"] in grade_map:
                    conn.execute(
                        "UPDATE grades SET grade = ? WHERE student_id = ? AND subject_id = ?",
                        (grade_value, student_id, subject["id"]),
                    )
                else:
                    conn.execute(
                        "INSERT INTO grades (student_id, subject_id, grade) VALUES (?, ?, ?)",
                        (student_id, subject["id"], grade_value),
                    )
        conn.commit()
        conn.close()

        log_action(
            current_username(),
            f"змінив оцінки з дисциплін: {student['last_name_UA']} {student['first_name_UA']} (ID {student_id})",
            group_ids=[gid],
            details=f"заповнено {filled} з {len(subjects)} предметів, група: {selected_group['name']}",
        )
        flash("Оцінки збережено", "success")
        if len(student_groups) > 1:
            return redirect(url_for('students.edit_grades',
                                    student_id=student_id, group_id=gid))
        return redirect(url_for('students.student_list'))

    conn.close()
    return render_template(
        'edit_grades.html',
        student=student,
        student_groups=student_groups,
        selected_group=selected_group,
        is_choice_view=False,
        subjects=subjects,
        grade_map=grade_map,
        is_reduced=is_reduced,
    )