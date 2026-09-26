"""
routes/students/grades.py
==========================
Виставлення оцінок студенту: звичайні предмети (edit_grades) і
активності - практики/курсові/атестації (edit_activities_grades), з
урахуванням скороченої програми (program_track_condition/
apply_track_overrides).
"""
from flask import render_template, request, redirect, url_for, session, flash
from routes.db import get_db
from routes.utils import log_action, login_required, permission_required, logger
from routes.helpers import current_username
from routes.students import students_bp
import sqlite3
from routes.utils import is_student_on_reduced_program, program_track_condition, apply_track_overrides
from datetime import datetime


@students_bp.route('/activities_grades/<int:student_id>', methods=['GET', 'POST'])
@login_required('')
def edit_activities_grades(student_id):
    """Масове виставлення оцінок студентам групи з практик/курсових/атестацій (аналог admin.manage_activities, але з боку картки студента/групи)."""
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

    if session.get('role') != 'admin' and student['group_id'] not in session.get('group_ids', []):
        conn.close()
        flash("Доступ заборонено: студент не належить до вашої групи", "error")
        return redirect(url_for('students.student_list'))

    is_reduced = bool(student['group_id']) and is_student_on_reduced_program(conn, student_id, student['group_id'])
    track_cond = program_track_condition(is_reduced)

    practices = [dict(r) for r in conn.execute(f"""
        SELECT id, code, name, credits, reduced_credits, type, reduced_type, position FROM practices WHERE group_id = ?{track_cond} ORDER BY position
    """, (student['group_id'],)).fetchall()]
    courseworks = [dict(r) for r in conn.execute(f"""
        SELECT id, code, name, credits, reduced_credits, type, reduced_type, position FROM courseworks WHERE group_id = ?{track_cond} ORDER BY position
    """, (student['group_id'],)).fetchall()]
    attestations = [dict(r) for r in conn.execute(f"""
        SELECT a.id, a.code, a.name, a.credits, a.reduced_credits, a.type, a.reduced_type, a.position, ag.name AS student_name
        FROM attestations a
        LEFT JOIN activity_grades ag ON ag.entity_id = a.id AND ag.entity_type = 'attestation' AND ag.student_id = ?
        WHERE a.group_id = ?{track_cond} ORDER BY position
    """, (student_id, student['group_id'])).fetchall()]

    for entity in practices + courseworks + attestations:
        apply_track_overrides(entity, is_reduced)

    existing_grades = conn.execute("""
        SELECT id, entity_id, entity_type, grade, name FROM activity_grades WHERE student_id = ?
    """, (student_id,)).fetchall()
    grade_map = {(g['entity_id'], g['entity_type']): {'id': g['id'], 'grade': g['grade'], 'name': g['name']}
                 for g in existing_grades}

    if request.method == 'POST':
        try:
            for entity_type, entities in [('practice', practices), ('coursework', courseworks), ('attestation', attestations)]:
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
                                """, (grade_value, student_name, grade_map[key]['id'], student_id, entity['id'], entity_type))
                            else:
                                conn.execute("""
                                    INSERT INTO activity_grades (student_id, entity_id, entity_type, grade, name)
                                    VALUES (?, ?, ?, ?, ?)
                                """, (student_id, entity['id'], entity_type, grade_value, student_name))
                        except ValueError:
                            conn.execute("""
                                DELETE FROM activity_grades WHERE student_id = ? AND entity_id = ? AND entity_type = ?
                            """, (student_id, entity['id'], entity_type))
                            flash(f"Некоректна оцінка для {entity['name']}: має бути числом", "error")
                    else:
                        if key in grade_map:
                            conn.execute("""
                                DELETE FROM activity_grades WHERE id = ? AND student_id = ? AND entity_id = ? AND entity_type = ?
                            """, (grade_map[key]['id'], student_id, entity['id'], entity_type))
                        else:
                            conn.execute("""
                                DELETE FROM activity_grades WHERE student_id = ? AND entity_id = ? AND entity_type = ?
                            """, (student_id, entity['id'], entity_type))

            conn.commit()
            flash("Оцінки успішно збережено", "success")
            log_action(
                current_username(),
                f"змінив активності: {student['last_name_UA']} {student['first_name_UA']} (ID {student_id})",
                group_ids=[student['group_id']],
                details=f"практики: {len(practices)}, курсові: {len(courseworks)}, атестації: {len(attestations)}"
            )
            conn.close()
            return redirect(url_for('students.student_list'))
        except Exception as e:
            conn.rollback()
            logger.error(f"Помилка при збереженні оцінок з активностей (student_id={student_id}): {e}", exc_info=True)
            flash(f"Помилка при збереженні оцінок: {str(e)}", "error")

    conn.close()
    return render_template(
        "edit_activities_grades.html",
        student=student, practices=practices, courseworks=courseworks,
        attestations=attestations, grade_map=grade_map
    )


@students_bp.route('/grades/<int:student_id>', methods=['GET', 'POST'])
@login_required('')
def edit_grades(student_id):
    """Форма виставлення/редагування оцінок одного студента з усіх предметів його групи."""
    conn = get_db()

    student = conn.execute("""
        SELECT s.*, g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS group_name
        FROM students s LEFT JOIN groups g ON s.group_id = g.id WHERE s.id = ?
    """, (student_id,)).fetchone()

    if not student:
        conn.close()
        flash("Студент не знайдений")
        return redirect(url_for('students.student_list'))

    subjects_query = "SELECT * FROM subjects WHERE group_id = ?"
    is_reduced = bool(student['group_id']) and is_student_on_reduced_program(conn, student_id, student['group_id'])
    subjects_query += program_track_condition(is_reduced)
    subjects = [dict(s) for s in conn.execute(subjects_query, (student['group_id'],)).fetchall()]
    for subject in subjects:
        apply_track_overrides(subject, is_reduced)
    existing_grades = conn.execute("SELECT subject_id, grade FROM grades WHERE student_id = ?", (student_id,)).fetchall()
    grade_map = {g['subject_id']: g['grade'] for g in existing_grades}

    if request.method == 'POST':
        filled = 0
        for subject in subjects:
            grade_value = request.form.get(f'grade_{subject["id"]}')
            if grade_value:
                filled += 1
                if subject["id"] in grade_map:
                    conn.execute("UPDATE grades SET grade = ? WHERE student_id = ? AND subject_id = ?",
                                 (grade_value, student_id, subject["id"]))
                else:
                    conn.execute("INSERT INTO grades (student_id, subject_id, grade) VALUES (?, ?, ?)",
                                 (student_id, subject["id"], grade_value))
        conn.commit()
        conn.close()

        log_action(
            current_username(),
            f"змінив оцінки з дисциплін: {student['last_name_UA']} {student['first_name_UA']} (ID {student_id})",
            group_ids=[student['group_id']],
            details=f"заповнено {filled} з {len(subjects)} предметів"
        )
        flash("Оцінки збережено")
        return redirect(url_for('students.student_list'))

    conn.close()
    return render_template("edit_grades.html", student=student, subjects=subjects, grade_map=grade_map)
