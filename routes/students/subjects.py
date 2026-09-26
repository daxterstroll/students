"""
routes/admin/subjects.py
=========================
Предмети, практика/атестація/курсові (+ підтримка скороченої
програми), імпорт предметів з Excel.
"""
from flask import render_template, request, redirect, url_for, flash, session, send_file
from routes.db import get_db
from routes.utils import log_action, permission_required, logger
from routes.helpers import current_username, sort_ukrainian
from routes.admin import admin_bp
import sqlite3
import os
from routes.utils import filter_students_for_item
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
from routes.admin import allowed_file
import time

UPLOAD_FOLDER = 'Uploads'


@admin_bp.route('/admin/manage_subjects', methods=['GET', 'POST'])
@permission_required('manage_subjects')
def manage_subjects():
    """CRUD-сторінка навчальних предметів групи: додавання/редагування/видалення/переміщення по позиції, а також масове виставлення оцінок студентам з цього предмету."""
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, name, start_year, study_form, program_credits,
               name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
        FROM groups WHERE archived = FALSE ORDER BY name, start_year
    """)
    groups = cursor.fetchall()

    selected_group_id = request.args.get('group_id')
    subjects = []
    students = []
    grades = []
    selected_subject_id = request.args.get('subject_id')

    if selected_group_id:
        cursor.execute('SELECT * FROM subjects WHERE group_id = ? ORDER BY position', (selected_group_id,))
        subjects = cursor.fetchall()
        if selected_subject_id:
            cursor.execute('SELECT * FROM students WHERE group_id = ?', (selected_group_id,))
            students = cursor.fetchall()
            selected_subject = cursor.execute('SELECT full_program_only, reduced_only FROM subjects WHERE id = ?', (selected_subject_id,)).fetchone()
            if selected_subject:
                # Предмет може стосуватись лише повної АБО лише
                # скороченої програми - прибираємо зі списку тих
                # студентів, кому він не стосується.
                students = filter_students_for_item(students, selected_subject, conn, selected_group_id)
            students = sort_ukrainian(
                students,
                key_func=lambda s: f"{s['last_name_UA']} {s['first_name_UA']} {s['middle_name_UA']}"
            )
            cursor.execute('SELECT id, student_id, subject_id, grade FROM grades WHERE subject_id = ?', (selected_subject_id,))
            grades = cursor.fetchall()

    if request.method == 'POST':
        action = request.form['action']
        group_id = request.form['group_id']

        if action == 'add':
            try:
                code = request.form['code'].strip()
                name = request.form['name'].strip()
                credits = int(request.form['credits'])
                type_ = request.form['type']
                position = int(request.form['position'])
                visibility = request.form.get('visibility', 'all')
                full_program_only = 1 if visibility == 'full_only' else 0
                reduced_only = 1 if visibility == 'reduced_only' else 0
                reduced_credits_raw = (request.form.get('reduced_credits') or '').strip()
                reduced_credits = int(reduced_credits_raw) if reduced_credits_raw else None
                reduced_type = request.form.get('reduced_type') or None
                if reduced_type not in (None, 'Залік', 'Екзамен'):
                    reduced_type = None
                if not code or not name or credits < 1 or position < 1 or type_ not in ['Залік', 'Екзамен']:
                    flash('Некорректные данные предмета', 'error')
                else:
                    cursor.execute('SELECT MAX(position) FROM subjects WHERE group_id = ?', (group_id,))
                    max_position = cursor.fetchone()[0] or 0
                    if position <= max_position:
                        cursor.execute('UPDATE subjects SET position = position + 1 WHERE position >= ? AND group_id = ?', (position, group_id))
                    cursor.execute('INSERT INTO subjects (code, name, credits, type, position, group_id, full_program_only, reduced_only, reduced_credits, reduced_type) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                                   (code, name, credits, type_, position, group_id, full_program_only, reduced_only, reduced_credits, reduced_type))
                    conn.commit()
                    log_action(current_username(),
                               f"додав предмет: {code} — {name} (група ID {group_id})",
                               details=f"кредити: {credits}, тип: {type_}, позиція: {position}, видимість: {visibility}, кредити скорочена: {reduced_credits if reduced_credits is not None else '-'}")
                    flash(f'Добавлен предмет {code}', 'success')
            except (KeyError, ValueError):
                flash('Некорректные данные предмета', 'error')

        elif action == 'edit':
            try:
                subject_id = request.form['subject_id']
                code = request.form['code'].strip()
                name = request.form['name'].strip()
                credits = int(request.form['credits'])
                type_ = request.form['type']
                position = int(request.form['position'])
                visibility = request.form.get('visibility', 'all')
                full_program_only = 1 if visibility == 'full_only' else 0
                reduced_only = 1 if visibility == 'reduced_only' else 0
                reduced_credits_raw = (request.form.get('reduced_credits') or '').strip()
                reduced_credits = int(reduced_credits_raw) if reduced_credits_raw else None
                reduced_type = request.form.get('reduced_type') or None
                if reduced_type not in (None, 'Залік', 'Екзамен'):
                    reduced_type = None
                if not code or not name or credits < 1 or position < 1 or type_ not in ['Залік', 'Екзамен']:
                    flash('Некорректные данные предмета', 'error')
                else:
                    cursor.execute('SELECT position FROM subjects WHERE id = ? AND group_id = ?', (subject_id, group_id))
                    cursor.execute('UPDATE subjects SET position = 0 WHERE id = ? AND group_id = ?', (subject_id, group_id))
                    cursor.execute('UPDATE subjects SET position = position + 1 WHERE position >= ? AND group_id = ? AND id != ?', (position, group_id, subject_id))
                    cursor.execute('UPDATE subjects SET code=?, name=?, credits=?, type=?, position=?, full_program_only=?, reduced_only=?, reduced_credits=?, reduced_type=? WHERE id=? AND group_id=?',
                                   (code, name, credits, type_, position, full_program_only, reduced_only, reduced_credits, reduced_type, subject_id, group_id))
                    cursor.execute('SELECT id, position FROM subjects WHERE group_id=? ORDER BY position, id', (group_id,))
                    for i, subj in enumerate(cursor.fetchall(), 1):
                        if subj['position'] != i:
                            cursor.execute('UPDATE subjects SET position=? WHERE id=? AND group_id=?', (i, subj['id'], group_id))
                    conn.commit()
                    log_action(current_username(),
                               f"редагував предмет: {code} — {name} (група ID {group_id})",
                               details=f"кредити: {credits}, тип: {type_}, позиція: {position}, видимість: {visibility}, кредити скорочена: {reduced_credits if reduced_credits is not None else '-'}")
                    flash(f'Обновлен предмет {code}', 'success')
            except (KeyError, ValueError):
                flash('Некорректные данные предмета', 'error')

        elif action == 'delete':
            try:
                subject_id = request.form['subject_id']
                cursor.execute('SELECT COUNT(*) FROM grades WHERE subject_id=?', (subject_id,))
                grade_count = cursor.fetchone()[0]
                if grade_count > 0:
                    cursor.execute('DELETE FROM grades WHERE subject_id=?', (subject_id,))
                cursor.execute('SELECT position, code, name FROM subjects WHERE id=? AND group_id=?', (subject_id, group_id))
                position_row = cursor.fetchone()
                if position_row is None:
                    flash('Предмет не найден', 'error')
                else:
                    position = position_row[0]
                    subj_code = position_row[1]
                    subj_name = position_row[2]
                    cursor.execute('DELETE FROM subjects WHERE id=? AND group_id=?', (subject_id, group_id))
                    cursor.execute('UPDATE subjects SET position=position-1 WHERE position>? AND group_id=?', (position, group_id))
                    conn.commit()
                    log_action(current_username(),
                               f"ВИДАЛИВ предмет: {subj_code} — {subj_name} (група ID {group_id})",
                               details=f"разом з {grade_count} оцінками" if grade_count > 0 else "оцінок не було")
                    flash(f'Предмет успешно удалён{"" if grade_count == 0 else f" вместе с {grade_count} оценками!"}', 'success')
            except Exception as e:
                conn.rollback()
                logger.error(f"Помилка при видаленні предмету ID {subject_id} (група {group_id}): {e}", exc_info=True)
                flash(f'Ошибка при удалении предмета: {str(e)}', 'error')

        elif action == 'move_up':
            try:
                subject_id = request.form['subject_id']
                cursor.execute('SELECT position FROM subjects WHERE id=? AND group_id=?', (subject_id, group_id))
                current_position = cursor.fetchone()[0]
                cursor.execute('SELECT id, position FROM subjects WHERE position<? AND group_id=? ORDER BY position DESC LIMIT 1', (current_position, group_id))
                prev_subject = cursor.fetchone()
                if prev_subject:
                    cursor.execute('UPDATE subjects SET position=? WHERE id=? AND group_id=?', (prev_subject['position'], subject_id, group_id))
                    cursor.execute('UPDATE subjects SET position=? WHERE id=? AND group_id=?', (current_position, prev_subject['id'], group_id))
                    conn.commit()
                    flash('Предмет перемещен вверх', 'success')
            except (KeyError, ValueError):
                flash('Ошибка при перемещении предмета', 'error')

        elif action == 'move_down':
            try:
                subject_id = request.form['subject_id']
                cursor.execute('SELECT position FROM subjects WHERE id=? AND group_id=?', (subject_id, group_id))
                current_position = cursor.fetchone()[0]
                cursor.execute('SELECT id, position FROM subjects WHERE position>? AND group_id=? ORDER BY position ASC LIMIT 1', (current_position, group_id))
                next_subject = cursor.fetchone()
                if next_subject:
                    cursor.execute('UPDATE subjects SET position=? WHERE id=? AND group_id=?', (next_subject['position'], subject_id, group_id))
                    cursor.execute('UPDATE subjects SET position=? WHERE id=? AND group_id=?', (current_position, next_subject['id'], group_id))
                    conn.commit()
                    flash('Предмет перемещен вниз', 'success')
            except (KeyError, ValueError):
                flash('Ошибка при перемещении предмета', 'error')

        elif action == 'edit_grades':
            try:
                subject_id = request.form['subject_id']
                subj_row = cursor.execute('SELECT name FROM subjects WHERE id=?', (subject_id,)).fetchone()
                subj_name = subj_row['name'] if subj_row else subject_id
                cursor.execute('SELECT id FROM students WHERE group_id=?', (group_id,))
                student_ids = [row['id'] for row in cursor.fetchall()]
                filled = 0
                for sid in student_ids:
                    grade = request.form.get(f'grade_{sid}')
                    grade_id = request.form.get(f'grade_id_{sid}')
                    if grade:
                        try:
                            grade = int(grade)
                            if not (0 <= grade <= 100):
                                flash(f'Оценка для студента {sid} должна быть от 0 до 100', 'error')
                                continue
                            if grade_id:
                                cursor.execute('UPDATE grades SET grade=? WHERE id=? AND student_id=? AND subject_id=?',
                                               (grade, grade_id, sid, subject_id))
                            else:
                                cursor.execute('INSERT INTO grades (student_id, subject_id, grade) VALUES (?, ?, ?)',
                                               (sid, subject_id, grade))
                            filled += 1
                        except ValueError:
                            flash(f'Некорректная оценка для студента {sid}', 'error')
                    else:
                        if grade_id:
                            cursor.execute('DELETE FROM grades WHERE id=? AND student_id=? AND subject_id=?',
                                           (grade_id, sid, subject_id))
                conn.commit()
                log_action(current_username(),
                           f"змінив оцінки з предмету: {subj_name} (група ID {group_id})",
                           details=f"заповнено {filled} з {len(student_ids)} студентів")
                flash('Оценки обновлены', 'success')
            except (KeyError, ValueError) as e:
                flash(f'Ошибка при обновлении оценок: {str(e)}', 'error')

        conn.close()
        return redirect(url_for('admin.manage_subjects', group_id=group_id))

    conn.close()
    return render_template('admin_subjects.html', groups=groups, selected_group_id=selected_group_id,
                           subjects=subjects, students=students, grades=grades, selected_subject_id=selected_subject_id)


@admin_bp.route('/admin/manage_activities', methods=['GET', 'POST'])
@permission_required('manage_activities')
def manage_activities():
    """CRUD-сторінка практик/курсових/атестацій (спільна логіка для трьох типів activity-сутностей через ALLOWED_TABLES) з масовим виставленням оцінок."""
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, name, start_year, study_form, program_credits,
               name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
        FROM groups WHERE archived = FALSE ORDER BY name, start_year
    """)
    groups = cursor.fetchall()

    selected_group_id = request.args.get('group_id')
    selected_entity_type = request.args.get('entity_type', 'practice')
    selected_entity_id = request.args.get('entity_id')
    entities = []
    students = []
    grades = []

    if selected_group_id:
        try:
            selected_group_id = int(selected_group_id)
            cursor.execute('SELECT id FROM groups WHERE id=?', (selected_group_id,))
            if not cursor.fetchone():
                flash('Обрана група не існує', 'error')
                selected_group_id = ''
            else:
                ALLOWED_TABLES = {'practice': 'practices', 'coursework': 'courseworks', 'attestation': 'attestations'}
                entity_table = ALLOWED_TABLES.get(selected_entity_type, 'practices')
                cursor.execute(f'SELECT * FROM {entity_table} WHERE group_id=? ORDER BY position', (selected_group_id,))
                entities = cursor.fetchall()

                if selected_entity_id:
                    try:
                        selected_entity_id = int(selected_entity_id)
                        cursor.execute('SELECT * FROM students WHERE group_id=?', (selected_group_id,))
                        students = cursor.fetchall()
                        selected_entity = cursor.execute(f'SELECT full_program_only, reduced_only FROM {entity_table} WHERE id = ?', (selected_entity_id,)).fetchone()
                        if selected_entity:
                            # Діяльність може стосуватись лише повної
                            # АБО лише скороченої програми - прибираємо
                            # зі списку тих, кому вона не стосується.
                            students = filter_students_for_item(students, selected_entity, conn, selected_group_id)
                        students = sort_ukrainian(
                            students,
                            key_func=lambda s: f"{s['last_name_UA']} {s['first_name_UA']} {s['middle_name_UA']}"
                        )
                        cursor.execute('SELECT id, student_id, entity_id, entity_type, grade, name FROM activity_grades WHERE entity_id=? AND entity_type=?',
                                       (selected_entity_id, selected_entity_type))
                        grades = cursor.fetchall()
                    except ValueError:
                        flash('Некоректний ID діяльності', 'error')
                        selected_entity_id = ''
        except ValueError:
            flash('Некоректний ID групи', 'error')
            selected_group_id = ''

    if request.method == 'POST':
        action = request.form.get('action')
        group_id = request.form.get('group_id')
        entity_type = request.form.get('entity_type', 'practice')
        ALLOWED_TABLES = {'practice': 'practices', 'coursework': 'courseworks', 'attestation': 'attestations'}
        if entity_type not in ALLOWED_TABLES:
            flash('Невірний тип діяльності', 'error')
            return redirect(url_for('admin.manage_activities'))
        entity_table = ALLOWED_TABLES[entity_type]

        try:
            if not group_id:
                flash('ID групи не вказано', 'error')
                return redirect(url_for('admin.manage_activities', group_id=selected_group_id, entity_type=entity_type))

            group_id = int(group_id)
            cursor.execute('SELECT id FROM groups WHERE id=?', (group_id,))
            if not cursor.fetchone():
                flash('Обрана група не існує', 'error')
                return redirect(url_for('admin.manage_activities', entity_type=entity_type))

            if action == 'add':
                code = request.form.get('code')
                name = request.form.get('name')
                credits = request.form.get('credits')
                type_ = request.form.get('type')
                position = request.form.get('position')
                visibility = request.form.get('visibility', 'all')
                full_program_only = 1 if visibility == 'full_only' else 0
                reduced_only = 1 if visibility == 'reduced_only' else 0
                reduced_credits_raw = (request.form.get('reduced_credits') or '').strip()
                reduced_credits = int(reduced_credits_raw) if reduced_credits_raw else None
                reduced_type = request.form.get('reduced_type') or None
                if reduced_type not in (None, 'Залік', 'Екзамен'):
                    reduced_type = None

                if not all([code, name, credits, type_, position]) and entity_type != 'attestation':
                    flash('Усі поля мають бути заповнені', 'error')
                    return redirect(url_for('admin.manage_activities', group_id=group_id, entity_type=entity_type))

                credits = int(credits) if credits else 0
                position = int(position) if position else 1
                if type_ not in ['Залік', 'Екзамен']:
                    flash('Невірний тип оцінки', 'error')
                    return redirect(url_for('admin.manage_activities', group_id=group_id, entity_type=entity_type))

                cursor.execute(f'SELECT MAX(position) FROM {entity_table} WHERE group_id=?', (group_id,))
                max_position = cursor.fetchone()[0] or 0
                if position <= max_position:
                    cursor.execute(f'UPDATE {entity_table} SET position=position+1 WHERE position>=? AND group_id=?', (position, group_id))

                cursor.execute(f'INSERT INTO {entity_table} (code, name, credits, type, position, group_id, full_program_only, reduced_only, reduced_credits, reduced_type) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                               (code, name or '', credits, type_, position, group_id, full_program_only, reduced_only, reduced_credits, reduced_type))
                conn.commit()
                log_action(current_username(),
                           f"додав {entity_type}: {code} — {name} (група ID {group_id})",
                           details=f"кредити: {credits}, тип: {type_}, видимість: {visibility}, кредити скорочена: {reduced_credits if reduced_credits is not None else '-'}")
                flash('Діяльність додано', 'success')

            elif action == 'edit':
                entity_id = request.form.get('entity_id')
                if not entity_id:
                    flash('ID діяльності не вказано', 'error')
                    return redirect(url_for('admin.manage_activities', group_id=group_id, entity_type=entity_type))

                entity_id = int(entity_id)
                code = request.form.get('code')
                name = request.form.get('name')
                credits = int(request.form.get('credits') or 0)
                type_ = request.form.get('type')
                position = int(request.form.get('position'))
                visibility = request.form.get('visibility', 'all')
                full_program_only = 1 if visibility == 'full_only' else 0
                reduced_only = 1 if visibility == 'reduced_only' else 0
                reduced_credits_raw = (request.form.get('reduced_credits') or '').strip()
                reduced_credits = int(reduced_credits_raw) if reduced_credits_raw else None
                reduced_type = request.form.get('reduced_type') or None
                if reduced_type not in (None, 'Залік', 'Екзамен'):
                    reduced_type = None

                cursor.execute(f'SELECT id FROM {entity_table} WHERE id=? AND group_id=?', (entity_id, group_id))
                if not cursor.fetchone():
                    flash('Діяльність з вказаним ID не знайдено', 'error')
                    return redirect(url_for('admin.manage_activities', group_id=group_id, entity_type=entity_type))

                cursor.execute(f'SELECT position FROM {entity_table} WHERE id=?', (entity_id,))
                current_position = cursor.fetchone()[0]
                cursor.execute(f'SELECT MAX(position) FROM {entity_table} WHERE group_id=?', (group_id,))
                max_position = cursor.fetchone()[0] or 0
                if position != current_position and position <= max_position:
                    cursor.execute(f'UPDATE {entity_table} SET position=position+1 WHERE position>=? AND group_id=? AND id!=?',
                                   (position, group_id, entity_id))

                cursor.execute(f'UPDATE {entity_table} SET code=?, name=?, credits=?, type=?, position=?, full_program_only=?, reduced_only=?, reduced_credits=?, reduced_type=? WHERE id=? AND group_id=?',
                               (code, name or '', credits, type_, position, full_program_only, reduced_only, reduced_credits, reduced_type, entity_id, group_id))
                conn.commit()
                log_action(current_username(),
                           f"редагував {entity_type}: {code} — {name} (група ID {group_id})")
                flash('Діяльність оновлено', 'success')

            elif action == 'delete':
                entity_id = int(request.form.get('entity_id'))
                cursor.execute(f'SELECT position, code, name FROM {entity_table} WHERE id=?', (entity_id,))
                row = cursor.fetchone()
                position = row[0]
                cursor.execute(f'DELETE FROM {entity_table} WHERE id=? AND group_id=?', (entity_id, group_id))
                cursor.execute(f'UPDATE {entity_table} SET position=position-1 WHERE position>? AND group_id=?', (position, group_id))
                cursor.execute('DELETE FROM activity_grades WHERE entity_id=? AND entity_type=?', (entity_id, entity_type))
                conn.commit()
                log_action(current_username(),
                           f"ВИДАЛИВ {entity_type}: {row[1]} — {row[2]} (група ID {group_id})")
                flash('Діяльність видалено', 'success')

            elif action == 'move_up':
                entity_id = int(request.form.get('entity_id'))
                cursor.execute(f'SELECT position FROM {entity_table} WHERE id=?', (entity_id,))
                current_position = cursor.fetchone()[0]
                if current_position > 1:
                    cursor.execute(f'UPDATE {entity_table} SET position=? WHERE position=? AND group_id=?',
                                   (current_position, current_position - 1, group_id))
                    cursor.execute(f'UPDATE {entity_table} SET position=? WHERE id=? AND group_id=?',
                                   (current_position - 1, entity_id, group_id))
                    conn.commit()
                    flash('Діяльність переміщено вгору', 'success')

            elif action == 'move_down':
                entity_id = int(request.form.get('entity_id'))
                cursor.execute(f'SELECT position FROM {entity_table} WHERE id=?', (entity_id,))
                current_position = cursor.fetchone()[0]
                cursor.execute(f'SELECT MAX(position) FROM {entity_table} WHERE group_id=?', (group_id,))
                max_position = cursor.fetchone()[0]
                if current_position < max_position:
                    cursor.execute(f'UPDATE {entity_table} SET position=? WHERE position=? AND group_id=?',
                                   (current_position, current_position + 1, group_id))
                    cursor.execute(f'UPDATE {entity_table} SET position=? WHERE id=? AND group_id=?',
                                   (current_position + 1, entity_id, group_id))
                    conn.commit()
                    flash('Діяльність переміщено вниз', 'success')

            elif action == 'edit_grades':
                entity_id = request.form['entity_id']
                entity_row = cursor.execute(f'SELECT name FROM {entity_table} WHERE id=?', (entity_id,)).fetchone()
                entity_name = entity_row['name'] if entity_row else entity_id
                cursor.execute('SELECT id FROM students WHERE group_id=?', (group_id,))
                student_ids = [row['id'] for row in cursor.fetchall()]
                filled = 0
                for sid in student_ids:
                    grade = request.form.get(f'grade_{sid}')
                    grade_id = request.form.get(f'grade_id_{sid}')
                    name_val = request.form.get(f'name_{sid}', '') if entity_type == 'attestation' else ''

                    if grade or name_val:
                        try:
                            grade_int = int(grade) if grade else None
                            if grade_int is not None and not (0 <= grade_int <= 100):
                                flash(f'Оценка для студента {sid} должна быть от 0 до 100', 'error')
                                continue
                            if grade_id:
                                cursor.execute('UPDATE activity_grades SET grade=?, name=? WHERE id=? AND student_id=? AND entity_id=? AND entity_type=?',
                                               (grade_int, name_val, grade_id, sid, entity_id, entity_type))
                            else:
                                cursor.execute('INSERT INTO activity_grades (student_id, entity_id, entity_type, grade, name) VALUES (?, ?, ?, ?, ?)',
                                               (sid, entity_id, entity_type, grade_int, name_val))
                            filled += 1
                        except ValueError:
                            flash(f'Некорректная оценка для студента {sid}', 'error')
                    else:
                        if grade_id:
                            cursor.execute('DELETE FROM activity_grades WHERE id=? AND student_id=? AND entity_id=? AND entity_type=?',
                                           (grade_id, sid, entity_id, entity_type))
                conn.commit()
                log_action(current_username(),
                           f"змінив оцінки з {entity_type}: {entity_name} (група ID {group_id})",
                           details=f"заповнено {filled} з {len(student_ids)} студентів")
                flash('Оценки обновлены', 'success')

            return redirect(url_for('admin.manage_activities', group_id=group_id, entity_type=entity_type))

        except (ValueError, sqlite3.Error) as e:
            conn.rollback()
            flash(f'Помилка: {e}', 'error')
            return redirect(url_for('admin.manage_activities', group_id=selected_group_id, entity_type=entity_type))

    conn.close()
    return render_template('admin_activities.html', groups=groups, selected_group_id=selected_group_id,
                           entities=entities, students=students, grades=grades,
                           selected_entity_id=selected_entity_id, entity_type=selected_entity_type)


@admin_bp.route('/admin/import_subjects', methods=['GET', 'POST'])
@permission_required('import_subjects')
def import_subjects():
    """Імпорт списку предметів групи з Excel-файлу."""
    conn = get_db()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            SELECT id, name, start_year, study_form, program_credits,
                   name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
            FROM groups WHERE archived = FALSE ORDER BY name, start_year
        """)
        groups = cursor.fetchall()
        if not groups:
            flash("Немає доступних груп", "warning")
    except sqlite3.Error as e:
        logger.error(f"Database error while fetching groups: {e}")
        flash("Помилка бази даних при отриманні груп", "error")
        groups = []

    selected_group_id = request.args.get('group_id', '')

    if request.method == 'POST':
        file = request.files.get('excel_file')
        group_id = request.form.get('group_id')

        if not file:
            flash("Будь ласка, виберіть файл", "error")
            return render_template('import_subjects.html', groups=groups, selected_group_id=selected_group_id)

        ext = file.filename.lower().split('.')[-1]
        if ext not in ['xlsx', 'xlsm', 'xltx', 'xltm']:
            flash("❗ Підтримуються тільки Excel файли формату .xlsx", "error")
            return render_template('import_subjects.html', groups=groups, selected_group_id=selected_group_id)

        if not group_id:
            flash("ID групи не вказано", "error")
            return render_template('import_subjects.html', groups=groups, selected_group_id=selected_group_id)

        try:
            group_id = int(group_id)
            cursor.execute("SELECT id FROM groups WHERE id=?", (group_id,))
            if not cursor.fetchone():
                flash("Обрана група не існує", "error")
                return render_template('import_subjects.html', groups=groups, selected_group_id=selected_group_id)
        except ValueError:
            flash("Некоректний ID групи", "error")
            return render_template('import_subjects.html', groups=groups, selected_group_id=selected_group_id)

        filename = f"subjects_{int(time.time())}.xlsx"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        file.save(filepath)

        try:
            try:
                wb = openpyxl.load_workbook(filepath, data_only=True)
                sheet = wb.active
            except InvalidFileException:
                flash("❗ Файл має неправильний формат. Збережіть його як Excel (*.xlsx)", "error")
                os.remove(filepath)
                return render_template('import_subjects.html', groups=groups, selected_group_id=selected_group_id)

            inserted = 0
            skipped = 0
            cursor.execute("SELECT MAX(position) FROM subjects WHERE group_id=?", (group_id,))
            max_position = cursor.fetchone()[0] or 0
            current_position = max_position + 1

            for i, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                if not row or all(cell is None for cell in row):
                    continue
                try:
                    code, name, credits, type_ = row[0], row[1], row[2], row[3]
                    if not all([code, name, credits, type_]):
                        skipped += 1
                        continue
                    code = str(code).strip()
                    name = str(name).strip()
                    type_ = str(type_).strip()
                    if type_ not in ['Залік', 'Екзамен']:
                        flash(f"❗ Невірний тип у рядку {i}: {type_}", "error")
                        skipped += 1
                        continue
                    credits = int(credits)
                    if credits < 1:
                        flash(f"❗ Некоректні кредити у рядку {i}", "error")
                        skipped += 1
                        continue

                    # Скорочена програма (колонки E-G, усі необов'язкові):
                    # видимість / кредити для скороченої / тип для скороченої.
                    full_program_only = 0
                    reduced_only = 0
                    visibility_raw = str(row[4]).strip().lower() if len(row) > 4 and row[4] not in (None, '') else ''
                    if visibility_raw.startswith('лише повна') or visibility_raw.startswith('повна') or visibility_raw == 'full_only':
                        full_program_only = 1
                    elif visibility_raw.startswith('лише скорочена') or visibility_raw.startswith('скорочена') or visibility_raw == 'reduced_only':
                        reduced_only = 1
                    elif visibility_raw and visibility_raw not in ('всі', 'все', 'all', 'усім'):
                        flash(f"⚠️ Рядок {i}: не розпізнано значення видимості '{row[4]}' - предмет імпортовано як звичайний (для всіх)")

                    reduced_credits = None
                    if len(row) > 5 and row[5] not in (None, ''):
                        try:
                            reduced_credits = int(row[5])
                        except (ValueError, TypeError):
                            flash(f"⚠️ Рядок {i}: некоректне значення кредитів для скороченої '{row[5]}' - поле пропущено")

                    reduced_type = None
                    if len(row) > 6 and row[6] not in (None, ''):
                        reduced_type_raw = str(row[6]).strip()
                        if reduced_type_raw in ('Залік', 'Екзамен'):
                            reduced_type = reduced_type_raw
                        else:
                            flash(f"⚠️ Рядок {i}: некоректний тип для скороченої '{row[6]}' - поле пропущено")

                    cursor.execute("SELECT id FROM subjects WHERE group_id=? AND code=?", (group_id, code))
                    if cursor.fetchone():
                        skipped += 1
                        continue
                    cursor.execute(
                        "INSERT INTO subjects (code, name, credits, type, position, group_id, full_program_only, reduced_only, reduced_credits, reduced_type) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (code, name, credits, type_, current_position, group_id, full_program_only, reduced_only, reduced_credits, reduced_type)
                    )
                    inserted += 1
                    current_position += 1
                except Exception as e:
                    logger.error(f"Row {i} error: {e}")
                    skipped += 1
                    continue

            conn.commit()
            log_action(
                current_username(),
                f"імпорт предметів з Excel: додано {inserted}, пропущено {skipped}",
                details=f"група ID: {group_id}, файл: {filename}"
            )
            flash(f"✅ Імпорт завершено. Додано: {inserted}, пропущено: {skipped}", "success")

        except Exception as e:
            conn.rollback()
            flash(f"⚠️ Помилка при імпорті Excel: {e}", "error")
            logger.error(f"Error importing Excel for group_id={group_id}: {e}")
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)
            conn.close()

        return redirect(url_for('admin.manage_subjects', group_id=group_id))

    conn.close()
    return render_template('import_subjects.html', groups=groups, selected_group_id=selected_group_id)
