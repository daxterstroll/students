"""
routes/admin/groups_courses.py
===============================
Групи, курси (kanban-дошка), переведення на наступний курс, архів
груп, оформлення випуску, відрахування, заморожені студенти.
"""
from flask import render_template, request, redirect, url_for, flash, session, send_file
from routes.db import get_db
from routes.utils import log_action, permission_required, logger
from routes.helpers import current_username, sort_ukrainian
from routes.admin import admin_bp
import sqlite3
import os
from routes.utils import get_attachments, save_multiple_attachments
from datetime import datetime
import difflib
import time


def compute_program_total_years(degree_level, program_credits):
    """Тривалість програми в роках - винесено окремо з
    compute_current_course, щоб можна було визначити, чи курс групи вже
    останній (для розмежування "Перевести на наступний курс" і
    "Оформити випуск")."""
    try:
        credits = int(program_credits)
    except (TypeError, ValueError):
        return 1
    if degree_level == 'Бакалавр':
        return 4 if credits == 240 else (3 if credits == 180 else max(1, credits // 60))
    elif degree_level == 'Магістр':
        return 2 if credits in (90, 120) else max(1, credits // 60)
    return max(1, credits // 60)


def compute_current_course(start_year, degree_level, program_credits):
    """
    Обчислює поточний курс групи з року вступу, ступеня і кредитів
    програми - та сама логіка тривалості навчання, що й
    analytics._compute_end_year (щоб не розходились). Використовується
    при створенні групи, щоб курс не доводилось вводити вручну.
    Повертає ціле число від 1 до тривалості програми (не більше).
    """
    try:
        start_year = int(start_year)
        credits = int(program_credits)
    except (TypeError, ValueError):
        return 1

    total_years = compute_program_total_years(degree_level, credits)

    current_year = datetime.now().year
    # Академічний рік починається у вересні - до вересня студенти
    # вступу поточного року року ще на 1 курсі, тобто рахуємо роки, що
    # минули з 1 вересня року вступу.
    if datetime.now().month < 9:
        current_year -= 1

    course = current_year - start_year + 1
    return max(1, min(course, total_years))


# Фіксовані пари УКР->EN (значень лише кілька, тому окремої таблиці в
# базі не потрібно, на відміну від спеціальностей) - джерело істини
# для полів "Ступінь"/"Найменування та статус закладу", щоб форма не
# могла надіслати неузгоджений англійський варіант.
# ПРИМІТКА: наразі ніде фактично не використовується (перевірено -
# жодного виклику по всьому проєкту), лишив на випадок, якщо форма
# створення групи раніше на це розраховувала чи використає в майбутньому.
INSTITUTION_NAME_STATUS_EN = {
    'Приватний вищий навчальний заклад «Європейський університет». Приватна форма власності. Міністерство освіти і науки України. Ліцензія серія ВО № 00228-022801 від 15/05/2017.':
        "Private Higher Educational Institution 'European University'. Private. Ministry of Education and Science of Ukraine. License series BO № 00228-022801 dated 15/05/2017.",
    'Львівська філія Приватного вищого навчального закладу «Європейський університет». Приватна форма власності. Міністерство освіти і науки України. Ліцензія серія ВО № 00228-022801 від 15/05/2017.':
        'Lviv Branch of Private Higher Education Establishment «European University». Private. Ministry of  Education and  Science of Ukraine. License series ВO № 00228-022801 from 15/05/2017.',
}


@admin_bp.route('/admin/manage_groups', methods=['GET', 'POST'])
@permission_required('manage_groups')
def manage_groups():
    """CRUD-сторінка навчальних груп: додавання, редагування та видалення (видалення заборонено, якщо є пов'язані студенти/предмети/активності)."""
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            start_year = request.form.get('start_year')
            study_form = request.form.get('study_form')
            program_credits = request.form.get('program_credits')
            degree_level = request.form.get('degree_level')
            specialty_code = request.form.get('specialty_code') or None
            degree_level_row = conn.execute(
                "SELECT name_en FROM degree_levels WHERE name_ua = ?", (degree_level,)
            ).fetchone()
            degree_level_en = degree_level_row['name_en'] if degree_level_row else None
            entry_requirements = request.form.get('entry_requirements')
            entry_requirements_en = request.form.get('entry_requirements_en')
            learning_outcomes = request.form.get('learning_outcomes')
            learning_outcomes_en = request.form.get('learning_outcomes_en')
            program_includes = request.form.get('program_includes')
            program_includes_en = request.form.get('program_includes_en')
            entry_requirements_reduced = request.form.get('entry_requirements_reduced')
            entry_requirements_reduced_en = request.form.get('entry_requirements_reduced_en')
            learning_outcomes_reduced = request.form.get('learning_outcomes_reduced')
            learning_outcomes_reduced_en = request.form.get('learning_outcomes_reduced_en')
            program_includes_reduced = request.form.get('program_includes_reduced')
            program_includes_reduced_en = request.form.get('program_includes_reduced_en')

            # Курс більше не вводиться вручну при створенні - обчислюється
            # з року вступу, ступеня і кредитів програми (скільки часу
            # вже минуло від 1 вересня року вступу). Якщо реальність
            # відрізняється від обчисленого (академічна відпустка тощо) -
            # можна поправити вручну пізніше через "Редагувати".
            course = compute_current_course(start_year, degree_level, program_credits)

            # Українська Й англійська назви спеціальності та галузі
            # знань більше не вводяться вручну на цій сторінці - усі 4
            # обчислюються з обраного коду спеціальності (Постанова КМУ
            # №1021), щоб гарантувати офіційне написання. Редагувати
            # самі значення каталогу (в т.ч. англійський відповідник) -
            # на сторінці "Спеціальності".
            specialty = None
            specialty_en = None
            knowledge_area = None
            knowledge_area_en = None
            specialty_short_name = None
            if specialty_code:
                cat_row = conn.execute("""
                    SELECT s.name_ua AS specialty_name, s.name_en AS specialty_name_en, s.short_name AS specialty_short_name,
                           k.code AS field_code, k.name_ua AS field_name, k.name_en AS field_name_en
                    FROM specialties s JOIN knowledge_fields k ON k.code = s.knowledge_field_code
                    WHERE s.code = ?
                """, (specialty_code,)).fetchone()
                if cat_row:
                    specialty = f"{specialty_code} {cat_row['specialty_name']}"
                    specialty_en = cat_row['specialty_name_en']
                    knowledge_area = f"{cat_row['field_code']} {cat_row['field_name']}"
                    knowledge_area_en = cat_row['field_name_en']
                    specialty_short_name = cat_row['specialty_short_name']

            # Освітня програма теж більше не вводиться вручну - вона
            # своя для кожної трійки "спеціальність + рік вступу +
            # ступінь" (та сама спеціальність і рік вступу можуть мати
            # зовсім іншу програму на іншому ступені), тому шукаємо в
            # окремому каталозі.
            educational_program = None
            educational_program_en = None
            if specialty_code and start_year and start_year.isdigit() and degree_level:
                ep_row = conn.execute(
                    "SELECT name_ua, name_en FROM educational_programs WHERE specialty_code = ? AND start_year = ? AND degree_level = ?",
                    (specialty_code, int(start_year), degree_level)
                ).fetchone()
                if ep_row:
                    educational_program = ep_row['name_ua']
                    educational_program_en = ep_row['name_en']

            # Назва кваліфікації - той самий принцип, що й освітня
            # програма: своя для кожної трійки "спеціальність + рік
            # вступу + ступінь".
            qualification_name = None
            qualification_name_en = None
            if specialty_code and start_year and start_year.isdigit() and degree_level:
                q_row = conn.execute(
                    "SELECT name_ua, name_en FROM qualification_names WHERE specialty_code = ? AND start_year = ? AND degree_level = ?",
                    (specialty_code, int(start_year), degree_level)
                ).fetchone()
                if q_row:
                    qualification_name = q_row['name_ua']
                    qualification_name_en = q_row['name_en']

            # Назва групи більше не вводиться вручну - вона завжди
            # складається зі скороченої назви спеціальності й курсу
            # (напр. "КН-4"), щоб назва групи автоматично відображала
            # її поточний курс. Для заочної форми додається літера "з"
            # до скороченої назви (напр. "ФБСз-4") - інакше денна й
            # заочна групи того самого курсу отримали б однакову назву
            # і зіштовхувались би через унікальність (назва + рік).
            # "м" для магістратури, "з" для заочної - у такому порядку
            # (напр. "КНм-1" денна магістратура, "КНмз-1" заочна
            # магістратура) - той самий принцип, що вже застосований
            # для форми навчання.
            name_short = specialty_short_name
            if name_short and degree_level == 'Магістр':
                name_short += 'м'
            if name_short and study_form == 'Заочна':
                name_short += 'з' 
            name = f"{name_short}-{course}" if name_short else None

            required_fields = [start_year, study_form, program_credits]
            if not all(required_fields):
                flash("Усі поля мають бути заповнені.", "error")
            elif not specialty_code:
                flash("Оберіть спеціальність зі списку.", "error")
            elif not specialty_short_name:
                flash("У обраної спеціальності не задано скорочену назву - додайте її на сторінці «Спеціальності», щоб можна було сформувати назву групи.", "error")
            elif not specialty_en:
                flash("У обраної спеціальності не задано англійську назву - додайте її на сторінці «Спеціальності» перед створенням групи.", "error")
            elif not educational_program:
                flash("Для цієї спеціальності й року вступу ще не додано освітню програму - додайте її на сторінці «Освітня програма».", "error")
            elif not qualification_name:
                flash("Для цієї спеціальності, ступеня і року вступу ще не додано назву кваліфікації - додайте її на сторінці «Назва кваліфікації».", "error")
            elif study_form not in ['Денна', 'Заочна']:
                flash("Форма навчання має бути 'Денна' або 'Заочна'.", "error")
            elif program_credits not in ['90', '120', '180', '240']:
                flash("Кількість кредитів має бути 90, 120, 180 або 240.", "error")
            else:
                try:
                    course = int(course)
                    start_year = int(start_year)
                    program_credits = int(program_credits)
                    current_year = datetime.now().year
                    if start_year < 2000 or start_year > current_year:
                        flash(f"Рік початку навчання має бути між 2000 і {current_year}.", "error")
                    else:
                        conn.execute("""
                            INSERT INTO groups (
                                name, course, start_year, study_form, program_credits,
                                qualification_name, degree_level, specialty, specialty_code, educational_program, knowledge_area,
                                qualification_name_en, degree_level_en, specialty_en, educational_program_en, knowledge_area_en,
                                entry_requirements, entry_requirements_en,
                                learning_outcomes, learning_outcomes_en, program_includes, program_includes_en,
                                entry_requirements_reduced, entry_requirements_reduced_en,
                                learning_outcomes_reduced, learning_outcomes_reduced_en,
                                program_includes_reduced, program_includes_reduced_en
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (name, course, start_year, study_form, program_credits,
                              qualification_name, degree_level, specialty, specialty_code, educational_program, knowledge_area,
                              qualification_name_en, degree_level_en, specialty_en, educational_program_en, knowledge_area_en,
                              entry_requirements, entry_requirements_en,
                              learning_outcomes, learning_outcomes_en, program_includes, program_includes_en,
                              entry_requirements_reduced, entry_requirements_reduced_en,
                              learning_outcomes_reduced, learning_outcomes_reduced_en,
                              program_includes_reduced, program_includes_reduced_en))
                        conn.commit()
                        flash("Групу додано успішно.", "success")
                        log_action(
                            current_username(),
                            f"додав групу: {name} ({start_year}, {study_form}, {program_credits} кредитів)",
                            details=f"спеціальність: {specialty or 'не вказано'}, ступінь: {degree_level or 'не вказано'}"
                        )
                except ValueError:
                    flash("Рік початку навчання або кредити мають бути числами.", "error")
                except sqlite3.IntegrityError as e:
                    if "UNIQUE constraint" in str(e):
                        flash("Група з такою назвою, роком початку, формою навчання і кількістю кредитів вже існує.", "error")
                    else:
                        logger.error(f"Помилка збереження групи: {e}", exc_info=True)
                        flash(f"Не вдалося зберегти групу через помилку даних: {e}", "error")

        elif action == 'edit':
            group_id = request.form.get('group_id')
            course = request.form.get('course')
            start_year = request.form.get('start_year')
            study_form = request.form.get('study_form')
            program_credits = request.form.get('program_credits')
            degree_level = request.form.get('degree_level')
            specialty_code = request.form.get('specialty_code') or None
            degree_level_row = conn.execute(
                "SELECT name_en FROM degree_levels WHERE name_ua = ?", (degree_level,)
            ).fetchone()
            degree_level_en = degree_level_row['name_en'] if degree_level_row else None
            entry_requirements = request.form.get('entry_requirements')
            entry_requirements_en = request.form.get('entry_requirements_en')
            learning_outcomes = request.form.get('learning_outcomes')
            learning_outcomes_en = request.form.get('learning_outcomes_en')
            program_includes = request.form.get('program_includes')
            program_includes_en = request.form.get('program_includes_en')
            entry_requirements_reduced = request.form.get('entry_requirements_reduced')
            entry_requirements_reduced_en = request.form.get('entry_requirements_reduced_en')
            learning_outcomes_reduced = request.form.get('learning_outcomes_reduced')
            learning_outcomes_reduced_en = request.form.get('learning_outcomes_reduced_en')
            program_includes_reduced = request.form.get('program_includes_reduced')
            program_includes_reduced_en = request.form.get('program_includes_reduced_en')

            specialty = None
            specialty_en = None
            knowledge_area = None
            knowledge_area_en = None
            specialty_short_name = None
            if specialty_code:
                cat_row = conn.execute("""
                    SELECT s.name_ua AS specialty_name, s.name_en AS specialty_name_en, s.short_name AS specialty_short_name,
                           k.code AS field_code, k.name_ua AS field_name, k.name_en AS field_name_en
                    FROM specialties s JOIN knowledge_fields k ON k.code = s.knowledge_field_code
                    WHERE s.code = ?
                """, (specialty_code,)).fetchone()
                if cat_row:
                    specialty = f"{specialty_code} {cat_row['specialty_name']}"
                    specialty_en = cat_row['specialty_name_en']
                    knowledge_area = f"{cat_row['field_code']} {cat_row['field_name']}"
                    knowledge_area_en = cat_row['field_name_en']
                    specialty_short_name = cat_row['specialty_short_name']

            educational_program = None
            educational_program_en = None
            if specialty_code and start_year and start_year.isdigit() and degree_level:
                ep_row = conn.execute(
                    "SELECT name_ua, name_en FROM educational_programs WHERE specialty_code = ? AND start_year = ? AND degree_level = ?",
                    (specialty_code, int(start_year), degree_level)
                ).fetchone()
                if ep_row:
                    educational_program = ep_row['name_ua']
                    educational_program_en = ep_row['name_en']

            # Назва кваліфікації - той самий принцип, що й освітня
            # програма: своя для кожної трійки "спеціальність + рік
            # вступу + ступінь".
            qualification_name = None
            qualification_name_en = None
            if specialty_code and start_year and start_year.isdigit() and degree_level:
                q_row = conn.execute(
                    "SELECT name_ua, name_en FROM qualification_names WHERE specialty_code = ? AND start_year = ? AND degree_level = ?",
                    (specialty_code, int(start_year), degree_level)
                ).fetchone()
                if q_row:
                    qualification_name = q_row['name_ua']
                    qualification_name_en = q_row['name_en']

            # "м" для магістратури, "з" для заочної - у такому порядку
            # (напр. "КНм-1" денна магістратура, "КНмз-1" заочна
            # магістратура) - той самий принцип, що вже застосований
            # для форми навчання.
            name_short = specialty_short_name
            if name_short and degree_level == 'Магістр':
                name_short += 'м'
            if name_short and study_form == 'Заочна':
                name_short += 'з' 
            name = f"{name_short}-{course}" if name_short and course else None

            required_fields = [group_id, start_year, study_form, program_credits, course]
            if not all(required_fields):
                flash("Усі поля мають бути заповнені.", "error")
            elif not specialty_code:
                flash("Оберіть спеціальність зі списку.", "error")
            elif not specialty_short_name:
                flash("У обраної спеціальності не задано скорочену назву - додайте її на сторінці «Спеціальності», щоб можна було сформувати назву групи.", "error")
            elif not specialty_en:
                flash("У обраної спеціальності не задано англійську назву - додайте її на сторінці «Спеціальності» перед створенням групи.", "error")
            elif not educational_program:
                flash("Для цієї спеціальності й року вступу ще не додано освітню програму - додайте її на сторінці «Освітня програма».", "error")
            elif not qualification_name:
                flash("Для цієї спеціальності, ступеня і року вступу ще не додано назву кваліфікації - додайте її на сторінці «Назва кваліфікації».", "error")
            elif not course.isdigit() or int(course) < 1 or int(course) > 6:
                flash("Курс має бути числом від 1 до 6.", "error")
            elif study_form not in ['Денна', 'Заочна']:
                flash("Форма навчання має бути 'Денна' або 'Заочна'.", "error")
            elif program_credits not in ['90', '120', '180', '240']:
                flash("Кількість кредитів має бути 90, 120, 180 або 240.", "error")
            else:
                try:
                    course = int(course)
                    start_year = int(start_year)
                    program_credits = int(program_credits)
                    current_year = datetime.now().year
                    if start_year < 2000 or start_year > current_year:
                        flash(f"Рік початку навчання має бути між 2000 і {current_year}.", "error")
                    else:
                        conn.execute("""
                            UPDATE groups SET
                                name=?, course=?, start_year=?, study_form=?, program_credits=?,
                                qualification_name=?, degree_level=?, specialty=?, specialty_code=?,
                                educational_program=?, knowledge_area=?,
                                qualification_name_en=?, degree_level_en=?, specialty_en=?,
                                educational_program_en=?, knowledge_area_en=?,
                                entry_requirements=?, entry_requirements_en=?,
                                learning_outcomes=?, learning_outcomes_en=?,
                                program_includes=?, program_includes_en=?,
                                entry_requirements_reduced=?, entry_requirements_reduced_en=?,
                                learning_outcomes_reduced=?, learning_outcomes_reduced_en=?,
                                program_includes_reduced=?, program_includes_reduced_en=?
                            WHERE id=?
                        """, (name, course, start_year, study_form, program_credits,
                              qualification_name, degree_level, specialty, specialty_code, educational_program, knowledge_area,
                              qualification_name_en, degree_level_en, specialty_en, educational_program_en, knowledge_area_en,
                              entry_requirements, entry_requirements_en,
                              learning_outcomes, learning_outcomes_en, program_includes, program_includes_en,
                              entry_requirements_reduced, entry_requirements_reduced_en,
                              learning_outcomes_reduced, learning_outcomes_reduced_en,
                              program_includes_reduced, program_includes_reduced_en,
                              group_id))
                        conn.commit()
                        flash("Групу відредаговано успішно.", "success")
                        log_action(
                            current_username(),
                            f"редагував групу: {name} (ID {group_id})",
                            details=f"рік: {start_year}, форма: {study_form}, кредити: {program_credits}"
                        )
                except ValueError:
                    flash("Рік початку навчання або кредити мають бути числами.", "error")
                except sqlite3.IntegrityError as e:
                    if "UNIQUE constraint" in str(e):
                        flash("Група з такою назвою, роком початку, формою навчання і кількістю кредитів вже існує.", "error")
                    else:
                        logger.error(f"Помилка збереження групи: {e}", exc_info=True)
                        flash(f"Не вдалося зберегти групу через помилку даних: {e}", "error")

        elif action == 'delete':
            group_id = request.form.get('group_id')
            group_row = conn.execute("SELECT name, start_year FROM groups WHERE id=?", (group_id,)).fetchone()
            counts = conn.execute("""
                SELECT
                    (SELECT COUNT(*) FROM students WHERE group_id=?) AS students,
                    (SELECT COUNT(*) FROM subjects WHERE group_id=?) AS subjects,
                    (SELECT COUNT(*) FROM practices WHERE group_id=?) AS practices,
                    (SELECT COUNT(*) FROM courseworks WHERE group_id=?) AS courseworks,
                    (SELECT COUNT(*) FROM attestations WHERE group_id=?) AS attestations
            """, (group_id, group_id, group_id, group_id, group_id)).fetchone()

            labels = {
                'students': 'студентів', 'subjects': 'предметів', 'practices': 'практик',
                'courseworks': 'курсових робіт', 'attestations': 'атестацій',
            }
            parts = [f"{counts[key]} {label}" for key, label in labels.items() if counts[key] > 0]

            if parts:
                flash(
                    f"Неможливо видалити групу «{group_row['name']}» - пов'язані дані: {', '.join(parts)}. "
                    f"Натисніть «Видалити групу разом з усім пов'язаним», щоб прибрати це остаточно.",
                    "error"
                )
            else:
                conn.execute("DELETE FROM groups WHERE id=?", (group_id,))
                conn.commit()
                flash("Групу видалено успішно.", "success")
                log_action(
                    current_username(),
                    f"ВИДАЛИВ групу: {group_row['name']} ({group_row['start_year']}) (ID {group_id})"
                )

        elif action == 'force_delete':
            group_id = request.form.get('group_id')
            group_row = conn.execute("SELECT name, start_year FROM groups WHERE id=?", (group_id,)).fetchone()
            if not group_row:
                flash("Групу не знайдено.", "error")
            else:
                student_count = conn.execute("SELECT COUNT(*) AS c FROM students WHERE group_id=?", (group_id,)).fetchone()['c']
                transfer_order_count = conn.execute(
                    "SELECT COUNT(*) AS c FROM course_transfer_order_groups WHERE group_id=?", (group_id,)
                ).fetchone()['c']

                if student_count > 0:
                    flash(
                        f"У групі «{group_row['name']}» досі {student_count} студент(ів) - "
                        f"спершу перенесіть чи видаліть їх окремо (каскадне видалення студентів разом з групою не підтримується).",
                        "error"
                    )
                elif transfer_order_count > 0:
                    # Це реальний наказ (документ) про переведення на курс,
                    # у якому фігурувала ця група - видаляти таку історію
                    # мовчки не можна, це зіпсує аудиторський слід наказу.
                    flash(
                        f"Групу «{group_row['name']}» не можна видалити - вона згадується в {transfer_order_count} "
                        f"наказ(ах) про переведення на курс. Видалення знищило б історію наказу.",
                        "error"
                    )
                else:
                    # Студентів У ГРУПІ ЗАРАЗ немає, але оцінки могли
                    # лишитись від студентів, яких раніше перевели чи
                    # заморозили з цієї групи (grades/activity_grades
                    # прив'язані до subject_id/entity_id конкретної
                    # групи, а не до student_id теперішньої групи) -
                    # тому спершу прибираємо такі "осиротілі" оцінки,
                    # інакше FOREIGN KEY constraint не дасть видалити
                    # самі предмети/практики/курсові/атестації.
                    subject_ids = [r['id'] for r in conn.execute("SELECT id FROM subjects WHERE group_id=?", (group_id,)).fetchall()]
                    if subject_ids:
                        placeholders = ','.join('?' for _ in subject_ids)
                        conn.execute(f"DELETE FROM grades WHERE subject_id IN ({placeholders})", subject_ids)

                    for table, entity_type in (('practices', 'practice'), ('courseworks', 'coursework'), ('attestations', 'attestation')):
                        entity_ids = [r['id'] for r in conn.execute(f"SELECT id FROM {table} WHERE group_id=?", (group_id,)).fetchall()]
                        if entity_ids:
                            placeholders = ','.join('?' for _ in entity_ids)
                            conn.execute(
                                f"DELETE FROM activity_grades WHERE entity_type=? AND entity_id IN ({placeholders})",
                                [entity_type] + entity_ids
                            )

                    # Права доступу викладачів до цієї групи - суто
                    # службові рядки, без самостійної цінності без групи.
                    conn.execute("DELETE FROM user_groups WHERE group_id=?", (group_id,))

                    # Записи заморожування/відрахування студентів, які
                    # КОЛИСЬ вийшли саме з цієї групи, - самі записи
                    # (причина, наказ, дата) лишаються цінними без
                    # прив'язки до конкретної групи, тому просто
                    # обнуляємо посилання, а не видаляємо весь запис.
                    conn.execute("UPDATE frozen_students SET previous_group_id = NULL WHERE previous_group_id=?", (group_id,))
                    conn.execute("UPDATE expulsion_order_students SET previous_group_id = NULL WHERE previous_group_id=?", (group_id,))

                    for table in ('subjects', 'practices', 'courseworks', 'attestations'):
                        conn.execute(f"DELETE FROM {table} WHERE group_id=?", (group_id,))
                    conn.execute("DELETE FROM groups WHERE id=?", (group_id,))
                    conn.commit()
                    flash(f"Групу «{group_row['name']}» і всі пов'язані дані видалено остаточно.", "success")
                    log_action(
                        current_username(),
                        f"ВИДАЛИВ групу разом з пов'язаними даними: {group_row['name']} ({group_row['start_year']}) (ID {group_id})"
                    )

    # Сортування списку груп - клікабельні заголовки таблиці (той самий
    # принцип, що й на сторінці студентів). Текстові поля - з
    # COLLATE UKRAINIAN, щоб сортувались за правильною абеткою.
    sortable_columns = {
        'id': 'g.id', 'name': 'g.name COLLATE UKRAINIAN', 'course': 'g.course',
        'start_year': 'g.start_year', 'study_form': 'g.study_form COLLATE UKRAINIAN',
        'program_credits': 'g.program_credits', 'specialty': 'g.specialty COLLATE UKRAINIAN',
        'degree_level': 'g.degree_level COLLATE UKRAINIAN',
        'educational_program': 'g.educational_program COLLATE UKRAINIAN',
        'knowledge_area': 'g.knowledge_area COLLATE UKRAINIAN',
        'qualification_name': 'g.qualification_name COLLATE UKRAINIAN',
        'student_count': 'student_count',
    }
    sort_by = request.args.get('sort_by', 'id')
    sort_order = request.args.get('sort_order', 'asc')
    if sort_by not in sortable_columns:
        sort_by = 'id'
    if sort_order not in ('asc', 'desc'):
        sort_order = 'asc'

    groups = conn.execute(f"""
        SELECT g.id, g.name, g.course, g.start_year, g.study_form, g.program_credits,
               g.qualification_name, g.degree_level, g.specialty, g.specialty_code, g.educational_program, g.knowledge_area,
               g.qualification_name_en, g.degree_level_en, g.specialty_en, g.educational_program_en, g.knowledge_area_en,
               g.entry_requirements, g.entry_requirements_en,
               g.learning_outcomes, g.learning_outcomes_en, g.program_includes, g.program_includes_en,
               g.entry_requirements_reduced, g.entry_requirements_reduced_en,
               g.learning_outcomes_reduced, g.learning_outcomes_reduced_en,
               g.program_includes_reduced, g.program_includes_reduced_en,
               g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS display_name,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = g.id) AS student_count,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = g.id
                    AND s.program_credits_override IS NOT NULL
                    AND s.program_credits_override < g.program_credits) AS reduced_count
        FROM groups g WHERE g.archived = FALSE ORDER BY {sortable_columns[sort_by]} {sort_order.upper()}
    """).fetchall()

    # Каталог ступенів для випадаючого списку "Ступінь" - та сама
    # логіка "активний/неактивний", що й у спеціальностей.
    degree_levels = conn.execute("""
        SELECT id, name_ua, name_en, is_active FROM degree_levels ORDER BY is_active DESC, id
    """).fetchall()

    # Освітні програми (спеціальність + рік вступу -> назва) - для
    # автопідстановки на JS-стороні за парою вже обраних полів.
    educational_programs_map = {}
    for row in conn.execute("SELECT specialty_code, start_year, degree_level, name_ua, name_en FROM educational_programs"):
        key = f"{row['specialty_code']}|{row['start_year']}|{row['degree_level']}"
        educational_programs_map[key] = {'name_ua': row['name_ua'], 'name_en': row['name_en'] or ''}

    qualification_names_map = {}
    for row in conn.execute("SELECT specialty_code, start_year, degree_level, name_ua, name_en FROM qualification_names"):
        key = f"{row['specialty_code']}|{row['start_year']}|{row['degree_level']}"
        qualification_names_map[key] = {'name_ua': row['name_ua'], 'name_en': row['name_en'] or ''}

    # Каталог спеціальностей (Постанова КМУ №266) для випадаючого
    # списку - згруповано по галузях знань, щоб було зручно шукати.
    specialty_catalog = conn.execute("""
        SELECT s.code, s.name_ua, s.short_name, s.name_en, s.is_active,
               k.code AS field_code, k.name_ua AS field_name, k.name_en AS field_name_en
        FROM specialties s JOIN knowledge_fields k ON k.code = s.knowledge_field_code
        ORDER BY s.is_active DESC, k.code, substr(s.code,1,1), CAST(substr(s.code,2) AS INTEGER)
    """).fetchall()

    conn.close()
    return render_template("manage_groups.html", groups=groups, specialty_catalog=specialty_catalog, degree_levels=degree_levels, educational_programs_map=educational_programs_map, qualification_names_map=qualification_names_map, sort_by=sort_by, sort_order=sort_order)


@admin_bp.route('/admin/courses')
@permission_required('manage_courses')
def courses():
    """
    Дошка "Курси": активні групи, згруповані по ступеню навчання і в
    межах ступеня - по поточному курсу (окремий рядок дошки на кожен
    ступінь, щоб 1 курс бакалаврів і 1 курс магістрів не змішувались в
    одній колонці). Наочний перегляд перед сезоном переведення на
    курс - поки що лише перегляд, без дій (додаються на наступних
    кроках).
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    groups = conn.execute("""
        SELECT id, name, course, start_year, study_form, program_credits,
               degree_level, specialty,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = groups.id AND COALESCE(s.archived, 0) = 0) AS student_count,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = groups.id AND COALESCE(s.archived, 0) = 0
                    AND s.program_credits_override IS NOT NULL
                    AND s.program_credits_override < groups.program_credits) AS reduced_count
        FROM groups
        WHERE archived = FALSE
        ORDER BY course, name
    """).fetchall()
    groups = [dict(g) for g in groups]
    for g in groups:
        g['is_final_course'] = g['course'] >= compute_program_total_years(g['degree_level'], g['program_credits'])

    # Порядок ступенів - за їхнім id в каталозі degree_levels (природна
    # ієрархія: молодший бакалавр -> бакалавр -> магістр -> доктор
    # філософії). Ступені, яких немає в каталозі (напр. видалені чи
    # перейменовані вручну), йдуть в кінець, за абеткою.
    degree_order = {
        row['name_ua']: row['id']
        for row in conn.execute("SELECT id, name_ua FROM degree_levels").fetchall()
    }

    groups_by_degree = {}
    for g in groups:
        groups_by_degree.setdefault(g['degree_level'] or '—', []).append(g)

    degree_sections = []
    for degree_level in sorted(groups_by_degree.keys(), key=lambda d: (degree_order.get(d, 999), d)):
        degree_groups = groups_by_degree[degree_level]
        groups_by_course = {}
        for g in degree_groups:
            groups_by_course.setdefault(g['course'], []).append(g)
        degree_sections.append({
            'degree_level': degree_level,
            'groups_by_course': groups_by_course,
            'courses_sorted': sorted(groups_by_course.keys()),
        })

    conn.close()

    return render_template(
        "courses.html",
        degree_sections=degree_sections,
    )


@admin_bp.route('/admin/course_transfer', methods=['GET', 'POST'])
@permission_required('manage_courses')
def course_transfer():
    """
    Крок 1 -> 2 процесу "Перевести на наступний курс".
    GET: показує активні групи (крім тих, що вже на останньому курсі
        своєї програми - для них окремий процес "Оформити випуск").
    POST: за обраними групами показує список усіх їхніх студентів для
        перегляду й, за потреби, виключення з причиною, перед
        остаточним підтвердженням (course_transfer_confirm).
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        group_ids = request.form.getlist('group_ids')
        if not group_ids:
            flash("Оберіть хоча б одну групу для переведення.", "error")
            conn.close()
            return redirect(url_for('admin.course_transfer'))

        placeholders = ','.join('?' for _ in group_ids)
        selected_groups = conn.execute(f"""
            SELECT id, name, course, degree_level, program_credits
            FROM groups WHERE id IN ({placeholders}) AND archived = FALSE
            ORDER BY name
        """, group_ids).fetchall()

        groups_with_students = []
        for g in selected_groups:
            students = conn.execute("""
                SELECT id, TRIM(last_name_UA || ' ' || first_name_UA) AS full_name, program_credits_override
                FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0
                ORDER BY last_name_UA COLLATE UKRAINIAN
            """, (g['id'],)).fetchall()
            groups_with_students.append({
                'id': g['id'], 'name': g['name'], 'course': g['course'],
                'course_to': g['course'] + 1,
                'program_credits': g['program_credits'],
                'students': students,
            })

        conn.close()
        return render_template("course_transfer_review.html", groups_with_students=groups_with_students)

    groups = conn.execute("""
        SELECT id, name, course, start_year, study_form, degree_level, program_credits,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = groups.id AND COALESCE(s.archived, 0) = 0) AS student_count
        FROM groups WHERE archived = FALSE ORDER BY course, name
    """).fetchall()
    conn.close()

    groups = [dict(g) for g in groups]
    for g in groups:
        g['is_final_course'] = g['course'] >= compute_program_total_years(g['degree_level'], g['program_credits'])
    # На цьому кроці показуємо лише групи, які МОЖУТЬ перейти на
    # наступний курс - випускні йдуть окремим процесом "Оформити випуск".
    transferable_groups = [g for g in groups if not g['is_final_course']]

    return render_template("course_transfer_select.html", groups=transferable_groups)


@admin_bp.route('/admin/course_transfer/confirm', methods=['POST'])
@permission_required('manage_courses')
def course_transfer_confirm():
    """Крок 3: остаточне підтвердження переведення - зберігає наказ,
    оновлює курс/назву обраних груп, виключених студентів заморожує."""
    order_number = (request.form.get('order_number') or '').strip()
    order_date = request.form.get('order_date')
    group_ids = [int(x) for x in request.form.getlist('group_ids')]

    if not order_number or not order_date or not group_ids:
        flash("Заповніть номер і дату наказу.", "error")
        return redirect(url_for('admin.course_transfer'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    # Перевіряємо, що для кожного виключеного студента вказано причину
    excluded_students = []  # (student_id, group_id, reason)
    for group_id in group_ids:
        students = conn.execute(
            "SELECT id FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0", (group_id,)
        ).fetchall()
        for s in students:
            included = request.form.get(f"include_{s['id']}") == '1'
            if not included:
                reason = (request.form.get(f"reason_{s['id']}") or '').strip()
                if not reason:
                    flash(f"Для виключеного студента (ID {s['id']}) не вказано причину.", "error")
                    conn.close()
                    return redirect(url_for('admin.course_transfer'))
                excluded_students.append((s['id'], group_id, reason))

    # Скан наказу - можна прикріпити кілька файлів будь-якого типу
    scan_files = request.files.getlist('scan_files')

    cur = conn.execute(
        "INSERT INTO course_transfer_orders (order_number, order_date, scan_file, created_by) VALUES (?, ?, NULL, ?)",
        (order_number, order_date, current_username())
    )
    order_id = cur.lastrowid
    saved_paths = save_multiple_attachments(conn, 'course_transfer_order', order_id, scan_files, 'course_transfer_orders', current_username())
    if saved_paths:
        conn.execute("UPDATE course_transfer_orders SET scan_file = ? WHERE id = ?", (saved_paths[0], order_id))

    transferred_groups_summary = []
    for group_id in group_ids:
        group = conn.execute("SELECT name, course, specialty_code FROM groups WHERE id = ?", (group_id,)).fetchone()
        if not group:
            continue
        course_from = group['course']
        course_to = course_from + 1

        conn.execute(
            "INSERT INTO course_transfer_order_groups (order_id, group_id, course_from, course_to) VALUES (?, ?, ?, ?)",
            (order_id, group_id, course_from, course_to)
        )

        short_name_row = conn.execute(
            "SELECT short_name FROM specialties WHERE code = ?", (group['specialty_code'],)
        ).fetchone()
        new_name = f"{short_name_row['short_name']}-{course_to}" if short_name_row and short_name_row['short_name'] else group['name']

        conn.execute("UPDATE groups SET course = ?, name = ? WHERE id = ?", (course_to, new_name, group_id))
        transferred_groups_summary.append(f"{group['name']} -> {new_name}")

    for student_id, group_id, reason in excluded_students:
        conn.execute(
            "INSERT INTO frozen_students (student_id, previous_group_id, order_id, reason, frozen_by) VALUES (?, ?, ?, ?, ?)",
            (student_id, group_id, order_id, reason, current_username())
        )
        conn.execute("UPDATE students SET group_id = NULL WHERE id = ?", (student_id,))

    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"наказ про переведення на курс №{order_number} від {order_date}",
        details=f"груп: {len(group_ids)}, заморожено студентів: {len(excluded_students)} | {', '.join(transferred_groups_summary)}"
    )
    flash(f"Переведення оформлено. Груп: {len(group_ids)}, заморожено студентів: {len(excluded_students)}.", "success")
    return redirect(url_for('admin.courses'))


@admin_bp.route('/admin/frozen_students')
@permission_required('manage_frozen_students')
def frozen_students():
    """
    Перегляд заморожених студентів: активні (ще не вирішено) окремо
    від історії (вже вирішено). Дія "Вирішити" - повернути студента в
    конкретну групу з коментарем. "Передати на відрахування" тут
    свідомо не робиться - це окремий процес "Наказ про відрахування",
    який сам підхопить студента звідси й закриє цей запис.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    active = conn.execute("""
        SELECT f.id, f.student_id, f.reason, f.document_file, f.frozen_at, f.frozen_by,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS student_name,
               g.name AS previous_group_name,
               o.order_number, o.order_date
        FROM frozen_students f
        JOIN students s ON s.id = f.student_id
        LEFT JOIN groups g ON g.id = f.previous_group_id
        LEFT JOIN course_transfer_orders o ON o.id = f.order_id
        WHERE f.resolved_at IS NULL
        ORDER BY f.frozen_at
    """).fetchall()

    history = conn.execute("""
        SELECT f.id, f.student_id, f.reason, f.document_file, f.frozen_at, f.frozen_by,
               f.resolved_at, f.resolution, f.resolved_by,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS student_name,
               g.name AS previous_group_name,
               o.order_number, o.order_date
        FROM frozen_students f
        JOIN students s ON s.id = f.student_id
        LEFT JOIN groups g ON g.id = f.previous_group_id
        LEFT JOIN course_transfer_orders o ON o.id = f.order_id
        WHERE f.resolved_at IS NOT NULL
        ORDER BY f.resolved_at DESC
    """).fetchall()

    active_groups = conn.execute(
        "SELECT id, name FROM groups WHERE archived = FALSE ORDER BY name"
    ).fetchall()

    # Вкладення (кілька файлів) для кожного запису - для старих записів
    # (до появи attachments) список буде порожній, і шаблон тоді сам
    # показує старе одиничне поле document_file як запасний варіант.
    attachments_by_frozen_id = {}
    for row in list(active) + list(history):
        attachments_by_frozen_id[row['id']] = get_attachments(conn, 'frozen_student', row['id'])

    conn.close()
    return render_template(
        "frozen_students.html", active=active, history=history, active_groups=active_groups,
        attachments_by_frozen_id=attachments_by_frozen_id,
    )


@admin_bp.route('/admin/frozen_students/<int:frozen_id>/resolve', methods=['POST'])
@permission_required('manage_frozen_students')
def resolve_frozen_student(frozen_id):
    """
    Крок 1 вирішення заморожування: перевіряє обрану групу і показує
    попередній перегляд оцінок, які можна перенести в нову групу.
    Предмети/практики/курсові/атестації належать конкретній групі
    (group_id), тому оцінка студента прив'язана до конкретного рядка
    старої групи - при переведенні в нову групу такі оцінки самі по
    собі НЕ переносяться, навіть якщо назва предмета та сама. Тут
    зіставляємо старі й нові за назвою (нечітке зіставлення) і
    пропонуємо перенести вибірково.
    """
    import difflib

    target_group_id = request.form.get('target_group_id')
    comment = (request.form.get('comment') or '').strip()

    conn = get_db()
    conn.row_factory = sqlite3.Row
    frozen = conn.execute(
        "SELECT student_id, previous_group_id FROM frozen_students WHERE id = ? AND resolved_at IS NULL", (frozen_id,)
    ).fetchone()
    if not frozen:
        flash("Запис не знайдено або вже вирішено.", "error")
        conn.close()
        return redirect(url_for('admin.frozen_students'))

    if not target_group_id:
        flash("Оберіть групу, до якої повернути студента.", "error")
        conn.close()
        return redirect(url_for('admin.frozen_students'))

    target_group = conn.execute("SELECT name FROM groups WHERE id = ?", (target_group_id,)).fetchone()
    if not target_group:
        flash("Обрану групу не знайдено.", "error")
        conn.close()
        return redirect(url_for('admin.frozen_students'))

    old_group_id = frozen['previous_group_id']
    matches = []

    if old_group_id:
        # Предмети - оцінка в grades за subject_id
        old_subjects = conn.execute("SELECT id, name FROM subjects WHERE group_id = ?", (old_group_id,)).fetchall()
        new_subjects = conn.execute("SELECT id, name FROM subjects WHERE group_id = ?", (target_group_id,)).fetchall()
        for old_s in old_subjects:
            grade_row = conn.execute(
                "SELECT grade FROM grades WHERE student_id = ? AND subject_id = ?", (frozen['student_id'], old_s['id'])
            ).fetchone()
            if not grade_row or not grade_row['grade']:
                continue
            best = max(new_subjects, key=lambda ns: difflib.SequenceMatcher(None, old_s['name'].lower(), ns['name'].lower()).ratio(), default=None)
            if best and difflib.SequenceMatcher(None, old_s['name'].lower(), best['name'].lower()).ratio() >= 0.85:
                matches.append({
                    'kind': 'subject', 'old_id': old_s['id'], 'new_id': best['id'],
                    'name': old_s['name'], 'new_name': best['name'], 'grade': grade_row['grade'],
                })

        # Практики/курсові/атестації - оцінка в activity_grades за (entity_id, entity_type)
        for kind, table in (('practice', 'practices'), ('coursework', 'courseworks'), ('attestation', 'attestations')):
            old_entities = conn.execute(f"SELECT id, name FROM {table} WHERE group_id = ?", (old_group_id,)).fetchall()
            new_entities = conn.execute(f"SELECT id, name FROM {table} WHERE group_id = ?", (target_group_id,)).fetchall()
            for old_e in old_entities:
                grade_row = conn.execute(
                    "SELECT grade FROM activity_grades WHERE student_id = ? AND entity_id = ? AND entity_type = ?",
                    (frozen['student_id'], old_e['id'], kind)
                ).fetchone()
                if not grade_row or grade_row['grade'] is None:
                    continue
                best = max(new_entities, key=lambda ne: difflib.SequenceMatcher(None, old_e['name'].lower(), ne['name'].lower()).ratio(), default=None)
                if best and difflib.SequenceMatcher(None, old_e['name'].lower(), best['name'].lower()).ratio() >= 0.85:
                    matches.append({
                        'kind': kind, 'old_id': old_e['id'], 'new_id': best['id'],
                        'name': old_e['name'], 'new_name': best['name'], 'grade': grade_row['grade'],
                    })

    conn.close()

    if not matches:
        # Нема чого переносити (чи взагалі не було старої групи) -
        # одразу виконуємо перенесення без окремого кроку підтвердження.
        return _do_resolve_frozen_student(frozen_id, frozen['student_id'], target_group_id, target_group['name'], comment, [])

    return render_template(
        "frozen_student_resolve_review.html",
        frozen_id=frozen_id, target_group_id=target_group_id, target_group_name=target_group['name'],
        comment=comment, matches=matches,
    )


def _do_resolve_frozen_student(frozen_id, student_id, target_group_id, target_group_name, comment, grades_to_transfer):
    """Виконує саме переведення: зміна групи, перенесення обраних оцінок, закриття заморожування."""
    conn = get_db()
    conn.row_factory = sqlite3.Row

    conn.execute("UPDATE students SET group_id = ? WHERE id = ?", (target_group_id, student_id))

    for m in grades_to_transfer:
        if m['kind'] == 'subject':
            conn.execute("INSERT INTO grades (student_id, subject_id, grade) VALUES (?, ?, ?)",
                         (student_id, m['new_id'], m['grade']))
        else:
            conn.execute("INSERT INTO activity_grades (student_id, entity_id, entity_type, grade, name) VALUES (?, ?, ?, ?, ?)",
                         (student_id, m['new_id'], m['kind'], m['grade'], m['new_name']))

    resolution_text = f"Повернено до групи «{target_group_name}»" + (f" - {comment}" if comment else "")
    if grades_to_transfer:
        resolution_text += f" (перенесено оцінок: {len(grades_to_transfer)})"

    conn.execute(
        "UPDATE frozen_students SET resolved_at = datetime('now', 'localtime'), resolution = ?, resolved_by = ? WHERE id = ?",
        (resolution_text, current_username(), frozen_id)
    )
    conn.commit()
    conn.close()

    log_action(current_username(), f"вирішив заморожування студента (ID {student_id}): {resolution_text}")
    flash("Студента повернено в групу." + (f" Перенесено оцінок: {len(grades_to_transfer)}." if grades_to_transfer else ""), "success")
    return redirect(url_for('admin.frozen_students'))


@admin_bp.route('/admin/frozen_students/<int:frozen_id>/resolve_confirm', methods=['POST'])
@permission_required('manage_frozen_students')
def resolve_frozen_student_confirm(frozen_id):
    """Крок 2: остаточне підтвердження з переліку оцінок, обраних для перенесення."""
    target_group_id = request.form.get('target_group_id')
    target_group_name = request.form.get('target_group_name')
    comment = request.form.get('comment') or ''

    conn = get_db()
    conn.row_factory = sqlite3.Row
    frozen = conn.execute("SELECT student_id FROM frozen_students WHERE id = ? AND resolved_at IS NULL", (frozen_id,)).fetchone()
    if not frozen:
        flash("Запис не знайдено або вже вирішено.", "error")
        conn.close()
        return redirect(url_for('admin.frozen_students'))
    student_id = frozen['student_id']
    conn.close()

    selected_keys = request.form.getlist('transfer')  # "kind|old_id|new_id|grade" (grade передається окремо)
    grades_to_transfer = []
    for key in selected_keys:
        kind, old_id, new_id = key.split('|')
        grade = request.form.get(f"grade_{key}")
        new_name = request.form.get(f"new_name_{key}")
        grades_to_transfer.append({'kind': kind, 'old_id': old_id, 'new_id': new_id, 'grade': grade, 'new_name': new_name})

    return _do_resolve_frozen_student(frozen_id, student_id, target_group_id, target_group_name, comment, grades_to_transfer)


@admin_bp.route('/admin/graduation', methods=['GET', 'POST'])
@permission_required('manage_courses')
def graduation():
    """
    Крок 1 -> 2 процесу "Оформити випуск" - той самий механізм, що й
    переведення на курс (вибір груп -> перегляд студентів з можливістю
    виключення), але лише для груп на останньому курсі своєї програми,
    і результат - архівування замість +1 до курсу.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        group_ids = request.form.getlist('group_ids')
        if not group_ids:
            flash("Оберіть хоча б одну групу для випуску.", "error")
            conn.close()
            return redirect(url_for('admin.graduation'))

        placeholders = ','.join('?' for _ in group_ids)
        selected_groups = conn.execute(f"""
            SELECT id, name FROM groups WHERE id IN ({placeholders}) AND archived = FALSE ORDER BY name
        """, group_ids).fetchall()

        groups_with_students = []
        for g in selected_groups:
            students = conn.execute("""
                SELECT id, TRIM(last_name_UA || ' ' || first_name_UA) AS full_name
                FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0
                ORDER BY last_name_UA COLLATE UKRAINIAN
            """, (g['id'],)).fetchall()
            groups_with_students.append({'id': g['id'], 'name': g['name'], 'students': students})

        conn.close()
        return render_template("graduation_review.html", groups_with_students=groups_with_students)

    groups = conn.execute("""
        SELECT id, name, course, start_year, study_form, degree_level, program_credits,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = groups.id AND COALESCE(s.archived, 0) = 0) AS student_count
        FROM groups WHERE archived = FALSE ORDER BY name
    """).fetchall()
    conn.close()

    # Показуємо лише групи, які ВЖЕ на останньому курсі своєї програми
    final_groups = [
        g for g in groups
        if g['course'] >= compute_program_total_years(g['degree_level'], g['program_credits'])
    ]

    return render_template("graduation_select.html", groups=final_groups)


@admin_bp.route('/admin/graduation/confirm', methods=['POST'])
@permission_required('manage_courses')
def graduation_confirm():
    """
    Підтвердження випуску: виключені студенти (з причиною) ідуть у
    frozen_students (без наказу переведення - для випуску формального
    наказу з номером/датою поки не передбачено, лише сама дія
    архівування), решта студентів і самі групи архівуються.
    """
    group_ids = [int(x) for x in request.form.getlist('group_ids')]
    if not group_ids:
        flash("Оберіть хоча б одну групу.", "error")
        return redirect(url_for('admin.graduation'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    excluded_students = []
    for group_id in group_ids:
        students = conn.execute(
            "SELECT id FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0", (group_id,)
        ).fetchall()
        for s in students:
            included = request.form.get(f"include_{s['id']}") == '1'
            if not included:
                reason = (request.form.get(f"reason_{s['id']}") or '').strip()
                if not reason:
                    flash(f"Для виключеного студента (ID {s['id']}) не вказано причину.", "error")
                    conn.close()
                    return redirect(url_for('admin.graduation'))
                excluded_students.append((s['id'], group_id, reason))

    graduated_names = []
    for group_id in group_ids:
        group = conn.execute("SELECT name FROM groups WHERE id = ?", (group_id,)).fetchone()
        if not group:
            continue

        for student_id, g_id, reason in excluded_students:
            if g_id == group_id:
                conn.execute(
                    "INSERT INTO frozen_students (student_id, previous_group_id, order_id, reason, frozen_by) VALUES (?, ?, NULL, ?, ?)",
                    (student_id, group_id, reason, current_username())
                )
                conn.execute("UPDATE students SET group_id = NULL WHERE id = ?", (student_id,))

        # Архівуємо групу і тих студентів, хто в ній лишився (виключені
        # вже відв'язані від group_id вище і архівування їх не торкнеться)
        conn.execute("UPDATE groups SET archived = TRUE WHERE id = ?", (group_id,))
        conn.execute("UPDATE students SET archived = TRUE WHERE group_id = ?", (group_id,))
        graduated_names.append(group['name'])

    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"оформив випуск груп: {', '.join(graduated_names)}",
        details=f"заморожено студентів: {len(excluded_students)}"
    )
    flash(f"Випуск оформлено. Груп: {len(graduated_names)}, заморожено студентів: {len(excluded_students)}.", "success")
    return redirect(url_for('admin.courses'))


@admin_bp.route('/admin/expulsion', methods=['GET', 'POST'])
@permission_required('manage_expulsion')
def expulsion():
    """
    Крок 1 -> 2 "Наказу про відрахування". Студенти для наказу можуть
    прийти з двох джерел одразу - з активних груп (обираєте групи,
    потім студентів) і зі списку "Заморожені студенти" (обираєте
    напряму) - обидва потрапляють в один спільний список для перегляду.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        group_ids = request.form.getlist('group_ids')
        frozen_ids = request.form.getlist('frozen_ids')

        if not group_ids and not frozen_ids:
            flash("Оберіть хоча б одну групу або одного замороженого студента.", "error")
            conn.close()
            return redirect(url_for('admin.expulsion'))

        groups_with_students = []
        if group_ids:
            placeholders = ','.join('?' for _ in group_ids)
            selected_groups = conn.execute(f"""
                SELECT id, name FROM groups WHERE id IN ({placeholders}) AND archived = FALSE ORDER BY name
            """, group_ids).fetchall()
            for g in selected_groups:
                students = conn.execute("""
                    SELECT id, TRIM(last_name_UA || ' ' || first_name_UA) AS full_name
                    FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0
                    ORDER BY last_name_UA COLLATE UKRAINIAN
                """, (g['id'],)).fetchall()
                groups_with_students.append({'id': g['id'], 'name': g['name'], 'students': students})

        frozen_students_list = []
        if frozen_ids:
            placeholders = ','.join('?' for _ in frozen_ids)
            frozen_students_list = conn.execute(f"""
                SELECT f.id AS frozen_id, f.student_id, f.reason, f.previous_group_id,
                       TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS full_name
                FROM frozen_students f JOIN students s ON s.id = f.student_id
                WHERE f.id IN ({placeholders}) AND f.resolved_at IS NULL
            """, frozen_ids).fetchall()

        conn.close()
        return render_template(
            "expulsion_review.html",
            groups_with_students=groups_with_students,
            frozen_students_list=frozen_students_list,
        )

    groups = conn.execute("""
        SELECT id, name,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = groups.id AND COALESCE(s.archived, 0) = 0) AS student_count
        FROM groups WHERE archived = FALSE ORDER BY name
    """).fetchall()

    frozen = conn.execute("""
        SELECT f.id, f.reason, f.frozen_at,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS student_name,
               g.name AS previous_group_name
        FROM frozen_students f
        JOIN students s ON s.id = f.student_id
        LEFT JOIN groups g ON g.id = f.previous_group_id
        WHERE f.resolved_at IS NULL
        ORDER BY f.frozen_at
    """).fetchall()

    conn.close()
    return render_template("expulsion_select.html", groups=groups, frozen=frozen)


@admin_bp.route('/admin/expulsion/confirm', methods=['POST'])
@permission_required('manage_expulsion')
def expulsion_confirm():
    """Остаточне підтвердження наказу про відрахування - обробляє
    студентів з обох джерел (активні групи + заморожені) в одній операції."""
    order_number = (request.form.get('order_number') or '').strip()
    order_date = request.form.get('order_date')
    group_ids = [int(x) for x in request.form.getlist('group_ids')]
    frozen_ids = [int(x) for x in request.form.getlist('frozen_ids')]

    if not order_number or not order_date:
        flash("Заповніть номер і дату наказу.", "error")
        return redirect(url_for('admin.expulsion'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    # Студенти з активних груп - включаються за чекбоксом (за
    # замовчуванням НЕ позначені - відрахування виключення, а не
    # правило, на відміну від переведення/випуску).
    to_expel = []  # (student_id, previous_group_id, reason, frozen_id_to_resolve)
    for group_id in group_ids:
        students = conn.execute(
            "SELECT id FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0", (group_id,)
        ).fetchall()
        for s in students:
            included = request.form.get(f"include_{s['id']}") == '1'
            if included:
                reason = (request.form.get(f"reason_{s['id']}") or '').strip()
                if not reason:
                    flash(f"Для студента (ID {s['id']}) не вказано причину відрахування.", "error")
                    conn.close()
                    return redirect(url_for('admin.expulsion'))
                to_expel.append((s['id'], group_id, reason, None))

    # Заморожені студенти - усі обрані на кроці 1 включаються, з
    # причиною, яку можна було відредагувати на кроці перегляду.
    for frozen_id in frozen_ids:
        frozen_row = conn.execute(
            "SELECT student_id, previous_group_id FROM frozen_students WHERE id = ? AND resolved_at IS NULL", (frozen_id,)
        ).fetchone()
        if not frozen_row:
            continue
        reason = (request.form.get(f"frozen_reason_{frozen_id}") or '').strip()
        if not reason:
            flash(f"Для замороженого студента (запис {frozen_id}) не вказано причину відрахування.", "error")
            conn.close()
            return redirect(url_for('admin.expulsion'))
        to_expel.append((frozen_row['student_id'], frozen_row['previous_group_id'], reason, frozen_id))

    if not to_expel:
        flash("Не обрано жодного студента для відрахування.", "error")
        conn.close()
        return redirect(url_for('admin.expulsion'))

    scan_files = request.files.getlist('scan_files')

    cur = conn.execute(
        "INSERT INTO expulsion_orders (order_number, order_date, scan_file, created_by) VALUES (?, ?, NULL, ?)",
        (order_number, order_date, current_username())
    )
    order_id = cur.lastrowid
    saved_paths = save_multiple_attachments(conn, 'expulsion_order', order_id, scan_files, 'expulsion_orders', current_username())
    if saved_paths:
        conn.execute("UPDATE expulsion_orders SET scan_file = ? WHERE id = ?", (saved_paths[0], order_id))

    for student_id, previous_group_id, reason, frozen_id in to_expel:
        conn.execute(
            "INSERT INTO expulsion_order_students (order_id, student_id, previous_group_id, reason) VALUES (?, ?, ?, ?)",
            (order_id, student_id, previous_group_id, reason)
        )
        conn.execute("UPDATE students SET archived = TRUE, group_id = NULL WHERE id = ?", (student_id,))
        if frozen_id:
            conn.execute(
                "UPDATE frozen_students SET resolved_at = datetime('now', 'localtime'), resolution = ?, resolved_by = ? WHERE id = ?",
                (f"Відраховано наказом №{order_number} від {order_date}", current_username(), frozen_id)
            )

    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"наказ про відрахування №{order_number} від {order_date}",
        details=f"відраховано студентів: {len(to_expel)}"
    )
    flash(f"Наказ про відрахування оформлено. Відраховано студентів: {len(to_expel)}.", "success")
    return redirect(url_for('admin.frozen_students'))


@admin_bp.route('/admin/archive/<int:group_id>', methods=['POST'])
@permission_required('archive')
def archive_group(group_id):
    """Архівує групу разом з усіма її студентами (archived=TRUE)."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, start_year FROM groups WHERE id=?", (group_id,))
    group_row = cursor.fetchone()
    if not group_row:
        flash('Група не знайдена', 'error')
        conn.close()
        return redirect(url_for('admin.manage_groups'))

    cursor.execute("UPDATE groups SET archived = TRUE WHERE id=?", (group_id,))
    cursor.execute("UPDATE students SET archived = TRUE WHERE group_id=?", (group_id,))
    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"заархівував групу: {group_row['name']} ({group_row['start_year']}) (ID {group_id})"
    )
    flash('Групу успішно заархівовано', 'success')
    return redirect(url_for('admin.manage_groups'))


@admin_bp.route('/admin/unarchive_group/<int:group_id>', methods=['POST'])
@permission_required('archive')
def unarchive_group(group_id):
    """Повертає архівну групу разом зі студентами назад в активні (archived=FALSE)."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, start_year FROM groups WHERE id=? AND archived=TRUE", (group_id,))
    group_row = cursor.fetchone()
    if not group_row:
        flash('Архівна група не знайдена', 'error')
        conn.close()
        return redirect(url_for('admin.archive'))

    cursor.execute("UPDATE groups SET archived = FALSE WHERE id=?", (group_id,))
    cursor.execute("UPDATE students SET archived = FALSE WHERE group_id=?", (group_id,))
    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"розархівував групу: {group_row['name']} ({group_row['start_year']}) (ID {group_id})"
    )
    flash('Групу успішно розархівовано', 'success')
    return redirect(url_for('admin.archive'))


@admin_bp.route('/admin/archive')
@permission_required('archive')
def archive():
    """Показує список заархівованих груп та їхніх студентів, згрупованих за роком випуску і ступенем навчання."""
    conn = get_db()
    conn.row_factory = sqlite3.Row

    groups = conn.execute("""
        SELECT g.id, g.name, g.start_year, g.study_form, g.program_credits, g.degree_level,
               g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS display_name,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = g.id AND s.archived = TRUE) AS student_count
        FROM groups g WHERE g.archived = TRUE ORDER BY g.start_year DESC, g.degree_level, g.name
    """).fetchall()
    groups = [dict(g) for g in groups]
    for g in groups:
        g['end_year'] = g['start_year'] + compute_program_total_years(g['degree_level'], g['program_credits'])

    # Порядок ступенів - за їхнім id в каталозі degree_levels (природна
    # ієрархія: молодший бакалавр -> бакалавр -> магістр -> доктор
    # філософії), той самий принцип, що й на дошці "Курси". Ступені,
    # яких немає в каталозі, йдуть в кінець, за абеткою.
    degree_order = {
        row['name_ua']: row['id']
        for row in conn.execute("SELECT id, name_ua FROM degree_levels").fetchall()
    }

    # Групуємо для компактнішого відображення: РІК ВИПУСКУ -> ступінь
    # -> список груп. Рік випуску - а не рік вступу, бо саме він
    # відповідає на природне питання "хто випустився у 2026-му" (а не
    # "хто вступив у 2022-му, і хай кожен сам порахує, коли випустився").
    groups_by_year_raw = {}
    for g in groups:
        groups_by_year_raw.setdefault(g['end_year'], {}).setdefault(g['degree_level'], []).append(g)

    groups_by_year = {}
    for year in sorted(groups_by_year_raw.keys(), reverse=True):
        degree_dict = groups_by_year_raw[year]
        groups_by_year[year] = {
            degree_level: degree_dict[degree_level]
            for degree_level in sorted(degree_dict.keys(), key=lambda d: (degree_order.get(d, 999), d or ''))
        }

    students_by_group = {}
    for group in groups:
        students = conn.execute("""
            SELECT id, last_name_UA, first_name_UA, birth_date, program_credits_override
            FROM students WHERE group_id=? AND archived=TRUE ORDER BY last_name_UA
        """, (group['id'],)).fetchall()
        students_by_group[group['id']] = students

    # Відраховані студенти - окремо від архівних груп, бо наказ про
    # відрахування може стосуватись студента з ГРУПИ, яка сама
    # лишається активною (лише сам студент стає архівним) - такий
    # студент інакше був би не видно ніде: ні в активному списку групи
    # (бо archived=TRUE), ні тут вище (бо його група не архівна).
    expelled_students = conn.execute("""
        SELECT eos.student_id, eos.reason, eo.order_number, eo.order_date,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS student_name,
               g.name AS previous_group_name
        FROM expulsion_order_students eos
        JOIN expulsion_orders eo ON eo.id = eos.order_id
        JOIN students s ON s.id = eos.student_id
        LEFT JOIN groups g ON g.id = eos.previous_group_id
        ORDER BY eo.order_date DESC
    """).fetchall()

    conn.close()
    log_action(current_username(), "переглянув список архівних груп")
    return render_template('archive.html', groups=groups, groups_by_year=groups_by_year, students_by_group=students_by_group, expelled_students=expelled_students)
