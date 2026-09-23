"""
routes/admin.py
================
Усі адміністративні розділи додатку: групи, предмети, практики/курсові/
атестації, документи про освіту, дипломи, акредитації, користувачі та
їхні права, журнал дій, імпорт з Excel та масова генерація .docx-
документів. Доступ до кожного маршруту обмежений через
`@permission_required('<назва_дозволу>')` (routes/utils.py).

Опис призначення кожної функції - див. докстрінг під відповідним
`def ...` нижче, або підсумкову таблицю в FUNCTIONS.md.
"""

from flask import Blueprint, render_template, request, redirect, url_for, flash, session, send_file
from datetime import datetime
import os
import sqlite3
import zipfile
import io
from werkzeug.security import generate_password_hash
from routes.db import get_db
from routes.utils import log_action, permission_required, is_student_on_reduced_program, save_multiple_attachments, get_attachments, filter_students_for_item
from routes.gen_docx import gen_doc
from routes import office_editor
from routes.utils import logger
from routes.helpers import current_username, sort_ukrainian
import openpyxl
from werkzeug.utils import secure_filename
import pandas as pd
import json
import uuid
import re
import threading
from openpyxl import load_workbook
from rapidfuzz import process, fuzz
from deep_translator import GoogleTranslator
from openpyxl.utils.exceptions import InvalidFileException
import time


from routes.utils import get_templates_with_metadata, TEMPLATE_FOLDER

translator = GoogleTranslator(source="auto", target="en")
translation_cache = {}

admin_bp = Blueprint('admin', __name__)

# Фіксовані пари УКР->EN (значень лише кілька, тому окремої таблиці в
# базі не потрібно, на відміну від спеціальностей) - джерело істини
# для полів "Ступінь"/"Найменування та статус закладу", щоб форма не
# могла надіслати неузгоджений англійський варіант.
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


INSTITUTION_NAME_STATUS_EN = {
    'Приватний вищий навчальний заклад «Європейський університет». Приватна форма власності. Міністерство освіти і науки України. Ліцензія серія ВО № 00228-022801 від 15/05/2017.':
        "Private Higher Educational Institution 'European University'. Private. Ministry of Education and Science of Ukraine. License series BO № 00228-022801 dated 15/05/2017.",
    'Львівська філія Приватного вищого навчального закладу «Європейський університет». Приватна форма власності. Міністерство освіти і науки України. Ліцензія серія ВО № 00228-022801 від 15/05/2017.':
        'Lviv Branch of Private Higher Education Establishment «European University». Private. Ministry of  Education and  Science of Ukraine. License series ВO № 00228-022801 from 15/05/2017.',
}

PERMISSIONS = [
    'manage_users',
    'view_logs',
    'group_export',
    'import_from_excel',
    'manage_education_documents',
    'manage_passport_documents',
    'import_passport_documents',
    'study_periods',
    'manage_groups',
    'manage_subjects',
    'manage_activities',
    'import_subjects',
    'archive',
    'manage_students',
    'manage_accreditations',
    'manage_diplomas',
    'import_education_docs',
    'manage_templates',
    'import_grades',
    'analytics',
    'manage_specialties',
    'manage_degree_levels',
    'manage_educational_programs',
    'manage_qualification_names',
    'manage_courses',
    'manage_frozen_students',
    'manage_expulsion',
    'manage_licenses',
    'manage_license_transfer',
    'license_report',
    'manage_pending_students'
]






@admin_bp.route('/admin/study_periods/bulk_assign', methods=['GET', 'POST'])
@permission_required('study_periods')
def manage_study_periods_bulk_assign():
    """Масове призначення періоду навчання (філія, дати, група) одразу декільком студентам — обраній групі цілком та/або довільному списку студентів."""
    db = get_db()
    cursor = db.cursor()

    # -------------------- POST --------------------
    if request.method == 'POST':
        mode = request.form.get('mode')
        group_id = request.form.get('group_id', type=int)
        student_ids_raw = request.form.get('student_ids_hidden', '')
        extra_student_ids = [int(x) for x in student_ids_raw.split(',') if x.strip().isdigit()]

        target_student_ids = set(extra_student_ids)

        if group_id and mode in ('whole_group', 'some_from_group'):
            cursor.execute("""
                SELECT id FROM students WHERE group_id = ? AND archived = FALSE
            """, (group_id,))
            target_student_ids.update(r['id'] for r in cursor.fetchall())

        if not target_student_ids:
            flash('Не обрано жодного студента', 'danger')
            return redirect(url_for('admin.manage_study_periods_bulk_assign'))

        filiya = (request.form.get('filiya') or '').strip()
        filiya_en = (request.form.get('filiya_en') or '').strip() or None
        group_name = (request.form.get('group_name') or '').strip() or None
        start_date = (request.form.get('start_date') or '').strip() or None
        end_date = (request.form.get('end_date') or '').strip() or None
        period_order = request.form.get('period_order', type=int) or 0
        note = (request.form.get('note') or '').strip() or None

        if not filiya:
            flash('Потрібно вказати філію', 'danger')
            return redirect(url_for('admin.manage_study_periods_bulk_assign',
                                  group_id=group_id or '',
                                  student_ids=','.join(map(str, extra_student_ids)),
                                  mode=mode))

        added = 0
        try:
            for student_id in target_student_ids:
                cursor.execute("""
                    INSERT INTO student_study_periods
                        (student_id, filiya, filiya_en, group_name, start_date, end_date, period_order, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (student_id, filiya, filiya_en, group_name, start_date, end_date, period_order, note))
                added += 1

            db.commit()
            log_action(
                current_username(),
                f"масово присвоїв період навчання ({filiya}) для {added} студент(ів)"
            )
            flash(f'Період навчання успішно додано {added} студент(ам)', 'success')
        except sqlite3.Error as e:
            db.rollback()
            logger.error(f"Помилка масового призначення періоду навчання (group_id={group_id}): {e}", exc_info=True)
            flash(f'Помилка збереження: {e}', 'danger')

        return redirect(url_for('admin.manage_study_periods_bulk_assign',
                              group_id=group_id or '',
                              student_ids=','.join(map(str, extra_student_ids)),
                              mode=mode))

    # -------------------- GET --------------------
    cursor.execute("SELECT id, name FROM groups WHERE archived = FALSE ORDER BY name")
    groups = cursor.fetchall()

    # Всі студенти з інформацією про групу (для JS-фільтрації)
    cursor.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA, 
               s.group_id, g.name AS group_name
        FROM students s
        LEFT JOIN groups g ON s.group_id = g.id
        WHERE s.archived = FALSE
        ORDER BY s.last_name_UA, s.first_name_UA
    """)
    all_students = cursor.fetchall()

    # Параметри
    mode = request.args.get('mode', 'any_students')
    group_id = request.args.get('group_id', type=int)
    student_ids_param = request.args.get('student_ids', '')
    extra_student_ids = [int(x) for x in student_ids_param.split(',') if x.strip().isdigit()]

    target_student_ids = set(extra_student_ids)

    if group_id and mode in ('whole_group', 'some_from_group'):
        cursor.execute("""
            SELECT id FROM students WHERE group_id = ? AND archived = FALSE
        """, (group_id,))
        target_student_ids.update(r['id'] for r in cursor.fetchall())

    # Прев'ю
    target_students = []
    if target_student_ids:
        placeholders = ','.join('?' for _ in target_student_ids)
        cursor.execute(f"""
            SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA, g.name AS group_name
            FROM students s
            LEFT JOIN groups g ON s.group_id = g.id
            WHERE s.id IN ({placeholders})
            ORDER BY s.last_name_UA, s.first_name_UA
        """, tuple(target_student_ids))
        target_students_rows = cursor.fetchall()

        cursor.execute(f"""
            SELECT student_id, filiya, filiya_en, group_name, start_date, end_date, period_order, note
            FROM student_study_periods
            WHERE student_id IN ({placeholders})
            ORDER BY student_id, period_order ASC, start_date ASC
        """, tuple(target_student_ids))
        periods_rows = cursor.fetchall()

        periods_by_student = {}
        for p in periods_rows:
            periods_by_student.setdefault(p['student_id'], []).append(p)

        for s in target_students_rows:
            s_dict = dict(s)
            s_dict['existing_periods'] = periods_by_student.get(s['id'], [])
            target_students.append(s_dict)

    return render_template(
        'manage_study_periods_bulk_assign.html',
        groups=groups,
        all_students=all_students,
        selected_group_id=group_id,
        selected_student_ids=extra_student_ids,
        target_students=target_students,
        selected_mode=mode
    )


@admin_bp.route('/admin/manage_diplomas', methods=['GET', 'POST'])
@permission_required('manage_diplomas')
def manage_diplomas():
    """Перегляд і масове збереження номерів диплома/додатка для всіх студентів обраної групи."""
    conn = get_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    if request.method == "POST":
        group_id = request.form.get("group_id")
        cursor.execute("SELECT s.id FROM students s WHERE s.group_id = ?", (group_id,))
        students = cursor.fetchall()

        for student in students:
            student_id = student['id']
            diploma_number = request.form.get(f'diploma_number_{student_id}', '').strip()
            appendix_number = request.form.get(f'appendix_number_{student_id}', '').strip()
            if diploma_number:
                diploma_number = diploma_number.zfill(6)
            cursor.execute("SELECT id FROM diplomas WHERE student_id=?", (student_id,))
            exists = cursor.fetchone()
            if exists:
                cursor.execute("""
                    UPDATE diplomas SET diploma_number=?, appendix_number=? WHERE student_id=?
                """, (diploma_number, appendix_number, student_id))
            else:
                cursor.execute("""
                    INSERT INTO diplomas(student_id, diploma_number, appendix_number) VALUES (?, ?, ?)
                """, (student_id, diploma_number, appendix_number))

        conn.commit()

        group_row = cursor.execute("SELECT name FROM groups WHERE id=?", (group_id,)).fetchone()
        group_name = group_row['name'] if group_row else f"ID {group_id}"
        log_action(
            current_username(),
            f"зберіг дипломи для групи: {group_name} (ID {group_id})",
            details=f"студентів оброблено: {len(students)}"
        )

        flash("Дані збережено")
        return redirect(url_for('admin.manage_diplomas', group_id=group_id))

    cursor.execute("SELECT id, name, start_year FROM groups WHERE archived = FALSE ORDER BY name")
    groups = cursor.fetchall()
    selected_group = request.args.get("group_id")
    if selected_group is not None:
        selected_group = int(selected_group)

    students = []
    if selected_group:
        cursor.execute("""
            SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
                   d.diploma_number, d.appendix_number
            FROM students s
            LEFT JOIN diplomas d ON s.id = d.student_id
            WHERE s.group_id = ?
        """, (selected_group,))
        students = cursor.fetchall()
        students = sort_ukrainian(
            students,
            key_func=lambda s: f"{s['last_name_UA']} {s['first_name_UA']} {s['middle_name_UA']}"
        )

    return render_template("manage_diplomas.html", groups=groups, students=students, selected_group=selected_group)


@admin_bp.route('/admin/manage_accreditations', methods=['GET', 'POST'])
@permission_required('manage_accreditations')
def manage_accreditations():
    """CRUD-сторінка довідників акредитацій (ступінь + спеціальність + текст українською/англійською) для підстановки в документи."""
    conn = get_db()
    cursor = conn.cursor()

    if request.method == 'POST' and 'add' in request.form:
        degree = request.form.get('degree')
        specialty = request.form.get('specialty')
        text_ua = request.form.get('text_ua')
        text_en = request.form.get('text_en')
        cursor.execute("""
            INSERT INTO accreditations (degree, specialty, text_ua, text_en) VALUES (?, ?, ?, ?)
        """, (degree, specialty, text_ua, text_en))
        conn.commit()
        log_action(
            current_username(),
            f"додав акредитацію: {degree} / {specialty}"
        )
        return redirect(url_for('admin.manage_accreditations'))

    if request.method == 'POST' and 'edit' in request.form:
        acc_id = request.form.get('id')
        degree = request.form.get('degree')
        specialty = request.form.get('specialty')
        text_ua = request.form.get('text_ua')
        text_en = request.form.get('text_en')
        cursor.execute("""
            UPDATE accreditations SET degree=?, specialty=?, text_ua=?, text_en=? WHERE id=?
        """, (degree, specialty, text_ua, text_en, acc_id))
        conn.commit()
        log_action(
            current_username(),
            f"редагував акредитацію ID {acc_id}: {degree} / {specialty}"
        )
        return redirect(url_for('admin.manage_accreditations'))

    if request.method == 'POST' and 'delete' in request.form:
        acc_id = request.form.get('id')
        row = cursor.execute("SELECT degree, specialty FROM accreditations WHERE id=?", (acc_id,)).fetchone()
        cursor.execute("DELETE FROM accreditations WHERE id=?", (acc_id,))
        conn.commit()
        log_action(
            current_username(),
            f"ВИДАЛИВ акредитацію ID {acc_id}: {row[0] if row else ''} / {row[1] if row else ''}"
        )
        return redirect(url_for('admin.manage_accreditations'))

    cursor.execute("SELECT id, degree, specialty, text_ua, text_en FROM accreditations ORDER BY degree, specialty")
    accreditations = cursor.fetchall()
    cursor.execute("SELECT DISTINCT specialty FROM groups WHERE archived = FALSE ORDER BY specialty")
    groups = cursor.fetchall()

    return render_template('manage_accreditations.html', accreditations=accreditations, groups=groups)


@admin_bp.route('/admin/manage_education_documents', methods=['GET', 'POST'])
@permission_required('manage_education_documents')
def manage_education_documents():
    """CRUD-сторінка документів про освіту студентів (атестат/диплом попереднього рівня) разом із додатковими даними про визнання іноземного документа (foreign_education_docs)."""
    db = get_db()
    cursor = db.cursor()

    cursor.execute("""
        SELECT id, last_name_UA, first_name_UA FROM students
        WHERE archived = FALSE ORDER BY last_name_UA, first_name_UA
    """)
    students = cursor.fetchall()

    cursor.execute("SELECT id, name FROM groups WHERE archived = FALSE ORDER BY name")
    groups = cursor.fetchall()

    # --- Отримуємо студента, якщо перейшли з картки ---
    student = None
    student_id_param = request.args.get('student_id', type=int)
    if student_id_param:
        cursor.execute("""
            SELECT id, last_name_UA, first_name_UA, middle_name_UA, group_id
            FROM students WHERE id = ?
        """, (student_id_param,))
        srow = cursor.fetchone()
        if srow:
            student = dict(srow)
            cursor.execute("SELECT name FROM groups WHERE id = ?", (student['group_id'],))
            grow = cursor.fetchone()
            student['group_name'] = grow['name'] if grow else ''

    selected_group_id = request.args.get('group_id', type=int)
    if not selected_group_id and student:
        selected_group_id = student['group_id']

    students_without_docs = []
    if selected_group_id:
        cursor.execute("""
            SELECT s.id, s.last_name_UA, s.first_name_UA
            FROM students s
            WHERE s.group_id = ? AND s.archived = FALSE
              AND s.id NOT IN (SELECT student_id FROM education_documents)
            ORDER BY s.last_name_UA, s.first_name_UA
        """, (selected_group_id,))
        students_without_docs = cursor.fetchall()

    # Основний запит документів
    cursor.execute("""
        SELECT g.id AS group_id, g.name AS group_name, s.id AS student_id,
               s.last_name_UA, s.first_name_UA,
               ed.id AS doc_id, ed.document_type, ed.document_type_en, ed.document_number,
               ed.institution_name, ed.institution_name_en, ed.country, ed.country_en, ed.completion_date,
               fed.reference_number, fed.reference_institution, fed.reference_institution_en,
               fed.reference_country, fed.reference_country_en, fed.reference_issue_date,
               fed.recognition_certificate_number, fed.recognition_issuer,
               fed.recognition_issuer_en, fed.recognition_date
        FROM education_documents ed
        INNER JOIN students s ON ed.student_id = s.id
        INNER JOIN groups g ON s.group_id = g.id
        LEFT JOIN foreign_education_docs fed ON ed.id = fed.education_doc_id
        WHERE s.archived = FALSE AND g.archived = FALSE
        ORDER BY g.name, s.last_name_UA, s.first_name_UA, ed.id
    """)
    rows = cursor.fetchall()

    documents_by_group = {}
    for row in rows:
        gid = row['group_id']
        if gid not in documents_by_group:
            documents_by_group[gid] = {'group_name': row['group_name'], 'docs': []}
        doc = dict(row)
        doc['attachments'] = get_attachments(db, 'education_document', doc['doc_id'])
        documents_by_group[gid]['docs'].append(doc)

    sorted_documents_by_group = sorted(documents_by_group.items(), key=lambda x: x[1]['group_name'])

    # Перевіряємо doc_id для студента
    student_doc_id = None
    if student:
        for gid, gdata in documents_by_group.items():
            for doc in gdata['docs']:
                if doc['student_id'] == student['id']:
                    student_doc_id = doc['doc_id']
                    break
            if student_doc_id:
                break

    # ====================== POST ======================
    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'delete':
            doc_id = request.form.get('doc_id')
            try:
                # Спершу самі файли сканів з диска, а тоді записи -
                # інакше видалення документа лишало б "осиротілі" файли.
                for att in get_attachments(db, 'education_document', doc_id):
                    att_path = os.path.join('static', att['file_path'])
                    if os.path.exists(att_path):
                        os.remove(att_path)
                cursor.execute("DELETE FROM attachments WHERE entity_type='education_document' AND entity_id=?", (doc_id,))
                cursor.execute("DELETE FROM foreign_education_docs WHERE education_doc_id = ?", (doc_id,))
                cursor.execute("DELETE FROM education_documents WHERE id = ?", (doc_id,))
                db.commit()
                log_action(current_username(), f"ВИДАЛИВ документ про освіту ID {doc_id}")
                flash('Документ успішно видалено', 'success')
            except sqlite3.Error as e:
                db.rollback()
                logger.error(f"Помилка видалення документа про освіту (ID {doc_id}): {e}", exc_info=True)
                flash(f'Помилка видалення: {e}', 'danger')

        elif action == 'delete_attachment':
            attachment_id = request.form.get('attachment_id')
            try:
                att = cursor.execute("SELECT file_path FROM attachments WHERE id=?", (attachment_id,)).fetchone()
                if att:
                    att_path = os.path.join('static', att['file_path'])
                    if os.path.exists(att_path):
                        os.remove(att_path)
                    cursor.execute("DELETE FROM attachments WHERE id=?", (attachment_id,))
                    db.commit()
                    flash('Скан видалено', 'success')
                else:
                    flash('Файл не знайдено', 'error')
            except Exception as e:
                db.rollback()
                logger.error(f"Помилка видалення вкладення (ID {attachment_id}): {e}", exc_info=True)
                flash(f'Помилка видалення файлу: {e}', 'danger')
            return redirect(url_for('admin.manage_education_documents', group_id=selected_group_id))

        elif action == 'edit':
            doc_id = request.form.get('doc_id')
            student_id = request.form.get('student_id')

            country = request.form.get('country') or ''
            country_en = request.form.get('country_en') or ''

            try:
                cursor.execute("SELECT student_id FROM education_documents WHERE id = ?", (doc_id,))
                existing = cursor.fetchone()
                if not existing:
                    flash('Документ не знайдено', 'danger')
                    return redirect(url_for('admin.manage_education_documents', group_id=selected_group_id))

                if student_id:
                    cursor.execute("SELECT id FROM students WHERE id = ? AND archived = FALSE", (student_id,))
                    if not cursor.fetchone():
                        flash('Обраний студент не існує або заархівований', 'danger')
                        return redirect(url_for('admin.manage_education_documents', group_id=selected_group_id))
                else:
                    student_id = existing[0]

                # Оновлення основного документа (зберігаємо country як ввели)
                cursor.execute("""
                    UPDATE education_documents SET
                        student_id=?, document_type=?, document_type_en=?, document_number=?,
                        institution_name=?, institution_name_en=?, country=?, country_en=?, completion_date=?
                    WHERE id=?
                """, (student_id, request.form.get('document_type'), request.form.get('document_type_en'),
                      request.form.get('document_number'), request.form.get('institution_name'),
                      request.form.get('institution_name_en'), country, country_en,
                      request.form.get('completion_date'), doc_id))

                # Foreign fields
                reference_number = request.form.get('reference_number') or None
                reference_institution = request.form.get('reference_institution') or None
                reference_institution_en = request.form.get('reference_institution_en') or None
                reference_country = request.form.get('reference_country') or None
                reference_country_en = request.form.get('reference_country_en') or None
                reference_issue_date = request.form.get('reference_issue_date') or None
                recognition_certificate_number = request.form.get('recognition_certificate_number') or None
                recognition_issuer = request.form.get('recognition_issuer') or None
                recognition_issuer_en = request.form.get('recognition_issuer_en') or None
                recognition_date = request.form.get('recognition_date') or None

                # Рішення приймаємо тільки на основі того, чи заповнені самі поля довідки/визнання,
                # а НЕ на основі значення country. Це прибирає випадкове видалення довідки
                # через неправильно передане/незмінене поле "Країна".
                has_foreign_data = any([
                    reference_number, reference_institution, reference_country,
                    reference_issue_date, recognition_certificate_number,
                    recognition_issuer, recognition_date
                ])

                cursor.execute("SELECT id FROM foreign_education_docs WHERE education_doc_id = ?", (doc_id,))
                foreign_exists = cursor.fetchone()

                if foreign_exists:
                    if has_foreign_data:
                        cursor.execute("""
                            UPDATE foreign_education_docs SET
                                reference_number=?, reference_institution=?, reference_institution_en=?,
                                reference_country=?, reference_country_en=?, reference_issue_date=?,
                                recognition_certificate_number=?, recognition_issuer=?,
                                recognition_issuer_en=?, recognition_date=?
                            WHERE education_doc_id=?
                        """, (reference_number, reference_institution, reference_institution_en,
                              reference_country, reference_country_en, reference_issue_date,
                              recognition_certificate_number, recognition_issuer,
                              recognition_issuer_en, recognition_date, doc_id))
                    else:
                        # Видаляємо запис тільки якщо користувач сам очистив усі поля довідки/визнання
                        cursor.execute("DELETE FROM foreign_education_docs WHERE education_doc_id = ?", (doc_id,))
                else:
                    if has_foreign_data:
                        cursor.execute("""
                            INSERT INTO foreign_education_docs (
                                education_doc_id, reference_number, reference_institution, reference_institution_en,
                                reference_country, reference_country_en, reference_issue_date,
                                recognition_certificate_number, recognition_issuer, recognition_issuer_en, recognition_date
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (doc_id, reference_number, reference_institution, reference_institution_en,
                              reference_country, reference_country_en, reference_issue_date,
                              recognition_certificate_number, recognition_issuer, recognition_issuer_en, recognition_date))

                db.commit()

                new_scans = [f for f in request.files.getlist('document_scans') if f and f.filename]
                if new_scans:
                    save_multiple_attachments(db, 'education_document', doc_id, new_scans, 'education_documents', current_username())
                    db.commit()

                log_action(
                    current_username(),
                    f"редагував документ про освіту ID {doc_id}",
                    details=f"країна: {country}"
                )
                flash('Документ успішно оновлено', 'success')

            except sqlite3.Error as e:
                db.rollback()
                logger.error(f"Помилка БД при редагуванні документа про освіту (ID {doc_id}): {e}", exc_info=True)
                flash(f'Помилка при редагуванні: {e}', 'danger')
            except Exception as e:
                db.rollback()
                logger.error(f"Непередбачена помилка при редагуванні документа про освіту (ID {doc_id}): {e}", exc_info=True)
                flash(f'Непередбачена помилка: {str(e)}', 'danger')

            return redirect(url_for('admin.manage_education_documents', group_id=selected_group_id))

        # === Додавання нового документа ===
        else:
            country = request.form.get('country') or ''
            country_en = request.form.get('country_en') or ''

            student_id = request.form.get('student_id')
            document_type = request.form.get('document_type')
            document_type_en = request.form.get('document_type_en')
            document_number = request.form.get('document_number')
            institution_name = request.form.get('institution_name')
            institution_name_en = request.form.get('institution_name_en')
            completion_date = request.form.get('completion_date')

            reference_number = request.form.get('reference_number') or None
            reference_institution = request.form.get('reference_institution') or None
            reference_institution_en = request.form.get('reference_institution_en') or None
            reference_country = request.form.get('reference_country') or None
            reference_country_en = request.form.get('reference_country_en') or None
            reference_issue_date = request.form.get('reference_issue_date') or None
            recognition_certificate_number = request.form.get('recognition_certificate_number') or None
            recognition_issuer = request.form.get('recognition_issuer') or None
            recognition_issuer_en = request.form.get('recognition_issuer_en') or None
            recognition_date = request.form.get('recognition_date') or None

            try:
                if not student_id:
                    raise ValueError("Не обрано студента")

                cursor.execute("""
                    INSERT INTO education_documents (
                        student_id, document_type, document_type_en, document_number,
                        institution_name, institution_name_en, country, country_en, completion_date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (student_id, document_type, document_type_en, document_number,
                      institution_name, institution_name_en, country, country_en, completion_date))

                education_doc_id = cursor.lastrowid

                has_foreign_data = any([
                    reference_number, reference_institution, reference_country,
                    reference_issue_date, recognition_certificate_number,
                    recognition_issuer, recognition_date
                ])

                if has_foreign_data:
                    cursor.execute("""
                        INSERT INTO foreign_education_docs (
                            education_doc_id, reference_number, reference_institution, reference_institution_en,
                            reference_country, reference_country_en, reference_issue_date,
                            recognition_certificate_number, recognition_issuer, recognition_issuer_en, recognition_date
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (education_doc_id, reference_number, reference_institution, reference_institution_en,
                          reference_country, reference_country_en, reference_issue_date,
                          recognition_certificate_number, recognition_issuer, recognition_issuer_en, recognition_date))

                db.commit()
                log_action(
                    current_username(),
                    f"додав документ про освіту: {document_type} №{document_number}",
                    details=f"країна: {country}"
                )

                new_scans = [f for f in request.files.getlist('document_scans') if f and f.filename]
                if new_scans:
                    save_multiple_attachments(db, 'education_document', education_doc_id, new_scans, 'education_documents', current_username())
                    db.commit()

                flash('Документ успішно додано', 'success')

            except (sqlite3.Error, ValueError) as e:
                db.rollback()
                flash(f'Помилка при додаванні: {str(e)}', 'danger')

        return redirect(url_for('admin.manage_education_documents', group_id=selected_group_id))

    cursor.close()
    return render_template(
        'manage_education_documents.html',
        groups=groups,
        selected_group_id=selected_group_id,
        students_without_docs=students_without_docs,
        documents_by_group=sorted_documents_by_group,
        students=students,
        student=student,
        student_doc_id=student_doc_id
    )
    
@admin_bp.route('/admin/manage_passport_documents', methods=['GET', 'POST'])
@permission_required('manage_passport_documents')
def manage_passport_documents():
    """CRUD-сторінка паспортних даних студентів (паспорт-книжка або
    ID-картка) - за тим самим зразком, що й документи про освіту, але
    без підтаблиці "іноземний документ" і з лімітом 5 сканів на запис."""
    db = get_db()
    cursor = db.cursor()
    MAX_FILES = 5

    cursor.execute("""
        SELECT id, last_name_UA, first_name_UA FROM students
        WHERE archived = FALSE ORDER BY last_name_UA, first_name_UA
    """)
    students = cursor.fetchall()

    cursor.execute("SELECT id, name FROM groups WHERE archived = FALSE ORDER BY name")
    groups = cursor.fetchall()

    student = None
    student_id_param = request.args.get('student_id', type=int)
    if student_id_param:
        cursor.execute("""
            SELECT id, last_name_UA, first_name_UA, middle_name_UA, group_id
            FROM students WHERE id = ?
        """, (student_id_param,))
        srow = cursor.fetchone()
        if srow:
            student = dict(srow)
            cursor.execute("SELECT name FROM groups WHERE id = ?", (student['group_id'],))
            grow = cursor.fetchone()
            student['group_name'] = grow['name'] if grow else ''

    selected_group_id = request.args.get('group_id', type=int)
    if not selected_group_id and student:
        selected_group_id = student['group_id']

    students_without_docs = []
    if selected_group_id:
        cursor.execute("""
            SELECT s.id, s.last_name_UA, s.first_name_UA
            FROM students s
            WHERE s.group_id = ? AND s.archived = FALSE
              AND s.id NOT IN (SELECT student_id FROM passport_documents)
            ORDER BY s.last_name_UA, s.first_name_UA
        """, (selected_group_id,))
        students_without_docs = cursor.fetchall()

    cursor.execute("""
        SELECT g.id AS group_id, g.name AS group_name, s.id AS student_id,
               s.last_name_UA, s.first_name_UA,
               pd.id AS doc_id, pd.document_type, pd.series, pd.number,
               pd.issued_by, pd.issue_date, pd.valid_until, pd.unique_number
        FROM passport_documents pd
        INNER JOIN students s ON pd.student_id = s.id
        INNER JOIN groups g ON s.group_id = g.id
        WHERE s.archived = FALSE AND g.archived = FALSE
        ORDER BY g.name, s.last_name_UA, s.first_name_UA, pd.id
    """)
    rows = cursor.fetchall()

    documents_by_group = {}
    for row in rows:
        gid = row['group_id']
        if gid not in documents_by_group:
            documents_by_group[gid] = {'group_name': row['group_name'], 'docs': []}
        doc = dict(row)
        doc['attachments'] = get_attachments(db, 'passport_document', doc['doc_id'])
        documents_by_group[gid]['docs'].append(doc)

    sorted_documents_by_group = sorted(documents_by_group.items(), key=lambda x: x[1]['group_name'])

    student_doc_id = None
    if student:
        for gid, gdata in documents_by_group.items():
            for doc in gdata['docs']:
                if doc['student_id'] == student['id']:
                    student_doc_id = doc['doc_id']
                    break
            if student_doc_id:
                break

    # ====================== POST ======================
    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'delete':
            doc_id = request.form.get('doc_id')
            try:
                for att in get_attachments(db, 'passport_document', doc_id):
                    att_path = os.path.join('static', att['file_path'])
                    if os.path.exists(att_path):
                        os.remove(att_path)
                cursor.execute("DELETE FROM attachments WHERE entity_type='passport_document' AND entity_id=?", (doc_id,))
                cursor.execute("DELETE FROM passport_documents WHERE id = ?", (doc_id,))
                db.commit()
                log_action(current_username(), f"ВИДАЛИВ паспортні дані ID {doc_id}")
                flash('Запис успішно видалено', 'success')
            except sqlite3.Error as e:
                db.rollback()
                logger.error(f"Помилка видалення паспортних даних (ID {doc_id}): {e}", exc_info=True)
                flash(f'Помилка видалення: {e}', 'danger')

        elif action == 'delete_attachment':
            attachment_id = request.form.get('attachment_id')
            try:
                att = cursor.execute("SELECT file_path FROM attachments WHERE id=?", (attachment_id,)).fetchone()
                if att:
                    att_path = os.path.join('static', att['file_path'])
                    if os.path.exists(att_path):
                        os.remove(att_path)
                    cursor.execute("DELETE FROM attachments WHERE id=?", (attachment_id,))
                    db.commit()
                    flash('Скан видалено', 'success')
                else:
                    flash('Файл не знайдено', 'error')
            except Exception as e:
                db.rollback()
                logger.error(f"Помилка видалення вкладення (ID {attachment_id}): {e}", exc_info=True)
                flash(f'Помилка видалення файлу: {e}', 'danger')
            return redirect(url_for('admin.manage_passport_documents', group_id=selected_group_id))

        elif action == 'edit':
            doc_id = request.form.get('doc_id')
            student_id = request.form.get('student_id')
            document_type = request.form.get('document_type')

            try:
                existing = cursor.execute("SELECT student_id FROM passport_documents WHERE id = ?", (doc_id,)).fetchone()
                if not existing:
                    flash('Запис не знайдено', 'danger')
                    return redirect(url_for('admin.manage_passport_documents', group_id=selected_group_id))

                if student_id:
                    if not cursor.execute("SELECT id FROM students WHERE id = ? AND archived = FALSE", (student_id,)).fetchone():
                        flash('Обраний студент не існує або заархівований', 'danger')
                        return redirect(url_for('admin.manage_passport_documents', group_id=selected_group_id))
                else:
                    student_id = existing[0]

                if document_type not in ('Паспорт (книжка)', 'ID-картка'):
                    raise ValueError("Невірний тип документа")

                existing_count = cursor.execute(
                    "SELECT COUNT(*) FROM attachments WHERE entity_type='passport_document' AND entity_id=?", (doc_id,)
                ).fetchone()[0]
                new_scans = [f for f in request.files.getlist('document_scans') if f and f.filename]
                if existing_count + len(new_scans) > MAX_FILES:
                    flash(f'Забагато файлів - максимум {MAX_FILES} на запис (вже є {existing_count}).', 'error')
                    return redirect(url_for('admin.manage_passport_documents', group_id=selected_group_id))

                cursor.execute("""
                    UPDATE passport_documents SET
                        student_id=?, document_type=?, series=?, number=?,
                        issued_by=?, issue_date=?, valid_until=?, unique_number=?
                    WHERE id=?
                """, (
                    student_id, document_type, request.form.get('series') or None, request.form.get('number'),
                    request.form.get('issued_by') or None, request.form.get('issue_date') or None,
                    request.form.get('valid_until') or None, request.form.get('unique_number') or None,
                    doc_id,
                ))
                db.commit()

                if new_scans:
                    save_multiple_attachments(db, 'passport_document', doc_id, new_scans, 'passport_documents', current_username())
                    db.commit()

                log_action(current_username(), f"редагував паспортні дані ID {doc_id}", details=document_type)
                flash('Дані успішно оновлено', 'success')

            except (sqlite3.Error, ValueError) as e:
                db.rollback()
                logger.error(f"Помилка редагування паспортних даних (ID {doc_id}): {e}", exc_info=True)
                flash(f'Помилка при редагуванні: {e}', 'danger')

            return redirect(url_for('admin.manage_passport_documents', group_id=selected_group_id))

        # === Додавання нового запису ===
        else:
            student_id = request.form.get('student_id')
            document_type = request.form.get('document_type')

            try:
                if not student_id:
                    raise ValueError("Не обрано студента")
                if document_type not in ('Паспорт (книжка)', 'ID-картка'):
                    raise ValueError("Невірний тип документа")

                new_scans = [f for f in request.files.getlist('document_scans') if f and f.filename]
                if len(new_scans) > MAX_FILES:
                    raise ValueError(f"Забагато файлів - максимум {MAX_FILES} на запис")

                cursor.execute("""
                    INSERT INTO passport_documents (
                        student_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    student_id, document_type, request.form.get('series') or None, request.form.get('number'),
                    request.form.get('issued_by') or None, request.form.get('issue_date') or None,
                    request.form.get('valid_until') or None, request.form.get('unique_number') or None,
                ))
                passport_doc_id = cursor.lastrowid
                db.commit()

                if new_scans:
                    save_multiple_attachments(db, 'passport_document', passport_doc_id, new_scans, 'passport_documents', current_username())
                    db.commit()

                log_action(current_username(), f"додав паспортні дані: {document_type} №{request.form.get('number')}")
                flash('Запис успішно додано', 'success')

            except (sqlite3.Error, ValueError) as e:
                db.rollback()
                flash(f'Помилка при додаванні: {str(e)}', 'danger')

        return redirect(url_for('admin.manage_passport_documents', group_id=selected_group_id))

    cursor.close()
    return render_template(
        'manage_passport_documents.html',
        groups=groups,
        selected_group_id=selected_group_id,
        students_without_docs=students_without_docs,
        documents_by_group=sorted_documents_by_group,
        students=students,
        student=student,
        student_doc_id=student_doc_id,
        max_files=MAX_FILES,
    )


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


@admin_bp.route('/admin/manage_specialties', methods=['GET', 'POST'])
@permission_required('manage_specialties')
def manage_specialties():
    """
    Каталог галузей знань і спеціальностей (Постанова КМУ №1021 від
    30.08.2024, чинна з 01.11.2024). Дозволяє:
      - вмикати/вимикати актуальність спеціальності для цього закладу
        (is_active) - неактивні не показуються у випадаючому списку на
        сторінці "Групи", але не видаляються (щоб не зламати наявні
        групи, які вже на них посилаються);
      - додавати власні спеціальності, яких немає в офіційному переліку
        (is_custom=1) - позначаються окремо, щоб було видно, що це не
        з держреєстру;
      - редагувати назву вже наявного запису (напр. якщо офіційний
        текст пізніше уточнили).
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'toggle_active':
            code = request.form.get('code')
            conn.execute("UPDATE specialties SET is_active = 1 - is_active WHERE code = ?", (code,))
            conn.commit()
            row = conn.execute("SELECT name_ua, is_active FROM specialties WHERE code = ?", (code,)).fetchone()
            log_action(
                current_username(),
                f"{'увімкнув' if row['is_active'] else 'вимкнув'} спеціальність: {code} {row['name_ua']}"
            )

        elif action == 'add':
            code = (request.form.get('code') or '').strip()
            name_ua = (request.form.get('name_ua') or '').strip()
            short_name = (request.form.get('short_name') or '').strip() or None
            name_en = (request.form.get('name_en') or '').strip() or None
            field_code = request.form.get('knowledge_field_code')

            if not code or not name_ua or not field_code:
                flash("Заповніть код, назву і галузь знань.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO specialties (code, name_ua, short_name, name_en, knowledge_field_code, is_active, is_custom) VALUES (?, ?, ?, ?, ?, 1, 1)",
                        (code, name_ua, short_name, name_en, field_code)
                    )
                    conn.commit()
                    flash(f"Спеціальність «{code} {name_ua}» додано.", "success")
                    log_action(current_username(), f"додав власну спеціальність: {code} {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Спеціальність з кодом «{code}» вже існує.", "error")

        elif action == 'edit':
            code = request.form.get('code')
            name_ua = (request.form.get('name_ua') or '').strip()
            short_name = (request.form.get('short_name') or '').strip() or None
            name_en = (request.form.get('name_en') or '').strip() or None
            field_code = request.form.get('knowledge_field_code')
            if not name_ua or not field_code:
                flash("Заповніть назву і галузь знань.", "error")
            else:
                conn.execute(
                    "UPDATE specialties SET name_ua = ?, short_name = ?, name_en = ?, knowledge_field_code = ? WHERE code = ?",
                    (name_ua, short_name, name_en, field_code, code)
                )
                conn.commit()
                flash("Спеціальність оновлено.", "success")
                log_action(current_username(), f"редагував спеціальність: {code} {name_ua}")

        elif action == 'edit_field':
            field_code = request.form.get('field_code')
            field_name_ua = (request.form.get('field_name_ua') or '').strip()
            field_name_en = (request.form.get('field_name_en') or '').strip() or None
            if not field_name_ua:
                flash("Заповніть назву галузі знань.", "error")
            else:
                conn.execute(
                    "UPDATE knowledge_fields SET name_ua = ?, name_en = ? WHERE code = ?",
                    (field_name_ua, field_name_en, field_code)
                )
                conn.commit()
                flash("Галузь знань оновлено.", "success")
                log_action(current_username(), f"редагував галузь знань: {field_code} {field_name_ua}")

        elif action == 'delete':
            code = request.form.get('code')
            in_use = conn.execute("SELECT COUNT(*) AS c FROM groups WHERE specialty_code = ?", (code,)).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - є {in_use} груп(и), що посилаються на цю спеціальність. "
                      f"Вимкніть актуальність замість видалення.", "error")
            else:
                row = conn.execute("SELECT name_ua FROM specialties WHERE code = ?", (code,)).fetchone()
                conn.execute("DELETE FROM specialties WHERE code = ?", (code,))
                conn.commit()
                flash("Спеціальність видалено.", "success")
                log_action(current_username(), f"видалив спеціальність: {code} {row['name_ua'] if row else ''}")

        elif action == 'add_field':
            field_code = (request.form.get('field_code') or '').strip()
            field_name_ua = (request.form.get('field_name_ua') or '').strip()
            field_name_en = (request.form.get('field_name_en') or '').strip() or None
            if not field_code or not field_name_ua:
                flash("Заповніть код і назву галузі знань.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO knowledge_fields (code, name_ua, name_en) VALUES (?, ?, ?)",
                        (field_code, field_name_ua, field_name_en)
                    )
                    conn.commit()
                    flash(f"Галузь знань «{field_code} {field_name_ua}» додано.", "success")
                    log_action(current_username(), f"додав галузь знань: {field_code} {field_name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Галузь знань з кодом «{field_code}» вже існує.", "error")

        elif action == 'delete_field':
            field_code = request.form.get('field_code')
            in_use = conn.execute(
                "SELECT COUNT(*) AS c FROM specialties WHERE knowledge_field_code = ?", (field_code,)
            ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - на цю галузь знань посилається {in_use} спеціальність(і). "
                      f"Спершу перенесіть або видаліть їх.", "error")
            else:
                row = conn.execute("SELECT name_ua FROM knowledge_fields WHERE code = ?", (field_code,)).fetchone()
                conn.execute("DELETE FROM knowledge_fields WHERE code = ?", (field_code,))
                conn.commit()
                flash("Галузь знань видалено.", "success")
                log_action(current_username(), f"видалив галузь знань: {field_code} {row['name_ua'] if row else ''}")

    fields = conn.execute("""
        SELECT k.code, k.name_ua, k.name_en,
               (SELECT COUNT(*) FROM specialties s WHERE s.knowledge_field_code = k.code) AS specialties_count
        FROM knowledge_fields k ORDER BY k.code
    """).fetchall()
    specialties = conn.execute("""
        SELECT s.code, s.name_ua, s.short_name, s.name_en, s.is_active, s.is_custom, k.code AS field_code, k.name_ua AS field_name,
               (SELECT COUNT(*) FROM groups g WHERE g.specialty_code = s.code) AS groups_count
        FROM specialties s JOIN knowledge_fields k ON k.code = s.knowledge_field_code
        ORDER BY k.code, substr(s.code,1,1), CAST(substr(s.code,2) AS INTEGER)
    """).fetchall()
    conn.close()

    return render_template("manage_specialties.html", fields=fields, specialties=specialties)


@admin_bp.route('/admin/manage_degree_levels', methods=['GET', 'POST'])
@permission_required('manage_degree_levels')
def manage_degree_levels():
    """
    Каталог ступенів (Бакалавр/Магістр тощо) - раніше жорстко прописані
    в коді двома варіантами, тепер редагований список: вмикання/
    вимикання актуальності, додавання нових ступенів, редагування
    назв. Той самий принцип, що й у "Спеціальностях".
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'toggle_active':
            level_id = request.form.get('id')
            conn.execute("UPDATE degree_levels SET is_active = 1 - is_active WHERE id = ?", (level_id,))
            conn.commit()
            row = conn.execute("SELECT name_ua, is_active FROM degree_levels WHERE id = ?", (level_id,)).fetchone()
            if row:
                log_action(
                    current_username(),
                    f"{'увімкнув' if row['is_active'] else 'вимкнув'} ступінь: {row['name_ua']}"
                )

        elif action == 'add':
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву ступеня.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO degree_levels (name_ua, name_en, is_active) VALUES (?, ?, 1)",
                        (name_ua, name_en)
                    )
                    conn.commit()
                    flash(f"Ступінь «{name_ua}» додано.", "success")
                    log_action(current_username(), f"додав ступінь: {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Ступінь «{name_ua}» вже існує.", "error")

        elif action == 'edit':
            level_id = request.form.get('id')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву ступеня.", "error")
            else:
                try:
                    conn.execute(
                        "UPDATE degree_levels SET name_ua = ?, name_en = ? WHERE id = ?",
                        (name_ua, name_en, level_id)
                    )
                    conn.commit()
                    flash("Ступінь оновлено.", "success")
                    log_action(current_username(), f"редагував ступінь: {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Ступінь «{name_ua}» вже існує.", "error")

        elif action == 'delete':
            level_id = request.form.get('id')
            row = conn.execute("SELECT name_ua FROM degree_levels WHERE id = ?", (level_id,)).fetchone()
            in_use = conn.execute(
                "SELECT COUNT(*) AS c FROM groups WHERE degree_level = ?", (row['name_ua'] if row else '',)
            ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - є {in_use} груп(и) з цим ступенем. Вимкніть актуальність замість видалення.", "error")
            else:
                conn.execute("DELETE FROM degree_levels WHERE id = ?", (level_id,))
                conn.commit()
                flash("Ступінь видалено.", "success")
                log_action(current_username(), f"видалив ступінь: {row['name_ua'] if row else ''}")

    degree_levels = conn.execute("""
        SELECT d.id, d.name_ua, d.name_en, d.is_active,
               (SELECT COUNT(*) FROM groups g WHERE g.degree_level = d.name_ua) AS groups_count
        FROM degree_levels d ORDER BY d.id
    """).fetchall()
    conn.close()

    return render_template("manage_degree_levels.html", degree_levels=degree_levels)


@admin_bp.route('/admin/manage_licenses', methods=['GET', 'POST'])
@permission_required('manage_licenses')
def manage_licenses():
    """
    Каталог ліцензій, під якими студенти вступають до закладу
    (Львівська, Київська, Польська тощо) - властивість СТУДЕНТА, а не
    групи, оскільки в одній групі можуть навчатись студенти з різних
    ліцензій. Той самий принцип CRUD, що й у "Ступенях"/"Спеціальностях".
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'toggle_active':
            license_id = request.form.get('id')
            conn.execute("UPDATE institution_licenses SET is_active = 1 - is_active WHERE id = ?", (license_id,))
            conn.commit()
            row = conn.execute("SELECT short_name_ua, is_active FROM institution_licenses WHERE id = ?", (license_id,)).fetchone()
            if row:
                log_action(current_username(), f"{'увімкнув' if row['is_active'] else 'вимкнув'} ліцензію: {row['short_name_ua']}")

        elif action == 'add':
            name_ua = (request.form.get('name_ua') or '').strip()
            short_name_ua = (request.form.get('short_name_ua') or '').strip() or None
            name_en = (request.form.get('name_en') or '').strip() or None
            short_name_en = (request.form.get('short_name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть повну назву ліцензії.", "error")
            else:
                conn.execute(
                    "INSERT INTO institution_licenses (name_ua, short_name_ua, name_en, short_name_en, is_active) VALUES (?, ?, ?, ?, 1)",
                    (name_ua, short_name_ua, name_en, short_name_en)
                )
                conn.commit()
                flash(f"Ліцензію «{short_name_ua or name_ua}» додано.", "success")
                log_action(current_username(), f"додав ліцензію: {short_name_ua or name_ua}")

        elif action == 'edit':
            license_id = request.form.get('id')
            name_ua = (request.form.get('name_ua') or '').strip()
            short_name_ua = (request.form.get('short_name_ua') or '').strip() or None
            name_en = (request.form.get('name_en') or '').strip() or None
            short_name_en = (request.form.get('short_name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть повну назву ліцензії.", "error")
            else:
                conn.execute(
                    "UPDATE institution_licenses SET name_ua = ?, short_name_ua = ?, name_en = ?, short_name_en = ? WHERE id = ?",
                    (name_ua, short_name_ua, name_en, short_name_en, license_id)
                )
                conn.commit()
                flash("Ліцензію оновлено.", "success")
                log_action(current_username(), f"редагував ліцензію (ID {license_id})")

        elif action == 'delete':
            license_id = request.form.get('id')
            row = conn.execute("SELECT short_name_ua, name_ua FROM institution_licenses WHERE id = ?", (license_id,)).fetchone()
            in_use = conn.execute(
                "SELECT COUNT(*) AS c FROM students WHERE license_id = ?", (license_id,)
            ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - {in_use} студент(ів) мають цю ліцензію. Вимкніть актуальність замість видалення.", "error")
            else:
                conn.execute("DELETE FROM institution_licenses WHERE id = ?", (license_id,))
                conn.commit()
                flash("Ліцензію видалено.", "success")
                log_action(current_username(), f"видалив ліцензію: {row['short_name_ua'] or row['name_ua'] if row else ''}")

    licenses = conn.execute("""
        SELECT l.id, l.name_ua, l.short_name_ua, l.name_en, l.short_name_en, l.is_active,
               (SELECT COUNT(*) FROM students s WHERE s.license_id = l.id) AS students_count
        FROM institution_licenses l ORDER BY l.id
    """).fetchall()
    conn.close()

    return render_template("manage_licenses.html", licenses=licenses)


@admin_bp.route('/admin/license_transfer', methods=['GET', 'POST'])
@permission_required('manage_license_transfer')
def license_transfer():
    """
    Крок 1 -> 2 переведення студентів між ліцензіями. Джерело студентів
    для наказу - з активних груп (позначили групу цілком) і/або окремі
    студенти з різних груп (позначили напряму) - обидва джерела
    зливаються в один спільний список на кроці перегляду.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        target_license_id = request.form.get('target_license_id')
        group_ids = request.form.getlist('group_ids')
        student_ids = request.form.getlist('student_ids')

        if not target_license_id:
            flash("Оберіть ліцензію призначення.", "error")
            conn.close()
            return redirect(url_for('admin.license_transfer'))
        if not group_ids and not student_ids:
            flash("Оберіть хоча б одну групу або одного студента.", "error")
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        target_license = conn.execute("SELECT id, short_name_ua, name_ua FROM institution_licenses WHERE id = ?", (target_license_id,)).fetchone()
        if not target_license:
            flash("Ліцензію призначення не знайдено.", "error")
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        # Збираємо студентів з обох джерел в один набір (без дублів,
        # якщо той самий студент прийшов і через групу, і напряму)
        student_id_set = set(int(x) for x in student_ids)
        for group_id in group_ids:
            rows = conn.execute(
                "SELECT id FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0", (group_id,)
            ).fetchall()
            student_id_set.update(r['id'] for r in rows)

        if not student_id_set:
            flash("У обраних групах немає активних студентів.", "error")
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        placeholders = ','.join('?' for _ in student_id_set)
        students = conn.execute(f"""
            SELECT s.id, TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS full_name,
                   g.name AS group_name, l.short_name_ua AS current_license_name
            FROM students s
            LEFT JOIN groups g ON g.id = s.group_id
            LEFT JOIN institution_licenses l ON l.id = s.license_id
            WHERE s.id IN ({placeholders})
            ORDER BY s.last_name_UA COLLATE UKRAINIAN
        """, list(student_id_set)).fetchall()

        conn.close()
        return render_template("license_transfer_review.html", students=students, target_license=target_license)

    licenses = conn.execute("SELECT id, name_ua, short_name_ua FROM institution_licenses WHERE is_active = 1 ORDER BY id").fetchall()
    groups = conn.execute("""
        SELECT id, name,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = groups.id AND COALESCE(s.archived, 0) = 0) AS student_count
        FROM groups WHERE archived = FALSE ORDER BY name
    """).fetchall()
    all_students = conn.execute("""
        SELECT s.id, TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS full_name,
               g.name AS group_name, l.short_name_ua AS current_license_name
        FROM students s
        LEFT JOIN groups g ON g.id = s.group_id
        LEFT JOIN institution_licenses l ON l.id = s.license_id
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA COLLATE UKRAINIAN
    """).fetchall()
    conn.close()

    return render_template("license_transfer_select.html", licenses=licenses, groups=groups, all_students=all_students)


@admin_bp.route('/admin/license_transfer/confirm', methods=['POST'])
@permission_required('manage_license_transfer')
def license_transfer_confirm():
    """Остаточне підтвердження наказу про переведення між ліцензіями."""
    order_number = (request.form.get('order_number') or '').strip()
    order_date = request.form.get('order_date')
    target_license_id = request.form.get('target_license_id')
    student_ids = [int(x) for x in request.form.getlist('student_ids')]

    if not order_number or not order_date or not target_license_id or not student_ids:
        flash("Заповніть номер, дату наказу і переконайтесь, що обрано студентів.", "error")
        return redirect(url_for('admin.license_transfer'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    scan_rel_path = None
    scan_file = request.files.get('scan_file')
    if scan_file and scan_file.filename:
        os.makedirs(os.path.join('static', 'uploads', 'license_transfer_orders'), exist_ok=True)
        ext = os.path.splitext(scan_file.filename)[1]
        safe_name = f"{uuid.uuid4().hex}{ext}"
        scan_file.save(os.path.join('static', 'uploads', 'license_transfer_orders', safe_name))
        scan_rel_path = f"uploads/license_transfer_orders/{safe_name}"

    cur = conn.execute(
        "INSERT INTO license_transfer_orders (order_number, order_date, scan_file, created_by) VALUES (?, ?, ?, ?)",
        (order_number, order_date, scan_rel_path, current_username())
    )
    order_id = cur.lastrowid

    for student_id in student_ids:
        prev = conn.execute("SELECT license_id FROM students WHERE id = ?", (student_id,)).fetchone()
        previous_license_id = prev['license_id'] if prev else None
        conn.execute(
            "INSERT INTO license_transfer_order_students (order_id, student_id, previous_license_id, new_license_id) VALUES (?, ?, ?, ?)",
            (order_id, student_id, previous_license_id, target_license_id)
        )
        conn.execute("UPDATE students SET license_id = ? WHERE id = ?", (target_license_id, student_id))

    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"наказ про переведення між ліцензіями №{order_number} від {order_date}",
        details=f"студентів: {len(student_ids)}, нова ліцензія ID {target_license_id}"
    )
    flash(f"Переведення між ліцензіями оформлено. Студентів: {len(student_ids)}.", "success")
    return redirect(url_for('admin.manage_licenses'))


@admin_bp.route('/admin/license_report')
@permission_required('license_report')
def license_report():
    """
    Звіт по ліцензіях: скільки студентів кожної групи належать до
    якої ліцензії (Ф=Філія/Львів, К=Київ, П=Польща тощо - будь-які
    активні ліцензії з каталогу). Два режими: "Список" (як у
    друкованих списках вступників - курс/спеціальність/студенти з
    позначкою ліцензії) і "Підсумки" (зведена таблиця з кількостями,
    як власноруч ведена Excel-таблиця). Перемикач архівних груп - той
    самий принцип, що й в Аналітиці.
    """
    include_archived = request.args.get('include_archived') == '1'
    view = request.args.get('view', 'list')

    conn = get_db()
    conn.row_factory = sqlite3.Row

    group_cond = "" if include_archived else "WHERE COALESCE(g.archived,0)=0"
    groups = conn.execute(f"""
        SELECT g.id, g.name, g.course, g.specialty, g.degree_level, g.study_form, g.archived, g.program_credits
        FROM groups g {group_cond}
        ORDER BY g.course, g.specialty COLLATE UKRAINIAN, g.name COLLATE UKRAINIAN
    """).fetchall()

    licenses = conn.execute(
        "SELECT id, short_name_ua FROM institution_licenses WHERE is_active = 1 ORDER BY id"
    ).fetchall()
    license_names = [l['short_name_ua'] for l in licenses]

    # Порядок ступенів - за їхнім id в каталозі degree_levels, той
    # самий принцип, що й на дошці "Курси" й в "Архіві груп".
    degree_order = {
        row['name_ua']: row['id']
        for row in conn.execute("SELECT id, name_ua FROM degree_levels").fetchall()
    }

    groups_with_students = []
    summary_rows = []
    grand_totals = {name: 0 for name in license_names}
    grand_totals['—'] = 0  # без ліцензії
    grand_total_all = 0

    # Підсумки по (ступінь + форма навчання) - як блоки "Бакалавр
    # денна"/"Бакалавр заочна" тощо у власноруч веденій таблиці
    breakdown = {}
    # Підсумки просто по курсах, незалежно від спеціальності/ступеня
    course_breakdown = {}

    for g in groups:
        students = conn.execute("""
            SELECT s.id, TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS full_name,
                   l.short_name_ua AS license_short_name, s.archived, s.program_credits_override
            FROM students s LEFT JOIN institution_licenses l ON l.id = s.license_id
            WHERE s.group_id = ? {archived_cond}
            ORDER BY s.last_name_UA COLLATE UKRAINIAN
        """.format(archived_cond="" if include_archived else "AND COALESCE(s.archived,0)=0"), (g['id'],)).fetchall()

        counts = {name: 0 for name in license_names}
        counts['—'] = 0
        for s in students:
            key = s['license_short_name'] or '—'
            counts[key] = counts.get(key, 0) + 1
            grand_totals[key] = grand_totals.get(key, 0) + 1
            grand_total_all += 1

        groups_with_students.append({'group': g, 'students': students, 'counts': {k: v for k, v in counts.items() if v}})

        group_total = sum(counts.values())
        summary_rows.append({'group': g, 'counts': counts, 'total': group_total})

        breakdown_key = (g['degree_level'] or '—', g['study_form'] or '—')
        if breakdown_key not in breakdown:
            breakdown[breakdown_key] = {name: 0 for name in license_names}
            breakdown[breakdown_key]['—'] = 0
        for name, cnt in counts.items():
            breakdown[breakdown_key][name] = breakdown[breakdown_key].get(name, 0) + cnt

        course_key = g['course']
        if course_key not in course_breakdown:
            course_breakdown[course_key] = {name: 0 for name in license_names}
            course_breakdown[course_key]['—'] = 0
        for name, cnt in counts.items():
            course_breakdown[course_key][name] = course_breakdown[course_key].get(name, 0) + cnt

    conn.close()

    # Групуємо для компактного вигляду в режимі "Список": ступінь -> список груп
    groups_by_degree_raw = {}
    for item in groups_with_students:
        groups_by_degree_raw.setdefault(item['group']['degree_level'] or 'Без ступеня', []).append(item)
    groups_by_degree = {
        degree_level: groups_by_degree_raw[degree_level]
        for degree_level in sorted(groups_by_degree_raw.keys(), key=lambda d: (degree_order.get(d, 999), d))
    }

    breakdown_rows = [
        {'degree_level': k[0], 'study_form': k[1], 'counts': v, 'total': sum(v.values())}
        for k, v in sorted(breakdown.items())
    ]
    course_breakdown_rows = [
        {'course': k, 'counts': v, 'total': sum(v.values())}
        for k, v in sorted(course_breakdown.items())
    ]

    return render_template(
        "license_report.html",
        view=view,
        include_archived=include_archived,
        license_names=license_names,
        groups_with_students=groups_with_students,
        groups_by_degree=groups_by_degree,
        summary_rows=summary_rows,
        course_breakdown_rows=course_breakdown_rows,
        breakdown_rows=breakdown_rows,
        grand_totals=grand_totals,
        grand_total_all=grand_total_all,
    )


@admin_bp.route('/admin/license_report/export_word')
@permission_required('license_report')
def license_report_export_word():
    """
    Вивантажує звіт по ліцензіях у .docx у форматі: "ВСТУП {рік}" ->
    "{курс} курс {форма навчання}" -> реальна назва групи -> нумерована
    таблиця "№ / Прізвище І.П. / Ліцензія" (ліцензія - колонка в
    таблиці, а не заголовок блоку). Денна й заочна форми - окремі
    файли (study_form обов'язковий параметр запиту).
    """
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from collections import OrderedDict

    include_archived = request.args.get('include_archived') == '1'
    study_form = request.args.get('study_form')  # 'Денна' або 'Заочна' - обов'язково, окремий файл на форму
    if study_form not in ('Денна', 'Заочна'):
        flash("Оберіть форму навчання (Денна чи Заочна) для вивантаження.", "error")
        return redirect(url_for('admin.license_report'))


    conn = get_db()
    conn.row_factory = sqlite3.Row
    group_cond = "AND COALESCE(g.archived,0)=0" if not include_archived else ""
    student_cond = "AND COALESCE(s.archived,0)=0" if not include_archived else ""

    rows = conn.execute(f"""
        SELECT g.start_year, g.course, g.name AS group_name, g.degree_level,
               g.specialty_code, sp.name_ua AS specialty_name_ua,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA || ' ' || COALESCE(s.middle_name_UA, '')) AS full_name,
               l.short_name_ua AS license_short_name
        FROM groups g
        JOIN students s ON s.group_id = g.id
        LEFT JOIN specialties sp ON sp.code = g.specialty_code
        LEFT JOIN institution_licenses l ON l.id = s.license_id
        WHERE g.study_form = ? {group_cond} {student_cond}
        ORDER BY g.start_year DESC, g.course, g.name COLLATE UKRAINIAN, s.last_name_UA COLLATE UKRAINIAN
    """, (study_form,)).fetchall()
    conn.close()

    # Групуємо: ступінь (спершу Бакалавр, потім Магістр, решта - після)
    # -> рік вступу -> курс -> реальна назва групи -> студенти (ім'я +
    # своя ліцензія як окрема колонка, а не заголовок блоку).
    DEGREE_ORDER = {'Бакалавр': 0, 'Магістр': 1}
    DEGREE_LABELS = {'Бакалавр': 'БАКАЛАВРИ', 'Магістр': 'МАГІСТРИ'}

    grouped = OrderedDict()
    for r in rows:
        degree_key = r['degree_level'] or 'Інше'
        section_key = (r['start_year'], r['course'])
        group_bucket = grouped.setdefault(degree_key, OrderedDict()).setdefault(section_key, OrderedDict()).setdefault(
            r['group_name'], {'specialty_code': r['specialty_code'], 'specialty_name': r['specialty_name_ua'], 'students': []}
        )
        group_bucket['students'].append((r['full_name'], r['license_short_name'] or 'без ліцензії'))

    # Сортуємо ключі ступенів: Бакалавр, потім Магістр, потім усе інше
    # за абеткою - самі роки/курси всередині зберігають порядок з SQL.
    ordered_degrees = sorted(grouped.keys(), key=lambda d: DEGREE_ORDER.get(d, 99))

    doc = Document()
    for section in doc.sections:
        section.left_margin = Cm(2)
        section.right_margin = Cm(1)

    study_form_text = {'Денна': 'денна форма навчання', 'Заочна': 'заочна форма навчання'}.get(study_form, study_form)

    for degree_key in ordered_degrees:
        degree_label = DEGREE_LABELS.get(degree_key, degree_key.upper())
        years = grouped[degree_key]

        for (start_year, course), groups_dict in years.items():
            p = doc.add_paragraph()
            run = p.add_run(f"ВСТУП {start_year}")
            run.bold = True
            run.font.size = Pt(16)

            p_degree = doc.add_paragraph()
            run_degree = p_degree.add_run(degree_label)
            run_degree.bold = True
            run_degree.font.size = Pt(16)

            p2 = doc.add_paragraph()
            run2 = p2.add_run(f"{course} курс {study_form_text}")
            run2.bold = True
            run2.font.size = Pt(16)

            for group_name, group_info in groups_dict.items():
                students = group_info['students']
                specialty_code = group_info['specialty_code'] or ''
                specialty_name = group_info['specialty_name'] or ''
                header_text = group_name
                if specialty_code or specialty_name:
                    header_text += f" ({specialty_code}, {specialty_name})"

                p3 = doc.add_paragraph()
                run3 = p3.add_run(header_text)
                run3.bold = True
                run3.font.size = Pt(14)

                table = doc.add_table(rows=1, cols=3)
                table.style = 'Table Grid'
                hdr = table.rows[0].cells
                hdr[0].text = 'П/н'
                hdr[1].text = 'Прізвище І.П.'
                hdr[2].text = 'Ліцензія'
                for cell in hdr:
                    for para in cell.paragraphs:
                        for r in para.runs:
                            r.bold = True

                table.columns[0].width = Cm(1.5)
                table.columns[1].width = Cm(9)
                table.columns[2].width = Cm(4.5)

                for i, (name, license_name) in enumerate(students, start=1):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(i)
                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[1].text = name
                    row_cells[2].text = license_name

                doc.add_paragraph()

            doc.add_paragraph()

    output_path = os.path.join('static', 'uploads', f"license_report_{uuid.uuid4().hex}.docx")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)

    log_action(current_username(), f"вивантажив звіт по ліцензіях у Word ({study_form})")
    download_name = f"Списки_студентів_{'денна' if study_form == 'Денна' else 'заочна'}.docx"
    return send_file(output_path, as_attachment=True, download_name=download_name)







@admin_bp.route('/admin/manage_educational_programs', methods=['GET', 'POST'])
@permission_required('manage_educational_programs')
def manage_educational_programs():
    """
    Каталог освітніх програм: для кожної трійки "спеціальність + рік
    вступу + ступінь" - своя назва програми (укр./англ.). Ступінь
    важливий окремо від року: та сама спеціальність і рік вступу
    можуть мати зовсім різну програму на рівні Молодшого спеціаліста/
    Бакалавра/Магістра. Використовується на сторінці "Групи" для
    автопідстановки замість вільного вводу.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            specialty_code = request.form.get('specialty_code')
            start_year = request.form.get('start_year')
            degree_level = request.form.get('degree_level')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None

            if not specialty_code or not start_year or not degree_level or not name_ua:
                flash("Оберіть спеціальність, ступінь, вкажіть рік і назву програми.", "error")
            elif not start_year.isdigit():
                flash("Рік має бути числом.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO educational_programs (specialty_code, start_year, degree_level, name_ua, name_en) VALUES (?, ?, ?, ?, ?)",
                        (specialty_code, int(start_year), degree_level, name_ua, name_en)
                    )
                    conn.commit()
                    flash(f"Освітню програму для {specialty_code} ({degree_level}, {start_year}) додано.", "success")
                    log_action(current_username(), f"додав освітню програму: {specialty_code} {degree_level} {start_year} - {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Для спеціальності {specialty_code}, ступеня «{degree_level}» і року {start_year} програма вже існує - відредагуйте наявну.", "error")

        elif action == 'edit':
            program_id = request.form.get('id')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву програми.", "error")
            else:
                conn.execute(
                    "UPDATE educational_programs SET name_ua = ?, name_en = ? WHERE id = ?",
                    (name_ua, name_en, program_id)
                )
                conn.commit()
                flash("Освітню програму оновлено.", "success")
                log_action(current_username(), f"редагував освітню програму ID {program_id}: {name_ua}")

        elif action == 'delete':
            program_id = request.form.get('id')
            row = conn.execute(
                "SELECT specialty_code, start_year, degree_level, name_ua FROM educational_programs WHERE id = ?", (program_id,)
            ).fetchone()
            in_use = 0
            if row:
                in_use = conn.execute(
                    "SELECT COUNT(*) AS c FROM groups WHERE specialty_code = ? AND start_year = ? AND degree_level = ?",
                    (row['specialty_code'], row['start_year'], row['degree_level'])
                ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - є {in_use} груп(и), що використовують цю програму.", "error")
            else:
                conn.execute("DELETE FROM educational_programs WHERE id = ?", (program_id,))
                conn.commit()
                flash("Освітню програму видалено.", "success")
                log_action(current_username(), f"видалив освітню програму: {row['name_ua'] if row else ''}")

    specialties = conn.execute("""
        SELECT s.code, s.name_ua, k.name_ua AS field_name
        FROM specialties s JOIN knowledge_fields k ON k.code = s.knowledge_field_code
        WHERE s.is_active = 1
        ORDER BY substr(s.code,1,1), CAST(substr(s.code,2) AS INTEGER)
    """).fetchall()

    degree_levels = conn.execute(
        "SELECT name_ua FROM degree_levels WHERE is_active = 1 ORDER BY id"
    ).fetchall()

    programs = conn.execute("""
        SELECT p.id, p.specialty_code, p.start_year, p.degree_level, p.name_ua, p.name_en,
               s.name_ua AS specialty_name,
               (SELECT COUNT(*) FROM groups g
                WHERE g.specialty_code = p.specialty_code AND g.start_year = p.start_year AND g.degree_level = p.degree_level) AS groups_count
        FROM educational_programs p JOIN specialties s ON s.code = p.specialty_code
        ORDER BY p.specialty_code, p.start_year DESC, p.degree_level COLLATE UKRAINIAN
    """).fetchall()

    # Групуємо програми по спеціальності для зручного відображення
    programs_by_specialty = {}
    for p in programs:
        programs_by_specialty.setdefault(p['specialty_code'], []).append(p)

    conn.close()

    return render_template(
        "manage_educational_programs.html",
        specialties=specialties,
        degree_levels=degree_levels,
        programs_by_specialty=programs_by_specialty,
    )


@admin_bp.route('/admin/manage_qualification_names', methods=['GET', 'POST'])
@permission_required('manage_qualification_names')
def manage_qualification_names():
    """
    Каталог назв кваліфікації: для кожної трійки "спеціальність + рік
    вступу + ступінь" - своя назва кваліфікації (укр./англ.). Той самий
    принцип, що й "Освітня програма" - формулювання може відрізнятись і
    рік від року, і від ступеня.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            specialty_code = request.form.get('specialty_code')
            start_year = request.form.get('start_year')
            degree_level = request.form.get('degree_level')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None

            if not specialty_code or not start_year or not degree_level or not name_ua:
                flash("Оберіть спеціальність, ступінь, вкажіть рік і назву кваліфікації.", "error")
            elif not start_year.isdigit():
                flash("Рік має бути числом.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO qualification_names (specialty_code, start_year, degree_level, name_ua, name_en) VALUES (?, ?, ?, ?, ?)",
                        (specialty_code, int(start_year), degree_level, name_ua, name_en)
                    )
                    conn.commit()
                    flash(f"Назву кваліфікації для {specialty_code} ({degree_level}, {start_year}) додано.", "success")
                    log_action(current_username(), f"додав назву кваліфікації: {specialty_code} {degree_level} {start_year} - {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Для спеціальності {specialty_code}, ступеня «{degree_level}» і року {start_year} назва кваліфікації вже існує - відредагуйте наявну.", "error")

        elif action == 'edit':
            qual_id = request.form.get('id')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву кваліфікації.", "error")
            else:
                conn.execute(
                    "UPDATE qualification_names SET name_ua = ?, name_en = ? WHERE id = ?",
                    (name_ua, name_en, qual_id)
                )
                conn.commit()
                flash("Назву кваліфікації оновлено.", "success")
                log_action(current_username(), f"редагував назву кваліфікації ID {qual_id}: {name_ua}")

        elif action == 'delete':
            qual_id = request.form.get('id')
            row = conn.execute(
                "SELECT specialty_code, start_year, degree_level, name_ua FROM qualification_names WHERE id = ?", (qual_id,)
            ).fetchone()
            in_use = 0
            if row:
                in_use = conn.execute(
                    "SELECT COUNT(*) AS c FROM groups WHERE specialty_code = ? AND start_year = ? AND degree_level = ?",
                    (row['specialty_code'], row['start_year'], row['degree_level'])
                ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - є {in_use} груп(и), що використовують цю назву кваліфікації.", "error")
            else:
                conn.execute("DELETE FROM qualification_names WHERE id = ?", (qual_id,))
                conn.commit()
                flash("Назву кваліфікації видалено.", "success")
                log_action(current_username(), f"видалив назву кваліфікації: {row['name_ua'] if row else ''}")

    specialties = conn.execute("""
        SELECT s.code, s.name_ua, k.name_ua AS field_name
        FROM specialties s JOIN knowledge_fields k ON k.code = s.knowledge_field_code
        WHERE s.is_active = 1
        ORDER BY substr(s.code,1,1), CAST(substr(s.code,2) AS INTEGER)
    """).fetchall()

    degree_levels = conn.execute(
        "SELECT name_ua FROM degree_levels WHERE is_active = 1 ORDER BY id"
    ).fetchall()

    qualifications = conn.execute("""
        SELECT q.id, q.specialty_code, q.start_year, q.degree_level, q.name_ua, q.name_en,
               s.name_ua AS specialty_name,
               (SELECT COUNT(*) FROM groups g
                WHERE g.specialty_code = q.specialty_code AND g.start_year = q.start_year AND g.degree_level = q.degree_level) AS groups_count
        FROM qualification_names q JOIN specialties s ON s.code = q.specialty_code
        ORDER BY q.specialty_code, q.start_year DESC, q.degree_level COLLATE UKRAINIAN
    """).fetchall()

    qualifications_by_specialty = {}
    for q in qualifications:
        qualifications_by_specialty.setdefault(q['specialty_code'], []).append(q)

    conn.close()

    return render_template(
        "manage_qualification_names.html",
        specialties=specialties,
        degree_levels=degree_levels,
        qualifications_by_specialty=qualifications_by_specialty,
    )


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


@admin_bp.route('/admin/pending_students')
@permission_required('manage_pending_students')
def pending_students():
    """
    Список заявок з публічної анкети самореєстрації (routes/public_apply.py).
    За замовчуванням - лише необроблені ("new"), опрацьовані ховаються
    в окрему вкладку історії. Для кожної заявки одразу видно, чи є
    ймовірний дублікат (та сама ПІБ+дата народження) серед уже
    наявних студентів або серед інших необроблених заявок.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    status_filter = request.args.get('status', 'new')
    if status_filter not in ('new', 'approved', 'rejected'):
        status_filter = 'new'

    rows = conn.execute("""
        SELECT * FROM pending_students WHERE status = ?
        ORDER BY created_at DESC
    """, (status_filter,)).fetchall()
    rows = [dict(r) for r in rows]
    for r in rows:
        r['scans_count'] = conn.execute(
            "SELECT COUNT(*) FROM attachments WHERE entity_type='pending_student' AND entity_id=?", (r['id'],)
        ).fetchone()[0]

    # Дублікати рахуємо лише для вкладки "нові" - для вже опрацьованих
    # заявок це неактуально.
    duplicates_by_id = {}
    if status_filter == 'new':
        for row in rows:
            existing_student = conn.execute("""
                SELECT id, last_name_UA, first_name_UA FROM students
                WHERE LOWER_UA(last_name_UA) = LOWER_UA(?) AND LOWER_UA(first_name_UA) = LOWER_UA(?)
                  AND birth_date = ?
            """, (row['last_name_UA'], row['first_name_UA'], row['birth_date'])).fetchone()

            other_pending_count = conn.execute("""
                SELECT COUNT(*) FROM pending_students
                WHERE status = 'new' AND id != ?
                  AND LOWER_UA(last_name_UA) = LOWER_UA(?) AND LOWER_UA(first_name_UA) = LOWER_UA(?)
                  AND birth_date = ?
            """, (row['id'], row['last_name_UA'], row['first_name_UA'], row['birth_date'])).fetchone()[0]

            if existing_student or other_pending_count:
                duplicates_by_id[row['id']] = {
                    'existing_student': existing_student,
                    'other_pending_count': other_pending_count,
                }

    counts = {
        s: conn.execute("SELECT COUNT(*) FROM pending_students WHERE status=?", (s,)).fetchone()[0]
        for s in ('new', 'approved', 'rejected')
    }

    conn.close()
    return render_template(
        'admin_pending_students.html',
        rows=rows, status_filter=status_filter, counts=counts,
        duplicates_by_id=duplicates_by_id,
    )


@admin_bp.route('/admin/pending_students/<int:pending_id>', methods=['GET', 'POST'])
@permission_required('manage_pending_students')
def pending_student_review(pending_id):
    """
    Перегляд однієї заявки: можна виправити будь-яке поле (студенти з
    телефону одруковуються), а тоді або підтвердити (обравши групу і,
    за потреби, ліцензію та кредити скороченої програми - саме тут
    заявка стає реальним студентом), або відхилити з приміткою.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    row = conn.execute("SELECT * FROM pending_students WHERE id=?", (pending_id,)).fetchone()
    if not row:
        conn.close()
        flash("Заявку не знайдено", "error")
        return redirect(url_for('admin.pending_students'))

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'save':
            fields = [
                'last_name_UA', 'first_name_UA', 'middle_name_UA', 'last_name_ENG', 'first_name_ENG', 'birth_date',
                'phone', 'phone_backup', 'email',
                'document_type', 'document_series', 'document_number', 'document_institution',
                'document_country', 'document_date',
                'passport_document_type', 'passport_series', 'passport_number', 'passport_issued_by',
                'passport_issue_date', 'passport_valid_until', 'passport_unique_number',
                'military_registration_number_drpvr', 'military_registration_document', 'military_issued_vod',
                'military_specialty_number', 'military_rank', 'military_address',
                'military_change_credentials', 'military_change_reason',
            ]
            values = {f: (request.form.get(f) or '').strip() or None for f in fields}

            if not values['last_name_UA'] or not values['first_name_UA'] or not values['birth_date']:
                flash("Прізвище, ім'я і дата народження обов'язкові", "error")
            else:
                set_clause = ", ".join(f"{k}=?" for k in fields)
                conn.execute(
                    f"UPDATE pending_students SET {set_clause} WHERE id=?",
                    list(values[k] for k in fields) + [pending_id]
                )
                conn.commit()
                flash("Зміни збережено", "success")
            conn.close()
            return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

        elif action == 'reject':
            note = (request.form.get('review_note') or '').strip() or None
            conn.execute(
                "UPDATE pending_students SET status='rejected', reviewed_by=?, reviewed_at=datetime('now','localtime'), review_note=? WHERE id=?",
                (current_username(), note, pending_id)
            )
            conn.commit()
            log_action(current_username(), f"відхилив заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} (заявка ID {pending_id})", details=note or '')
            conn.close()
            flash("Заявку відхилено", "success")
            return redirect(url_for('admin.pending_students'))

        elif action == 'delete':
            # Видаляє саму заявку і її файли (фото, скани) - реального
            # студента, якщо заявку вже підтвердили, це НЕ чіпає: він
            # уже самостійний запис, не залежний від заявки.
            if row['photo_path']:
                photo_abs = os.path.join('static', row['photo_path'])
                if os.path.exists(photo_abs):
                    os.remove(photo_abs)
            for att in get_attachments(conn, 'pending_student', pending_id):
                att_abs = os.path.join('static', att['file_path'])
                if os.path.exists(att_abs):
                    os.remove(att_abs)
            for att in get_attachments(conn, 'pending_student_passport', pending_id):
                att_abs = os.path.join('static', att['file_path'])
                if os.path.exists(att_abs):
                    os.remove(att_abs)
            for att in get_attachments(conn, 'pending_student_military', pending_id):
                att_abs = os.path.join('static', att['file_path'])
                if os.path.exists(att_abs):
                    os.remove(att_abs)
            conn.execute("DELETE FROM attachments WHERE entity_type='pending_student' AND entity_id=?", (pending_id,))
            conn.execute("DELETE FROM attachments WHERE entity_type='pending_student_passport' AND entity_id=?", (pending_id,))
            conn.execute("DELETE FROM attachments WHERE entity_type='pending_student_military' AND entity_id=?", (pending_id,))
            conn.execute("DELETE FROM pending_students WHERE id=?", (pending_id,))
            conn.commit()
            log_action(current_username(), f"видалив заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} (заявка ID {pending_id})")
            conn.close()
            flash("Заявку видалено", "success")
            return redirect(url_for('admin.pending_students', status=row['status']))

        elif action == 'approve':
            group_id = request.form.get('group_id')
            if not group_id:
                flash("Оберіть групу, щоб підтвердити заявку", "error")
                conn.close()
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
            group_id = int(group_id)
            license_id = request.form.get('license_id') or None
            program_credits_override_raw = (request.form.get('program_credits_override') or '').strip()
            program_credits_override = int(program_credits_override_raw) if program_credits_override_raw else None

            cur = conn.execute("""
                INSERT INTO students (
                    last_name_UA, first_name_UA, middle_name_UA, last_name_ENG, first_name_ENG, birth_date,
                    group_id, license_id, phone, phone_backup, email, program_credits_override
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row['last_name_UA'], row['first_name_UA'], row['middle_name_UA'], row['last_name_ENG'], row['first_name_ENG'],
                row['birth_date'], group_id, license_id, row['phone'], row['phone_backup'], row['email'],
                program_credits_override,
            ))
            student_id = cur.lastrowid

            # Фото - копіюємо байти з "карантинної" папки в стандартне
            # місце зберігання фото студентів (той самий шлях, що й
            # для звичайного завантаження фото на картці студента).
            if row['photo_path']:
                try:
                    from routes.photo import photo_path_for_student, _ensure_dir, PHOTOS_DIR
                    _ensure_dir()
                    src_path = os.path.join('static', row['photo_path'])
                    if os.path.exists(src_path):
                        dest_path = photo_path_for_student(student_id)
                        with open(src_path, 'rb') as src, open(dest_path, 'wb') as dst:
                            dst.write(src.read())
                        conn.execute("UPDATE students SET photo=? WHERE id=?",
                                     (f"uploads/photos/student_{student_id}.jpg", student_id))
                except Exception as e:
                    logger.error(f"Не вдалося перенести фото заявки {pending_id} студенту {student_id}: {e}")

            if row['document_type'] or row['document_number']:
                doc_number = ' '.join(x for x in [row['document_series'], row['document_number']] if x) or ''
                cur_doc = conn.execute("""
                    INSERT INTO education_documents (
                        student_id, document_type, document_type_en, document_number,
                        institution_name, institution_name_en, country, country_en, completion_date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    student_id, row['document_type'] or '', '', doc_number,
                    row['document_institution'] or '', '', row['document_country'] or '', '',
                    row['document_date'] or '',
                ))
                # Скани, завантажені разом із заявкою, - це скани саме
                # цього документа: переприв'язуємо (сам файл на диску
                # лишається на місці, змінюється лише запис у attachments).
                conn.execute(
                    "UPDATE attachments SET entity_type='education_document', entity_id=? WHERE entity_type='pending_student' AND entity_id=?",
                    (cur_doc.lastrowid, pending_id)
                )
            else:
                # Немає окремого документа про освіту, куди прив'язати
                # скани, - лишаємо їх загальними вкладеннями студента.
                conn.execute(
                    "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student' AND entity_id=?",
                    (student_id, pending_id)
                )

            if row['passport_number']:
                cur_passport = conn.execute("""
                    INSERT INTO passport_documents (
                        student_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    student_id, row['passport_document_type'] or 'Паспорт (книжка)', row['passport_series'],
                    row['passport_number'], row['passport_issued_by'], row['passport_issue_date'],
                    row['passport_valid_until'], row['passport_unique_number'],
                ))
                conn.execute(
                    "UPDATE attachments SET entity_type='passport_document', entity_id=? WHERE entity_type='pending_student_passport' AND entity_id=?",
                    (cur_passport.lastrowid, pending_id)
                )
            else:
                conn.execute(
                    "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student_passport' AND entity_id=?",
                    (student_id, pending_id)
                )

            if any([row['military_registration_number_drpvr'], row['military_registration_document'], row['military_rank']]):
                cur_military = conn.execute("""
                    INSERT INTO military (
                        student_id, registration_number_of_the_DRPVR, military_registration_document, issued_VOD,
                        military_accounting_specialty_number, military_rank, address_of_residence,
                        change_credentials, reason_for_changing_credentials
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    student_id, row['military_registration_number_drpvr'], row['military_registration_document'],
                    row['military_issued_vod'], row['military_specialty_number'], row['military_rank'],
                    row['military_address'], row['military_change_credentials'], row['military_change_reason'],
                ))
                conn.execute(
                    "UPDATE attachments SET entity_type='military', entity_id=? WHERE entity_type='pending_student_military' AND entity_id=?",
                    (cur_military.lastrowid, pending_id)
                )
            else:
                conn.execute(
                    "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student_military' AND entity_id=?",
                    (student_id, pending_id)
                )

            conn.execute(
                "UPDATE pending_students SET status='approved', reviewed_by=?, reviewed_at=datetime('now','localtime'), resulting_student_id=? WHERE id=?",
                (current_username(), student_id, pending_id)
            )
            conn.commit()
            log_action(
                current_username(),
                f"підтвердив заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} -> студент ID {student_id}",
                group_ids=[group_id],
            )
            conn.close()
            flash(f"Студента {row['last_name_UA']} {row['first_name_UA']} додано", "success")
            return redirect(url_for('students.student_details', student_id=student_id))

        elif action == 'update_existing':
            # Не створює нового студента - обраними пунктами оновлює
            # ВЖЕ НАЯВНОГО (той самий, на якого вказує "можливий
            # дублікат"). Кожен пункт - окрема галочка, щоб адмін сам
            # вирішував, що саме брати з заявки, а що лишити як є.
            existing_id = request.form.get('existing_student_id')
            if not existing_id:
                flash("Не вказано, якого студента оновлювати", "error")
                conn.close()
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
            existing_id = int(existing_id)

            updated_parts = []

            if request.form.get('update_personal'):
                conn.execute("""
                    UPDATE students SET last_name_UA=?, first_name_UA=?, middle_name_UA=?,
                                         last_name_ENG=?, first_name_ENG=?, birth_date=?
                    WHERE id=?
                """, (
                    row['last_name_UA'], row['first_name_UA'], row['middle_name_UA'],
                    row['last_name_ENG'], row['first_name_ENG'], row['birth_date'], existing_id
                ))
                updated_parts.append('особисті дані')

            if request.form.get('update_contacts'):
                conn.execute(
                    "UPDATE students SET phone=?, phone_backup=?, email=? WHERE id=?",
                    (row['phone'], row['phone_backup'], row['email'], existing_id)
                )
                updated_parts.append('контакти')

            if request.form.get('update_photo') and row['photo_path']:
                try:
                    from routes.photo import photo_path_for_student, _ensure_dir
                    _ensure_dir()
                    src_path = os.path.join('static', row['photo_path'])
                    if os.path.exists(src_path):
                        dest_path = photo_path_for_student(existing_id)
                        with open(src_path, 'rb') as src, open(dest_path, 'wb') as dst:
                            dst.write(src.read())
                        conn.execute("UPDATE students SET photo=? WHERE id=?",
                                     (f"uploads/photos/student_{existing_id}.jpg", existing_id))
                        updated_parts.append('фото')
                except Exception as e:
                    logger.error(f"Не вдалося перенести фото заявки {pending_id} студенту {existing_id}: {e}")

            doc_id_for_scans = None
            if request.form.get('add_document') and (row['document_type'] or row['document_number']):
                doc_number = ' '.join(x for x in [row['document_series'], row['document_number']] if x) or ''
                cur_doc = conn.execute("""
                    INSERT INTO education_documents (
                        student_id, document_type, document_type_en, document_number,
                        institution_name, institution_name_en, country, country_en, completion_date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    existing_id, row['document_type'] or '', '', doc_number,
                    row['document_institution'] or '', '', row['document_country'] or '', '',
                    row['document_date'] or '',
                ))
                doc_id_for_scans = cur_doc.lastrowid
                updated_parts.append('документ про освіту (додано як новий)')

            if doc_id_for_scans:
                conn.execute(
                    "UPDATE attachments SET entity_type='education_document', entity_id=? WHERE entity_type='pending_student' AND entity_id=?",
                    (doc_id_for_scans, pending_id)
                )
            # Будь-які скани, що лишились непереприв'язаними (документ
            # не додавали чи галочку не ставили) - чіпляємо як загальні
            # файли студента, щоб не загубились.
            conn.execute(
                "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student' AND entity_id=?",
                (existing_id, pending_id)
            )

            passport_id_for_scans = None
            if request.form.get('add_passport') and row['passport_number']:
                cur_passport = conn.execute("""
                    INSERT INTO passport_documents (
                        student_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    existing_id, row['passport_document_type'] or 'Паспорт (книжка)', row['passport_series'],
                    row['passport_number'], row['passport_issued_by'], row['passport_issue_date'],
                    row['passport_valid_until'], row['passport_unique_number'],
                ))
                passport_id_for_scans = cur_passport.lastrowid
                updated_parts.append('паспортні дані (додано як новий запис)')

            if passport_id_for_scans:
                conn.execute(
                    "UPDATE attachments SET entity_type='passport_document', entity_id=? WHERE entity_type='pending_student_passport' AND entity_id=?",
                    (passport_id_for_scans, pending_id)
                )
            conn.execute(
                "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student_passport' AND entity_id=?",
                (existing_id, pending_id)
            )

            if request.form.get('add_military') and any([
                row['military_registration_number_drpvr'], row['military_registration_document'], row['military_rank']
            ]):
                existing_military = conn.execute("SELECT id FROM military WHERE student_id=?", (existing_id,)).fetchone()
                if existing_military:
                    conn.execute("""
                        UPDATE military SET registration_number_of_the_DRPVR=?, military_registration_document=?,
                                             issued_VOD=?, military_accounting_specialty_number=?, military_rank=?,
                                             address_of_residence=?, change_credentials=?, reason_for_changing_credentials=?
                        WHERE id=?
                    """, (
                        row['military_registration_number_drpvr'], row['military_registration_document'],
                        row['military_issued_vod'], row['military_specialty_number'], row['military_rank'],
                        row['military_address'], row['military_change_credentials'], row['military_change_reason'],
                        existing_military['id']
                    ))
                    military_id_for_scans = existing_military['id']
                else:
                    cur_military = conn.execute("""
                        INSERT INTO military (
                            student_id, registration_number_of_the_DRPVR, military_registration_document, issued_VOD,
                            military_accounting_specialty_number, military_rank, address_of_residence,
                            change_credentials, reason_for_changing_credentials
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        existing_id, row['military_registration_number_drpvr'], row['military_registration_document'],
                        row['military_issued_vod'], row['military_specialty_number'], row['military_rank'],
                        row['military_address'], row['military_change_credentials'], row['military_change_reason'],
                    ))
                    military_id_for_scans = cur_military.lastrowid
                updated_parts.append('військовий облік')
                conn.execute(
                    "UPDATE attachments SET entity_type='military', entity_id=? WHERE entity_type='pending_student_military' AND entity_id=?",
                    (military_id_for_scans, pending_id)
                )
            conn.execute(
                "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student_military' AND entity_id=?",
                (existing_id, pending_id)
            )

            if not updated_parts:
                conn.rollback()
                conn.close()
                flash("Не обрано жодного пункту для оновлення", "error")
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

            conn.execute(
                "UPDATE pending_students SET status='approved', reviewed_by=?, reviewed_at=datetime('now','localtime'), resulting_student_id=? WHERE id=?",
                (current_username(), existing_id, pending_id)
            )
            conn.commit()
            log_action(
                current_username(),
                f"оновив наявного студента (ID {existing_id}) даними із заявки на реєстрацію {pending_id}",
                details=', '.join(updated_parts)
            )
            conn.close()
            flash(f"Оновлено: {', '.join(updated_parts)}", "success")
            return redirect(url_for('students.student_details', student_id=existing_id))

        conn.close()
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    groups = conn.execute("""
        SELECT id, name, start_year, study_form, program_credits,
               name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
        FROM groups WHERE archived = FALSE ORDER BY name COLLATE UKRAINIAN
    """).fetchall()
    licenses = conn.execute("SELECT id, short_name_ua, name_ua FROM institution_licenses WHERE is_active=1 ORDER BY id").fetchall()

    existing_student = conn.execute("""
        SELECT id, last_name_UA, first_name_UA FROM students
        WHERE LOWER_UA(last_name_UA) = LOWER_UA(?) AND LOWER_UA(first_name_UA) = LOWER_UA(?) AND birth_date = ?
    """, (row['last_name_UA'], row['first_name_UA'], row['birth_date'])).fetchone()

    scans = get_attachments(conn, 'pending_student', pending_id)
    passport_scans = get_attachments(conn, 'pending_student_passport', pending_id)
    military_scans = get_attachments(conn, 'pending_student_military', pending_id)

    conn.close()
    return render_template(
        'admin_pending_student_review.html',
        row=row, groups=groups, licenses=licenses, existing_student=existing_student,
        scans=scans, passport_scans=passport_scans, military_scans=military_scans,
    )


@admin_bp.route('/admin/pending_students/<int:pending_id>/photo', methods=['POST'])
@permission_required('manage_pending_students')
def pending_student_photo(pending_id):
    """
    Ручна (пере)обрізка фото абітурієнта на сторінці перегляду заявки -
    коли автоматична обрізка по центру (яку робить сама публічна форма)
    вийшла невдало. Той самий Cropper.js-підхід, що й на картці
    студента (students.upload_photo), лише зберігає результат у
    "карантинну" папку заявки, а не в фото реального студента.
    """
    conn = get_db()
    row = conn.execute("SELECT photo_path FROM pending_students WHERE id=?", (pending_id,)).fetchone()
    if not row:
        conn.close()
        flash("Заявку не знайдено", "error")
        return redirect(url_for('admin.pending_students'))

    file = request.files.get('photo_file')
    if not file or file.filename == '':
        conn.close()
        flash("Оберіть файл фотографії", "error")
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    try:
        crop_box = (
            float(request.form['crop_x']),
            float(request.form['crop_y']),
            float(request.form['crop_w']),
            float(request.form['crop_h']),
        )
    except (KeyError, ValueError):
        conn.close()
        flash("Некоректні дані обрізки фото - спробуйте ще раз", "error")
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    from routes.photo import process_and_save_pending_photo_with_crop
    try:
        new_path = process_and_save_pending_photo_with_crop(file.read(), crop_box, old_rel_path=row['photo_path'])
    except ValueError as e:
        conn.close()
        flash(str(e), "error")
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    conn.execute("UPDATE pending_students SET photo_path=? WHERE id=?", (new_path, pending_id))
    conn.commit()
    conn.close()
    flash("Фото оновлено", "success")
    return redirect(url_for('admin.pending_student_review', pending_id=pending_id))


@admin_bp.route('/admin/update_requests')
@permission_required('manage_students')
def update_requests():
    """
    Список заявок на оновлення даних наявного студента (з токен-
    посилань, які адмін сам створює на картці студента - routes/
    students.generate_update_link). За замовчуванням - лише ті, що
    студент уже заповнив і чекають на модерацію (status='submitted').
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    status_filter = request.args.get('status', 'submitted')
    if status_filter not in ('pending', 'submitted', 'approved', 'rejected'):
        status_filter = 'submitted'

    rows = conn.execute("""
        SELECT ur.*, s.last_name_UA, s.first_name_UA
        FROM update_requests ur
        JOIN students s ON s.id = ur.student_id
        WHERE ur.status = ?
        ORDER BY ur.created_at DESC
    """, (status_filter,)).fetchall()

    counts = {
        s: conn.execute("SELECT COUNT(*) FROM update_requests WHERE status=?", (s,)).fetchone()[0]
        for s in ('pending', 'submitted', 'approved', 'rejected')
    }

    conn.close()
    return render_template('admin_update_requests.html', rows=rows, status_filter=status_filter, counts=counts)


@admin_bp.route('/admin/update_requests/<int:request_id>', methods=['GET', 'POST'])
@permission_required('manage_students')
def update_request_review(request_id):
    """Перегляд однієї заявки на оновлення - вибіркове застосування
    полів до наявного студента (та сама логіка "яке поле застосувати",
    що й для дублікатів заявок на реєстрацію)."""
    conn = get_db()
    conn.row_factory = sqlite3.Row
    row = conn.execute("""
        SELECT ur.*, s.last_name_UA, s.first_name_UA, s.middle_name_UA
        FROM update_requests ur JOIN students s ON s.id = ur.student_id
        WHERE ur.id = ?
    """, (request_id,)).fetchone()
    if not row:
        conn.close()
        flash("Заявку не знайдено", "error")
        return redirect(url_for('admin.update_requests'))

    allowed_fields = json.loads(row['allowed_fields'])
    submitted = json.loads(row['submitted_data']) if row['submitted_data'] else {}

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'delete':
            for entity_type in ('update_request_passport', 'update_request_education', 'update_request_military'):
                for att in get_attachments(conn, entity_type, request_id):
                    att_path = os.path.join('static', att['file_path'])
                    if os.path.exists(att_path):
                        os.remove(att_path)
                conn.execute("DELETE FROM attachments WHERE entity_type=? AND entity_id=?", (entity_type, request_id))
            if row['photo_path']:
                photo_abs = os.path.join('static', row['photo_path'])
                if os.path.exists(photo_abs):
                    os.remove(photo_abs)
            conn.execute("DELETE FROM update_requests WHERE id=?", (request_id,))
            conn.commit()
            log_action(current_username(), f"видалив заявку на оновлення даних ID {request_id}")
            conn.close()
            flash("Заявку видалено", "success")
            return redirect(url_for('admin.update_requests'))

        elif action == 'reject':
            note = (request.form.get('review_note') or '').strip() or None
            conn.execute(
                "UPDATE update_requests SET status='rejected', reviewed_by=?, reviewed_at=datetime('now','localtime'), review_note=? WHERE id=?",
                (current_username(), note, request_id)
            )
            conn.commit()
            log_action(current_username(), f"відхилив заявку на оновлення даних ID {request_id}", details=note or '')
            conn.close()
            flash("Заявку відхилено", "success")
            return redirect(url_for('admin.update_requests'))

        elif action == 'apply':
            student_id = row['student_id']
            applied_parts = []

            simple_updates = {}
            for key in ('last_name_UA', 'first_name_UA', 'middle_name_UA', 'phone', 'phone_backup', 'email', 'tax_id', 'edebo_code'):
                if key in allowed_fields and request.form.get(f'apply_{key}'):
                    simple_updates[key] = submitted.get(key)
            if simple_updates:
                set_clause = ", ".join(f"{k}=?" for k in simple_updates)
                conn.execute(f"UPDATE students SET {set_clause} WHERE id=?", list(simple_updates.values()) + [student_id])
                applied_parts.append(", ".join(simple_updates.keys()))

            if 'photo' in allowed_fields and row['photo_path'] and request.form.get('apply_photo'):
                from routes.photo import photo_path_for_student, _ensure_dir
                _ensure_dir()
                src_path = os.path.join('static', row['photo_path'])
                if os.path.exists(src_path):
                    dest_path = photo_path_for_student(student_id)
                    with open(src_path, 'rb') as src, open(dest_path, 'wb') as dst:
                        dst.write(src.read())
                    conn.execute("UPDATE students SET photo=? WHERE id=?", (f"uploads/photos/student_{student_id}.jpg", student_id))
                    applied_parts.append('фото')

            if 'passport' in allowed_fields and submitted.get('passport') and request.form.get('apply_passport'):
                p = submitted['passport']
                cur_p = conn.execute("""
                    INSERT INTO passport_documents (student_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (student_id, p.get('document_type') or 'Паспорт (книжка)', p.get('series'), p.get('number'),
                      p.get('issued_by'), p.get('issue_date'), p.get('valid_until'), p.get('unique_number')))
                conn.execute(
                    "UPDATE attachments SET entity_type='passport_document', entity_id=? WHERE entity_type='update_request_passport' AND entity_id=?",
                    (cur_p.lastrowid, request_id)
                )
                applied_parts.append('паспортні дані (новий запис)')

            if 'education_document' in allowed_fields and submitted.get('education_document') and request.form.get('apply_education_document'):
                d = submitted['education_document']
                cur_d = conn.execute("""
                    INSERT INTO education_documents (student_id, document_type, document_type_en, document_number,
                        institution_name, institution_name_en, country, country_en, completion_date)
                    VALUES (?, ?, '', ?, ?, '', ?, '', ?)
                """, (student_id, d.get('document_type') or '', d.get('number') or '', d.get('institution') or '',
                      d.get('country') or '', d.get('completion_date') or ''))
                conn.execute(
                    "UPDATE attachments SET entity_type='education_document', entity_id=? WHERE entity_type='update_request_education' AND entity_id=?",
                    (cur_d.lastrowid, request_id)
                )
                applied_parts.append('документ про освіту (новий запис)')

            if 'military' in allowed_fields and submitted.get('military') and request.form.get('apply_military'):
                m = submitted['military']
                existing_military = conn.execute("SELECT id FROM military WHERE student_id=?", (student_id,)).fetchone()
                if existing_military:
                    conn.execute("""
                        UPDATE military SET registration_number_of_the_DRPVR=?, military_registration_document=?,
                            issued_VOD=?, military_accounting_specialty_number=?, military_rank=?, address_of_residence=?
                        WHERE id=?
                    """, (m.get('registration_number_of_the_DRPVR'), m.get('military_registration_document'),
                          m.get('issued_VOD'), m.get('military_accounting_specialty_number'), m.get('military_rank'),
                          m.get('address_of_residence'), existing_military['id']))
                    military_id = existing_military['id']
                else:
                    cur_m = conn.execute("""
                        INSERT INTO military (student_id, registration_number_of_the_DRPVR, military_registration_document,
                            issued_VOD, military_accounting_specialty_number, military_rank, address_of_residence)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (student_id, m.get('registration_number_of_the_DRPVR'), m.get('military_registration_document'),
                          m.get('issued_VOD'), m.get('military_accounting_specialty_number'), m.get('military_rank'),
                          m.get('address_of_residence')))
                    military_id = cur_m.lastrowid
                conn.execute(
                    "UPDATE attachments SET entity_type='military', entity_id=? WHERE entity_type='update_request_military' AND entity_id=?",
                    (military_id, request_id)
                )
                applied_parts.append('військові дані')

            if not applied_parts:
                conn.rollback()
                conn.close()
                flash("Не обрано жодного пункту для застосування", "error")
                return redirect(url_for('admin.update_request_review', request_id=request_id))

            conn.execute(
                "UPDATE update_requests SET status='approved', reviewed_by=?, reviewed_at=datetime('now','localtime') WHERE id=?",
                (current_username(), request_id)
            )
            conn.commit()
            log_action(
                current_username(),
                f"застосував заявку на оновлення даних: {row['last_name_UA']} {row['first_name_UA']} (ID {student_id})",
                details=", ".join(applied_parts)
            )
            conn.close()
            flash(f"Застосовано: {', '.join(applied_parts)}", "success")
            return redirect(url_for('students.student_details', student_id=student_id))

        conn.close()
        return redirect(url_for('admin.update_request_review', request_id=request_id))

    passport_scans = get_attachments(conn, 'update_request_passport', request_id)
    education_scans = get_attachments(conn, 'update_request_education', request_id)
    military_scans = get_attachments(conn, 'update_request_military', request_id)

    conn.close()
    return render_template(
        'admin_update_request_review.html',
        row=row, allowed_fields=allowed_fields, submitted=submitted,
        passport_scans=passport_scans, education_scans=education_scans, military_scans=military_scans,
    )


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


@admin_bp.route('/admin/view_logs')
@permission_required('view_logs')
def view_logs():
    """Показує журнал дій (app.log) у зручному розібраному вигляді: дата/час/рівень/користувач/дія, з можливістю фільтрації на фронтенді."""
    current_dir = os.path.dirname(__file__)
    project_root = os.path.dirname(current_dir)
    log_file_path = os.path.join(project_root, 'app.log')

    parsed_logs = []

    if os.path.exists(log_file_path):
        try:
            with open(log_file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    entry = {'raw': line, 'date': '', 'time': '', 'level': 'INFO', 'username': '', 'action': line}
                    m = re.match(r'(\d{4}-\d{2}-\d{2})\s*\|\s*(\d{2}:\d{2}:\d{2})\s*\|\s*(\w+)\s*\|\s*(.*)', line)
                    if m:
                        entry['date'] = m.group(1)
                        entry['time'] = m.group(2)
                        entry['level'] = m.group(3)
                        rest = m.group(4).strip()
                        entry['action'] = rest
                        u = re.match(r'👤\s*([^\s-][^-]*?)\s+-\s+(.*)', rest)
                        if u:
                            entry['username'] = u.group(1).strip()
                            entry['action'] = u.group(2).strip()
                    parsed_logs.append(entry)
        except Exception as e:
            logger.error(f"Помилка при читанні логів: {e}")

    parsed_logs.reverse()
    usernames = sorted({e['username'] for e in parsed_logs if e['username']})

    log_action(current_username(), "переглянув журнал дій")

    from datetime import date
    return render_template('view_logs.html', logs=parsed_logs, usernames=usernames,
                           now=date.today().strftime('%Y-%m-%d'))


@admin_bp.route('/admin/users', methods=['GET', 'POST'])
@permission_required('manage_users')
def manage_users():
    """Список користувачів системи та редагування їхніх прав доступу (is_admin + список дозволів permissions)."""
    conn = get_db()
    try:
        users = conn.execute("""
            SELECT u.id, u.username, u.role, u.is_admin, u.permissions,
                   GROUP_CONCAT(g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)', ', ') AS group_names
            FROM users u
            LEFT JOIN user_groups ug ON u.id = ug.user_id
            LEFT JOIN groups g ON ug.group_id = g.id
            GROUP BY u.id ORDER BY u.username
        """).fetchall()

        perm_names_ua = {
            'manage_users': 'Список користувачів та управління правами',
            'view_logs': 'Журнал дій',
            'group_export': 'Масова генерація документів',
            'import_from_excel': 'Інпорт студентів',
            'manage_education_documents': 'Управління документами про освіту',
            'manage_passport_documents': 'Управління паспортними даними',
            'import_passport_documents': 'Імпорт паспортних даних',
            'study_periods': 'Періоди навчання',
            'manage_groups': 'Управління групами',
            'manage_subjects': 'Предмети',
            'manage_activities': 'Управління діяльностями',
            'import_subjects': 'Імпорт предметів з Excel',
            'archive': 'Управління архівом',
            'manage_students': 'Управління студентами (Видалення студента та його війс. док.)',
            'manage_accreditations': 'Управління акредетаціями',
            'manage_diplomas': 'Управління номерами диплому і додатку',
            'import_education_docs': 'Управління імпортом документів',
            'manage_templates': 'Управління шаблонами документів',
            'import_grades': 'Імпорт оцінок з Excel',
            'analytics': 'Аналітика',
            'manage_specialties': 'Спеціальності',
            'manage_degree_levels': 'Ступені',
            'manage_educational_programs': 'Освітня програма',
            'manage_qualification_names': 'Назва кваліфікації',
            'manage_courses': 'Курси (та переведення на курс / випуск)',
            'manage_frozen_students': 'Заморожені студенти',
            'manage_expulsion': 'Наказ про відрахування',
            'manage_licenses': 'Ліцензії',
            'manage_license_transfer': 'Перевести між ліцензіями',
            'license_report': 'Звіт по ліцензіях',
            'manage_pending_students': 'Заявки на реєстрацію (публічна анкета)',
        }

        if request.method == 'POST':
            user_id = request.form.get('user_id')
            if not user_id:
                flash('Не вказано користувача', 'danger')
                return redirect(url_for('admin.manage_users'))

            is_admin = 1 if 'is_admin' in request.form else 0
            selected_perms = [p for p in PERMISSIONS if p in request.form]

            target_user = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()

            conn.execute("UPDATE users SET is_admin=?, permissions=? WHERE id=?",
                         (is_admin, json.dumps(selected_perms), user_id))
            conn.commit()

            log_action(
                current_username(),
                f"змінив права: {target_user['username']} (ID {user_id})",
                details=f"is_admin: {bool(is_admin)}, дозволи: {', '.join(selected_perms) or 'жодного'}"
            )
            flash('Права успішно оновлено', 'success')
            return redirect(url_for('admin.manage_users'))

        return render_template('manage_users.html', users=users, permissions=PERMISSIONS, perm_names_ua=perm_names_ua)

    except sqlite3.Error as e:
        logger.error(f"Помилка бази даних у manage_users: {e}", exc_info=True)
        flash(f'Помилка бази даних: {e}', 'danger')
        return redirect(url_for('admin.manage_users'))
    finally:
        conn.close()


@admin_bp.route('/admin/users/add', methods=['GET', 'POST'])
@permission_required('manage_users')
def add_user():
    """Форма створення нового користувача (логін/пароль/роль/групи)."""
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        role = request.form.get('role')
        group_ids = request.form.getlist('group_id')

        if not all([username, password, role]):
            flash('Заповніть усі обовязкові поля', 'danger')
            return redirect(url_for('admin.add_user'))

        is_admin = 1 if role == 'admin' else 0
        permissions = json.dumps([])

        conn = get_db()
        try:
            exists = conn.execute("SELECT 1 FROM users WHERE username=?", (username,)).fetchone()
            if exists:
                flash(f'Користувач з іменем "{username}" вже існує', 'danger')
                return redirect(url_for('admin.add_user'))

            cursor = conn.execute(
                "INSERT INTO users (username, password_hash, role, is_admin, permissions) VALUES (?, ?, ?, ?, ?)",
                (username, generate_password_hash(password), role, is_admin, permissions)
            )

            user_id = cursor.lastrowid

            for gid in group_ids:
                if gid:
                    conn.execute("INSERT INTO user_groups (user_id, group_id) VALUES (?, ?)", (user_id, gid))

            conn.commit()
            log_action(
                current_username(),
                f"додав користувача: {username} (ID {user_id})",
                details=f"роль: {role}, груп: {len(group_ids)}"
            )
            flash('Користувача успішно додано', 'success')
            return redirect(url_for('admin.manage_users'))

        except sqlite3.IntegrityError as e:
            logger.error(f"Помилка БД при додаванні користувача '{username}': {e}", exc_info=True)
            flash(f'Помилка бази даних: {e}', 'danger')
        finally:
            conn.close()

    conn = get_db()
    try:
        groups = conn.execute("""
            SELECT id, name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
            FROM groups ORDER BY name, start_year
        """).fetchall()
    finally:
        conn.close()

    return render_template('add_user.html', groups=groups)


@admin_bp.route('/admin/users/<int:user_id>/edit', methods=['GET', 'POST'])
@permission_required('manage_users')
def edit_user(user_id):
    """Форма редагування ролі та прив'язаних груп існуючого користувача."""
    conn = get_db()
    try:
        if request.method == 'POST':
            role = request.form.get('role')
            group_ids = request.form.getlist('group_id')

            if not role:
                flash('Роль обовязкова', 'danger')
                return redirect(url_for('admin.edit_user', user_id=user_id))

            is_admin = 1 if role == 'admin' else 0
            user_row = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()

            conn.execute("UPDATE users SET role=?, is_admin=? WHERE id=?", (role, is_admin, user_id))
            conn.execute("DELETE FROM user_groups WHERE user_id=?", (user_id,))
            for gid in group_ids:
                if gid:
                    conn.execute("INSERT INTO user_groups (user_id, group_id) VALUES (?, ?)", (user_id, gid))

            conn.commit()
            log_action(
                current_username(),
                f"змінив роль/групи: {user_row['username']} (ID {user_id})",
                details=f"роль: {role}, груп: {len(group_ids)}"
            )
            flash('Дані користувача оновлено', 'success')
            return redirect(url_for('admin.manage_users'))

        user = conn.execute("SELECT id, username, role FROM users WHERE id=?", (user_id,)).fetchone()
        if not user:
            flash('Користувача не знайдено', 'danger')
            return redirect(url_for('admin.manage_users'))

        current_groups = conn.execute("SELECT group_id FROM user_groups WHERE user_id=?", (user_id,)).fetchall()
        current_group_ids = [row['group_id'] for row in current_groups]

        groups = conn.execute("""
            SELECT id, name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
            FROM groups ORDER BY name, start_year
        """).fetchall()

        return render_template('edit_user.html', user=user, groups=groups, current_group_ids=current_group_ids)
    finally:
        conn.close()


@admin_bp.route('/admin/users/<int:user_id>/change-password', methods=['GET', 'POST'])
@permission_required('manage_users')
def change_password(user_id):
    """Форма зміни пароля вказаного користувача адміністратором."""
    if request.method == 'POST':
        password = request.form.get('password')
        if not password or len(password) < 6:
            flash('Пароль повинен бути не коротшим 6 символів', 'danger')
            return redirect(url_for('admin.change_password', user_id=user_id))

        conn = get_db()
        try:
            target = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()
            conn.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(password), user_id))
            conn.commit()
            log_action(
                current_username(),
                f"змінив пароль: {target['username']} (ID {user_id})"
            )
            flash('Пароль успішно змінено', 'success')
            return redirect(url_for('admin.manage_users'))
        finally:
            conn.close()

    conn = get_db()
    try:
        user = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()
        if not user:
            flash('Користувача не знайдено', 'danger')
            return redirect(url_for('admin.manage_users'))
    finally:
        conn.close()

    return render_template('change_password.html', user_id=user_id, username=user['username'])


@admin_bp.route('/admin/users/<int:user_id>/delete', methods=['POST'])
@permission_required('manage_users')
def delete_user(user_id):
    """Видалення користувача та його зв'язків з групами (user_groups)."""
    conn = get_db()
    try:
        username_row = conn.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()
        if not username_row:
            flash('Користувача не знайдено', 'danger')
            return redirect(url_for('admin.manage_users'))

        conn.execute("DELETE FROM users WHERE id=?", (user_id,))
        conn.execute("DELETE FROM user_groups WHERE user_id=?", (user_id,))
        conn.commit()

        log_action(
            current_username(),
            f"ВИДАЛИВ користувача: {username_row['username']} (ID {user_id})"
        )
        flash('Користувача успішно видалено', 'success')
    except sqlite3.Error as e:
        logger.error(f"Помилка БД при видаленні користувача (ID {user_id}): {e}", exc_info=True)
        flash(f'Помилка при видаленні: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('admin.manage_users'))


@admin_bp.route('/admin/templates', methods=['GET', 'POST'])
@permission_required('manage_templates')
def manage_templates():
    """
    Сторінка управління Word-шаблонами (папка template_word/): перегляд
    списку, завантаження нового шаблону з необов'язковим описом і
    позначкою "тільки для адміністратора".
    """
    conn = get_db()

    if request.method == 'POST':
        file = request.files.get('template_file')
        display_name = (request.form.get('display_name') or '').strip()
        description = (request.form.get('description') or '').strip()
        admin_only = 1 if request.form.get('admin_only') == 'on' else 0
        # Галочка "Показувати в списках" (за замовчуванням увімкнена у формі):
        # знята галочка = шаблон прихований зі списків вибору при генерації.
        hidden = 0 if request.form.get('visible') == 'on' else 1

        if not file or file.filename == '':
            flash('Оберіть файл шаблону', 'danger')
            return redirect(url_for('admin.manage_templates'))

        filename = secure_filename(file.filename)
        if not filename.lower().endswith('.docx'):
            flash('Шаблон повинен бути файлом .docx', 'danger')
            return redirect(url_for('admin.manage_templates'))

        # Якщо назву не вказано - у списках вибору показується ім'я файлу
        if not display_name:
            display_name = filename

        os.makedirs(TEMPLATE_FOLDER, exist_ok=True)
        dest_path = os.path.join(TEMPLATE_FOLDER, filename)
        is_replace = os.path.exists(dest_path)

        try:
            file.save(dest_path)
            conn.execute("""
                INSERT INTO document_templates (filename, display_name, description, admin_only, hidden, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
                ON CONFLICT(filename) DO UPDATE SET
                    display_name = excluded.display_name,
                    description = excluded.description,
                    admin_only = excluded.admin_only,
                    hidden = excluded.hidden,
                    uploaded_by = excluded.uploaded_by,
                    uploaded_at = excluded.uploaded_at
            """, (filename, display_name, description, admin_only, hidden, current_username()))
            conn.commit()
            log_action(
                current_username(),
                f"{'оновив' if is_replace else 'завантажив новий'} шаблон документа: {filename}",
                details=f"тільки для адміністратора: {'так' if admin_only else 'ні'}"
            )
            flash(f"Шаблон «{filename}» успішно {'оновлено' if is_replace else 'завантажено'}", 'success')
        except Exception as e:
            logger.error(f"Помилка при завантаженні шаблону {filename}: {e}", exc_info=True)
            flash(f'Помилка при завантаженні шаблону: {e}', 'danger')
        finally:
            conn.close()

        return redirect(url_for('admin.manage_templates'))

    # -------------------- GET --------------------
    # Показуємо усі фізично наявні файли в template_word/, приєднуючи
    # метадані з БД, якщо вони є (файл міг бути покладений вручну, без
    # завантаження через цю сторінку - таким теж не повинно ламати список).
    rows = conn.execute("SELECT * FROM document_templates").fetchall()
    meta_by_filename = {row['filename']: dict(row) for row in rows}
    conn.close()

    templates = []
    if os.path.isdir(TEMPLATE_FOLDER):
        for f in sorted(os.listdir(TEMPLATE_FOLDER)):
            if not f.lower().endswith('.docx'):
                continue
            full_path = os.path.join(TEMPLATE_FOLDER, f)
            meta = meta_by_filename.get(f, {})
            templates.append({
                'filename': f,
                'display_name': meta.get('display_name') or f,
                'description': meta.get('description') or '',
                'admin_only': bool(meta.get('admin_only', 0)),
                'hidden': bool(meta.get('hidden', 0)),
                'uploaded_by': meta.get('uploaded_by') or '',
                'uploaded_at': meta.get('uploaded_at') or '',
                'size_kb': round(os.path.getsize(full_path) / 1024, 1),
            })

    return render_template('manage_templates.html', templates=templates)


@admin_bp.route('/admin/templates/<filename>/toggle_visibility', methods=['POST'])
@permission_required('manage_templates')
def toggle_template_visibility(filename):
    """
    Перемикає видимість шаблону в списках вибору при генерації
    (прихований <-> видимий), без потреби перезавантажувати файл.
    Для файлів, які лежать у template_word/ без запису в БД (додані
    вручну), запис створюється автоматично.
    """
    filename = secure_filename(filename)
    full_path = os.path.join(TEMPLATE_FOLDER, filename)

    if not os.path.isfile(full_path):
        flash(f"Файл шаблону «{filename}» не знайдено", 'danger')
        return redirect(url_for('admin.manage_templates'))

    conn = get_db()
    try:
        row = conn.execute("SELECT hidden FROM document_templates WHERE filename = ?", (filename,)).fetchone()
        if row is None:
            # Файл без метаданих (покладений вручну) - створюємо запис одразу прихованим
            conn.execute(
                "INSERT INTO document_templates (filename, display_name, hidden, uploaded_by) VALUES (?, ?, 1, ?)",
                (filename, filename, current_username())
            )
            new_hidden = 1
        else:
            new_hidden = 0 if row['hidden'] else 1
            conn.execute("UPDATE document_templates SET hidden = ? WHERE filename = ?", (new_hidden, filename))
        conn.commit()
        log_action(
            current_username(),
            f"{'приховав' if new_hidden else 'зробив видимим'} шаблон документа: {filename}"
        )
        flash(f"Шаблон «{filename}» тепер {'прихований зі' if new_hidden else 'видимий у'} списках вибору", 'success')
    except Exception as e:
        logger.error(f"Помилка при зміні видимості шаблону {filename}: {e}", exc_info=True)
        flash(f'Помилка при зміні видимості шаблону: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('admin.manage_templates'))


@admin_bp.route('/admin/templates/<filename>/download')
@permission_required('manage_templates')
def download_template(filename):
    """Віддає файл шаблону з template_word/ на завантаження (напр., щоб відредагувати його у Word і завантажити оновлену версію назад)."""
    filename = secure_filename(filename)
    full_path = os.path.join(TEMPLATE_FOLDER, filename)

    if not os.path.isfile(full_path):
        flash(f"Файл шаблону «{filename}» не знайдено", 'danger')
        return redirect(url_for('admin.manage_templates'))

    return send_file(full_path, as_attachment=True, download_name=filename)


@admin_bp.route('/admin/templates/<filename>/delete', methods=['POST'])
@permission_required('manage_templates')
def delete_template(filename):
    """Видаляє шаблон - сам файл із template_word/ та його метадані з БД."""
    filename = secure_filename(filename)
    full_path = os.path.join(TEMPLATE_FOLDER, filename)

    conn = get_db()
    try:
        conn.execute("DELETE FROM document_templates WHERE filename = ?", (filename,))
        conn.commit()
        if os.path.exists(full_path):
            os.remove(full_path)
        log_action(current_username(), f"видалив шаблон документа: {filename}")
        flash(f"Шаблон «{filename}» видалено", 'success')
    except Exception as e:
        logger.error(f"Помилка при видаленні шаблону {filename}: {e}", exc_info=True)
        flash(f'Помилка при видаленні шаблону: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('admin.manage_templates'))


@admin_bp.route('/admin/export_photos', methods=['GET', 'POST'])
@permission_required('group_export')
def export_photos():
    """
    Масове вивантаження фото студентів групи одним ZIP-архівом - для
    друку студентських квитків тощо. Можна забрати всіх студентів
    групи з фото одразу, або зняти позначку з окремих і завантажити
    лише вибраних. Кожен файл у архіві називається "Прізвище_Ім'я_По
    батькові.jpg" (по батькові пропускається, якщо не вказано) - готово
    вставляти в будь-яку програму верстки квитків без перейменування.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        group_id = request.form.get('group_id')
        student_ids = request.form.getlist('student_ids')
        if not student_ids:
            flash('Оберіть хоча б одного студента з фото', 'error')
            return redirect(url_for('admin.export_photos', group_id=group_id))

        placeholders = ','.join('?' for _ in student_ids)
        students = conn.execute(f"""
            SELECT id, last_name_UA, first_name_UA, middle_name_UA, photo
            FROM students WHERE id IN ({placeholders}) AND photo IS NOT NULL
        """, student_ids).fetchall()

        group = conn.execute("SELECT name FROM groups WHERE id=?", (group_id,)).fetchone()
        conn.close()

        if not students:
            flash('У жодного з обраних студентів немає завантаженого фото', 'error')
            return redirect(url_for('admin.export_photos', group_id=group_id))

        buffer = io.BytesIO()
        used_names = {}
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for s in students:
                photo_path = os.path.join('static', s['photo'])
                if not os.path.exists(photo_path):
                    continue
                ext = os.path.splitext(s['photo'])[1] or '.jpg'
                name_parts = [s['last_name_UA'], s['first_name_UA']]
                if s['middle_name_UA']:
                    name_parts.append(s['middle_name_UA'])
                base_name = '_'.join(p.strip() for p in name_parts if p and p.strip())
                # Про всяк випадок - якщо в групі раптом двоє тезок з
                # однаковим ПІБ, другий файл не повинен мовчки
                # перезаписати перший у архіві.
                arcname = f"{base_name}{ext}"
                if arcname in used_names:
                    used_names[arcname] += 1
                    arcname = f"{base_name}_{used_names[arcname]}{ext}"
                else:
                    used_names[arcname] = 1
                zf.write(photo_path, arcname=arcname)

        buffer.seek(0)
        log_action(current_username(), f"вивантажив фото студентів (ZIP): {len(students)} шт., група ID {group_id}")
        safe_group_name = (group['name'] if group else 'group').replace(' ', '_')
        download_name = f"Фото_{safe_group_name}_{datetime.now().strftime('%Y-%m-%d')}.zip"
        return send_file(buffer, as_attachment=True, download_name=download_name, mimetype='application/zip')

    # ====================== GET ======================
    groups = conn.execute("""
        SELECT id, name, start_year, study_form, program_credits,
               name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
        FROM groups WHERE archived = FALSE ORDER BY name COLLATE UKRAINIAN
    """).fetchall()

    selected_group_id = request.args.get('group_id', type=int)
    students = []
    if selected_group_id:
        students = conn.execute("""
            SELECT id, last_name_UA, first_name_UA, middle_name_UA, photo
            FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0
            ORDER BY last_name_UA COLLATE UKRAINIAN
        """, (selected_group_id,)).fetchall()
        students = sort_ukrainian(students, key_func=lambda s: f"{s['last_name_UA']} {s['first_name_UA']} {s['middle_name_UA']}")

    conn.close()
    return render_template(
        'export_photos.html',
        groups=groups, selected_group_id=selected_group_id, students=students,
    )


@admin_bp.route('/admin/group_export', methods=['GET', 'POST'])
@permission_required('group_export')
def group_export():
    """Сторінка вибору параметрів (група і/або рік народження, шаблон .docx) перед масовою генерацією документів студентів."""
    conn = get_db()
    conn.row_factory = sqlite3.Row

    groups = conn.execute("""
        SELECT id, name, start_year, study_form, program_credits,
               name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
        FROM groups WHERE archived = FALSE ORDER BY name, start_year
    """).fetchall()

    available_templates = get_templates_with_metadata(is_admin=session.get('is_admin', False))
    default_template = available_templates[0]['path'] if available_templates else ''

    current_year = datetime.now().year
    years = list(range(1980, current_year + 1))
    students = []
    selected_group_id = request.args.get('group_id', type=int) if request.method == 'GET' else request.form.get('group_id', type=int)
    selected_year = request.args.get('birth_year', type=int) if request.method == 'GET' else request.form.get('birth_year', type=int)
    selected_template = request.args.get('template', default_template) if request.method == 'GET' else request.form.get('template', default_template)

    if selected_group_id:
        group_check = conn.execute("SELECT id FROM groups WHERE id=? AND archived=FALSE", (selected_group_id,)).fetchone()
        if not group_check:
            flash('Обрана група не існує або є архівною.', 'error')
            selected_group_id = None

    if request.method == 'POST':
        if not selected_group_id and not selected_year:
            flash('Будь ласка, оберіть групу або рік народження.', 'error')
        else:
            active_students = request.form.getlist('active_students')
            return redirect(url_for('admin.generate_group_docs', group_id=selected_group_id,
                                    birth_year=selected_year, template=selected_template,
                                    active_students=','.join(active_students)))

    if selected_group_id or selected_year:
        base_query = "SELECT * FROM students WHERE archived = FALSE"
        params = []
        if selected_group_id:
            base_query += " AND group_id=?"
            params.append(selected_group_id)
        if selected_year:
            base_query += " AND SUBSTR(birth_date, 7, 4) >= ?"
            params.append(str(selected_year))
        try:
            students = conn.execute(base_query, params).fetchall()
        except Exception as e:
            logger.error(f"Помилка при отриманні студентів: {e}")
            conn.close()
            return "Помилка бази даних", 500

    conn.close()
    return render_template('group_export.html', students=students, groups=groups, years=years,
                           selected_group_id=selected_group_id, selected_year=selected_year,
                           selected_template=selected_template,
                           available_templates=available_templates)


UPLOAD_FOLDER = 'Uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    """Перевіряє, чи має файл дозволене розширення (.xlsx) для імпорту."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


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


@admin_bp.route('/admin/import_passport_documents', methods=['GET', 'POST'])
@permission_required('import_passport_documents')
def import_passport_documents():
    """
    Масовий імпорт паспортних даних з Excel: кожен рядок - один
    студент (знаходиться нечітким пошуком за ПІБ серед усіх активних
    студентів, routes/admin.fuzzy_find_student). Якщо студент уже має
    паспортні дані - додає ще один запис (історія), не перезаписує.
    """
    if request.method == 'POST':
        file = request.files.get('excel_file')
        if not file or not allowed_file(file.filename):
            flash("⚠️ Оберіть коректний файл .xlsx", "error")
            return redirect(url_for('admin.import_passport_documents'))

        filename = f"passport_import_{int(time.time())}.xlsx"
        filepath = os.path.join('static', 'uploads', 'tmp', filename)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        file.save(filepath)

        conn = get_db()
        cursor = conn.cursor()
        inserted, skipped = 0, 0
        try:
            wb = load_workbook(filepath)
            sheet = wb.active
            for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                if not row or all(cell is None for cell in row):
                    continue
                try:
                    fio = row[0]
                    document_type = str(row[1]).strip() if len(row) > 1 and row[1] else None
                    number = str(row[3]).strip() if len(row) > 3 and row[3] not in (None, '') else None

                    if not fio or not document_type or not number:
                        flash(f"❗ Рядок {i}: не заповнено ПІБ, тип документа чи номер - пропущено")
                        skipped += 1
                        continue
                    if document_type not in ('Паспорт (книжка)', 'ID-картка'):
                        flash(f"❗ Рядок {i}: невідомий тип документа '{document_type}' (очікується 'Паспорт (книжка)' або 'ID-картка') - пропущено")
                        skipped += 1
                        continue

                    student_id, matched_name, score = fuzzy_find_student(cursor, str(fio).strip())
                    if not student_id:
                        flash(f"❗ Рядок {i}: студента не знайдено за ПІБ '{fio}' - пропущено")
                        skipped += 1
                        continue

                    series = str(row[2]).strip() if len(row) > 2 and row[2] not in (None, '') else None
                    issued_by = str(row[4]).strip() if len(row) > 4 and row[4] not in (None, '') else None
                    issue_date = str(row[5]).strip().replace('-', '.') if len(row) > 5 and row[5] not in (None, '') else None
                    valid_until = str(row[6]).strip().replace('-', '.') if len(row) > 6 and row[6] not in (None, '') else None
                    unique_number = str(row[7]).strip() if len(row) > 7 and row[7] not in (None, '') else None

                    cursor.execute("""
                        INSERT INTO passport_documents (
                            student_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (student_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number))
                    inserted += 1
                except Exception as e:
                    logger.error(f"Row {i} error (passport import): {e}")
                    flash(f"❗ Рядок {i}: помилка обробки - {e}")
                    skipped += 1
                    continue

            conn.commit()
            log_action(
                current_username(),
                f"імпорт паспортних даних з Excel: додано {inserted}, пропущено {skipped}",
                details=f"файл: {filename}"
            )
            flash(f"✅ Імпорт завершено. Додано: {inserted}, пропущено: {skipped}", "success")
        except Exception as e:
            conn.rollback()
            flash(f"⚠️ Помилка при імпорті Excel: {e}", "error")
            logger.error(f"Error importing passport Excel: {e}")
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)
            conn.close()

        return redirect(url_for('admin.import_passport_documents'))

    return render_template('import_passport_documents.html')


# ============================================================
# Масова генерація документів у фоні - щоб не тримати HTTP-запит
# (і воркер waitress) заблокованим на весь час генерації, коли
# студентів у групі багато. Прогрес зберігається в пам'яті процесу
# (той самий підхід, що й _SESSIONS у office_editor.py) і опитується
# сторінкою прогресу через AJAX, поки не завершиться.
# ============================================================
_JOBS = {}
_JOBS_LOCK = threading.Lock()
_JOB_TTL_SECONDS = 3 * 3600


def _cleanup_old_jobs():
    now = time.time()
    with _JOBS_LOCK:
        stale = [jid for jid, j in _JOBS.items() if now - j['started_at'] > _JOB_TTL_SECONDS]
        for jid in stale:
            del _JOBS[jid]


def _run_generation_job(job_id, students_data, selected_template, batch_id, user_id, username, group_name, birth_year):
    """Виконується в окремому потоці: генерує документи по одному,
    оновлюючи прогрес у _JOBS[job_id] після кожного студента - навіть
    якщо один документ впаде з помилкою, решта продовжують генеруватись."""
    job = _JOBS[job_id]
    for student_dict, military_dict in students_data:
        with _JOBS_LOCK:
            job['current_name'] = f"{student_dict.get('last_name_UA','')} {student_dict.get('first_name_UA','')}"
        filename = f"{student_dict['last_name_UA']}_{student_dict['first_name_UA']}.docx".replace(" ", "_")
        full_path = os.path.join(office_editor.SESSIONS_DIR, f"{uuid.uuid4().hex}.docx")
        try:
            gen_doc(student_dict, military_dict, template=selected_template, out=full_path, user_name=username)
            doc_id = office_editor.create_editing_session(full_path, filename, user_id, batch_id=batch_id)
            with _JOBS_LOCK:
                job['succeeded'].append({
                    'doc_id': doc_id,
                    'name': f"{student_dict['last_name_UA']} {student_dict['first_name_UA']}",
                    'filename': filename,
                })
        except Exception as e:
            logger.error(f"Помилка при генерації документа для {student_dict.get('last_name_UA', '')}: {e}", exc_info=True)
            with _JOBS_LOCK:
                job['failed'].append({
                    'name': f"{student_dict.get('last_name_UA','')} {student_dict.get('first_name_UA','')}",
                    'error': str(e),
                })
        finally:
            with _JOBS_LOCK:
                job['done'] += 1

    with _JOBS_LOCK:
        job['complete'] = True
        job['current_name'] = ''

    log_action(
        username,
        f"масова генерація документів: {group_name}",
        details=f"шаблон: {selected_template}, рік нар.: {birth_year or 'всі'}, "
                f"успішно: {len(job['succeeded'])}, з помилкою: {len(job['failed'])}"
    )


@admin_bp.route('/admin/generate_group_docs', methods=['GET', 'POST'])
@permission_required('group_export')
def generate_group_docs():
    """Генерує .docx-документи (за обраним шаблоном) для всіх студентів, що підпадають під фільтр (група і/або рік народження) у фоновому потоці, показуючи прогрес-бар, а потім відкриває сторінку перегляду/редагування кожного в ONLYOFFICE перед завантаженням підсумкового ZIP-архіву (routes/office_editor.py)."""
    _cleanup_old_jobs()
    group_id = request.args.get('group_id', type=int) if request.method == 'GET' else request.form.get('group_id', type=int)
    birth_year = request.args.get('birth_year', type=int) if request.method == 'GET' else request.form.get('birth_year', type=int)
    selected_template = request.args.get('template', '') if request.method == 'GET' else request.form.get('template', '')
    active_students = request.args.get('active_students', '').split(',') if request.args.get('active_students') else []

    allowed_paths = {t['path'] for t in get_templates_with_metadata(is_admin=session.get('is_admin', False))}
    if selected_template not in allowed_paths:
        flash("У вас немає прав для генерації документів цим шаблоном", "danger")
        return redirect(url_for('admin.group_export'))

    if not group_id and not birth_year:
        flash('Оберіть групу або рік народження для генерації документів.', 'error')
        return redirect(url_for('admin.group_export'))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    base_query = """
        SELECT s.*,
               g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS group_name,
               g.study_form, g.start_year, g.program_credits,
               g.qualification_name, g.degree_level, g.specialty, g.educational_program, g.knowledge_area,
               g.qualification_name_en, g.degree_level_en, g.specialty_en, g.educational_program_en, g.knowledge_area_en,
               il.name_ua AS institution_name_and_status, il.name_en AS institution_name_and_status_en,
               il.short_name_ua AS license_short_name_ua, il.short_name_en AS license_short_name_en,
               g.entry_requirements, g.entry_requirements_en,
               g.learning_outcomes, g.learning_outcomes_en, g.program_includes, g.program_includes_en,
               g.entry_requirements_reduced, g.entry_requirements_reduced_en,
               g.learning_outcomes_reduced, g.learning_outcomes_reduced_en,
               g.program_includes_reduced, g.program_includes_reduced_en
        FROM students s LEFT JOIN groups g ON s.group_id = g.id
                         LEFT JOIN institution_licenses il ON s.license_id = il.id
        WHERE s.archived = FALSE
    """
    params = []
    if group_id:
        base_query += " AND s.group_id=?"
        params.append(group_id)
    if birth_year:
        base_query += " AND SUBSTR(s.birth_date, 7, 4) >= ?"
        params.append(str(birth_year))

    try:
        students = conn.execute(base_query, params).fetchall()
        if not students:
            conn.close()
            return "Студенты не найдены по заданным фильтрам", 404
    except Exception as e:
        logger.error(f"Ошибка при выполнении SQL-запроса: {e}")
        conn.close()
        return "Ошибка базы данных", 500

    if active_students and active_students[0]:
        students = [s for s in students if str(s['id']) in active_students]

    group_name = "Зі всіх груп"
    if group_id and students:
        group_name = students[0]['group_name'] if students[0]['group_name'] else f"Група_{group_id}"

    batch_id = office_editor.new_batch_id()

    # Забираємо всі дані студентів (і військові дані) заздалегідь, поки
    # з'єднання з базою відкрите в цьому запиті - фоновий потік більше
    # не звертатиметься до conn з цього обробника (SQLite-з'єднання
    # прив'язане до потоку, у якому було створене).
    students_data = []
    for student in students:
        student_dict = dict(student)
        military = conn.execute("SELECT * FROM military WHERE student_id=?", (student['id'],)).fetchone()
        students_data.append((student_dict, dict(military) if military else {}))
    conn.close()

    job_id = uuid.uuid4().hex
    with _JOBS_LOCK:
        _JOBS[job_id] = {
            'total': len(students_data),
            'done': 0,
            'current_name': '',
            'succeeded': [],
            'failed': [],
            'complete': False,
            'batch_id': batch_id,
            'group_name': group_name,
            'started_at': time.time(),
        }

    thread = threading.Thread(
        target=_run_generation_job,
        args=(job_id, students_data, selected_template, batch_id, session['user_id'],
              current_username(), group_name, birth_year),
        daemon=True,
    )
    thread.start()

    return render_template('group_generate_progress.html', job_id=job_id, group_name=group_name, total=len(students_data))


@admin_bp.route('/admin/generate_group_docs/status/<job_id>')
@permission_required('group_export')
def generate_group_docs_status(job_id):
    """JSON-статус фонової генерації - опитується сторінкою прогресу через AJAX."""
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if not job:
            return {'error': 'not_found'}, 404
        return {
            'total': job['total'],
            'done': job['done'],
            'current_name': job['current_name'],
            'complete': job['complete'],
            'succeeded_count': len(job['succeeded']),
            'failed_count': len(job['failed']),
        }


@admin_bp.route('/admin/generate_group_docs/result/<job_id>')
@permission_required('group_export')
def generate_group_docs_result(job_id):
    """Показує підсумок завершеної фонової генерації: перелік готових документів (з переходом у ONLYOFFICE) і, за наявності, список тих, що не вдалося згенерувати."""
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job or not job['complete']:
        flash('Завдання генерації не знайдено або ще не завершено', 'warning')
        return redirect(url_for('admin.group_export'))

    if not job['succeeded']:
        flash('Не вдалося згенерувати жодного документа', 'danger')
        return redirect(url_for('admin.group_export'))

    return render_template(
        'group_docs_preview.html',
        items=job['succeeded'],
        failed=job['failed'],
        batch_id=job['batch_id'],
        group_name=job['group_name'],
    )


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


TEMP_PREVIEW_FOLDER = "temp_preview"

def save_preview_to_file(preview):
    """Зберігає проміжний результат парсингу Excel-файлу імпорту документів про освіту у тимчасовий JSON-файл (щоб не тримати великі дані в сесії) і повертає його ID."""
    os.makedirs(TEMP_PREVIEW_FOLDER, exist_ok=True)
    preview_id = str(uuid.uuid4())
    path = os.path.join(TEMP_PREVIEW_FOLDER, f"{preview_id}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(preview, f, ensure_ascii=False)
    return preview_id

def load_preview_from_file(preview_id):
    """Завантажує раніше збережений (save_preview_to_file) прев'ю-результат імпорту за його ID."""
    path = os.path.join(TEMP_PREVIEW_FOLDER, f"{preview_id}.json")
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def fuzzy_find_student(cursor, full_name, threshold=80):
    """Нечіткий пошук студента за повним ім'ям (ПІБ) серед активних студентів, використовуючи rapidfuzz; повертає ID найкращого збігу, якщо він проходить поріг схожості threshold."""
    cursor.execute("SELECT id, last_name_UA, first_name_UA, middle_name_UA FROM students WHERE archived=FALSE")
    students = cursor.fetchall()
    names = ["{} {} {}".format(s["last_name_UA"], s["first_name_UA"], s["middle_name_UA"]) for s in students]
    matches = process.extract(full_name, names, scorer=fuzz.ratio, limit=3)
    for match_name, score, idx in matches:
        if score >= threshold:
            return students[idx]["id"], match_name, score
    return None, None, None

def translate_to_en(text):
    """Перекладає текст на англійську через GoogleTranslator з кешуванням результатів (translation_cache), щоб не робити повторні запити до перекладача для однакового тексту."""
    if not text:
        return ""
    if text in translation_cache:
        return translation_cache[text]
    try:
        result = translator.translate(text)
        translation_cache[text] = result
        return result
    except Exception as e:
        logger.warning(f"Помилка перекладу '{text}': {e}")
        return text
        
def find_country(text):
    """Визначає країну (укр./англ. назву) за ключовими словами в тексті установи/довідки; за замовчуванням вважає, що це Україна."""
    countries = {
        "польща": "Poland",
        "poland": "Poland",
        "німеччина": "Germany",
        "germany": "Germany",
        "чех": "Czech Republic",
        "czech": "Czech Republic",
        "словач": "Slovakia",
        "slovakia": "Slovakia"
    }

    lower = text.lower()

    for key, en in countries.items():
        if key in lower:
            return key.capitalize(), en

    if "україна " in lower:
        return "Україна ", "Ukraine"

    return "Україна", "Ukraine"

def parse_document(text: str):
    """Розбирає текстовий опис документа про освіту (формат 'Тип документа Номер; ДД.ММ.РРРР; Ким видано: Установа') на структуровані поля (тип, номер, дата, установа, країна) для імпорту."""
    text = text.strip()
    pattern = re.compile(
        r'^(?P<type>[^;]+?)\s*;\s*(?P<date>\d{2}\.\d{2}\.\d{4})\s*;\s*Ким видано:\s*(?P<institution>.+?)$',
        re.IGNORECASE | re.UNICODE
    )
    match = pattern.search(text)
    if not match:
        return None

    full_prefix = match.group("type").strip()
    parts = re.split(r'\s{2,}', full_prefix.strip())
    if len(parts) < 2:
        doc_type = full_prefix
        doc_number = ""
    else:
        doc_type = " ".join(parts[:-1]).strip()
        doc_number = parts[-1].strip()

    completion_date = match.group("date").strip()
    completion_date = completion_date.replace(".", "/")
    institution = match.group("institution")
    country, country_en = find_country(institution)

    return {
        "document_type": doc_type,
        "document_type_en": translate_to_en(doc_type) or "",
        "document_number": doc_number,
        "completion_date": completion_date,
        "institution_name": institution,
        "institution_name_en": translate_to_en(institution) or "",
        "country": country,
        "country_en": country_en
    }

def parse_reference_cell_ua(text):
    """Розбирає комірку Excel з даними довідки про визнання документа (номер; установа; країна; дата) на окремі поля."""
    if not text or not str(text).strip():
        return {
            "reference_number": "",
            "reference_institution": "",
            "reference_institution_en": "",
            "reference_country": "",
            "reference_country_en": "",
            "reference_issue_date": "",
        }
    parts = [p.strip() for p in str(text).split(";")]
    parts += [""] * (4 - len(parts))
    reference_number, reference_institution, reference_country, reference_issue_date = parts[:4]
    reference_issue_date = reference_issue_date.replace(".", "/")
    return {
        "reference_number": reference_number,
        "reference_institution": reference_institution,
        "reference_institution_en": translate_to_en(reference_institution) or "",
        "reference_country": reference_country,
        "reference_country_en": translate_to_en(reference_country) or "",
        "reference_issue_date": reference_issue_date,
    }

def parse_recognition_cell_ua(text):
    """Розбирає комірку Excel з даними про визнання (номер сертифіката; орган; дата) на окремі поля."""
    if not text or not str(text).strip():
        return {
            "recognition_certificate_number": "",
            "recognition_issuer": "",
            "recognition_issuer_en": "",
            "recognition_date": "",
        }
    parts = [p.strip() for p in str(text).split(";")]
    parts += [""] * (3 - len(parts))
    recognition_certificate_number, recognition_issuer, recognition_date = parts[:3]
    recognition_date = recognition_date.replace(".", "/")
    return {
        "recognition_certificate_number": recognition_certificate_number,
        "recognition_issuer": recognition_issuer,
        "recognition_issuer_en": translate_to_en(recognition_issuer) or "",
        "recognition_date": recognition_date,
    }

def import_documents_preview(file_path, db):
    """Читає Excel-файл з документами про освіту, для кожного рядка знаходить студента (fuzzy_find_student), парсить документ (parse_document) та формує прев'ю-список змін (нові/оновлення) для підтвердження користувачем перед комітом у import_docs_commit."""
    wb = load_workbook(file_path)
    sheet = wb.active
    cursor = db.cursor()
    preview_rows = []
    for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        fio = row[0]
        document_text = row[1]
        reference_ua_text = row[2] if len(row) > 2 else None
        recognition_ua_text = row[3] if len(row) > 3 else None

        if not fio or not document_text:
            preview_rows.append({"row_index": row_index, "error": "Пусті дані"})
            continue
        student_id, matched_name, score = fuzzy_find_student(cursor, fio)
        if not student_id:
            preview_rows.append({"row_index": row_index, "error": f"Студент не знайдений: {fio}"})
            continue
        data = parse_document(document_text)
        if not data:
            preview_rows.append({"row_index": row_index, "error": f"Не вдалося розпізнати документ: {document_text}"})
            continue

        data.update(parse_reference_cell_ua(reference_ua_text))
        data.update(parse_recognition_cell_ua(recognition_ua_text))


        cursor.execute("""
            SELECT id, document_type, document_type_en, document_number, completion_date,
                   institution_name, institution_name_en, country, country_en
            FROM education_documents WHERE student_id=?
        """, (student_id,))
        existing_doc = cursor.fetchone()

        row_data = {"row_index": row_index, "student_id": student_id, "matched_name": matched_name, "score": score, **data}

        if existing_doc:
            row_data.update({
                "status": "Оновлення", "status_class": "text-warning",
                "existing_info": f"(існує: {existing_doc['document_number'] or 'без номера'})",
                "has_document": True,
                "existing_doc_id": existing_doc["id"],
                "old_document_type": existing_doc["document_type"] or "",
                "old_document_type_en": existing_doc["document_type_en"] or "",
                "old_document_number": existing_doc["document_number"] or "",
                "old_completion_date": existing_doc["completion_date"] or "",
                "old_institution_name": existing_doc["institution_name"] or "",
                "old_institution_name_en": existing_doc["institution_name_en"] or "",
                "old_country": existing_doc["country"] or "",
                "old_country_en": existing_doc["country_en"] or "",
            })

            cursor.execute("""
                SELECT reference_number, reference_institution, reference_institution_en,
                       reference_country, reference_country_en, reference_issue_date,
                       recognition_certificate_number, recognition_issuer, recognition_issuer_en, recognition_date
                FROM foreign_education_docs WHERE education_doc_id=?
            """, (existing_doc["id"],))
            existing_foreign = cursor.fetchone()
            if existing_foreign:
                row_data.update({
                    "old_reference_number": existing_foreign["reference_number"] or "",
                    "old_reference_institution": existing_foreign["reference_institution"] or "",
                    "old_reference_institution_en": existing_foreign["reference_institution_en"] or "",
                    "old_reference_country": existing_foreign["reference_country"] or "",
                    "old_reference_country_en": existing_foreign["reference_country_en"] or "",
                    "old_reference_issue_date": existing_foreign["reference_issue_date"] or "",
                    "old_recognition_certificate_number": existing_foreign["recognition_certificate_number"] or "",
                    "old_recognition_issuer": existing_foreign["recognition_issuer"] or "",
                    "old_recognition_issuer_en": existing_foreign["recognition_issuer_en"] or "",
                    "old_recognition_date": existing_foreign["recognition_date"] or "",
                })
        else:
            row_data.update({"status": "Новий", "status_class": "text-success", "existing_info": "", "has_document": False})

        preview_rows.append(row_data)
    return preview_rows
    
@admin_bp.route('/admin/import_education_docs_preview', methods=['GET', 'POST'])
@permission_required('import_education_docs')
def import_docs_preview():
    """Приймає завантажений Excel-файл, будує прев'ю через import_documents_preview і показує сторінку підтвердження перед фактичним імпортом."""
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            flash("Файл не обрано", "danger")
            return redirect(url_for('admin.manage_education_documents'))

        filename = secure_filename(file.filename)
        path = os.path.join("uploads", filename)
        os.makedirs("uploads", exist_ok=True)
        file.save(path)

        db = get_db()
        preview = import_documents_preview(path, db)
        preview_id = save_preview_to_file(preview)

        return render_template("import_education_preview.html", preview=preview, preview_id=preview_id)

    return render_template("import_education_upload.html")

@admin_bp.route('/admin/import_docs_commit', methods=['POST'])
@permission_required('import_education_docs')
def import_docs_commit():
    """Записує в БД документи про освіту (та дані про визнання), підтверджені користувачем на сторінці прев'ю (тільки позначені чекбоксом рядки)."""
    preview_id = request.form.get("preview_id")
    preview = load_preview_from_file(preview_id)
    db = get_db()
    added = 0
    updated = 0

    for row in preview:
        row_index = row["row_index"]
        if f"add_{row_index}" not in request.form:
            continue

        if not row.get("error"):
            row["document_type"]       = request.form.get(f"document_type_{row_index}",       row["document_type"])
            row["document_type_en"]    = request.form.get(f"document_type_en_{row_index}",    row.get("document_type_en", "")) or ""
            row["document_number"]     = request.form.get(f"document_number_{row_index}",     row["document_number"])
            row["completion_date"]     = request.form.get(f"completion_date_{row_index}",     row["completion_date"])
            row["institution_name"]    = request.form.get(f"institution_name_{row_index}",    row["institution_name"])
            row["institution_name_en"] = request.form.get(f"institution_name_en_{row_index}", row.get("institution_name_en", "")) or ""
            row["country"]             = request.form.get(f"country_{row_index}",             row["country"])
            row["country_en"]          = request.form.get(f"country_en_{row_index}",          row["country_en"])

            reference_number = request.form.get(f"reference_number_{row_index}", row.get("reference_number", "")) or None
            reference_institution = request.form.get(f"reference_institution_{row_index}", row.get("reference_institution", "")) or None
            reference_institution_en = request.form.get(f"reference_institution_en_{row_index}", row.get("reference_institution_en", "")) or None
            reference_country = request.form.get(f"reference_country_{row_index}", row.get("reference_country", "")) or None
            reference_country_en = request.form.get(f"reference_country_en_{row_index}", row.get("reference_country_en", "")) or None
            reference_issue_date = request.form.get(f"reference_issue_date_{row_index}", row.get("reference_issue_date", "")) or None
            recognition_certificate_number = request.form.get(f"recognition_certificate_number_{row_index}", row.get("recognition_certificate_number", "")) or None
            recognition_issuer = request.form.get(f"recognition_issuer_{row_index}", row.get("recognition_issuer", "")) or None
            recognition_issuer_en = request.form.get(f"recognition_issuer_en_{row_index}", row.get("recognition_issuer_en", "")) or None
            recognition_date = request.form.get(f"recognition_date_{row_index}", row.get("recognition_date", "")) or None

            country_lower = (row["country"] or "").lower()

            is_foreign = not (country_lower in ('україна', 'ukraine'))

            cursor = db.cursor()
            cursor.execute("SELECT id FROM education_documents WHERE student_id=?", (row["student_id"],))
            existing = cursor.fetchone()

            doc_type_en  = row.get("document_type_en", "") or ""
            inst_name_en = row.get("institution_name_en", "") or ""
            country_en   = row.get("country_en", "") or ""

            if existing:
                education_doc_id = existing["id"]
                cursor.execute("""
                    UPDATE education_documents SET
                        document_type=?, document_type_en=?, document_number=?, completion_date=?,
                        institution_name=?, institution_name_en=?, country=?, country_en=?
                    WHERE id=?
                """, (row["document_type"], doc_type_en, row["document_number"], row["completion_date"],
                      row["institution_name"], inst_name_en, row["country"], row["country_en"], education_doc_id))
                updated += 1
            else:
                cursor.execute("""
                    INSERT INTO education_documents (
                        student_id, document_type, document_type_en, document_number,
                        completion_date, institution_name, institution_name_en, country, country_en
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (row["student_id"], row["document_type"], doc_type_en, row["document_number"],
                      row["completion_date"], row["institution_name"], inst_name_en, row["country"], country_en))
                education_doc_id = cursor.lastrowid
                added += 1

            # --- Foreign / recognition data ---
            cursor.execute(
                "SELECT id FROM foreign_education_docs WHERE education_doc_id = ?",
                (education_doc_id,)
            )
            foreign_exists = cursor.fetchone()

            # Если заполнена хотя бы одна колонка аккредитации/признания
            has_foreign_data = any([
                reference_number,
                reference_institution,
                reference_country,
                reference_issue_date,
                recognition_certificate_number,
                recognition_issuer,
                recognition_date
            ])

            if foreign_exists:
                if has_foreign_data:
                    cursor.execute("""
                        UPDATE foreign_education_docs SET
                            reference_number=?, reference_institution=?, reference_institution_en=?,
                            reference_country=?, reference_country_en=?, reference_issue_date=?,
                            recognition_certificate_number=?, recognition_issuer=?,
                            recognition_issuer_en=?, recognition_date=?
                        WHERE education_doc_id=?
                    """, (
                        reference_number,
                        reference_institution,
                        reference_institution_en,
                        reference_country,
                        reference_country_en,
                        reference_issue_date,
                        recognition_certificate_number,
                        recognition_issuer,
                        recognition_issuer_en,
                        recognition_date,
                        education_doc_id
                    ))
                else:
                    # если все поля пустые - удаляем старую запись
                    cursor.execute(
                        "DELETE FROM foreign_education_docs WHERE education_doc_id=?",
                        (education_doc_id,)
                    )

            else:
                if has_foreign_data:
                    cursor.execute("""
                        INSERT INTO foreign_education_docs (
                            education_doc_id,
                            reference_number,
                            reference_institution,
                            reference_institution_en,
                            reference_country,
                            reference_country_en,
                            reference_issue_date,
                            recognition_certificate_number,
                            recognition_issuer,
                            recognition_issuer_en,
                            recognition_date
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        education_doc_id,
                        reference_number,
                        reference_institution,
                        reference_institution_en,
                        reference_country,
                        reference_country_en,
                        reference_issue_date,
                        recognition_certificate_number,
                        recognition_issuer,
                        recognition_issuer_en,
                        recognition_date
                    ))

    db.commit()

    log_action(
        current_username(),
        f"імпорт документів про освіту: додано {added}, оновлено {updated}"
    )

    flash(f"Додано записів: {added}, Оновлено записів: {updated}", "success")
    return redirect(url_for('admin.manage_education_documents'))