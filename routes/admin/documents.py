"""
routes/admin/documents.py
==========================
Документи про освіту й паспортні дані студентів (+ скани-вкладення),
масовий імпорт документів про освіту (двоетапний: попередній перегляд
-> підтвердження) і паспортних даних (одноетапний), разом із
допоміжними функціями розпізнавання Excel-комірок для імпорту.
"""
from flask import render_template, request, redirect, url_for, flash, session, send_file, jsonify
from routes.db import get_db
from routes.utils import log_action, permission_required, logger
from routes.helpers import current_username, sort_ukrainian
from routes.admin import admin_bp
import sqlite3
import os
from routes.utils import get_attachments, save_multiple_attachments
from routes.admin import allowed_file
from werkzeug.utils import secure_filename
import pandas as pd
import json
import uuid
import re
import time
from datetime import datetime, date
from openpyxl import load_workbook
from rapidfuzz import process, fuzz
import requests
from routes.config import DEEPL_API_KEY
from routes.passport_types import (
    PASSPORT_TYPES, PASSPORT_TYPES_ID_STYLE, PASSPORT_TYPE_DEFAULT,
    normalize_passport_type, normalize_passport_number,
    passport_template_context,
)
from routes.enrollments_service import ensure_enrollments_table

def _excel_date_to_ua(value):
    """Excel повертає datetime для клітинок з форматом дати.
    str(datetime) → '2020-05-15 00:00:00' — треба прибрати час і переставити
    складові на ДД.ММ.РРРР. Рядки лишаємо як є (але з '.' замість '-')."""
    if value in (None, ''):
        return None
    if isinstance(value, datetime):
        return value.strftime('%d.%m.%Y')
    if isinstance(value, date):
        return value.strftime('%d.%m.%Y')
    return str(value).strip().replace('-', '.')


ALLOWED_EXTENSIONS = {'xlsx'}

# DeepL - офіційний платний/безкоштовний API (ключ з .env, див.
# routes/config.py DEEPL_API_KEY) замість GoogleTranslator/MyMemory -
# ті були безкоштовними неофіційними сервісами, які агресивно блокували
# за частотою запитів (TooManyRequests), тому відмовились від обох.
#
# Запит робимо напряму через requests, а не через бібліотеку
# deep_translator.DeeplTranslator - вона досі авторизується старим
# способом (auth_key у query-параметрах GET-запиту), який DeepL більше
# не приймає: сервер у відповідь повертає 403 "Missing Authorization
# header, expected 'Authorization: DeepL-Auth-Key <API key>'". Ключ при
# цьому цілком робочий - перевірено напряму через curl з тим самим
# ключем, спрацювало лише з заголовком Authorization.
DEEPL_API_URL = "https://api-free.deepl.com/v2/translate"
translation_cache = {}

# Папка для тимчасових JSON-файлів, у яких між кроком попереднього
# перегляду і кроком підтвердження зберігається розібраний Excel-файл
# імпорту документів про освіту. Тримаємо в static/uploads/tmp/ - там
# уже лежать інші тимчасові файли імпортів (напр. паспортні дані), і
# ця папка вже існує/створюється; прямий доступ через /static/ до неї
# закритий в app.py (block_templates_static_access).
TEMP_PREVIEW_FOLDER = os.path.join('static', 'uploads', 'tmp', 'preview')

@admin_bp.route('/admin/manage_education_documents', methods=['GET', 'POST'])
@permission_required('manage_education_documents')
def manage_education_documents():
    """CRUD-сторінка документів про освіту студентів (атестат/диплом попереднього рівня) разом із додатковими даними про визнання іноземного документа (foreign_education_docs)."""
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
            SELECT DISTINCT s.id, s.last_name_UA, s.first_name_UA
            FROM students s
            LEFT JOIN student_enrollments e
                   ON e.student_id = s.id
                  AND e.group_id = ?
                  AND e.status = 'active'
            WHERE COALESCE(s.archived, 0) = 0
              AND (e.id IS NOT NULL OR s.group_id = ?)
              AND s.id NOT IN (
                  SELECT student_id FROM education_documents
                  WHERE group_id = ? OR group_id IS NULL
              )
            ORDER BY s.last_name_UA COLLATE UKRAINIAN, s.first_name_UA
        """, (selected_group_id, selected_group_id, selected_group_id))
        students_without_docs = cursor.fetchall()

    # Основний запит документів
    cursor.execute("""
        SELECT COALESCE(ed.group_id, s.group_id) AS group_id,
               g.name AS group_name, s.id AS student_id,
               s.last_name_UA, s.first_name_UA,
               ed.id AS doc_id, ed.group_id AS doc_group_id,
               ed.document_type, ed.document_type_en, ed.document_number,
               ed.institution_name, ed.institution_name_en, ed.country, ed.country_en, ed.completion_date,
               fed.reference_number, fed.reference_institution, fed.reference_institution_en,
               fed.reference_country, fed.reference_country_en, fed.reference_issue_date,
               fed.recognition_certificate_number, fed.recognition_issuer,
               fed.recognition_issuer_en, fed.recognition_date
        FROM education_documents ed
        INNER JOIN students s ON ed.student_id = s.id
        LEFT JOIN groups g ON g.id = COALESCE(ed.group_id, s.group_id)
        LEFT JOIN foreign_education_docs fed ON ed.id = fed.education_doc_id
        WHERE s.archived = FALSE AND (g.archived = FALSE OR g.archived IS NULL)
        ORDER BY g.name, s.last_name_UA, s.first_name_UA, ed.id
    """)
    rows = cursor.fetchall()
    
        # Enrollment'и обраного студента — для випадаючого списку груп
    # у формі додавання/редагування документа (щоб при паралельному
    # навчанні документ можна було прив'язати до конкретної групи).
    student_enrollments = []
    if student:
        try:
            from routes.enrollments_service import list_student_enrollments
            student_enrollments = [
                e for e in list_student_enrollments(student['id'], conn=db)
                if e.get('group_id')
            ]
        except Exception as logger_exc:
            logger.warning(f"Не вдалося завантажити enrollments: {logger_exc}")

    documents_by_group = {}
    for row in rows:
        gid = row['group_id']
        if gid not in documents_by_group:
            documents_by_group[gid] = {
                'group_name': row['group_name'] or 'Без групи',
                'docs': [],
            }
        doc = dict(row)
        doc['attachments'] = get_attachments(db, 'education_document', doc['doc_id'])
        documents_by_group[gid]['docs'].append(doc)

    # ── Безпечне сортування ──
    # group_name може бути None, якщо:
    #   - студент без групи (заморожений, в архіві), і
    #   - документ не прив'язаний до групи (group_id IS NULL у education_documents);
    #   - групу, до якої був прив'язаний документ, видалили, а FK не спрацював.
    # Без цього `sorted` падає з TypeError: '<' not supported between str and None.
    sorted_documents_by_group = sorted(
        documents_by_group.items(),
        key=lambda x: (x[1]['group_name'] or '').lower(),
    )

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
                doc_group_id = request.form.get('doc_group_id', type=int)

                cursor.execute("""
                    UPDATE education_documents SET
                        student_id=?, group_id=?, document_type=?, document_type_en=?, document_number=?,
                        institution_name=?, institution_name_en=?, country=?, country_en=?, completion_date=?
                    WHERE id=?
                """, (student_id, doc_group_id, request.form.get('document_type'), request.form.get('document_type_en'),
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
                if len(new_scans) > MAX_FILES:
                    raise ValueError(f"Забагато файлів за раз - максимум {MAX_FILES} за одну загрузку")
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

                doc_group_id = request.form.get('doc_group_id', type=int)

                cursor.execute("""
                    INSERT INTO education_documents (
                        student_id, group_id, document_type, document_type_en, document_number,
                        institution_name, institution_name_en, country, country_en, completion_date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (student_id, doc_group_id, document_type, document_type_en, document_number,
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
                if len(new_scans) > MAX_FILES:
                    raise ValueError(f"Забагато файлів за раз - максимум {MAX_FILES} за одну загрузку")
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
        student_doc_id=student_doc_id,
        student_enrollments=student_enrollments,
    )

def _normalize_name_for_match(s):
    """
    Нормалізує ПІБ для нечіткого порівняння:
      - нижній регістр;
      - усі варіанти апострофа → один символ;
      - усі варіанти дефіса/тире → пробіл;
      - пунктуація → пробіл;
      - зайві пробіли → один.
    """
    if not s:
        return ''
    s = str(s).lower().strip()
    for ch in ('\u2019', '\u02BC', '\u2018', '`', '\u02BB', '\u2032', '\u055A', "'"):
        s = s.replace(ch, "'")
    for ch in ('\u2014', '\u2013', '\u2012', '\u2212', '\uFF0D', '\u2011', '-'):
        s = s.replace(ch, ' ')
    s = ''.join(c if (c.isalnum() or c.isspace() or c == "'") else ' ' for c in s)
    return ' '.join(s.split())


def fuzzy_find_student(cursor, full_name, threshold=95):
    """
    Нечіткий пошук студента за ПІБ.

    Чому саме так:
      * нормалізація регістру / апострофів / дефісів / пробілів — прибирає
        розбіжності, які в старій версії «з'їдали» 5–20 пунктів схожості;
      * token_set_ratio — коротший ПІБ (напр. без по батькові) повністю
        «входить» у довший → 100. token_sort_ratio цей кейс НЕ вирішує
        (дає ~74 через різницю в довжині рядків);
      * поріг 90 — свідомо високий, бо token_set_ratio і так дає 100 для
        будь-якого «subset»-матчу; нижчий поріг почне пропускати
        «іваненко» → «іващенко» як хибні збіги;
      * захист від неоднозначності: якщо 2+ студенти мають однаково
        високий бал (типова ситуація, коли у файлі лише «Прізвище Ім'я»
        без по батькові, а в базі двоє таких) — повертаємо None, щоб
        рядок позначився як «студента не знайдено», а не підставив
        першого-ліпшого.
    """
    if full_name is None:
        return None, None, None
    target = _normalize_name_for_match(full_name)
    if not target:
        return None, None, None

    cursor.execute(
        "SELECT id, last_name_UA, first_name_UA, middle_name_UA "
        "FROM students WHERE archived=FALSE"
    )
    students = cursor.fetchall()
    if not students:
        return None, None, None

    names = []
    for s in students:
        parts = (s["last_name_UA"], s["first_name_UA"], s["middle_name_UA"])
        joined = ' '.join(str(p) for p in parts if p and str(p).strip())
        names.append(_normalize_name_for_match(joined))

    # 1) точний збіг після нормалізації — найнадійніший
    for idx, name in enumerate(names):
        if name == target:
            return students[idx]["id"], name, 100.0

    # 2) нечітке — token_set_ratio (subset-толерантний)
    matches = process.extract(
        target, names, scorer=fuzz.token_set_ratio, limit=3,
    )
    if not matches:
        return None, None, None

    best_name, best_score, best_idx = matches[0]
    if best_score < threshold:
        return None, None, None

    # 3) Захист від неоднозначності: якщо «другий претендент» має майже
    #    такий самий бал — це найчастіше «Прізвище Ім'я» без по батькові
    #    при двох студентах з таким ПІБ. Не вгадуємо — повертаємо «не знайдено».
    if len(matches) > 1:
        second_name, second_score, _ = matches[1]
        if second_score >= threshold and (best_score - second_score) < 5:
            return None, None, None

    return students[best_idx]["id"], best_name, best_score


@admin_bp.route('/admin/manage_passport_documents', methods=['GET', 'POST'])
@permission_required('manage_passport_documents')
def manage_passport_documents():
    """CRUD-сторінка паспортних даних студентів (паспорт-книжка або
    ID-картка) - за тим самим зразком, що й документи про освіту, але
    без підтаблиці "іноземний документ" і з лімітом 5 сканів на запис."""
    db = get_db()
    cursor = db.cursor()
    # Ліміт лише на ОДНУ загрузку. Загальна кількість сканів на
    # запис не обмежується — можна додавати новими партіями.
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
        # УВАГА: паспорти прив'язані до студента (passport_documents.student_id),
        # а не до групи — тому перевіряємо наявність БУДЬ-ЯКОГО паспорта.
        # Раніше тут помилково запитувалась education_documents (копіпаст),
        # через що дропдаун показував не тих студентів.
        cursor.execute("""
            SELECT DISTINCT s.id, s.last_name_UA, s.first_name_UA
            FROM students s
            LEFT JOIN student_enrollments e
                   ON e.student_id = s.id
                  AND e.group_id = ?
                  AND e.status = 'active'
            WHERE COALESCE(s.archived, 0) = 0
              AND (e.id IS NOT NULL OR s.group_id = ?)
              AND s.id NOT IN (
                  SELECT student_id FROM passport_documents
                  WHERE group_id = ? OR group_id IS NULL
              )
            ORDER BY s.last_name_UA COLLATE UKRAINIAN, s.first_name_UA
        """, (selected_group_id, selected_group_id, selected_group_id))
        students_without_docs = cursor.fetchall()

    cursor.execute("""
        SELECT COALESCE(pd.group_id, s.group_id) AS group_id,
               g.name AS group_name, s.id AS student_id,
               s.last_name_UA, s.first_name_UA,
               pd.id AS doc_id, pd.group_id AS doc_group_id,
               pd.document_type, pd.series, pd.number,
               pd.issued_by, pd.issue_date, pd.valid_until, pd.unique_number
        FROM passport_documents pd
        INNER JOIN students s ON pd.student_id = s.id
        LEFT JOIN groups g ON g.id = COALESCE(pd.group_id, s.group_id)
        WHERE s.archived = FALSE AND (g.archived = FALSE OR g.archived IS NULL)
        ORDER BY g.name, s.last_name_UA, s.first_name_UA, pd.id
    """)
    rows = cursor.fetchall()

    documents_by_group = {}
    for row in rows:
        gid = row['group_id']
        if gid not in documents_by_group:
            documents_by_group[gid] = {
                'group_name': row['group_name'] or 'Без групи',
                'docs': [],
            }
        doc = dict(row)
        doc['attachments'] = get_attachments(db, 'passport_document', doc['doc_id'])
        documents_by_group[gid]['docs'].append(doc)

    # group_name може бути None: паспорт прив'язаний до групи, якої вже
    # немає (видалена), або у студента немає основної групи (архівний/
    # заморожений без enrollment'ів). Без `or ''` sorted() падає з
    # TypeError: '<' not supported between instances of 'str' and 'NoneType'.
    sorted_documents_by_group = sorted(
        documents_by_group.items(),
        key=lambda x: (x[1]['group_name'] or '').lower(),
    )

    # Enrollments усіх студентів у поточному списку — для селектора
    # «Для якої групи» в модалках редагування.
    student_ids_in_view = set()
    for gdata in documents_by_group.values():
        for doc in gdata['docs']:
            student_ids_in_view.add(doc['student_id'])

    enrollments_by_student = {}
    if student_ids_in_view:
        placeholders = ','.join('?' * len(student_ids_in_view))
        enroll_rows = db.execute(f"""
            SELECT e.student_id, e.group_id, e.is_primary, e.status,
                   g.name AS group_name
            FROM student_enrollments e
            JOIN groups g ON g.id = e.group_id
            WHERE e.student_id IN ({placeholders})
            ORDER BY e.is_primary DESC, e.id
        """, list(student_ids_in_view)).fetchall()
        for r in enroll_rows:
            enrollments_by_student.setdefault(r['student_id'], []).append(dict(r))

    # Для add-форми (коли прийшли з картки студента)
    student_enrollments = []
    if student:
        student_enrollments = enrollments_by_student.get(student['id'], [])
    
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

                if document_type not in PASSPORT_TYPES:
                    raise ValueError("Невірний тип документа")

                new_scans = [f for f in request.files.getlist('document_scans') if f and f.filename]
                # Ліміт — лише на одноразову загрузку (не більше
                # MAX_FILES нових файлів за раз). Загальна кількість
                # сканів на запис не обмежується.
                if len(new_scans) > MAX_FILES:
                    flash(f'Забагато файлів за раз - максимум {MAX_FILES} за одну загрузку (обрано {len(new_scans)}).', 'error')
                    return redirect(url_for('admin.manage_passport_documents', group_id=selected_group_id))

                # "Дійсний до" стосується лише ID-картки (книжка в Україні
                # безстрокова) - JS ховає поле для книжки, але сервер не
                # повинен покладатись лише на це (форму можна надіслати і
                # з вимкненим JS), тому підстраховуємось і тут.
                valid_until = request.form.get('valid_until') or None if document_type in PASSPORT_TYPES_ID_STYLE else None

                doc_group_id = request.form.get('doc_group_id', type=int)

                cursor.execute("""
                    UPDATE passport_documents SET
                        student_id=?, group_id=?, document_type=?, series=?, number=?,
                        issued_by=?, issue_date=?, valid_until=?, unique_number=?
                    WHERE id=?
                """, (
                    student_id, doc_group_id, document_type, request.form.get('series') or None,
                    normalize_passport_number(request.form.get('number'), document_type),
                    request.form.get('issued_by') or None, request.form.get('issue_date') or None,
                    valid_until, request.form.get('unique_number') or None,
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
                if document_type not in PASSPORT_TYPES:
                    raise ValueError("Невірний тип документа")

                new_scans = [f for f in request.files.getlist('document_scans') if f and f.filename]
                if len(new_scans) > MAX_FILES:
                    raise ValueError(f"Забагато файлів за раз - максимум {MAX_FILES} за одну загрузку")

                valid_until = request.form.get('valid_until') or None if document_type in PASSPORT_TYPES_ID_STYLE else None

                doc_group_id = request.form.get('doc_group_id', type=int)

                cursor.execute("""
                    INSERT INTO passport_documents (
                        student_id, group_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    student_id, doc_group_id, document_type, request.form.get('series') or None,
                    normalize_passport_number(request.form.get('number'), document_type),
                    request.form.get('issued_by') or None, request.form.get('issue_date') or None,
                    valid_until, request.form.get('unique_number') or None,
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
        student_enrollments=student_enrollments,
        enrollments_by_student=enrollments_by_student,
        **passport_template_context(),
    )


# routes/admin/documents.py — на початку файлу вже є: import os, json, uuid, time
# додайте константу поруч із TEMP_PREVIEW_FOLDER:

PASSPORT_REPORTS_DIR = os.path.join(os.getcwd(), 'static', 'uploads', 'import_reports')

# Каталог для сесій імпорту паспортів (між кроками)
PASSPORT_IMPORT_SESSIONS_DIR = os.path.join(
    os.getcwd(), 'static', 'uploads', 'import_sessions', 'passports'
)
PASSPORT_IMPORT_TTL = 6 * 3600  # 6 годин


def _cleanup_passport_import_sessions():
    if not os.path.isdir(PASSPORT_IMPORT_SESSIONS_DIR):
        return
    now = time.time()
    for name in os.listdir(PASSPORT_IMPORT_SESSIONS_DIR):
        path = os.path.join(PASSPORT_IMPORT_SESSIONS_DIR, name)
        try:
            if os.path.isfile(path) and now - os.path.getmtime(path) > PASSPORT_IMPORT_TTL:
                os.remove(path)
        except OSError:
            pass


def _passport_state_path(token):
    token = re.sub(r'[^0-9a-f]', '', token or '')
    return os.path.join(PASSPORT_IMPORT_SESSIONS_DIR, f'{token}.json')


def _save_passport_state(token, state):
    os.makedirs(PASSPORT_IMPORT_SESSIONS_DIR, exist_ok=True)
    with open(_passport_state_path(token), 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, default=str)


def _load_passport_state(token):
    token = re.sub(r'[^0-9a-f]', '', token or '')
    if not token:
        return None
    path = _passport_state_path(token)
    if not os.path.isfile(path):
        return None
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def _determine_target_enrollment_for_passport(conn, student_id, group_from_excel):
    """
    Визначає цільову групу для паспорта.

    Пріоритет:
      1. Якщо в колонці I явно вказано group_id — перевіряємо, що це
         активна група студента (primary або enrollment). Якщо ні —
         повертаємо reason='excel_group_invalid' (помилка рядка).
      2. Інакше — активні enrollments студента:
         - 0 → fallback на students.group_id (legacy)
         - 1 → цей enrollment
         - >1 → needs_choice=True, у превью треба обрати вручну
    """
    if group_from_excel:
        ok = conn.execute("""
            SELECT 1 FROM groups g WHERE g.id = ? AND COALESCE(g.archived, 0) = 0
              AND (
                  ? = (SELECT group_id FROM students WHERE id = ?)
                  OR EXISTS (
                      SELECT 1 FROM student_enrollments e
                      WHERE e.student_id = ? AND e.group_id = ?
                        AND e.status = 'active'
                  )
              )
        """, (group_from_excel, group_from_excel, student_id,
              student_id, group_from_excel)).fetchone()
        if not ok:
            return {
                'target_group_id': None,
                'groups_available': [],
                'needs_choice': False,
                'reason': 'excel_group_invalid',
            }
        return {
            'target_group_id': int(group_from_excel),
            'groups_available': [],
            'needs_choice': False,
            'reason': 'excel_column',
        }

    rows = conn.execute("""
        SELECT e.group_id, e.is_primary, g.name AS group_name
        FROM student_enrollments e
        JOIN groups g ON g.id = e.group_id
        WHERE e.student_id = ? AND e.status = 'active'
        ORDER BY e.is_primary DESC, e.id
    """, (student_id,)).fetchall()

    if not rows:
        row = conn.execute("""
            SELECT s.group_id, g.name AS group_name
            FROM students s LEFT JOIN groups g ON g.id = s.group_id
            WHERE s.id = ?
        """, (student_id,)).fetchone()
        if row and row['group_id']:
            return {
                'target_group_id': row['group_id'],
                'groups_available': [{
                    'group_id': row['group_id'],
                    'group_name': row['group_name'] or f"#{row['group_id']}",
                    'is_primary': True,
                }],
                'needs_choice': False,
                'reason': 'legacy_primary',
            }
        return {
            'target_group_id': None,
            'groups_available': [],
            'needs_choice': False,
            'reason': 'no_active_enrollments',
        }

    if len(rows) == 1:
        r = rows[0]
        return {
            'target_group_id': r['group_id'],
            'groups_available': [{
                'group_id': r['group_id'],
                'group_name': r['group_name'] or f"#{r['group_id']}",
                'is_primary': bool(r['is_primary']),
            }],
            'needs_choice': False,
            'reason': 'single_enrollment',
        }

    return {
        'target_group_id': None,
        'groups_available': [{
            'group_id': r['group_id'],
            'group_name': r['group_name'] or f"#{r['group_id']}",
            'is_primary': bool(r['is_primary']),
        } for r in rows],
        'needs_choice': True,
        'reason': 'multiple_enrollments',
    }


def _build_passport_import_preview(filepath, conn):
    """Парсить Excel, будує превью (БЕЗ ЗАПИСУ В БД)."""
    preview = []
    global_errors = []

    try:
        wb = load_workbook(filepath, data_only=True, read_only=True)
        sheet = wb.active
    except Exception as e:
        logger.error(f'[passport_import] open xlsx: {e}', exc_info=True)
        global_errors.append({'reason': f'Не вдалося відкрити файл: {e}'})
        return preview, global_errors

    row_counter = 0
    try:
        for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            row_counter += 1
            if not row or all(cell is None for cell in row):
                continue
            item = _build_passport_row_preview(i, row, conn)
            if item:
                preview.append(item)
    finally:
        try:
            wb.close()
        except Exception:
            pass

    if row_counter == 0:
        global_errors.append({'reason': 'У файлі немає рядків з даними'})

    return preview, global_errors


def _build_passport_row_preview(i, row, conn):
    fio = row[0] if len(row) > 0 else None
    document_type = str(row[1]).strip() if len(row) > 1 and row[1] else None
    series = str(row[2]).strip() if len(row) > 2 and row[2] not in (None, '') else None
    number = str(row[3]).strip() if len(row) > 3 and row[3] not in (None, '') else None
    issued_by = str(row[4]).strip() if len(row) > 4 and row[4] not in (None, '') else None
    issue_date = _excel_date_to_ua(row[5]) if len(row) > 5 else None
    valid_until = _excel_date_to_ua(row[6]) if len(row) > 6 else None
    unique_number = str(row[7]).strip() if len(row) > 7 and row[7] not in (None, '') else None
    group_id_raw = row[8] if len(row) > 8 else None
    number = normalize_passport_number(number, document_type)
    base = {
        'row': i,
        'fio_excel': str(fio or '').strip() or None,
        'student_id': None,
        'student_name': None,
        'matched_score': None,
        'document_type': document_type,
        'series': series,
        'number': number,
        'issued_by': issued_by,
        'issue_date': issue_date,
        'valid_until': valid_until,
        'unique_number': unique_number,
        'requested_group_id': None,
        'target_group_id': None,
        'groups_available': [],
        'needs_choice': False,
        'target_reason': None,
        'status': 'error',
        'error': None,
        'warning': None,
        'existing_same_type': None,
    }
  
    # ── Нормалізуємо номер: Excel втрачає провідні нулі (039628 → 39628).
    # Робимо це ДО перевірки заповненості й ДО пошуку дубліката, щоб
    # "39628" і "039628" вважались одним паспортом.
    if document_type in PASSPORT_TYPES:
        normalized_number = normalize_passport_number(number, document_type)
        if normalized_number and normalized_number != number:
            base['number_original'] = number  # для інформації
            base['number'] = normalized_number
            number = normalized_number  # для подальших перевірок

    if not fio or not document_type or not number:
        base['error'] = 'Не заповнено ПІБ, тип документа чи номер'
        return base

    if document_type not in PASSPORT_TYPES:
        base['error'] = f"Невідомий тип документа: «{document_type}»"
        return base

    if group_id_raw not in (None, ''):
        try:
            base['requested_group_id'] = int(group_id_raw)
        except (ValueError, TypeError):
            base['error'] = f"Некоректний ID групи в колонці I: «{group_id_raw}»"
            return base

    # Пошук студента
    student_id, matched_name, score = fuzzy_find_student(conn.cursor(), str(fio).strip())
    if not student_id:
        base['error'] = 'Студента не знайдено за ПІБ (або неоднозначно)'
        return base

    base['student_id'] = student_id
    base['student_name'] = matched_name
    base['matched_score'] = round(score, 1) if score is not None else None

    # valid_until має сенс лише для ID-типів
    if document_type not in PASSPORT_TYPES_ID_STYLE:
        base['valid_until'] = None
    
    

    # Цільова група
    target_info = _determine_target_enrollment_for_passport(
        conn, student_id, base['requested_group_id']
    )

    if target_info['reason'] == 'excel_group_invalid':
        base['error'] = (
            f"Вказана в колонці I група ID {base['requested_group_id']} "
            f"не належить студенту або недоступна"
        )
        return base

    base['target_group_id'] = target_info['target_group_id']
    base['groups_available'] = target_info['groups_available']
    base['needs_choice'] = target_info['needs_choice']
    base['target_reason'] = target_info['reason']

    # Дублікат (той самий паспорт уже є в цій групі).
    # Порівнюємо через NORM_PASSPORT з ОБОХ боків — на випадок, якщо
    # в БД лежать старі записи без нормалізації ("39628" замість
    # "039628"). Після міграції (init_db.py) це не критично, але
    # захищаємось і для ненормалізованих даних.
    if base['target_group_id'] is not None:
        duplicate = conn.execute("""
            SELECT id FROM passport_documents
            WHERE student_id = ?
              AND COALESCE(group_id, -1) = COALESCE(?, -1)
              AND document_type = ?
              AND COALESCE(series, '') = COALESCE(?, '')
              AND NORM_PASSPORT(number, document_type) = NORM_PASSPORT(?, document_type)
        """, (student_id, base['target_group_id'], document_type, series, number)).fetchone()
        if duplicate:
            base['status'] = 'skip'
            base['warning'] = (
                f'Паспорт з таким самим номером уже є в цій групі '
                f'(ID запису {duplicate["id"]}) — буде пропущено'
            )
            return base

    # Попередження: є інший паспорт того самого типу в цій групі
    if base['target_group_id'] is not None:
        same_type = conn.execute("""
            SELECT id, number, issued_by, issue_date FROM passport_documents
            WHERE student_id = ?
              AND COALESCE(group_id, -1) = COALESCE(?, -1)
              AND document_type = ?
            ORDER BY id DESC LIMIT 1
        """, (student_id, base['target_group_id'], document_type)).fetchone()
        if same_type:
            base['existing_same_type'] = {
                'number': same_type['number'] or '',
                'issued_by': same_type['issued_by'] or '',
                'issue_date': same_type['issue_date'] or '',
            }
            base['warning'] = (
                f"У студента вже є «{document_type}» у цій групі "
                f"(номер {same_type['number'] or '—'}, виданий {same_type['issue_date'] or '—'}). "
                f"Цей паспорт буде ДОДАНО ЯК НОВИЙ ЗАПИС (історія)."
            )

    base['status'] = 'new'
    return base

@admin_bp.route('/admin/import_passport_documents', methods=['GET', 'POST'])
@permission_required('import_passport_documents')
def import_passport_documents():
    _cleanup_passport_import_sessions()

    if request.method == 'POST':
        file = request.files.get('excel_file')
        if not file or not allowed_file(file.filename):
            flash("⚠️ Оберіть коректний файл .xlsx", "error")
            return redirect(url_for('admin.import_passport_documents'))

        filename = secure_filename(file.filename)
        filepath = os.path.join('uploads', f'{uuid.uuid4().hex}_{filename}')
        os.makedirs('uploads', exist_ok=True)
        file.save(filepath)

        conn = get_db()
        try:
            ensure_enrollments_table(conn)
            preview, global_errors = _build_passport_import_preview(filepath, conn)
        finally:
            try:
                os.remove(filepath)
            except OSError:
                pass
            conn.close()

        if global_errors:
            for ge in global_errors:
                flash(ge['reason'], 'error')
            return redirect(url_for('admin.import_passport_documents'))

        if not preview:
            flash('У файлі не знайдено жодного рядка для обробки', 'warning')
            return redirect(url_for('admin.import_passport_documents'))

        token = uuid.uuid4().hex[:32]
        _save_passport_state(token, {
            'preview': preview,
            'filename': filename,
            'user_id': session.get('user_id'),
            'created_at': time.time(),
        })

        return redirect(url_for('admin.import_passport_preview', token=token))

    return render_template('import_passport_documents.html', **passport_template_context())


@admin_bp.route('/admin/import_passport_documents/preview/<token>')
@permission_required('import_passport_documents')
def import_passport_preview(token):
    state = _load_passport_state(token)
    if not state:
        flash('Сесію імпорту не знайдено або вона застаріла', 'warning')
        return redirect(url_for('admin.import_passport_documents'))

    if state.get('user_id') != session.get('user_id') and not session.get('is_admin'):
        flash('Сесію імпорту створено іншим користувачем', 'error')
        return redirect(url_for('admin.import_passport_documents'))

    preview = state['preview']
    stats = {
        'new': sum(1 for p in preview if p['status'] == 'new'),
        'skip': sum(1 for p in preview if p['status'] == 'skip'),
        'errors': sum(1 for p in preview if p['status'] == 'error'),
        'needs_choice': sum(1 for p in preview if p.get('needs_choice')),
    }

    return render_template(
        'import_passport_preview.html',
        token=token,
        preview=preview,
        stats=stats,
        filename=state.get('filename', ''),
    )


@admin_bp.route('/admin/import_passport_documents/commit/<token>', methods=['POST'])
@permission_required('import_passport_documents')
def import_passport_commit(token):
    state = _load_passport_state(token)
    if not state:
        flash('Сесію імпорту не знайдено або вона застаріла', 'warning')
        return redirect(url_for('admin.import_passport_documents'))

    if state.get('user_id') != session.get('user_id') and not session.get('is_admin'):
        flash('Сесію імпорту створено іншим користувачем', 'error')
        return redirect(url_for('admin.import_passport_documents'))

    # Валідація: якщо needs_choice і чекбокс стоїть — radio обов'язковий
    missing_choices = []
    for item in state['preview']:
        if item['status'] != 'new':
            continue
        if not item.get('needs_choice'):
            continue
        row = item['row']
        if request.form.get(f'apply_{row}') and not request.form.get(f'target_group_{row}'):
            missing_choices.append(f"рядок {row}")

    if missing_choices:
        flash(
            'Не обрано групу для: ' + ', '.join(missing_choices) +
            '. Оберіть групу біля кожного рядка з паралельним навчанням.',
            'error',
        )
        return redirect(url_for('admin.import_passport_preview', token=token))

    conn = get_db()
    stats = {'inserted': 0, 'skipped': 0, 'errors': 0, 'warnings': []}

    try:
        for item in state['preview']:
            row = item['row']

            if item['status'] == 'error':
                stats['errors'] += 1
                continue

            if item['status'] == 'skip':
                stats['skipped'] += 1
                continue

            if not request.form.get(f'apply_{row}'):
                stats['skipped'] += 1
                continue

            # Визначаємо цільову групу
            if item.get('needs_choice'):
                tgid_raw = request.form.get(f'target_group_{row}')
                try:
                    target_group_id = int(tgid_raw) if tgid_raw else None
                except (ValueError, TypeError):
                    target_group_id = None
                if not target_group_id:
                    stats['warnings'].append({
                        'row': row,
                        'fio': item.get('fio_excel') or '',
                        'reason': 'не обрано групу — пропущено',
                    })
                    stats['skipped'] += 1
                    continue
            else:
                target_group_id = item.get('target_group_id')

            try:
                conn.execute("""
                    INSERT INTO passport_documents (
                        student_id, group_id, document_type, series, number,
                        issued_by, issue_date, valid_until, unique_number
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    item['student_id'], target_group_id,
                    item['document_type'], item['series'],
                    normalize_passport_number(item['number'], item['document_type']),
                    item['issued_by'], item['issue_date'],
                    item['valid_until'], item['unique_number'],
                ))
                stats['inserted'] += 1
            except sqlite3.IntegrityError as e:
                logger.warning(f'[passport_import] row {row} integrity: {e}')
                stats['warnings'].append({
                    'row': row,
                    'fio': item.get('fio_excel') or '',
                    'reason': f'Не вдалося додати (можливо, дублікат): {e}',
                })
                stats['skipped'] += 1
            except Exception as e:
                logger.error(f'[passport_import] row {row}: {e}', exc_info=True)
                stats['warnings'].append({
                    'row': row,
                    'fio': item.get('fio_excel') or '',
                    'reason': f'Помилка: {e}',
                })
                stats['skipped'] += 1

        conn.commit()
    except Exception as e:
        conn.rollback()
        logger.error(f'[passport_import] global: {e}', exc_info=True)
        flash(f'Помилка імпорту: {e}', 'error')
        conn.close()
        return redirect(url_for('admin.import_passport_preview', token=token))
    finally:
        conn.close()

    log_action(
        current_username(),
        f"імпорт паспортних даних з Excel: додано {stats['inserted']}, "
        f"пропущено {stats['skipped']}, помилок {stats['errors']}",
        details=f"файл: {state.get('filename')}, попереджень: {len(stats['warnings'])}",
    )

    try:
        os.remove(_passport_state_path(token))
    except OSError:
        pass

    report_id = None
    if stats['warnings'] or stats['errors']:
        os.makedirs(PASSPORT_REPORTS_DIR, exist_ok=True)
        report_id = uuid.uuid4().hex
        with open(os.path.join(PASSPORT_REPORTS_DIR, f'passport_{report_id}.json'), 'w', encoding='utf-8') as f:
            json.dump({
                'filename': state.get('filename'),
                'inserted': stats['inserted'],
                'skipped': stats['skipped'],
                'skipped_details': [
                    {'row': p['row'], 'fio': p.get('fio_excel'), 'reason': p.get('error')}
                    for p in state['preview'] if p['status'] == 'error'
                ] + stats['warnings'],
                'when': datetime.now().strftime('%d.%m.%Y %H:%M'),
            }, f, ensure_ascii=False)
        return redirect(url_for('admin.import_passport_report', report_id=report_id))

    flash(
        f"✅ Імпорт завершено. Додано: {stats['inserted']}, пропущено: {stats['skipped']}",
        'success',
    )
    return redirect(url_for('admin.import_passport_documents'))

@admin_bp.route('/admin/import_passport_documents/report/<report_id>')
@permission_required('import_passport_documents')
def import_passport_report(report_id):
    """Детальний звіт — окрема сторінка, щоб не тримати 1300+ flash у сесії."""
    report_id = re.sub(r'[^0-9a-f]', '', report_id or '')
    path = os.path.join(PASSPORT_REPORTS_DIR, f'passport_{report_id}.json')
    if not report_id or not os.path.isfile(path):
        flash('Звіт не знайдено або він застарів', 'warning')
        return redirect(url_for('admin.import_passport_documents'))
    with open(path, encoding='utf-8') as f:
        report = json.load(f)
    return render_template('import_passport_report.html', report=report)
    
    

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


def translate_to_en(text):
    """
    Перекладає текст на англійську через офіційний DeepL API (прямий
    HTTP-запит - див. коментар біля DEEPL_API_URL вище, чому не через
    бібліотеку deep_translator), з кешуванням результатів
    (translation_cache), щоб не робити повторний запит до перекладача
    для однакового тексту вдруге.

    Якщо DEEPL_API_KEY не задано (routes/config.py вже попередив про
    це в лозі при старті) або сам запит впав (напр. вичерпано місячну
    квоту безкоштовного плану - 500 000 символів) - повертає порожній
    рядок, щоб поле просто лишилось незаповненим і адмін вписав
    переклад вручну, а не щоб застосунок падав.
    """
    if not text:
        return ""
    if text in translation_cache:
        return translation_cache[text]
    if not DEEPL_API_KEY:
        return ""

    try:
        response = requests.post(
            DEEPL_API_URL,
            headers={"Authorization": f"DeepL-Auth-Key {DEEPL_API_KEY}"},
            data={"text": text, "target_lang": "EN"},
            timeout=10,
        )
        response.raise_for_status()
        result = response.json()["translations"][0]["text"]
        translation_cache[text] = result
        return result
    except Exception as e:
        logger.warning(f"Помилка перекладу DeepL '{text}': {e}")
        return ""


@admin_bp.route('/admin/translate_text', methods=['POST'])
@permission_required('manage_education_documents')
def translate_text_endpoint():
    """
    AJAX-ендпоінт для автозаповнення англійських полів у формі документа
    про освіту (Тип документа/Назва закладу/Країна) - перевикористовує
    той самий translate_to_en() з кешуванням, що й масовий імпорт, тому
    однакові формулювання (напр. "Диплом молодшого спеціаліста") не
    викликають DeepL повторно.
    """
    text = (request.get_json(silent=True) or {}).get('text', '').strip()
    if not text:
        return jsonify({'translation': ''})
    return jsonify({'translation': translate_to_en(text)})


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

def _determine_target_enrollment_for_education_doc(cursor, student_id, group_from_excel):
    """
    Визначає цільову групу для документа про освіту.

    Пріоритет:
      1. Якщо в колонці E явно вказано group_id — перевіряємо, що це
         активна група студента.
      2. Інакше — активні enrollments:
         - 0 → fallback на students.group_id (legacy)
         - 1 → цей enrollment (автоматично)
         - >1 → needs_choice=True (адмін обирає радіо-кнопкою)
    """
    if group_from_excel:
        ok = cursor.execute("""
            SELECT 1 FROM groups g WHERE g.id = ? AND COALESCE(g.archived, 0) = 0
              AND (
                  ? = (SELECT group_id FROM students WHERE id = ?)
                  OR EXISTS (
                      SELECT 1 FROM student_enrollments e
                      WHERE e.student_id = ? AND e.group_id = ?
                        AND e.status = 'active'
                  )
              )
        """, (group_from_excel, group_from_excel, student_id,
              student_id, group_from_excel)).fetchone()
        if not ok:
            return {
                'target_group_id': None,
                'groups_available': [],
                'needs_choice': False,
                'reason': 'excel_group_invalid',
            }
        return {
            'target_group_id': int(group_from_excel),
            'groups_available': [],
            'needs_choice': False,
            'reason': 'excel_column',
        }

    rows = cursor.execute("""
        SELECT e.group_id, e.is_primary, g.name AS group_name
        FROM student_enrollments e
        JOIN groups g ON g.id = e.group_id
        WHERE e.student_id = ? AND e.status = 'active'
        ORDER BY e.is_primary DESC, e.id
    """, (student_id,)).fetchall()

    if not rows:
        row = cursor.execute("""
            SELECT s.group_id, g.name AS group_name
            FROM students s LEFT JOIN groups g ON g.id = s.group_id
            WHERE s.id = ?
        """, (student_id,)).fetchone()
        if row and row['group_id']:
            return {
                'target_group_id': row['group_id'],
                'groups_available': [{
                    'group_id': row['group_id'],
                    'group_name': row['group_name'] or f"#{row['group_id']}",
                    'is_primary': True,
                }],
                'needs_choice': False,
                'reason': 'legacy_primary',
            }
        return {
            'target_group_id': None,
            'groups_available': [],
            'needs_choice': False,
            'reason': 'no_active_enrollments',
        }

    if len(rows) == 1:
        r = rows[0]
        return {
            'target_group_id': r['group_id'],
            'groups_available': [{
                'group_id': r['group_id'],
                'group_name': r['group_name'] or f"#{r['group_id']}",
                'is_primary': bool(r['is_primary']),
            }],
            'needs_choice': False,
            'reason': 'single_enrollment',
        }

    return {
        'target_group_id': None,
        'groups_available': [{
            'group_id': r['group_id'],
            'group_name': r['group_name'] or f"#{r['group_id']}",
            'is_primary': bool(r['is_primary']),
        } for r in rows],
        'needs_choice': True,
        'reason': 'multiple_enrollments',
    }


def import_documents_preview(file_path, db):
    """Читає Excel, будує прев'ю (БЕЗ ЗАПИСУ В БД)."""
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

        student_id, matched_name, score = fuzzy_find_student(
            cursor, str(fio).strip() if fio is not None else None
        )

        # ── Колонка E (опційна): ID групи документа ──
        doc_group_id_raw = row[4] if len(row) > 4 else None
        requested_group_id = None
        if doc_group_id_raw not in (None, ''):
            try:
                requested_group_id = int(doc_group_id_raw)
            except (ValueError, TypeError):
                preview_rows.append({
                    "row_index": row_index,
                    "error": (f"Некоректний ID групи документа: "
                              f"'{doc_group_id_raw}'"),
                })
                continue

        if not student_id:
            preview_rows.append({
                "row_index": row_index,
                "error": f"Студент не знайдений: {fio}",
            })
            continue

        data = parse_document(document_text)
        if not data:
            preview_rows.append({
                "row_index": row_index,
                "error": f"Не вдалося розпізнати документ: {document_text}",
            })
            continue

        data.update(parse_reference_cell_ua(reference_ua_text))
        data.update(parse_recognition_cell_ua(recognition_ua_text))

        # ── Визначаємо цільову групу ──
        target_info = _determine_target_enrollment_for_education_doc(
            cursor, student_id, requested_group_id
        )

        if target_info['reason'] == 'excel_group_invalid':
            preview_rows.append({
                "row_index": row_index,
                "error": (
                    f"Вказана в колонці E група ID {requested_group_id} "
                    f"не належить студенту або недоступна"
                ),
            })
            continue

        target_group_id = target_info['target_group_id']
        groups_available = target_info['groups_available']
        needs_choice = target_info['needs_choice']
        target_reason = target_info['reason']

        # ── Пошук існуючого документа ──
        # Для needs_choice (група не визначена) документ не шукаємо —
        # під час commit його пошукаємо вже для обраної групи.
        existing_doc = None
        if not needs_choice:
            if target_group_id is not None:
                cursor.execute("""
                    SELECT id, document_type, document_type_en, document_number,
                           completion_date, institution_name, institution_name_en,
                           country, country_en, group_id
                    FROM education_documents
                    WHERE student_id = ? AND group_id = ?
                    ORDER BY id DESC LIMIT 1
                """, (student_id, target_group_id))
                existing_doc = cursor.fetchone()
            else:
                # Немає груп у студента — шукаємо legacy (group_id IS NULL)
                cursor.execute("""
                    SELECT id, document_type, document_type_en, document_number,
                           completion_date, institution_name, institution_name_en,
                           country, country_en, group_id
                    FROM education_documents
                    WHERE student_id = ? AND group_id IS NULL
                    ORDER BY id DESC LIMIT 1
                """, (student_id,))
                existing_doc = cursor.fetchone()

        row_data = {
            "row_index": row_index,
            "student_id": student_id,
            "matched_name": matched_name,
            "score": round(score, 1) if score is not None else None,
            "requested_group_id": requested_group_id,
            "target_group_id": target_group_id,
            "groups_available": groups_available,
            "needs_choice": needs_choice,
            "target_reason": target_reason,
            **data,
        }

        if existing_doc:
            row_data.update({
                "status": "Оновлення", "status_class": "text-warning",
                "existing_info": f"(існує: {existing_doc['document_number'] or 'без номера'})",
                "has_document": True,
                "existing_doc_id": existing_doc["id"],
                "existing_group_id": existing_doc["group_id"],
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
                       recognition_certificate_number, recognition_issuer,
                       recognition_issuer_en, recognition_date
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
            row_data.update({
                "status": "Новий", "status_class": "text-success",
                "existing_info": "", "has_document": False,
            })

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

        # Список усіх активних груп — для випадаючого списку
        # «Група документа» в модалці предпросмотра (важливо для
        # паралельного навчання: диплом бакалавра → магістратура).
        cursor = db.cursor()
        cursor.execute(
            "SELECT id, name FROM groups WHERE archived = FALSE "
            "ORDER BY name COLLATE UKRAINIAN"
        )
        all_groups = [dict(r) for r in cursor.fetchall()]

        return render_template(
            "import_education_preview.html",
            preview=preview,
            preview_id=preview_id,
            all_groups=all_groups,
        )

    return render_template("import_education_upload.html")


@admin_bp.route('/admin/import_docs_commit', methods=['POST'])
@permission_required('import_education_docs')
def import_docs_commit():
    """Записує в БД документи про освіту, підтверджені користувачем на сторінці прев'ю."""
    preview_id = request.form.get("preview_id")
    preview = load_preview_from_file(preview_id)
    db = get_db()
    added = 0
    updated = 0
    skipped_no_group = []

    try:
        edits_json = json.loads(request.form.get("edits_json") or "{}")
    except (ValueError, TypeError):
        edits_json = {}
    if not isinstance(edits_json, dict):
        edits_json = {}

    for row in preview:
        row_index = row["row_index"]
        if f"add_{row_index}" not in request.form:
            continue

        # ── Застосовуємо редагування з модалки (без групи — група
        #    обирається радіо-кнопкою в таблиці, а не в модалці) ──
        edits_for_row = edits_json.get(str(row_index)) or edits_json.get(row_index)
        if edits_for_row:
            for _field in (
                'document_type', 'document_type_en',
                'institution_name', 'institution_name_en',
                'document_number', 'completion_date',
                'country', 'country_en',
                'reference_number', 'reference_institution',
                'reference_institution_en', 'reference_country',
                'reference_country_en', 'reference_issue_date',
                'recognition_certificate_number', 'recognition_issuer',
                'recognition_issuer_en', 'recognition_date',
            ):
                if _field in edits_for_row:
                    row[_field] = edits_for_row[_field]

        if row.get("error"):
            continue

        # ── Визначаємо цільову групу ──
        if row.get("needs_choice"):
            tgid_raw = request.form.get(f"target_group_{row_index}")
            try:
                target_group_id = int(tgid_raw) if tgid_raw else None
            except (ValueError, TypeError):
                target_group_id = None

            if not target_group_id:
                skipped_no_group.append({
                    'row': row_index,
                    'fio': row.get('matched_name') or '',
                    'reason': 'не обрано групу для паралельного студента',
                })
                continue
        else:
            target_group_id = row.get("target_group_id")
            if target_group_id is None and row.get("requested_group_id") is not None:
                target_group_id = row.get("requested_group_id")

        # ── Пошук існуючого документа для обраної групи ──
        if target_group_id is not None:
            cursor = db.cursor()
            cursor.execute("""
                SELECT id, group_id FROM education_documents
                WHERE student_id = ? AND group_id = ?
                ORDER BY id DESC LIMIT 1
            """, (row["student_id"], target_group_id))
            existing = cursor.fetchone()
        else:
            # Немає групи — legacy-запис із group_id IS NULL
            cursor = db.cursor()
            cursor.execute("""
                SELECT id, group_id FROM education_documents
                WHERE student_id = ? AND group_id IS NULL
                ORDER BY id DESC LIMIT 1
            """, (row["student_id"]))
            existing = cursor.fetchone()

        doc_type_en  = row.get("document_type_en", "") or ""
        inst_name_en = row.get("institution_name_en", "") or ""
        country_en   = row.get("country_en", "") or ""

        if existing:
            education_doc_id = existing["id"]
            cursor.execute("""
                UPDATE education_documents SET
                    document_type=?, document_type_en=?, document_number=?, completion_date=?,
                    institution_name=?, institution_name_en=?, country=?, country_en=?,
                    group_id=?
                WHERE id=?
            """, (row["document_type"], doc_type_en, row["document_number"], row["completion_date"],
                  row["institution_name"], inst_name_en, row["country"], row["country_en"],
                  target_group_id, education_doc_id))
            updated += 1
        else:
            cursor.execute("""
                INSERT INTO education_documents (
                    student_id, group_id, document_type, document_type_en, document_number,
                    completion_date, institution_name, institution_name_en, country, country_en
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (row["student_id"], target_group_id, row["document_type"], doc_type_en, row["document_number"],
                  row["completion_date"], row["institution_name"], inst_name_en, row["country"], country_en))
            education_doc_id = cursor.lastrowid
            added += 1

        # --- Foreign / recognition data ---
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

        has_foreign_data = any([
            reference_number, reference_institution, reference_country,
            reference_issue_date, recognition_certificate_number,
            recognition_issuer, recognition_date
        ])

        cursor.execute(
            "SELECT id FROM foreign_education_docs WHERE education_doc_id = ?",
            (education_doc_id,)
        )
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
                """, (
                    reference_number, reference_institution, reference_institution_en,
                    reference_country, reference_country_en, reference_issue_date,
                    recognition_certificate_number, recognition_issuer,
                    recognition_issuer_en, recognition_date, education_doc_id
                ))
            else:
                cursor.execute(
                    "DELETE FROM foreign_education_docs WHERE education_doc_id=?",
                    (education_doc_id,)
                )
        elif has_foreign_data:
            cursor.execute("""
                INSERT INTO foreign_education_docs (
                    education_doc_id,
                    reference_number, reference_institution, reference_institution_en,
                    reference_country, reference_country_en, reference_issue_date,
                    recognition_certificate_number, recognition_issuer,
                    recognition_issuer_en, recognition_date
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                education_doc_id, reference_number, reference_institution,
                reference_institution_en, reference_country, reference_country_en,
                reference_issue_date, recognition_certificate_number,
                recognition_issuer, recognition_issuer_en, recognition_date
            ))

    db.commit()

    log_action(
        current_username(),
        f"імпорт документів про освіту: додано {added}, оновлено {updated}"
        + (f", пропущено {len(skipped_no_group)} (не обрано групу)" if skipped_no_group else "")
    )

    msg = f"Додано записів: {added}, Оновлено записів: {updated}"
    if skipped_no_group:
        msg += f". Пропущено: {len(skipped_no_group)} (не обрано групу для паралельного студента)"
        for s in skipped_no_group:
            flash(f"Рядок {s['row']} ({s['fio']}): {s['reason']}", 'warning')
    flash(msg, "success")
    return redirect(url_for('admin.manage_education_documents'))