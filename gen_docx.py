import os
import re
import sqlite3
from datetime import datetime
from docxtpl import DocxTemplate
from utils import log_action, logger as global_logger
from db import get_db
from datetime import datetime
from docxtpl import RichText


def _to_richtext_multiline(text, separator=";", font_size_pt=8, font_name='Times New Roman'):
    """
    Превращает 'Назва 1; Назва 2' в RichText с реальным переносом строки
    после каждого separator, с заданным размером и шрифтом.
    """
    rt = RichText()
    if not text:
        rt.add("", size=font_size_pt * 2, font=font_name)
        return rt

    parts = [p.strip() for p in text.split(separator) if p.strip()]
    for i, part in enumerate(parts):
        if i > 0:
            rt.add(separator + " ", size=font_size_pt * 2, font=font_name)
            rt.add("\n", size=font_size_pt * 2, font=font_name)
        rt.add(part, size=font_size_pt * 2, font=font_name)
    return rt

def _format_date_ddmmyyyy(date_str):
    """Приводит дату из БД (YYYY-MM-DD, DD.MM.YYYY или уже DD/MM/YYYY) к формату DD/MM/YYYY."""
    if not date_str:
        return ""
    date_str = date_str.strip()
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(date_str, fmt).strftime('%d/%m/%Y')
        except ValueError:
            continue
    return date_str  # неизвестный формат — вернуть как есть

def get_study_periods(cursor, student_id, lang='ua',
                       default_filiya="Львівська філія Приватного вищого навчального закладу «Європейський університет»",
                       default_filiya_en="Lviv Branch of Private Higher Education Establishment «European University»"):
    """
    Единый хелпер для подстановки в docx-шаблон.
    Возвращает кортеж (names_text, dates_text):
      names_text -> "Назва 1; Назва 2"
      dates_text -> "01/09/2022-31/05/2025; 01/09/2025-31/05/2026"

    lang='ua' -> використовує filiya
    lang='en' -> використовує filiya_en

    Якщо записів немає — повертає дефолтну філію і порожні дати.
    """
    cursor.execute("""
        SELECT filiya, filiya_en, group_name, start_date, end_date
        FROM student_study_periods
        WHERE student_id = ?
        ORDER BY period_order ASC, start_date ASC
    """, (student_id,))
    periods = cursor.fetchall()

    if not periods:
        default_name = default_filiya if lang == 'ua' else default_filiya_en
        return default_name, ""

    names = []
    dates = []
    for p in periods:
        name = p['filiya'] if lang == 'ua' else (p['filiya_en'] or p['filiya'])
        names.append(name)

        start = _format_date_ddmmyyyy(p['start_date'])
        end = _format_date_ddmmyyyy(p['end_date'])
        if start and end:
            dates.append(f"{start}-{end}")
        elif start:
            dates.append(f"{start}-т.ч." if lang == 'ua' else f"{start}-present")
        else:
            dates.append("")

    return "; ".join(names), "; ".join(dates)

def insert_subjects_table(doc, student_id):
    """Вставка таблицы с предметами и оценками студента."""
    # global_logger.debug(f"Запуск insert_subjects_table для student_id={student_id}")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    student = conn.execute("""
        SELECT s.*, 
               g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS group_name
        FROM students s
        LEFT JOIN groups g ON s.group_id = g.id
        WHERE s.id = ?
    """, (student_id,)).fetchone()

    if not student:
        global_logger.error(f"Студент с ID {student_id} не найден")
        conn.close()
        return False

    subjects = conn.execute("""
        SELECT s.code, s.name, s.credits, s.type, g.grade
        FROM subjects s
        LEFT JOIN grades g ON s.id = g.subject_id AND g.student_id = ?
        WHERE s.group_id = ?
        ORDER BY s.position
    """, (student_id, student['group_id'])).fetchall()

    if not subjects:
        # global_logger.warning(f"Предметы для студента ID {student_id}, group_id={student['group_id']} не найдены")
        conn.close()
        return False

    table = doc.add_table(rows=len(subjects) + 1, cols=5)
    table.style = 'Table Grid'

    headers = ['Код', 'Назва', 'Кредити', 'Тип', 'Оцінка']
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        cell.paragraphs[0].style.font.size = Pt(10)
        cell.paragraphs[0].style.font.name = 'Times New Roman'

    for i, subject in enumerate(subjects, 1):
        row = table.rows[i]
        row.cells[0].text = subject['code'] or ''
        row.cells[1].text = subject['name'] or ''
        row.cells[2].text = str(subject['credits']) or ''
        row.cells[3].text = subject['type'] or ''
        row.cells[4].text = str(subject['grade']) if subject['grade'] is not None else ''
        for cell in row.cells:
            cell.paragraphs[0].style.font.size = Pt(10)
            cell.paragraphs[0].style.font.name = 'Times New Roman'

    conn.close()
    global_logger.debug(f"Таблица с предметами для студента ID {student_id} успешно вставлена")
    return True

def clean_text(text):
    """Очищает текст от непечатаемых символов, сохраняя переносы строк и тире (—), и приводит к строке."""
    if text is None:
        return ''
    text = str(text).encode('utf-8').decode('utf-8', errors='ignore')
    # text = re.sub(r'[^\x20-\x7E\xA0-\xFF\u0400-\u04FF\u2014\u2013\n]', '', text)
    # global_logger.debug(f"clean_text input: '{text}', output: '{text.strip()}'")
    return text.strip()

def format_grade(grade, subject_type):
    """Преобразует числовую оценку в текстовую форму по формуле."""
    try:
        grade = int(grade)
        if not 0 <= grade <= 100:
            return "Ошибка: введите число от 0 до 100" if subject_type == "Залік" else ""
        
        if subject_type == "Залік":
            if 60 <= grade <= 100:
                letter = 'A' if 90 <= grade <= 100 else 'B' if 82 <= grade <= 89 else 'C' if 74 <= grade <= 81 else 'D' if 64 <= grade <= 73 else 'E' if 60 <= grade <= 63 else ''
                return f"Зараховано / Passed {grade} {letter}"
            return "Не зараховано / Fail"
        elif subject_type == "Екзамен":
            if 90 <= grade <= 100:
                letter = 'A'
                return f"Відмінно / Excellent {grade} {letter}"
            elif 74 <= grade <= 89:
                letter = 'B' if 82 <= grade <= 89 else 'C' if 74 <= grade <= 81 else ''
                return f"Добре / Good {grade} {letter}"
            elif 60 <= grade <= 73:
                letter = 'D' if 64 <= grade <= 73 else 'E' if 60 <= grade <= 63 else ''
                return f"Задовільно / Satisfactory {grade} {letter}"
            elif 35 <= grade <= 59:
                return f"Незадовільно / Fail {grade} Fx"
            elif 1 <= grade <= 34:
                return f"Незадовільно / Fail {grade} F"
            return "Незадовільно / Fail"
        return ""
    except (ValueError, TypeError):
        return "Ошибка: введите число от 0 до 100" if subject_type == "Залік" else ""

def get_subjects_grades(student_id, group_id):
    """Получение данных о предметах и их оценках."""
    # global_logger.debug(f"Запуск get_subjects_grades для student_id={student_id}, group_id={group_id}")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        results = conn.execute("""
            SELECT s.id, s.code, s.name, s.credits, s.type, s.position, IFNULL(g.grade, '') AS grade
            FROM subjects s
            LEFT JOIN grades g ON g.subject_id = s.id AND g.student_id = ?
            WHERE s.group_id = ?
            ORDER BY s.position
        """, (student_id, group_id)).fetchall()
        subjects = [dict(r) for r in results]
        # global_logger.debug(f"Получено {len(subjects)} предметов: {subjects}")
        valid_subjects = []
        for subject in subjects:
            if all(key in subject for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade']):
                subject = {k: clean_text(v) for k, v in subject.items()}
                subject['grade'] = format_grade(subject['grade'], subject['type']) if subject['grade'] else ''
                valid_subjects.append(subject)
            else:
                global_logger.warning(f"Неполные данные предмета пропущены: {subject}")
        return valid_subjects
    except Exception as e:
        global_logger.error(f"[get_subjects_grades] Ошибка: {e}")
        return []
    finally:
        conn.close()

def get_practice_data(student_id, group_id):
    """Получение данных о практиках и их оценках."""
    # global_logger.debug(f"Запуск get_practice_data для student_id={student_id}, group_id={group_id}")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        results = conn.execute("""
            SELECT p.id, p.code, p.name, p.credits, p.type, p.position, IFNULL(ag.grade, '') AS grade
            FROM practices p
            LEFT JOIN activity_grades ag ON ag.entity_id = p.id AND ag.entity_type = 'practice' AND ag.student_id = ?
            WHERE p.group_id = ?
            ORDER BY p.position
        """, (student_id, group_id)).fetchall()
        practices = [dict(r) for r in results]
        # global_logger.debug(f"Получено {len(practices)} практик: {practices}")
        valid_practices = []
        for practice in practices:
            if all(key in practice for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade']):
                practice = {k: clean_text(v) for k, v in practice.items()}
                practice['grade'] = format_grade(practice['grade'], practice['type']) if practice['grade'] else ''
                valid_practices.append(practice)
            else:
                global_logger.warning(f"Неполные данные практики пропущены: {practice}")
        return valid_practices
    except Exception as e:
        global_logger.error(f"[get_practice_data] Ошибка: {e}")
        return []
    finally:
        conn.close()

def get_coursework_data(student_id, group_id):
    """Получение данных о курсовых работах и их оценках."""
    # global_logger.debug(f"Запуск get_coursework_data для student_id={student_id}, group_id={group_id}")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        results = conn.execute("""
            SELECT c.id, c.code, c.name, c.credits, c.type, c.position, IFNULL(ag.grade, '') AS grade
            FROM courseworks c
            LEFT JOIN activity_grades ag ON ag.entity_id = c.id AND ag.entity_type = 'coursework' AND ag.student_id = ?
            WHERE c.group_id = ?
            ORDER BY c.position
        """, (student_id, group_id)).fetchall()
        courseworks = [dict(r) for r in results]
        # global_logger.debug(f"Получено {len(courseworks)} курсовых работ: {courseworks}")
        valid_courseworks = []
        for coursework in courseworks:
            if all(key in coursework for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade']):
                coursework = {k: clean_text(v) for k, v in coursework.items()}
                coursework['grade'] = format_grade(coursework['grade'], coursework['type']) if coursework['grade'] else ''
                valid_courseworks.append(coursework)
            else:
                global_logger.warning(f"Неполные данные курсовой работы пропущены: {coursework}")
        return valid_courseworks
    except Exception as e:
        global_logger.error(f"[get_coursework_data] Ошибка: {e}")
        return []
    finally:
        conn.close()

def get_attestation_data(student_id, group_id):
    """Получение данных об аттестациях и их оценках."""
    # global_logger.debug(f"Запуск get_attestation_data для student_id={student_id}, group_id={group_id}")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        results = conn.execute("""
            SELECT a.id, a.code, a.name, a.credits, a.type, a.position, IFNULL(ag.grade, '') AS grade, IFNULL(ag.name, '') AS student_name
            FROM attestations a
            LEFT JOIN activity_grades ag ON ag.entity_id = a.id AND ag.entity_type = 'attestation' AND ag.student_id = ?
            WHERE a.group_id = ?
            ORDER BY a.position
        """, (student_id, group_id)).fetchall()
        attestations = [dict(r) for r in results]
        # global_logger.debug(f"Получено {len(attestations)} аттестаций: {attestations}")
        valid_attestations = []
        for attestation in attestations:
            if all(key in attestation for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade', 'student_name']):
                attestation = {k: clean_text(v) for k, v in attestation.items()}
                attestation['grade'] = format_grade(attestation['grade'], attestation['type']) if attestation['grade'] else ''
                valid_attestations.append(attestation)
            else:
                global_logger.warning(f"Неполные данные аттестации пропущены: {attestation}")
        return valid_attestations
    except Exception as e:
        global_logger.error(f"[get_attestation_data] Ошибка: {e}")
        return []
    finally:
        conn.close()

def gen_doc(student: dict, military: dict, template='template.docx', out='out.docx', user_name='Система'):
    """Генерирует документ для студента на основе шаблона."""
    global_logger.debug(f"Запуск gen_doc: student_id={student.get('id', 'unknown')}, template={template}, out={out}")
    
    # Проверка входных данных
    # global_logger.debug(f"Входные данные student: {dict(student)}")
    # if military:
        # global_logger.debug(f"Входные данные military: {dict(military)}")
    # else:
        # global_logger.debug("Данные military отсутствуют")

    # Проверка существования шаблона
    if not os.path.exists(template):
        global_logger.error(f"Шаблон {template} не найден")
        raise FileNotFoundError(f"Шаблон {template} не найден")

    try:
        doc = DocxTemplate(template)
        global_logger.debug(f"Шаблон {template} успешно загружен")
    except Exception as e:
        global_logger.error(f"Ошибка при загрузке шаблона {template}: {str(e)}")
        raise

    # Получение данных о документах об образовании
    conn = get_db()
    conn.row_factory = sqlite3.Row
    education_docs = conn.execute("""
        SELECT ed.document_type, ed.document_number, ed.institution_name, ed.country, ed.completion_date,
               ed.document_type_en, ed.institution_name_en, ed.country_en,
               fed.reference_number, fed.reference_institution, fed.reference_country, fed.reference_issue_date,
               fed.reference_institution_en, fed.reference_country_en,
               fed.recognition_certificate_number, fed.recognition_issuer, fed.recognition_date,
               fed.recognition_issuer_en
        FROM education_documents ed
        LEFT JOIN foreign_education_docs fed ON ed.id = fed.education_doc_id
        WHERE ed.student_id = ?
        ORDER BY ed.id DESC LIMIT 1
    """, (student['id'],)).fetchone()
    conn.close()

    # Преобразование словарей
    student_dict = {k: clean_text(v) for k, v in dict(student).items()}
    military_dict = {k: clean_text(v) for k, v in dict(military).items()} if military else {}
    
    # Добавление данных об образовании в student_dict
    if education_docs:
        for key, value in dict(education_docs).items():
            student_dict[key] = clean_text(value) if value else ''

    # Форматирование birth_date
    birth_date = student_dict.get('birth_date', '')
    # global_logger.debug(f"Исходная birth_date: '{birth_date}', тип: {type(birth_date)}")
    if birth_date:
        try:
            date_obj = datetime.strptime(birth_date, '%d.%m.%Y')
            birth_date = date_obj.strftime('%d/%m/%Y')
            # global_logger.debug(f"Отформатированная birth_date: '{birth_date}'")
        except ValueError:
            try:
                date_obj = datetime.strptime(birth_date, '%Y-%m-%d')
                birth_date = date_obj.strftime('%d/%m/%Y')
                # global_logger.debug(f"Отформатированная birth_date (альтернативный формат): '{birth_date}'")
            except ValueError:
                # global_logger.warning(f"Неизвестный формат birth_date: '{birth_date}', оставляем как есть")
                birth_date = student_dict['birth_date']
    
    student_dict['birth_date'] = birth_date

    # Вычисление study_years на основе program_credits
    program_credits = student_dict.get('program_credits', '')
    study_years = ''
    try:
        credits = int(program_credits)
        if credits == 240:
            study_years = '4'
        elif credits == 180:
            study_years = '3'
        elif credits == 120:
            study_years = '2'
        elif credits == 90:  # Для магистратуры
            study_years = '1.5'
        else:
            study_years = str(credits // 60)  # Общее правило: 60 кредитов = 1 год
        # global_logger.debug(f"program_credits: {program_credits}, study_years: {study_years}")
    except (ValueError, TypeError):
        # global_logger.warning(f"Невалидное значение program_credits: '{program_credits}', study_years оставлено пустым")
        study_years = ''
    
    student_dict['study_years'] = study_years
    
    # Вычисление study_form_eu на основе study_form
    if 'adddiplom' in template.lower():
        study_form = student_dict.get('study_form', '')
        study_form_eu = ''
        if study_form == 'Денна':
            study_form_eu = 'Full'
        elif study_form == 'Заочна':
            study_form_eu = 'Part'
        else:
            study_form_eu = study_form
        # global_logger.debug(f"study_form: {study_form}, study_form_eu: {study_form_eu}")
    
        student_dict['study_form_eu'] = study_form_eu
    
    # Вычисление end_year на основе start_year, program_credits и degree_level
    end_year = ''
    start_year = student_dict.get('start_year', '')
    program_credits = student_dict.get('program_credits', '')
    degree_level = student_dict.get('degree_level', '')
    try:
        if program_credits and start_year:
            credits = int(program_credits)
            year = int(start_year)
            if degree_level == 'Бакалавр':
                if credits == 240:
                    end_year = str(year + 4)
                elif credits == 180:
                    end_year = str(year + 3)
            elif degree_level == 'Магістр':
                if credits == 90:
                    end_year = str(year + 2)  # Магистратура 1.5-2 года
                elif credits == 120:
                    end_year = str(year + 2)
            else:
                end_year = str(year + (credits // 60))  # Общее правило
        # global_logger.debug(f"start_year: {start_year}, program_credits: {program_credits}, degree_level: {degree_level}, end_year: {end_year}")
    except (ValueError, TypeError) as e:
        global_logger.warning(f"Ошибка при расчёте end_year: start_year='{start_year}', program_credits='{program_credits}', degree_level='{degree_level}', ошибка: {str(e)}")
        end_year = ''
    
    student_dict['end_year'] = end_year
    
    student_dict['end_year_short'] = student_dict['end_year'][-2:] if student_dict['end_year'] else ''

    # Проверка новых полей, включая недавно добавленные
    new_fields = [
        'qualification_name', 'degree_level', 'specialty', 'educational_program', 'knowledge_area',
        'qualification_name_en', 'degree_level_en', 'specialty_en', 'educational_program_en', 'knowledge_area_en',
        'program_credits', 'study_years', 'study_form', 'study_form_eu', 'start_year', 'end_year', 'end_year_short', 
        'institution_name_and_status', 'institution_name_and_status_en',
        'entry_requirements', 'entry_requirements_en',
        'learning_outcomes', 'learning_outcomes_en', 'program_includes', 'program_includes_en',
        'document_type', 'document_number', 'institution_name', 'country', 'completion_date',
        'document_type_en', 'institution_name_en', 'country_en',
        'reference_number', 'reference_institution', 'reference_country', 'reference_issue_date',
        'reference_institution_en', 'reference_country_en',
        'recognition_certificate_number', 'recognition_issuer', 'recognition_date',
        'recognition_issuer_en'
    ]
    # global_logger.debug(f"Новые поля в student_dict: {[(k, student_dict.get(k, '')) for k in new_fields]}")

    if degree_level == "Магістр":
        student_dict['top_qualification_text'] = "- підготовка кваліфікаційної роботи / preparation of qualification work"
        student_dict['bottom_qualification_text'] = "- захист кваліфікаційної роботи / defense of qualification work"

    # Обработка текста для полей с разделением на отдельные строки по \n и удалением лишнего \n
    fields_to_process = ['program_includes', 'program_includes_en', 'learning_outcomes', 'learning_outcomes_en']
    for field in fields_to_process:
        if field in student_dict and student_dict[field]:
            lines = student_dict[field].split('\n')
            cleaned_lines = [line.strip() for line in lines if line.strip()]
            student_dict[field] = cleaned_lines
            # global_logger.debug(f"Обработан {field} как список: {student_dict[field]}")

    # Объединяем словари
    context = {**student_dict, **military_dict}
    # global_logger.debug(f"Контекст перед рендерингом: {context}")

    # Данные для диплома
    try:
        if 'adddiplom' in template.lower() and 'group_id' in student_dict and 'id' in student_dict:
            context['subjects_grades'] = get_subjects_grades(student_dict['id'], student_dict['group_id']) or []
            context['practice_data'] = get_practice_data(student_dict['id'], student_dict['group_id']) or []
            context['coursework_data'] = get_coursework_data(student_dict['id'], student_dict['group_id']) or []
            context['attestation_data'] = get_attestation_data(student_dict['id'], student_dict['group_id']) or []
            # global_logger.debug(f"Данные для диплома: subjects_grades={context['subjects_grades']}, "
                        # f"practice_data={context['practice_data']}, "
                        # f"coursework_data={context['coursework_data']}, "
                        # f"attestation_data={context['attestation_data']}")
    except Exception as e:
        global_logger.error(f"Ошибка при получении данных для диплома для студента ID {student_dict.get('id', 'unknown')}: {e}")
        raise
    
    # -----------------------------
    # Автоподстановка аккредитации для диплома с отладкой
    # -----------------------------
    try:
        if 'adddiplom' in template.lower():
            conn = get_db()
            conn.row_factory = sqlite3.Row
            acc = conn.execute("""
                SELECT text_ua, text_en
                FROM accreditations
                WHERE degree = ? AND specialty = ?
                ORDER BY id DESC LIMIT 1
            """, (
                student_dict.get('degree_level', ''),
                student_dict.get('specialty', '')
            )).fetchone()
            conn.close()

            if acc:
                context['accreditation_text'] = acc['text_ua']
                context['accreditation_text_en'] = acc['text_en']
                # Вывод для отладки
                #print("=== Автоподстановка аккредитации ===")
                #print("UA:", acc['text_ua'])
                #print("EN:", acc['text_en'])
                #global_logger.debug(f"Автоподстановка аккредитации: UA='{acc['text_ua']}', EN='{acc['text_en']}'")
            else:
                context['accreditation_text'] = ''
                context['accreditation_text_en'] = ''
                #print("=== Автоподстановка аккредитации ===")
                #print("Аккредитация не найдена для ступені и спеціальності:")
                #print("degree_level:", student_dict.get('degree_level', ''))
                #print("specialty:", student_dict.get('specialty', ''))
                #global_logger.debug(f"Аккредитация не найдена: degree_level='{student_dict.get('degree_level', '')}', specialty='{student_dict.get('specialty', '')}'")

    except Exception as e:
        global_logger.error(f"Ошибка при получении аккредитации для диплома студента ID {student_dict.get('id', 'unknown')}: {e}")
        context['accreditation_text'] = ''
        context['accreditation_text_en'] = ''

    # Проверка на диплом с отличием с отладочной информацией
    context['diploma_with_honor_text'] = student_dict.get('last_name_UA', '')
    context['diploma_with_honor_text_en'] = student_dict.get('last_name_en', '')

    if 'adddiplom' in template.lower():

        all_grades = []

        if context.get('subjects_grades'):
            all_grades += context['subjects_grades']

        if context.get('practice_data'):
            all_grades += context['practice_data']

        total_grades = 0
        excellent_count = 0
        satisfactory_count = 0

        for g in all_grades:
            grade_text = g.get('grade', '')

            if not grade_text:
                continue

            total_grades += 1

            if ' A' in grade_text:
                excellent_count += 1

            elif ' D' in grade_text or ' E' in grade_text:
                satisfactory_count += 1

        # аттестация
        attestation_grade = ''

        if context.get('attestation_data'):
            attestation_grade = next(
                (g.get('grade', '') for g in context['attestation_data'] if g.get('grade')),
                ''
            )

        diploma_with_honours = (
            total_grades > 0
            and excellent_count / total_grades >= 0.75
            and satisfactory_count == 0
            and ' A' in attestation_grade
        )

        if diploma_with_honours:
            context['diploma_with_honor_text'] = 'Диплом з відзнакою'
            context['diploma_with_honor_text_en'] = 'Diploma with honours'
        else:
            context['diploma_with_honor_text'] = 'Інформація відсутня'
            context['diploma_with_honor_text_en'] = 'Information is absent'

    
    
    # -----------------------------
    # Подтягиваем диплом и додаток для Word
    # -----------------------------
    conn = get_db()
    conn.row_factory = sqlite3.Row
    diploma_row = conn.execute("""
        SELECT diploma_number, appendix_number
        FROM diplomas
        WHERE student_id = ?
        ORDER BY id DESC LIMIT 1
    """, (student['id'],)).fetchone()
    conn.close()
    
# -----------------------------
    # Період навчання (назви окремо, дати окремо, з переносом рядків)
    # -----------------------------
    conn = get_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        names_ua, dates_ua = get_study_periods(
            cursor, student_dict['id'], lang='ua',
            default_filiya=student_dict.get('institution_name', 'Львівська філія'),
            default_filiya_en=student_dict.get('institution_name_en', 'Lviv branch')
        )
        names_en, dates_en = get_study_periods(
            cursor, student_dict['id'], lang='en',
            default_filiya=student_dict.get('institution_name', 'Львівська філія'),
            default_filiya_en=student_dict.get('institution_name_en', 'Lviv branch')
        )
    except Exception as e:
        global_logger.error(f"Ошибка при получении периодов навчання для студента ID {student_dict.get('id', 'unknown')}: {e}")
        names_ua = student_dict.get('institution_name', '')
        dates_ua = ''
        names_en = student_dict.get('institution_name_en', '')
        dates_en = ''
    finally:
        conn.close()

    context['study_period_names'] = _to_richtext_multiline(names_ua, font_size_pt=8, font_name='Times New Roman')
    context['study_period_dates'] = _to_richtext_multiline(dates_ua, font_size_pt=8, font_name='Times New Roman')
    context['study_period_names_en'] = _to_richtext_multiline(names_en, font_size_pt=8, font_name='Times New Roman')
    context['study_period_dates_en'] = _to_richtext_multiline(dates_en, font_size_pt=8, font_name='Times New Roman')
    
    

    if diploma_row:
        # дополняем нулями до 6 цифр
        diploma_number = diploma_row['diploma_number'] or ''
        appendix_number = diploma_row['appendix_number'] or ''

        diploma_number = diploma_number.zfill(6) if diploma_number else ''
        appendix_number = appendix_number.zfill(6) if appendix_number else ''

        context['diploma_number'] = diploma_number
        context['appendix_number'] = appendix_number
    else:
        context['diploma_number'] = ''
        context['appendix_number'] = ''
        
    try:
        doc.render(context)
        global_logger.debug("Шаблон успешно отрендерен")
    except Exception as e:
        global_logger.error(f"Ошибка при рендеринге документа для студента ID {student_dict.get('id', 'unknown')}: {e}")
        raise

    student_name = f"{student_dict.get('last_name_UA', '')} {student_dict.get('first_name_UA', '')}".strip()
    log_action(user_name, f"згенерував документ '{out}' для студента {student_name}", student_dict.get('group_id'))
    
    try:
        doc.save(out)
        global_logger.debug(f"Документ сохранён как {out}")
    except Exception as e:
        global_logger.error(f"Ошибка при сохранении документа {out}: {e}")
        raise

    return out