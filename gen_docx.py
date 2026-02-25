import os
import re
import sqlite3
from datetime import datetime
from docxtpl import DocxTemplate

from utils import log_action, logger as global_logger
from db import get_db

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

    # Проверка новых полей, включая недавно добавленные
    new_fields = [
        'qualification_name', 'degree_level', 'specialty', 'educational_program', 'knowledge_area',
        'qualification_name_en', 'degree_level_en', 'specialty_en', 'educational_program_en', 'knowledge_area_en',
        'program_credits', 'study_years', 'study_form', 'study_form_eu', 'start_year', 'end_year',
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

    # Проверка на диплом с отличием с отладочной информацией
    context['diploma_with_honor_text'] = student_dict.get('last_name_UA', '')  # Значение по умолчанию
    context['diploma_with_honor_text_en'] = student_dict.get('last_name_en', '')  # Значение по умолчанию
    if 'adddiplom' in template.lower():
        global_logger.debug(f"Проверка диплома с отличием для шаблона {template}")
        if context.get('subjects_grades') and context.get('practice_data') and \
           context.get('coursework_data') and context.get('attestation_data'):
            all_grades = (context['subjects_grades'] + context['practice_data'] + context['coursework_data'])
            total_grades = len(all_grades)
            # global_logger.debug(f"Всего оценок: {total_grades}, данные: {all_grades}")
            if total_grades > 0:
                excellent_count = sum(1 for grade in all_grades if 'Відмінно / Excellent' in grade.get('grade', ''))
                good_count = sum(1 for grade in all_grades if 'Добре / Good' in grade.get('grade', ''))
                other_count = total_grades - excellent_count - good_count
                attestation_grade = next((grade.get('grade', '') for grade in context['attestation_data'] if grade.get('grade')), '')
                # global_logger.debug(f"Отличных: {excellent_count}, Хороших: {good_count}, Других: {other_count}, Аттестация: {attestation_grade}")

                if excellent_count / total_grades >= 0.75 and other_count == 0 and 'Відмінно / Excellent' in attestation_grade:
                    context['diploma_with_honor_text'] = 'Диплом з відзнакою'
                    context['diploma_with_honor_text_en'] = 'Diploma with honours'
                    # global_logger.debug("Критерий выполнен: Диплом з відзнакою / Diploma with honours")
                else:
                    context['diploma_with_honor_text'] = 'Інформація відсутня'
                    context['diploma_with_honor_text_en'] = 'Information is absent'
                    # global_logger.debug("Критерий не выполнен: Информація відсутня / Information is absent")
            else:
                global_logger.debug("Нет оценок для анализа")
        else:
            global_logger.debug("Отсутствуют данные для проверки диплома с отличием")

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