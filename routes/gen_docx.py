
import os
import re
import sqlite3
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from routes.utils import log_action, logger as global_logger, is_student_on_reduced_program, program_track_condition, apply_track_overrides
from routes.db import get_db
from datetime import datetime
from docxtpl import RichText
from docx.shared import Mm

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

def _parse_date(date_str):
    """Повертає datetime або None, якщо формат не розпізнано."""
    if not date_str:
        return None
    date_str = date_str.strip()
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None

# З цієї дати (включно) і далі - формат "з ... до ..." / "from ... to ...".
# До неї (напр. 29/09/2026) - звичайний "01/09/2022-29/09/2026".
LONG_DATES_FORMAT_FROM = datetime(2026, 9, 30)

def get_study_periods(cursor, student_id, group_id=None, lang='ua',
                       default_filiya="Львівська філія Приватного вищого навчального закладу «Європейський університет»",
                       default_filiya_en="Lviv Branch of Private Higher Education Establishment «European University»"):
    """
    Повертає (names, dates) — періоди навчання САМЕ ДЛЯ ЦІЄЇ ГРУПИ
    (enrollment'а студента), бо в кожної групи періоди можуть бути свої
    (наприклад паралельне навчання в двох закладах з різними датами).

    Фільтр: `group_id = ? OR group_id IS NULL` — щоб старі періоди,
    які з якоїсь причини не прив'язані до конкретної групи, все одно
    підтягувалися (легасі-поведінка).
    """
    if group_id:
        cursor.execute("""
            SELECT filiya, filiya_en, group_name, start_date, end_date
            FROM student_study_periods
            WHERE student_id = ? AND (group_id = ? OR group_id IS NULL)
            ORDER BY period_order ASC, start_date ASC
        """, (student_id, group_id))
    else:
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

    end_dates = [d for d in (_parse_date(p['end_date']) for p in periods) if d]
    use_long_format = bool(end_dates) and max(end_dates) >= LONG_DATES_FORMAT_FROM

    names = []
    dates = []
    for p in periods:
        name = p['filiya'] if lang == 'ua' else (p['filiya_en'] or p['filiya'])
        names.append(name)

        start = _format_date_ddmmyyyy(p['start_date'])
        end = _format_date_ddmmyyyy(p['end_date'])
        if start and end:
            if use_long_format:
                dates.append(f"з {start} до {end}" if lang == 'ua' else f"from {start} to {end}")
            else:
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
    """Получение данных о предметах и их оценках.

    Якщо студент навчається за скороченою програмою відносно своєї
    групи (program_credits_override менше за program_credits групи -
    напр. вступив одразу на 2 курс з визнанням частини кредитів
    диплома бакалавра):
      - предмети з subjects.full_program_only=1 виключаються зі
        списку - вони стосуються лише повної програми;
      - для спільних предметів, де subjects.reduced_credits заповнено,
        підміняється значення credits - той самий предмет може мати
        меншу кількість кредитів у скороченій програмі (напр. курсова
        робота 20 кредитів у повній, 10 - у скороченій).
    """
    # global_logger.debug(f"Запуск get_subjects_grades для student_id={student_id}, group_id={group_id}")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        is_reduced_program = is_student_on_reduced_program(conn, student_id, group_id)

        query = """
            SELECT s.id, s.code, s.name, s.credits, s.reduced_credits, s.type, s.reduced_type, s.position, IFNULL(g.grade, '') AS grade
            FROM subjects s
            LEFT JOIN grades g ON g.subject_id = s.id AND g.student_id = ?
            WHERE s.group_id = ?
        """
        params = [student_id, group_id]
        query += program_track_condition(is_reduced_program, 's')
        query += " ORDER BY s.position"

        results = conn.execute(query, params).fetchall()
        subjects = [dict(r) for r in results]
        # global_logger.debug(f"Получено {len(subjects)} предметов: {subjects}")
        valid_subjects = []
        for subject in subjects:
            if all(key in subject for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade']):
                apply_track_overrides(subject, is_reduced_program)
                subject.pop('reduced_credits', None)
                subject.pop('reduced_type', None)
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
    """Получение данных о практиках и их оценках.

    Той самий принцип скороченої програми, що й у get_subjects_grades:
    виключення full_program_only=1 і підміна кредитів через
    reduced_credits, якщо заповнено.
    """
    # global_logger.debug(f"Запуск get_practice_data для student_id={student_id}, group_id={group_id}")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        is_reduced_program = is_student_on_reduced_program(conn, student_id, group_id)

        query = """
            SELECT p.id, p.code, p.name, p.credits, p.reduced_credits, p.type, p.reduced_type, p.position, IFNULL(ag.grade, '') AS grade
            FROM practices p
            LEFT JOIN activity_grades ag ON ag.entity_id = p.id AND ag.entity_type = 'practice' AND ag.student_id = ?
            WHERE p.group_id = ?
        """
        params = [student_id, group_id]
        query += program_track_condition(is_reduced_program, 'p')
        query += " ORDER BY p.position"

        results = conn.execute(query, params).fetchall()
        practices = [dict(r) for r in results]
        # global_logger.debug(f"Получено {len(practices)} практик: {practices}")
        valid_practices = []
        for practice in practices:
            if all(key in practice for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade']):
                apply_track_overrides(practice, is_reduced_program)
                practice.pop('reduced_credits', None)
                practice.pop('reduced_type', None)
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
    """Получение данных о курсовых работах и их оценках.

    Той самий принцип скороченої програми, що й у get_subjects_grades.
    """
    # global_logger.debug(f"Запуск get_coursework_data для student_id={student_id}, group_id={group_id}")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        is_reduced_program = is_student_on_reduced_program(conn, student_id, group_id)

        query = """
            SELECT c.id, c.code, c.name, c.credits, c.reduced_credits, c.type, c.reduced_type, c.position, IFNULL(ag.grade, '') AS grade
            FROM courseworks c
            LEFT JOIN activity_grades ag ON ag.entity_id = c.id AND ag.entity_type = 'coursework' AND ag.student_id = ?
            WHERE c.group_id = ?
        """
        params = [student_id, group_id]
        query += program_track_condition(is_reduced_program, 'c')
        query += " ORDER BY c.position"

        results = conn.execute(query, params).fetchall()
        courseworks = [dict(r) for r in results]
        # global_logger.debug(f"Получено {len(courseworks)} курсовых работ: {courseworks}")
        valid_courseworks = []
        for coursework in courseworks:
            if all(key in coursework for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade']):
                apply_track_overrides(coursework, is_reduced_program)
                coursework.pop('reduced_credits', None)
                coursework.pop('reduced_type', None)
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
    """Получение данных об аттестациях и их оценках.

    Той самий принцип скороченої програми, що й у get_subjects_grades.
    """
    # global_logger.debug(f"Запуск get_attestation_data для student_id={student_id}, group_id={group_id}")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        is_reduced_program = is_student_on_reduced_program(conn, student_id, group_id)

        query = """
            SELECT a.id, a.code, a.name, a.credits, a.reduced_credits, a.type, a.reduced_type, a.position, IFNULL(ag.grade, '') AS grade, IFNULL(ag.name, '') AS student_name
            FROM attestations a
            LEFT JOIN activity_grades ag ON ag.entity_id = a.id AND ag.entity_type = 'attestation' AND ag.student_id = ?
            WHERE a.group_id = ?
        """
        params = [student_id, group_id]
        query += program_track_condition(is_reduced_program, 'a')
        query += " ORDER BY a.position"

        results = conn.execute(query, params).fetchall()
        attestations = [dict(r) for r in results]
        # global_logger.debug(f"Получено {len(attestations)} аттестаций: {attestations}")
        valid_attestations = []
        for attestation in attestations:
            if all(key in attestation for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade', 'student_name']):
                apply_track_overrides(attestation, is_reduced_program)
                attestation.pop('reduced_credits', None)
                attestation.pop('reduced_type', None)
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

def apply_reduced_program_texts(student_dict):
    """Якщо студент навчається за скороченою програмою відносно своєї
    групи (program_credits_override менше за program_credits групи -
    напр. вступив одразу на 2 курс з визнанням частини кредитів
    попереднього диплома):
      - підміняє "Вимоги для вступу", "Програма підготовки включає" і
        "Програмні результати навчання" на відповідні _reduced-варіанти
        групи, якщо вони заповнені (порожній _reduced-текст лишає
        звичайний текст групи - щоб не ламати документи для груп, де
        скорочений варіант ще не заповнили);
      - підміняє саме значення student_dict['program_credits'] на
        program_credits_override, щоб розрахунок "study_years"/
        "end_year" нижче за текстом (а також пункт додатку "Тривалість
        освітньої програми в кредитах та/або роках") показував РЕАЛЬНУ
        кількість кредитів студента, а не кредити групи.
    Викликається до розбиття program_includes/learning_outcomes на
    рядки і до обчислення study_years/end_year, щоб обидва спирались
    вже на підмінене значення.
    """
    override = student_dict.get('program_credits_override')
    base = student_dict.get('program_credits')
    try:
        is_reduced = bool(override) and bool(base) and int(override) < int(base)
    except (TypeError, ValueError):
        is_reduced = False

    if not is_reduced:
        return

    for field in ('entry_requirements', 'entry_requirements_en',
                  'program_includes', 'program_includes_en',
                  'learning_outcomes', 'learning_outcomes_en'):
        # Колонки в БД називаються "..._reduced" (укр.) і
        # "..._reduced_en" (англ.) - тобто "_reduced" завжди ПЕРЕД
        # "_en", а не в кінці назви поля.
        if field.endswith('_en'):
            reduced_key = f"{field[:-3]}_reduced_en"
        else:
            reduced_key = f"{field}_reduced"
        reduced_value = student_dict.get(reduced_key)
        if reduced_value:
            student_dict[field] = reduced_value

    student_dict['program_credits'] = override


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

    # Получение данных о документах об образовании — для КОНКРЕТНОЇ групи,
    # для якої генерується документ. Fallback на group_id IS NULL — щоб
    # старі «загальні» документи підтягувалися, коли для групи свого
    # немає. Пріоритет: спочатку документ конкретної групи, потім NULL.
    _group_id = student['group_id'] if 'group_id' in student.keys() else None
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
          AND (ed.group_id = ? OR ed.group_id IS NULL)
        ORDER BY (ed.group_id = ?) DESC, ed.id DESC
        LIMIT 1
    """, (student['id'], _group_id, _group_id)).fetchone()
    conn.close()

    # Преобразование словарей
    student_dict = {k: clean_text(v) for k, v in dict(student).items()}
    military_dict = {k: clean_text(v) for k, v in dict(military).items()} if military else {}

    # Періоди навчання — для ТІЄЇ групи, для якої генерується документ.
    # Робимо це ТУТ, ПІСЛЯ student_dict, щоб student_dict['group_id']
    # уже був доступний (при паралельному навчанні в кожної групи свої
    # періоди). НЕ дублювати цей блок у кінці функції.
    conn = get_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        current_group_id = student_dict.get('group_id')
        names_ua, dates_ua = get_study_periods(
            cursor, student_dict['id'], group_id=current_group_id, lang='ua'
        )
        names_en, dates_en = get_study_periods(
            cursor, student_dict['id'], group_id=current_group_id, lang='en'
        )
    except Exception as e:
        global_logger.error(
            f"Ошибка при получении периодов навчання для студента ID "
            f"{student_dict.get('id', 'unknown')}: {e}"
        )
        names_ua = "Львівська філія Приватного вищого навчального закладу «Європейський університет»"
        dates_ua = ''
        names_en = "Lviv Branch of Private Higher Education Establishment «European University»"
        dates_en = ''
    finally:
        conn.close()

    # Якщо студент на скороченій програмі...
    apply_reduced_program_texts(student_dict)

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
    # (раніше обчислювалось лише для шаблонів з "adddiplom" у назві -
    # тепер доступне у будь-якому шаблоні)
    study_form = student_dict.get('study_form', '')
    study_form_eu = ''
    if study_form == 'Денна':
        study_form_eu = 'Full'
    elif study_form == 'Заочна':
        study_form_eu = 'Part'
    elif study_form == 'Дистанційна':
        study_form_eu = 'Distance'
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

    # Об'єднаем словари
    context = {**student_dict, **military_dict}
    # global_logger.debug(f"Контекст перед рендерингом: {context}")

    # Фото студента (3х4, якщо завантажене на сторінці студента) - як
    # InlineImage для шаблонів, у яких є картинка-плейсхолдер
    # {{ student_photo }}. Якщо фото немає - ключ просто не додається в
    # контекст, і плейсхолдер у шаблоні лишиться порожнім (не ламає
    # рендеринг решти документа).
    photo_rel = student_dict.get('photo')
    if photo_rel:
        photo_full_path = os.path.join(os.getcwd(), 'static', photo_rel)
        if os.path.exists(photo_full_path):
            try:
                context['student_photo'] = InlineImage(doc, photo_full_path, width=Mm(30))
            except Exception as e:
                global_logger.error(f"Не вдалося вставити фото студента (ID {student_dict.get('id')}): {e}")

    # Дані для диплома (список оцінок з предметів, практик, курсових,
    # атестацій) - раніше збиралися лише для шаблонів з "adddiplom" у
    # назві, тепер доступні у будь-якому шаблоні.
    try:
        if 'group_id' in student_dict and 'id' in student_dict:
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
    # (раніше обчислювалось лише для шаблонів з "adddiplom" у назві,
    # тепер доступне у будь-якому шаблоні)
    # -----------------------------
    try:
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

    # Проверка на диплом с отличием
    # (раніше обчислювалось лише для "adddiplom"; за замовчуванням тут
    # раніше стояло прізвище студента - це виглядало як забутий
    # налагоджувальний код, а не навмисна поведінка, тому прибрано)
    context['diploma_with_honor_text'] = ''
    context['diploma_with_honor_text_en'] = ''

    all_grades = []

    if context.get('subjects_grades'):
        all_grades += context['subjects_grades']

    if context.get('practice_data'):
        all_grades += context['practice_data']

    total_grades = 0
    excellent_count = 0
    satisfactory_count = 0
    fail_count = 0

    for g in all_grades:
        grade_text = g.get('grade', '')

        if not grade_text:
            continue

        total_grades += 1

        if ' A' in grade_text:
            excellent_count += 1

        elif ' D' in grade_text or ' E' in grade_text:
            satisfactory_count += 1

        # Fx / F (або текст без літери, напр. "Незадовільно / Fail") -
        # це провал, а не «просто не A/D/E». Без цієї гілки студент із
        # 45 Fx і рештою A отримував «Диплом з відзнакою».
        elif ' Fx' in grade_text or ' F' in grade_text or 'Незадовільно' in grade_text or 'Fail' in grade_text:
            fail_count += 1

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
        and fail_count == 0
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
    # Номери диплома/додатка беруться для КОНКРЕТНОЇ групи, під яку
    # генерується документ (student_dict['group_id']), а не «будь-який»
    # номер студента. Якщо для цієї групи номер ще не заводили — у
    # документі номер буде порожнім, і адмін побачить, що треба заповнити
    # (Довідники → Номери диплому / додатку → ця група).
    diploma_row = conn.execute("""
        SELECT diploma_number, appendix_number
        FROM diplomas
        WHERE student_id = ? AND group_id = ?
        ORDER BY id DESC LIMIT 1
    """, (student['id'], student_dict.get('group_id'))).fetchone()
    conn.close()
    
    # Якщо періодів немає (дати порожні) — формуємо дати та назву на основі даних студента
    if not dates_ua:
        start_year = student_dict.get('start_year', '')
        end_year = student_dict.get('end_year', '')
        
        if start_year and end_year:
            if end_year >= "2025":
                dates_ua = f"з 01/09/{start_year} до 30/09/{end_year}"
                dates_en = f"from 01/09/{start_year} to 30/09/{end_year}"
            else:
                dates_ua = f"01/09/{start_year}–30/09/{end_year}"
                dates_en = f"01/09/{start_year}–30/09/{end_year}"

        # Раніше тут було жорстке порівняння з двома захардкодженими
        # рядками (Львівська філія / основний заклад) - ламалось би для
        # будь-якої третьої ліцензії (напр. Польської). Тепер
        # institution_name_and_status/_en беруться напряму з ліцензії
        # студента (routes/students.py, routes/admin.py) - тому просто
        # використовуємо їх як є, без порівняння.
        
        def _institution_short_name(full_text):
            """'Назва «Х». Приватна форма... Ліцензія...' -> 'Назва «Х»' (до першої »).
            Якщо лапок немає - повертає '', і тоді лишається назва за замовчуванням."""
            if not full_text:
                return ''
            end = full_text.find('»')
            return full_text[:end + 1].strip() if end != -1 else ''    
        
        
        names_ua = _institution_short_name(student_dict.get('institution_name_and_status', '')) or names_ua
        names_en = _institution_short_name(student_dict.get('institution_name_and_status_en', '')) or names_en

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



# ── Відомість сесії (той самий стек що gen_doc: DocxTemplate) ──────

def _vidomist_template_path():
    candidates = [
        os.path.join('static', 'templates', 'word', 'vidomist_sheet.docx'),
        os.path.join(os.path.dirname(__file__), '..', 'static', 'templates', 'word', 'vidomist_sheet.docx'),
    ]
    for p in candidates:
        ap = os.path.abspath(p)
        if os.path.isfile(ap):
            return ap
    return os.path.abspath(candidates[0])


def _vidomist_format_specialty(code, name):
    code = (code or '').strip()
    name = (name or '').strip()
    if not code and not name:
        return ''
    if not code:
        return name
    if not name:
        return code
    if name == code or name.startswith(code + ' '):
        return name
    parts = name.split(None, 1)
    if parts and parts[0].upper() == code.upper():
        return name
    return f'{code} {name}'.strip()


def _vidomist_roman(course):
    m = {1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI'}
    try:
        return m.get(int(course), str(course or ''))
    except (TypeError, ValueError):
        return str(course or '')


def _vidomist_semester_word(sem):
    m = {1: 'перший', 2: 'другий', 3: 'третій'}
    try:
        return m.get(int(sem), str(sem or ''))
    except (TypeError, ValueError):
        return str(sem or '')


def _vidomist_ects(points):
    if points is None or points == '':
        return ''
    try:
        p = int(points)
    except (TypeError, ValueError):
        return ''
    if p >= 90: return 'A'
    if p >= 82: return 'B'
    if p >= 74: return 'C'
    if p >= 64: return 'D'
    if p >= 60: return 'E'
    if p >= 35: return 'FX'
    return 'F'


def _vidomist_national(points, control_form):
    if points is None or points == '':
        return ''
    try:
        p = int(points)
    except (TypeError, ValueError):
        return ''
    if (control_form or '') == 'Залік':
        return 'Зараховано' if p >= 60 else 'Незараховано'
    if p >= 90: return 'Відмінно / Excellent'
    if p >= 82: return 'Добре/ Good'
    if p >= 74: return 'Добре/ Good'
    if p >= 64: return 'Задовільно/ Satisfactory'
    if p >= 60: return 'Задовільно/ Satisfactory'
    return 'Незадовільно/ Fail'


def _vidomist_build_context(sheet, students):
    """Контекст для DocxTemplate — ті самі ключі що {{ }} у шаблоні."""
    sheet = dict(sheet) if not isinstance(sheet, dict) else sheet
    control = sheet.get('control_form') or 'Екзамен'
    teacher = sheet.get('teacher_name') or ''

    rows = []
    for i, st in enumerate(students or []):
        if not isinstance(st, dict):
            try:
                st = dict(st)
            except Exception:
                st = {}
        pts = st.get('points')
        nat = st.get('national_grade') or _vidomist_national(pts, control)
        ects = st.get('ects') or _vidomist_ects(pts)
        def _v(x):
            return '' if x is None else x
        rows.append({
            'n': i + 1,
            'full_name': st.get('full_name') or '',
            'edebo_code': st.get('edebo_code') or '',
            'work_topic': st.get('work_topic') or '',
            'points_mid': _v(st.get('points_mid')),
            'points_current': _v(st.get('points_current')),
            'points_exam': _v(st.get('points_exam')),
            'points': _v(pts),
            'national_grade': nat or '',
            'ects': ects or '',
            'teacher_sign': st.get('teacher_sign') or teacher or '',
        })

    buckets = {k: 0 for k in ('A', 'B', 'C', 'D', 'E', 'FX', 'F', 'absent')}
    for s in rows:
        if s['points'] == '' or s['points'] is None:
            buckets['absent'] += 1
        else:
            key = (s['ects'] or 'F').upper()
            buckets[key if key in buckets else 'F'] += 1
    total = max(len(rows), 1)
    quality = buckets['A'] + buckets['B'] + buckets['C']
    absolute = quality + buckets['D'] + buckets['E']

    form = sheet.get('study_form') or ''
    num = sheet.get('statement_number') or ''
    inst = (sheet.get('institution_name_and_status') or
            'ПВНЗ „Європейський університет“\nЛьвівська філія')
    if '  ' in inst and '\n' not in inst and 'філія' in inst.lower():
        # розбити заклад / філія
        import re
        m = re.match(r'^(.*університет[^\"]*[\"“]?)\s+(Львівська\s+філія)\s*$', inst, re.I)
        if m:
            inst = m.group(1).strip() + '\n' + m.group(2).strip()

    ctx = {
        'institution': inst,
        'faculty': sheet.get('faculty') or sheet.get('department') or 'Інформаційні системи та технології',
        'specialty': _vidomist_format_specialty(sheet.get('specialty_code'), sheet.get('specialty')),
        'course': _vidomist_roman(sheet.get('course')),
        'group_name': sheet.get('group_name') or '',
        'study_form': form,
        'study_form_part': f'({form})' if form else '',
        'academic_year': sheet.get('academic_year') or '',
        'degree_level': sheet.get('degree_level') or '',
        'statement_title': f'ВІДОМІСТЬ  № {num}.'.strip() if num else 'ВІДОМІСТЬ',
        'statement_number': num,
        'exam_date': sheet.get('exam_date') or '',
        'discipline_name': sheet.get('discipline_name') or '',
        'semester_word': _vidomist_semester_word(sheet.get('semester')),
        'semester': str(sheet.get('semester') or ''),
        'control_form': (control or '').lower(),
        'hours_credits': sheet.get('hours_credits') or '',
        'teacher_name': teacher,
        'license_name': sheet.get('license_name') or None,
        'students': rows,
        'stats_a_label': 'А (90-100)', 'stats_a_count': str(buckets['A']),
        'stats_a_pct': f"{buckets['A']/total:.1%}",
        'stats_b_label': 'В(82-89)', 'stats_b_count': str(buckets['B']),
        'stats_b_pct': f"{buckets['B']/total:.1%}",
        'stats_c_label': 'С(74-81)', 'stats_c_count': str(buckets['C']),
        'stats_c_pct': f"{buckets['C']/total:.1%}",
        'stats_d_label': 'D(64-73)', 'stats_d_count': str(buckets['D']),
        'stats_d_pct': f"{buckets['D']/total:.1%}",
        'stats_e_label': 'E(60-63)', 'stats_e_count': str(buckets['E']),
        'stats_e_pct': f"{buckets['E']/total:.1%}",
        'stats_fx_label': 'FX(35-59)', 'stats_fx_count': str(buckets['FX']),
        'stats_fx_pct': f"{buckets['FX']/total:.1%}",
        'stats_f_label': 'F(1-34)', 'stats_f_count': str(buckets['F']),
        'stats_f_pct': f"{buckets['F']/total:.1%}",
        'stats_abs_label': "Не з'явився", 'stats_abs_count': str(buckets['absent']),
        'stats_abs_pct': f"{buckets['absent']/total:.1%}",
        'stats_quality': f"{quality/total:.1%}",
        'stats_absolute': f"{absolute/total:.1%}",
    }
    return ctx, rows


def _vidomist_fill_students_table(doc, students_rows):
    """
    Як insert_subjects_table: після render дописати рядки студентів.
    Шукає таблицю з «Прізвище» / «№» у шапці; клонує останній (зразок) рядок.
    """
    from copy import deepcopy
    from docx.document import Document as DocType

    table = None
    for t in doc.tables:
        if not t.rows:
            continue
        head = ' '.join((c.text or '') for c in t.rows[0].cells).lower()
        if 'прізвище' in head or '№' in head or 'з/п' in head or 'з\\п' in head:
            table = t
            break
    if table is None:
        return

    header_rows = 1
    if len(table.rows) >= 2:
        head2 = ' '.join((c.text or '') for c in table.rows[1].cells).lower()
        if any(x in head2 for x in ('міжсесій', 'поточн', 'іспит', 'ects', 'підсумков', 'нац')):
            header_rows = 2

    sample_idx = header_rows if len(table.rows) > header_rows else max(0, header_rows - 1)
    sample_tr = deepcopy(table.rows[sample_idx]._tr)
    n_grid = len(table.rows[sample_idx].cells)

    # map columns by header text
    texts = [''] * n_grid
    for ri in range(min(header_rows, len(table.rows))):
        for i, c in enumerate(table.rows[ri].cells):
            if i >= n_grid:
                break
            tx = (c.text or '').strip().lower().replace('\n', ' ')
            if tx:
                texts[i] = (texts[i] + ' ' + tx).strip() if texts[i] else tx

    def find(*needles):
        for i, tx in enumerate(texts):
            for n in needles:
                if n in tx:
                    return i
        return None

    roles = {
        'n': find('№', 'з\\п', 'з/п') if find('№', 'з\\п', 'з/п') is not None else 0,
        'name': find('прізвище', 'піб', 'студент'),
        'edebo': find('індив', 'план'),
        'mid': find('міжсесій'),
        'current': find('поточн'),
        'exam': find('іспит', 'залік'),
        'total': find('підсумков', '100 баль'),
        'national': find('нац', 'національ'),
        'ects': find('ects'),
        'sign': find('підпис'),
    }

    # remove data rows after header
    tbl = table._tbl
    while len(table.rows) > header_rows:
        tbl.remove(table.rows[-1]._tr)

    def write_cell(row, idx, value):
        if idx is None or idx < 0 or idx >= len(row.cells):
            return
        cell = row.cells[idx]
        text = '' if value is None else str(value)
        if not cell.paragraphs:
            return
        p = cell.paragraphs[0]
        if p.runs:
            p.runs[0].text = text
            for r in p.runs[1:]:
                r.text = ''
        else:
            p.add_run(text)

    for s in students_rows:
        tbl.append(deepcopy(sample_tr))
        row = table.rows[-1]
        write_cell(row, roles.get('n'), s.get('n'))
        write_cell(row, roles.get('name'), s.get('full_name'))
        write_cell(row, roles.get('edebo'), s.get('edebo_code'))
        write_cell(row, roles.get('mid'), s.get('points_mid'))
        write_cell(row, roles.get('current'), s.get('points_current'))
        write_cell(row, roles.get('exam'), s.get('points_exam'))
        write_cell(row, roles.get('total'), s.get('points'))
        write_cell(row, roles.get('national'), s.get('national_grade'))
        write_cell(row, roles.get('ects'), s.get('ects'))
        write_cell(row, roles.get('sign'), s.get('teacher_sign'))


def gen_vidomist_sheet(sheet, students, template=None, out=None, user_name='Система'):
    """
    Відомість сесії — той самий механізм що gen_doc:
      DocxTemplate(template).render(context) → save

    Таблиця студентів: без {%tr %} (Word ламає endfor) —
    після render рядки додаються як у insert_subjects_table.
    """
    if template is None:
        template = _vidomist_template_path()
    if not os.path.exists(template):
        global_logger.error(f"Шаблон відомості не знайдено: {template}")
        raise FileNotFoundError(f"Шаблон відомості не знайдено: {template}")

    context, rows = _vidomist_build_context(sheet, students)

    try:
        doc = DocxTemplate(template)
        global_logger.debug(f"Відомість: шаблон {template} завантажено")
    except Exception as e:
        global_logger.error(f"Помилка завантаження шаблону відомості: {e}")
        raise

    try:
        doc.render(context)
        global_logger.debug("Відомість: render OK")
    except Exception as e:
        global_logger.error(f"Помилка render відомості: {e}")
        raise

    # Дописати студентів у таблицю (стиль рядка-зразка з шаблону)
    try:
        _vidomist_fill_students_table(doc, rows)
    except Exception as e:
        global_logger.warning(f"Відомість: заповнення таблиці студентів: {e}")

    disc = (sheet.get('discipline_name') if isinstance(sheet, dict) else None) or ''
    log_action(user_name, f"згенерував відомість '{disc}'", None)

    if out:
        try:
            doc.save(out)
            global_logger.debug(f"Відомість збережено: {out}")
        except Exception as e:
            global_logger.error(f"Помилка збереження відомості {out}: {e}")
            raise
        return out

    from io import BytesIO
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf



# ── Загальна відомість сесії (шаблон + таблиця кодом) ──────────────

def _vidomist_summary_template_path():
    candidates = [
        os.path.join('static', 'templates', 'word', 'vidomist_summary.docx'),
        os.path.join(os.path.dirname(__file__), '..', 'static', 'templates', 'word', 'vidomist_summary.docx'),
    ]
    for p in candidates:
        ap = os.path.abspath(p)
        if os.path.isfile(ap):
            return ap
    return os.path.abspath(candidates[0])


def _vidomist_insert_summary_table(document, sheets, order, names, grades):
    """Вставка таблиці на місце маркера. Без merge комірок (стабільний XML)."""
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    MARKERS = ('<<<GRADES_TABLE>>>', '{{ grades_table }}')
    marker = None
    for p in document.paragraphs:
        text = p.text or ''
        if any(m in text for m in MARKERS):
            marker = p
            break
    if marker is None:
        for p in document.paragraphs:
            if 'декан' in (p.text or '').lower():
                marker = p
                break

    sheets = sheets or []
    order = order or []
    n_disc = len(sheets)
    cols = 2 + n_disc * 3 + 1
    if cols < 3:
        cols = 3

    def set_text(cell, text, bold=False, size=8, left=False):
        cell.text = ''
        p = cell.paragraphs[0]
        run = p.add_run('' if text is None else str(text))
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.bold = bold
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if left else WD_ALIGN_PARAGRAPH.CENTER

    table = document.add_table(rows=2 + len(order), cols=cols)
    try:
        table.style = 'Table Grid'
    except Exception:
        pass

    set_text(table.rows[0].cells[0], '№', bold=True)
    set_text(table.rows[0].cells[1], "Прізвище, ім'я та по батькові", bold=True)
    for i, sh in enumerate(sheets):
        base = 2 + i * 3
        title = (sh.get('discipline_name') or '')[:40]
        ctrl = sh.get('control_form') or ''
        # без merge: назва в першій з трьох колонок, підписи в рядку 1
        set_text(table.rows[0].cells[base], f'{title} ({ctrl})', bold=True, size=7)
        set_text(table.rows[0].cells[base + 1], '', bold=True, size=7)
        set_text(table.rows[0].cells[base + 2], '', bold=True, size=7)
        for j, h in enumerate(['нац.', 'бали', 'ECTS']):
            set_text(table.rows[1].cells[base + j], h, bold=True, size=7)
    avg_col = 2 + n_disc * 3
    set_text(table.rows[0].cells[avg_col], 'Середній бал', bold=True, size=8)
    set_text(table.rows[1].cells[0], '', size=7)
    set_text(table.rows[1].cells[1], '', size=7)
    set_text(table.rows[1].cells[avg_col], 'бали', bold=True, size=7)

    for ri, sid in enumerate(order):
        row = table.rows[ri + 2]
        set_text(row.cells[0], ri + 1, size=8)
        set_text(row.cells[1], names.get(sid, ''), size=8, left=True)
        pts_list = []
        for i, sh in enumerate(sheets):
            base = 2 + i * 3
            g = (grades.get(sid) or {}).get(sh['id']) or {}
            set_text(row.cells[base], g.get('national_grade') or '', size=7)
            pts = g.get('points')
            set_text(row.cells[base + 1], '' if pts is None else pts, size=8)
            set_text(row.cells[base + 2], g.get('ects') or '', size=8)
            if pts is not None and pts != '':
                try:
                    pts_list.append(float(pts))
                except (TypeError, ValueError):
                    pass
        avg = round(sum(pts_list) / len(pts_list), 2) if pts_list else ''
        set_text(row.cells[avg_col], avg, bold=True, size=8)

    if marker is not None:
        # перенести таблицю одразу ПІСЛЯ маркера, маркер прибрати
        try:
            marker._p.addnext(table._tbl)
        except Exception:
            marker._element.addnext(table._tbl)
        try:
            parent = marker._element.getparent()
            parent.remove(marker._element)
        except Exception:
            for r in marker.runs:
                r.text = ''


def gen_vidomist_summary(session_meta, sheets, by_sheet_students, template=None, out=None, user_name='Система'):
    """
    1) DocxTemplate.render шапки
    2) save → відкрити як чистий Document
    3) вставити таблицю (валідний docx)
    """
    from io import BytesIO
    from docx import Document as DocxDocument

    session_meta = dict(session_meta) if not isinstance(session_meta, dict) else session_meta
    sheets = [dict(s) if not isinstance(s, dict) else s for s in (sheets or [])]

    def _g(st, key, default=None):
        if isinstance(st, dict):
            return st.get(key, default)
        try:
            return st[key]
        except Exception:
            return default

    order, names, grades = [], {}, {}
    for sh in sheets:
        for st in (by_sheet_students.get(sh['id']) or []):
            sid = _g(st, 'student_id') or _g(st, 'full_name')
            try:
                sid_key = int(sid)
            except (TypeError, ValueError):
                sid_key = sid
            if sid_key not in names:
                order.append(sid_key)
                names[sid_key] = _g(st, 'full_name') or ''
                grades[sid_key] = {}
            grades[sid_key][sh['id']] = {
                'national_grade': _g(st, 'national_grade') or '',
                'points': _g(st, 'points'),
                'ects': _g(st, 'ects') or '',
            }

    form = session_meta.get('study_form') or ''
    ctx = {
        'institution': session_meta.get('institution_name_and_status')
            or 'ПВНЗ „Європейський університет“\\nЛьвівська філія',
        'group_name': session_meta.get('group_name') or '',
        'course': _vidomist_roman(session_meta.get('course')),
        'study_form': form,
        'academic_year': session_meta.get('academic_year') or '',
        'semester': str(session_meta.get('semester') or ''),
        'degree_level': session_meta.get('degree_level') or '',
        'specialty': _vidomist_format_specialty(
            session_meta.get('specialty_code'), session_meta.get('specialty')),
        'session_name': session_meta.get('name') or session_meta.get('session_name') or '',
        'grades_table': '<<<GRADES_TABLE>>>',
    }

    if template is None:
        template = _vidomist_summary_template_path()
    if not os.path.exists(template):
        global_logger.error(f"Шаблон загальної відомості не знайдено: {template}")
        raise FileNotFoundError(f"Шаблон загальної відомості не знайдено: {template}")

    try:
        tpl = DocxTemplate(template)
        tpl.render(ctx)
        tmp = BytesIO()
        tpl.save(tmp)
        tmp.seek(0)
        document = DocxDocument(tmp)
    except Exception as e:
        global_logger.error(f"Помилка render загальної відомості: {e}")
        raise

    try:
        _vidomist_insert_summary_table(document, sheets, order, names, grades)
    except Exception as e:
        global_logger.warning(f"Загальна відомість: таблиця: {e}", exc_info=True)

    try:
        log_action(user_name, f"згенерував загальну відомість сесії '{ctx.get('session_name')}'", None)
    except Exception:
        pass

    if out:
        document.save(out)
        return out
    buf = BytesIO()
    document.save(buf)
    buf.seek(0)
    return buf
