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
    from routes.diploma_supplements import ensure_diploma_supplement_schema
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        ensure_diploma_supplement_schema(conn)
        is_reduced_program = is_student_on_reduced_program(conn, student_id, group_id)

        query = """
            SELECT s.id, s.code, s.name, COALESCE(s.name_en, '') AS name_en,
                   s.credits, s.reduced_credits, s.type, s.reduced_type, s.position, IFNULL(g.grade, '') AS grade
            FROM subjects s
            LEFT JOIN grades g ON g.subject_id = s.id AND g.student_id = ?
            WHERE s.group_id = ?
        """
        params = [student_id, group_id]
        query += program_track_condition(is_reduced_program, 's')
        query += " ORDER BY s.position"

        results = conn.execute(query, params).fetchall()
        subjects = [dict(r) for r in results]
        valid_subjects = []
        for subject in subjects:
            if all(key in subject for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade']):
                apply_track_overrides(subject, is_reduced_program)
                subject.pop('reduced_credits', None)
                subject.pop('reduced_type', None)
                subject = {k: clean_text(v) for k, v in subject.items()}
                subject['grade'] = format_grade(subject['grade'], subject['type']) if subject['grade'] else ''
                subject['display_name'] = (
                    f"{subject['name']} / {subject['name_en']}" if subject.get('name_en') else subject['name']
                )
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
    from routes.diploma_supplements import ensure_diploma_supplement_schema
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        ensure_diploma_supplement_schema(conn)
        is_reduced_program = is_student_on_reduced_program(conn, student_id, group_id)

        query = """
            SELECT p.id, p.code, p.name, COALESCE(p.name_en, '') AS name_en,
                   p.credits, p.reduced_credits, p.type, p.reduced_type, p.position, IFNULL(ag.grade, '') AS grade
            FROM practices p
            LEFT JOIN activity_grades ag ON ag.entity_id = p.id AND ag.entity_type = 'practice' AND ag.student_id = ?
            WHERE p.group_id = ?
        """
        params = [student_id, group_id]
        query += program_track_condition(is_reduced_program, 'p')
        query += " ORDER BY p.position"

        results = conn.execute(query, params).fetchall()
        practices = [dict(r) for r in results]
        valid_practices = []
        for practice in practices:
            if all(key in practice for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade']):
                apply_track_overrides(practice, is_reduced_program)
                practice.pop('reduced_credits', None)
                practice.pop('reduced_type', None)
                practice = {k: clean_text(v) for k, v in practice.items()}
                practice['grade'] = format_grade(practice['grade'], practice['type']) if practice['grade'] else ''
                practice['display_name'] = (
                    f"{practice['name']} / {practice['name_en']}" if practice.get('name_en') else practice['name']
                )
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
    from routes.diploma_supplements import ensure_diploma_supplement_schema
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        ensure_diploma_supplement_schema(conn)
        is_reduced_program = is_student_on_reduced_program(conn, student_id, group_id)

        query = """
            SELECT c.id, c.code, c.name, COALESCE(c.name_en, '') AS name_en,
                   c.credits, c.reduced_credits, c.type, c.reduced_type, c.position, IFNULL(ag.grade, '') AS grade
            FROM courseworks c
            LEFT JOIN activity_grades ag ON ag.entity_id = c.id AND ag.entity_type = 'coursework' AND ag.student_id = ?
            WHERE c.group_id = ?
        """
        params = [student_id, group_id]
        query += program_track_condition(is_reduced_program, 'c')
        query += " ORDER BY c.position"

        results = conn.execute(query, params).fetchall()
        courseworks = [dict(r) for r in results]
        valid_courseworks = []
        for coursework in courseworks:
            if all(key in coursework for key in ['id', 'code', 'name', 'credits', 'type', 'position', 'grade']):
                apply_track_overrides(coursework, is_reduced_program)
                coursework.pop('reduced_credits', None)
                coursework.pop('reduced_type', None)
                coursework = {k: clean_text(v) for k, v in coursework.items()}
                coursework['grade'] = format_grade(coursework['grade'], coursework['type']) if coursework['grade'] else ''
                coursework['display_name'] = (
                    f"{coursework['name']} / {coursework['name_en']}" if coursework.get('name_en') else coursework['name']
                )
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
    from routes.diploma_supplements import ensure_diploma_supplement_schema
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        ensure_diploma_supplement_schema(conn)
        is_reduced_program = is_student_on_reduced_program(conn, student_id, group_id)

        query = """
            SELECT a.id, a.code, a.name, COALESCE(a.name_en, '') AS att_name_en,
                   a.credits, a.reduced_credits, a.type, a.reduced_type,
                   a.position, ag.id AS ag_id, IFNULL(ag.grade, '') AS grade,
                   IFNULL(ag.name, '') AS student_name, IFNULL(ag.name_en, '') AS topic_en
            FROM attestations a
            LEFT JOIN activity_grades ag
                   ON ag.entity_id = a.id AND ag.entity_type = 'attestation' AND ag.student_id = ?
            WHERE a.group_id = ?
        """
        params = [student_id, group_id]
        query += program_track_condition(is_reduced_program, 'a')
        query += " ORDER BY a.position"

        results = conn.execute(query, params).fetchall()
        attestations = [dict(r) for r in results]
        valid_attestations = []
        cache_dirty = False

        # Імпорт тут, а не зверху модуля — щоб уникнути циклічної залежності
        # (documents.py імпортує admin/__init__ тощо).
        from routes.admin.documents import translate_to_en

        for attestation in attestations:
            if all(key in attestation for key in ('id', 'code', 'name', 'credits', 'type',
                                                   'position', 'grade', 'student_name')):
                raw_ag_id = attestation.pop('ag_id', None)
                saved_att_en = (attestation.pop('att_name_en', '') or '').strip()
                saved_topic_en = (attestation.pop('topic_en', '') or '').strip()
                apply_track_overrides(attestation, is_reduced_program)
                attestation.pop('reduced_credits', None)
                attestation.pop('reduced_type', None)
                attestation = {k: clean_text(v) for k, v in attestation.items()}
                attestation['grade'] = format_grade(attestation['grade'], attestation['type']) if attestation['grade'] else ''

                # ── Складання комбінованого рядка «Назва атестації «Тема» / Назва атестації (EN) «Тема (EN)»» ──
                name_ua = (attestation.get('name') or '').strip()
                topic_ua = (attestation.get('student_name') or '').strip()
                name_en = saved_att_en
                topic_en = saved_topic_en

                if name_ua and not name_en:
                    name_en = translate_to_en(name_ua) or ''
                if topic_ua and not topic_en:
                    topic_en = translate_to_en(topic_ua) or ''
                    if topic_en and raw_ag_id:
                        try:
                            conn.execute(
                                "UPDATE activity_grades SET name_en = ? WHERE id = ? AND COALESCE(TRIM(name_en), '') = ''",
                                (topic_en, raw_ag_id),
                            )
                            cache_dirty = True
                        except Exception:
                            pass

                attestation['name_en'] = name_en
                attestation['topic_ua'] = topic_ua
                attestation['topic_en'] = topic_en

                if name_ua and topic_ua:
                    ua_part = f'{name_ua} «{topic_ua}»'
                    if name_en and topic_en:
                        en_part = f'{name_en} «{topic_en}»'
                        attestation['display_name'] = f'{ua_part} / {en_part}'
                    elif name_en:
                        attestation['display_name'] = f'{ua_part} / {name_en}'
                    else:
                        # Переклад недоступний (напр. ліміт DeepL) — лишаємо лише укр. частину
                        attestation['display_name'] = ua_part
                elif topic_ua:
                    # Назва атестації порожня — виводимо лише тему
                    attestation['display_name'] = f'{topic_ua} / {topic_en}' if topic_en else topic_ua
                else:
                    attestation['display_name'] = f'{name_ua} / {name_en}' if (name_ua and name_en) else name_ua

                valid_attestations.append(attestation)
            else:
                global_logger.warning(f"Неполные данные аттестации пропущены: {attestation}")
        if cache_dirty:
            conn.commit()
        return valid_attestations
    except Exception as e:
        global_logger.error(f"[get_attestation_data] Ошибка: {e}")
        return []
    finally:
        conn.close()

def apply_reduced_program_texts(student_dict):
    """Якщо студент навчається за скороченою програмою - підміняє тексти."""
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
        if field.endswith('_en'):
            reduced_key = f"{field[:-3]}_reduced_en"
        else:
            reduced_key = f"{field}_reduced"
        reduced_value = student_dict.get(reduced_key)
        if reduced_value:
            student_dict[field] = reduced_value

    student_dict['program_credits'] = override


def gen_doc(
    student: dict,
    military: dict,
    template='template.docx',
    out='out.docx',
    user_name='Система',
    *,
    record_log=False,
    capture_context=None,
):
    """Генерирует документ для студента на основе шаблона."""
    from routes.diploma_supplements import (
        compute_student_diploma_metrics,
        record_diploma_generation,
    )
    global_logger.debug(f"Запуск gen_doc: student_id={student.get('id', 'unknown')}, template={template}, out={out}")

    if not os.path.exists(template):
        global_logger.error(f"Шаблон {template} не найден")
        raise FileNotFoundError(f"Шаблон {template} не найден")

    try:
        doc = DocxTemplate(template)
    except Exception as e:
        global_logger.error(f"Ошибка при загрузке шаблона {template}: {str(e)}")
        raise

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

    student_dict = {k: clean_text(v) for k, v in dict(student).items()}
    military_dict = {k: clean_text(v) for k, v in dict(military).items()} if military else {}

    conn = get_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        current_group_id = student_dict.get('group_id')
        names_ua, dates_ua = get_study_periods(cursor, student_dict['id'], group_id=current_group_id, lang='ua')
        names_en, dates_en = get_study_periods(cursor, student_dict['id'], group_id=current_group_id, lang='en')
    except Exception as e:
        global_logger.error(f"Ошибка при получении периодов навчання: {e}")
        names_ua = "Львівська філія Приватного вищого навчального закладу «Європейський університет»"
        dates_ua = ''
        names_en = "Lviv Branch of Private Higher Education Establishment «European University»"
        dates_en = ''
    finally:
        conn.close()

    apply_reduced_program_texts(student_dict)

    if education_docs:
        for key, value in dict(education_docs).items():
            student_dict[key] = clean_text(value) if value else ''

    birth_date = student_dict.get('birth_date', '')
    if birth_date:
        try:
            date_obj = datetime.strptime(birth_date, '%d.%m.%Y')
            birth_date = date_obj.strftime('%d/%m/%Y')
        except ValueError:
            try:
                date_obj = datetime.strptime(birth_date, '%Y-%m-%d')
                birth_date = date_obj.strftime('%d/%m/%Y')
            except ValueError:
                birth_date = student_dict['birth_date']
    student_dict['birth_date'] = birth_date

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
        elif credits == 90:
            study_years = '1.5'
        else:
            study_years = str(credits // 60)
    except (ValueError, TypeError):
        study_years = ''
    student_dict['study_years'] = study_years

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
    student_dict['study_form_eu'] = study_form_eu

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
                if credits in (90, 120):
                    end_year = str(year + 2)
            else:
                end_year = str(year + (credits // 60))
    except (ValueError, TypeError) as e:
        global_logger.warning(f"Ошибка при расчёте end_year: {e}")
        end_year = ''
    student_dict['end_year'] = end_year
    student_dict['end_year_short'] = student_dict['end_year'][-2:] if student_dict['end_year'] else ''

    if degree_level == "Магістр":
        student_dict['top_qualification_text'] = "- підготовка кваліфікаційної роботи / preparation of qualification work"
        student_dict['bottom_qualification_text'] = "- захист кваліфікаційної роботи / defense of qualification work"

    fields_to_process = ['program_includes', 'program_includes_en', 'learning_outcomes', 'learning_outcomes_en']
    for field in fields_to_process:
        if field in student_dict and student_dict[field]:
            lines = student_dict[field].split('\n')
            cleaned_lines = [line.strip() for line in lines if line.strip()]
            student_dict[field] = cleaned_lines

    context = {**student_dict, **military_dict}
    if not (context.get('faculty') or '').strip():
        try:
            conn_f = get_db()
            conn_f.row_factory = sqlite3.Row
            if student_dict.get('group_id'):
                frow = conn_f.execute("""
                    SELECT COALESCE(NULLIF(TRIM(g.faculty), ''), fl.name_ua, '') AS faculty,
                           COALESCE(NULLIF(TRIM(g.faculty_en), ''), fl.name_en, '') AS faculty_en
                    FROM groups g
                    LEFT JOIN specialties sp ON sp.code = g.specialty_code
                    LEFT JOIN faculties fl ON fl.id = COALESCE(g.faculty_id, sp.faculty_id)
                    WHERE g.id = ?
                """, (student_dict['group_id'],)).fetchone()
                if frow:
                    context['faculty'] = frow['faculty'] or ''
                    if not (context.get('faculty_en') or '').strip():
                        context['faculty_en'] = frow['faculty_en'] or ''
            elif student_dict.get('specialty_code'):
                from routes.utils import lookup_specialty_faculty
                f_ua, f_en = lookup_specialty_faculty(conn_f, student_dict.get('specialty_code'))
                context['faculty'] = f_ua
                if not (context.get('faculty_en') or '').strip():
                    context['faculty_en'] = f_en
            conn_f.close()
        except Exception as e:
            global_logger.warning(f"Не вдалося визначити факультет у gen_doc: {e}")
    context.setdefault('faculty', '')
    context.setdefault('faculty_en', '')

    photo_rel = student_dict.get('photo')
    if photo_rel:
        photo_full_path = os.path.join(os.getcwd(), 'static', photo_rel)
        if os.path.exists(photo_full_path):
            try:
                context['student_photo'] = InlineImage(doc, photo_full_path, width=Mm(30))
            except Exception as e:
                global_logger.error(f"Не вдалося вставити фото студента (ID {student_dict.get('id')}): {e}")

    try:
        if 'group_id' in student_dict and 'id' in student_dict:
            context['subjects_grades'] = get_subjects_grades(student_dict['id'], student_dict['group_id']) or []
            context['practice_data'] = get_practice_data(student_dict['id'], student_dict['group_id']) or []
            context['coursework_data'] = get_coursework_data(student_dict['id'], student_dict['group_id']) or []
            context['attestation_data'] = get_attestation_data(student_dict['id'], student_dict['group_id']) or []
    except Exception as e:
        global_logger.error(f"Ошибка при получении данных для диплома: {e}")
        raise

    try:
        conn = get_db()
        conn.row_factory = sqlite3.Row
        acc = conn.execute("""
            SELECT text_ua, text_en
            FROM accreditations
            WHERE degree = ? AND specialty = ?
            ORDER BY id DESC LIMIT 1
        """, (student_dict.get('degree_level', ''), student_dict.get('specialty', ''))).fetchone()
        conn.close()
        if acc:
            context['accreditation_text'] = acc['text_ua']
            context['accreditation_text_en'] = acc['text_en']
        else:
            context['accreditation_text'] = ''
            context['accreditation_text_en'] = ''
    except Exception as e:
        global_logger.error(f"Ошибка при получении аккредитации: {e}")
        context['accreditation_text'] = ''
        context['accreditation_text_en'] = ''

    metrics = {}
    if student_dict.get('id'):
        try:
            conn_m = get_db()
            conn_m.row_factory = sqlite3.Row
            try:
                metrics = compute_student_diploma_metrics(
                    conn_m,
                    int(student_dict['id']),
                    int(student_dict['group_id']) if student_dict.get('group_id') else None,
                )
            finally:
                conn_m.close()
        except Exception as e:
            global_logger.warning(f"Помилка розрахунку метрик диплома: {e}")

    if metrics:
        diploma_with_honours = bool(metrics.get('final_honours'))
    else:
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
            elif ' Fx' in grade_text or ' F' in grade_text or 'Незадовільно' in grade_text or 'Fail' in grade_text:
                fail_count += 1

        attestation_grade = ''
        if context.get('attestation_data'):
            attestation_grade = next(
                (g.get('grade', '') for g in context['attestation_data'] if g.get('grade')), ''
            )

        diploma_with_honours = (
            total_grades > 0
            and excellent_count / total_grades >= 0.75
            and satisfactory_count == 0
            and fail_count == 0
            and ' A' in attestation_grade
        )

    context['diploma_with_honours_bool'] = diploma_with_honours
    if diploma_with_honours:
        context['diploma_with_honor_text'] = 'Диплом з відзнакою'
        context['diploma_with_honor_text_en'] = 'Diploma with honours'
    else:
        context['diploma_with_honor_text'] = 'Інформація відсутня'
        context['diploma_with_honor_text_en'] = 'Information is absent'

    context['gpa_100'] = metrics['gpa_100'] if metrics.get('gpa_100') is not None else ''
    context['weighted_gpa_100'] = metrics['weighted_gpa_100'] if metrics.get('weighted_gpa_100') is not None else ''
    context['gpa_5'] = metrics['gpa_5'] if metrics.get('gpa_5') is not None else ''
    context['total_ects_credits'] = metrics.get('total_ects_credits') or ''
    context['total_hours'] = metrics.get('total_hours') or ''
    context['excellent_pct'] = metrics.get('excellent_pct', 0)
    context['excellent_count'] = metrics.get('excellent_count', 0)
    context['total_graded_count'] = metrics.get('graded_count', 0)
    context['graduation_order_number'] = metrics.get('graduation_order_number', '')
    context['graduation_order_date'] = metrics.get('graduation_order_date', '')
    context['graduation_order_date_slash'] = metrics.get('graduation_order_date_slash', '')

    conn = get_db()
    conn.row_factory = sqlite3.Row
    diploma_row = conn.execute("""
        SELECT diploma_number, appendix_number
        FROM diplomas
        WHERE student_id = ? AND group_id = ?
        ORDER BY id DESC LIMIT 1
    """, (student['id'], student_dict.get('group_id'))).fetchone()
    conn.close()

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

        def _institution_short_name(full_text):
            if not full_text:
                return ''

            quote_pairs = [
                ('«', '»'),
                ('“', '”'),
                ('„', '“'),
                ('"', '"'),
                ("'", "'"),
            ]

            for open_q, close_q in quote_pairs:
                start = full_text.find(open_q)
                if start == -1:
                    continue
                end = full_text.find(close_q, start + len(open_q))
                if end == -1:
                    continue
                return full_text[:end + len(close_q)].strip()

            return ''

        names_ua = _institution_short_name(student_dict.get('institution_name_and_status', '')) or names_ua
        names_en = _institution_short_name(student_dict.get('institution_name_and_status_en', '')) or names_en

    context['study_period_names'] = _to_richtext_multiline(names_ua, font_size_pt=8, font_name='Times New Roman')
    context['study_period_dates'] = _to_richtext_multiline(dates_ua, font_size_pt=8, font_name='Times New Roman')
    context['study_period_names_en'] = _to_richtext_multiline(names_en, font_size_pt=8, font_name='Times New Roman')
    context['study_period_dates_en'] = _to_richtext_multiline(dates_en, font_size_pt=8, font_name='Times New Roman')

    if diploma_row:
        diploma_number = diploma_row['diploma_number'] or ''
        appendix_number = diploma_row['appendix_number'] or ''
        diploma_number = diploma_number.zfill(6) if diploma_number else ''
        appendix_number = appendix_number.zfill(6) if appendix_number else ''
        context['diploma_number'] = diploma_number
        context['appendix_number'] = appendix_number
    else:
        context['diploma_number'] = ''
        context['appendix_number'] = ''

    if isinstance(capture_context, dict):
        capture_context.update(context)

    try:
        doc.render(context)
    except Exception as e:
        global_logger.error(f"Ошибка при рендеринге документа: {e}")
        raise

    student_name = f"{student_dict.get('last_name_UA', '')} {student_dict.get('first_name_UA', '')}".strip()
    log_action(user_name, f"згенерував документ '{out}' для студента {student_name}", student_dict.get('group_id'))

    try:
        doc.save(out)
    except Exception as e:
        global_logger.error(f"Ошибка при сохранении документа {out}: {e}")
        raise

    if record_log and student_dict.get('id'):
        try:
            conn_log = get_db()
            try:
                record_diploma_generation(
                    conn_log,
                    int(student_dict['id']),
                    int(student_dict['group_id']) if student_dict.get('group_id') else None,
                    template,
                    user_name,
                    metrics=metrics or None,
                )
            finally:
                conn_log.close()
        except Exception as e:
            global_logger.warning(f"Не вдалося записати в журнал генерації дипломів: {e}")

    return out



# ── Відомість сесії (той самий стек що gen_doc: DocxTemplate) ──────

def _vidomist_template_path(is_fulltime=False):
    filename = 'vidomist_sheet_denna.docx' if is_fulltime else 'vidomist_sheet.docx'
    candidates = [
        os.path.join('static', 'templates', 'word', filename),
        os.path.join(os.path.dirname(__file__), '..', 'static', 'templates', 'word', filename),
    ]
    for p in candidates:
        ap = os.path.abspath(p)
        if os.path.isfile(ap):
            return ap
    fallback = [
        os.path.join('static', 'templates', 'word', 'vidomist_sheet.docx'),
        os.path.join(os.path.dirname(__file__), '..', 'static', 'templates', 'word', 'vidomist_sheet.docx'),
    ]
    for p in fallback:
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

def _fmt_exam_date_ua(iso_date_str):
    """YYYY-MM-DD → «15 грудня 2026». Если не ISO — вернуть как есть."""
    if not iso_date_str:
        return ''
    s = str(iso_date_str).strip()
    import re as _re
    m = _re.match(r'^(\d{4})-(\d{2})-(\d{2})$', s)
    if not m:
        return s
    year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
    months_ua = {
        1: 'січня', 2: 'лютого', 3: 'березня', 4: 'квітня',
        5: 'травня', 6: 'червня', 7: 'липня', 8: 'серпня',
        9: 'вересня', 10: 'жовтня', 11: 'листопада', 12: 'грудня',
    }
    name = months_ua.get(month)
    if not name:
        return s
    return f"{day} {name} {year}"


def _is_fulltime_form(study_form) -> bool:
    """True якщо форма навчання — денна."""
    s = (study_form or '').strip().lower()
    return s in ('денна', 'denna', 'full-time', 'full', 'очна')


def _format_teacher_short(full_or_short: str) -> str:
    """
    Приводить ПІБ викладача («ПІБ (повне)» з картотеки) до скороченого формату
    «Гушак О. О.» для колонки «Підпис викладача».
    Прибирає наукові ступені/посади (к.ф.-м.н., доц., проф. тощо).
    """
    if not full_or_short:
        return ''
    s = str(full_or_short).strip()
    import re as _re
    # Якщо вже скорочений формат «Прізвище І. П.» або «Прізвище І.П.»
    m = _re.match(r'^([А-ЯІЇЄҐA-Z][а-яіїєґa-z\'’\-]+)\s+([А-ЯІЇЄҐA-Z])\.\s*([А-ЯІЇЄҐA-Z])\.?$', s)
    if m:
        return f"{m.group(1)} {m.group(2)}. {m.group(3)}."

    degree_re = _re.compile(r'^[кд]\.[а-яіїєґ\.\-]*н\.?$', _re.I)
    skip = {
        'доц', 'доц.', 'доцент', 'проф', 'проф.', 'професор',
        'ст', 'ст.', 'викл', 'викл.', 'ст.викл', 'ст.викл.',
        'асист', 'асист.', 'асистент', 'phd', 'ph.d', 'ph.d.'
    }

    raw_parts = _re.split(r'[\s,]+', s)
    clean = []
    for p in raw_parts:
        p_strip = p.strip(' ,;')
        if not p_strip:
            continue
        pl = p_strip.lower().rstrip('.')
        if pl in skip or degree_re.match(p_strip):
            continue
        clean.append(p_strip)
    if not clean:
        return s
    if len(clean) >= 3:
        return f"{clean[0]} {clean[1][0].upper()}. {clean[2][0].upper()}."
    if len(clean) == 2:
        m2 = _re.match(r'^([А-ЯІЇЄҐA-Z])\.?\s*([А-ЯІЇЄҐA-Z])\.?$', clean[1])
        if m2:
            return f"{clean[0]} {m2.group(1)}. {m2.group(2)}."
        return f"{clean[0]} {clean[1][0].upper()}."
    return clean[0]


def _vidomist_build_context(sheet, students):
    """Контекст для DocxTemplate — ті самі ключі що {{ }} у шаблоні."""
    sheet = dict(sheet) if not isinstance(sheet, dict) else sheet
    control = sheet.get('control_form') or 'Екзамен'
    teacher_short = sheet.get('teacher_short')
    teacher_full = sheet.get('teacher_full')
    teacher_raw = sheet.get('teacher_name') or ''
    teacher_ref_id = sheet.get('teacher_ref_id')

    # Якщо teacher_short або teacher_full не підтягнуті через JOIN, довантажуємо з картотеки teachers
    if teacher_ref_id and (not teacher_short or not teacher_full):
        try:
            from routes.db import get_db
            _conn = get_db()
            _t = _conn.execute(
                "SELECT short_name, full_name FROM teachers WHERE id = ?",
                (teacher_ref_id,)
            ).fetchone()
            _conn.close()
            if _t:
                teacher_short = teacher_short or _t['short_name']
                teacher_full = teacher_full or _t['full_name']
        except Exception:
            pass

    # У шапці «Викладач ...»: "Скорочене ім'я" з картотеки (напр. «к.т.н., Гущак О.М.»)
    teacher_header = (teacher_short or teacher_raw or teacher_full or '').strip()

    # У колонці таблиці «Підпис викладача»: скорочується з «ПІБ (повне)» (напр. «Гушак О. О.»)
    teacher_sign = _format_teacher_short(
        teacher_full or teacher_short or teacher_raw
    )
    form = sheet.get('study_form') or ''
    is_fulltime = _is_fulltime_form(form)

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
            'points_mid': '' if is_fulltime else _v(st.get('points_mid')),
            'points_current': '' if is_fulltime else _v(st.get('points_current')),
            'points_exam': '' if is_fulltime else _v(st.get('points_exam')),
            'points': _v(pts),
            'national_grade': nat or '',
            'ects': ects or '',
            'teacher_sign': st.get('teacher_sign') or teacher_sign,
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

    # form / is_fulltime вже визначені вище
    num = sheet.get('statement_number') or ''
    inst = (sheet.get('institution_name_and_status') or
            'ПВНЗ „Європейський університет“')
    if '  ' in inst and '\n' not in inst and 'філія' in inst.lower():
        import re
        m = re.match(r'^(.*університет[^\"]*[\"“]?)\s+(Львівська\s+філія)\s*$', inst, re.I)
        if m:
            inst = m.group(1).strip() + '\n' + m.group(2).strip()

    from routes.utils import resolve_vidomist_faculty
    knowledge_area = (sheet.get('knowledge_area') or '').strip()
    _conn_v = None
    try:
        sp_ref = (sheet.get('specialty_code') or sheet.get('specialty') or '').strip()
        need_kf = not knowledge_area and bool(sheet.get('specialty_code'))
        need_fac = not (sheet.get('faculty') or '').strip() and bool(sp_ref)
        if need_kf or need_fac:
            from routes.db import get_db
            _conn_v = get_db()
        if need_kf and _conn_v is not None:
            _kf = _conn_v.execute(
                """
                SELECT k.code, k.name_ua
                FROM specialties s
                JOIN knowledge_fields k ON k.code = s.knowledge_field_code
                WHERE s.code = ?
                """,
                (sheet.get('specialty_code'),),
            ).fetchone()
            if _kf:
                knowledge_area = f"{_kf['code']} {_kf['name_ua']}".strip()
        faculty_ua, faculty_en, _ = resolve_vidomist_faculty(
            {**sheet, 'knowledge_area': knowledge_area},
            conn=_conn_v,
        )
    except Exception:
        faculty_ua = (sheet.get('faculty') or '').strip() or knowledge_area or (sheet.get('department') or '').strip() or ''
        faculty_en = (sheet.get('faculty_en') or '').strip()
    finally:
        if _conn_v is not None:
            try:
                _conn_v.close()
            except Exception:
                pass

    ctx = {
        'institution': inst,
        'faculty': faculty_ua,
        'faculty_en': faculty_en,
        'knowledge_area': knowledge_area,
        'specialty': _vidomist_format_specialty(sheet.get('specialty_code'), sheet.get('specialty')),
        'course': _vidomist_roman(sheet.get('course')),
        'group_name': sheet.get('group_name') or '',
        'study_form': form,
        'study_form_part': f'({form})' if form else '',
        'academic_year': sheet.get('academic_year') or '',
        'degree_level': sheet.get('degree_level') or '',
        'statement_title': f'ВІДОМІСТЬ  № {num}.'.strip() if num else 'ВІДОМІСТЬ',
        'statement_number': num,
        'exam_date': _fmt_exam_date_ua(sheet.get('exam_date')),
        'discipline_name': sheet.get('discipline_name') or '',
        'semester_word': _vidomist_semester_word(sheet.get('semester')),
        'semester': str(sheet.get('semester') or ''),
        'control_form': (control or '').lower(),
        'hours_credits': sheet.get('hours_credits') or '',
        'teacher_name': teacher_header,
        'teacher_sign': teacher_sign,
        'is_fulltime': is_fulltime,
        'entity_kind': sheet.get('entity_kind') or '',
        'include_topic': (sheet.get('entity_kind') or '').strip() == 'attestation',
        'total_label': 'Кількість балів за 100 бальною' if is_fulltime else 'Підсумкова оцінка',
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


def _vidomist_insert_topic_column(table, after_grid_col, header_rows=2, title='Тема роботи'):
    """
    Вставляє нову колонку (за замовчуванням «Тема роботи») у таблицю відомості
    одразу після колонки з індексом `after_grid_col` (№ індив. навч. плану),
    пропорційно стискаючи ширини інших колонок, щоб таблиця не виходила за поля.
    Коректно обробляє gridSpan, vMerge та tblGrid/tcW.
    """
    from copy import deepcopy
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tbl = table._tbl
    tblGrid = tbl.tblGrid
    if tblGrid is None:
        return

    gridCols = tblGrid.findall(qn('w:gridCol'))
    if not gridCols or after_grid_col < 0 or after_grid_col >= len(gridCols):
        return

    widths = []
    for gc in gridCols:
        try:
            widths.append(int(gc.get(qn('w:w')) or 1200))
        except (TypeError, ValueError):
            widths.append(1200)

    total_w = sum(widths)
    if total_w <= 0:
        total_w = 1200 * len(widths)
        widths = [1200] * len(widths)

    topic_w = min(2400, max(1500, int(round(total_w * 0.17))))
    reducible = [max(0, w - (400 if i == 0 else 550)) for i, w in enumerate(widths)]
    total_reducible = sum(reducible)

    shrunk = list(widths)
    if total_reducible > 0:
        take = min(topic_w, total_reducible)
        taken = 0
        for i, red in enumerate(reducible):
            if red <= 0:
                continue
            delta = int(round(take * (red / total_reducible)))
            delta = min(delta, red, take - taken)
            shrunk[i] -= delta
            taken += delta
        rem = take - taken
        if rem > 0:
            order_idx = sorted(range(len(widths)), key=lambda i: shrunk[i] - (400 if i == 0 else 550), reverse=True)
            for i in order_idx:
                if rem <= 0:
                    break
                avail = shrunk[i] - (400 if i == 0 else 550)
                if avail > 0:
                    step = min(avail, rem)
                    shrunk[i] -= step
                    rem -= step

    for gc, w in zip(gridCols, shrunk):
        gc.set(qn('w:w'), str(w))

    new_gc = OxmlElement('w:gridCol')
    new_gc.set(qn('w:w'), str(topic_w))
    gridCols[after_grid_col].addnext(new_gc)

    new_widths = shrunk[:after_grid_col + 1] + [topic_w] + shrunk[after_grid_col + 1:]

    def _set_tc_width(tc, w_val):
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)
        tcW = tcPr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW')
            tcPr.insert(0, tcW)
        tcW.set(qn('w:w'), str(w_val))
        tcW.set(qn('w:type'), 'dxa')

    def _get_span(tc):
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            return 1
        gs = tcPr.find(qn('w:gridSpan'))
        if gs is None:
            return 1
        try:
            return max(1, int(gs.get(qn('w:val')) or 1))
        except (TypeError, ValueError):
            return 1

    def _set_span(tc, val):
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)
        gs = tcPr.find(qn('w:gridSpan'))
        if val <= 1:
            if gs is not None:
                tcPr.remove(gs)
        else:
            if gs is None:
                gs = OxmlElement('w:gridSpan')
                tcPr.append(gs)
            gs.set(qn('w:val'), str(val))

    def _prepare_new_tc(ref_tc, row_idx):
        new_tc = deepcopy(ref_tc)
        tcPr = new_tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            new_tc.insert(0, tcPr)

        for tag in ('w:gridSpan', 'w:noWrap', 'w:vMerge'):
            el = tcPr.find(qn(tag))
            if el is not None:
                tcPr.remove(el)

        if row_idx == 0 and header_rows > 1:
            vm = OxmlElement('w:vMerge')
            vm.set(qn('w:val'), 'restart')
            tcPr.append(vm)
        elif 0 < row_idx < header_rows:
            vm = OxmlElement('w:vMerge')
            tcPr.append(vm)

        _set_tc_width(new_tc, topic_w)

        ps = new_tc.findall(qn('w:p'))
        if not ps:
            p = OxmlElement('w:p')
            new_tc.append(p)
        else:
            p = ps[0]
            for extra_p in ps[1:]:
                new_tc.remove(extra_p)

        runs = p.findall(qn('w:r'))
        if row_idx == 0:
            if runs:
                first_r = runs[0]
                for extra_r in runs[1:]:
                    p.remove(extra_r)
                for child in list(first_r):
                    if child.tag != qn('w:rPr'):
                        first_r.remove(child)
                t_el = OxmlElement('w:t')
                t_el.text = title
                first_r.append(t_el)
            else:
                r = OxmlElement('w:r')
                rPr = OxmlElement('w:rPr')
                rFonts = OxmlElement('w:rFonts')
                rFonts.set(qn('w:ascii'), 'Times New Roman')
                rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                rFonts.set(qn('w:cs'), 'Times New Roman')
                rPr.append(rFonts)
                rPr.append(OxmlElement('w:b'))
                r.append(rPr)
                t_el = OxmlElement('w:t')
                t_el.text = title
                r.append(t_el)
                p.append(r)
        elif row_idx < header_rows:
            for r in runs:
                p.remove(r)
        else:
            pPr = p.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                p.insert(0, pPr)
            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                pPr.append(jc)
            jc.set(qn('w:val'), 'left')

            if runs:
                first_r = runs[0]
                for extra_r in runs[1:]:
                    p.remove(extra_r)
                for child in list(first_r):
                    if child.tag != qn('w:rPr'):
                        first_r.remove(child)
                t_el = OxmlElement('w:t')
                t_el.text = ''
                first_r.append(t_el)

        return new_tc

    for ri, row in enumerate(table.rows):
        tcs = row._tr.findall(qn('w:tc'))
        old_g = 0
        new_g = 0
        for tc in tcs:
            span = _get_span(tc)
            if old_g + span - 1 == after_grid_col:
                _set_tc_width(tc, sum(new_widths[new_g:new_g + span]))
                new_tc = _prepare_new_tc(tc, ri)
                tc.addnext(new_tc)
                new_g += span + 1
                old_g += span
            elif old_g <= after_grid_col < old_g + span - 1:
                _set_span(tc, span + 1)
                _set_tc_width(tc, sum(new_widths[new_g:new_g + span + 1]))
                new_g += span + 1
                old_g += span
            else:
                _set_tc_width(tc, sum(new_widths[new_g:new_g + span]))
                new_g += span
                old_g += span


def _vidomist_fill_students_table(doc, students_rows, is_fulltime=False, include_topic=False):
    """
    Як insert_subjects_table: після render дописати рядки студентів.
    Шукає таблицю з «Прізвище» / «№» у шапці; клонує останній (зразок) рядок.
    Підтримує як окремий 7-колонковий шаблон (денна форма), так і 10-колонковий (заочна).
    Якщо `include_topic=True` (тип «Атестація»), додає колонку «Тема роботи» одразу після
    «№ індив. навч. плану» (якщо її ще немає у шаблоні) і заповнює її з `work_topic`.
    """
    from copy import deepcopy

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
        if any(x in head2 for x in ('міжсесій', 'поточн', 'іспит', 'ects', 'підсумков', 'нац', '100 баль')):
            header_rows = 2

    def _collect_header_texts():
        s_idx = header_rows if len(table.rows) > header_rows else max(0, header_rows - 1)
        cols_cnt = len(table.rows[s_idx].cells)
        res = [''] * cols_cnt
        for ri in range(min(header_rows, len(table.rows))):
            for i, c in enumerate(table.rows[ri].cells):
                if i >= cols_cnt:
                    break
                tx = (c.text or '').strip().lower().replace('\n', ' ')
                if tx:
                    res[i] = (res[i] + ' ' + tx).strip() if res[i] else tx
        return s_idx, res

    sample_idx, texts = _collect_header_texts()

    def find(*needles):
        for i, tx in enumerate(texts):
            for n in needles:
                if n in tx:
                    return i
        return None

    if include_topic and find('тема') is None:
        edebo_col = find('індив', 'план')
        name_col = find('прізвище', 'піб', 'студент')
        after_col = edebo_col if edebo_col is not None else (name_col if name_col is not None else 1)
        _vidomist_insert_topic_column(table, after_col, header_rows=header_rows, title='Тема роботи')
        sample_idx, texts = _collect_header_texts()

    sample_tr = deepcopy(table.rows[sample_idx]._tr)

    roles = {
        'n': find('№', 'з\\п', 'з/п') if find('№', 'з\\п', 'з/п') is not None else 0,
        'name': find('прізвище', 'піб', 'студент'),
        'edebo': find('індив', 'план'),
        'topic': find('тема') if include_topic else None,
        'mid': find('міжсесій'),
        'current': find('поточн'),
        'exam': find('іспит', 'залік'),
        'total': find('100 баль', 'кількість балів', 'підсумков'),
        'national': find('нац', 'національ'),
        'ects': find('ects'),
        'sign': find('підпис'),
    }

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
            p.runs[0].font.name = 'Times New Roman'
            for r in p.runs[1:]:
                r.text = ''
        else:
            r = p.add_run(text)
            r.font.name = 'Times New Roman'

    for s in students_rows:
        tbl.append(deepcopy(sample_tr))
        row = table.rows[-1]
        write_cell(row, roles.get('n'), s.get('n'))
        write_cell(row, roles.get('name'), s.get('full_name'))
        write_cell(row, roles.get('edebo'), s.get('edebo_code'))
        write_cell(row, roles.get('topic'), s.get('work_topic'))
        write_cell(row, roles.get('mid'), s.get('points_mid'))
        write_cell(row, roles.get('current'), s.get('points_current'))
        write_cell(row, roles.get('exam'), s.get('points_exam'))
        write_cell(row, roles.get('total'), s.get('points'))
        write_cell(row, roles.get('national'), s.get('national_grade'))
        write_cell(row, roles.get('ects'), s.get('ects'))
        write_cell(row, roles.get('sign'), s.get('teacher_sign'))


def gen_vidomist_sheet(sheet, students, template=None, out=None, user_name='Система'):
    """Відомість сесії — DocxTemplate.render + дописування рядків таблиці."""
    sheet_dict = dict(sheet) if not isinstance(sheet, dict) else sheet
    form = sheet_dict.get('study_form') or ''
    is_fulltime = _is_fulltime_form(form)
    include_topic = (sheet_dict.get('entity_kind') or '').strip() == 'attestation'

    if template is None:
        template = _vidomist_template_path(is_fulltime=is_fulltime)
    if not os.path.exists(template):
        global_logger.error(f"Шаблон відомості не знайдено: {template}")
        raise FileNotFoundError(f"Шаблон відомості не знайдено: {template}")

    context, rows = _vidomist_build_context(sheet_dict, students)

    try:
        doc = DocxTemplate(template)
    except Exception as e:
        global_logger.error(f"Помилка завантаження шаблону відомості: {e}")
        raise

    try:
        doc.render(context)
    except Exception as e:
        global_logger.error(f"Помилка render відомості: {e}")
        raise

    try:
        _vidomist_fill_students_table(
            doc,
            rows,
            is_fulltime=bool(context.get('is_fulltime')),
            include_topic=include_topic,
        )
    except Exception as e:
        global_logger.warning(f"Відомість: заповнення таблиці студентів: {e}")

    disc = (sheet.get('discipline_name') if isinstance(sheet, dict) else None) or ''
    log_action(user_name, f"згенерував відомість '{disc}'", None)

    if out:
        try:
            doc.save(out)
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


def _ensure_sectpr(document):
    """Гарантує, що в документі є body-level w:sectPr.

    Деякі .docx-шаблони (згенеровані не Word'ом, конвертовані,
    пошкоджені ручним редагуванням XML) не мають w:sectPr у
    w:body. python-docx у такому разі має порожній
    document.sections, і document.add_table() падає з:
        IndexError: list index out of range
        (document._block_width → self.sections[-1])

    Ця функція вставляє стандартний A4 landscape sectPr з полями
    ~1 см, якщо його немає. Word і далі вільно редагує документ.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    try:
        _ = document.sections[-1]
        return  # sectPr уже є
    except (IndexError, KeyError):
        pass

    body = document.element.body
    if body.find(qn('w:sectPr')) is not None:
        return

    sectPr = OxmlElement('w:sectPr')

    pgSz = OxmlElement('w:pgSz')
    pgSz.set(qn('w:w'), '16838')
    pgSz.set(qn('w:h'), '11906')
    pgSz.set(qn('w:orient'), 'landscape')
    sectPr.append(pgSz)

    pgMar = OxmlElement('w:pgMar')
    pgMar.set(qn('w:top'),    '567')
    pgMar.set(qn('w:right'),  '567')
    pgMar.set(qn('w:bottom'), '567')
    pgMar.set(qn('w:left'),   '567')
    pgMar.set(qn('w:header'), '708')
    pgMar.set(qn('w:footer'), '708')
    pgMar.set(qn('w:gutter'), '0')
    sectPr.append(pgMar)

    body.append(sectPr)


def _short_national_grade(nat):
    """Скорочення національної оцінки для компактного відображення в зведеній таблиці."""
    if not nat:
        return ''
    s = str(nat).strip()
    if '/' in s:
        s = s.split('/')[0].strip()
    low = s.lower()
    if 'відмінно' in low:
        return 'відм.'
    if 'добре' in low:
        return 'добре'
    if 'задовільно' in low and 'не' not in low:
        return 'задов.'
    if 'незадовільно' in low:
        return 'незад.'
    if 'зараховано' in low and 'не' not in low:
        return 'зар.'
    if 'незараховано' in low:
        return 'незар.'
    return s[:6]


def _vidomist_insert_summary_table(document, sheets, order, names, grades):
    """Вставляє зведену таблицю замість маркера `<<<GRADES_TABLE>>>`
    (або перед словом «декан», якщо маркера немає).

    Фінальна перевірка гарантує, що:
      - після КОЖНОЇ <w:tbl> у body стоїть <w:p> (вимога OOXML);
      - <w:sectPr> — рівно ОДИН і він — ОСТАННІЙ елемент <w:body>
        (Word цього вимагає строго, LibreOffice прощає).
    """
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    _ensure_sectpr(document)

    # Гарантуємо альбомну сторінку A4 з полями 1.0 см
    for s in document.sections:
        try:
            s.orientation = WD_ORIENT.LANDSCAPE
            s.page_width = Cm(29.7)
            s.page_height = Cm(21.0)
            s.left_margin = Cm(1.0)
            s.right_margin = Cm(1.0)
            s.top_margin = Cm(1.0)
            s.bottom_margin = Cm(1.0)
        except Exception:
            pass

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
    cols = max(3, 2 + n_disc * 3 + 1)
    avg_col = 2 + n_disc * 3 if n_disc > 0 else 2

    # Розрахунок ширин колонок під робочу ширину сторінки (A4 landscape ~15704 dxa)
    W_TOTAL = 15704
    w_num = 450
    w_avg = 700 if n_disc > 0 else 0
    w_avail = W_TOTAL - w_num - w_avg

    if n_disc == 0:
        w_name = w_avail
        col_widths = [w_num, w_name, 700]
    else:
        if n_disc <= 3:
            w_name = 3800
        elif n_disc <= 5:
            w_name = 3200
        elif n_disc <= 7:
            w_name = 2700
        elif n_disc <= 9:
            w_name = 2300
        else:
            w_name = 2000

        w_disc_total = w_avail - w_name
        n_subcols = 3 * n_disc
        w_sub_base = int(w_disc_total / n_subcols)
        rem = w_disc_total - (w_sub_base * n_subcols)

        col_widths = [w_num, w_name]
        for i in range(n_disc):
            for j in range(3):
                extra = 1 if (i * 3 + j) < rem else 0
                col_widths.append(w_sub_base + extra)
        col_widths.append(w_avg)

    # Розміри шрифтів залежно від кількості предметів
    if n_disc >= 8:
        disc_font_size = 6.5
        sub_font_size = 6.0
        data_font_size = 6.5
        name_font_size = 7.5
        hdr_font_size = 7.0
    elif n_disc >= 5:
        disc_font_size = 7.0
        sub_font_size = 6.5
        data_font_size = 7.0
        name_font_size = 8.0
        hdr_font_size = 7.5
    else:
        disc_font_size = 7.5
        sub_font_size = 7.0
        data_font_size = 7.5
        name_font_size = 8.5
        hdr_font_size = 8.0

    table = document.add_table(rows=2 + len(order), cols=cols)
    try:
        table.style = 'Table Grid'
    except Exception:
        pass

    # Властивості таблиці: фіксована ширина, центр, компактні внутрішні поля комірок
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), str(W_TOTAL))
    tblW.set(qn('w:type'), 'dxa')

    tblLayout = tblPr.find(qn('w:tblLayout'))
    if tblLayout is None:
        tblLayout = OxmlElement('w:tblLayout')
        tblPr.append(tblLayout)
    tblLayout.set(qn('w:type'), 'fixed')

    jc = tblPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        tblPr.append(jc)
    jc.set(qn('w:val'), 'center')

    tblCellMar = tblPr.find(qn('w:tblCellMar'))
    if tblCellMar is None:
        tblCellMar = OxmlElement('w:tblCellMar')
        tblPr.append(tblCellMar)
    tblCellMar.clear()
    for side, val in [('top', '30'), ('bottom', '30'), ('left', '35'), ('right', '35')]:
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'), val)
        m.set(qn('w:type'), 'dxa')
        tblCellMar.append(m)

    # Оновлення сітки колонок таблиці (w:tblGrid)
    tblGrid = table._tbl.tblGrid
    tblGrid.clear()
    for w in col_widths:
        col_elem = OxmlElement('w:gridCol')
        col_elem.set(qn('w:w'), str(w))
        tblGrid.append(col_elem)

    def set_cell(cell, text, bold=False, size=7.5, align=WD_ALIGN_PARAGRAPH.CENTER, width_dxa=None, fill_hex=None):
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        if text is not None and text != '':
            run = p.add_run(str(text))
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.bold = bold

        tcPr = cell._tc.get_or_add_tcPr()
        vAlign = tcPr.find(qn('w:vAlign'))
        if vAlign is None:
            vAlign = OxmlElement('w:vAlign')
            tcPr.append(vAlign)
        vAlign.set(qn('w:val'), 'center')

        if width_dxa:
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.append(tcW)
            tcW.set(qn('w:w'), str(width_dxa))
            tcW.set(qn('w:type'), 'dxa')

        if fill_hex:
            shd = tcPr.find(qn('w:shd'))
            if shd is None:
                shd = OxmlElement('w:shd')
                tcPr.append(shd)
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), fill_hex)

    def mark_row_header(row):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn('w:tblHeader')) is None:
            trPr.append(OxmlElement('w:tblHeader'))
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))

    def mark_row_cant_split(row):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))

    HDR_FILL = 'F2F2F2'

    # Об'єднання комірок заголовка (рядок 0 та 1)
    # 1) № (вертикально рядки 0 і 1)
    c_num = table.cell(0, 0).merge(table.cell(1, 0))
    set_cell(c_num, '№', bold=True, size=hdr_font_size, width_dxa=w_num, fill_hex=HDR_FILL)

    # 2) ПІБ (вертикально рядки 0 і 1)
    c_name = table.cell(0, 1).merge(table.cell(1, 1))
    set_cell(c_name, "Прізвище, ім'я\nпо батькові", bold=True, size=hdr_font_size, width_dxa=w_name, fill_hex=HDR_FILL)

    # 3) Дисципліни (рядок 0: об'єднання 3 колонок; рядок 1: нац, бали, ECTS)
    for i, sh in enumerate(sheets):
        base = 2 + i * 3
        c_disc = table.cell(0, base).merge(table.cell(0, base + 2))
        title = (sh.get('discipline_name') or '').strip()
        ctrl = (sh.get('control_form') or '').strip()
        disc_text = f"{title}\n({ctrl})" if ctrl else title
        w_disc_hdr = col_widths[base] + col_widths[base + 1] + col_widths[base + 2]
        set_cell(c_disc, disc_text, bold=True, size=disc_font_size, width_dxa=w_disc_hdr, fill_hex=HDR_FILL)

        set_cell(table.cell(1, base), 'нац.', bold=True, size=sub_font_size, width_dxa=col_widths[base], fill_hex=HDR_FILL)
        set_cell(table.cell(1, base + 1), 'бали', bold=True, size=sub_font_size, width_dxa=col_widths[base + 1], fill_hex=HDR_FILL)
        set_cell(table.cell(1, base + 2), 'ECTS', bold=True, size=sub_font_size, width_dxa=col_widths[base + 2], fill_hex=HDR_FILL)

    # 4) Середній бал (вертикально рядки 0 і 1)
    if n_disc > 0:
        c_avg = table.cell(0, avg_col).merge(table.cell(1, avg_col))
        set_cell(c_avg, "Сер.\nбал", bold=True, size=hdr_font_size, width_dxa=w_avg, fill_hex=HDR_FILL)

    mark_row_header(table.rows[0])
    mark_row_header(table.rows[1])

    # Заповнення рядків студентів
    for ri, sid in enumerate(order):
        row = table.rows[ri + 2]
        mark_row_cant_split(row)

        set_cell(row.cells[0], ri + 1, size=data_font_size, width_dxa=w_num)
        set_cell(row.cells[1], names.get(sid, ''), size=name_font_size, align=WD_ALIGN_PARAGRAPH.LEFT, width_dxa=w_name)

        pts_list = []
        for i, sh in enumerate(sheets):
            base = 2 + i * 3
            g = (grades.get(sid) or {}).get(sh['id']) or {}
            nat_short = _short_national_grade(g.get('national_grade'))
            pts = g.get('points')
            ects = g.get('ects') or ''

            set_cell(row.cells[base], nat_short, size=data_font_size, width_dxa=col_widths[base])
            set_cell(row.cells[base + 1], '' if pts is None else pts, size=data_font_size, width_dxa=col_widths[base + 1])
            set_cell(row.cells[base + 2], ects, size=data_font_size, width_dxa=col_widths[base + 2])

            if pts is not None and pts != '':
                try:
                    pts_list.append(float(pts))
                except (TypeError, ValueError):
                    pass

        if n_disc > 0:
            avg = round(sum(pts_list) / len(pts_list), 2) if pts_list else ''
            set_cell(row.cells[avg_col], avg, bold=True, size=data_font_size, width_dxa=w_avg)

    # ── Переміщення таблиці на місце маркера ──
    if marker is not None:
        try:
            marker._p.addnext(table._tbl)
        except Exception:
            pass
        try:
            for r in list(marker.runs):
                r._r.getparent().remove(r._r)
        except Exception:
            pass
        try:
            marker._p.getparent().remove(marker._p)
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════
    #  ФІНАЛЬНА НОРМАЛІЗАЦІЯ <w:body> ПІД ВИМОГИ WORD
    # ═══════════════════════════════════════════════════════════════
    body = document._element.body

    # 1) після КОЖНОЇ <w:tbl> має бути <w:p>
    for child in list(body):
        if child.tag == qn('w:tbl'):
            nxt = child.getnext()
            if nxt is None or nxt.tag != qn('w:p'):
                child.addnext(OxmlElement('w:p'))

    # 2) <w:sectPr> має бути РІВНО ОДИН і ОСТАННІЙ елемент body.
    #    Прибираємо будь-які дублікати sectPr у body.
    all_sectPrs = body.findall(qn('w:sectPr'))
    if len(all_sectPrs) > 1:
        global_logger.warning(
            f"[vidomist_summary] Знайдено {len(all_sectPrs)} <w:sectPr> у body — "
            f"залишаю тільки останній"
        )
        for extra in all_sectPrs[:-1]:
            body.remove(extra)
        all_sectPrs = [all_sectPrs[-1]]

    # 3) Якщо sectPr є — переміщуємо його в САМИЙ КІНЕЦЬ body
    #    (Word жорстко вимагає, щоб sectPr був останнім).
    if all_sectPrs:
        sectPr = all_sectPrs[0]
        if body[-1] is not sectPr:
            body.remove(sectPr)
            body.append(sectPr)
            global_logger.debug(
                "[vidomist_summary] <w:sectPr> переміщено в кінець body"
            )

def gen_vidomist_summary(session_meta, sheets, by_sheet_students, template=None, out=None, user_name='Система'):
    """Загальна відомість сесії: DocxTemplate.render + вставка таблиці."""
    from io import BytesIO
    from docx.oxml.ns import qn

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
            pts = _g(st, 'points')
            ctrl = sh.get('control_form')
            nat = _g(st, 'national_grade') or _vidomist_national(pts, ctrl)
            ects = _g(st, 'ects') or _vidomist_ects(pts)
            grades[sid_key][sh['id']] = {
                'national_grade': nat or '',
                'points': pts,
                'ects': ects or '',
            }

    try:
        from routes.helpers import sort_ukrainian
        order = sort_ukrainian(order, lambda sid: names.get(sid, ''))
    except Exception:
        pass

    from routes.utils import resolve_vidomist_faculty
    _conn_sum = None
    try:
        sp_sum_ref = (session_meta.get('specialty_code') or session_meta.get('specialty') or '').strip()
        if not (session_meta.get('faculty') or '').strip() and sp_sum_ref:
            from routes.db import get_db
            _conn_sum = get_db()
        sum_faculty, sum_faculty_en, sum_ka = resolve_vidomist_faculty(session_meta, conn=_conn_sum)
    except Exception:
        sum_ka = (session_meta.get('knowledge_area') or '').strip()
        sum_faculty = (session_meta.get('faculty') or '').strip() or sum_ka or ''
        sum_faculty_en = (session_meta.get('faculty_en') or '').strip()
    finally:
        if _conn_sum is not None:
            try:
                _conn_sum.close()
            except Exception:
                pass

    form = session_meta.get('study_form') or ''
    ctx = {
        'institution': session_meta.get('institution_name_and_status')
            or 'ПВНЗ „Європейський університет“',
        'faculty': sum_faculty,
        'faculty_en': sum_faculty_en,
        'knowledge_area': sum_ka,
        'group_name': session_meta.get('group_name') or '',
        'course': _vidomist_roman(session_meta.get('course')),
        'study_form': form,
        'academic_year': session_meta.get('academic_year') or '',
        'semester': str(session_meta.get('semester') or ''),
        'degree_level': session_meta.get('degree_level') or '',
        'specialty': _vidomist_format_specialty(
            session_meta.get('specialty_code'), session_meta.get('specialty')),
        'session_name': session_meta.get('name') or session_meta.get('session_name') or '',
        # 'grades_table' НАВМИСНО відсутній: замість Jinja-плейсхолдера
        # {{ grades_table }} вписуйте в шаблон звичайний текст
        # <<<GRADES_TABLE>>> — код його знайде сам і замінить
        # таблицею. Через {{ }} docxtpl розбиває маркер на кілька runs
        # і ламає XML, який Word не пробачає (а LibreOffice прощає).
    }

    if template is None:
        template = _vidomist_summary_template_path()
    if not os.path.exists(template):
        global_logger.error(f"Шаблон загальної відомості не знайдено: {template}")
        raise FileNotFoundError(f"Шаблон загальної відомості не знайдено: {template}")

    # ── БЕЗ roundtrip через tpl.save() → DocxDocument() ──
    # DocxTemplate успадкований від Document, тому працюємо з tpl
    # напряму. Раніше робили дві серіалізації поспіль — Word часто
    # цього не прощає (LibreOffice — прощає).
    try:
        document = DocxTemplate(template)
        document.render(ctx)
    except Exception as e:
        global_logger.error(f"Помилка render загальної відомості: {e}")
        raise

    try:
        _vidomist_insert_summary_table(document, sheets, order, names, grades)
    except Exception as e:
        global_logger.error(
            f"Загальна відомість: НЕ вдалося вставити таблицю: {e}",
            exc_info=True,
        )
        raise

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

    # ─── Діагностика: структура + zip/XML ───
    body = document._element.body
    sectPrs = body.findall(qn('w:sectPr'))
    last_tag = body[-1].tag.split('}')[-1] if len(body) else '—'
    global_logger.info(
        f"[vidomist_summary] sections={len(document.sections)}, "
        f"paragraphs={len(document.paragraphs)}, "
        f"tables={len(document.tables)}, "
        f"sectPr_count={len(sectPrs)}, "
        f"last_body_child={last_tag}, "
        f"sheets={len(sheets)}, students={len(order)}"
    )

    import zipfile
    import xml.etree.ElementTree as ET
    try:
        with zipfile.ZipFile(buf) as zf:
            names_in_zip = zf.namelist()
            required = ['[Content_Types].xml', 'word/document.xml',
                        'word/_rels/document.xml.rels']
            missing = [r for r in required if r not in names_in_zip]
            ET.fromstring(zf.read('word/document.xml'))
            global_logger.info(
                f"[vidomist_summary] zip OK: {len(names_in_zip)} entries, "
                f"missing={missing or 'немає'}"
            )
    except Exception as e:
        global_logger.error(
            f"[vidomist_summary] ФАЙЛ НЕВАЛІДНИЙ: {e}", exc_info=True
        )

    buf.seek(0)
    return buf
