"""Template-preserving Excel export for session statements."""
from __future__ import annotations

import os
import re
from copy import copy
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

from routes.security import sanitize_spreadsheet_cell

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
TEMPLATE_DIR = os.path.join(ROOT, 'static', 'templates', 'excel')


def _get(row, key, default=None):
    if row is None:
        return default
    try:
        value = row.get(key, default) if isinstance(row, dict) else row[key]
    except (KeyError, IndexError, TypeError):
        return default
    return default if value is None else value


def _is_day_form(study_form):
    form = (study_form or '').strip().casefold()
    return form in {'денна', 'очна', 'denna', 'full-time', 'full time', 'full'}


def _template_path(study_form):
    filename = 'vidomist_template_denna.xlsx' if _is_day_form(study_form) else 'vidomist_template_zaochna.xlsx'
    path = os.path.join(TEMPLATE_DIR, filename)
    if not os.path.isfile(path):
        raise FileNotFoundError(f'Не знайдено шаблон Excel: {path}')
    return path


def _find_table(ws):
    """Find header/data rows using labels in the supplied workbook."""
    for r in range(1, min(ws.max_row, 50) + 1):
        labels = [str(ws.cell(r, c).value or '').strip().casefold() for c in range(1, ws.max_column + 1)]
        row_text = ' '.join(labels)
        if ('прізвище' in row_text or 'піб' in row_text or 'студент' in row_text) and ('№' in row_text or 'з/п' in row_text or 'з\\п' in row_text):
            # Templates have a second tier of headers; data starts after it.
            next_text = ' '.join(str(ws.cell(r + 1, c).value or '').casefold() for c in range(1, ws.max_column + 1))
            header_rows = 2 if any(term in next_text for term in ('бал', 'ects', 'поточ', 'іспит', 'нац', 'підсум')) else 1
            return r, r + header_rows, header_rows
    raise ValueError('У шаблоні не знайдено заголовок таблиці студентів')


def _column_roles(ws, header_row, header_rows, is_day):
    labels = []
    for c in range(1, ws.max_column + 1):
        labels.append(' '.join(str(ws.cell(r, c).value or '').casefold() for r in range(header_row, header_row + header_rows)))

    def find(*terms):
        return next((i + 1 for i, label in enumerate(labels) if any(t in label for t in terms)), None)

    roles = {'n': find('№', 'з/п', 'з\\п') or 1,
             'name': find('прізвище', 'піб', 'студент') or 2,
             'edebo': find('індивідуаль', 'навчального плану', 'едебо')}
    if is_day:
        roles.update(total=find('100 баль', 'кількість балів', 'підсумков', 'бали'),
                     national=find('національ', 'нац.'), ects=find('ects'), sign=find('підпис'))
    else:
        roles.update(mid=find('міжсес'), current=find('поточ'), exam=find('іспит', 'екзамен', 'залік'),
                     total=find('підсумков', '100 баль', 'кількість балів'),
                     national=find('національ', 'нац.'), ects=find('ects'), sign=find('підпис'))

    # Some templates use merged headings; fallback column positions match their
    # respective printed layouts and don't touch any template styling.
    defaults = ({'national': 4, 'total': 5, 'ects': 6, 'sign': 7} if is_day else
                {'mid': 4, 'current': 5, 'exam': 6, 'total': 7, 'national': 8, 'ects': 9, 'sign': 10})
    for role, col in defaults.items():
        roles.setdefault(role, col)
    roles.setdefault('edebo', 3)
    return roles


def _set_header_values(ws, sheet):
    """Update values only; retain template merges, dimensions, and cell styles."""
    course_map = {1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI'}
    try:
        course = course_map.get(int(_get(sheet, 'course')), str(_get(sheet, 'course') or ''))
    except (TypeError, ValueError):
        course = str(_get(sheet, 'course') or '')
    exam_date = str(_get(sheet, 'exam_date') or '')
    match = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', exam_date)
    if match:
        months = ('', 'січня', 'лютого', 'березня', 'квітня', 'травня', 'червня', 'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня')
        y, m, d = map(int, match.groups())
        if 1 <= m <= 12:
            exam_date = f'{d} {months[m]} {y}'
    semester_names = {1: 'перший', 2: 'другий', 3: 'третій'}
    try:
        semester = semester_names.get(int(_get(sheet, 'semester')), str(_get(sheet, 'semester') or ''))
    except (TypeError, ValueError):
        semester = str(_get(sheet, 'semester') or '')
    values = {
        'A1': _get(sheet, 'institution_name_and_status') or 'ПВНЗ «Європейський університет»\nЛьвівська філія',
        'A3': f"Інститут, факультет, відділення      {_get(sheet, 'faculty') or _get(sheet, 'department') or 'Інформаційні системи та технології'}",
        'A4': f"Спеціальність           {_get(sheet, 'specialty_code') or ''} {_get(sheet, 'specialty') or ''}".strip(),
        'A5': f"Курс    {course}   Група       {_get(sheet, 'group_name') or ''}     ({_get(sheet, 'study_form') or ''})   {_get(sheet, 'academic_year') or ''}   навчальний рік",
        'A6': f"Освітній рівень  {_get(sheet, 'degree_level') or ''}",
        'A7': f"ВІДОМІСТЬ  № {_get(sheet, 'statement_number')}." if _get(sheet, 'statement_number') else 'ВІДОМІСТЬ',
        'A8': f'Дата проведення                                       {exam_date}',
        'A10': f"Назва навчальної дисципліни           {_get(sheet, 'discipline_name') or ''}",
        'A12': f'за   {semester}   навчальний семестр.',
        'A13': f"Форма семестрового контролю  {(_get(sheet, 'control_form') or '').lower()}     Загальна кількість годин   {_get(sheet, 'hours_credits') or ''}",
        'A15': f"Викладач                                                   {_get(sheet, 'teacher_short') or _get(sheet, 'teacher_name') or _get(sheet, 'teacher_full') or ''}",
    }
    for cell, value in values.items():
        ws[cell] = sanitize_spreadsheet_cell(value)


def _copy_row(ws, source, target):
    for col in range(1, ws.max_column + 1):
        src, dst = ws.cell(source, col), ws.cell(target, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
    ws.row_dimensions[target].height = ws.row_dimensions[source].height


def _resize_student_area(ws, data_start, footer_row, required_count, sample_row):
    """Resize only the student area; retain/reposition all footer merges and styles."""
    capacity = max(footer_row - data_start, 0)
    delta = required_count - capacity
    if delta == 0:
        return footer_row

    merged = [str(rng) for rng in ws.merged_cells.ranges]
    for rng in merged:
        ws.unmerge_cells(rng)
    if delta > 0:
        # Translate formulas below the insertion point before inserting rows.
        for row in ws.iter_rows(min_row=footer_row, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.value = Translator(cell.value, origin=cell.coordinate).translate_formula(cell.coordinate)
        ws.insert_rows(footer_row, delta)
        for r in range(footer_row, footer_row + delta):
            _copy_row(ws, sample_row, r)
        shifted = []
        for rng in merged:
            from openpyxl.worksheet.cell_range import CellRange
            cr = CellRange(rng)
            if cr.min_row >= footer_row:
                cr.shift(row_shift=delta)
            elif cr.max_row >= footer_row:
                cr.max_row += delta
            shifted.append(str(cr))
    else:
        remove_start = data_start + required_count
        remove_count = -delta
        ws.delete_rows(remove_start, remove_count)
        shifted = []
        from openpyxl.worksheet.cell_range import CellRange
        for rng in merged:
            cr = CellRange(rng)
            if cr.min_row >= remove_start + remove_count:
                cr.shift(row_shift=delta)
            elif cr.max_row >= remove_start:
                cr.max_row = remove_start - 1
            if cr.max_row >= cr.min_row:
                shifted.append(str(cr))
    for rng in shifted:
        ws.merge_cells(rng)
    return footer_row + delta


def _grade_values(points, control):
    if points is None or points == '':
        return '', ''
    try:
        p = int(points)
    except (TypeError, ValueError):
        return '', ''
    if control == 'Залік':
        national = 'Зараховано' if p >= 60 else 'Незараховано'
    elif p >= 90:
        national = 'Відмінно / Excellent'
    elif p >= 74:
        national = 'Добре/ Good'
    elif p >= 60:
        national = 'Задовільно/ Satisfactory'
    else:
        national = 'Незадовільно/ Fail'
    ects = 'A' if p >= 90 else 'B' if p >= 82 else 'C' if p >= 74 else 'D' if p >= 64 else 'E' if p >= 60 else 'FX' if p >= 35 else 'F'
    return national, ects


def build_statement_xlsx(sheet, students) -> BytesIO:
    sheet = dict(sheet) if sheet is not None and not isinstance(sheet, dict) else (sheet or {})
    students = [dict(s) if not isinstance(s, dict) else s for s in (students or [])]
    template = _template_path(_get(sheet, 'study_form'))
    wb = load_workbook(template)
    ws = wb.active
    header_row, data_start, header_rows = _find_table(ws)
    is_day = _is_day_form(_get(sheet, 'study_form'))
    roles = _column_roles(ws, header_row, header_rows, is_day)

    # Locate the existing signature/statistics footer, preserving its cells and
    # formatting. The footer follows the reserved blank student rows in template.
    footer_row = None
    for r in range(data_start, ws.max_row + 1):
        text = ' '.join(str(ws.cell(r, c).value or '').casefold() for c in range(1, ws.max_column + 1))
        if 'декан' in text or 'результати оцінювання' in text:
            footer_row = r
            break
    if footer_row is None:
        footer_row = max(data_start + 1, ws.max_row - 14)

    sample_row = data_start
    _resize_student_area(ws, data_start, footer_row, max(len(students), 1), sample_row)
    _set_header_values(ws, sheet)

    teacher = _get(sheet, 'teacher_full') or _get(sheet, 'teacher_short') or _get(sheet, 'teacher_name') or ''
    try:
        from routes.gen_docx import _format_teacher_short
        teacher_sign = _format_teacher_short(teacher)
    except Exception:
        teacher_sign = teacher
    control = _get(sheet, 'control_form') or 'Екзамен'

    for idx, student in enumerate(students):
        row = data_start + idx
        if row != sample_row:
            _copy_row(ws, sample_row, row)
        points = _get(student, 'points')
        nat, ects = _grade_values(points, control)
        values = {
            'n': idx + 1,
            'name': sanitize_spreadsheet_cell(_get(student, 'full_name') or ''),
            'edebo': sanitize_spreadsheet_cell(_get(student, 'edebo_code') or ''),
            'mid': _get(student, 'points_mid') if not is_day else '',
            'current': _get(student, 'points_current') if not is_day else '',
            'exam': _get(student, 'points_exam') if not is_day else '',
            'total': points if points is not None else '',
            'national': sanitize_spreadsheet_cell(_get(student, 'national_grade') or nat),
            'ects': sanitize_spreadsheet_cell(_get(student, 'ects') or ects),
            'sign': sanitize_spreadsheet_cell(_get(student, 'teacher_sign') or teacher_sign),
        }
        for role, value in values.items():
            col = roles.get(role)
            if col:
                ws.cell(row, col).value = value

    if not students:
        ws.cell(data_start, roles['n']).value = 1
        for role in roles:
            if role != 'n':
                ws.cell(data_start, roles[role]).value = ''

    # Keep workbook sizing/print settings from the chosen template untouched.
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output