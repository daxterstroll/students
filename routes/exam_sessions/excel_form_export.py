"""Excel export for session statements using form-specific workbook templates."""
from __future__ import annotations

import os
import re
import zipfile
from copy import copy
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
_TEMPLATE_DIR = os.path.join(_ROOT, 'static', 'templates', 'excel')
_DAY_TEMPLATE = 'vidomist_template_denna.xlsx'
_PART_TIME_TEMPLATE = 'vidomist_template_zaochna.xlsx'
_FALLBACK_TEMPLATE = 'vidomist_template.xlsx'
_FONT_NAME = 'Times New Roman'


def _get(row, key, default=None):
    if row is None:
        return default
    if isinstance(row, dict):
        value = row.get(key, default)
        return default if value is None else value
    try:
        value = row[key]
        return default if value is None else value
    except (KeyError, IndexError, TypeError):
        return default


def _is_fulltime(study_form):
    form = (study_form or '').strip().casefold()
    return form in {'денна', 'очна', 'denna', 'full-time', 'full time', 'full'}


def _template_for(study_form):
    filename = _DAY_TEMPLATE if _is_fulltime(study_form) else _PART_TIME_TEMPLATE
    path = os.path.join(_TEMPLATE_DIR, filename)
    if os.path.isfile(path):
        return path
    fallback = os.path.join(_TEMPLATE_DIR, _FALLBACK_TEMPLATE)
    if os.path.isfile(fallback):
        return fallback
    raise FileNotFoundError(f'Не знайдено Excel-шаблон відомості: {path}')


def _fmt_exam_date(value):
    if not value:
        return ''
    text = str(value).strip()
    match = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', text)
    if not match:
        return text
    year, month, day = map(int, match.groups())
    months = ('', 'січня', 'лютого', 'березня', 'квітня', 'травня', 'червня',
              'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня')
    return f'{day} {months[month]} {year}' if 1 <= month <= 12 else text


def _roman_course(course):
    values = {1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI'}
    try:
        return values.get(int(course), str(course or ''))
    except (TypeError, ValueError):
        return str(course or '')


def _header_rows_and_roles(ws, fulltime):
    """Find the existing table header and map columns by their labels."""
    header_row = None
    header_text = []
    for row_idx in range(1, min(ws.max_row, 40) + 1):
        values = [str(ws.cell(row_idx, col).value or '').strip().casefold()
                  for col in range(1, ws.max_column + 1)]
        joined = ' '.join(values)
        if ('прізвище' in joined or 'піб' in joined or 'студент' in joined) and (
            '№' in joined or 'з/п' in joined or 'з\\п' in joined or 'з\п' in joined
        ):
            header_row = row_idx
            header_text = values
            break
    if header_row is None:
        # Both supplied templates have the standard two-row header in this area.
        header_row = 17
        header_text = [str(ws.cell(header_row, c).value or '').strip().casefold()
                       for c in range(1, ws.max_column + 1)]

    second_text = [str(ws.cell(header_row + 1, c).value or '').strip().casefold()
                   for c in range(1, ws.max_column + 1)]
    combined = [(header_text[i] + ' ' + second_text[i]).strip()
                for i in range(max(len(header_text), len(second_text)))]

    def find(*terms):
        for idx, label in enumerate(combined, start=1):
            if any(term in label for term in terms):
                return idx
        return None

    roles = {
        'n': find('№', 'з/п', 'з\\п', 'з\п') or 1,
        'name': find('прізвище', 'піб', 'по батькові', 'студент') or 2,
        'edebo': find('індивідуаль', 'навчального плану', 'едебо'),
    }
    if fulltime:
        roles.update({
            'mid': None, 'current': None, 'exam': None,
            'total': find('100 баль', 'кількість балів', 'підсумков', 'бали'),
            'national': find('національ', 'нац.'),
            'ects': find('ects'),
            'sign': find('підпис'),
        })
    else:
        roles.update({
            'mid': find('міжсес'),
            'current': find('поточ'),
            'exam': find('іспит', 'екзамен', 'залік'),
            'total': find('100 баль', 'кількість балів', 'підсумков', 'бали'),
            'national': find('національ', 'нац.'),
            'ects': find('ects'),
            'sign': find('підпис'),
        })
    # Header text in merged cells can be empty in secondary columns. Use the
    # established layouts as a safe fallback, while keeping header detection.
    if roles['name'] == 2 and roles['edebo'] is None:
        roles['edebo'] = 3
    if fulltime:
        roles['national'] = roles['national'] or 4
        roles['total'] = roles['total'] or 5
        roles['ects'] = roles['ects'] or 6
        roles['sign'] = roles['sign'] or 7
    else:
        roles['mid'] = roles['mid'] or 4
        roles['current'] = roles['current'] or 5
        roles['exam'] = roles['exam'] or 6
        roles['total'] = roles['total'] or 7
        roles['national'] = roles['national'] or 8
        roles['ects'] = roles['ects'] or 9
        roles['sign'] = roles['sign'] or 10
    roles = {key: value for key, value in roles.items() if value is not None}
    return header_row, header_row + 2, roles


def _set_metadata(ws, sheet, col_count):
    """Fill metadata in the template's existing header cells and merge layout."""
    course = _roman_course(_get(sheet, 'course'))
    group = _get(sheet, 'group_name') or ''
    form = _get(sheet, 'study_form') or ''
    year = _get(sheet, 'academic_year') or ''
    institution = _get(sheet, 'institution_name_and_status') or 'ПВНЗ «Європейський університет»\nЛьвівська філія'
    values = {
        1: institution,
        3: f"Інститут, факультет, відділення      {_get(sheet, 'faculty') or _get(sheet, 'department') or 'Інформаційні системи та технології'}",
        4: f"Спеціальність           {_get(sheet, 'specialty_code') or ''} {_get(sheet, 'specialty') or ''}".strip(),
        5: f"Курс    {course}   Група       {group}     ({form})   {year}   навчальний рік",
        6: f"Освітній рівень  {_get(sheet, 'degree_level') or ''}",
        7: f"ВІДОМІСТЬ  № {_get(sheet, 'statement_number')}." if _get(sheet, 'statement_number') else 'ВІДОМІСТЬ',
        8: f"Дата проведення                                       {_fmt_exam_date(_get(sheet, 'exam_date'))}",
        10: f"Назва навчальної дисципліни           {_get(sheet, 'discipline_name') or ''}",
        12: f"за   {_get(sheet, 'semester') or ''}   навчальний семестр.",
        13: f"Форма семестрового контролю  {(_get(sheet, 'control_form') or '').lower()}     Загальна кількість годин   {_get(sheet, 'hours_credits') or ''}",
        15: f"Викладач                                                   {_get(sheet, 'teacher_short') or _get(sheet, 'teacher_name') or _get(sheet, 'teacher_full') or ''}",
    }
    for row, value in values.items():
        cell = ws.cell(row=row, column=1)
        cell.value = value
        cell.font = copy(cell.font)
        cell.font = Font(name=_FONT_NAME, size=cell.font.sz or 11,
                         bold=cell.font.bold, italic=cell.font.italic)


def _copy_row_style(ws, source_row, target_row, col_count):
    if source_row == target_row:
        return
    for col in range(1, col_count + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.protection:
            target.protection = copy(source.protection)
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def _format_grade(points, control_form):
    if points is None or points == '':
        return ''
    try:
        value = int(points)
    except (TypeError, ValueError):
        return ''
    if control_form == 'Залік':
        return 'Зараховано' if value >= 60 else 'Незараховано'
    if value >= 90:
        return 'Відмінно / Excellent'
    if value >= 74:
        return 'Добре / Good'
    if value >= 60:
        return 'Задовільно / Satisfactory'
    return 'Незадовільно / Fail'


def _ects(points):
    if points is None or points == '':
        return ''
    try:
        points = int(points)
    except (TypeError, ValueError):
        return ''
    return 'A' if points >= 90 else 'B' if points >= 82 else 'C' if points >= 74 else 'D' if points >= 64 else 'E' if points >= 60 else 'FX' if points >= 35 else 'F'


def _write_footer(ws, first_student, last_student, n_students, roles, col_count, sample_row):
    """Replace the template's old footer with a clean footer after the data."""
    from openpyxl.styles import Alignment, Font

    footer_start = last_student + 3
    bold = Font(name=_FONT_NAME, size=12, bold=True)
    normal = Font(name=_FONT_NAME, size=11)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ects_col = roles.get('ects')
    points_col = roles.get('total')
    grades_range = (f'${get_column_letter(ects_col)}${first_student}:'
                    f'${get_column_letter(ects_col)}${last_student}') if ects_col else None
    points_range = (f'${get_column_letter(points_col)}${first_student}:'
                    f'${get_column_letter(points_col)}${last_student}') if points_col else None

    _copy_row_style(ws, sample_row, footer_start, col_count)
    ws.cell(footer_start, min(4, col_count)).value = 'Декан факультету_____________________'
    ws.cell(footer_start, min(4, col_count)).font = bold
    _copy_row_style(ws, sample_row, footer_start + 2, col_count)
    ws.cell(footer_start + 2, min(4, col_count)).value = 'Викладач____________________________'
    ws.cell(footer_start + 2, min(4, col_count)).font = bold

    stats_row = footer_start + 4
    _copy_row_style(ws, sample_row, stats_row, col_count)
    for col, value in ((1, 'Результати оцінювання'), (3, 'Кількість'), (4, '%'),
                       (5, 'Якість успішності,%'), (6, 'Абсолютна успішність,%')):
        if col <= col_count:
            cell = ws.cell(stats_row, col)
            cell.value = value
            cell.font = bold
            cell.alignment = center

    bands = [('А (90-100)', 'A'), ('В (82-89)', 'B'), ('С (74-81)', 'C'),
             ('D (64-73)', 'D'), ('E (60-63)', 'E'), ('FX (35-59)', 'FX'),
             ('F (1-34)', 'F'), ("Не з'явився", None)]
    denom = max(n_students, 1)
    first_band = stats_row + 1
    for offset, (label, letter) in enumerate(bands):
        row = first_band + offset
        _copy_row_style(ws, sample_row, row, col_count)
        ws.cell(row, 1).value = label
        ws.cell(row, 1).font = bold
        if col_count >= 3 and grades_range:
            formula = (f'=COUNTIF({grades_range},"{letter}")' if letter else
                       f'=COUNTBLANK({points_range})' if points_range else '=0')
            ws.cell(row, 3).value = formula
        if col_count >= 4:
            ws.cell(row, 4).value = f'=C{row}/{denom}'
            ws.cell(row, 4).number_format = '0%'
        if offset == 0 and col_count >= 5:
            ws.cell(row, 5).value = f'=D{first_band}+D{first_band+1}+D{first_band+2}'
            ws.cell(row, 5).number_format = '0%'
        if offset == 0 and col_count >= 6:
            ws.cell(row, 6).value = f'=D{first_band}+D{first_band+1}+D{first_band+2}+D{first_band+3}+D{first_band+4}'
            ws.cell(row, 6).number_format = '0%'


def build_statement_xlsx(sheet, students) -> BytesIO:
    """Generate a study-form-specific statement from the matching XLSX template."""
    sheet = dict(sheet) if sheet is not None and not isinstance(sheet, dict) else (sheet or {})
    students = [dict(student) if not isinstance(student, dict) else student
                for student in (students or [])]
    fulltime = _is_fulltime(_get(sheet, 'study_form'))
    template_path = _template_for(_get(sheet, 'study_form'))
    workbook = load_workbook(template_path)
    ws = workbook.active

    header_row, data_start, roles = _header_rows_and_roles(ws, fulltime)
    col_count = max(ws.max_column, max(roles.values(), default=1))
    sample_row = data_start

    # Remove stale template content and merged regions below the table header.
    # Keeping old footer cells here caused Excel's repair warning and overlapping
    # signature/statistics blocks when the student count changed.
    for merged in list(ws.merged_cells.ranges):
        if merged.max_row >= data_start:
            ws.unmerge_cells(str(merged))
    old_max_row = ws.max_row
    for row in ws.iter_rows(min_row=data_start, max_row=old_max_row, max_col=col_count):
        for cell in row:
            cell.value = None

    _set_metadata(ws, sheet, col_count)
    control = _get(sheet, 'control_form') or 'Екзамен'
    teacher_name = (_get(sheet, 'teacher_full') or _get(sheet, 'teacher_short')
                    or _get(sheet, 'teacher_name') or '')
    try:
        from routes.gen_docx import _format_teacher_short
        teacher_sign = _format_teacher_short(teacher_name)
    except Exception:
        teacher_sign = teacher_name

    for idx, student in enumerate(students):
        row = data_start + idx
        _copy_row_style(ws, sample_row, row, col_count)
        points = _get(student, 'points')
        values = {
            'n': idx + 1,
            'name': _get(student, 'full_name') or '',
            'edebo': _get(student, 'edebo_code') or '',
            'mid': _get(student, 'points_mid') if not fulltime else '',
            'current': _get(student, 'points_current') if not fulltime else '',
            'exam': _get(student, 'points_exam') if not fulltime else '',
            'total': points if points is not None and points != '' else '',
            'national': _get(student, 'national_grade') or _format_grade(points, control),
            'ects': _get(student, 'ects') or _ects(points),
            'sign': _get(student, 'teacher_sign') or teacher_sign,
        }
        for role, value in values.items():
            col = roles.get(role)
            if col and col <= col_count:
                cell = ws.cell(row, col)
                cell.value = value
                cell.font = copy(cell.font)
                cell.font = Font(name=_FONT_NAME, size=cell.font.sz or 11,
                                 bold=cell.font.bold, italic=cell.font.italic)
                if not cell.alignment:
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        if ws.row_dimensions[row].height is None:
            ws.row_dimensions[row].height = 16.5

    last_student = data_start + max(len(students), 1) - 1
    if not students:
        _copy_row_style(ws, sample_row, data_start, col_count)
        ws.cell(data_start, roles.get('n', 1)).value = 1
    _write_footer(ws, data_start, last_student, len(students), roles, col_count, sample_row)

    safe_title = str(_get(sheet, 'discipline_name') or 'Відомість')[:31]
    for char in '[]:*?/\\':
        safe_title = safe_title.replace(char, '_')
    ws.title = safe_title or 'Відомість'

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    # Verify both the XLSX package and workbook XML are readable before serving.
    with zipfile.ZipFile(output) as archive:
        if archive.testzip() is not None:
            raise ValueError('Згенерований Excel-файл містить пошкоджені дані')
    output.seek(0)
    checked = load_workbook(output, read_only=True, data_only=False)
    checked.close()
    output.seek(0)
    return output
