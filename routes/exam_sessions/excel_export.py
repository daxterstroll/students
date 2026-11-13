"""
Експорт відомості в Excel за фізичним шаблоном
static/templates/excel/vidomist_template.xlsx

Якщо шаблон відсутній — будує аркуш з нуля в тому ж стилі (Times New Roman).
"""
from __future__ import annotations

import os
from copy import copy
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

_TEMPLATE_CANDIDATES = [
    os.path.join('static', 'templates', 'excel', 'vidomist_template.xlsx'),
    os.path.join(os.path.dirname(__file__), '..', '..', 'static', 'templates', 'excel', 'vidomist_template.xlsx'),
]

_FONT = 'Times New Roman'
_thin = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin'),
)


def _find_template():
    for p in _TEMPLATE_CANDIDATES:
        ap = os.path.abspath(p)
        if os.path.isfile(ap):
            return ap
    return None


def _get(row, key, default=None):
    if row is None:
        return default
    if isinstance(row, dict):
        return row.get(key, default)
    try:
        if key in row.keys():
            val = row[key]
            return default if val is None else val
    except Exception:
        pass
    try:
        return row[key]
    except Exception:
        return default


def _roman_course(course) -> str:
    mapping = {1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI'}
    try:
        return mapping.get(int(course), str(course or ''))
    except (TypeError, ValueError):
        return str(course or '')


def _semester_word(sem) -> str:
    mapping = {1: 'перший', 2: 'другий', 3: 'третій'}
    try:
        return mapping.get(int(sem), str(sem or ''))
    except (TypeError, ValueError):
        return str(sem or '')


def _ects_formula(row: int) -> str:
    e = f'E{row}'
    return (
        f'=IF({e}="","",IF({e}<35,"F",IF({e}<60,"FX",'
        f'IF({e}<=63,"E",IF({e}<=73,"D",IF({e}<=81,"C",'
        f'IF({e}<=89,"B","A")))))))'
    )


def _national_from_points(points, control_form: str):
    if points is None or points == '':
        return ''
    try:
        p = int(points)
    except (TypeError, ValueError):
        return ''
    if control_form == 'Залік':
        return 'Зараховано' if p >= 60 else 'Незараховано'
    if p >= 90:
        return 'Відмінно / Excellent'
    if p >= 82:
        return 'Добре/ Good'
    if p >= 74:
        return 'Добре/ Good'
    if p >= 64:
        return 'Задовільно/ Satisfactory'
    if p >= 60:
        return 'Задовільно/ Satisfactory'
    if p >= 35:
        return 'Незадовільно/ Fail'
    return 'Незадовільно/ Fail'


def _unmerge_from_row(ws, from_row: int):
    """Зняти всі merge, що торкаються рядків >= from_row (інакше openpyxl падає)."""
    to_remove = []
    for mr in list(ws.merged_cells.ranges):
        if mr.max_row >= from_row:
            to_remove.append(str(mr))
    for coord in to_remove:
        try:
            ws.unmerge_cells(coord)
        except Exception:
            pass


def _set(ws, row, col, value, font=None, align=None, border=None, keep_style=True):
    cell = ws.cell(row=row, column=col)
    prev_font = cell.font
    prev_align = cell.alignment
    prev_border = cell.border
    cell.value = value
    if font is not None:
        cell.font = font
    elif keep_style and prev_font is not None:
        cell.font = prev_font
    if align is not None:
        cell.alignment = align
    elif keep_style and prev_align is not None:
        cell.alignment = prev_align
    if border is not None:
        cell.border = border
    elif keep_style and prev_border is not None:
        cell.border = prev_border
    return cell


def _fill_header(ws, sheet: dict):
    f_title = Font(name=_FONT, size=14, bold=True)
    f13 = Font(name=_FONT, size=13)
    f13b = Font(name=_FONT, size=13, bold=True)
    f14b = Font(name=_FONT, size=14, bold=True)
    f16b = Font(name=_FONT, size=16, bold=True)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    inst = _get(sheet, 'institution_name_and_status') or \
        'ПВНЗ „Європейський університет“                                                 Львівська філія'
    _set(ws, 1, 1, inst, f_title, center)
    try:
        ws.merge_cells('A1:G1')
    except Exception:
        pass

    faculty = _get(sheet, 'faculty') or _get(sheet, 'department') or 'Інформаційні системи та технології'
    _set(ws, 3, 1, f'Інститут, факультет, відділення      {faculty}', f13)
    try:
        ws.merge_cells('A3:G3')
    except Exception:
        pass

    spec_code = _get(sheet, 'specialty_code') or ''
    spec_name = _get(sheet, 'specialty') or ''
    _set(ws, 4, 1, f'Спеціальність           {spec_code} {spec_name}'.strip(), f13b)
    try:
        ws.merge_cells('A4:G4')
    except Exception:
        pass

    course = _roman_course(_get(sheet, 'course'))
    group = _get(sheet, 'group_name') or ''
    form = _get(sheet, 'study_form') or ''
    year = _get(sheet, 'academic_year') or ''
    form_part = f'({form})' if form else ''
    _set(
        ws, 5, 1,
        f'Курс    {course}   Група       {group}     {form_part}                {year}   навчальний рік',
        f13,
    )
    try:
        ws.merge_cells('A5:G5')
    except Exception:
        pass

    degree = _get(sheet, 'degree_level') or ''
    _set(ws, 6, 1, f'Освітній рівень  {degree}', f13)
    try:
        ws.merge_cells('A6:G6')
    except Exception:
        pass

    num = _get(sheet, 'statement_number') or ''
    title = f'ВІДОМІСТЬ  № {num}.'.strip() if num else 'ВІДОМІСТЬ'
    _set(ws, 7, 1, title, f16b, center)
    try:
        ws.merge_cells('A7:G7')
    except Exception:
        pass

    exam_date = _get(sheet, 'exam_date') or ''
    exam_date = _fmt_exam_date_ua(dict(sheet).get('exam_date') if not isinstance(sheet, dict) else sheet.get('exam_date'))
    _set(ws, 8, 1, f'Дата проведення                                       {exam_date}', f14b)
    try:
        ws.merge_cells('A8:G8')
    except Exception:
        pass

    disc = _get(sheet, 'discipline_name') or ''
    _set(ws, 10, 1, f'Назва навчальної диспипліни           {disc}', Font(name=_FONT, size=14))
    try:
        ws.merge_cells('A10:G10')
    except Exception:
        pass

    sem_word = _semester_word(_get(sheet, 'semester'))
    _set(ws, 12, 1, f'за   {sem_word}   навчальний семестр.', f13)
    try:
        ws.merge_cells('A12:G12')
    except Exception:
        pass

    control = (_get(sheet, 'control_form') or '').lower()
    hours = _get(sheet, 'hours_credits') or ''
    _set(
        ws, 13, 1,
        f'Форма семестрового контролю  {control}     Загальна кількість годин   {hours}',
        f13,
    )
    try:
        ws.merge_cells('A13:G13')
    except Exception:
        pass

    teacher = _get(sheet, 'teacher_name') or ''
    _set(ws, 15, 1, f'Викладач                                                   {teacher}', f13)
    try:
        ws.merge_cells('A15:G15')
    except Exception:
        pass

    # Заголовок таблиці (рядки 17–18)
    f12b = Font(name=_FONT, size=12, bold=True)
    ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)
    headers_r17 = [
        (1, '№ з\\п'),
        (2, 'Прізвище та ініціали студента'),
        (3, 'Номер індивідуального навчального плану'),
        (4, 'Оцінка'),
    ]
    for col, text in headers_r17:
        _set(ws, 17, col, text, f12b, ctr, _thin)

    headers_r18 = [
        (4, 'за національною шкалою'),
        (5, 'кількість балів за 100 бальною шкалою'),
        (6, 'ECTS'),
        (7, 'Підпис викладача'),
    ]
    for col, text in headers_r18:
        _set(ws, 18, col, text, f12b, ctr, _thin)

    # merges як у прикладі
    for rng in ('A17:A18', 'B17:B18', 'C17:C18', 'D17:F17'):
        try:
            ws.merge_cells(rng)
        except Exception:
            pass
    ws.row_dimensions[18].height = 43.5


def _write_students(ws, sheet: dict, students: list, start: int = 19) -> int:
    f11 = Font(name=_FONT, size=11)
    f12 = Font(name=_FONT, size=12)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center')
    teacher = _get(sheet, 'teacher_name') or ''
    control = _get(sheet, 'control_form') or 'Екзамен'

    n = len(students)
    if n == 0:
        # один порожній рядок
        for c in range(1, 8):
            _set(ws, start, c, None, f11, center, _thin)
        _set(ws, start, 1, 1, f11, center, _thin)
        _set(ws, start, 6, _ects_formula(start), f12, center, _thin)
        return start

    for i, st in enumerate(students):
        r = start + i
        points = _get(st, 'points')
        nat = _get(st, 'national_grade') or _national_from_points(points, control)
        code = _get(st, 'edebo_code') or ''
        name = _get(st, 'full_name') or ''
        sign = _get(st, 'teacher_sign') or teacher

        _set(ws, r, 1, i + 1, f11, center, _thin)
        _set(ws, r, 2, name, f11, left, _thin)
        _set(ws, r, 3, code, f11, center, _thin)
        _set(ws, r, 4, nat, f12, center, _thin)
        _set(ws, r, 5, points if points is not None and points != '' else '', f11, center, _thin)
        _set(ws, r, 6, _ects_formula(r), f12, center, _thin)
        _set(ws, r, 7, sign, f11, center, _thin)
        ws.row_dimensions[r].height = 16.5

    return start + n - 1


def _write_footer(ws, first_student: int, last_student: int, n_students: int):
    f14b = Font(name=_FONT, size=14, bold=True)
    f12b = Font(name=_FONT, size=12, bold=True)
    f12 = Font(name=_FONT, size=12)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    sig = last_student + 3
    _set(ws, sig, 4, 'Декан факультету_____________________', f14b)
    _set(ws, sig + 2, 4, 'Викладач____________________________', f14b)

    stats_header = sig + 4
    for col, text in [
        (1, 'Результати оцінювання'),
        (3, 'Кількість'),
        (4, '%'),
        (5, 'Якість успішності,%'),
        (6, 'Абсолютна успішність,%'),
    ]:
        _set(ws, stats_header, col, text, f12b, center)

    total = max(n_students, 1)
    rng = f'$F${first_student}:$F${last_student}'
    bands = [
        ('А (90-100)', 'A'),
        ('В(82-89)', 'B'),
        ('С(74-81)', 'C'),
        ('D(64-73)', 'D'),
        ('E(60-63)', 'E'),
        ('FX(35-59)', 'FX'),
        ('F(1-34)', 'F'),
        ("Не з'явився", None),
    ]
    first_band = stats_header + 1
    for i, (label, letter) in enumerate(bands):
        r = first_band + i
        _set(ws, r, 1, label, f12b, center)
        if letter:
            _set(ws, r, 3, f'=COUNTIF({rng},"{letter}")', f12, center)
        else:
            _set(ws, r, 3, f'=COUNTBLANK($E${first_student}:$E${last_student})', f12, center)
        cell = _set(ws, r, 4, f'=C{r}/{total}', f12, center)
        cell.number_format = '0%'

    cell = _set(
        ws, first_band, 5,
        f'=D{first_band}+D{first_band+1}+D{first_band+2}',
        f12, center,
    )
    cell.number_format = '0%'
    cell = _set(
        ws, first_band, 6,
        f'=D{first_band}+D{first_band+1}+D{first_band+2}+D{first_band+3}+D{first_band+4}',
        f12, center,
    )
    cell.number_format = '0%'


def _apply_column_widths(ws):
    widths = {'A': 5.75, 'B': 29.38, 'C': 15.75, 'D': 14.63, 'E': 14.25, 'F': 8, 'G': 16.13}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def build_statement_xlsx(sheet, students) -> BytesIO:
    """
    sheet — dict або Row з метаданими відомості.
    students — list of dict/Row.
    """
    sheet = dict(sheet) if sheet is not None and not isinstance(sheet, dict) else (sheet or {})
    students = list(students or [])
    # нормалізуємо студентів у dict
    norm_students = []
    for st in students:
        if isinstance(st, dict):
            norm_students.append(st)
        else:
            try:
                norm_students.append(dict(st))
            except Exception:
                norm_students.append({
                    'full_name': _get(st, 'full_name', ''),
                    'edebo_code': _get(st, 'edebo_code', ''),
                    'points': _get(st, 'points'),
                    'national_grade': _get(st, 'national_grade'),
                    'ects': _get(st, 'ects'),
                    'teacher_sign': _get(st, 'teacher_sign'),
                    'work_topic': _get(st, 'work_topic'),
                })

    template_path = _find_template()
    if template_path:
        wb = load_workbook(template_path)
        ws = wb.active
        # критично: прибрати merge нижче шапки, інакше запис у B19 падає
        _unmerge_from_row(ws, 17)
        # очистити область даних шаблону
        for row in ws.iter_rows(min_row=17, max_row=min(ws.max_row, 80), max_col=7):
            for cell in row:
                # лише значення — шрифт/межі з шаблону лишаються
                cell.value = None
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = 'Відомість'

    _apply_column_widths(ws)
    _fill_header(ws, sheet)
    first = 19
    last = _write_students(ws, sheet, norm_students, start=first)
    _write_footer(ws, first, last, len(norm_students))

    safe = str(_get(sheet, 'discipline_name') or 'Відомість')[:25]
    for ch in r'[]:*?/\\':
        safe = safe.replace(ch, '_')
    ws.title = safe or 'Відомість'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_group_course_summary_xlsx(group_meta: dict, sessions_blocks: list) -> BytesIO:
    """
    Зведена відомість по групі за всі сесії (як «Книга1»).

    sessions_blocks: [
      {
        'academic_year': '...',
        'semester': 1,
        'sheets': [ {id, discipline_name, control_form, hours_credits}, ... ],
        'grades': { student_id: { sheet_id: {national, points, ects} } },
      },
      ...
    ]
    students_order: list of {id, full_name} — передається в group_meta['students']
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    students = group_meta.get('students') or []
    wb = Workbook()
    ws = wb.active
    ws.title = 'Зведена'

    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    bold = Font(name='Times New Roman', size=10, bold=True)
    normal = Font(name='Times New Roman', size=9)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    header_fill = PatternFill('solid', fgColor='D9E2F3')
    avg_fill = PatternFill('solid', fgColor='FFF2CC')

    # Row 1: semester titles
    # Row 2: discipline names
    # Row 3: control form
    # Row 4: credits
    # Row 5: subheaders оцінка/бали/ECTS
    # Row 6+: students

    # columns: A=№ B=ПІБ then for each sheet 3 cols, then semester avg, ...
    ws['A1'] = group_meta.get('title') or 'Зведена відомість'
    ws['A1'].font = Font(name='Times New Roman', size=14, bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)

    ws['A2'] = (
        f"Група {group_meta.get('group_name') or ''} · курс {_roman(group_meta.get('course'))} · "
        f"{group_meta.get('study_form') or ''} · {group_meta.get('degree_level') or ''} · "
        f"{group_meta.get('specialty_code') or ''} {group_meta.get('specialty') or ''}"
    )
    ws['A2'].font = Font(name='Times New Roman', size=11)

    start_col = 3  # C
    # header rows 4-8
    header_sem = 4
    header_disc = 5
    header_ctrl = 6
    header_cred = 7
    header_sub = 8
    data_start = 9

    ws.cell(header_sub, 1, '№ з/п').font = bold
    ws.cell(header_sub, 2, "Прізвище, ім'я та по батькові студента").font = bold
    for c in (1, 2):
        ws.cell(header_sub, c).fill = header_fill
        ws.cell(header_sub, c).border = thin
        ws.cell(header_sub, c).alignment = center

    col = start_col
    block_ranges = []  # (sem_label, start_col, end_col_before_avg, avg_col)

    for block in sessions_blocks:
        sheets = block.get('sheets') or []
        if not sheets:
            continue
        sem = block.get('semester')
        year = block.get('academic_year') or ''
        sem_label = f"{sem} семестр" + (f" ({year})" if year else '')
        block_start = col

        for sh in sheets:
            # 3 columns per discipline
            disc = (sh.get('discipline_name') or '')[:40]
            ctrl = (sh.get('control_form') or '')[:1]  # З/Е
            if (sh.get('control_form') or '').lower().startswith('екз'):
                ctrl = 'Е'
            elif (sh.get('control_form') or '').lower().startswith('зал'):
                ctrl = 'З'
            credits = sh.get('hours_credits') or ''

            ws.cell(header_disc, col, disc).font = bold
            ws.cell(header_disc, col).alignment = center
            ws.merge_cells(start_row=header_disc, start_column=col, end_row=header_disc, end_column=col + 2)

            ws.cell(header_ctrl, col, ctrl).font = normal
            ws.cell(header_ctrl, col).alignment = center
            ws.merge_cells(start_row=header_ctrl, start_column=col, end_row=header_ctrl, end_column=col + 2)

            ws.cell(header_cred, col, credits).font = normal
            ws.cell(header_cred, col).alignment = center
            ws.merge_cells(start_row=header_cred, start_column=col, end_row=header_cred, end_column=col + 2)

            for i, h in enumerate(('оцінка', 'бали', 'ECTS')):
                cell = ws.cell(header_sub, col + i, h)
                cell.font = bold
                cell.alignment = center
                cell.fill = header_fill
                cell.border = thin

            for r in range(header_disc, header_sub):
                for c in range(col, col + 3):
                    ws.cell(r, c).border = thin
                    ws.cell(r, c).fill = header_fill

            col += 3

        avg_col = col
        ws.cell(header_disc, avg_col, 'Середній бал за семестр').font = bold
        ws.cell(header_disc, avg_col).alignment = center
        ws.cell(header_disc, avg_col).fill = avg_fill
        for r in range(header_disc, header_sub + 1):
            ws.cell(r, avg_col).border = thin
            ws.cell(r, avg_col).fill = avg_fill
        ws.cell(header_sub, avg_col, 'бали').font = bold
        ws.cell(header_sub, avg_col).alignment = center

        # semester label merge over block
        ws.cell(header_sem, block_start, sem_label).font = bold
        ws.cell(header_sem, block_start).alignment = center
        if avg_col > block_start:
            ws.merge_cells(start_row=header_sem, start_column=block_start, end_row=header_sem, end_column=avg_col)
        for c in range(block_start, avg_col + 1):
            ws.cell(header_sem, c).border = thin
            ws.cell(header_sem, c).fill = header_fill

        block_ranges.append({
            'sheets': sheets,
            'grades': block.get('grades') or {},
            'start': block_start,
            'avg_col': avg_col,
        })
        col = avg_col + 1

    # overall average column
    overall_col = col
    ws.cell(header_disc, overall_col, 'Середній бал за курс').font = bold
    ws.cell(header_disc, overall_col).fill = avg_fill
    ws.cell(header_disc, overall_col).alignment = center
    for r in range(header_disc, header_sub + 1):
        ws.cell(r, overall_col).border = thin
        ws.cell(r, overall_col).fill = avg_fill
    ws.cell(header_sub, overall_col, 'бали').font = bold

    # data rows
    for ri, st in enumerate(students):
        row = data_start + ri
        sid = st['id']
        ws.cell(row, 1, ri + 1).font = normal
        ws.cell(row, 1).alignment = center
        ws.cell(row, 1).border = thin
        ws.cell(row, 2, st.get('full_name') or '').font = normal
        ws.cell(row, 2).alignment = left
        ws.cell(row, 2).border = thin

        all_pts = []
        for br in block_ranges:
            sem_pts = []
            c = br['start']
            for sh in br['sheets']:
                g = (br['grades'].get(sid) or {}).get(sh['id']) or {}
                nat = g.get('national') or ''
                pts = g.get('points')
                ects = g.get('ects') or ''
                ws.cell(row, c, nat).font = normal
                ws.cell(row, c).alignment = center
                ws.cell(row, c).border = thin
                cell_pts = ws.cell(row, c + 1, pts if pts is not None and pts != '' else '')
                cell_pts.font = normal
                cell_pts.alignment = center
                cell_pts.border = thin
                ws.cell(row, c + 2, ects).font = normal
                ws.cell(row, c + 2).alignment = center
                ws.cell(row, c + 2).border = thin
                if pts is not None and pts != '':
                    try:
                        sem_pts.append(float(pts))
                    except (TypeError, ValueError):
                        pass
                c += 3
            avg = round(sum(sem_pts) / len(sem_pts), 2) if sem_pts else ''
            cell = ws.cell(row, br['avg_col'], avg)
            cell.font = bold
            cell.alignment = center
            cell.border = thin
            cell.fill = avg_fill
            all_pts.extend(sem_pts)

        overall = round(sum(all_pts) / len(all_pts), 2) if all_pts else ''
        cell = ws.cell(row, overall_col, overall)
        cell.font = bold
        cell.alignment = center
        cell.border = thin
        cell.fill = avg_fill

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 28
    for c in range(3, overall_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 8

    ws.row_dimensions[header_disc].height = 45
    ws.freeze_panes = 'C9'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _roman(course):
    m = {1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI'}
    try:
        return m.get(int(course), str(course or ''))
    except (TypeError, ValueError):
        return str(course or '')


def build_session_summary_xlsx(session_meta, sheets, by_sheet_students) -> BytesIO:
    """
    Загальна відомість сесії в Excel — усі дисципліни × усі студенти.
    Повторює структуру gen_vidomist_summary (Word), але у форматі .xlsx:

      ┌──────────────────────────────────────────────────────────────┐
      │                   ВІДОМІСТЬ <назва сесії>                     │
      │ ПВНЗ «Європейський університет», Львівська філія              │
      │ Група: … | Курс: … | Форма: … | Ступінь: …                   │
      │ Спеціальність: …                                             │
      │ Навчальний рік: … | Семестр: …                                │
      ├────┬────────────┬───────┬───────┬───────┬─────┬──────────────┤
      │    │            │ Дисц1                │ …   │              │
      │ №  │    ПІБ     ├───────┼───────┼───────┼─────┤ Середній бал │
      │    │            │ нац.  │ бали  │ ECTS  │ …   │              │
      ├────┼────────────┼───────┼───────┼───────┼─────┼──────────────┤
      │ 1  │ Іваненко… │   A   │  92   │   A   │ …   │     92.0     │
      └────┴────────────┴───────┴───────┴───────┴─────┴──────────────┘

    Параметри — ті самі, що й у `gen_vidomist_summary`:
      session_meta        — dict з полями сесії (name, group_name, course,
                            study_form, degree_level, specialty, specialty_code,
                            academic_year, semester, institution_name_and_status)
      sheets              — список dict (id, discipline_name, control_form, …)
      by_sheet_students   — {sheet_id: [{student_id, full_name, points,
                            national_grade, ects}, …], …}
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    def _g(st, key, default=None):
        if isinstance(st, dict):
            return st.get(key, default)
        try:
            return st[key]
        except Exception:
            return default

    session_meta = dict(session_meta) if not isinstance(session_meta, dict) else session_meta
    sheets = [dict(s) if not isinstance(s, dict) else s for s in (sheets or [])]

    # ── порядок студентів + оцінки ──
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

    n_disc = len(sheets)
    n_students = len(order)
    # № + ПІБ + 3×n_disc + Середній бал
    cols = 2 + n_disc * 3 + 1
    avg_col = 3 + n_disc * 3

    # ── стилі ──
    FONT_NAME = 'Times New Roman'
    f_title  = Font(name=FONT_NAME, size=14, bold=True)
    f_header = Font(name=FONT_NAME, size=9, bold=True)
    f_body   = Font(name=FONT_NAME, size=9)
    f_bold   = Font(name=FONT_NAME, size=9, bold=True)

    thin = Side(style='thin', color='888888')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left   = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    header_fill = PatternFill('solid', fgColor='E8E0CC')  # світло-кремовий
    zebra_fill  = PatternFill('solid', fgColor='F7F3EA')  # дуже світлий

    wb = Workbook()
    ws = wb.active
    ws.title = 'Загальна відомість'

    def put(row, col, value, font=None, align=None, fill=None, border_=None):
        cell = ws.cell(row=row, column=col, value=value)
        if font:    cell.font = font
        if align:   cell.alignment = align
        if fill:    cell.fill = fill
        if border_: cell.border = border_
        return cell

    # ── шапка документа ──
    r = 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
    title = 'ВІДОМІСТЬ'
    if session_meta.get('name'):
        title += ' — ' + session_meta['name']
    put(r, 1, title, f_title, center)
    ws.row_dimensions[r].height = 24
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
    inst = (session_meta.get('institution_name_and_status')
            or 'ПВНЗ «Європейський університет», Львівська філія')
    inst = inst.replace('\n', ', ')
    put(r, 1, inst, f_body, left)
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
    parts = []
    if session_meta.get('group_name'):  parts.append(f"Група: {session_meta['group_name']}")
    if session_meta.get('course'):      parts.append(f"Курс: {session_meta['course']}")
    if session_meta.get('study_form'):  parts.append(f"Форма: {session_meta['study_form']}")
    if session_meta.get('degree_level'):parts.append(f"Ступінь: {session_meta['degree_level']}")
    put(r, 1, ' | '.join(parts) if parts else '', f_body, left)
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
    spec_code = session_meta.get('specialty_code') or ''
    spec_name = session_meta.get('specialty') or ''
    spec_text = f"Спеціальність: {spec_code} {spec_name}".strip().rstrip(':').strip()
    if spec_code or spec_name:
        put(r, 1, spec_text, f_body, left)
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
    ys = []
    if session_meta.get('academic_year'): ys.append(f"Навчальний рік: {session_meta['academic_year']}")
    if session_meta.get('semester'):      ys.append(f"Семестр: {session_meta['semester']}")
    put(r, 1, ' | '.join(ys) if ys else '', f_body, left)
    r += 1

    r += 1  # порожній рядок

    # ── таблиця ──
    header_row_1 = r
    header_row_2 = r + 1
    data_start   = r + 2

    # Рядок 1: № | ПІБ | Дисципліна 1 (merge 3) | … | Середній бал (merge 2)
    put(header_row_1, 1, '№', f_header, center, header_fill, border)
    put(header_row_1, 2, "Прізвище, ім'я та по батькові",
        f_header, center, header_fill, border)

    for i, sh in enumerate(sheets):
        base_col = 3 + i * 3
        title_d = sh.get('discipline_name') or ''
        ctrl    = sh.get('control_form') or ''
        text = f'{title_d} ({ctrl})' if ctrl else title_d
        put(header_row_1, base_col, text, f_header, center, header_fill, border)
        if base_col + 2 > base_col:
            ws.merge_cells(
                start_row=header_row_1, start_column=base_col,
                end_row=header_row_1,   end_column=base_col + 2,
            )
        for cc in range(base_col, base_col + 3):
            ws.cell(row=header_row_1, column=cc).border = border

    put(header_row_1, avg_col, 'Середній бал',
        f_header, center, header_fill, border)
    ws.merge_cells(
        start_row=header_row_1, start_column=avg_col,
        end_row=header_row_2,   end_column=avg_col,
    )
    ws.cell(row=header_row_2, column=avg_col).border = border

    # вертикальне об'єднання № і ПІБ
    ws.merge_cells(start_row=header_row_1, start_column=1, end_row=header_row_2, end_column=1)
    ws.merge_cells(start_row=header_row_1, start_column=2, end_row=header_row_2, end_column=2)
    ws.cell(row=header_row_2, column=1).border = border
    ws.cell(row=header_row_2, column=2).border = border

    # Рядок 2: нац. | бали | ECTS для кожного предмета
    for i in range(n_disc):
        base_col = 3 + i * 3
        for j, h in enumerate(['нац.', 'бали', 'ECTS']):
            put(header_row_2, base_col + j, h, f_header, center, header_fill, border)

    ws.row_dimensions[header_row_1].height = 40
    ws.row_dimensions[header_row_2].height = 20

    # ── дані ──
    for ri, sid in enumerate(order):
        row = data_start + ri
        is_zebra = ri % 2 == 1
        fill = zebra_fill if is_zebra else None

        put(row, 1, ri + 1, f_body, center, fill, border)
        put(row, 2, names.get(sid, ''), f_body, left, fill, border)

        pts_list = []
        for i, sh in enumerate(sheets):
            base_col = 3 + i * 3
            g = (grades.get(sid) or {}).get(sh['id']) or {}
            put(row, base_col,     g.get('national_grade') or '', f_body, center, fill, border)
            pts = g.get('points')
            put(row, base_col + 1, '' if pts is None else pts,  f_body, center, fill, border)
            put(row, base_col + 2, g.get('ects') or '',        f_body, center, fill, border)
            if pts is not None and pts != '':
                try:
                    pts_list.append(float(pts))
                except (TypeError, ValueError):
                    pass

        avg = round(sum(pts_list) / len(pts_list), 2) if pts_list else ''
        put(row, avg_col, avg, f_bold, center, fill, border)

        ws.row_dimensions[row].height = 16

    # ── ширини колонок ──
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 32
    for i in range(n_disc):
        base_col = 3 + i * 3
        for j in range(3):
            ws.column_dimensions[get_column_letter(base_col + j)].width = 8
    ws.column_dimensions[get_column_letter(avg_col)].width = 10

    # ── зручності ──
    ws.freeze_panes = ws.cell(row=data_start, column=3)   # шапка не прокручується
    ws.print_title_rows = f'{header_row_1}:{header_row_2}'  # шапка на кожній сторінці при друці

    # ── налаштування сторінки ──
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ── підсумковий блок ──
    footer_row = data_start + n_students + 1
    ws.merge_cells(start_row=footer_row, start_column=1,
                   end_row=footer_row, end_column=cols)
    put(footer_row, 1, f"Усього студентів: {n_students} | "
                       f"Дисциплін: {n_disc}",
        f_bold, left)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf