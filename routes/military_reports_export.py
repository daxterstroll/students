"""
routes/military_reports_export.py
==================================
Генерація .xlsx-документів військового обліку:

  - Додаток 5 (Список персонального військового обліку) — 4 категорії:
      officers   — офіцери
      rank       — рядові / сержанти / старшини
      women      — жінки
      conscripts — призовники
  - Додаток 12 (Відомість оперативного обліку) — зведена таблиця

Джерело шаблону — реальні CSV-файли, надані адміністратором
(структура колонок і заголовків відтворена 1:1).

Дата в шапці («станом на») приймається ззовні — визначається в
routes.military_reports (annual / current / custom).

За приміткою до Списків (Постанова №812) записи відрахованих
зберігаються в списках до кінця поточного року — тобто при генерації
звіту за рік N показуються всі, хто відрахований у році N (незалежно
від того, на яку дату дивимось).

Логіка фільтрації для Додатка 5 (спільна з UI):
  1. Архівні / відчислені (students.archived = 1) → не показуються
  2. Заморожені (frozen_students з resolved_at IS NULL) → не показуються
  3. Вибулі за графою 18 (note_18 ≠ NULL) → не показуються:
     - Виключено за віком
     - Звільнено (відраховано)
     - Військовослужбовець
  4. Жінки без військових даних → не показуються
  5. Фільтр базової ліцензії (якщо не show_all_licenses)
  6. Категорія визначається через _detect_category
"""

import io
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from routes.utils import logger


# ═══════════════════════════════════════════════════════════════════════
# Спільні стилі
# ═══════════════════════════════════════════════════════════════════════

THIN = Side(style='thin', color='000000')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT_TOP = Alignment(horizontal='left', vertical='top', wrap_text=True)
FONT_BASE = Font(name='Times New Roman', size=9)
FONT_BOLD = Font(name='Times New Roman', size=9, bold=True)
FONT_HEADER = Font(name='Times New Roman', size=9, bold=True)
FONT_TITLE = Font(name='Times New Roman', size=12, bold=True)
FONT_SMALL = Font(name='Times New Roman', size=8, italic=True)
FILL_HEADER = PatternFill('solid', fgColor='D9D9D9')


def _format_date_ua(d):
    """date → 'ДД.ММ.РРРР'"""
    return d.strftime('%d.%m.%Y') if d else ''


def _row_height_for_wrap(row, lines=2):
    """Підганяє висоту рядка під перенесений текст."""
    row.height = 12 * lines + 4


# ═══════════════════════════════════════════════════════════════════════
# Додаток 5 — Список персонального військового обліку
# ═══════════════════════════════════════════════════════════════════════

# Заголовки 24 колонок Додатка 5 (з реального CSV адміністратора)
DODATOK_5_HEADERS = [
    'Порядковий номер',
    'Категорія військового обов\'язку',
    'Військове звання',
    'Прізвище, ім\'я, по батькові',
    'Число, місяць, рік народження',
    'Реєстраційний номер облікового запису в Єдиному державному реєстрі призовників, військовозобов\'язаних та резервістів',
    'Реєстраційний номер облікової картки платника податків (за наявності)',
    'Військово-облікова спеціальність',
    'Реквізити військово-облікового документа (серія (за наявності), номер, ким та коли видано, час та дата формування)',
    'Реквізити паспорта громадянина України (серія (за наявності), номер, ким і коли видано)',
    'Адреса задекларованого (зареєстрованого) місця проживання',
    'Найменування територіального центру комплектування та соціальної підтримки, органу СБУ, відповідного підрозділу розвідувального органу, в якому перебуває на військовому обліку',
    'Відомості про наявність відстрочки від призову на військову службу (направлення для проходження базової військової служби)',
    'Відомості про перебування на спеціальному військовому обліку',
    'Відомості про проходження військової служби',
    'Відомості про наявність мобілізаційного розпорядження (дата видачі)',
    'Посада, реквізити акта про прийняття/звільнення на посаду/з посади',
    'Реквізити повідомлення про прийняття/звільнення на посаду/з посади',
    'Ознака філії',
    'Примітка',
    '7. Зміна облікових даних (колонка для повідомлення Д4)',
    '8. Підстава для зміни облікових даних (колонка для повідомлення Д4)',
    'З 1 січня {year} р. зараховано на навчання (для Додатку 12)',
    'З 1 січня {year} р. відраховано з навчання (для Додатку 12)',
]

# Ширини колонок (у символах) — підібрані емпірично під альбомну A4
DODATOK_5_WIDTHS = [
    6,   # 1 №
    16,  # 2 Категорія
    16,  # 3 Звання
    28,  # 4 ПІБ
    12,  # 5 Дата народження
    20,  # 6 Реєстр ЄДР
    15,  # 7 РНОКПП
    14,  # 8 ВОС
    32,  # 9 ВОД
    32,  # 10 Паспорт
    32,  # 11 Адреса
    30,  # 12 ТЦК
    25,  # 13 Відстрочка
    14,  # 14 Спец. облік
    14,  # 15 Служба
    16,  # 16 Моб. розпорядження
    18,  # 17 Посада
    18,  # 18 Повідомлення
    10,  # 19 Філія
    20,  # 20 Примітка
    20,  # 21 Зміна даних
    20,  # 22 Підстава
    20,  # 23 Зараховано
    20,  # 24 Відраховано
]

CATEGORY_TITLE = {
    'officers': (
        'СПИСОК персонального військового обліку військовозобов\'язаних '
        '(резервістів) офіцерського складу'
    ),
    'rank': (
        'СПИСОК персонального військового обліку військовозобов\'язаних '
        '(резервістів) рядового, сержантського та старшинського складу'
    ),
    'women': (
        'СПИСОК персонального військового обліку військовозобов\'язаних '
        'та резервістів з числа жінок'
    ),
    'conscripts': (
        'СПИСОК персонального військового обліку призовників'
    ),
}


def _write_dodatok_5_header(ws, category, report_date, report_year):
    """Шапка документа (до 8-го рядка включно)."""
    # Рядок 1-2: "Додаток 5 / до Порядку" — праворуч
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    c = ws.cell(row=1, column=1, value='Додаток 5')
    c.font = FONT_BOLD
    c.alignment = Alignment(horizontal='left')

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    c = ws.cell(row=2, column=1, value='до Порядку')
    c.font = FONT_BASE
    c.alignment = Alignment(horizontal='left')

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=4)
    c = ws.cell(
        row=3, column=1,
        value='(в редакції постанови Кабінету Міністрів України '
              'від 30 липня 2025 р. № 916, зі змінами від 10.06.2026 № 812)',
    )
    c.font = FONT_SMALL
    c.alignment = Alignment(horizontal='left', wrap_text=True)
    ws.row_dimensions[3].height = 24

    # Рядок 5: заголовок
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=24)
    c = ws.cell(
        row=5, column=1,
        value=f"{CATEGORY_TITLE.get(category, 'СПИСОК')} "
              f"станом на {_format_date_ua(report_date)}",
    )
    c.font = FONT_TITLE
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[5].height = 36

    # Рядок 7: нумерація колонок 1..24
    for i in range(1, 25):
        cell = ws.cell(row=7, column=i, value=i)
        cell.font = FONT_BASE
        cell.alignment = CENTER
        cell.border = BORDER
        cell.fill = FILL_HEADER
    ws.row_dimensions[7].height = 18

    # Рядок 8: заголовки колонок
    for i, h in enumerate(DODATOK_5_HEADERS, start=1):
        text = h.format(year=report_year) if '{year}' in h else h
        cell = ws.cell(row=8, column=i, value=text)
        cell.font = FONT_HEADER
        cell.alignment = CENTER
        cell.border = BORDER
        cell.fill = FILL_HEADER
    ws.row_dimensions[8].height = 80


def _write_dodatok_5_row(ws, row_num, data):
    """Один рядок даних Додатка 5."""
    values = [
        row_num,                              # 1 №
        data.get('status_badge') or '',       # 2 Категорія
        data.get('military_rank') or '',      # 3 Звання
        data.get('full_name') or '',          # 4 ПІБ
        data.get('birth_date') or '',         # 5 Дата народження
        data.get('registry_number') or '',    # 6 Реєстр ЄДР
        data.get('tax_id') or '',             # 7 РНОКПП
        data.get('vos') or '',                # 8 ВОС
        data.get('military_document') or '',  # 9 ВОД
        data.get('passport') or '',           # 10 Паспорт
        data.get('address') or '',            # 11 Адреса
        data.get('tck_name') or '',           # 12 ТЦК
        data.get('deferral') or '',           # 13 Відстрочка
        '',                                   # 14 Спец. облік
        '',                                   # 15 Служба
        '',                                   # 16 Моб. розпорядження
        '',                                   # 17 Посада
        '',                                   # 18 Повідомлення
        data.get('license_short') or '',      # 19 Філія
        data.get('note_18') or '',            # 20 Примітка
        data.get('change_credentials') or '', # 21 Зміна даних
        data.get('reason_for_change') or '',  # 22 Підстава
        data.get('enrolled_note') or '',      # 23 Зараховано
        data.get('expelled_note') or '',      # 24 Відраховано
    ]
    row = ws.max_row + 1
    for col, val in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = FONT_BASE
        cell.border = BORDER
        cell.alignment = CENTER if col in (1, 5, 19) else LEFT_TOP
    # Висота — під кількість рядків у найдовшій комірці (наближено)
    max_len = max(len(str(v)) for v in values if v) if any(values) else 10
    lines = max(1, min(6, (max_len // 30) + 1))
    _row_height_for_wrap(ws.row_dimensions[row], lines=lines)


def generate_dodatok_5_xlsx(category, report_date, rows, base_license_label=''):
    """
    Генерує .xlsx Додатка 5 для однієї категорії.

    :param category: 'officers' | 'rank' | 'women' | 'conscripts'
    :param report_date: datetime.date — дата в шапці («станом на»)
    :param rows: список dict з полями:
        full_name, birth_date, tax_id, status_badge, military_rank,
        registry_number, vos, military_document, passport, address,
        tck_name, deferral, note_18, change_credentials,
        reason_for_change, license_short, enrolled_note, expelled_note
    :param base_license_label: коротка назва базової ліцензії (для шапки)
    :return: BytesIO з .xlsx
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Додаток 5'

    report_year = report_date.year

    # Ширини колонок
    for i, w in enumerate(DODATOK_5_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Шапка
    _write_dodatok_5_header(ws, category, report_date, report_year)

    # Дані
    for i, r in enumerate(rows, start=1):
        _write_dodatok_5_row(ws, i, r)

    # Налаштування сторінки (альбомна A4)
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = '7:8'  # повторювати шапку при друці

    # Закріпити перший стовпець і шапку
    ws.freeze_panes = 'B9'

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
    
# ═══════════════════════════════════════════════════════════════════════
# Додаток 12 — Відомість оперативного обліку
# ═══════════════════════════════════════════════════════════════════════

DODATOK_12_HEADERS_TEMPLATE = [
    'Порядковий номер',
    'Найменування групи',
    'Усього',
    'Мають мобілізаційні розпорядження',
    'Заброньовані',
    'Військовозобов\'язані, які не заброньовані та не мають мобілізаційних розпоряджень',
    'З 1 січня {year} р. зараховано на навчання',
    'З 1 січня {year} р. відраховано з навчання',
]

DODATOK_12_WIDTHS = [
    8,   # №
    60,  # Група
    10,  # Усього
    18,  # Моб. розпорядження
    14,  # Заброньовані
    22,  # Не заброньовані
    16,  # Зараховано
    16,  # Відраховано
]


def _write_dodatok_12_header(ws, report_date, report_year):
    """Шапка Додатка 12 (до 11-го рядка)."""
    # Рядок 1-2: "Додаток 12 / до Порядку"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    c = ws.cell(row=1, column=1, value='Додаток 12')
    c.font = FONT_BOLD
    c.alignment = Alignment(horizontal='left')

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=2)
    c = ws.cell(row=2, column=1, value='до Порядку')
    c.font = FONT_BASE
    c.alignment = Alignment(horizontal='left')

    # Рядок 4: "ВІДОМІСТЬ"
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=8)
    c = ws.cell(row=4, column=1, value='ВІДОМІСТЬ')
    c.font = FONT_TITLE
    c.alignment = Alignment(horizontal='center')

    # Рядок 5-6: опис
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=8)
    c = ws.cell(
        row=5, column=1,
        value=(
            "оперативного обліку призовників, військовозобов'язаних "
            "та резервістів у"
        ),
    )
    c.font = FONT_BASE
    c.alignment = Alignment(horizontal='center')

    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=8)
    c = ws.cell(
        row=6, column=1,
        value='Приватному вищому навчальному закладі «Європейський університет»',
    )
    c.font = FONT_BASE
    c.alignment = Alignment(horizontal='center', wrap_text=True)
    ws.row_dimensions[6].height = 18

    # Рядок 7: "(назва органу...)"
    ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=8)
    c = ws.cell(
        row=7, column=1,
        value=(
            '(назва державного органу, органу місцевого самоврядування, '
            'підприємства, установи, організації)'
        ),
    )
    c.font = FONT_SMALL
    c.alignment = Alignment(horizontal='center')

    # Рядок 9: дата
    ws.merge_cells(start_row=9, start_column=1, end_row=9, end_column=8)
    c = ws.cell(row=9, column=1, value=_format_date_ua(report_date))
    c.font = FONT_BASE
    c.alignment = Alignment(horizontal='center')

    # Рядок 10: "(дата внесення...)"
    ws.merge_cells(start_row=10, start_column=1, end_row=10, end_column=8)
    c = ws.cell(
        row=10, column=1,
        value='(дата внесення останніх змін до відомості оперативного обліку)',
    )
    c.font = FONT_SMALL
    c.alignment = Alignment(horizontal='center')

    # Рядок 12: заголовки колонок
    for i, h in enumerate(DODATOK_12_HEADERS_TEMPLATE, start=1):
        text = h.format(year=report_year) if '{year}' in h else h
        cell = ws.cell(row=12, column=i, value=text)
        cell.font = FONT_HEADER
        cell.alignment = CENTER
        cell.border = BORDER
        cell.fill = FILL_HEADER
    ws.row_dimensions[12].height = 80


def _write_dodatok_12_row(ws, values, bold=False, fill=None):
    """Один рядок даних Додатка 12."""
    row = ws.max_row + 1
    for col, val in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = FONT_BOLD if bold else FONT_BASE
        cell.border = BORDER
        cell.alignment = CENTER if col != 2 else LEFT_TOP
        if fill:
            cell.fill = fill
    _row_height_for_wrap(ws.row_dimensions[row], lines=2)


def generate_dodatok_12_xlsx(data, report_date):
    """
    Генерує .xlsx Додатка 12 (зведена відомість оперативного обліку).

    :param data: dict з полями:
        officers, rank, women, conscripts        — кількості по 4 групах
        excluded_by_age                          — виключені за віком
        dismissed_this_year                      — відраховані в поточному році
        mobilized                                — мобілізовані
        enrolled: {officers, rank, women, conscripts}  — зараховані цього року
        expelled: {officers, rank, women, conscripts}  — відраховані цього року
        base_license_label                       — коротка назва базової ліцензії
        show_all_licenses                        — чи враховувати всі ліцензії
    :param report_date: datetime.date — «станом на»
    :return: BytesIO з .xlsx
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Додаток 12'

    report_year = report_date.year

    # Ширини колонок
    for i, w in enumerate(DODATOK_12_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Шапка
    _write_dodatok_12_header(ws, report_date, report_year)

    # ── Дані ──
    officers = data.get('officers', 0)
    rank = data.get('rank', 0)
    women = data.get('women', 0)
    conscripts = data.get('conscripts', 0)
    excluded_by_age = data.get('excluded_by_age', 0)
    dismissed_this_year = data.get('dismissed_this_year', 0)
    mobilized = data.get('mobilized', 0)

    enrolled = data.get('enrolled', {}) or {}
    expelled = data.get('expelled', {}) or {}
    enrolled_officers = enrolled.get('officers', 0)
    enrolled_rank = enrolled.get('rank', 0)
    enrolled_women = enrolled.get('women', 0)
    enrolled_conscripts = enrolled.get('conscripts', 0)
    expelled_officers = expelled.get('officers', 0)
    expelled_rank = expelled.get('rank', 0)
    expelled_women = expelled.get('women', 0)
    expelled_conscripts = expelled.get('conscripts', 0)

    mobilized_total = officers + rank + women
    total_row_1 = mobilized_total
    total_all = mobilized_total + conscripts

    fill_secondary = PatternFill('solid', fgColor='F2F2F2')
    fill_warning = PatternFill('solid', fgColor='FFF2CC')
    fill_info = PatternFill('solid', fgColor='DEEBF7')

    # Рядок 1 — Військовозобов'язані (у т.ч. резервісти)
    _write_dodatok_12_row(ws, [
        '1.', "Військовозобов'язані (у тому числі резервісти), із них:",
        total_row_1, '—', '—', '—',
        enrolled_officers + enrolled_rank + enrolled_women,
        expelled_officers + expelled_rank + expelled_women,
    ], bold=True, fill=fill_secondary)

    # Рядок 1.1 — Офіцери
    _write_dodatok_12_row(ws, [
        '1.1', '— офіцерського складу',
        officers, '—', '—', '—',
        enrolled_officers, expelled_officers,
    ])

    # Рядок 1.2 — Рядові / сержанти / старшини
    _write_dodatok_12_row(ws, [
        '1.2', '— рядового, сержантського та старшинського складу',
        rank, '—', '—', '—',
        enrolled_rank, expelled_rank,
    ])

    # Рядок 1.3 — Жінки
    _write_dodatok_12_row(ws, [
        '1.3', '— жінки',
        women, '—', '—', '—',
        enrolled_women, expelled_women,
    ])

    # Рядок 2 — Призовники
    _write_dodatok_12_row(ws, [
        '2.', 'Призовники',
        conscripts, '—', '—', '—',
        enrolled_conscripts, expelled_conscripts,
    ], bold=True, fill=fill_secondary)

    # Порожній рядок-роздільник
    _write_dodatok_12_row(ws, ['', '', '', '', '', '', '', ''])

    # Рядок 3 — Виключено за віком (Постанова №812)
    _write_dodatok_12_row(ws, [
        '3.', 'Виключено з військового обліку за віком',
        excluded_by_age, '—', '—', '—', '—', '—',
    ], fill=fill_secondary)

    # Рядок 4 — Звільнено (відраховано)
    _write_dodatok_12_row(ws, [
        '4.', f'Звільнено з роботи, завершено навчання (відраховано) — за {report_year} рік',
        dismissed_this_year, '—', '—', '—', '—', '—',
    ], fill=fill_warning)

    # Рядок 5 — Мобілізовані
    _write_dodatok_12_row(ws, [
        '5.', 'Військовослужбовці (мобілізовані)',
        mobilized, '—', '—', '—', '—', '—',
    ], fill=fill_info)

    # ── Підсумкові рядки ──
    _write_dodatok_12_row(ws, [
        '', 'Разом (рядки 1 + 2):',
        total_all, '—', '—', '—',
        enrolled_officers + enrolled_rank + enrolled_women + enrolled_conscripts,
        expelled_officers + expelled_rank + expelled_women + expelled_conscripts,
    ], bold=True, fill=fill_secondary)

    _write_dodatok_12_row(ws, [
        '', 'Разом вибулих (рядки 3 + 4 + 5):',
        excluded_by_age + dismissed_this_year + mobilized, '—', '—', '—', '—', '—',
    ], bold=True, fill=fill_secondary)

    # ── Підпис ──
    last_row = ws.max_row + 2
    ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
    c = ws.cell(
        row=last_row, column=1,
        value='Особа, відповідальна за ведення військового обліку:',
    )
    c.font = FONT_BASE
    c.alignment = Alignment(horizontal='left')

    ws.merge_cells(start_row=last_row, start_column=6, end_row=last_row, end_column=8)
    c = ws.cell(
        row=last_row, column=6,
        value='Оксана ГОЛОВКО / Юлія БЕРЕЗНЮК',
    )
    c.font = FONT_BASE
    c.alignment = Alignment(horizontal='center')

    ws.merge_cells(start_row=last_row + 1, start_column=6, end_row=last_row + 1, end_column=8)
    c = ws.cell(row=last_row + 1, column=6, value='(підпис)')
    c.font = FONT_SMALL
    c.alignment = Alignment(horizontal='center')

    ws.merge_cells(start_row=last_row + 3, start_column=1, end_row=last_row + 3, end_column=5)
    c = ws.cell(row=last_row + 3, column=1, value='Ректор:')
    c.font = FONT_BASE
    c.alignment = Alignment(horizontal='left')

    ws.merge_cells(start_row=last_row + 3, start_column=6, end_row=last_row + 3, end_column=8)
    c = ws.cell(row=last_row + 3, column=6, value='Олена ТИМОШЕНКО')
    c.font = FONT_BASE
    c.alignment = Alignment(horizontal='center')

    ws.merge_cells(start_row=last_row + 4, start_column=6, end_row=last_row + 4, end_column=8)
    c = ws.cell(row=last_row + 4, column=6, value='(підпис)')
    c.font = FONT_SMALL
    c.alignment = Alignment(horizontal='center')

    # Налаштування сторінки
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out