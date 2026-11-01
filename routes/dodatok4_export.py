"""
Генерація xlsx «Додаток 4 — Повідомлення про зміну облікових даних».

Шаблон: static/military_forms/dodatok_4.xlsx
(скопіюйте dodatok_4_template.xlsx з artifacts).
"""

from __future__ import annotations

import io
import os
from copy import copy

from openpyxl import load_workbook


def _project_root():
    # routes/ -> project root
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _template_path():
    """Шукає шаблон у static/example_excel, потім static/military_forms."""
    root = _project_root()
    candidates = [
        os.path.join(root, 'static', 'example_excel', 'dodatok_4.xlsx'),
        os.path.join(root, 'static', 'example_excel', 'Додаток 4.xlsx'),
        os.path.join(root, 'static', 'example_excel', 'додаток 4.xlsx'),
        os.path.join(root, 'static', 'military_forms', 'dodatok_4.xlsx'),
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return candidates[0]  # для повідомлення про помилку


def generate_dodatok_4_xlsx(data: dict) -> io.BytesIO:
    """
    data keys:
      tck_header, full_name, registry_number, military_document,
      vos_code, military_rank, birth_date, change_text, basis_text,
      tck_name, address, tax_id (РНКОПП/РНОКПП),
      rector_name (optional)
    """
    from openpyxl.styles import Alignment

    path = _template_path()
    if not os.path.isfile(path):
        raise FileNotFoundError(
            f'Немає шаблону Додатка 4: {path}. '
            f'Покладіть шаблон у static/example_excel/dodatok_4.xlsx (або «Додаток 4.xlsx»)'
        )

    wb = load_workbook(path)
    ws = wb.active

    def set_cell(coord, value, wrap=False, row_height=None):
        if value is None:
            value = ''
        cell = ws[coord]
        cell.value = value
        if wrap:
            cell.alignment = Alignment(
                wrap_text=True,
                vertical='top',
                horizontal=cell.alignment.horizontal if cell.alignment else 'left',
            )
            if row_height is not None:
                ws.row_dimensions[cell.row].height = row_height

    # Шапка — ТЦК (одержувач)
    set_cell('D4', data.get('tck_header') or data.get('tck_name') or '', wrap=True, row_height=36)

    # РНКОПП / РНОКПП (правий верх) — ІПН студента
    set_cell('I2', data.get('tax_id') or data.get('rnkopp') or '')

    # 1. ПІБ
    set_cell('A12', data.get('full_name') or '')

    # 2. Реєстр. номер
    set_cell('G14', data.get('registry_number') or '')

    # 3. ВОД
    set_cell('F15', data.get('military_document') or '')

    # 4. ВОС
    set_cell('G17', data.get('vos_code') or '')

    # 5. Звання
    set_cell('D18', data.get('military_rank') or '')

    # 6. Дата народження
    set_cell('D19', data.get('birth_date') or '')

    # 7. Зміна облікових даних
    set_cell('E20', data.get('change_text') or '', wrap=True, row_height=30)

    # 8. Підстава
    set_cell('A23', data.get('basis_text') or '', wrap=True, row_height=36)

    # 9. Перебування на обліку (ТЦК) — довгі назви
    set_cell('A26', data.get('tck_name') or '', wrap=True, row_height=48)

    # 10. Адреса
    set_cell('A29', data.get('address') or '', wrap=True, row_height=36)

    if data.get('rector_name'):
        set_cell('G34', data['rector_name'])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
