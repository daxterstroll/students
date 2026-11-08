"""
Експорт відомості сесії — через routes.gen_docx.gen_vidomist_sheet
(той самий DocxTemplate стек що додатки / gen_doc).
"""
from __future__ import annotations

from io import BytesIO


def build_sheet_docx(sheet, students) -> BytesIO:
    from routes.gen_docx import gen_vidomist_sheet
    result = gen_vidomist_sheet(sheet, students, out=None)
    if isinstance(result, BytesIO):
        return result
    # якщо раптом шлях
    bio = BytesIO()
    with open(result, 'rb') as f:
        bio.write(f.read())
    bio.seek(0)
    return bio


def build_session_summary_docx(session_meta: dict, sheets: list, by_sheet_students: dict) -> BytesIO:
    """Загальна відомість — через gen_docx.gen_vidomist_summary (шаблон + таблиця)."""
    from routes.gen_docx import gen_vidomist_summary
    result = gen_vidomist_summary(session_meta, sheets, by_sheet_students, out=None)
    if isinstance(result, BytesIO):
        return result
    bio = BytesIO()
    with open(result, 'rb') as f:
        bio.write(f.read())
    bio.seek(0)
    return bio
