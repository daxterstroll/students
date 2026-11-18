"""
routes.exam_sessions
====================
Сесійні відомості (окремий контур від subjects/grades).

Підмодулі:
  - schema.py   — ensure_* таблиці
  - service.py  — бізнес-логіка, скоуп, права, шкала ECTS
  - admin.py    — UI відповідального (сесії, відомості, затвердження, Excel)
  - public.py   — форма викладача за токеном
  - transfer.py — перенос затверджених оцінок у grades/activity_grades

Blueprint реєструється в app.py як exam_sessions_bp.
"""
from flask import Blueprint, abort, request

exam_sessions_bp = Blueprint('exam_sessions', __name__)

from . import admin  # noqa: E402, F401
from . import public  # noqa: E402, F401
from . import transfer  # noqa: E402, F401
from . import history  # noqa: E402, F401


@exam_sessions_bp.before_request
def disable_statement_excel_exports():
    """Keep the exam statement workflow Word-only; other Excel reports remain."""
    if request.endpoint in {
        'exam_sessions.sheet_excel',
        'exam_sessions.session_summary_xlsx',
    }:
        abort(404)


@exam_sessions_bp.after_app_request
def hide_statement_excel_links(response):
    """Remove links to disabled Excel exports from statement pages."""
    if not response.mimetype or response.mimetype != 'text/html':
        return response

    path = request.path or ''
    if not (
        path.startswith('/admin/exam-sessions/sheets/')
        or path.startswith('/admin/exam-sessions/')
    ):
        return response

    try:
        body = response.get_data(as_text=True)
    except (RuntimeError, UnicodeDecodeError):
        return response

    import re
    body = re.sub(
        r'\s*<a\b(?=[^>]*href=["\'][^"\']*/admin/exam-sessions/sheets/\d+/excel[^"\']*["\'])[^>]*>.*?</a>',
        '', body, flags=re.IGNORECASE | re.DOTALL,
    )
    body = re.sub(
        r'\s*<a\b(?=[^>]*href=["\'][^"\']*/admin/exam-sessions/\d+/summary-xlsx[^"\']*["\'])[^>]*>.*?</a>',
        '', body, flags=re.IGNORECASE | re.DOTALL,
    )
    response.set_data(body)
    return response
