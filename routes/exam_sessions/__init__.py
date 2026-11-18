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
def disable_sheet_excel_export():
    """Excel remains available for the complete session summary, only sheet Excel is disabled."""
    if request.endpoint == 'exam_sessions.sheet_excel':
        abort(404)


@exam_sessions_bp.after_app_request
def hide_sheet_excel_link(response):
    """Hide only the individual sheet Excel link; keep session summary Excel link."""
    if not response.mimetype or response.mimetype != 'text/html':
        return response

    path = request.path or ''
    if not path.startswith('/admin/exam-sessions/sheets/'):
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
    response.set_data(body)
    return response
