"""
routes.exam_sessions
====================
Сесійні відомості (окремий контур від subjects/grades).

Підмодулі:
  - schema.py   — ensure_* таблиці
  - service.py  — бізнес-логіка, скоуп, шкала ECTS
  - admin.py    — UI відповідального (сесії, відомості, затвердження, Excel)
  - public.py   — форма викладача за токеном
  - transfer.py — перенос затверджених оцінок у grades/activity_grades

Blueprint реєструється в app.py як exam_sessions_bp.
"""
from flask import Blueprint

exam_sessions_bp = Blueprint('exam_sessions', __name__)

from . import admin  # noqa: E402, F401
from . import public  # noqa: E402, F401
from . import transfer  # noqa: E402, F401
from . import history  # noqa: E402, F401

# admin imports build_statement_xlsx directly. Replace that symbol with the
# form-aware implementation after admin's routes have been declared.
from .excel_form_export import build_statement_xlsx as _build_statement_xlsx  # noqa: E402
admin.build_statement_xlsx = _build_statement_xlsx
