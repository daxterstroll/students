"""
routes/exam_sessions
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
