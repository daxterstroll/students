"""
routes/students/__init__.py
============================
Раніше весь функціонал жив в одному файлі routes/students.py (~1750
рядків) - розбито на тематичні підмодулі нижче, кожен реєструє свої
маршрути на тому самому students_bp, тому жодна назва ендпоінта
(`students.student_details`, `students.edit_student` тощо) не
змінилась - усі виклики url_for('students.xxx') у шаблонах і далі
працюють без змін.

Підмодулі:
  - core.py           Список студентів, картка студента, додати/
                      редагувати/видалити, фото, заморозка, періоди
                      навчання, посилання на оновлення даних,
                      генерація документів для одного студента, FAQ
  - military.py       Військовий облік студента (+ скани)
  - grades.py         Виставлення оцінок (предмети й активності)
  - import_export.py  Масовий імпорт студентів з Excel
"""

from flask import Blueprint

students_bp = Blueprint('students', __name__)

from . import core
from . import military
from . import grades
from . import import_export
