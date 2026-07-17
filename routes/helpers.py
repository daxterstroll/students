"""
routes/helpers.py
==================
Спільні допоміжні функції, які раніше були продубльовані в декількох
файлах (admin.py, students.py, auth.py). Мета цього модуля - усунути
дублювання коду ("не повторюйся" / DRY) та зробити поведінку однаковою
у всьому додатку.

Якщо потрібно змінити, наприклад, як визначається ім'я поточного
користувача для логів, або як сортуються українські прізвища - це
робиться в ОДНОМУ місці (тут), а не в 40+ місцях по всьому проєкту.
"""

import locale

from flask import session

from routes.utils import logger


def current_username() -> str:
    """
    Повертає ім'я користувача з поточної сесії для використання в
    log_action(...) та повідомленнях.

    Раніше вираз `session.get('username', 'невідомо')` був
    продубльований ~50 разів в admin.py та students.py. Тепер
    достатньо викликати current_username().
    """
    return session.get('username', 'невідомо')


def sort_ukrainian(items, key_func):
    """
    Безпечно сортує список рядків (наприклад, студентів) за українським
    алфавітом.

    Раніше в трьох різних місцях admin.py (manage_diplomas,
    manage_subjects, manage_activities) був майже однаковий код:
        locale.setlocale(locale.LC_COLLATE, 'uk_UA.UTF-8')
        sorted(students, key=lambda s: locale.strxfrm(...))
    При цьому у двох з трьох місць НЕ було try/except - якщо на
    сервері не встановлена локаль 'uk_UA.UTF-8' (що дуже ймовірно на
    "чистому" Linux/Docker-образі), виклик кидав необроблений
    `locale.Error` і вся сторінка падала з 500 помилкою.

    Ця функція:
    1. Пробує встановити українську локаль (UTF-8, потім Windows-варіант).
    2. Якщо жодна локаль не доступна - НЕ падає, а сортує звичайним
       Python-порядком (та один раз пише попередження в лог, а не на
       кожен виклик).

    :param items: список елементів для сортування (наприклад, Row-об'єктів)
    :param key_func: функція, що повертає рядок для порівняння,
                     наприклад: lambda s: f"{s['last_name_UA']} {s['first_name_UA']}"
    """
    for loc in ('uk_UA.UTF-8', 'Ukrainian_Ukraine.1251'):
        try:
            locale.setlocale(locale.LC_COLLATE, loc)
            return sorted(items, key=lambda item: locale.strxfrm(key_func(item)))
        except locale.Error:
            continue

    logger.warning(
        "Українська локаль (uk_UA.UTF-8 / Ukrainian_Ukraine.1251) недоступна "
        "на цьому сервері - використано стандартне сортування Python "
        "замість коректного українського алфавітного порядку."
    )
    return sorted(items, key=key_func)
