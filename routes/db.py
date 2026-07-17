"""
routes/db.py
============
Робота з підключенням до SQLite та кастомна колація (правила сортування)
для українських рядків, яка використовується в SQL-запитах через
`ORDER BY ... COLLATE UKRAINIAN`.
"""

import logging
import locale
import os
import sqlite3

logger = logging.getLogger('Students')

# --------------------------------------------------------------------------
# Налаштування української локалі виконується ОДИН РАЗ при імпорті модуля,
# а не при кожному порівнянні рядків під час сортування.
#
# РАНІШЕ: ukrainian_collation() викликав locale.setlocale(...) на кожне
# порівняння двох рядків - а SQLite викликає функцію колації дуже часто
# під час ORDER BY (O(n log n) разів на запит). Це:
#   1) було вкрай неефективно;
#   2) не мало жодної обробки помилок - якщо на сервері не встановлена
#      локаль 'uk_UA.UTF-8' (типова ситуація на "чистому" Linux/Docker),
#      будь-який запит з ORDER BY ... COLLATE UKRAINIAN падав з
#      необробленим locale.Error і повертав користувачу 500-у помилку.
# --------------------------------------------------------------------------
_UA_COLLATION_AVAILABLE = False
for _loc in ('uk_UA.UTF-8', 'Ukrainian_Ukraine.1251'):
    try:
        locale.setlocale(locale.LC_COLLATE, _loc)
        _UA_COLLATION_AVAILABLE = True
        break
    except locale.Error:
        continue

if not _UA_COLLATION_AVAILABLE:
    logger.warning(
        "Українська локаль (uk_UA.UTF-8 / Ukrainian_Ukraine.1251) не "
        "встановлена на сервері. Сортування 'COLLATE UKRAINIAN' в SQL-запитах "
        "буде використовувати простий посимвольний порядок замість "
        "коректного українського алфавіту, доки локаль не буде встановлена "
        "в ОС (напр. `sudo locale-gen uk_UA.UTF-8`)."
    )


def ukrainian_collation(str1, str2):
    """
    Функція колації для SQLite (використовується як `COLLATE UKRAINIAN`
    в SQL-запитах, напр. в students.py: `ORDER BY s.last_name_UA COLLATE UKRAINIAN`).

    Порівнює два рядки за правилами української локалі, якщо вона доступна
    на сервері; інакше безпечно повертається до звичайного порівняння
    рядків Python, замість того, щоб кидати виняток і ламати запит.
    """
    if _UA_COLLATION_AVAILABLE:
        try:
            return locale.strcoll(str1, str2)
        except locale.Error:
            pass
    return -1 if str1 < str2 else (1 if str1 > str2 else 0)


def get_db():
    """
    Створює нове з'єднання з базою даних SQLite (students.db у корені
    проєкту, на рівень вище папки routes/).

    Кожен виклик повертає НОВЕ з'єднання - викликач відповідає за те, щоб
    закрити його (conn.close()), бажано в блоці try/finally, щоб з'єднання
    закривалось навіть при помилці всередині запиту.

    Налаштовує:
      - row_factory = sqlite3.Row, щоб рядки можна було читати як за
        індексом (row[0]), так і за назвою колонки (row['name']);
      - кастомну колацію "UKRAINIAN" для сортування українського тексту;
      - PRAGMA foreign_keys = ON, щоб працювали ON DELETE CASCADE
        та інші зовнішні ключі, визначені в схемі (init_db.py).
    """
    # Поднимаемся из routes в корень проекта students
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    DB_PATH = os.path.join(BASE_DIR, 'students.db')

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.create_collation("UKRAINIAN", ukrainian_collation)
    conn.execute("PRAGMA foreign_keys = ON")

    return conn
