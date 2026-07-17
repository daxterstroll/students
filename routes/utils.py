"""
routes/utils.py
================
Спільні утиліти: логер додатку, логування дій користувачів (log_action),
декоратори перевірки авторизації/прав (login_required, permission_required),
транслітерація українських імен та пошук шаблонів .docx.

Логування: усі модулі, які хочуть писати помилки/події в app.log,
ПОВИННІ використовувати саме логер `logger`, визначений нижче
(`from routes.utils import logger`), а не голий `import logging`.
Лише `logger` тут налаштований з файловим обробником (app.log) та
консольним виводом; звичайний кореневий логер Python (`logging.error(...)`)
обробників не має і в файл нічого не пише.
"""

from functools import wraps
from flask import session, redirect, url_for, flash
import logging
import os
from routes.db import get_db
import json
import re

# Настройка пути к файлу
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
log_file_path = os.path.join(PROJECT_ROOT, 'app.log')

# Настройка глобального логгера
logger = logging.getLogger('Students')
logger.setLevel(logging.DEBUG)

# Удаляем все существующие обработчики, чтобы избежать конфликтов
if logger.handlers:
    logger.handlers.clear()

# Создаем обработчики
file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
console_handler = logging.StreamHandler()

# Форматирование
formatter = logging.Formatter('%(asctime)s %(levelname)s | %(message)s ', datefmt='%Y-%m-%d | %H:%M:%S |')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Добавляем обработчики к логгеру
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Проверка создания файла логов
try:
    with open(log_file_path, 'a', encoding='utf-8'):
        pass
    logger.debug(f"Логирование инициализировано. Файл логов: {log_file_path}")
except Exception as e:
    logger.error(f"Ошибка при доступе к файлу логов {log_file_path}: {e}")
    print(f"Ошибка при доступе к файлу логов: {e}")

def log_action(username, action, group_ids=None, mode=None, details=None):
    """
    Записує дію користувача в лог-файл (app.log) у вигляді:
    "YYYY-MM-DD | HH:MM:SS | INFO | 👤 username - дія (групи: ...) | деталі"

    Параметри:
        username   - ім'я користувача, що виконав дію (див. helpers.current_username())
        action     - короткий опис дії, наприклад "додав студента"
        group_ids  - ID груп, яких стосується дія (для не-адмінів ці ID
                     перетворюються на людські назви груп)
        mode       - необов'язковий режим/контекст дії
        details    - додаткові деталі (наприклад, кількість оброблених рядків)

    ВАЖЛИВО: раніше ця функція не мала обробки помилок. Якщо get_db()
    не міг відкрити файл БД (немає прав доступу, диск переповнений тощо)
    або запит до groups падав, виняток "вилітав" з log_action() і міг
    зламати весь HTTP-запит - навіть якщо основна дія (наприклад,
    збереження студента) вже завершилась успішно. Тепер збій самого
    логування ніколи не ламає основний функціонал, але обов'язково
    фіксується як ERROR в лог, щоб про проблему було відомо.
    """
    conn = None
    try:
        conn = get_db()
        role = session.get('role', 'невідомо')

        group_names_str = ''
        if group_ids is not None and role != 'admin':
            if not isinstance(group_ids, list):
                group_ids = [group_ids] if group_ids else []
            if group_ids:
                placeholders = ','.join('?' for _ in group_ids)
                group_names = conn.execute(
                    f"""
                    SELECT name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
                    FROM groups
                    WHERE id IN ({placeholders})
                    ORDER BY name, start_year
                    """,
                    group_ids
                ).fetchall()
                group_names_str = ', '.join([row['display_name'] for row in group_names]) if group_names else 'немає груп'

        log_msg = f"👤 {username} - {action}"
        if mode:
            log_msg += f" (режим: {mode})"
        if group_names_str:
            log_msg += f" (групи: {group_names_str})"
        if details:
            log_msg += f" | {details}"

        logger.info(log_msg)

    except Exception as e:
        # Логування НІКОЛИ не повинно "ронити" запит користувача, що його викликав.
        logger.error(f"Не вдалося записати дію в лог (username={username}, action={action}): {e}", exc_info=True)
    finally:
        if conn is not None:
            conn.close()
    
def login_required(role=None):
    """
    Декоратор: пускає на сторінку тільки авторизованих користувачів.
    Стара версія для зворотної сумісності - для нових перевірок прав
    краще використовувати permission_required() нижче.

    Args:
        role (str, optional): якщо вказано (напр. 'admin'), доступ
            дозволено лише користувачам з саме такою роллю; інакше -
            403 Forbidden. Якщо None/'' - достатньо просто бути залогіненим.

    Поведінка при відсутності сесії: редірект на сторінку логіну.
    """
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user_id' not in session:
                return redirect(url_for('auth.login'))
            if role and session.get('role') != role:
                return "403 Forbidden", 403
            return f(*args, **kwargs)
        return decorated_function
    return decorator

def permission_required(permission=None):
    """
    Декоратор: перевіряє авторизацію та (за потреби) конкретний дозвіл
    користувача (permission) із таблиці users.permissions (JSON-список).

    - Якщо permission=None: достатньо просто бути залогіненим.
    - Якщо permission вказано: потрібен is_admin=1 АБО цей дозвіл
      присутній у списку permissions користувача.
    - Значення is_admin/permissions спочатку читаються з сесії (кеш),
      і лише якщо їх там ще немає - підвантажуються з БД один раз і
      кешуються в сесії (щоб не робити зайвий SELECT на кожен запит).

    Незалогінених користувачів редіректить на логін;
    користувачів без потрібного дозволу - на список студентів.
    """
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user_id' not in session:
                flash('Потрібна авторизація', 'danger')
                return redirect(url_for('auth.login'))

            # Спочатку перевіряємо сесію (кеш)
            is_admin = session.get('is_admin', False)
            perms = session.get('permissions', [])

            # Якщо немає в сесії, завантажуємо з БД.
            # ВИПРАВЛЕНО: раніше тут було `if not hasattr(session, 'is_admin')`.
            # Flask-об'єкт session - це словникоподібний контейнер, і
            # реального python-атрибута 'is_admin' на ньому ніколи не буває,
            # тож ця умова була ЗАВЖДИ істинною. Через це кеш у сесії ніколи
            # не спрацьовував і кожен захищений запит робив зайвий SELECT
            # до таблиці users. Правильна перевірка - через `in`.
            if 'is_admin' not in session:
                try:
                    conn = get_db()
                    try:
                        user = conn.execute("""
                            SELECT role, is_admin, permissions 
                            FROM users 
                            WHERE id = ?
                        """, (session['user_id'],)).fetchone()
                    finally:
                        conn.close()
                except Exception as e:
                    logger.error(f"Не вдалося перевірити права користувача (user_id={session.get('user_id')}): {e}", exc_info=True)
                    flash('Помилка бази даних під час перевірки прав доступу', 'danger')
                    return redirect(url_for('auth.login'))

                if not user:
                    session.clear()
                    return redirect(url_for('auth.login'))
                is_admin = bool(user['is_admin']) or (user['role'] == 'admin')
                perms = json.loads(user['permissions'] or '[]')
                session['is_admin'] = is_admin
                session['permissions'] = perms

            if permission is None:
                return f(*args, **kwargs)

            if is_admin or permission in perms:
                return f(*args, **kwargs)

            flash('Недостатньо прав', 'danger')
            return redirect(url_for('students.student_list'))  # або інший маршрут

        return decorated_function
    return decorator
  
def transliterate_ukrainian(text: str) -> str:
    """
    Транслітерація українського тексту згідно з Постановою КМУ №55-2010
    https://zakon.rada.gov.ua/laws/show/55-2010-п

    Використовується для автоматичного заповнення англомовних версій
    прізвища/імені студента при імпорті з Excel (див. generate_english_name).
    Повертає порожній рядок, якщо на вхід передано не-рядок або None.
    """

    if not text or not isinstance(text, str):
        return ""

    # базовая таблица
    base = {
        'а':'a','б':'b','в':'v','г':'h','ґ':'g','д':'d','е':'e',
        'ж':'zh','з':'z','и':'y','і':'i','й':'i','к':'k','л':'l',
        'м':'m','н':'n','о':'o','п':'p','р':'r','с':'s','т':'t',
        'у':'u','ф':'f','х':'kh','ц':'ts','ч':'ch','ш':'sh','щ':'shch',
        'ь':'','’':'','\'':'',

        'А':'A','Б':'B','В':'V','Г':'H','Ґ':'G','Д':'D','Е':'E',
        'Ж':'Zh','З':'Z','И':'Y','І':'I','Й':'I','К':'K','Л':'L',
        'М':'M','Н':'N','О':'O','П':'P','Р':'R','С':'S','Т':'T',
        'У':'U','Ф':'F','Х':'Kh','Ц':'Ts','Ч':'Ch','Ш':'Sh','Щ':'Shch',
        'Ь':''
    }

    # йотированные
    special_start = {
        'є':'ye','ї':'yi','ю':'yu','я':'ya',
        'Є':'Ye','Ї':'Yi','Ю':'Yu','Я':'Ya'
    }

    special_other = {
        'є':'ie','ї':'i','ю':'iu','я':'ia',
        'Є':'Ie','Ї':'I','Ю':'Iu','Я':'Ia'
    }

    result = []
    i = 0
    length = len(text)

    while i < length:

        # правило зг
        if i + 1 < length and text[i:i+2].lower() == "зг":
            if text[i].isupper():
                result.append("Zgh")
            else:
                result.append("zgh")
            i += 2
            continue

        char = text[i]

        # начало слова
        start = i == 0 or text[i-1] in " -'’"

        # йотированные
        if char in special_start:
            if start:
                result.append(special_start[char])
            else:
                result.append(special_other[char])
        else:
            result.append(base.get(char, char))

        i += 1

    return "".join(result)

# Пример использования для генерации полного имени
def generate_english_name(last_name_ua, first_name_ua):
    """
    Транслітерує українські прізвище та ім'я в англійський варіант
    (використовуючи transliterate_ukrainian). Викликається при імпорті
    студентів з Excel, коли англомовних ПІБ немає у файлі.

    Повертає кортеж (last_name_eng, first_name_eng).
    """
    last_name_eng = transliterate_ukrainian(last_name_ua)
    first_name_eng = transliterate_ukrainian(first_name_ua)
    return last_name_eng, first_name_eng


TEMPLATE_FOLDER = os.path.join(os.getcwd(), 'template_word')


def get_available_templates():
    """
    Повертає список шляхів (у форматі 'template_word/файл.docx') до всіх
    .docx-шаблонів у папці template_word - вони показуються користувачу
    як варіанти шаблону при масовій генерації документів
    (admin.group_export / admin.generate_group_docs).

    Якщо папки template_word не існує - повертає порожній список
    (а не кидає помилку).
    """
    if not os.path.isdir(TEMPLATE_FOLDER):
        return []
    files = [
        f for f in os.listdir(TEMPLATE_FOLDER)
        if f.lower().endswith('.docx') and os.path.isfile(os.path.join(TEMPLATE_FOLDER, f))
    ]
    files.sort()
    return [f"template_word/{f}" for f in files]