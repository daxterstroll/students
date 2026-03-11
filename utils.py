from functools import wraps
from flask import session, redirect, url_for, flash
import logging
import os
from db import get_db
import json

# Настройка пути к файлу
project_root = os.path.dirname(os.path.abspath(__file__))
log_file_path = os.path.join(project_root, 'app.log')

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

def log_action(username, action, group_ids=None, mode=None):
    """Логирование действий пользователя."""
    conn = get_db()
    role = session.get('role')  # Получаем роль пользователя из сессии

    # Определяем строку с группами только если есть group_ids и роль не admin
    group_names_str = ''
    if group_ids is not None and role != 'admin':
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

    conn.close()

    # Формируем лог с учетом режима, если он передан
    if mode:
        logger.info(f"👤 {username} - {action} (режим: {mode})")
    elif group_names_str:
        logger.info(f"👤 {username} - {action} (групи: {group_names_str})")
    else:
        logger.info(f"👤 {username} - {action}")

def login_required(role=None):
    """Декоратор для проверки авторизации и роли пользователя (стара версія для сумісності).
    Args:
        role (str, optional): Требуемая роль пользователя (например, 'admin').
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
    """Розширений декоратор для перевірки дозволів.
    - Якщо permission=None: тільки перевірка логіну.
    - Якщо permission вказано: потрібен is_admin=1 або цей дозвіл в permissions.
    - Зворотна сумісність з role='admin'.
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

            # Якщо немає в сесії, завантажуємо з БД
            if not hasattr(session, 'is_admin'):
                conn = get_db()
                user = conn.execute("""
                    SELECT role, is_admin, permissions 
                    FROM users 
                    WHERE id = ?
                """, (session['user_id'],)).fetchone()
                conn.close()
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
    
    
import re

def transliterate_ukrainian(text: str) -> str:
    """
    Транслитерация украинского текста согласно Постановлению КМУ №55-2010
    https://zakon.rada.gov.ua/laws/show/55-2010-п
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
    last_name_eng = transliterate_ukrainian(last_name_ua)
    first_name_eng = transliterate_ukrainian(first_name_ua)
    return last_name_eng, first_name_eng