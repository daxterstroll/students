from functools import wraps
from flask import session, redirect, url_for
import logging
import os
from db import get_db

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
    """
    Декоратор для ограничения доступа к маршрутам.
    
    Аргументы:
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
    
    
def transliterate_ukrainian(text):
    """Транслитерация украинского текста по правилам Постановления №55-2010."""
    if not text or not isinstance(text, str):
        return ""

    # Правила транслитерации согласно Постановлению №55-2010
    translit_rules = {
        'а': 'a', 'б': 'b', 'в': 'v', 'г': 'h', 'ґ': 'g',
        'д': 'd', 'е': 'e', 'є': 'ye', 'ж': 'zh', 'з': 'z',
        'и': 'y', 'і': 'i', 'ї': 'yi', 'й': 'y', 'к': 'k',
        'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p',
        'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f',
        'х': 'kh', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'shch',
        'ь': '', 'ю': 'yu', 'я': 'ya', 'є': 'ie', 'ї': 'i',
        'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'H', 'Ґ': 'G',
        'Д': 'D', 'Е': 'E', 'Є': 'Ye', 'Ж': 'Zh', 'З': 'Z',
        'И': 'Y', 'І': 'I', 'Ї': 'Yi', 'Й': 'Y', 'К': 'K',
        'Л': 'L', 'М': 'M', 'Н': 'N', 'О': 'O', 'П': 'P',
        'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U', 'Ф': 'F',
        'Х': 'Kh', 'Ц': 'Ts', 'Ч': 'Ch', 'Ш': 'Sh', 'Щ': 'Shch',
        'Ь': '', 'Ю': 'Yu', 'Я': 'Ya', 'Є': 'Ie', 'Ї': 'I'
    }

    result = ''
    i = 0
    while i < len(text):
        char = text[i]
        if i + 1 < len(text):
            # Проверяем сочетания для особых случаев (например, 'зг' -> 'zgh')
            bigram = text[i:i+2].lower()
            if bigram in {'зг': 'zgh', 'ЗГ': 'Zgh'}:
                result += translit_rules.get(bigram[0], bigram[0]) + 'gh'
                i += 2
                continue
        # Одиночный символ
        result += translit_rules.get(char, char)
        i += 1

    return result

# Пример использования для генерации полного имени
def generate_english_name(last_name_ua, first_name_ua):
    last_name_eng = transliterate_ukrainian(last_name_ua)
    first_name_eng = transliterate_ukrainian(first_name_ua)
    return last_name_eng, first_name_eng