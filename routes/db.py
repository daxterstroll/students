"""
routes/db.py
============
Робота з підключенням до SQLite, кастомною колацією (правила сортування)
для українських рядків (`COLLATE UKRAINIAN`) та прозорим перемиканням між
робочою базою (`students.db`) і тестовою копією (`students_test_sandbox.db`).
"""

import json
import locale
import logging
import os
import sqlite3
import threading
import time

logger = logging.getLogger('Students')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REAL_DB_PATH = os.path.join(BASE_DIR, 'students.db')
TEST_DB_PATH = os.path.join(BASE_DIR, 'students_test_sandbox.db')
TEST_STATE_PATH = os.path.join(BASE_DIR, 'test_mode_state.json')

# --------------------------------------------------------------------------
# Кеш стану тестового режиму (за mtime_ns + size файлу-прапора), щоб
# get_db() не читав JSON на кожному запиті без потреби.
# --------------------------------------------------------------------------
_TEST_STATE_LOCK = threading.Lock()
_TEST_STATE_CACHE = {
    'sig': None,
    'state': {'active': False},
}

# --------------------------------------------------------------------------
# Налаштування української локалі виконується ОДИН РАЗ при імпорті модуля,
# а не при кожному порівнянні рядків під час сортування.
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
    Функція колації для SQLite (`COLLATE UKRAINIAN`).
    Порівнює два рядки за правилами української локалі, якщо вона доступна.
    """
    if _UA_COLLATION_AVAILABLE:
        try:
            return locale.strcoll(str1, str2)
        except locale.Error:
            pass
    return -1 if str1 < str2 else (1 if str1 > str2 else 0)


def normalize_name_for_match(s):
    """Нормалізує рядок ФІО для нечіткого порівняння в SQL."""
    if s is None:
        return ''
    s = str(s).lower()
    for ch in ('\u2019', '\u02BC', '\u2018', '`', '\u02BB', '\u2032', '\u055A', '\uFF07'):
        s = s.replace(ch, "'")
    for ch in ('\u2014', '\u2013', '\u2012', '\u2212', '\uFF0D', '\u2011', '-'):
        s = s.replace(ch, ' ')
    s = ' '.join(s.split())
    return s.strip()


def _invalidate_test_state_cache():
    with _TEST_STATE_LOCK:
        _TEST_STATE_CACHE['sig'] = None


def get_test_mode_state() -> dict:
    """
    Повертає актуальний стан тестового режиму:
      {'active': bool, 'started_at': str, 'started_by': str, ...}
    Якщо файл-прапор відсутній або сама копія БД не існує — active=False.
    """
    with _TEST_STATE_LOCK:
        try:
            st = os.stat(TEST_STATE_PATH)
            sig = (st.st_mtime_ns, st.st_size, os.path.exists(TEST_DB_PATH))
        except OSError:
            sig = ('missing', os.path.exists(TEST_DB_PATH))

        if _TEST_STATE_CACHE['sig'] == sig:
            return dict(_TEST_STATE_CACHE['state'])

        state = {'active': False}
        if os.path.exists(TEST_STATE_PATH) and os.path.exists(TEST_DB_PATH):
            try:
                with open(TEST_STATE_PATH, 'r', encoding='utf-8') as f:
                    raw = json.load(f)
                if isinstance(raw, dict) and raw.get('active'):
                    state = dict(raw)
                    state['active'] = True
            except Exception as e:
                logger.warning(f"Не вдалося прочитати test_mode_state.json: {e}")

        _TEST_STATE_CACHE['sig'] = sig
        _TEST_STATE_CACHE['state'] = state
        return dict(state)


def is_test_mode_active() -> bool:
    """True, якщо зараз увімкнено тестовий режим (пісочниця БД)."""
    return bool(get_test_mode_state().get('active'))


def set_test_mode_state(state: dict) -> None:
    """Атомарно зберігає файл-прапор тестового режиму."""
    payload = dict(state or {})
    payload['active'] = True
    tmp_path = f"{TEST_STATE_PATH}.tmp.{os.getpid()}"
    with open(tmp_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, TEST_STATE_PATH)
    _invalidate_test_state_cache()


def clear_test_mode_state() -> None:
    """Вимикає тестовий режим (видаляє файл-прапор)."""
    for p in (TEST_STATE_PATH,):
        try:
            if os.path.exists(p):
                os.remove(p)
        except OSError as e:
            logger.warning(f"Не вдалося видалити {p}: {e}")
    _invalidate_test_state_cache()


def cleanup_test_db_files(retries: int = 8) -> bool:
    """
    Видаляє файли тимчасової тестової БД (включно з -wal і -shm).
    На Windows файл може утримуватись долі секунди потоком, що закривається,
    тому робимо кілька коротких спроб.
    """
    if is_test_mode_active():
        return False

    targets = [TEST_DB_PATH, f"{TEST_DB_PATH}-wal", f"{TEST_DB_PATH}-shm"]
    all_gone = True
    for path in targets:
        if not os.path.exists(path):
            continue
        removed = False
        for attempt in range(max(1, retries)):
            try:
                os.remove(path)
                removed = True
                break
            except OSError:
                if attempt + 1 < retries:
                    time.sleep(0.1)
        if not removed and os.path.exists(path):
            all_gone = False
            logger.warning(f"Тимчасовий файл тестової БД ще зайнятий: {path}")
    return all_gone


def get_active_db_path(force_real: bool = False) -> str:
    """Повертає шлях до активної SQLite БД (реальна або тестова пісочниця)."""
    if force_real:
        return REAL_DB_PATH
    if is_test_mode_active() and os.path.exists(TEST_DB_PATH):
        return TEST_DB_PATH
    return REAL_DB_PATH


def _configure_sqlite_connection(conn: sqlite3.Connection) -> sqlite3.Connection:
    conn.row_factory = sqlite3.Row
    conn.create_collation("UKRAINIAN", ukrainian_collation)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.create_function("LOWER_UA", 1, lambda s: s.lower() if s else s)
    conn.create_function("NORM_NAME", 1, normalize_name_for_match)

    try:
        from routes.passport_types import normalize_passport_number
        conn.create_function(
            "NORM_PASSPORT", 2,
            lambda n, t: (normalize_passport_number(n, t) or '')
        )
    except Exception as _e:
        logger.warning(f"Не вдалося зареєструвати SQL-функцію NORM_PASSPORT: {_e}")

    try:
        from routes.passport_types import (
            normalize_passport_series, normalize_doc_number_letters,
        )
        conn.create_function(
            "NORM_SERIES", 1,
            lambda s: (normalize_passport_series(s) or '')
        )
        conn.create_function(
            "NORM_DOC_NUM", 1,
            lambda s: (normalize_doc_number_letters(s) or '')
        )
    except Exception as _e:
        logger.warning(f"Не вдалося зареєструвати SQL-функції NORM_SERIES/NORM_DOC_NUM: {_e}")

    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA busy_timeout = 30000")
    return conn


def get_db(force_real: bool = False):
    """
    Створює нове з'єднання з базою даних SQLite.
    Якщо увімкнено тестовий режим (і не передано force_real=True),
    автоматично підключається до тестової копії `students_test_sandbox.db`.
    """
    db_path = get_active_db_path(force_real=force_real)
    conn = sqlite3.connect(db_path, timeout=30)
    return _configure_sqlite_connection(conn)


def get_real_db():
    """Завжди підключається до реальної БД `students.db` (для резервних копій і сесій)."""
    return get_db(force_real=True)
