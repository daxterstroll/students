"""
routes/utils.py
================
Спільні утиліти: логер додатку, логування дій користувачів (log_action),
декоратори перевірки авторизації/прав (login_required, permission_required),
транслітерація українських імен та пошук шаблонів .docx.

Логування: усі модулі, які хочуть писати помилки/події в лог,
ПОВИННІ використовувати саме логер `logger`, визначений нижче
(`from routes.utils import logger`), а не голий `import logging`.
Лише `logger` тут налаштований з файловим обробником та консольним
виводом; звичайний кореневий логер Python (`logging.error(...)`)
обробників не має і в файл нічого не пише.

Логи лежать у папці logs/ окремим файлом на кожен місяць
(logs/app-2026-10.log і т.д.) - раніше все писалось в один app.log,
який ріс необмежено і який /admin/view_logs доводилось повністю
вичитувати в пам'ять на кожен запит. MonthlyFileHandler нижче сам
перемикається на новий файл, коли починається новий місяць; стара
накопичена app.log (якщо лишилась з попередньої версії) розбивається
на місячні файли один раз при старті - див. _migrate_legacy_log().
"""

from functools import wraps
from flask import session, redirect, url_for, flash, g, request, has_request_context
import logging
import os
from datetime import datetime
from routes.db import get_db
import json
import re
import uuid

# Настройка пути к файлу
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGS_DIR = os.path.join(PROJECT_ROOT, 'logs')
os.makedirs(LOGS_DIR, exist_ok=True)

LEGACY_LOG_PATH = os.path.join(PROJECT_ROOT, 'app.log')
_LOG_LINE_DATE_RE = re.compile(r'^(\d{4})-(\d{2})-\d{2}\s*\|')


def monthly_log_path(dt=None):
    """Шлях до файлу логів конкретного місяця (за замовчуванням - поточного)."""
    dt = dt or datetime.now()
    return os.path.join(LOGS_DIR, f"app-{dt.strftime('%Y-%m')}.log")

def title_case_ukrainian(text):
    """
    Нормалізує ПІБ до вигляду «Прізвище Ім'я По-батькові»
    (перша літера кожного слова — велика, решта — маленькі).

    Коректно обробляє:
      - апострофи: В'ЯЧЕСЛАВ → В'ячеслав (після апострофа — маленька)
      - дефіси: АННА-МАРІЯ → Анна-Марія (після дефіса — велика)
    """
    if not text:
        return ''

    def _cap_word(word):
        if not word:
            return word
        result = []
        cap_next = True
        for ch in word:
            if ch in ("'", '’', 'ʼ', '`'):
                result.append(ch)
                cap_next = False
            elif ch == '-':
                result.append(ch)
                cap_next = True
            elif cap_next:
                result.append(ch.upper())
                cap_next = False
            else:
                result.append(ch.lower())
        return ''.join(result)

    return ' '.join(_cap_word(w) for w in text.split())


def _migrate_legacy_log():
    """
    Одноразова міграція: якщо з попередньої версії лишився старий
    суцільний app.log у корені проєкту - розбиває його рядки по
    місяцях (за датою на початку кожного рядка) у нові файли в logs/,
    а сам старий файл перейменовує в app.log.bak (не видаляє - про
    всяк випадок). Якщо app.log.bak вже існує, нічого не робить -
    міграція вже відбулась раніше.
    """
    if not os.path.exists(LEGACY_LOG_PATH):
        return
    backup_path = LEGACY_LOG_PATH + '.bak'
    if os.path.exists(backup_path):
        return
    try:
        buckets = {}
        with open(LEGACY_LOG_PATH, 'r', encoding='utf-8') as f:
            for line in f:
                m = _LOG_LINE_DATE_RE.match(line)
                year_month = f"{m.group(1)}-{m.group(2)}" if m else 'unknown'
                buckets.setdefault(year_month, []).append(line)

        for year_month, lines in buckets.items():
            target = os.path.join(LOGS_DIR, f"app-{year_month}.log")
            # Дописуємо в кінець - якщо файл цього місяця вже існує
            # (малоймовірно для першого запуску, але про всяк випадок),
            # старі записи не губляться.
            with open(target, 'a', encoding='utf-8') as out:
                out.writelines(lines)

        os.rename(LEGACY_LOG_PATH, backup_path)
    except Exception as e:
        # Логер ще не налаштований на цьому етапі - друкуємо прямо в консоль.
        print(f"Не вдалося розбити старий app.log по місяцях: {e}")


_migrate_legacy_log()


class MonthlyFileHandler(logging.FileHandler):
    """
    FileHandler, що сам перемикається на файл нового місяця щойно
    починається новий місяць - без цього довелось би перезапускати
    застосунок першого числа, щоб лог почав писатися в новий файл.
    """
    def __init__(self, encoding=None):
        self._current_year_month = datetime.now().strftime('%Y-%m')
        super().__init__(monthly_log_path(), encoding=encoding)

    def emit(self, record):
        now_year_month = datetime.now().strftime('%Y-%m')
        if now_year_month != self._current_year_month:
            self._current_year_month = now_year_month
            if self.stream:
                self.stream.close()
                self.stream = None
            self.baseFilename = monthly_log_path()
            self.stream = self._open()
        super().emit(record)


# Настройка глобального логгера
logger = logging.getLogger('Students')
logger.setLevel(logging.DEBUG)

# Удаляем все существующие обработчики, чтобы избежать конфликтов
if logger.handlers:
    logger.handlers.clear()

# Создаем обработчики
file_handler = MonthlyFileHandler(encoding='utf-8')
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
    with open(monthly_log_path(), 'a', encoding='utf-8'):
        pass
    logger.debug(f"Логирование инициализировано. Файл логов: {monthly_log_path()}")
except Exception as e:
    logger.error(f"Ошибка при доступе к файлу логов {monthly_log_path()}: {e}")
    print(f"Ошибка при доступе к файлу логов: {e}")

_AUDIT_SCHEMA_ENSURED = False

AUDIT_CATEGORY_UA = {
    'auth': 'Вхід / вихід',
    'security': 'Безпека',
    'create': 'Створення',
    'update': 'Редагування',
    'delete': 'Видалення',
    'permissions': 'Права та доступ',
    'sessions': 'Сесії та оцінки',
    'orders': 'Накази',
    'export': 'Генерація / експорт',
    'import': 'Імпорт даних',
    'view': 'Перегляд',
    'system': 'Система',
}

AUDIT_SECTION_BY_PREFIX = [
    ('/login', 'Авторизація'),
    ('/logout', 'Авторизація'),
    ('/admin/users', 'Користувачі та права'),
    ('/admin/view_logs', 'Журнал дій'),
    ('/admin/download_log', 'Журнал дій'),
    ('/admin/export_logs', 'Журнал дій'),
    ('/admin/manage_templates', 'Шаблони документів'),
    ('/admin/manage_education_documents', 'Документи про освіту'),
    ('/admin/import_docs', 'Документи про освіту'),
    ('/admin/manage_passport_documents', 'Паспортні дані'),
    ('/admin/import_passport', 'Паспортні дані'),
    ('/admin/manage_faculties', 'Довідники'),
    ('/admin/manage_specialties', 'Довідники'),
    ('/admin/manage_degree_levels', 'Довідники'),
    ('/admin/manage_licenses', 'Довідники'),
    ('/admin/manage_educational_programs', 'Довідники'),
    ('/admin/manage_qualification_names', 'Довідники'),
    ('/admin/manage_accreditations', 'Довідники'),
    ('/admin/orders', 'Накази'),
    ('/admin/enrollment_order', 'Накази'),
    ('/admin/expulsion', 'Накази'),
    ('/admin/renewal', 'Накази'),
    ('/admin/pending_students', 'Заявки на реєстрацію'),
    ('/admin/update_requests', 'Заявки на оновлення'),
    ('/admin/study_periods', 'Періоди навчання'),
    ('/admin/diplomas', 'Дипломи та додатки'),
    ('/admin/groups', 'Групи та курси'),
    ('/admin/courses', 'Групи та курси'),
    ('/admin/archive', 'Групи та курси'),
    ('/admin/subjects', 'Предмети та активності'),
    ('/admin/activities', 'Предмети та активності'),
    ('/admin/group_export', 'Генерація та експорт'),
    ('/admin/custom_export', 'Генерація та експорт'),
    ('/admin/license_report', 'Звіт по ліцензіях'),
    ('/admin/export_photos', 'Фото студентів'),
    ('/import_grades', 'Імпорт оцінок'),
    ('/military-reports', 'Військовий облік'),
    ('/session-grade', 'Сесії та відомості'),
    ('/analytics', 'Аналітика'),
    ('/office', 'Редактор документів'),
    ('/photo', 'Фото студентів'),
    ('/apply', 'Публічні форми'),
    ('/update-info', 'Публічні форми'),
    ('/grades/', 'Оцінки та план'),
    ('/activities/', 'Оцінки та план'),
    ('/students', 'Студенти'),
]


def ensure_audit_logs_table(conn):
    """Ідемпотентно створює таблицю audit_logs та індекси для детального аудиту."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            log_uid TEXT,
            created_at TEXT NOT NULL,
            level TEXT NOT NULL DEFAULT 'INFO',
            username TEXT,
            user_id INTEGER,
            user_role TEXT,
            category TEXT,
            action_type TEXT,
            section TEXT,
            action_text TEXT NOT NULL,
            entity_type TEXT,
            entity_id TEXT,
            entity_name TEXT,
            entity_url TEXT,
            group_ids_json TEXT,
            group_names TEXT,
            mode TEXT,
            changes_json TEXT,
            details_json TEXT,
            details_text TEXT,
            route_path TEXT,
            endpoint TEXT,
            http_method TEXT,
            ip_address TEXT,
            user_agent TEXT
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_logs_created_at ON audit_logs(created_at)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_logs_username ON audit_logs(username)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_logs_category ON audit_logs(category)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_logs_section ON audit_logs(section)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_logs_level ON audit_logs(level)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_logs_uid ON audit_logs(log_uid)")


def _extract_client_ip():
    if not has_request_context():
        return None
    from routes.security import client_ip
    ip = client_ip()
    return None if not ip or ip == 'unknown' else ip


def _short_user_agent(ua_raw):
    if not ua_raw:
        return None
    ua = str(ua_raw).strip()
    if not ua:
        return None
    browser = 'Інший браузер'
    for pattern, name in (
        (r'Edg/(\d+)', 'Edge'),
        (r'OPR/(\d+)', 'Opera'),
        (r'Chrome/(\d+)', 'Chrome'),
        (r'Firefox/(\d+)', 'Firefox'),
        (r'Version/(\d+).*Safari', 'Safari'),
        (r'curl/([\d.]+)', 'curl'),
        (r'python-requests/([\d.]+)', 'Python'),
    ):
        m = re.search(pattern, ua)
        if m:
            browser = f"{name} {m.group(1)}"
            break
    os_name = ''
    if 'Windows NT 10' in ua or 'Windows NT 11' in ua:
        os_name = 'Windows'
    elif 'Windows' in ua:
        os_name = 'Windows'
    elif 'Android' in ua:
        os_name = 'Android'
    elif 'iPhone' in ua or 'iPad' in ua:
        os_name = 'iOS'
    elif 'Mac OS X' in ua or 'Macintosh' in ua:
        os_name = 'macOS'
    elif 'Linux' in ua:
        os_name = 'Linux'
    return f"{browser} · {os_name}" if os_name else browser


def _normalize_diff_val(val, empty_placeholder='—'):
    if val is None:
        return empty_placeholder
    if isinstance(val, bool):
        return 'Так' if val else 'Ні'
    s = str(val).strip()
    return s if s else empty_placeholder


def build_field_changes(old_data, new_data, field_labels, *, formatters=None, empty_placeholder='—'):
    """
    Порівнює словники/sqlite3.Row (old_data vs new_data) за картою підписів field_labels
    і повертає список змін у форматі:
      [{'field': key, 'label': '...', 'old': '...', 'new': '...'}, ...]
    """
    old_dict = dict(old_data) if old_data is not None else {}
    new_dict = dict(new_data) if new_data is not None else {}
    formatters = formatters or {}
    changes = []

    for field, label in field_labels.items():
        if field not in new_dict:
            continue
        raw_old = old_dict.get(field)
        raw_new = new_dict.get(field)

        norm_old = '' if raw_old is None else str(raw_old).strip()
        norm_new = '' if raw_new is None else str(raw_new).strip()
        if isinstance(raw_old, bool):
            norm_old = '1' if raw_old else '0'
        if isinstance(raw_new, bool):
            norm_new = '1' if raw_new else '0'

        if norm_old == norm_new:
            continue

        fmt = formatters.get(field)
        if callable(fmt):
            disp_old = _normalize_diff_val(fmt(raw_old), empty_placeholder)
            disp_new = _normalize_diff_val(fmt(raw_new), empty_placeholder)
        else:
            disp_old = _normalize_diff_val(raw_old, empty_placeholder)
            disp_new = _normalize_diff_val(raw_new, empty_placeholder)

        if disp_old == disp_new:
            continue

        changes.append({
            'field': field,
            'label': label,
            'old': disp_old,
            'new': disp_new,
        })
    return changes


def _normalize_changes_list(changes):
    if not changes:
        return []
    out = []
    for item in changes:
        if isinstance(item, dict):
            label = str(item.get('label') or item.get('field') or 'Поле').strip()
            field = str(item.get('field') or label).strip()
            old_v = _normalize_diff_val(item.get('old'))
            new_v = _normalize_diff_val(item.get('new'))
            out.append({'field': field, 'label': label, 'old': old_v, 'new': new_v})
        elif isinstance(item, (tuple, list)) and len(item) >= 3:
            label = str(item[0]).strip()
            out.append({
                'field': label,
                'label': label,
                'old': _normalize_diff_val(item[1]),
                'new': _normalize_diff_val(item[2]),
            })
    return out


_INLINE_ARROW_RE = re.compile(r'([^;|\n]+?)\s*(?:→|->)\s*([^;|\n]+)')


def parse_inline_changes_from_text(text):
    """Витягує прості дифи «було → стало» із текстового рядка details."""
    if not text:
        return []
    changes = []
    for part in re.split(r'[;\n]+', str(text)):
        part = part.strip()
        if not part or ('→' not in part and '->' not in part):
            continue
        m = _INLINE_ARROW_RE.search(part)
        if not m:
            continue
        left = m.group(1).strip()
        right = m.group(2).strip()
        label = 'Значення'
        if ':' in left and not left.lower().startswith('http'):
            maybe_lbl, maybe_old = left.split(':', 1)
            if 1 <= len(maybe_lbl.strip()) <= 45:
                label = maybe_lbl.strip()
                left = maybe_old.strip()
        if right.lower().startswith('нова група:'):
            right = right.split(':', 1)[1].strip()
        if left.lower().startswith('стара група:'):
            label = 'Група'
            left = left.split(':', 1)[1].strip()
        if left or right:
            changes.append({
                'field': label.lower().replace(' ', '_'),
                'label': label,
                'old': left or '—',
                'new': right or '—',
            })
    return changes


def infer_audit_metadata(action_text, details_text='', route_path='', endpoint='', view_args=None):
    """Визначає section, category, action_type, entity_type, entity_id, entity_url за контекстом."""
    act = (action_text or '').strip()
    act_low = act.lower()
    path = (route_path or '').strip()
    ep = (endpoint or '').strip()
    view_args = view_args or {}

    section = None
    for prefix, sec_name in AUDIT_SECTION_BY_PREFIX:
        if path.startswith(prefix):
            section = sec_name
            break
    if not section:
        if 'студент' in act_low:
            section = 'Студенти'
        elif 'паспорт' in act_low:
            section = 'Паспортні дані'
        elif 'освіт' in act_low and 'документ' in act_low:
            section = 'Документи про освіту'
        elif 'військ' in act_low or 'тцк' in act_low or 'додаток' in act_low or 'резерв+' in act_low:
            section = 'Військовий облік'
        elif 'наказ' in act_low:
            section = 'Накази'
        elif 'сесі' in act_low or 'відомост' in act_low or 'викладач' in act_low:
            section = 'Сесії та відомості'
        elif 'користувач' in act_low or 'пароль' in act_low:
            section = 'Користувачі та права'
        elif 'шаблон' in act_low:
            section = 'Шаблони документів'
        elif 'лог' in act_low or 'журнал' in act_low:
            section = 'Журнал дій'
        elif 'ввійшов' in act_low or 'вийшов' in act_low or 'входу' in act_low:
            section = 'Авторизація'
        elif 'оцінк' in act_low or 'активност' in act_low:
            section = 'Оцінки та план'
        elif 'факультет' in act_low or 'спеціальн' in act_low or 'ступін' in act_low or 'ліцензі' in act_low or 'кваліфікац' in act_low or 'програм' in act_low or 'галуз' in act_low:
            section = 'Довідники'
        else:
            section = 'Система'

    # action_type & category
    action_type = 'other'
    category = 'system'
    if 'невдала спроба входу' in act_low or 'блокування спроб входу' in act_low or 'rate limit' in act_low:
        action_type = 'rate_limit' if 'блокуван' in act_low or 'rate' in act_low else 'login_failed'
        category = 'security'
    elif 'ввійшов у систему' in act_low:
        action_type = 'login'
        category = 'auth'
    elif 'вийшов із системи' in act_low:
        action_type = 'logout'
        category = 'auth'
    elif 'переглянув' in act_low:
        action_type = 'view'
        category = 'view'
    elif 'видалив' in act_low:
        action_type = 'delete'
        category = 'delete'
    elif 'відхилив' in act_low:
        action_type = 'reject'
        category = 'update'
    elif 'імпорт' in act_low or 'пакетний імпорт' in act_low:
        action_type = 'import'
        category = 'import'
    elif 'згенерував' in act_low or 'вивантажив' in act_low or 'завантажив лог' in act_low or 'експорт' in act_low:
        action_type = 'export'
        category = 'export'
    elif 'повторно відкрив' in act_low or 'розблокував' in act_low:
        action_type = 'reopen'
        category = 'sessions'
    elif 'затвердив' in act_low or 'підтвердив' in act_low or 'застосував' in act_low:
        action_type = 'approve'
        category = 'sessions' if ('відомост' in act_low or 'сесі' in act_low) else 'update'
    elif 'об\'єднав' in act_low or 'об’єднав' in act_low:
        action_type = 'merge'
        category = 'orders'
    elif 'скопіював' in act_low:
        action_type = 'copy'
        category = 'create'
    elif any(w in act_low for w in ('додав', 'створив', 'завантажив новий', 'зарахував', 'призначив', 'синхронізував')):
        action_type = 'create'
        category = 'permissions' if section == 'Користувачі та права' else 'create'
    elif any(w in act_low for w in ('редагував', 'оновив', 'змінив', 'прив\'язав', 'відв\'язав', 'вирішив', 'розархівував', 'переніс', 'перевів', 'увімкнув', 'вимкнув', 'активував', 'деактивував', 'приховав', 'зробив видимим', 'скинув', 'відрахував', 'випустив')):
        action_type = 'update'
        if section == 'Користувачі та права':
            category = 'permissions'
        elif section == 'Сесії та відомості':
            category = 'sessions'
        elif section == 'Накази':
            category = 'orders'
        else:
            category = 'update'

    if category in ('create', 'update', 'delete'):
        if section == 'Користувачі та права' and category != 'delete':
            category = 'permissions'

    # entity_type, entity_id, entity_url
    entity_type = None
    entity_id = None
    entity_url = None
    if 'student_id' in view_args and view_args['student_id']:
        entity_type = 'student'
        entity_id = str(view_args['student_id'])
    elif 'existing_id' in view_args and view_args['existing_id']:
        entity_type = 'student'
        entity_id = str(view_args['existing_id'])
    elif 'order_id' in view_args and view_args['order_id']:
        entity_type = 'order'
        entity_id = str(view_args['order_id'])
    elif 'sheet_id' in view_args and view_args['sheet_id']:
        entity_type = 'exam_sheet'
        entity_id = str(view_args['sheet_id'])
    elif 'session_id' in view_args and view_args['session_id']:
        entity_type = 'exam_session'
        entity_id = str(view_args['session_id'])
    elif 'user_id' in view_args and view_args['user_id']:
        entity_type = 'user'
        entity_id = str(view_args['user_id'])
    elif 'pending_id' in view_args and view_args['pending_id']:
        entity_type = 'pending_student'
        entity_id = str(view_args['pending_id'])
    elif 'request_id' in view_args and view_args['request_id']:
        entity_type = 'update_request'
        entity_id = str(view_args['request_id'])
    elif 'teacher_id' in view_args and view_args['teacher_id']:
        entity_type = 'teacher'
        entity_id = str(view_args['teacher_id'])
    elif 'group_id' in view_args and view_args['group_id']:
        entity_type = 'group'
        entity_id = str(view_args['group_id'])

    if not entity_id:
        m = re.search(r'(?:ID\s*|student=|#)(\d+)', act)
        if m:
            entity_id = m.group(1)

    if not entity_type:
        if 'студент' in act_low or 'student=' in act_low:
            entity_type = 'student'
        elif 'наказ' in act_low:
            entity_type = 'order'
        elif 'відомост' in act_low:
            entity_type = 'exam_sheet'
        elif 'сесі' in act_low:
            entity_type = 'exam_session'
        elif 'користувач' in act_low:
            entity_type = 'user'
        elif 'паспорт' in act_low:
            entity_type = 'passport_document'
        elif 'документ про освіту' in act_low:
            entity_type = 'education_document'
        elif 'шаблон' in act_low:
            entity_type = 'template'
        elif 'ліцензі' in act_low:
            entity_type = 'license'
        elif 'факультет' in act_low:
            entity_type = 'faculty'
        elif 'спеціальн' in act_low:
            entity_type = 'specialty'
        elif 'ступін' in act_low:
            entity_type = 'degree_level'
        elif 'освітн' in act_low and 'програм' in act_low:
            entity_type = 'educational_program'
        elif 'кваліфікац' in act_low:
            entity_type = 'qualification_name'
        elif 'тцк' in act_low:
            entity_type = 'tck'
        elif 'викладач' in act_low:
            entity_type = 'teacher'

    if action_type != 'delete' and entity_id and str(entity_id).isdigit():
        eid = int(entity_id)
        if entity_type == 'student':
            entity_url = f"/students/{eid}"
        elif entity_type == 'order':
            entity_url = f"/admin/orders/{eid}"
        elif entity_type == 'exam_session':
            entity_url = f"/session-grade/sessions/{eid}"
        elif entity_type == 'exam_sheet':
            entity_url = f"/session-grade/sheets/{eid}"
        elif entity_type == 'user':
            entity_url = f"/admin/users/{eid}/edit"
        elif entity_type == 'pending_student':
            entity_url = f"/admin/pending_students/{eid}"
        elif entity_type == 'update_request':
            entity_url = f"/admin/update_requests/{eid}"

    return {
        'section': section,
        'category': category,
        'action_type': action_type,
        'entity_type': entity_type,
        'entity_id': entity_id,
        'entity_url': entity_url,
    }


def log_action(
    username,
    action,
    group_ids=None,
    mode=None,
    details=None,
    *,
    level='INFO',
    category=None,
    action_type=None,
    section=None,
    entity_type=None,
    entity_id=None,
    entity_name=None,
    entity_url=None,
    changes=None,
    meta=None,
    ip_address=None,
    route_path=None,
    endpoint=None,
    http_method=None,
    user_agent=None,
    conn=None,
):
    """
    Записує дію користувача в структуровану таблицю `audit_logs` та у місячний
    лог-файл (logs/app-YYYY-MM.log). Повністю зворотно сумісна зі старими
    викликами log_action(username, action, group_ids=..., mode=..., details=...).
    """
    global _AUDIT_SCHEMA_ENSURED
    own_conn = None
    log_uid = uuid.uuid4().hex[:12]
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    level_norm = (str(level or 'INFO').upper() if level else 'INFO')
    if level_norm not in ('DEBUG', 'INFO', 'WARNING', 'ERROR'):
        level_norm = 'INFO'

    try:
        db = conn if conn is not None else get_db()
        if conn is None:
            own_conn = db

        in_req = has_request_context()
        if in_req:
            try:
                user_id = session.get('user_id')
                role = session.get('role') or ('admin' if session.get('is_admin') else None)
            except RuntimeError:
                user_id = None
                role = None
            req_path = route_path or request.path
            req_ep = endpoint or request.endpoint
            req_method = http_method or request.method
            req_ip = ip_address or _extract_client_ip()
            req_ua = user_agent or _short_user_agent(request.headers.get('User-Agent', ''))
            req_view_args = getattr(request, 'view_args', None) or {}
        else:
            user_id = None
            role = None
            req_path = route_path
            req_ep = endpoint
            req_method = http_method
            req_ip = ip_address
            req_ua = user_agent
            req_view_args = {}

        # Завжди визначаємо людські назви груп, якщо передано group_ids
        norm_gids = []
        group_names_str = ''
        if group_ids is not None:
            raw_list = group_ids if isinstance(group_ids, (list, tuple, set)) else ([group_ids] if group_ids else [])
            for g_item in raw_list:
                if g_item is None or g_item == '':
                    continue
                try:
                    norm_gids.append(int(g_item))
                except (TypeError, ValueError):
                    pass
            if norm_gids:
                placeholders = ','.join('?' for _ in norm_gids)
                group_rows = db.execute(
                    f"""
                    SELECT name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кр.)' AS display_name
                    FROM groups
                    WHERE id IN ({placeholders})
                    ORDER BY name, start_year
                    """,
                    norm_gids,
                ).fetchall()
                group_names_str = ', '.join([r['display_name'] for r in group_rows]) if group_rows else ''

        details_str = ''
        if isinstance(details, dict):
            if meta is None:
                meta = details
            details_str = '; '.join(f"{k}: {v}" for k, v in details.items() if v not in (None, ''))
        elif details is not None:
            details_str = str(details).strip()

        inferred = infer_audit_metadata(
            action,
            details_str,
            route_path=req_path or '',
            endpoint=req_ep or '',
            view_args=req_view_args,
        )

        final_section = section or inferred['section']
        final_category = category or inferred['category']
        final_action_type = action_type or inferred['action_type']
        final_entity_type = entity_type or inferred['entity_type']
        final_entity_id = str(entity_id) if entity_id is not None else inferred['entity_id']
        final_entity_url = entity_url if entity_url is not None else inferred['entity_url']

        norm_changes = _normalize_changes_list(changes)
        if not norm_changes and details_str:
            norm_changes = parse_inline_changes_from_text(details_str)

        changes_json = json.dumps(norm_changes, ensure_ascii=False) if norm_changes else None
        meta_json = json.dumps(meta, ensure_ascii=False) if meta else None
        gids_json = json.dumps(norm_gids, ensure_ascii=False) if norm_gids else None

        # Запис у таблицю БД (не блокує файловий лог у разі винятку)
        try:
            if not _AUDIT_SCHEMA_ENSURED:
                ensure_audit_logs_table(db)
                _AUDIT_SCHEMA_ENSURED = True
            try:
                db.execute(
                    """
                    INSERT INTO audit_logs (
                        log_uid, created_at, level, username, user_id, user_role,
                        category, action_type, section, action_text,
                        entity_type, entity_id, entity_name, entity_url,
                        group_ids_json, group_names, mode,
                        changes_json, details_json, details_text,
                        route_path, endpoint, http_method, ip_address, user_agent
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        log_uid,
                        now_str,
                        level_norm,
                        str(username or 'невідомо'),
                        user_id,
                        role,
                        final_category,
                        final_action_type,
                        final_section,
                        str(action or ''),
                        final_entity_type,
                        final_entity_id,
                        str(entity_name) if entity_name else None,
                        final_entity_url,
                        gids_json,
                        group_names_str or None,
                        str(mode) if mode else None,
                        changes_json,
                        meta_json,
                        details_str or None,
                        req_path,
                        req_ep,
                        req_method,
                        req_ip,
                        req_ua,
                    ),
                )
            except Exception:
                ensure_audit_logs_table(db)
                _AUDIT_SCHEMA_ENSURED = True
                db.execute(
                    """
                    INSERT INTO audit_logs (
                        log_uid, created_at, level, username, user_id, user_role,
                        category, action_type, section, action_text,
                        entity_type, entity_id, entity_name, entity_url,
                        group_ids_json, group_names, mode,
                        changes_json, details_json, details_text,
                        route_path, endpoint, http_method, ip_address, user_agent
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        log_uid,
                        now_str,
                        level_norm,
                        str(username or 'невідомо'),
                        user_id,
                        role,
                        final_category,
                        final_action_type,
                        final_section,
                        str(action or ''),
                        final_entity_type,
                        final_entity_id,
                        str(entity_name) if entity_name else None,
                        final_entity_url,
                        gids_json,
                        group_names_str or None,
                        str(mode) if mode else None,
                        changes_json,
                        meta_json,
                        details_str or None,
                        req_path,
                        req_ep,
                        req_method,
                        req_ip,
                        req_ua,
                    ),
                )
            if own_conn is not None:
                own_conn.commit()
        except Exception as db_exc:
            logger.warning(f"audit_logs DB insert fallback ({username} / {action}): {db_exc}")

        # Людський запис у місячний текстовий файл .log
        log_msg = f"👤 {username} - {action}"
        if mode:
            log_msg += f" (режим: {mode})"
        if final_section:
            where_part = f"{final_section} · {req_path}" if req_path else final_section
            log_msg += f" (де: {where_part})"
        if group_names_str:
            log_msg += f" (групи: {group_names_str})"
        if req_ip and 'ip:' not in (details_str or '').lower():
            log_msg += f" (IP: {req_ip})"
        if norm_changes:
            diff_preview = '; '.join(
                f"{c['label']}: «{c['old']}» → «{c['new']}»" for c in norm_changes[:8]
            )
            if len(norm_changes) > 8:
                diff_preview += f" (+ще {len(norm_changes) - 8})"
            log_msg += f" | Зміни: {diff_preview}"
        if details_str:
            log_msg += f" | {details_str}"
        log_msg += f" [ref:{log_uid}]"

        if level_norm == 'ERROR':
            logger.error(log_msg)
        elif level_norm == 'WARNING':
            logger.warning(log_msg)
        elif level_norm == 'DEBUG':
            logger.debug(log_msg)
        else:
            logger.info(log_msg)

    except Exception as e:
        # Логування НІКОЛИ не повинно "ронити" запит користувача, що його викликав.
        logger.error(f"Не вдалося записати дію в лог (username={username}, action={action}): {e}", exc_info=True)
    finally:
        if own_conn is not None:
            own_conn.close()
    
def _sync_session_auth():
    """
    Синхронізує role, is_admin і permissions поточної сесії з БД один раз
    за HTTP-запит (кеш у flask.g). Якщо користувача видалили — очищає
    сесію і повертає False.
    """
    user_id = session.get('user_id')
    if not user_id:
        return False
    if getattr(g, '_auth_synced', False):
        return True

    try:
        conn = get_db()
        try:
            user = conn.execute(
                """
                SELECT role, is_admin, permissions
                FROM users
                WHERE id = ?
                """,
                (user_id,),
            ).fetchone()
        finally:
            conn.close()
    except Exception as e:
        logger.error(f"Не вдалося перевірити права користувача (user_id={user_id}): {e}", exc_info=True)
        return None

    if not user:
        session.clear()
        return False

    try:
        perms = json.loads(user['permissions'] or '[]')
    except Exception:
        perms = []
    if not isinstance(perms, list):
        perms = []

    is_admin = bool(user['is_admin']) or (user['role'] == 'admin')
    role = 'admin' if is_admin else (user['role'] or 'user')
    session['role'] = role
    session['is_admin'] = is_admin
    session['permissions'] = perms
    g._auth_synced = True
    return True


def login_required(role=None):
    """
    Декоратор: пускає на сторінку тільки авторизаних користувачів.
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
            synced = _sync_session_auth()
            if synced is None:
                flash('Помилка бази даних під час перевірки прав доступу', 'danger')
                return redirect(url_for('auth.login'))
            if not synced:
                return redirect(url_for('auth.login'))
            if role and session.get('role') != role:
                return "403 Forbidden", 403
            return f(*args, **kwargs)
        return decorated_function
    return decorator

def permission_required(*permissions):
    """
    Декоратор: перевіряє авторизацію та (за потреби) наявність ХОЧА Б ОДНОГО
    із зазначених дозволів користувача із таблиці users.permissions (JSON-список).

    - Якщо дозволів не передано або передано лише None: достатньо бути залогіненим.
    - Якщо передано один або кілька дозволів (напр. @permission_required('manage_students', 'manage_military_reports')):
      потрібен is_admin=1 АБО наявність хоча б одного з цих дозволів у permissions користувача.
    - Права синхронізуються з БД один раз на HTTP-запит (через _sync_session_auth),
      тому відкликання доступу чи видалення користувача діє негайно.

    Незалогінених користувачів редіректить на логін;
    користувачів без потрібного дозволу - на список студентів.
    """
    required_perms = tuple(p for p in permissions if p)

    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user_id' not in session:
                flash('Потрібна авторизація', 'danger')
                return redirect(url_for('auth.login'))

            synced = _sync_session_auth()
            if synced is None:
                flash('Помилка бази даних під час перевірки прав доступу', 'danger')
                return redirect(url_for('auth.login'))
            if not synced:
                flash('Потрібна авторизація', 'danger')
                return redirect(url_for('auth.login'))

            is_admin = bool(session.get('is_admin', False))
            perms = session.get('permissions') or []

            if not required_perms:
                return f(*args, **kwargs)

            if is_admin or any(p in perms for p in required_perms):
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


def get_templates_with_metadata(is_admin=False):
    """
    Повертає список шаблонів разом з їхніми метаданими (опис,
    відображувана назва, чи "тільки для адміністратора") - для
    сторінок генерації документа (students.generate,
    admin.group_export) замість жорстко "зашитих" перевірок на кшталт
    `if 'adddiplom' in tpl.lower()`.

    Кожен елемент списку - словник:
        {'path': 'template_word/файл.docx', 'filename': 'файл.docx',
         'display_name': ..., 'description': ..., 'admin_only': bool}

    Шаблони, для яких немає запису в таблиці document_templates
    (наприклад, файли, покладені в папку вручну, без завантаження через
    сторінку управління шаблонами), все одно показуються - просто без
    опису і НЕ як адмінські (admin_only=False), із назвою за
    замовчуванням (ім'я файлу).

    :param is_admin: якщо False - шаблони з admin_only=True не включаються в результат.
    """
    from routes.db import get_db  # локальний імпорт, щоб уникнути циклічної залежності

    paths = get_available_templates()
    if not paths:
        return []

    conn = get_db()
    rows = conn.execute("SELECT filename, display_name, description, admin_only, hidden FROM document_templates").fetchall()
    meta_by_filename = {row['filename']: dict(row) for row in rows}

    result = []
    for path in paths:
        filename = path.split('/')[-1]
        meta = meta_by_filename.get(filename, {})
        # Приховані шаблони не показуються в списках вибору взагалі
        # (керується галочкою "Показувати в списках" на сторінці
        # управління шаблонами; файли без запису в БД - видимі).
        if meta.get('hidden', 0):
            continue
        admin_only = bool(meta.get('admin_only', 0))
        if admin_only and not is_admin:
            continue
        result.append({
            'path': path,
            'filename': filename,
            'display_name': meta.get('display_name') or filename,
            'description': meta.get('description') or '',
            'admin_only': admin_only,
        })
    return result


def save_multiple_attachments(conn, entity_type, entity_id, files, upload_subdir, uploaded_by=None):
    """Зберігає список файлів (з request.files.getlist(...)) на диск і
    реєструє кожен у спільній таблиці "attachments".
    Перевіряє безпечні розширення та заголовки файлів (захист від XSS/виконаних файлів)."""
    import uuid as _uuid
    from routes.security import validate_upload_file, SAFE_ATTACHMENT_EXTENSIONS

    saved_paths = []
    upload_dir = os.path.join('static', 'uploads', upload_subdir)
    os.makedirs(upload_dir, exist_ok=True)
    for f in (files or []):
        if not f or not f.filename:
            continue
        ok, err = validate_upload_file(f, SAFE_ATTACHMENT_EXTENSIONS)
        if not ok:
            raise ValueError(f"Файл «{f.filename}»: {err}")
        ext = os.path.splitext(f.filename)[1].lower()
        safe_name = f"{_uuid.uuid4().hex}{ext}"
        f.save(os.path.join(upload_dir, safe_name))
        rel_path = f"uploads/{upload_subdir}/{safe_name}"
        conn.execute(
            "INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_by) VALUES (?, ?, ?, ?, ?)",
            (entity_type, entity_id, rel_path, f.filename, uploaded_by)
        )
        saved_paths.append(rel_path)
    return saved_paths


def get_attachments(conn, entity_type, entity_id):
    """Повертає всі вкладення для entity_type+entity_id, найновіші
    останніми (в порядку завантаження)."""
    return conn.execute(
        "SELECT id, file_path, original_name, uploaded_at, uploaded_by FROM attachments "
        "WHERE entity_type = ? AND entity_id = ? ORDER BY id",
        (entity_type, entity_id)
    ).fetchall()
    
def is_student_on_reduced_program(conn, student_id, group_id):
    """Чи навчається студент за скороченою програмою відносно КОНКРЕТНОЇ
    групи (override менше за program_credits групи).

    Override тепер зберігається в student_enrollments.program_credits_override
    (per-group), а не в students.program_credits_override — інакше
    паралельний студент не міг би мати різні override у різних групах.
    Fallback на students.program_credits_override — для студентів без
    enrollment'а (legacy / до міграції).
    """
    row = conn.execute("""
        SELECT
            COALESCE(e.program_credits_override, s.program_credits_override) AS override_credits,
            g.program_credits AS group_credits
        FROM students s
        JOIN groups g ON g.id = ?
        LEFT JOIN student_enrollments e
               ON e.student_id = s.id AND e.group_id = g.id
        WHERE s.id = ?
    """, (group_id, student_id)).fetchone()
    if not row or not row['override_credits'] or not row['group_credits']:
        return False
    try:
        return int(row['override_credits']) < int(row['group_credits'])
    except (TypeError, ValueError):
        return False

def program_track_condition(is_reduced, alias=''):
    """SQL-умова (починається з " AND ..."), що обмежує запит до
    subjects/practices/courseworks/attestations пунктами, видимими на
    ФАКТИЧНІЙ програмі студента:
      - студент на повній програмі (is_reduced=False): ховає пункти,
        позначені reduced_only=1 (існують лише для скороченої);
      - студент на скороченій (is_reduced=True): ховає пункти,
        позначені full_program_only=1 (існують лише для повної).

    alias - псевдонім таблиці в запиті (напр. "s" -> "s.full_program_only");
    порожній рядок, якщо запит без псевдоніма.

    Єдина точка цієї перевірки - раніше one-off умова
    "AND full_program_only = 0" була написана вручну в ~10 різних
    місцях (генерація документів, картка студента, форми оцінок), і
    коли з'явився дзеркальний reduced_only, довелось би повторити
    правку в кожному з них. Тепер запити просто додають
    program_track_condition(is_reduced, alias) до WHERE.

    ВАЖЛИВО - архітектурне обмеження: цей механізм (full_program_only /
    reduced_only / reduced_credits / reduced_type на предметі +
    program_credits_override на студенті) розрахований РІВНО на дві
    програми - повну і скорочену. Якщо колись з'явиться третій варіант
    навчального плану (напр. ще коротша програма), цих двох булевих
    прапорців і двох "reduced_*"-полів не вистачить - знадобиться
    окрема таблиця "навчальний план" з plan_id на предметах і студентах
    замість двох фіксованих полів. Прив'язуватись до буквально двох
    варіантів у новому коді небезпечно.
    """
    prefix = f"{alias}." if alias else ""
    if is_reduced:
        return f" AND {prefix}full_program_only = 0"
    return f" AND {prefix}reduced_only = 0"


def apply_track_overrides(item, is_reduced):
    """Підміняє credits/type пункту (предмета/практики/курсової/
    атестації) на його reduced_credits/reduced_type, якщо студент на
    скороченій програмі і ці поля заповнені. item має бути dict (не
    sqlite3.Row - Row не підтримує присвоєння). Повертає той самий
    item для зручності ланцюжкового виклику.

    Спільна для генерації документів (gen_docx.py), картки студента і
    форм виставлення оцінок (students.py) - раніше та сама пара
    "if is_reduced and item.get('reduced_X'): item['Y'] = ..." була
    продубльована в кожному з цих місць окремо.
    """
    if is_reduced:
        if item.get('reduced_credits') is not None:
            item['credits'] = item['reduced_credits']
        if item.get('reduced_type'):
            item['type'] = item['reduced_type']
    return item


def filter_students_for_item(students, item, conn, group_id):
    """Для масового виставлення оцінки по ОДНОМУ пункту (предмет/
    практика/...) - фільтрує список студентів групи так, щоб
    залишились лише ті, кому цей пункт реально стосується:
      - пункт full_program_only=1 -> лишає тільки студентів НЕ на
        скороченій програмі;
      - пункт reduced_only=1 -> лишає тільки студентів НА скороченій;
      - інакше - лишає всіх без змін.
    students - список sqlite3.Row/dict з полем 'id'. item - dict-подібний
    з полями full_program_only/reduced_only (з рядка subjects/...).
    """
    if item['full_program_only']:
        return [s for s in students if not is_student_on_reduced_program(conn, s['id'], group_id)]
    if item['reduced_only']:
        return [s for s in students if is_student_on_reduced_program(conn, s['id'], group_id)]
    return students
    
import re as _re_phone

# Дозволені роздільники в номерах телефонів (пробіли, дефіси, дужки,
# точки, неперервний пробіл) — прибираються при нормалізації.
_PHONE_SEPARATORS_RE = _re_phone.compile(r'[\s\-\(\)\.\u00a0]')


def normalize_phone_ua(value):
    """
    Нормалізує український номер телефону до канонічного вигляду
    +380XXXXXXXXX (12 цифр з плюсом).

    Приймає всі типові варіанти введення:
        0971234567          → +380971234567
        +380971234567       → +380971234567
        380971234567        → +380971234567
        80971234567         → +380971234567
        097 123 45 67       → +380971234567
        (097) 123-45-67     → +380971234567
        +38 (097) 123 45 67 → +380971234567
        +380-97-123-45-67   → +380971234567

    Якщо номер не схожий на український (напр. іноземний студент
    з польським/німецьким номером) — повертає його з прибраними
    зайвими пробілами, БЕЗ примусової нормалізації (щоб не зіпсувати).

    None / порожній рядок → None.
    """
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None

    digits = _PHONE_SEPARATORS_RE.sub('', s).lstrip('+')

    # 12 цифр з 380 — вже канон (380 + 9 цифр)
    if len(digits) == 12 and digits.startswith('380'):
        return '+' + digits

    # 10 цифр, починається з 0 — типовий український мобільний
    # 067/068/096/097/098/050/066/095/099/063/073/093/091/092/094 тощо
    if len(digits) == 10 and digits.startswith('0'):
        return '+38' + digits

    # 11 цифр, починається з 80 — «+» пропущено
    if len(digits) == 11 and digits.startswith('80'):
        return '+3' + digits

    # Не схоже на український формат (напр. іноземний номер) —
    # повертаємо канонічний вигляд: без роздільників, з провідним '+',
    # якщо він був.
    has_plus = s.lstrip().startswith('+')
    return ('+' if has_plus else '') + digits
    
def normalize_phone_pair(phone, phone_backup):
    """
    Нормалізує ОБИДВА телефони (основний і резервний) і прибирає
    безглуздий дублікат, якщо після нормалізації вони виявились
    однаковими (студент ввів той самий номер у різних форматах).

    Повертає кортеж (phone, phone_backup), де phone_backup
    гарантовано None, якщо він збігається з phone.

    Це той випадок, коли "0971234567" і "+380971234567" — насправді
    одне й те саме, і зберігати його двічі немає сенсу.
    """
    norm_phone = normalize_phone_ua(phone)
    norm_backup = normalize_phone_ua(phone_backup)

    if norm_backup and norm_phone and norm_backup == norm_phone:
        norm_backup = None

    return norm_phone, norm_backup


# ═══════════════════════════════════════════════════════════════════════════════
# Факультети: довідник, міграція схеми, початкове заповнення та резолвінг
# ═══════════════════════════════════════════════════════════════════════════════

FACULTY_SEED = (
    ('Факультет економіки та менеджменту', 'ФЕМ', 'Faculty of Economics and Management', 10),
    ('Факультет інформаційних систем та технологій', 'ФІСТ', 'Faculty of Information Systems and Technologies', 20),
    ('Факультет психології та соціальних технологій', 'ФПСТ', 'Faculty of Psychology and Social Technologies', 30),
    ('Юридичний факультет', 'ЮФ', 'Faculty of Law', 40),
    ('Факультет іноземної філології', 'ФІФ', 'Faculty of Foreign Philology', 50),
)

# Відповідність кодів спеціальностей (КМУ №1021 та попередні шифри) скороченій назві факультету
SPECIALTY_CODE_TO_FACULTY_SHORT = {
    # ФЕМ
    'C1': 'ФЕМ', 'D1': 'ФЕМ', 'D2': 'ФЕМ', 'D3': 'ФЕМ', 'D4': 'ФЕМ',
    'D5': 'ФЕМ', 'D6': 'ФЕМ', 'D7': 'ФЕМ', 'J2': 'ФЕМ', 'J3': 'ФЕМ',
    '051': 'ФЕМ', '071': 'ФЕМ', '072': 'ФЕМ', '073': 'ФЕМ', '075': 'ФЕМ',
    '076': 'ФЕМ', '281': 'ФЕМ', '292': 'ФЕМ',
    # ФІСТ
    'F1': 'ФІСТ', 'F2': 'ФІСТ', 'F3': 'ФІСТ', 'F4': 'ФІСТ', 'F5': 'ФІСТ',
    'F6': 'ФІСТ', 'F7': 'ФІСТ',
    '113': 'ФІСТ', '121': 'ФІСТ', '122': 'ФІСТ', '124': 'ФІСТ', '125': 'ФІСТ', '126': 'ФІСТ',
    # ФПСТ
    'C4': 'ФПСТ', 'C5': 'ФПСТ', 'I4': 'ФПСТ', 'I10': 'ФПСТ',
    '053': 'ФПСТ', '231': 'ФПСТ',
    # ЮФ
    'D8': 'ЮФ', 'D9': 'ЮФ', 'K9': 'ЮФ',
    '081': 'ЮФ', '262': 'ЮФ', '293': 'ЮФ',
    # ФІФ
    'B11': 'ФІФ', '035': 'ФІФ',
}


def ensure_faculty_schema(conn):
    """
    Ідемпотентно створює таблицю `faculties`, додає колонки `faculty_id` у `specialties`
    та `faculty_id`, `faculty`, `faculty_en` у `groups`, заповнює 5 університетських
    факультетів і виконує одноразовий початковий backfill (відстежується у `_patch_migrations`).
    Безпечно викликати повторно на будь-якому з'єднанні SQLite.
    """
    conn.execute("""
        CREATE TABLE IF NOT EXISTS faculties (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name_ua TEXT NOT NULL UNIQUE,
            short_name TEXT,
            name_en TEXT,
            sort_order INTEGER NOT NULL DEFAULT 100,
            is_active INTEGER NOT NULL DEFAULT 1
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_faculties_active_order ON faculties(is_active, sort_order, id)")

    for name_ua, short_name, name_en, sort_order in FACULTY_SEED:
        conn.execute(
            "INSERT OR IGNORE INTO faculties (name_ua, short_name, name_en, sort_order, is_active) VALUES (?, ?, ?, ?, 1)",
            (name_ua, short_name, name_en, sort_order),
        )

    def _table_columns(table_name):
        return {r[1] for r in conn.execute(f"PRAGMA table_info({table_name})").fetchall()}

    spec_cols = _table_columns('specialties')
    if spec_cols and 'faculty_id' not in spec_cols:
        conn.execute("ALTER TABLE specialties ADD COLUMN faculty_id INTEGER")

    group_cols = _table_columns('groups')
    if group_cols:
        if 'faculty_id' not in group_cols:
            conn.execute("ALTER TABLE groups ADD COLUMN faculty_id INTEGER")
        if 'faculty' not in group_cols:
            conn.execute("ALTER TABLE groups ADD COLUMN faculty TEXT")
        if 'faculty_en' not in group_cols:
            conn.execute("ALTER TABLE groups ADD COLUMN faculty_en TEXT")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS _patch_migrations (
            key TEXT PRIMARY KEY,
            applied_at TEXT NOT NULL DEFAULT (datetime('now'))
        )
    """)

    done = conn.execute(
        "SELECT 1 FROM _patch_migrations WHERE key = 'faculties_initial_backfill_v1'"
    ).fetchone()

    if not done:
        fac_rows = conn.execute("SELECT id, short_name, name_ua, name_en FROM faculties").fetchall()
        by_short = {(r[1] or '').strip().upper(): (r[0], r[2], r[3] or '') for r in fac_rows if r[1]}

        # 1. Прив'язуємо спеціальності без faculty_id за картою кодів
        if spec_cols:
            spec_rows = conn.execute("SELECT code, faculty_id FROM specialties").fetchall()
            for sp_row in spec_rows:
                sp_code = (sp_row[0] or '').strip().upper()
                if sp_row[1] is not None:
                    continue
                short = SPECIALTY_CODE_TO_FACULTY_SHORT.get(sp_code)
                if short and short in by_short:
                    fid, _, _ = by_short[short]
                    conn.execute(
                        "UPDATE specialties SET faculty_id = ? WHERE code = ? AND faculty_id IS NULL",
                        (fid, sp_row[0]),
                    )

        # 2. Заповнюємо групи без факультету: спочатку зі спеціальності, далі за картою шифрів
        if group_cols:
            spec_fac_map = {}
            if spec_cols:
                for row in conn.execute("""
                    SELECT UPPER(TRIM(s.code)) AS code, f.id AS fid, f.name_ua, f.name_en
                    FROM specialties s
                    JOIN faculties f ON f.id = s.faculty_id
                """).fetchall():
                    spec_fac_map[row[0]] = (row[1], row[2], row[3] or '')

            has_sp_code = 'specialty_code' in group_cols
            has_sp_text = 'specialty' in group_cols
            sel_cols = ["id"]
            if has_sp_code:
                sel_cols.append("specialty_code")
            if has_sp_text:
                sel_cols.append("specialty")
            sel_cols.extend(["faculty_id", "faculty"])
            grp_rows = conn.execute(
                f"SELECT {', '.join(sel_cols)} FROM groups"
            ).fetchall()
            for g_row in grp_rows:
                idx = 1
                g_sp_code = g_row[idx] if has_sp_code else None
                if has_sp_code:
                    idx += 1
                g_sp_text = g_row[idx] if has_sp_text else None
                if has_sp_text:
                    idx += 1
                gid, g_fid, g_fac = g_row[0], g_row[idx], g_row[idx + 1]
                if g_fid is not None or (g_fac or '').strip():
                    continue
                code_norm = (g_sp_code or '').strip().upper()
                if not code_norm and g_sp_text:
                    code_norm = (g_sp_text or '').strip().split()[0].upper()

                target = spec_fac_map.get(code_norm)
                if not target:
                    short = SPECIALTY_CODE_TO_FACULTY_SHORT.get(code_norm)
                    if short and short in by_short:
                        target = by_short[short]
                if target:
                    fid, f_ua, f_en = target
                    conn.execute(
                        "UPDATE groups SET faculty_id = ?, faculty = ?, faculty_en = ? WHERE id = ?",
                        (fid, f_ua, f_en or None, gid),
                    )

        conn.execute(
            "INSERT OR IGNORE INTO _patch_migrations (key) VALUES ('faculties_initial_backfill_v1')"
        )

    conn.commit()


def lookup_specialty_faculty(conn, specialty_code):
    """Повертає (name_ua, name_en) факультету за кодом спеціальності або ('', '')."""
    if conn is None or not specialty_code:
        return '', ''
    code_clean = str(specialty_code).strip()
    if not code_clean:
        return '', ''
    candidates = [code_clean]
    first_tok = code_clean.split()[0]
    if first_tok and first_tok != code_clean:
        candidates.append(first_tok)
    try:
        for cand in candidates:
            row = conn.execute("""
                SELECT f.name_ua, f.name_en
                FROM specialties s
                JOIN faculties f ON f.id = s.faculty_id
                WHERE UPPER(TRIM(s.code)) = UPPER(TRIM(?))
                LIMIT 1
            """, (cand,)).fetchone()
            if row:
                return (row[0] or '').strip(), (row[1] or '').strip()
    except Exception:
        pass
    return '', ''


def resolve_vidomist_faculty(sheet, conn=None):
    """
    Визначає назву факультету для відомостей/експортів сесій за пріоритетом:
      1. Явний факультет групи/відомості (`sheet['faculty']`)
      2. Типовий факультет спеціальності з БД (`specialties.faculty_id`)
      3. Резервно — галузь знань (`sheet['knowledge_area']`)
      4. `sheet['department']` -> ''
    Повертає кортеж (faculty_ua, faculty_en, knowledge_area).
    """
    def _g(key):
        if sheet is None:
            return ''
        try:
            v = sheet[key]
        except (KeyError, IndexError, TypeError):
            v = sheet.get(key) if isinstance(sheet, dict) else None
        return (str(v).strip() if v is not None else '')

    faculty_ua = _g('faculty')
    faculty_en = _g('faculty_en')
    knowledge_area = _g('knowledge_area')
    specialty_code = _g('specialty_code') or _g('specialty')

    if not faculty_ua and conn is not None and specialty_code:
        sp_ua, sp_en = lookup_specialty_faculty(conn, specialty_code)
        if sp_ua:
            faculty_ua = sp_ua
            if not faculty_en and sp_en:
                faculty_en = sp_en

    if not faculty_ua:
        faculty_ua = knowledge_area or _g('department') or ''

    return faculty_ua, faculty_en, knowledge_area

