"""
routes/security.py
==================
Централізований модуль безпеки застосунку:
  - CSRF-захист (per-session токен, перевірка заголовка/форми для POST/PUT/PATCH/DELETE);
  - In-memory sliding-window Rate Limiter для захисту /login і публічних маршрутів;
  - Валідація завантажуваних файлів (білий список розширень + magic bytes + перевірка .docx);
  - Захист від Path Traversal (нормалізація і перевірка меж директорії, валідація токенів);
  - Єдина перевірка області видимості студента (IDOR-захист через user_scope);
  - Політика паролів та білий список ролей;
  - Санітизація комірок CSV/XLSX від Formula Injection;
  - Контроль доступу до /static/uploads/ з перевіркою сесії, прав і скоупа.
"""

from collections import deque
import hmac
import io
import os
import re
import secrets
import threading
import time
import zipfile

from flask import (
    abort,
    jsonify,
    request,
    send_file,
    session,
)

from routes.config import PROJECT_ROOT

# ═══════════════════════════════════════════════════════════════════════
# 1. CSRF-захист
# ═══════════════════════════════════════════════════════════════════════

# Зовнішні webhook-и (наприклад, ONLYOFFICE Document Server, який
# автентифікується власним JWT-підписом без браузерної сесії).
CSRF_EXEMPT_PREFIXES = (
    '/office/callback/',
)


def generate_csrf_token() -> str:
    """Повертає поточний CSRF-токен сесії або генерує новий."""
    token = session.get('_csrf_token')
    if not token or not isinstance(token, str):
        token = secrets.token_urlsafe(32)
        session['_csrf_token'] = token
        session.modified = True
    return token


def _is_json_or_ajax_request() -> bool:
    if request.is_json:
        return True
    if (request.headers.get('X-Requested-With') or '').lower() == 'xmlhttprequest':
        return True
    accept = (request.headers.get('Accept') or '').lower()
    if 'application/json' in accept and 'text/html' not in accept:
        return True
    return False


def validate_csrf_request():
    """
    Глобальна перевірка CSRF для всіх state-changing запитів.
    Повертає Response 403 при провалі перевірки або None при успіху.
    """
    if request.method not in ('POST', 'PUT', 'PATCH', 'DELETE'):
        return None

    for prefix in CSRF_EXEMPT_PREFIXES:
        if request.path.startswith(prefix):
            return None

    expected = session.get('_csrf_token')
    submitted = (
        request.headers.get('X-CSRFToken')
        or request.headers.get('X-CSRF-Token')
        or request.form.get('csrf_token')
        or ''
    )

    if (
        not expected
        or not isinstance(expected, str)
        or not isinstance(submitted, str)
        or not submitted
        or not hmac.compare_digest(expected, submitted)
    ):
        if _is_json_or_ajax_request():
            return jsonify({
                'ok': False,
                'error': 'Помилка перевірки безпеки (CSRF). Оновіть сторінку та спробуйте ще раз.',
            }), 403
        return (
            'Помилка перевірки безпеки (CSRF). Поверніться назад, оновіть сторінку і повторіть дію.',
            403,
        )

    return None


# ═══════════════════════════════════════════════════════════════════════
# 2. Sliding-window Rate Limiter (thread-safe, без зовнішніх залежностей)
# ═══════════════════════════════════════════════════════════════════════

_RATE_LOCK = threading.Lock()
_RATE_BUCKETS: dict[tuple[str, str], deque[float]] = {}
_LAST_CLEANUP_AT = 0.0


def client_ip() -> str:
    return (request.remote_addr or 'unknown').strip()


def check_rate_limit(bucket: str, key: str, limit: int, window_seconds: int) -> tuple[bool, int]:
    """
    Перевіряє, чи не перевищено ліміт запитів (limit за window_seconds).
    Повертає (allowed, retry_after_seconds).
    """
    global _LAST_CLEANUP_AT
    now = time.monotonic()
    cutoff = now - window_seconds
    composite = (bucket, str(key))

    with _RATE_LOCK:
        # Періодичне прибирання старих ключів, щоб пам'ять не росла нескінченно
        if now - _LAST_CLEANUP_AT > 300 or len(_RATE_BUCKETS) > 5000:
            stale = []
            for k, dq in _RATE_BUCKETS.items():
                while dq and dq[0] <= cutoff:
                    dq.popleft()
                if not dq:
                    stale.append(k)
            for k in stale:
                _RATE_BUCKETS.pop(k, None)
            _LAST_CLEANUP_AT = now

        dq = _RATE_BUCKETS.get(composite)
        if dq is None:
            dq = deque()
            _RATE_BUCKETS[composite] = dq

        while dq and dq[0] <= cutoff:
            dq.popleft()

        if len(dq) >= limit:
            retry_after = max(1, int(window_seconds - (now - dq[0])) + 1)
            return False, retry_after

        dq.append(now)
        return True, 0


# ═══════════════════════════════════════════════════════════════════════
# 3. Валідація завантажуваних файлів та токенів
# ═══════════════════════════════════════════════════════════════════════

SAFE_ATTACHMENT_EXTENSIONS = frozenset({
    '.pdf',
    '.jpg', '.jpeg', '.png', '.webp',
    '.doc', '.docx', '.xls', '.xlsx', '.txt',
})

SAFE_SCAN_EXTENSIONS = frozenset({'.pdf', '.jpg', '.jpeg', '.png', '.webp'})

_UUID_RE = re.compile(r'^[0-9a-fA-F]{32}$|^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
_HEX_RE = re.compile(r'^[0-9a-fA-F]{16,64}$')


def is_safe_uuid(value: str) -> bool:
    return bool(value and isinstance(value, str) and _UUID_RE.fullmatch(value.strip()))


def is_safe_hex_token(value: str) -> bool:
    return bool(value and isinstance(value, str) and _HEX_RE.fullmatch(value.strip()))


def resolve_under_dir(base_dir: str, rel_path: str):
    """
    Повертає безпечний канонічний шлях усередині base_dir або None
    при спробі обходу директорії / некоректному шляху.
    """
    if not rel_path or not isinstance(rel_path, str):
        return None
    if '\x00' in rel_path:
        return None
    clean_rel = rel_path.replace('\\', '/').lstrip('/')
    if not clean_rel:
        return None
    real_base = os.path.realpath(base_dir)
    candidate = os.path.realpath(os.path.join(real_base, clean_rel))
    if candidate != real_base and not candidate.startswith(real_base + os.sep):
        return None
    return candidate


def _check_magic_header(ext: str, header: bytes) -> bool:
    """Перевіряє відповідність перших байтів очікуваному формату."""
    if ext == '.pdf':
        return b'%PDF-' in header[:16]
    if ext in ('.jpg', '.jpeg'):
        return header.startswith(b'\xff\xd8\xff')
    if ext == '.png':
        return header.startswith(b'\x89PNG\r\n\x1a\n')
    if ext == '.webp':
        return len(header) >= 12 and header[:4] == b'RIFF' and header[8:12] == b'WEBP'
    if ext in ('.docx', '.xlsx'):
        return header.startswith(b'PK\x03\x04')
    if ext in ('.doc', '.xls'):
        return header.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1') or header.startswith(b'PK\x03\x04')
    if ext == '.txt':
        return b'\x00' not in header
    return False


def validate_upload_file(file_storage, allowed_exts=SAFE_ATTACHMENT_EXTENSIONS) -> tuple[bool, str]:
    """
    Перевіряє FileStorage на дозволене розширення та базовий заголовок файлу.
    Завжди повертає курсор потоку на початок.
    """
    if not file_storage or not getattr(file_storage, 'filename', ''):
        return False, 'Файл не обрано'

    filename = file_storage.filename.strip()
    ext = os.path.splitext(filename)[1].lower()
    if not ext or ext not in allowed_exts:
        return False, f"Недопустимий формат файлу ({ext or 'без розширення'})"

    stream = file_storage.stream
    pos = stream.tell()
    try:
        header = stream.read(64)
    finally:
        stream.seek(pos)

    if not header:
        return False, 'Файл порожній'

    if not _check_magic_header(ext, header):
        return False, f"Вміст файлу не відповідає формату {ext}"

    return True, ''


def validate_docx_bytes(raw_bytes: bytes, max_uncompressed: int = 50 * 1024 * 1024) -> tuple[bool, str]:
    """
    Перевіряє, що завантажений шаблон є коректним OOXML (.docx) архівом
    і не містить zip/XML-бомби надмірного розміру після розпакування.
    """
    if not raw_bytes or not raw_bytes.startswith(b'PK\x03\x04'):
        return False, 'Файл не є коректним документом .docx'
    try:
        with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
            names = set(zf.namelist())
            if '[Content_Types].xml' not in names or 'word/document.xml' not in names:
                return False, 'Архів не має структури Word-документа (.docx)'
            total_size = sum(info.file_size for info in zf.infolist())
            if total_size > max_uncompressed:
                return False, 'Документ перевищує допустимий розмір після розпакування'
    except zipfile.BadZipFile:
        return False, 'Пошкоджений файл .docx'
    return True, ''


# ═══════════════════════════════════════════════════════════════════════
# 4. Паролі та ролі
# ═══════════════════════════════════════════════════════════════════════

ALLOWED_USER_ROLES = frozenset({'admin', 'user'})

_COMMON_WEAK_PASSWORDS = frozenset({
    '1234567890', '0123456789', '12345678901', 'password123',
    'qwerty1234', 'qwertyuiop', 'admin12345', 'administrator',
    'letmein123', 'welcome123', 'passw0rd12', '1111111111',
})


def validate_password_strength(password: str) -> tuple[bool, str]:
    if not password or len(password) < 10:
        return False, 'Пароль має містити щонайменше 10 символів'
    if password.lower() in _COMMON_WEAK_PASSWORDS:
        return False, 'Обраний пароль занадто простий, оберіть надійніший'
    has_letter = any(ch.isalpha() for ch in password)
    has_digit = any(ch.isdigit() for ch in password)
    if not (has_letter and has_digit):
        return False, 'Пароль має містити щонайменше одну літеру та одну цифру'
    return True, ''


# ═══════════════════════════════════════════════════════════════════════
# 5. Авторизація та область видимості (IDOR-захист)
# ═══════════════════════════════════════════════════════════════════════

def user_can_access_student(conn, student_id: int, user_id=None, is_admin=None) -> bool:
    """
    Перевіряє, чи має поточний користувач доступ до студента student_id
    відповідно до його актуального скоупа (user_groups, user_specialties,
    user_degree_levels, user_knowledge_areas).
    """
    if is_admin is None:
        is_admin = bool(session.get('is_admin', False)) or session.get('role') == 'admin'
    if is_admin:
        return True

    if user_id is None:
        user_id = session.get('user_id')
    if not user_id:
        return False

    from routes.user_scope import get_allowed_group_ids
    allowed_gids = set(get_allowed_group_ids(conn, user_id))
    if not allowed_gids:
        return False

    st_row = conn.execute(
        "SELECT group_id FROM students WHERE id = ?",
        (student_id,),
    ).fetchone()
    if not st_row:
        return False

    student_gids = set()
    if st_row['group_id'] is not None:
        student_gids.add(int(st_row['group_id']))

    for r in conn.execute(
        "SELECT group_id FROM student_enrollments WHERE student_id = ? AND group_id IS NOT NULL",
        (student_id,),
    ).fetchall():
        student_gids.add(int(r['group_id']))

    return bool(student_gids & allowed_gids)


def has_any_permission(*perms: str) -> bool:
    if bool(session.get('is_admin', False)) or session.get('role') == 'admin':
        return True
    user_perms = set(session.get('permissions') or [])
    return any(p in user_perms for p in perms)


# ═══════════════════════════════════════════════════════════════════════
# 6. Захист CSV / XLSX від Formula Injection
# ═══════════════════════════════════════════════════════════════════════

_PURE_INTL_PHONE_RE = re.compile(r'^\+\d{7,15}$')


def sanitize_spreadsheet_cell(value, *, for_csv: bool = False):
    """
    Запобігає виконанню формул у Excel / LibreOffice / Google Sheets, якщо
    текстова комірка починається з '=', '+', '-', '@', '\t' або '\r'.
    Числа (int/float) і чисті міжнародні телефони формату +380... для XLSX
    залишаються незмінними.
    """
    if value is None or isinstance(value, (int, float, bool)):
        return value
    s = str(value)
    if not s:
        return s

    first = s[0]
    stripped_first = s.lstrip()[0] if s.lstrip() else ''
    dangerous = ('=', '@', '\t', '\r')
    if first in ('\t', '\r') or stripped_first in dangerous:
        return "'" + s

    if stripped_first in ('+', '-'):
        # В CSV будь-який провідний +/- може інтерпретуватися як формула.
        # В XLSX зберігаємо валідні телефони виду +380XXXXXXXXX без апострофа.
        if not for_csv and _PURE_INTL_PHONE_RE.fullmatch(s.strip()):
            return s
        return "'" + s

    return s


# ═══════════════════════════════════════════════════════════════════════
# 7. Захищена віддача файлів /static/uploads/
# ═══════════════════════════════════════════════════════════════════════

_BLOCKED_STATIC_PREFIXES = (
    '/static/templates/',
    '/static/uploads/tmp/',
    '/static/uploads/tmp_military_pdf/',
    '/static/uploads/tmp_military_bulk/',
    '/static/uploads/import_reports/',
    '/static/uploads/import_sessions/',
    '/static/uploads/office_sessions/',
    '/static/uploads/photo_bulk_sessions/',
    '/static/uploads/license_reports/',
)

_SAFE_INLINE_IMAGE_OR_PDF = frozenset({'.jpg', '.jpeg', '.png', '.webp', '.pdf'})
_STUDENT_PHOTO_RE = re.compile(r'^photos/student_(\d+)\.jpg$')


def _attachment_student_id(conn, rel_upload_path: str):
    """Пробує знайти student_id, до якого належить файл у attachments."""
    db_rel = f"uploads/{rel_upload_path}"
    row = conn.execute(
        "SELECT entity_type, entity_id FROM attachments WHERE file_path = ? LIMIT 1",
        (db_rel,),
    ).fetchone()
    if not row:
        return None
    etype = row['entity_type']
    eid = row['entity_id']
    if etype == 'student':
        return int(eid)
    if etype == 'military':
        m = conn.execute("SELECT student_id FROM military WHERE id = ?", (eid,)).fetchone()
        return int(m['student_id']) if m else None
    if etype == 'education_document':
        d = conn.execute("SELECT student_id FROM education_documents WHERE id = ?", (eid,)).fetchone()
        return int(d['student_id']) if d else None
    if etype == 'passport_document':
        p = conn.execute("SELECT student_id FROM passport_documents WHERE id = ?", (eid,)).fetchone()
        return int(p['student_id']) if p else None
    if etype == 'frozen_student':
        f = conn.execute("SELECT student_id FROM frozen_students WHERE id = ?", (eid,)).fetchone()
        return int(f['student_id']) if f else None
    return None


def authorize_and_serve_static():
    """
    Викликається з app.before_request:
      - повністю блокує доступ до /static/templates/ і тимчасових/сесійних папок;
      - перехоплює будь-які запити до /static/uploads/ та перевіряє сесію,
        права доступу і прив'язку студента перед віддачею файлу.
    """
    path = request.path or ''
    for prefix in _BLOCKED_STATIC_PREFIXES:
        if path.startswith(prefix):
            abort(404)

    if not path.startswith('/static/uploads/'):
        return None

    # Без авторизації будь-кий доступ до /static/uploads/ закритий
    if not session.get('user_id'):
        abort(404)

    subpath = path[len('/static/uploads/'):]
    if not subpath or '..' in subpath or '\\' in subpath or '\x00' in subpath:
        abort(404)

    uploads_root = os.path.join(PROJECT_ROOT, 'static', 'uploads')
    disk_path = resolve_under_dir(uploads_root, subpath)
    if not disk_path or not os.path.isfile(disk_path):
        abort(404)

    is_admin = bool(session.get('is_admin', False)) or session.get('role') == 'admin'

    from routes.db import get_db
    conn = get_db()
    try:
        allowed = False
        top_folder = subpath.split('/', 1)[0]

        if is_admin:
            allowed = True
        elif top_folder == 'photos':
            m = _STUDENT_PHOTO_RE.fullmatch(subpath)
            if m:
                sid = int(m.group(1))
                if (
                    has_any_permission('manage_military_reports', 'manage_pending_students', 'manage_update_requests')
                    or user_can_access_student(conn, sid)
                ):
                    allowed = True
        elif top_folder == 'military_records':
            if has_any_permission('manage_military_reports'):
                allowed = True
            elif has_any_permission('manage_students', 'view_student_card'):
                sid = _attachment_student_id(conn, subpath)
                allowed = (sid is not None and user_can_access_student(conn, sid))
        elif top_folder == 'passport_documents':
            if has_any_permission('manage_passport_documents'):
                allowed = True
            elif has_any_permission('manage_students', 'view_student_card'):
                sid = _attachment_student_id(conn, subpath)
                allowed = (sid is not None and user_can_access_student(conn, sid))
        elif top_folder == 'education_documents':
            if has_any_permission('manage_education_documents'):
                allowed = True
            elif has_any_permission('manage_students', 'view_student_card'):
                sid = _attachment_student_id(conn, subpath)
                allowed = (sid is not None and user_can_access_student(conn, sid))
        elif top_folder in ('orders', 'course_transfer_orders'):
            allowed = has_any_permission('manage_orders', 'manage_students')
        elif top_folder == 'frozen_students':
            if has_any_permission('manage_students', 'view_student_card'):
                sid = _attachment_student_id(conn, subpath)
                allowed = (sid is None and has_any_permission('manage_students')) or (sid is not None and user_can_access_student(conn, sid))
        elif top_folder in ('pending_students', 'pending_pdfs'):
            allowed = has_any_permission('manage_pending_students')
        elif top_folder == 'pending_photos':
            allowed = has_any_permission(
                'manage_pending_students', 'manage_update_requests',
                'manage_students', 'manage_military_reports',
            )
        elif top_folder == 'update_requests':
            allowed = has_any_permission('manage_update_requests')
    finally:
        conn.close()

    if not allowed:
        abort(403)

    ext = os.path.splitext(disk_path)[1].lower()
    as_attachment = ext not in _SAFE_INLINE_IMAGE_OR_PDF
    resp = send_file(disk_path, as_attachment=as_attachment)
    resp.headers['X-Content-Type-Options'] = 'nosniff'
    resp.headers['Cache-Control'] = 'private, no-store'
    return resp
