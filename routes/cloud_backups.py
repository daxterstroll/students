"""
routes/cloud_backups.py
=======================
Хмарне резервне копіювання реальної бази даних SQLite (`students.db`) у зовнішні сховища:
  1. Google Диск (OAuth 2.0 Refresh Token для @gmail.com / Workspace або Service Account JSON).
  2. S3-сумісне сховище (Cloudflare R2, Backblaze B2, AWS S3, локальний MinIO) — чистий AWS SigV4.
  3. WebDAV (Nextcloud, ownCloud, Synology NAS, Hetzner Storage Box).
  4. Telegram-канал / приватний чат (стиснений .zip до 49 МБ).

Принципи безпеки:
  - Конфігурація і секрети зберігаються у `cloud_backup_config.json` у корені проєкту
    (поза `students.db` та поза папкою `backups/`), тож відновлення БД чи тестовий режим не зачіпають ключі.
  - У браузер і шаблони передаються ТІЛЬКИ масковані прапорці (`has_*`) — жоден секрет не віддається назад на клієнт.
  - Зовнішні URL проходять перевірку схеми (`http`/`https`) та захист від SSRF на cloud-metadata (`169.254.0.0/16`).
  - Додаткові pip-залежності не потрібні (використовується лише `requests` і стандартна бібліотека Python).
"""

from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import hmac
import ipaddress
import json
import os
import re
import secrets
import threading
import time
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile

import requests

from routes.config import PROJECT_ROOT
from routes.db import get_real_db
from routes.security import resolve_under_dir
from routes.utils import logger

CLOUD_CONFIG_PATH = os.path.join(PROJECT_ROOT, 'cloud_backup_config.json')
BACKUPS_DIR = os.path.join(PROJECT_ROOT, 'backups')

_CLOUD_CFG_LOCK = threading.RLock()
_BACKUP_NAME_RE = re.compile(r'^[A-Za-z0-9_.-]+\.db$')
_GDRIVE_FOLDER_URL_RE = re.compile(r'/folders/([A-Za-z0-9_-]{10,})')
_GDRIVE_ID_RE = re.compile(r'^[A-Za-z0-9_-]{10,128}$')

PROVIDER_ORDER = ('gdrive', 's3', 'webdav', 'telegram')

PROVIDER_LABELS = {
    'gdrive': 'Google Диск',
    's3': 'S3-сховище (R2 / B2 / MinIO / AWS)',
    'webdav': 'WebDAV (Nextcloud / NAS)',
    'telegram': 'Telegram-канал / чат',
}

TELEGRAM_MAX_BYTES = 49 * 1024 * 1024  # запас до ліміту Telegram Bot API (50 МБ)
_EMPTY_SHA256 = hashlib.sha256(b'').hexdigest()

DEFAULT_CLOUD_CONFIG = {
    'gdrive': {
        'enabled': False,
        'auto_upload': False,
        'compress_zip': True,
        'auth_mode': 'oauth',  # 'oauth' | 'service_account'
        'client_id': '',
        'client_secret': '',
        'refresh_token': '',
        'service_account_json': '',
        'folder_id': '',
        'keep_remote': 14,
        'account_email': '',
        'last_status': '',
        'last_error': '',
        'last_uploaded_at': '',
        'last_file': '',
        'last_trigger': '',
    },
    's3': {
        'enabled': False,
        'auto_upload': False,
        'compress_zip': True,
        'preset': 'r2',  # 'r2' | 'b2' | 'aws' | 'minio' | 'custom'
        'endpoint_url': '',
        'region': 'auto',
        'bucket': '',
        'prefix': 'students-backups/',
        'access_key_id': '',
        'secret_access_key': '',
        'path_style': True,
        'keep_remote': 14,
        'last_status': '',
        'last_error': '',
        'last_uploaded_at': '',
        'last_file': '',
        'last_trigger': '',
    },
    'webdav': {
        'enabled': False,
        'auto_upload': False,
        'compress_zip': True,
        'base_url': '',
        'remote_dir': 'students-backups',
        'username': '',
        'password': '',
        'verify_ssl': True,
        'keep_remote': 14,
        'last_status': '',
        'last_error': '',
        'last_uploaded_at': '',
        'last_file': '',
        'last_trigger': '',
    },
    'telegram': {
        'enabled': False,
        'auto_upload': False,
        'compress_zip': True,
        'token_source': 'military_bot',  # 'military_bot' | 'custom'
        'bot_token': '',
        'chat_id': '',
        'caption_prefix': '🗄 Бекап БД',
        'last_status': '',
        'last_error': '',
        'last_uploaded_at': '',
        'last_file': '',
        'last_trigger': '',
    },
}

_SECRET_FIELDS = {
    'gdrive': ('client_secret', 'refresh_token', 'service_account_json'),
    's3': ('secret_access_key',),
    'webdav': ('password',),
    'telegram': ('bot_token',),
}


# ═══════════════════════════════════════════════════════════════════════════════
# 1. Завантаження, збереження та маскування конфігурації
# ═══════════════════════════════════════════════════════════════════════════════

def _clone_default_config() -> dict:
    return {k: dict(v) for k, v in DEFAULT_CLOUD_CONFIG.items()}


def get_cloud_config_raw() -> dict:
    with _CLOUD_CFG_LOCK:
        cfg = _clone_default_config()
        if os.path.exists(CLOUD_CONFIG_PATH):
            try:
                with open(CLOUD_CONFIG_PATH, 'r', encoding='utf-8') as f:
                    raw = json.load(f)
                if isinstance(raw, dict):
                    for prov in PROVIDER_ORDER:
                        if isinstance(raw.get(prov), dict):
                            cfg[prov].update(raw[prov])
            except Exception as e:
                logger.warning(f"Не вдалося прочитати cloud_backup_config.json: {e}")
        return cfg


def _write_cloud_config_raw(cfg: dict) -> None:
    with _CLOUD_CFG_LOCK:
        tmp = f"{CLOUD_CONFIG_PATH}.tmp.{os.getpid()}.{secrets.token_hex(4)}"
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        try:
            os.chmod(tmp, 0o600)
        except OSError:
            pass
        os.replace(tmp, CLOUD_CONFIG_PATH)


def _update_provider_runtime_status(
    provider: str,
    ok: bool,
    message_or_error: str,
    filename: str = '',
    trigger: str = '',
) -> None:
    if provider not in PROVIDER_ORDER:
        return
    with _CLOUD_CFG_LOCK:
        cfg = get_cloud_config_raw()
        p = cfg[provider]
        p['last_status'] = 'ok' if ok else 'error'
        if ok:
            p['last_error'] = ''
            p['last_uploaded_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            if filename:
                p['last_file'] = filename
            if trigger:
                p['last_trigger'] = trigger
        else:
            p['last_error'] = (message_or_error or 'Невідома помилка')[:400]
        _write_cloud_config_raw(cfg)


def _mask_preview(val: str) -> str:
    s = (val or '').strip()
    if not s:
        return ''
    if len(s) <= 8:
        return '••••' + s[-2:] if len(s) >= 2 else '••••'
    return s[:4] + '••••' + s[-4:]


def _is_provider_configured(provider: str, p: dict) -> bool:
    if provider == 'gdrive':
        if p.get('auth_mode') == 'service_account':
            return bool((p.get('service_account_json') or '').strip())
        return bool(
            (p.get('client_id') or '').strip()
            and (p.get('client_secret') or '').strip()
            and (p.get('refresh_token') or '').strip()
        )
    if provider == 's3':
        return bool(
            (p.get('endpoint_url') or '').strip()
            and (p.get('bucket') or '').strip()
            and (p.get('access_key_id') or '').strip()
            and (p.get('secret_access_key') or '').strip()
        )
    if provider == 'webdav':
        return bool(
            (p.get('base_url') or '').strip()
            and (p.get('username') or '').strip()
            and (p.get('password') or '').strip()
        )
    if provider == 'telegram':
        has_token = bool((p.get('bot_token') or '').strip()) if p.get('token_source') == 'custom' else bool(_get_military_bot_token())
        return bool(has_token and (p.get('chat_id') or '').strip())
    return False


def get_cloud_config_masked() -> dict:
    """Повертає конфігурацію без жодних сирих секретів (лише прапорці has_* та несекретні поля)."""
    raw = get_cloud_config_raw()
    masked: dict = {}
    for prov in PROVIDER_ORDER:
        p = dict(raw[prov])
        secrets_set = _SECRET_FIELDS.get(prov, ())
        for sf in secrets_set:
            val = (p.pop(sf, '') or '').strip()
            p[f'has_{sf}'] = bool(val)
            p[f'{sf}_preview'] = _mask_preview(val) if val and sf != 'service_account_json' else ('JSON завантажено' if val else '')
        p['configured'] = _is_provider_configured(prov, raw[prov])
        p['label'] = PROVIDER_LABELS[prov]
        if prov == 'telegram':
            p['has_military_bot_token'] = bool(_get_military_bot_token())
        masked[prov] = p
    return masked


def save_cloud_provider_config(provider: str, form_data: dict) -> tuple[bool, str]:
    if provider not in PROVIDER_ORDER:
        return False, 'Невідомий провайдер хмарного сховища'

    with _CLOUD_CFG_LOCK:
        cfg = get_cloud_config_raw()
        p = cfg[provider]

        enabled = bool(form_data.get('enabled'))
        auto_upload = bool(form_data.get('auto_upload'))
        compress_zip = bool(form_data.get('compress_zip'))

        if provider == 'gdrive':
            auth_mode = (form_data.get('auth_mode') or 'oauth').strip()
            if auth_mode not in ('oauth', 'service_account'):
                auth_mode = 'oauth'
            client_id = (form_data.get('client_id') or '').strip()
            folder_id = extract_gdrive_folder_id(form_data.get('folder_id') or '')
            keep_remote = _clamp_int(form_data.get('keep_remote'), default=14, lo=0, hi=200)

            client_secret = (form_data.get('client_secret') or '').strip() or p.get('client_secret', '')
            refresh_token = (form_data.get('refresh_token') or '').strip() or p.get('refresh_token', '')
            sa_raw = (form_data.get('service_account_json') or '').strip()
            if sa_raw:
                ok_sa, sa_err = _validate_service_account_json(sa_raw)
                if not ok_sa:
                    return False, sa_err
                service_account_json = sa_raw
            else:
                service_account_json = p.get('service_account_json', '')

            if auth_mode == 'service_account' and folder_id == '':
                # Нагадуємо тільки якщо вмикається сервісний акаунт без папки
                pass

            p.update({
                'enabled': enabled,
                'auto_upload': auto_upload and enabled,
                'compress_zip': compress_zip,
                'auth_mode': auth_mode,
                'client_id': client_id,
                'client_secret': client_secret,
                'refresh_token': refresh_token,
                'service_account_json': service_account_json,
                'folder_id': folder_id,
                'keep_remote': keep_remote,
            })

        elif provider == 's3':
            endpoint_url = (form_data.get('endpoint_url') or '').strip().rstrip('/')
            if endpoint_url:
                ok_u, u_err = validate_remote_http_url(endpoint_url)
                if not ok_u:
                    return False, u_err
            preset = (form_data.get('preset') or 'r2').strip()
            if preset not in ('r2', 'b2', 'aws', 'minio', 'custom'):
                preset = 'custom'
            region = (form_data.get('region') or '').strip() or ('auto' if preset == 'r2' else 'us-east-1')
            bucket = (form_data.get('bucket') or '').strip()
            if bucket and not re.fullmatch(r'[a-z0-9][a-z0-9._-]{1,62}[a-z0-9]', bucket):
                return False, 'Некоректна назва бакету S3 (допускаються латинські малі літери, цифри, ".", "-", "_")'
            prefix = (form_data.get('prefix') or 'students-backups/').strip().lstrip('/')
            if prefix and not prefix.endswith('/'):
                prefix += '/'
            access_key_id = (form_data.get('access_key_id') or '').strip()
            secret_access_key = (form_data.get('secret_access_key') or '').strip() or p.get('secret_access_key', '')
            path_style = bool(form_data.get('path_style'))
            keep_remote = _clamp_int(form_data.get('keep_remote'), default=14, lo=0, hi=200)

            p.update({
                'enabled': enabled,
                'auto_upload': auto_upload and enabled,
                'compress_zip': compress_zip,
                'preset': preset,
                'endpoint_url': endpoint_url,
                'region': region,
                'bucket': bucket,
                'prefix': prefix,
                'access_key_id': access_key_id,
                'secret_access_key': secret_access_key,
                'path_style': path_style,
                'keep_remote': keep_remote,
            })

        elif provider == 'webdav':
            base_url = (form_data.get('base_url') or '').strip()
            if base_url:
                ok_u, u_err = validate_remote_http_url(base_url)
                if not ok_u:
                    return False, u_err
                if not base_url.endswith('/'):
                    base_url += '/'
            remote_dir = (form_data.get('remote_dir') or 'students-backups').strip().strip('/')
            if '..' in remote_dir.split('/'):
                return False, 'Некоректна назва віддаленої теки WebDAV'
            username = (form_data.get('username') or '').strip()
            password = (form_data.get('password') or '').strip() or p.get('password', '')
            verify_ssl = bool(form_data.get('verify_ssl'))
            keep_remote = _clamp_int(form_data.get('keep_remote'), default=14, lo=0, hi=200)

            p.update({
                'enabled': enabled,
                'auto_upload': auto_upload and enabled,
                'compress_zip': compress_zip,
                'base_url': base_url,
                'remote_dir': remote_dir,
                'username': username,
                'password': password,
                'verify_ssl': verify_ssl,
                'keep_remote': keep_remote,
            })

        elif provider == 'telegram':
            token_source = (form_data.get('token_source') or 'military_bot').strip()
            if token_source not in ('military_bot', 'custom'):
                token_source = 'military_bot'
            bot_token = (form_data.get('bot_token') or '').strip() or p.get('bot_token', '')
            chat_id = (form_data.get('chat_id') or '').strip()
            caption_prefix = (form_data.get('caption_prefix') or '🗄 Бекап БД').strip()[:80]

            p.update({
                'enabled': enabled,
                'auto_upload': auto_upload and enabled,
                'compress_zip': compress_zip,
                'token_source': token_source,
                'bot_token': bot_token,
                'chat_id': chat_id,
                'caption_prefix': caption_prefix,
            })

        if p.get('enabled') and not _is_provider_configured(provider, p):
            p['enabled'] = False
            p['auto_upload'] = False
            _write_cloud_config_raw(cfg)
            return False, f"Параметри для «{PROVIDER_LABELS[provider]}» збережено, але провайдер не активовано — заповніть усі обов'язкові ключі/поля."

        _write_cloud_config_raw(cfg)
        return True, f"Налаштування «{PROVIDER_LABELS[provider]}» збережено."


def disconnect_cloud_provider(provider: str) -> tuple[bool, str]:
    if provider not in PROVIDER_ORDER:
        return False, 'Невідомий провайдер'
    with _CLOUD_CFG_LOCK:
        cfg = get_cloud_config_raw()
        cfg[provider] = dict(DEFAULT_CLOUD_CONFIG[provider])
        _write_cloud_config_raw(cfg)
    return True, f"Провайдер «{PROVIDER_LABELS[provider]}» відключено, ключі доступу очищено."


# ═══════════════════════════════════════════════════════════════════════════════
# 2. Допоміжні функції безпеки та підготовки файлу
# ═══════════════════════════════════════════════════════════════════════════════

def _clamp_int(raw, default: int, lo: int, hi: int) -> int:
    try:
        val = int(str(raw).strip())
    except (TypeError, ValueError):
        return default
    return max(lo, min(hi, val))


def validate_remote_http_url(url: str) -> tuple[bool, str]:
    """
    Перевіряє, що URL має безпечну схему http/https і не веде до адрес
    хмарних метаданих (169.254.0.0/16, metadata.google.internal тощо).
    Локальні NAS/MinIO у мережі університету (192.168.*, 10.*, .local) дозволені.
    """
    s = (url or '').strip()
    if not s:
        return False, 'URL не може бути порожнім'
    try:
        parsed = urllib.parse.urlsplit(s)
    except Exception:
        return False, 'Некоректний формат URL'

    if parsed.scheme.lower() not in ('http', 'https'):
        return False, 'Дозволені лише протоколи http:// або https://'
    if parsed.username or parsed.password:
        return False, 'Не вказуйте логін/пароль безпосередньо в URL — використовуйте окремі поля'
    host = (parsed.hostname or '').strip().lower()
    if not host:
        return False, 'Не вказано домен або IP-адресу сервера'
    if host in ('metadata.google.internal', 'metadata.goog', '0.0.0.0'):
        return False, 'Заборонена адреса сервера метаданих'

    try:
        ip_obj = ipaddress.ip_address(host)
        if ip_obj.is_link_local or ip_obj.is_unspecified or ip_obj.is_multicast:
            return False, f'Заборонена мережева адреса: {host}'
    except ValueError:
        pass

    return True, ''


def extract_gdrive_folder_id(raw: str) -> str:
    s = (raw or '').strip()
    if not s:
        return ''
    m = _GDRIVE_FOLDER_URL_RE.search(s)
    if m:
        return m.group(1)
    if _GDRIVE_ID_RE.fullmatch(s):
        return s
    # Якщо передали URL виду ?id=...
    try:
        parsed = urllib.parse.urlsplit(s)
        qs = urllib.parse.parse_qs(parsed.query)
        if qs.get('id') and _GDRIVE_ID_RE.fullmatch(qs['id'][0]):
            return qs['id'][0]
    except Exception:
        pass
    return s


def _validate_service_account_json(raw: str) -> tuple[bool, str]:
    try:
        data = json.loads(raw)
    except Exception:
        return False, 'Файл або JSON Service Account має некоректний синтаксис'
    if not isinstance(data, dict) or data.get('type') != 'service_account':
        return False, 'Очікується JSON типу "service_account"'
    if not data.get('client_email') or not data.get('private_key'):
        return False, 'У JSON Service Account відсутні client_email або private_key'
    return True, ''


def _get_military_bot_token() -> str:
    env_tok = (os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
    try:
        conn = get_real_db()
        try:
            row = conn.execute("SELECT bot_token FROM telegram_bot_settings WHERE id = 1").fetchone()
            if row and row['bot_token']:
                return str(row['bot_token']).strip()
        finally:
            conn.close()
    except Exception:
        pass
    return env_tok


def _resolve_backup_path(filename: str) -> str | None:
    if not filename or not _BACKUP_NAME_RE.fullmatch(filename):
        return None
    path = resolve_under_dir(BACKUPS_DIR, filename)
    if not path or not os.path.isfile(path):
        return None
    return path


def _prepare_upload_artifact(src_path: str, compress_zip: bool) -> tuple[str, str, str, bool]:
    """
    Повертає (upload_path, remote_filename, content_type, should_cleanup).
    """
    base_name = os.path.basename(src_path)
    if not compress_zip:
        return src_path, base_name, 'application/x-sqlite3', False

    os.makedirs(BACKUPS_DIR, exist_ok=True)
    tmp_zip = os.path.join(BACKUPS_DIR, f".cloud_tmp_{secrets.token_hex(6)}_{base_name}.zip")
    with zipfile.ZipFile(tmp_zip, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        zf.write(src_path, arcname=base_name)
    return tmp_zip, f"{base_name}.zip", 'application/zip', True


def _format_size_human(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} Б"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} КБ"
    if num_bytes < 1024 * 1024 * 1024:
        return f"{num_bytes / (1024 * 1024):.2f} МБ"
    return f"{num_bytes / (1024 * 1024 * 1024):.2f} ГБ"


# ═══════════════════════════════════════════════════════════════════════════════
# 3. Провайдер 1: Google Диск (OAuth 2.0 + опціонально Service Account)
# ═══════════════════════════════════════════════════════════════════════════════

GDRIVE_AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth'
GDRIVE_TOKEN_URL = 'https://oauth2.googleapis.com/token'
GDRIVE_DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.file'


def build_gdrive_auth_url(client_id: str, redirect_uri: str, state: str) -> str:
    params = {
        'client_id': client_id.strip(),
        'redirect_uri': redirect_uri.strip(),
        'response_type': 'code',
        'scope': GDRIVE_DRIVE_SCOPE,
        'access_type': 'offline',
        'prompt': 'consent',
        'include_granted_scopes': 'true',
        'state': state,
    }
    return f"{GDRIVE_AUTH_URL}?{urllib.parse.urlencode(params)}"


def extract_oauth_code_from_input(raw_input: str) -> tuple[str, str]:
    """
    Приймає або чистий authorization code, або повний URL редіректу з Google
    (наприклад, http://127.0.0.1/?code=4/0A...&state=...&scope=...).
    Повертає (code, state_from_url).
    """
    s = (raw_input or '').strip()
    if not s:
        return '', ''
    if s.startswith('http://') or s.startswith('https://') or 'code=' in s:
        if not (s.startswith('http://') or s.startswith('https://')):
            s = 'http://127.0.0.1/?' + s.lstrip('?')
        parsed = urllib.parse.urlsplit(s)
        qs = urllib.parse.parse_qs(parsed.query)
        code = (qs.get('code') or [''])[0].strip()
        state = (qs.get('state') or [''])[0].strip()
        return code, state
    return s, ''


def exchange_gdrive_oauth_code(
    client_id: str,
    client_secret: str,
    code: str,
    redirect_uri: str,
) -> tuple[bool, dict, str]:
    try:
        resp = requests.post(
            GDRIVE_TOKEN_URL,
            data={
                'client_id': client_id.strip(),
                'client_secret': client_secret.strip(),
                'code': code.strip(),
                'redirect_uri': redirect_uri.strip(),
                'grant_type': 'authorization_code',
            },
            timeout=25,
        )
        data = resp.json() if resp.content else {}
        if resp.status_code != 200 or not data.get('access_token'):
            err = data.get('error_description') or data.get('error') or f"HTTP {resp.status_code}"
            return False, {}, f"Google OAuth помилка обміну коду: {err}"
        return True, data, ''
    except Exception as e:
        return False, {}, f"Помилка з'єднання з Google OAuth: {e}"


def _gdrive_get_access_token(p: dict) -> tuple[bool, str, str]:
    auth_mode = p.get('auth_mode') or 'oauth'
    if auth_mode == 'service_account':
        sa_raw = (p.get('service_account_json') or '').strip()
        ok_sa, sa_err = _validate_service_account_json(sa_raw)
        if not ok_sa:
            return False, '', sa_err
        sa_data = json.loads(sa_raw)
        try:
            import jwt  # PyJWT
        except ImportError:
            return False, '', 'Бібліотека PyJWT не встановлена на сервері. Використовуйте режим OAuth 2.0.'
        now = int(time.time())
        claim = {
            'iss': sa_data['client_email'],
            'scope': 'https://www.googleapis.com/auth/drive',
            'aud': GDRIVE_TOKEN_URL,
            'iat': now,
            'exp': now + 3600,
        }
        try:
            assertion = jwt.encode(claim, sa_data['private_key'], algorithm='RS256')
        except Exception as e:
            return False, '', (
                f"Не вдалося підписати JWT для Service Account ({e}). "
                "Для Service Account потрібен пакет cryptography, або перемкніться на OAuth 2.0."
            )
        try:
            resp = requests.post(
                GDRIVE_TOKEN_URL,
                data={
                    'grant_type': 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                    'assertion': assertion,
                },
                timeout=25,
            )
            data = resp.json() if resp.content else {}
            if resp.status_code != 200 or not data.get('access_token'):
                err = data.get('error_description') or data.get('error') or f"HTTP {resp.status_code}"
                return False, '', f"Помилка автентифікації Service Account: {err}"
            return True, data['access_token'], ''
        except Exception as e:
            return False, '', f"Помилка запиту токена Google: {e}"

    # Стандартний OAuth 2.0 (refresh_token)
    cid = (p.get('client_id') or '').strip()
    csec = (p.get('client_secret') or '').strip()
    rtok = (p.get('refresh_token') or '').strip()
    if not (cid and csec and rtok):
        return False, '', 'Не заповнені Client ID, Client Secret або Refresh Token Google'

    try:
        resp = requests.post(
            GDRIVE_TOKEN_URL,
            data={
                'client_id': cid,
                'client_secret': csec,
                'refresh_token': rtok,
                'grant_type': 'refresh_token',
            },
            timeout=25,
        )
        data = resp.json() if resp.content else {}
        if resp.status_code != 200 or not data.get('access_token'):
            err = data.get('error_description') or data.get('error') or f"HTTP {resp.status_code}"
            return False, '', f"Не вдалося оновити токен Google Drive: {err}"
        return True, data['access_token'], ''
    except Exception as e:
        return False, '', f"Помилка з'єднання з Google OAuth: {e}"


def _gdrive_test(p: dict) -> tuple[bool, str]:
    ok, token, err = _gdrive_get_access_token(p)
    if not ok:
        return False, err

    headers = {'Authorization': f'Bearer {token}'}
    try:
        about_resp = requests.get(
            'https://www.googleapis.com/drive/v3/about',
            params={'fields': 'user(emailAddress,displayName),storageQuota(limit,usage)'},
            headers=headers,
            timeout=20,
        )
        if about_resp.status_code != 200:
            return False, f"Google Drive повернув HTTP {about_resp.status_code}: {about_resp.text[:200]}"
        about = about_resp.json()
        user_info = about.get('user') or {}
        email = user_info.get('emailAddress') or user_info.get('displayName') or 'підключено'
        quota = about.get('storageQuota') or {}
        quota_hint = ''
        if quota.get('limit'):
            used = int(quota.get('usage') or 0)
            limit = int(quota.get('limit') or 0)
            quota_hint = f" · зайнято {_format_size_human(used)} з {_format_size_human(limit)}"

        folder_id = (p.get('folder_id') or '').strip()
        folder_hint = ''
        if folder_id:
            f_resp = requests.get(
                f"https://www.googleapis.com/drive/v3/files/{urllib.parse.quote(folder_id, safe='')}",
                params={'fields': 'id,name,mimeType', 'supportsAllDrives': 'true'},
                headers=headers,
                timeout=20,
            )
            if f_resp.status_code != 200:
                return False, f"Акаунт ({email}) працює, але вказана тека Google Диску недоступна (HTTP {f_resp.status_code})"
            f_data = f_resp.json()
            if f_data.get('mimeType') != 'application/vnd.google-apps.folder':
                return False, 'Вказаний Google Drive ID не є текою'
            folder_hint = f" · тека «{f_data.get('name')}»"

        # Збережемо email акаунта у несекретному полі
        with _CLOUD_CFG_LOCK:
            cfg = get_cloud_config_raw()
            cfg['gdrive']['account_email'] = email
            _write_cloud_config_raw(cfg)

        return True, f"Підключення успішне: {email}{quota_hint}{folder_hint}"
    except Exception as e:
        return False, f"Помилка перевірки Google Drive: {e}"


def _gdrive_upload(p: dict, upload_path: str, remote_name: str, content_type: str) -> tuple[bool, str]:
    ok, token, err = _gdrive_get_access_token(p)
    if not ok:
        return False, err

    file_size = os.path.getsize(upload_path)
    folder_id = (p.get('folder_id') or '').strip()
    metadata: dict = {
        'name': remote_name,
        'appProperties': {'source': 'students_backup'},
    }
    if folder_id:
        metadata['parents'] = [folder_id]

    try:
        # 1. Ініціалізуємо Resumable Upload сесію
        init_resp = requests.post(
            'https://www.googleapis.com/upload/drive/v3/files',
            params={'uploadType': 'resumable', 'supportsAllDrives': 'true'},
            headers={
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json; charset=UTF-8',
                'X-Upload-Content-Type': content_type,
                'X-Upload-Content-Length': str(file_size),
            },
            json=metadata,
            timeout=30,
        )
        if init_resp.status_code not in (200, 201):
            return False, f"Google Drive не дозволив старт завантаження (HTTP {init_resp.status_code}): {init_resp.text[:250]}"
        session_uri = init_resp.headers.get('Location')
        if not session_uri:
            return False, 'Google Drive не повернув URL сесії завантаження'

        # 2. Завантажуємо файл частинами по 4 МБ (кратність 256 КБ) або одним PUT для малих файлів
        chunk_size = 4 * 1024 * 1024
        uploaded_file_id = ''
        with open(upload_path, 'rb') as f:
            if file_size <= chunk_size:
                put_resp = requests.put(
                    session_uri,
                    headers={
                        'Content-Length': str(file_size),
                        'Content-Type': content_type,
                    },
                    data=f,
                    timeout=180,
                )
                if put_resp.status_code not in (200, 201):
                    return False, f"Помилка завантаження в Google Drive (HTTP {put_resp.status_code}): {put_resp.text[:250]}"
                if put_resp.content:
                    uploaded_file_id = (put_resp.json() or {}).get('id', '')
            else:
                offset = 0
                while offset < file_size:
                    buf = f.read(chunk_size)
                    if not buf:
                        break
                    end = offset + len(buf) - 1
                    put_resp = requests.put(
                        session_uri,
                        headers={
                            'Content-Length': str(len(buf)),
                            'Content-Type': content_type,
                            'Content-Range': f'bytes {offset}-{end}/{file_size}',
                        },
                        data=buf,
                        timeout=180,
                    )
                    if put_resp.status_code in (200, 201):
                        if put_resp.content:
                            uploaded_file_id = (put_resp.json() or {}).get('id', '')
                        break
                    if put_resp.status_code != 308:
                        return False, f"Збій передачі чанку в Google Drive (HTTP {put_resp.status_code}): {put_resp.text[:250]}"
                    offset += len(buf)

        # 3. Ротація старих копій у Google Drive (якщо увімкнено)
        keep_remote = int(p.get('keep_remote') or 0)
        if keep_remote > 0:
            _gdrive_cleanup(token, folder_id, keep_remote)

        suffix = f" (ID: {uploaded_file_id})" if uploaded_file_id else ''
        return True, f"Завантажено у Google Диск як «{remote_name}»{suffix}"
    except Exception as e:
        return False, f"Помилка завантаження у Google Drive: {e}"


def _gdrive_cleanup(token: str, folder_id: str, keep_remote: int) -> None:
    try:
        q_parts = [
            "trashed = false",
            "name contains 'students_'",
            "appProperties has { key='source' and value='students_backup' }",
        ]
        if folder_id:
            safe_fid = folder_id.replace("'", "")
            q_parts.append(f"'{safe_fid}' in parents")

        resp = requests.get(
            'https://www.googleapis.com/drive/v3/files',
            params={
                'q': ' and '.join(q_parts),
                'orderBy': 'createdTime desc',
                'pageSize': 100,
                'fields': 'files(id,name,createdTime)',
                'supportsAllDrives': 'true',
                'includeItemsFromAllDrives': 'true',
            },
            headers={'Authorization': f'Bearer {token}'},
            timeout=25,
        )
        if resp.status_code != 200:
            return
        files = (resp.json() or {}).get('files') or []
        for old in files[keep_remote:]:
            fid = old.get('id')
            if not fid:
                continue
            requests.delete(
                f"https://www.googleapis.com/drive/v3/files/{urllib.parse.quote(fid, safe='')}",
                params={'supportsAllDrives': 'true'},
                headers={'Authorization': f'Bearer {token}'},
                timeout=15,
            )
    except Exception as e:
        logger.warning(f"Ротація файлів у Google Drive завершилась із попередженням: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# 4. Провайдер 2: S3-сумісне сховище (Cloudflare R2, B2, MinIO, AWS S3 — SigV4)
# ═══════════════════════════════════════════════════════════════════════════════

def _sign_hmac(key: bytes, msg: str) -> bytes:
    return hmac.new(key, msg.encode('utf-8'), hashlib.sha256).digest()


def _s3_signing_key(secret_key: str, date_stamp: str, region: str, service: str = 's3') -> bytes:
    k_date = _sign_hmac(('AWS4' + secret_key).encode('utf-8'), date_stamp)
    k_region = _sign_hmac(k_date, region)
    k_service = _sign_hmac(k_region, service)
    return _sign_hmac(k_service, 'aws4_request')


def _s3_build_url_and_headers(
    method: str,
    endpoint_url: str,
    region: str,
    bucket: str,
    object_key: str,
    access_key: str,
    secret_key: str,
    path_style: bool = True,
    query_params: dict | None = None,
    payload_hash: str = _EMPTY_SHA256,
    extra_headers: dict | None = None,
) -> tuple[str, dict]:
    parsed = urllib.parse.urlsplit(endpoint_url.strip().rstrip('/'))
    scheme = parsed.scheme.lower()
    base_host = parsed.netloc
    base_path = parsed.path.rstrip('/')

    clean_key = (object_key or '').lstrip('/')
    encoded_key = urllib.parse.quote(clean_key, safe='/-_.~')
    encoded_bucket = urllib.parse.quote(bucket.strip(), safe='-_.~')

    if path_style:
        host_header = base_host
        canonical_uri = f"{base_path}/{encoded_bucket}"
        if clean_key:
            canonical_uri += f"/{encoded_key}"
    else:
        host_header = f"{bucket.strip()}.{base_host}"
        canonical_uri = f"{base_path}/{encoded_key}" if clean_key else (f"{base_path}/" if base_path else '/')

    if not canonical_uri.startswith('/'):
        canonical_uri = '/' + canonical_uri

    qp = sorted(
        (
            urllib.parse.quote(str(k), safe='-_.~'),
            urllib.parse.quote(str(v), safe='-_.~'),
        )
        for k, v in (query_params or {}).items()
    )
    canonical_qs = '&'.join(f"{k}={v}" for k, v in qp)

    now_utc = datetime.now(timezone.utc)
    amz_date = now_utc.strftime('%Y%m%dT%H%M%SZ')
    date_stamp = amz_date[:8]

    canonical_headers = (
        f"host:{host_header}\n"
        f"x-amz-content-sha256:{payload_hash}\n"
        f"x-amz-date:{amz_date}\n"
    )
    signed_headers = 'host;x-amz-content-sha256;x-amz-date'
    canonical_request = '\n'.join([
        method.upper(),
        canonical_uri,
        canonical_qs,
        canonical_headers,
        signed_headers,
        payload_hash,
    ])

    eff_region = (region or 'us-east-1').strip()
    scope = f"{date_stamp}/{eff_region}/s3/aws4_request"
    string_to_sign = '\n'.join([
        'AWS4-HMAC-SHA256',
        amz_date,
        scope,
        hashlib.sha256(canonical_request.encode('utf-8')).hexdigest(),
    ])
    signing_key = _s3_signing_key(secret_key, date_stamp, eff_region, 's3')
    signature = hmac.new(signing_key, string_to_sign.encode('utf-8'), hashlib.sha256).hexdigest()

    headers = {
        'Host': host_header,
        'x-amz-date': amz_date,
        'x-amz-content-sha256': payload_hash,
        'Authorization': (
            f"AWS4-HMAC-SHA256 Credential={access_key}/{scope}, "
            f"SignedHeaders={signed_headers}, Signature={signature}"
        ),
    }
    if extra_headers:
        headers.update(extra_headers)

    full_url = f"{scheme}://{host_header}{canonical_uri}"
    if canonical_qs:
        full_url += f"?{canonical_qs}"
    return full_url, headers


def _file_sha256_hex(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _parse_s3_list_xml(xml_text: str) -> list[tuple[str, str]]:
    """Повертає [(key, last_modified), ...] з XML відповіді ListObjectsV2 без DTD/XXE."""
    clean = re.sub(r'<!DOCTYPE[^>]*>|<!ENTITY[^>]*>', '', xml_text or '', flags=re.I)
    root = ET.fromstring(clean)
    items: list[tuple[str, str]] = []
    for elem in root.iter():
        if elem.tag.endswith('Contents'):
            key = ''
            lm = ''
            for child in elem:
                if child.tag.endswith('Key') and child.text:
                    key = child.text
                elif child.tag.endswith('LastModified') and child.text:
                    lm = child.text
            if key:
                items.append((key, lm))
    return items


def _s3_test(p: dict) -> tuple[bool, str]:
    ok_u, u_err = validate_remote_http_url(p.get('endpoint_url') or '')
    if not ok_u:
        return False, u_err

    bucket = (p.get('bucket') or '').strip()
    prefix = (p.get('prefix') or '').strip().lstrip('/')
    if not bucket:
        return False, 'Вкажіть назву бакету S3'

    try:
        url, headers = _s3_build_url_and_headers(
            method='GET',
            endpoint_url=p['endpoint_url'],
            region=p.get('region') or 'auto',
            bucket=bucket,
            object_key='',
            access_key=p['access_key_id'],
            secret_key=p['secret_access_key'],
            path_style=bool(p.get('path_style', True)),
            query_params={'list-type': '2', 'max-keys': '10', 'prefix': prefix},
        )
        resp = requests.get(url, headers=headers, timeout=25)
        if resp.status_code != 200:
            return False, f"S3 повернув HTTP {resp.status_code}: {resp.text[:250]}"
        items = _parse_s3_list_xml(resp.text)
        return True, f"Підключення до бакету «{bucket}» успішне (знайдено об'єктів за префіксом: {len(items)})"
    except Exception as e:
        return False, f"Помилка з'єднання з S3: {e}"


def _s3_upload(p: dict, upload_path: str, remote_name: str, content_type: str) -> tuple[bool, str]:
    ok_u, u_err = validate_remote_http_url(p.get('endpoint_url') or '')
    if not ok_u:
        return False, u_err

    bucket = (p.get('bucket') or '').strip()
    prefix = (p.get('prefix') or '').strip().lstrip('/')
    if prefix and not prefix.endswith('/'):
        prefix += '/'
    object_key = f"{prefix}{remote_name}"

    try:
        payload_hash = _file_sha256_hex(upload_path)
        url, headers = _s3_build_url_and_headers(
            method='PUT',
            endpoint_url=p['endpoint_url'],
            region=p.get('region') or 'auto',
            bucket=bucket,
            object_key=object_key,
            access_key=p['access_key_id'],
            secret_key=p['secret_access_key'],
            path_style=bool(p.get('path_style', True)),
            payload_hash=payload_hash,
            extra_headers={'Content-Type': content_type},
        )
        with open(upload_path, 'rb') as f:
            resp = requests.put(url, headers=headers, data=f, timeout=300)
        if resp.status_code not in (200, 201, 204):
            return False, f"Помилка завантаження в S3 (HTTP {resp.status_code}): {resp.text[:250]}"

        keep_remote = int(p.get('keep_remote') or 0)
        if keep_remote > 0:
            _s3_cleanup(p, prefix, keep_remote)

        return True, f"Завантажено у бакет «{bucket}» як «{object_key}»"
    except Exception as e:
        return False, f"Помилка завантаження в S3: {e}"


def _s3_cleanup(p: dict, prefix: str, keep_remote: int) -> None:
    try:
        list_url, list_headers = _s3_build_url_and_headers(
            method='GET',
            endpoint_url=p['endpoint_url'],
            region=p.get('region') or 'auto',
            bucket=p['bucket'],
            object_key='',
            access_key=p['access_key_id'],
            secret_key=p['secret_access_key'],
            path_style=bool(p.get('path_style', True)),
            query_params={'list-type': '2', 'prefix': f"{prefix}students_"},
        )
        resp = requests.get(list_url, headers=list_headers, timeout=25)
        if resp.status_code != 200:
            return
        items = _parse_s3_list_xml(resp.text)
        backups_only = [
            (k, lm) for k, lm in items
            if os.path.basename(k).startswith('students_') and (k.endswith('.db') or k.endswith('.db.zip'))
        ]
        backups_only.sort(key=lambda x: (x[1], x[0]), reverse=True)
        for old_key, _ in backups_only[keep_remote:]:
            del_url, del_headers = _s3_build_url_and_headers(
                method='DELETE',
                endpoint_url=p['endpoint_url'],
                region=p.get('region') or 'auto',
                bucket=p['bucket'],
                object_key=old_key,
                access_key=p['access_key_id'],
                secret_key=p['secret_access_key'],
                path_style=bool(p.get('path_style', True)),
            )
            requests.delete(del_url, headers=del_headers, timeout=15)
    except Exception as e:
        logger.warning(f"Ротація файлів у S3 завершилась із попередженням: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# 5. Провайдер 3: WebDAV (Nextcloud, ownCloud, Synology NAS, Storage Box)
# ═══════════════════════════════════════════════════════════════════════════════

def _webdav_join_url(base_url: str, *parts: str) -> str:
    base = base_url.strip().rstrip('/')
    segs = []
    for part in parts:
        for seg in str(part or '').replace('\\', '/').split('/'):
            seg = seg.strip()
            if seg and seg not in ('.', '..'):
                segs.append(urllib.parse.quote(seg, safe='-_.~'))
    return f"{base}/" + '/'.join(segs) if segs else f"{base}/"


def _webdav_ensure_dir(base_url: str, remote_dir: str, auth: tuple[str, str], verify_ssl: bool) -> tuple[bool, str]:
    clean_dir = (remote_dir or '').strip().strip('/')
    if not clean_dir:
        return True, ''
    acc = []
    for seg in clean_dir.split('/'):
        seg = seg.strip()
        if not seg or seg in ('.', '..'):
            continue
        acc.append(seg)
        dir_url = _webdav_join_url(base_url, *acc) + '/'
        resp = requests.request('MKCOL', dir_url, auth=auth, verify=verify_ssl, timeout=20)
        if resp.status_code not in (200, 201, 204, 301, 302, 405):
            return False, f"Не вдалося створити теку WebDAV «{'/'.join(acc)}» (HTTP {resp.status_code})"
    return True, ''


def _parse_webdav_propfind(xml_text: str) -> list[str]:
    clean = re.sub(r'<!DOCTYPE[^>]*>|<!ENTITY[^>]*>', '', xml_text or '', flags=re.I)
    root = ET.fromstring(clean)
    names: list[str] = []
    for elem in root.iter():
        if elem.tag.endswith('href') and elem.text:
            href = urllib.parse.unquote(elem.text.rstrip('/'))
            name = os.path.basename(href)
            if name.startswith('students_') and (name.endswith('.db') or name.endswith('.db.zip')):
                names.append(name)
    return sorted(set(names), reverse=True)


def _webdav_test(p: dict) -> tuple[bool, str]:
    ok_u, u_err = validate_remote_http_url(p.get('base_url') or '')
    if not ok_u:
        return False, u_err

    auth = (p.get('username') or '', p.get('password') or '')
    verify_ssl = bool(p.get('verify_ssl', True))
    try:
        ok_dir, dir_err = _webdav_ensure_dir(p['base_url'], p.get('remote_dir') or '', auth, verify_ssl)
        if not ok_dir:
            return False, dir_err
        target_dir_url = _webdav_join_url(p['base_url'], p.get('remote_dir') or '')
        resp = requests.request(
            'PROPFIND',
            target_dir_url,
            headers={'Depth': '1'},
            auth=auth,
            verify=verify_ssl,
            timeout=25,
        )
        if resp.status_code not in (200, 207):
            return False, f"WebDAV повернув HTTP {resp.status_code}: {resp.text[:250]}"
        files = _parse_webdav_propfind(resp.text)
        return True, f"З'єднання WebDAV успішне (знайдено копій у теці: {len(files)})"
    except Exception as e:
        return False, f"Помилка з'єднання з WebDAV: {e}"


def _webdav_upload(p: dict, upload_path: str, remote_name: str, content_type: str) -> tuple[bool, str]:
    ok_u, u_err = validate_remote_http_url(p.get('base_url') or '')
    if not ok_u:
        return False, u_err

    auth = (p.get('username') or '', p.get('password') or '')
    verify_ssl = bool(p.get('verify_ssl', True))
    remote_dir = p.get('remote_dir') or ''
    try:
        ok_dir, dir_err = _webdav_ensure_dir(p['base_url'], remote_dir, auth, verify_ssl)
        if not ok_dir:
            return False, dir_err

        put_url = _webdav_join_url(p['base_url'], remote_dir, remote_name)
        with open(upload_path, 'rb') as f:
            resp = requests.put(
                put_url,
                data=f,
                headers={'Content-Type': content_type},
                auth=auth,
                verify=verify_ssl,
                timeout=300,
            )
        if resp.status_code not in (200, 201, 204):
            return False, f"Помилка завантаження WebDAV (HTTP {resp.status_code}): {resp.text[:250]}"

        keep_remote = int(p.get('keep_remote') or 0)
        if keep_remote > 0:
            _webdav_cleanup(p['base_url'], remote_dir, auth, verify_ssl, keep_remote)

        return True, f"Завантажено на WebDAV як «{remote_dir}/{remote_name}»"
    except Exception as e:
        return False, f"Помилка завантаження WebDAV: {e}"


def _webdav_cleanup(base_url: str, remote_dir: str, auth: tuple[str, str], verify_ssl: bool, keep_remote: int) -> None:
    try:
        dir_url = _webdav_join_url(base_url, remote_dir)
        resp = requests.request('PROPFIND', dir_url, headers={'Depth': '1'}, auth=auth, verify=verify_ssl, timeout=25)
        if resp.status_code not in (200, 207):
            return
        files = _parse_webdav_propfind(resp.text)
        for old_name in files[keep_remote:]:
            del_url = _webdav_join_url(base_url, remote_dir, old_name)
            requests.delete(del_url, auth=auth, verify=verify_ssl, timeout=15)
    except Exception as e:
        logger.warning(f"Ротація файлів WebDAV завершилась із попередженням: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# 6. Провайдер 4: Telegram-канал / приватний чат
# ═══════════════════════════════════════════════════════════════════════════════

def _telegram_resolve_token(p: dict) -> str:
    if (p.get('token_source') or 'military_bot') == 'custom':
        return (p.get('bot_token') or '').strip()
    return _get_military_bot_token()


def _telegram_test(p: dict) -> tuple[bool, str]:
    token = _telegram_resolve_token(p)
    chat_id = (p.get('chat_id') or '').strip()
    if not token:
        return False, 'Не вказано токен Telegram-бота'
    if not chat_id:
        return False, 'Вкажіть Chat ID або @channelname для відправки копій'

    try:
        me = requests.get(f"https://api.telegram.org/bot{token}/getMe", timeout=15).json()
        if not me.get('ok'):
            return False, f"Невалідний токен бота: {me.get('description') or 'помилка'}"
        bot_user = (me.get('result') or {}).get('username') or 'bot'

        ch = requests.get(
            f"https://api.telegram.org/bot{token}/getChat",
            params={'chat_id': chat_id},
            timeout=15,
        ).json()
        if not ch.get('ok'):
            return False, f"Бот @{bot_user} активний, але чат «{chat_id}» недоступний: {ch.get('description') or 'переконайтеся, що бота додано адміністратором'}"
        res = ch.get('result') or {}
        title = res.get('title') or res.get('username') or str(res.get('id') or chat_id)
        return True, f"Підключення успішне: бот @{bot_user} → чат «{title}»"
    except Exception as e:
        return False, f"Помилка перевірки Telegram: {e}"


def _telegram_upload(p: dict, upload_path: str, remote_name: str, raw_db_size: int) -> tuple[bool, str]:
    token = _telegram_resolve_token(p)
    chat_id = (p.get('chat_id') or '').strip()
    if not token or not chat_id:
        return False, 'Не налаштовано токен або Chat ID Telegram'

    upload_size = os.path.getsize(upload_path)
    if upload_size > TELEGRAM_MAX_BYTES:
        return False, (
            f"Розмір файлу ({_format_size_human(upload_size)}) перевищує ліміт Telegram Bot API (50 МБ). "
            "Використайте Google Диск, S3 або WebDAV."
        )

    prefix = (p.get('caption_prefix') or '🗄 Бекап БД').strip()
    caption = (
        f"{prefix}\n"
        f"Файл: {remote_name}\n"
        f"Стиснений розмір: {_format_size_human(upload_size)} (БД: {_format_size_human(raw_db_size)})\n"
        f"Час: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    try:
        with open(upload_path, 'rb') as f:
            resp = requests.post(
                f"https://api.telegram.org/bot{token}/sendDocument",
                data={'chat_id': chat_id, 'caption': caption},
                files={'document': (remote_name, f)},
                timeout=180,
            )
        data = resp.json() if resp.content else {}
        if resp.status_code != 200 or not data.get('ok'):
            return False, f"Telegram повернув помилку: {data.get('description') or resp.text[:200]}"
        return True, f"Надіслано в Telegram ({chat_id}) як «{remote_name}»"
    except Exception as e:
        return False, f"Помилка відправки у Telegram: {e}"


# ═══════════════════════════════════════════════════════════════════════════════
# 7. Публічний API для перевірки, ручного та автоматичного вивантаження
# ═══════════════════════════════════════════════════════════════════════════════

def test_cloud_provider(provider: str) -> tuple[bool, str]:
    if provider not in PROVIDER_ORDER:
        return False, 'Невідомий провайдер'
    cfg = get_cloud_config_raw()
    p = cfg[provider]
    if not _is_provider_configured(provider, p):
        return False, f"Провайдер «{PROVIDER_LABELS[provider]}» ще не налаштовано (заповніть обов'язкові поля та збережіть)"

    if provider == 'gdrive':
        ok, msg = _gdrive_test(p)
    elif provider == 's3':
        ok, msg = _s3_test(p)
    elif provider == 'webdav':
        ok, msg = _webdav_test(p)
    else:
        ok, msg = _telegram_test(p)

    _update_provider_runtime_status(provider, ok, msg if not ok else '')
    return ok, msg


def upload_backup_to_provider(
    filename: str,
    provider: str,
    trigger: str = 'manual',
) -> tuple[bool, str, str]:
    """
    Вивантажує вказаний файл резервної копії БД (`students_*.db`) у вибране хмарне сховище.
    Повертає (ok, message, remote_name).
    """
    if provider not in PROVIDER_ORDER:
        return False, 'Невідомий провайдер хмарного сховища', ''

    src_path = _resolve_backup_path(filename)
    if not src_path:
        return False, 'Файл резервної копії не знайдено', ''

    cfg = get_cloud_config_raw()
    p = cfg[provider]
    if not _is_provider_configured(provider, p):
        return False, f"Провайдер «{PROVIDER_LABELS[provider]}» не налаштований", ''

    raw_size = os.path.getsize(src_path)
    compress_zip = bool(p.get('compress_zip', True))
    upload_path, remote_name, content_type, should_cleanup = _prepare_upload_artifact(src_path, compress_zip)

    try:
        if provider == 'gdrive':
            ok, msg = _gdrive_upload(p, upload_path, remote_name, content_type)
        elif provider == 's3':
            ok, msg = _s3_upload(p, upload_path, remote_name, content_type)
        elif provider == 'webdav':
            ok, msg = _webdav_upload(p, upload_path, remote_name, content_type)
        else:
            ok, msg = _telegram_upload(p, upload_path, remote_name, raw_size)
    finally:
        if should_cleanup and os.path.exists(upload_path):
            try:
                os.remove(upload_path)
            except OSError:
                pass

    _update_provider_runtime_status(
        provider=provider,
        ok=ok,
        message_or_error='' if ok else msg,
        filename=filename,
        trigger=trigger,
    )
    return ok, msg, remote_name


def upload_backup_to_enabled_providers(
    filename: str,
    trigger: str = 'manual',
    only_auto: bool = False,
) -> list[dict]:
    """
    Вивантажує бекап в усі увімкнені хмарні провайдери (або лише ті, де увімкнено `auto_upload`).
    Повертає список результатів [{provider, label, ok, message, remote_name}, ...].
    """
    cfg = get_cloud_config_raw()
    results = []
    for prov in PROVIDER_ORDER:
        p = cfg[prov]
        if not p.get('enabled'):
            continue
        if only_auto and not p.get('auto_upload'):
            continue
        if not _is_provider_configured(prov, p):
            continue
        ok, msg, remote_name = upload_backup_to_provider(filename, prov, trigger=trigger)
        results.append({
            'provider': prov,
            'label': PROVIDER_LABELS[prov],
            'ok': ok,
            'message': msg,
            'remote_name': remote_name,
        })
    return results
