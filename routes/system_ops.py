"""
routes/system_ops.py
====================
Системні інструменти адміністрування, надійність та швидка навігація:
  1. Тестовий режим («Пісочниця БД»):
     - Миттєве створення ізольованої копії робочої SQLite БД через sqlite3.Connection.backup();
     - Блокування вихідних повідомлень Telegram-бота під час тестування;
     - Повернення до незміненої реальної БД з автоматичним прибиранням тестових файлів.
  2. Гаряче резервне копіювання SQLite БД (WAL-безпечне) + авто-ротація + завантаження/відновлення.
  3. Моніторинг активних сесій користувачів і примусовий вихід (відкликання сеансу).
  4. Глобальний швидкий пошук (Ctrl+K) по студентах, групах і наказах (з урахуванням скоупу).
  5. Історія змін конкретної картки студента (з audit_logs та military_history).
"""

from __future__ import annotations

from datetime import datetime, timedelta
import gc
import hmac
import json
import os
import re
import secrets
import sqlite3
import threading
import time

from flask import (
    Blueprint,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

from routes.cloud_backups import (
    PROVIDER_LABELS,
    PROVIDER_ORDER,
    build_gdrive_auth_url,
    disconnect_cloud_provider,
    exchange_gdrive_oauth_code,
    extract_oauth_code_from_input,
    get_cloud_config_masked,
    get_cloud_config_raw,
    save_cloud_provider_config,
    test_cloud_provider,
    upload_backup_to_enabled_providers,
    upload_backup_to_provider,
)
from routes.config import PROJECT_ROOT
from routes.db import (
    REAL_DB_PATH,
    TEST_DB_PATH,
    cleanup_test_db_files,
    clear_test_mode_state,
    get_db,
    get_real_db,
    get_test_mode_state,
    is_test_mode_active,
    set_test_mode_state,
)
from routes.helpers import current_username
from routes.security import (
    client_ip,
    has_any_permission,
    resolve_under_dir,
    user_can_access_student,
)
from routes.user_scope import get_allowed_group_ids
from routes.utils import (
    _short_user_agent,
    ensure_audit_logs_table,
    log_action,
    logger,
    login_required,
    permission_required,
)

system_ops_bp = Blueprint('system_ops', __name__)

BACKUPS_DIR = os.path.join(PROJECT_ROOT, 'backups')
SETTINGS_PATH = os.path.join(PROJECT_ROOT, 'system_ops_settings.json')
BACKUPS_META_PATH = os.path.join(BACKUPS_DIR, 'backups_meta.json')

_BACKUP_LOCK = threading.Lock()
_SETTINGS_LOCK = threading.Lock()
_SESSIONS_SCHEMA_LOCK = threading.Lock()
_SESSIONS_SCHEMA_READY = False

_SCHEDULER_LOCK = threading.Lock()
_SCHEDULER_THREAD: threading.Thread | None = None
_SCHEDULER_STOP = threading.Event()

_BACKUP_NAME_RE = re.compile(r'^[A-Za-z0-9_.-]+\.db$')

DEFAULT_SYSTEM_SETTINGS = {
    'auto_backup_enabled': True,
    'auto_backup_hour': 3,
    'auto_backup_keep_days': 14,
    'auto_backup_max_files': 30,
    'last_auto_backup_date': '',
    'last_auto_backup_at': '',
    'last_auto_backup_file': '',
    'last_auto_backup_error': '',
}


# ═══════════════════════════════════════════════════════════════════════════════
# 1. Налаштування резервування (поза БД — не зникають при відновленні чи тесті)
# ═══════════════════════════════════════════════════════════════════════════════

def _ensure_backups_dir() -> None:
    os.makedirs(BACKUPS_DIR, exist_ok=True)


def get_system_settings() -> dict:
    with _SETTINGS_LOCK:
        cfg = dict(DEFAULT_SYSTEM_SETTINGS)
        if os.path.exists(SETTINGS_PATH):
            try:
                with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    raw = json.load(f)
                if isinstance(raw, dict):
                    cfg.update(raw)
            except Exception as e:
                logger.warning(f"Не вдалося прочитати system_ops_settings.json: {e}")
        return cfg


def save_system_settings(updates: dict) -> dict:
    with _SETTINGS_LOCK:
        cfg = dict(DEFAULT_SYSTEM_SETTINGS)
        if os.path.exists(SETTINGS_PATH):
            try:
                with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    raw = json.load(f)
                if isinstance(raw, dict):
                    cfg.update(raw)
            except Exception:
                pass
        cfg.update(updates or {})
        tmp = f"{SETTINGS_PATH}.tmp.{os.getpid()}"
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        os.replace(tmp, SETTINGS_PATH)
        return cfg


def _load_backups_meta() -> dict:
    _ensure_backups_dir()
    if not os.path.exists(BACKUPS_META_PATH):
        return {}
    try:
        with open(BACKUPS_META_PATH, 'r', encoding='utf-8') as f:
            raw = json.load(f)
        return raw if isinstance(raw, dict) else {}
    except Exception:
        return {}


def _save_backups_meta(meta: dict) -> None:
    _ensure_backups_dir()
    tmp = f"{BACKUPS_META_PATH}.tmp.{os.getpid()}"
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    os.replace(tmp, BACKUPS_META_PATH)


def record_backup_cloud_upload(filename: str, results: list[dict]) -> None:
    """Фіксує результати вивантаження резервної копії у хмарні сховища."""
    if not filename or not results:
        return
    with _BACKUP_LOCK:
        meta = _load_backups_meta()
        entry = meta.get(filename)
        if not isinstance(entry, dict):
            entry = {'filename': filename}
        uploads = entry.get('cloud_uploads')
        if not isinstance(uploads, dict):
            uploads = {}
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for r in results:
            prov = r.get('provider')
            if not prov:
                continue
            uploads[prov] = {
                'ok': bool(r.get('ok')),
                'at': now_str,
                'message': (r.get('message') or '')[:240],
                'remote_name': r.get('remote_name') or '',
            }
        entry['cloud_uploads'] = uploads
        meta[filename] = entry
        _save_backups_meta(meta)


def _format_size_human(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} Б"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} КБ"
    if num_bytes < 1024 * 1024 * 1024:
        return f"{num_bytes / (1024 * 1024):.2f} МБ"
    return f"{num_bytes / (1024 * 1024 * 1024):.2f} ГБ"


# ═══════════════════════════════════════════════════════════════════════════════
# 2. Гаряче резервне копіювання SQLite (WAL-безпечне) та відновлення
# ═══════════════════════════════════════════════════════════════════════════════

def _copy_sqlite_hot(src_path: str, dst_path: str) -> None:
    """
    Виконує консистентне копіювання SQLite БД навіть під час активних запитів і WAL-режиму.
    """
    for side in (dst_path, f"{dst_path}-wal", f"{dst_path}-shm"):
        if os.path.exists(side):
            try:
                os.remove(side)
            except OSError:
                pass

    src_conn = None
    dst_conn = None
    try:
        src_conn = sqlite3.connect(src_path, timeout=30)
        dst_conn = sqlite3.connect(dst_path, timeout=30)
        src_conn.backup(dst_conn, pages=256, sleep=0.02)
        dst_conn.commit()
    finally:
        if dst_conn is not None:
            try:
                dst_conn.close()
            except Exception:
                pass
        if src_conn is not None:
            try:
                src_conn.close()
            except Exception:
                pass


def _verify_sqlite_file(db_path: str) -> tuple[bool, str]:
    """Перевіряє, що файл є валідною SQLite БД із ключовими таблицями."""
    conn = None
    try:
        conn = sqlite3.connect(db_path, timeout=15)
        qc = conn.execute("PRAGMA quick_check(1)").fetchone()
        if not qc or str(qc[0]).lower() != 'ok':
            return False, f"Перевірка цілісності повернула: {qc[0] if qc else 'помилка'}"
        tables = {
            r[0]
            for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        if 'users' not in tables or 'students' not in tables:
            return False, "У файлі відсутні обов'язкові таблиці (users / students)"
        return True, 'ok'
    except Exception as e:
        return False, str(e)
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass
        for side in (f"{db_path}-wal", f"{db_path}-shm"):
            if os.path.exists(side):
                try:
                    os.remove(side)
                except OSError:
                    pass


def create_sqlite_backup(
    kind: str = 'manual',
    created_by: str = 'system',
    note: str = '',
) -> tuple[bool, dict | None, str]:
    """
    Створює гарячу резервну копію РЕАЛЬНОЇ бази даних (`students.db`).
    Завжди читає реальну БД, навіть якщо зараз увімкнено тестовий режим.
    """
    if kind not in ('manual', 'auto', 'pre_restore', 'pre_test'):
        kind = 'manual'

    if not os.path.exists(REAL_DB_PATH):
        return False, None, 'Файл основної бази даних students.db не знайдено'

    _ensure_backups_dir()
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"students_{kind}_{stamp}.db"
    final_path = os.path.join(BACKUPS_DIR, filename)

    with _BACKUP_LOCK:
        # Унікалізуємо ім'я, якщо протягом тієї ж секунди вже був бекап
        idx = 1
        while os.path.exists(final_path):
            filename = f"students_{kind}_{stamp}_{idx}.db"
            final_path = os.path.join(BACKUPS_DIR, filename)
            idx += 1

        tmp_path = os.path.join(BACKUPS_DIR, f".tmp_{filename}.{secrets.token_hex(4)}")
        try:
            _copy_sqlite_hot(REAL_DB_PATH, tmp_path)
            ok_chk, chk_msg = _verify_sqlite_file(tmp_path)
            if not ok_chk:
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
                return False, None, f"Копія не пройшла перевірку: {chk_msg}"

            os.replace(tmp_path, final_path)
            st = os.stat(final_path)
            meta = _load_backups_meta()
            entry = {
                'filename': filename,
                'kind': kind,
                'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'created_by': created_by or 'system',
                'note': (note or '').strip()[:240],
                'size_bytes': st.st_size,
            }
            meta[filename] = entry
            _save_backups_meta(meta)

            if kind == 'auto':
                rotate_auto_backups()

            return True, entry, ''
        except Exception as e:
            logger.error(f"Помилка створення резервної копії БД: {e}", exc_info=True)
            for side in (tmp_path, f"{tmp_path}-wal", f"{tmp_path}-shm"):
                if os.path.exists(side):
                    try:
                        os.remove(side)
                    except OSError:
                        pass
            return False, None, str(e)


def rotate_auto_backups() -> int:
    """
    Видаляє застарілі автоматичні копії (`students_auto_*.db`) згідно з
    налаштуваннями `auto_backup_keep_days` та `auto_backup_max_files`,
    завжди зберігаючи мінімум 3 найсвіжіші автокопії. Ручні копії не чіпає.
    """
    cfg = get_system_settings()
    keep_days = max(1, int(cfg.get('auto_backup_keep_days') or 14))
    max_files = max(3, int(cfg.get('auto_backup_max_files') or 30))

    _ensure_backups_dir()
    auto_files = []
    for fn in os.listdir(BACKUPS_DIR):
        if fn.startswith('students_auto_') and fn.endswith('.db'):
            full = os.path.join(BACKUPS_DIR, fn)
            if os.path.isfile(full):
                auto_files.append((fn, full, os.path.getmtime(full)))

    auto_files.sort(key=lambda x: x[2], reverse=True)
    if len(auto_files) <= 3:
        return 0

    cutoff_ts = time.time() - keep_days * 86400
    deleted = 0
    meta = _load_backups_meta()
    meta_changed = False

    for pos, (fn, full, mtime) in enumerate(auto_files):
        if pos < 3:
            continue
        if pos >= max_files or mtime < cutoff_ts:
            try:
                os.remove(full)
                deleted += 1
                if fn in meta:
                    meta.pop(fn, None)
                    meta_changed = True
            except OSError as e:
                logger.warning(f"Не вдалося видалити старий бекап {fn}: {e}")

    if meta_changed:
        _save_backups_meta(meta)
    return deleted


def list_sqlite_backups() -> list[dict]:
    _ensure_backups_dir()
    meta = _load_backups_meta()
    kind_labels = {
        'manual': 'Ручна копія',
        'auto': 'Автоматична',
        'pre_restore': 'Перед відновленням',
        'pre_test': 'Перед тест-режимом',
        'external': 'Зовнішній файл',
    }
    items = []
    for fn in os.listdir(BACKUPS_DIR):
        if not fn.endswith('.db') or fn.startswith('.'):
            continue
        full = os.path.join(BACKUPS_DIR, fn)
        if not os.path.isfile(full):
            continue
        st = os.stat(full)
        info = meta.get(fn) or {}
        kind = info.get('kind')
        if not kind:
            if fn.startswith('students_auto_'):
                kind = 'auto'
            elif fn.startswith('students_pre_restore_'):
                kind = 'pre_restore'
            elif fn.startswith('students_pre_test_'):
                kind = 'pre_test'
            elif fn.startswith('students_manual_'):
                kind = 'manual'
            else:
                kind = 'external'

        created_at = info.get('created_at') or datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
        items.append({
            'filename': fn,
            'kind': kind,
            'kind_label': kind_labels.get(kind, kind),
            'created_at': created_at,
            'created_by': info.get('created_by') or '—',
            'note': info.get('note') or '',
            'size_bytes': st.st_size,
            'size_human': _format_size_human(st.st_size),
            'mtime': st.st_mtime,
            'cloud_uploads': info.get('cloud_uploads') if isinstance(info.get('cloud_uploads'), dict) else {},
        })

    items.sort(key=lambda x: x['mtime'], reverse=True)
    return items


def _resolve_backup_file(filename: str) -> str | None:
    if not filename or not _BACKUP_NAME_RE.fullmatch(filename):
        return None
    return resolve_under_dir(BACKUPS_DIR, filename)


def restore_sqlite_backup(filename: str, actor: str) -> tuple[bool, str]:
    """
    Відновлює реальну базу даних `students.db` із обраної копії.
    Спочатку автоматично створює страхувальну копію `pre_restore`, потім
    атомарно перезаписує сторінки реальної БД через SQLite Backup API.
    """
    if is_test_mode_active():
        return False, 'Спочатку вийдіть із тестового режиму, щоб відновити реальну БД'

    target_path = _resolve_backup_file(filename)
    if not target_path or not os.path.isfile(target_path):
        return False, 'Файл резервної копії не знайдено'

    ok_chk, chk_msg = _verify_sqlite_file(target_path)
    if not ok_chk:
        return False, f"Файл копії пошкоджено або він несумісний: {chk_msg}"

    # Створюємо страхувальний знімок перед відновленням
    ok_pre, pre_entry, pre_err = create_sqlite_backup(
        kind='pre_restore',
        created_by=actor,
        note=f"Автознімок перед відновленням із {filename}",
    )
    if not ok_pre:
        return False, f"Не вдалося створити страхувальну копію перед відновленням: {pre_err}"

    # Ставимо на паузу Telegram-воркер на час заміни сторінок БД
    bot_was_running = False
    try:
        from routes.military_telegram import get_worker_status, stop_bot_worker, ensure_bot_started_if_enabled
        bot_was_running = bool(get_worker_status().get('running'))
        if bot_was_running:
            stop_bot_worker(disable_in_db=False)
    except Exception:
        ensure_bot_started_if_enabled = None

    with _BACKUP_LOCK:
        src_conn = None
        dst_conn = None
        try:
            src_conn = sqlite3.connect(target_path, timeout=30)
            dst_conn = sqlite3.connect(REAL_DB_PATH, timeout=30)
            src_conn.backup(dst_conn, pages=256, sleep=0.02)
            dst_conn.commit()
            try:
                dst_conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            except Exception:
                pass
        except Exception as e:
            logger.error(f"Помилка відновлення БД із {filename}: {e}", exc_info=True)
            return False, f"Помилка відновлення БД: {e}"
        finally:
            if dst_conn is not None:
                try:
                    dst_conn.close()
                except Exception:
                    pass
            if src_conn is not None:
                try:
                    src_conn.close()
                except Exception:
                    pass

    if bot_was_running and ensure_bot_started_if_enabled:
        try:
            ensure_bot_started_if_enabled()
        except Exception:
            pass

    pre_name = pre_entry['filename'] if pre_entry else '—'
    return True, f"Базу даних успішно відновлено з «{filename}». Попередній стан збережено у «{pre_name}»."


def _backup_scheduler_loop() -> None:
    logger.info("Системний планувальник авторезервування БД запущено")
    while not _SCHEDULER_STOP.is_set():
        try:
            cfg = get_system_settings()
            if cfg.get('auto_backup_enabled'):
                now = datetime.now()
                today_iso = now.strftime('%Y-%m-%d')
                target_hour = int(cfg.get('auto_backup_hour') if cfg.get('auto_backup_hour') is not None else 3)
                if now.hour >= target_hour and cfg.get('last_auto_backup_date') != today_iso:
                    ok, entry, err = create_sqlite_backup(
                        kind='auto',
                        created_by='auto_scheduler',
                        note='Щоденна автоматична резервна копія',
                    )
                    if ok and entry:
                        save_system_settings({
                            'last_auto_backup_date': today_iso,
                            'last_auto_backup_at': entry['created_at'],
                            'last_auto_backup_file': entry['filename'],
                            'last_auto_backup_error': '',
                        })
                        logger.info(f"Авторезервування БД завершено: {entry['filename']}")
                        try:
                            cloud_results = upload_backup_to_enabled_providers(
                                entry['filename'],
                                trigger='auto',
                                only_auto=True,
                            )
                            if cloud_results:
                                record_backup_cloud_upload(entry['filename'], cloud_results)
                                for cr in cloud_results:
                                    if cr['ok']:
                                        logger.info(f"Авто-вивантаження у {cr['label']}: {cr['message']}")
                                    else:
                                        logger.warning(f"Помилка авто-вивантаження у {cr['label']}: {cr['message']}")
                        except Exception as ce:
                            logger.error(f"Помилка хмарного авто-вивантаження: {ce}", exc_info=True)
                    else:
                        save_system_settings({
                            'last_auto_backup_error': err or 'Невідома помилка',
                        })
                        logger.error(f"Помилка авторезервування БД: {err}")
        except Exception as e:
            logger.error(f"Помилка планувальника авторезервування БД: {e}", exc_info=True)

        _SCHEDULER_STOP.wait(60.0)


def ensure_backup_scheduler_started() -> None:
    """Одноразово запускає фоновий потік щоденного авторезервування БД."""
    global _SCHEDULER_THREAD
    with _SCHEDULER_LOCK:
        if _SCHEDULER_THREAD and _SCHEDULER_THREAD.is_alive():
            return
        # Якщо після аварійного виходу лишились хвости неактивного тест-режиму — приберемо
        cleanup_test_db_files(retries=2)
        _SCHEDULER_STOP.clear()
        _SCHEDULER_THREAD = threading.Thread(
            target=_backup_scheduler_loop,
            name='sqlite-auto-backup',
            daemon=True,
        )
        _SCHEDULER_THREAD.start()


# ═══════════════════════════════════════════════════════════════════════════════
# 3. Тестовий режим («Пісочниця БД»)
# ═══════════════════════════════════════════════════════════════════════════════

def enter_test_mode(username: str, user_id: int | None) -> tuple[bool, str]:
    if is_test_mode_active():
        return False, 'Тестовий режим уже увімкнено'
    if not os.path.exists(REAL_DB_PATH):
        return False, 'Основний файл БД students.db не знайдено'

    # Зупиняємо фонового Telegram-бота на час тестування (без вимкнення в реальній БД)
    tg_was_running = False
    try:
        from routes.military_telegram import get_worker_status, stop_bot_worker
        tg_was_running = bool(get_worker_status().get('running'))
        if tg_was_running:
            stop_bot_worker(disable_in_db=False)
    except Exception:
        pass

    with _BACKUP_LOCK:
        try:
            _copy_sqlite_hot(REAL_DB_PATH, TEST_DB_PATH)
            ok_chk, chk_msg = _verify_sqlite_file(TEST_DB_PATH)
            if not ok_chk:
                cleanup_test_db_files(retries=3)
                return False, f"Не вдалося підготувати тестову копію БД: {chk_msg}"

            set_test_mode_state({
                'active': True,
                'started_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'started_by': username or 'admin',
                'started_by_user_id': user_id,
                'telegram_was_running': tg_was_running,
            })
        except Exception as e:
            logger.error(f"Помилка запуску тестового режиму: {e}", exc_info=True)
            clear_test_mode_state()
            cleanup_test_db_files(retries=3)
            return False, f"Не вдалося створити копію БД: {e}"

    log_action(
        username,
        "увімкнув тестовий режим (створено ізольовану копію БД)",
        level='WARNING',
        category='system',
        section='Система',
        details=f"копія БД: {os.path.basename(TEST_DB_PATH)}",
    )
    return True, 'Тестовий режим увімкнено! Усі дії тепер виконуються в копії БД і не вплинуть на реальну базу.'


def reset_test_mode_copy(username: str) -> tuple[bool, str]:
    """Перезаливає тестову копію зі свіжого знімка реальної БД, лишаючись у тест-режимі."""
    state = get_test_mode_state()
    if not state.get('active'):
        return False, 'Тестовий режим зараз не активний'

    with _BACKUP_LOCK:
        tmp_sandbox = f"{TEST_DB_PATH}.refresh"
        try:
            _copy_sqlite_hot(REAL_DB_PATH, tmp_sandbox)
            ok_chk, chk_msg = _verify_sqlite_file(tmp_sandbox)
            if not ok_chk:
                return False, f"Помилка перевірки копії: {chk_msg}"

            # Копіюємо сторінки в наявний файл, щоб не ламати відкриті дескриптори
            _copy_sqlite_hot(tmp_sandbox, TEST_DB_PATH)
            for p in (tmp_sandbox, f"{tmp_sandbox}-wal", f"{tmp_sandbox}-shm"):
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except OSError:
                        pass

            state['reset_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            state['reset_by'] = username or 'admin'
            set_test_mode_state(state)
        except Exception as e:
            logger.error(f"Помилка оновлення тестової копії: {e}", exc_info=True)
            return False, f"Не вдалося оновити тестову копію: {e}"

    log_action(
        username,
        "скинув тестову копію БД до стану реальної бази",
        level='WARNING',
        category='system',
        section='Система',
    )
    return True, 'Тестову копію БД оновлено до актуального стану реальної бази.'


def exit_test_mode(username: str) -> tuple[bool, str]:
    state = get_test_mode_state()
    if not state.get('active'):
        cleanup_test_db_files(retries=3)
        clear_test_mode_state()
        return True, 'Тестовий режим уже вимкнено — використовується реальна база даних.'

    clear_test_mode_state()
    gc.collect()
    cleanup_test_db_files(retries=10)

    # Примусово оновимо кеш прав поточної сесії вже з реальної БД
    try:
        from flask import has_request_context
        if has_request_context():
            session.pop('perm_check_at', None)
            if hasattr(session, 'modified'):
                session.modified = True
    except Exception:
        pass

    # Повертаємо Telegram-бота, якщо він увімкнений у реальній БД
    try:
        from routes.military_telegram import ensure_bot_started_if_enabled
        ensure_bot_started_if_enabled()
    except Exception as e:
        logger.warning(f"Не вдалося відновити роботу Telegram-бота після тест-режиму: {e}")

    log_action(
        username,
        "вимкнув тестовий режим і повернув реальну базу даних",
        level='INFO',
        category='system',
        section='Система',
        details=f"тест тривав з {state.get('started_at', '—')}",
    )
    return True, 'Тестовий режим завершено! Повернуто реальну базу даних, усі тестові зміни скасовано.'


# ═══════════════════════════════════════════════════════════════════════════════
# 4. Активні сесії користувачів (у реальній БД `students.db`)
# ═══════════════════════════════════════════════════════════════════════════════

def ensure_user_sessions_table(conn) -> None:
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_sessions (
            sid TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL,
            username TEXT NOT NULL,
            role TEXT,
            is_admin INTEGER NOT NULL DEFAULT 0,
            ip_address TEXT,
            user_agent TEXT,
            current_path TEXT,
            created_at TEXT NOT NULL,
            last_seen_at TEXT NOT NULL,
            is_revoked INTEGER NOT NULL DEFAULT 0,
            revoked_by TEXT,
            revoked_at TEXT
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_user_sessions_user ON user_sessions(user_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_user_sessions_seen ON user_sessions(last_seen_at)")


def _ensure_sessions_schema_once(conn) -> None:
    global _SESSIONS_SCHEMA_READY
    if _SESSIONS_SCHEMA_READY:
        return
    with _SESSIONS_SCHEMA_LOCK:
        if _SESSIONS_SCHEMA_READY:
            return
        ensure_user_sessions_table(conn)
        conn.commit()
        _SESSIONS_SCHEMA_READY = True


def register_user_session(
    user_id: int,
    username: str,
    role: str | None,
    is_admin: bool,
    ip_address: str | None = None,
    user_agent: str | None = None,
    current_path: str | None = None,
) -> str:
    sid = secrets.token_urlsafe(24)
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ua_short = _short_user_agent(user_agent) or (str(user_agent)[:80] if user_agent else 'Невідомий браузер')
    ip_val = ip_address or client_ip()

    conn = None
    try:
        conn = get_real_db()
        _ensure_sessions_schema_once(conn)
        conn.execute(
            """
            INSERT INTO user_sessions (
                sid, user_id, username, role, is_admin,
                ip_address, user_agent, current_path,
                created_at, last_seen_at, is_revoked
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
            """,
            (
                sid,
                int(user_id),
                str(username or 'user'),
                role or ('admin' if is_admin else 'user'),
                1 if is_admin else 0,
                ip_val,
                ua_short,
                (current_path or '/')[:200],
                now_str,
                now_str,
            ),
        )
        conn.commit()
    except Exception as e:
        logger.warning(f"Не вдалося зареєструвати сесію користувача {username}: {e}")
    finally:
        if conn is not None:
            conn.close()
    return sid


def mark_session_logged_out(sid: str | None, actor: str | None = None) -> None:
    if not sid:
        return
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = None
    try:
        conn = get_real_db()
        _ensure_sessions_schema_once(conn)
        conn.execute(
            """
            UPDATE user_sessions
            SET is_revoked = 1,
                revoked_by = COALESCE(revoked_by, ?),
                revoked_at = COALESCE(revoked_at, ?)
            WHERE sid = ?
            """,
            (actor or 'logout', now_str, sid),
        )
        conn.commit()
    except Exception as e:
        logger.warning(f"Не вдалося закрити запис сесії {sid}: {e}")
    finally:
        if conn is not None:
            conn.close()


def enforce_and_touch_session():
    """
    Викликається з @app.before_request:
      - Якщо поточну сесію відкликано адміністратором — миттєво скидає сесію і повертає на /login.
      - Інакше періодично (раз на 30 сек) оновлює last_seen_at, IP і поточний шлях.
    """
    if request.path.startswith('/static/'):
        return None
    user_id = session.get('user_id')
    if not user_id:
        return None

    sid = session.get('session_sid')
    if not sid:
        sid = register_user_session(
            user_id=user_id,
            username=session.get('username') or 'user',
            role=session.get('role'),
            is_admin=bool(session.get('is_admin', False)),
            ip_address=client_ip(),
            user_agent=request.headers.get('User-Agent', ''),
            current_path=request.path,
        )
        session['session_sid'] = sid
        session['_sid_touch_at'] = time.time()
        session.modified = True
        return None

    conn = None
    try:
        conn = get_real_db()
        _ensure_sessions_schema_once(conn)
        row = conn.execute(
            "SELECT is_revoked, revoked_by FROM user_sessions WHERE sid = ?",
            (sid,),
        ).fetchone()

        if row is not None and int(row['is_revoked'] or 0) == 1:
            session.clear()
            if request.path.startswith('/api/'):
                return jsonify({
                    'ok': False,
                    'error': 'Ваш сеанс було завершено адміністратором.',
                }), 401
            flash('Ваш сеанс було завершено адміністратором. Увійдіть повторно.', 'warning')
            return redirect(url_for('auth.login'))

        now_ts = time.time()
        last_touch = float(session.get('_sid_touch_at') or 0)
        if row is None:
            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ua_short = _short_user_agent(request.headers.get('User-Agent', '')) or 'Браузер'
            conn.execute(
                """
                INSERT OR IGNORE INTO user_sessions (
                    sid, user_id, username, role, is_admin,
                    ip_address, user_agent, current_path,
                    created_at, last_seen_at, is_revoked
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
                """,
                (
                    sid,
                    int(user_id),
                    str(session.get('username') or 'user'),
                    session.get('role') or 'user',
                    1 if session.get('is_admin') else 0,
                    client_ip(),
                    ua_short,
                    request.path[:200],
                    now_str,
                    now_str,
                ),
            )
            conn.commit()
            session['_sid_touch_at'] = now_ts
            session.modified = True
        elif now_ts - last_touch >= 30:
            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            conn.execute(
                """
                UPDATE user_sessions
                SET last_seen_at = ?, ip_address = ?, current_path = ?
                WHERE sid = ? AND is_revoked = 0
                """,
                (now_str, client_ip(), request.path[:200], sid),
            )
            conn.commit()
            session['_sid_touch_at'] = now_ts
            session.modified = True
    except Exception as e:
        logger.warning(f"Помилка перевірки активної сесії: {e}")
    finally:
        if conn is not None:
            conn.close()
    return None


def list_active_sessions(current_sid: str | None = None) -> list[dict]:
    conn = None
    try:
        conn = get_real_db()
        _ensure_sessions_schema_once(conn)
        cutoff = (datetime.now() - timedelta(hours=14)).strftime('%Y-%m-%d %H:%M:%S')
        rows = conn.execute(
            """
            SELECT sid, user_id, username, role, is_admin,
                   ip_address, user_agent, current_path,
                   created_at, last_seen_at
            FROM user_sessions
            WHERE is_revoked = 0 AND last_seen_at >= ?
            ORDER BY last_seen_at DESC
            """,
            (cutoff,),
        ).fetchall()

        now = datetime.now()
        out = []
        for r in rows:
            d = dict(r)
            d['is_current'] = bool(current_sid and d['sid'] == current_sid)
            d['online_recent'] = False
            d['idle_label'] = '—'
            try:
                dt_seen = datetime.strptime(d['last_seen_at'], '%Y-%m-%d %H:%M:%S')
                diff_sec = max(0, int((now - dt_seen).total_seconds()))
                d['online_recent'] = diff_sec <= 300
                if diff_sec < 60:
                    d['idle_label'] = 'щойно'
                elif diff_sec < 3600:
                    d['idle_label'] = f"{diff_sec // 60} хв тому"
                else:
                    d['idle_label'] = f"{diff_sec // 3600} год тому"
            except Exception:
                pass
            out.append(d)
        return out
    finally:
        if conn is not None:
            conn.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 5. Веб-маршрути адміністрування (/admin/system)
# ═══════════════════════════════════════════════════════════════════════════════

@system_ops_bp.route('/admin/system')
@permission_required('manage_system')
def system_dashboard():
    tab = (request.args.get('tab') or 'sandbox').strip()
    if tab not in ('sandbox', 'backups', 'cloud', 'sessions'):
        tab = 'sandbox'

    settings = get_system_settings()
    backups = list_sqlite_backups()
    test_mode = get_test_mode_state()
    active_sessions = list_active_sessions(current_sid=session.get('session_sid'))
    cloud_config = get_cloud_config_masked()
    cloud_enabled_count = sum(1 for p in PROVIDER_ORDER if cloud_config[p].get('enabled'))

    real_db_size = _format_size_human(os.path.getsize(REAL_DB_PATH)) if os.path.exists(REAL_DB_PATH) else '—'
    test_db_size = _format_size_human(os.path.getsize(TEST_DB_PATH)) if os.path.exists(TEST_DB_PATH) else '—'

    gdrive_callback_url = url_for('system_ops.cloud_gdrive_callback_route', _external=True)
    host_only = (request.host or '').split(':', 1)[0].lower()
    is_local_intranet_host = host_only.endswith('.local') or host_only in ('localhost', '127.0.0.1')
    gdrive_manual_auth_url = session.get('gdrive_manual_auth_url', '')

    return render_template(
        'system_tools.html',
        active_tab=tab,
        settings=settings,
        backups=backups,
        test_mode=test_mode,
        active_sessions=active_sessions,
        real_db_size=real_db_size,
        test_db_size=test_db_size,
        cloud_config=cloud_config,
        cloud_provider_order=PROVIDER_ORDER,
        cloud_provider_labels=PROVIDER_LABELS,
        cloud_enabled_count=cloud_enabled_count,
        gdrive_callback_url=gdrive_callback_url,
        is_local_intranet_host=is_local_intranet_host,
        gdrive_manual_auth_url=gdrive_manual_auth_url,
    )


@system_ops_bp.route('/admin/system/test-mode/start', methods=['POST'])
@permission_required('manage_system')
def start_test_mode_route():
    ok, msg = enter_test_mode(current_username(), session.get('user_id'))
    flash(msg, 'warning' if ok else 'danger')
    return redirect(request.form.get('next') or url_for('system_ops.system_dashboard', tab='sandbox'))


@system_ops_bp.route('/admin/system/test-mode/reset', methods=['POST'])
@permission_required('manage_system')
def reset_test_mode_route():
    ok, msg = reset_test_mode_copy(current_username())
    flash(msg, 'success' if ok else 'danger')
    return redirect(request.form.get('next') or url_for('system_ops.system_dashboard', tab='sandbox'))


@system_ops_bp.route('/admin/system/test-mode/stop', methods=['POST'])
@permission_required('manage_system')
def stop_test_mode_route():
    ok, msg = exit_test_mode(current_username())
    flash(msg, 'success' if ok else 'danger')
    next_url = request.form.get('next') or ''
    if next_url and next_url.startswith('/') and not next_url.startswith('//'):
        return redirect(next_url)
    return redirect(url_for('system_ops.system_dashboard', tab='sandbox'))


@system_ops_bp.route('/admin/system/backups/create', methods=['POST'])
@permission_required('manage_system')
def create_backup_route():
    note = (request.form.get('note') or '').strip()
    upload_to_cloud = request.form.get('upload_to_cloud') == '1'
    ok, entry, err = create_sqlite_backup(
        kind='manual',
        created_by=current_username(),
        note=note,
    )
    if ok and entry:
        log_action(
            current_username(),
            f"створив резервну копію БД: {entry['filename']}",
            category='system',
            section='Система',
            details=note or f"розмір: {_format_size_human(entry['size_bytes'])}",
        )
        flash(f"Резервну копію «{entry['filename']}» створено ({_format_size_human(entry['size_bytes'])}).", 'success')
        if upload_to_cloud:
            results = upload_backup_to_enabled_providers(entry['filename'], trigger='manual', only_auto=False)
            if results:
                record_backup_cloud_upload(entry['filename'], results)
                ok_list = [r['label'] for r in results if r['ok']]
                err_list = [f"{r['label']}: {r['message']}" for r in results if not r['ok']]
                if ok_list:
                    flash(f"Успішно вивантажено в хмару: {', '.join(ok_list)}.", 'success')
                if err_list:
                    flash("Помилки вивантаження в хмару: " + ' | '.join(err_list), 'danger')
            else:
                flash('Жодне хмарне сховище ще не увімкнено у вкладці «Хмарні копії».', 'warning')
    else:
        flash(f"Не вдалося створити резервну копію: {err}", 'danger')
    return redirect(url_for('system_ops.system_dashboard', tab='backups'))


@system_ops_bp.route('/admin/system/backups/settings', methods=['POST'])
@permission_required('manage_system')
def update_backup_settings_route():
    enabled = request.form.get('auto_backup_enabled') == '1'
    hour = request.form.get('auto_backup_hour', type=int)
    keep_days = request.form.get('auto_backup_keep_days', type=int)
    max_files = request.form.get('auto_backup_max_files', type=int)

    hour = 3 if hour is None else max(0, min(23, hour))
    keep_days = 14 if keep_days is None else max(1, min(365, keep_days))
    max_files = 30 if max_files is None else max(3, min(200, max_files))

    save_system_settings({
        'auto_backup_enabled': enabled,
        'auto_backup_hour': hour,
        'auto_backup_keep_days': keep_days,
        'auto_backup_max_files': max_files,
    })
    rotate_auto_backups()
    log_action(
        current_username(),
        "оновив налаштування авторезервування БД",
        category='system',
        section='Система',
        details=f"увімкнено={enabled}, година={hour:02d}:00, зберігати_днів={keep_days}, макс_файлів={max_files}",
    )
    flash('Налаштування автоматичного резервування збережено.', 'success')
    return redirect(url_for('system_ops.system_dashboard', tab='backups'))


@system_ops_bp.route('/admin/system/backups/<filename>/download')
@permission_required('manage_system')
def download_backup_route(filename):
    path = _resolve_backup_file(filename)
    if not path or not os.path.isfile(path):
        abort(404)
    log_action(
        current_username(),
        f"завантажив резервну копію БД: {filename}",
        category='export',
        section='Система',
    )
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@system_ops_bp.route('/admin/system/backups/<filename>/restore', methods=['POST'])
@permission_required('manage_system')
def restore_backup_route(filename):
    ok, msg = restore_sqlite_backup(filename, current_username())
    if ok:
        log_action(
            current_username(),
            f"відновив БД із резервної копії: {filename}",
            level='WARNING',
            category='system',
            section='Система',
        )
        flash(msg, 'success')
    else:
        flash(msg, 'danger')
    return redirect(url_for('system_ops.system_dashboard', tab='backups'))


@system_ops_bp.route('/admin/system/backups/<filename>/delete', methods=['POST'])
@permission_required('manage_system')
def delete_backup_route(filename):
    path = _resolve_backup_file(filename)
    if not path or not os.path.isfile(path):
        flash('Файл резервної копії не знайдено.', 'danger')
        return redirect(url_for('system_ops.system_dashboard', tab='backups'))

    try:
        os.remove(path)
        meta = _load_backups_meta()
        if filename in meta:
            meta.pop(filename, None)
            _save_backups_meta(meta)
        log_action(
            current_username(),
            f"видалив резервну копію БД: {filename}",
            level='WARNING',
            category='delete',
            section='Система',
        )
        flash(f"Резервну копію «{filename}» видалено.", 'success')
    except Exception as e:
        flash(f"Не вдалося видалити файл: {e}", 'danger')
    return redirect(url_for('system_ops.system_dashboard', tab='backups'))


@system_ops_bp.route('/admin/system/backups/<filename>/upload-cloud', methods=['POST'])
@permission_required('manage_system')
def upload_backup_cloud_route(filename):
    provider = (request.form.get('provider') or 'all').strip()
    next_tab = request.form.get('next_tab') or 'backups'
    if next_tab not in ('backups', 'cloud'):
        next_tab = 'backups'

    if provider == 'all':
        results = upload_backup_to_enabled_providers(filename, trigger='manual', only_auto=False)
        if not results:
            flash('Немає жодного увімкненого та налаштованого хмарного сховища.', 'warning')
            return redirect(url_for('system_ops.system_dashboard', tab=next_tab))
    else:
        ok, msg, remote_name = upload_backup_to_provider(filename, provider, trigger='manual')
        results = [{
            'provider': provider,
            'label': PROVIDER_LABELS.get(provider, provider),
            'ok': ok,
            'message': msg,
            'remote_name': remote_name,
        }]

    record_backup_cloud_upload(filename, results)
    ok_msgs = [f"{r['label']}: {r['message']}" for r in results if r['ok']]
    err_msgs = [f"{r['label']}: {r['message']}" for r in results if not r['ok']]

    if ok_msgs:
        log_action(
            current_username(),
            f"вивантажив резервну копію БД {filename} у хмару",
            category='system',
            section='Система',
            details='; '.join(ok_msgs)[:250],
        )
        flash(' · '.join(ok_msgs), 'success')
    if err_msgs:
        flash(' · '.join(err_msgs), 'danger')
    return redirect(url_for('system_ops.system_dashboard', tab=next_tab))


@system_ops_bp.route('/admin/system/cloud/upload-latest', methods=['POST'])
@permission_required('manage_system')
def cloud_upload_latest_route():
    provider = (request.form.get('provider') or 'all').strip()
    create_new = request.form.get('create_new') == '1'

    target_filename = ''
    if create_new:
        ok, entry, err = create_sqlite_backup(
            kind='manual',
            created_by=current_username(),
            note='Знімок перед хмарним вивантаженням',
        )
        if not ok or not entry:
            flash(f"Не вдалося створити свіжу копію БД: {err}", 'danger')
            return redirect(url_for('system_ops.system_dashboard', tab='cloud'))
        target_filename = entry['filename']
    else:
        backups = list_sqlite_backups()
        if not backups:
            flash('Локальних копій ще немає — спочатку створіть резервну копію або увімкніть опцію «Створити нову».', 'warning')
            return redirect(url_for('system_ops.system_dashboard', tab='cloud'))
        target_filename = backups[0]['filename']

    if provider == 'all':
        results = upload_backup_to_enabled_providers(target_filename, trigger='manual', only_auto=False)
        if not results:
            flash('Спочатку налаштуйте й увімкніть хоча б одне хмарне сховище.', 'warning')
            return redirect(url_for('system_ops.system_dashboard', tab='cloud'))
    else:
        ok, msg, remote_name = upload_backup_to_provider(target_filename, provider, trigger='manual')
        results = [{
            'provider': provider,
            'label': PROVIDER_LABELS.get(provider, provider),
            'ok': ok,
            'message': msg,
            'remote_name': remote_name,
        }]

    record_backup_cloud_upload(target_filename, results)
    ok_msgs = [f"{r['label']}: {r['message']}" for r in results if r['ok']]
    err_msgs = [f"{r['label']}: {r['message']}" for r in results if not r['ok']]
    if ok_msgs:
        log_action(
            current_username(),
            f"вивантажив резервну копію {target_filename} у хмарне сховище",
            category='system',
            section='Система',
            details='; '.join(ok_msgs)[:250],
        )
        flash(' · '.join(ok_msgs), 'success')
    if err_msgs:
        flash(' · '.join(err_msgs), 'danger')
    return redirect(url_for('system_ops.system_dashboard', tab='cloud'))


@system_ops_bp.route('/admin/system/cloud/<provider>/save', methods=['POST'])
@permission_required('manage_system')
def cloud_save_provider_route(provider):
    if provider not in PROVIDER_ORDER:
        abort(404)

    form_data = dict(request.form)
    form_data['enabled'] = request.form.get('enabled') == '1'
    form_data['auto_upload'] = request.form.get('auto_upload') == '1'
    form_data['compress_zip'] = request.form.get('compress_zip') == '1'
    form_data['path_style'] = request.form.get('path_style') == '1'
    form_data['verify_ssl'] = request.form.get('verify_ssl') == '1'

    if provider == 'gdrive':
        sa_file = request.files.get('service_account_file')
        if sa_file and sa_file.filename:
            raw_bytes = sa_file.read(128 * 1024 + 1)
            if len(raw_bytes) > 128 * 1024:
                flash('JSON-файл Service Account занадто великий (макс. 128 КБ).', 'danger')
                return redirect(url_for('system_ops.system_dashboard', tab='cloud'))
            try:
                form_data['service_account_json'] = raw_bytes.decode('utf-8')
            except UnicodeDecodeError:
                flash('Файл Service Account повинен бути у кодуванні UTF-8 JSON.', 'danger')
                return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    ok, msg = save_cloud_provider_config(provider, form_data)
    if ok:
        log_action(
            current_username(),
            f"оновив налаштування хмарного сховища: {PROVIDER_LABELS[provider]}",
            category='system',
            section='Система',
            details=f"enabled={form_data['enabled']}, auto_upload={form_data['auto_upload']}",
        )
        flash(msg, 'success')
    else:
        flash(msg, 'warning')
    return redirect(url_for('system_ops.system_dashboard', tab='cloud'))


@system_ops_bp.route('/admin/system/cloud/<provider>/test', methods=['POST'])
@permission_required('manage_system')
def cloud_test_provider_route(provider):
    if provider not in PROVIDER_ORDER:
        abort(404)
    ok, msg = test_cloud_provider(provider)
    log_action(
        current_username(),
        f"перевірив з'єднання з хмарним сховищем: {PROVIDER_LABELS[provider]}",
        level='INFO' if ok else 'WARNING',
        category='system',
        section='Система',
        details=msg[:240],
    )
    flash(msg, 'success' if ok else 'danger')
    return redirect(url_for('system_ops.system_dashboard', tab='cloud'))


@system_ops_bp.route('/admin/system/cloud/<provider>/disconnect', methods=['POST'])
@permission_required('manage_system')
def cloud_disconnect_provider_route(provider):
    if provider not in PROVIDER_ORDER:
        abort(404)
    ok, msg = disconnect_cloud_provider(provider)
    if provider == 'gdrive':
        session.pop('gdrive_manual_auth_url', None)
        session.pop('gdrive_oauth_state', None)
        session.pop('gdrive_oauth_redirect_uri', None)
    log_action(
        current_username(),
        f"відключив хмарне сховище: {PROVIDER_LABELS[provider]}",
        level='WARNING',
        category='system',
        section='Система',
    )
    flash(msg, 'info' if ok else 'danger')
    return redirect(url_for('system_ops.system_dashboard', tab='cloud'))


@system_ops_bp.route('/admin/system/cloud/gdrive/auth-start', methods=['POST'])
@permission_required('manage_system')
def cloud_gdrive_auth_start_route():
    # За потреби зберігаємо client_id / client_secret прямо зі стартової форми OAuth
    cur = get_cloud_config_raw()['gdrive']
    client_id = (request.form.get('client_id') or '').strip()
    client_secret = (request.form.get('client_secret') or '').strip()
    folder_id = (request.form.get('folder_id') or '').strip()
    if client_id or client_secret or folder_id:
        save_cloud_provider_config('gdrive', {
            'enabled': bool(cur.get('enabled', False)),
            'auto_upload': bool(cur.get('auto_upload', False)),
            'compress_zip': bool(cur.get('compress_zip', True)),
            'auth_mode': 'oauth',
            'client_id': client_id or cur.get('client_id', ''),
            'client_secret': client_secret,
            'folder_id': folder_id if folder_id else cur.get('folder_id', ''),
            'keep_remote': cur.get('keep_remote', 14),
        })

    cfg = get_cloud_config_raw()['gdrive']
    cid = (cfg.get('client_id') or '').strip()
    csec = (cfg.get('client_secret') or '').strip()
    if not cid or not csec:
        flash('Спочатку введіть і збережіть Google Client ID та Client Secret.', 'warning')
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    redirect_mode = (request.form.get('redirect_mode') or 'auto').strip()
    if redirect_mode == 'manual_loopback':
        redirect_uri = 'http://127.0.0.1'
    else:
        redirect_uri = url_for('system_ops.cloud_gdrive_callback_route', _external=True)

    state = secrets.token_urlsafe(24)
    session['gdrive_oauth_state'] = state
    session['gdrive_oauth_redirect_uri'] = redirect_uri
    auth_url = build_gdrive_auth_url(cid, redirect_uri, state)

    if redirect_mode == 'manual_loopback':
        session['gdrive_manual_auth_url'] = auth_url
        flash(
            'Посилання для входу в Google згенеровано нижче. Перейдіть за ним, підтвердіть доступ, '
            'а коли браузер переадресує на http://127.0.0.1/?code=... — скопіюйте повний URL із адресного рядка та вставте в крок 2.',
            'info',
        )
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    session.pop('gdrive_manual_auth_url', None)
    return redirect(auth_url)


@system_ops_bp.route('/admin/system/cloud/gdrive/callback')
@permission_required('manage_system')
def cloud_gdrive_callback_route():
    err_param = (request.args.get('error') or '').strip()
    if err_param:
        flash(f"Google відхилив авторизацію: {err_param}", 'danger')
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    expected_state = session.pop('gdrive_oauth_state', '')
    got_state = (request.args.get('state') or '').strip()
    if not expected_state or not got_state or not hmac.compare_digest(expected_state, got_state):
        flash('Невалідний або прострочений параметр безпеки OAuth (state). Спробуйте підключити ще раз.', 'danger')
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    code = (request.args.get('code') or '').strip()
    if not code:
        flash('Google не повернув код авторизації.', 'danger')
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    redirect_uri = session.pop('gdrive_oauth_redirect_uri', '') or url_for('system_ops.cloud_gdrive_callback_route', _external=True)
    return _finish_gdrive_oauth_exchange(code, redirect_uri)


@system_ops_bp.route('/admin/system/cloud/gdrive/exchange-code', methods=['POST'])
@permission_required('manage_system')
def cloud_gdrive_exchange_code_route():
    raw_input = request.form.get('oauth_code_or_url') or ''
    code, url_state = extract_oauth_code_from_input(raw_input)
    if not code:
        flash('Не знайдено параметр code у вставленому рядку. Скопіюйте повний URL після редіректу Google або сам код.', 'warning')
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    expected_state = session.get('gdrive_oauth_state', '')
    if url_state and expected_state and not hmac.compare_digest(expected_state, url_state):
        flash('Параметр безпеки state у вставленому URL не збігається з поточною сесією.', 'danger')
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    session.pop('gdrive_oauth_state', None)
    session.pop('gdrive_manual_auth_url', None)
    redirect_uri = session.pop('gdrive_oauth_redirect_uri', '') or 'http://127.0.0.1'
    return _finish_gdrive_oauth_exchange(code, redirect_uri)


def _finish_gdrive_oauth_exchange(code: str, redirect_uri: str):
    cfg = get_cloud_config_raw()['gdrive']
    cid = (cfg.get('client_id') or '').strip()
    csec = (cfg.get('client_secret') or '').strip()
    if not cid or not csec:
        flash('Відсутні Client ID або Client Secret Google.', 'danger')
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    ok, token_data, err = exchange_gdrive_oauth_code(cid, csec, code, redirect_uri)
    if not ok:
        flash(err, 'danger')
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    rtok = (token_data.get('refresh_token') or '').strip() or cfg.get('refresh_token', '')
    if not rtok:
        flash(
            'Google повернув access_token без refresh_token. Видаліть доступ застосунку в налаштуваннях Google-акаунту '
            'або повторіть вхід із вибором акаунту.',
            'warning',
        )
        return redirect(url_for('system_ops.system_dashboard', tab='cloud'))

    save_cloud_provider_config('gdrive', {
        'enabled': True,
        'auto_upload': bool(cfg.get('auto_upload', False)),
        'compress_zip': bool(cfg.get('compress_zip', True)),
        'auth_mode': 'oauth',
        'client_id': cid,
        'client_secret': csec,
        'refresh_token': rtok,
        'folder_id': cfg.get('folder_id', ''),
        'keep_remote': cfg.get('keep_remote', 14),
    })
    ok_t, msg_t = test_cloud_provider('gdrive')
    log_action(
        current_username(),
        "підключив Google Диск для хмарного резервування БД",
        category='system',
        section='Система',
        details=msg_t[:240],
    )
    flash(f"Google Диск підключено! {msg_t}" if ok_t else f"Токен збережено, але перевірка повернула: {msg_t}", 'success' if ok_t else 'warning')
    return redirect(url_for('system_ops.system_dashboard', tab='cloud'))


@system_ops_bp.route('/admin/system/sessions/<sid>/revoke', methods=['POST'])
@permission_required('manage_system')
def revoke_session_route(sid):
    current_sid = session.get('session_sid')
    if current_sid and sid == current_sid:
        flash('Не можна примусово завершити власну поточну сесію (скористайтеся кнопкою «Вийти»).', 'warning')
        return redirect(url_for('system_ops.system_dashboard', tab='sessions'))

    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    actor = current_username()
    conn = get_real_db()
    try:
        _ensure_sessions_schema_once(conn)
        row = conn.execute(
            "SELECT username, ip_address FROM user_sessions WHERE sid = ?",
            (sid,),
        ).fetchone()
        if not row:
            flash('Сесію не знайдено.', 'warning')
            return redirect(url_for('system_ops.system_dashboard', tab='sessions'))
        conn.execute(
            """
            UPDATE user_sessions
            SET is_revoked = 1, revoked_by = ?, revoked_at = ?
            WHERE sid = ?
            """,
            (actor, now_str, sid),
        )
        conn.commit()
    finally:
        conn.close()

    log_action(
        actor,
        f"примусово завершив сесію користувача {row['username']}",
        level='WARNING',
        category='security',
        section='Система',
        details=f"IP: {row['ip_address'] or '—'}",
    )
    flash(f"Сесію користувача «{row['username']}» завершено.", 'success')
    return redirect(url_for('system_ops.system_dashboard', tab='sessions'))


@system_ops_bp.route('/admin/system/sessions/revoke-others', methods=['POST'])
@permission_required('manage_system')
def revoke_other_sessions_route():
    current_sid = session.get('session_sid') or ''
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    actor = current_username()
    conn = get_real_db()
    try:
        _ensure_sessions_schema_once(conn)
        cur = conn.execute(
            """
            UPDATE user_sessions
            SET is_revoked = 1, revoked_by = ?, revoked_at = ?
            WHERE is_revoked = 0 AND sid != ?
            """,
            (actor, now_str, current_sid),
        )
        conn.commit()
        count = cur.rowcount or 0
    finally:
        conn.close()

    log_action(
        actor,
        f"примусово завершив усі інші активні сесії ({count})",
        level='WARNING',
        category='security',
        section='Система',
    )
    flash(f"Завершено інших активних сесій: {count}.", 'success')
    return redirect(url_for('system_ops.system_dashboard', tab='sessions'))


# ═══════════════════════════════════════════════════════════════════════════════
# 6. Глобальний швидкий пошук (Ctrl+K)
# ═══════════════════════════════════════════════════════════════════════════════

@system_ops_bp.route('/api/global-search')
@login_required('')
def global_search_api():
    q = (request.args.get('q') or '').strip()
    if len(q) < 2:
        return jsonify({'ok': True, 'query': q, 'students': [], 'groups': [], 'orders': []})

    q_low = q.lower()
    tokens = [t for t in q_low.split() if t]
    is_admin = bool(session.get('is_admin', False))

    can_students = is_admin or has_any_permission(
        'manage_students',
        'view_student_card',
        'manage_military_reports',
        'study_periods',
        'group_export',
        'manage_education_documents',
        'manage_passport_documents',
        'manage_orders',
        'edit_grades',
    )
    can_all_students = is_admin or has_any_permission('manage_military_reports')
    can_groups = is_admin or has_any_permission('manage_groups', 'manage_students', 'manage_courses')
    can_orders = is_admin or has_any_permission('manage_orders')

    conn = get_db()
    try:
        allowed_gids = [] if is_admin else list(get_allowed_group_ids(conn, session.get('user_id')))

        students_res = []
        if can_students and (can_all_students or allowed_gids):
            scope_sql = ""
            scope_params: list = []
            if not can_all_students:
                ph = ','.join('?' for _ in allowed_gids)
                scope_sql = (
                    f" AND (s.group_id IN ({ph}) "
                    f"OR EXISTS (SELECT 1 FROM student_enrollments se_sc "
                    f"WHERE se_sc.student_id = s.id AND se_sc.group_id IN ({ph})))"
                )
                scope_params = allowed_gids + allowed_gids

            # Токенний пошук по ПІБ + збіг за документами/ідентифікаторами
            fio_conds = []
            fio_params: list = []
            for tok in tokens[:5]:
                fio_conds.append(
                    "LOWER_UA(COALESCE(s.last_name_UA,'') || ' ' || COALESCE(s.first_name_UA,'') || ' ' || COALESCE(s.middle_name_UA,'')) LIKE ?"
                )
                fio_params.append(f"%{tok}%")
            fio_sql = "(" + " AND ".join(fio_conds) + ")" if fio_conds else "1=0"

            q_like = f"%{q_low}%"
            q_no_spaces = re.sub(r'\s+', '', q_low)
            q_compact = f"%{q_no_spaces}%"

            sql_students = f"""
                SELECT DISTINCT
                    s.id,
                    s.last_name_UA,
                    s.first_name_UA,
                    s.middle_name_UA,
                    s.tax_id,
                    s.edebo_code,
                    s.phone,
                    COALESCE(s.archived, 0) AS archived,
                    g.name AS group_name,
                    g.start_year AS group_year,
                    (SELECT 1 FROM frozen_students fz WHERE fz.student_id = s.id AND fz.resolved_at IS NULL LIMIT 1) AS is_frozen,
                    (SELECT m.registry_number FROM military m WHERE m.student_id = s.id LIMIT 1) AS registry_number,
                    (SELECT COALESCE(p.series,'') || COALESCE(p.number,'')
                     FROM passport_documents p
                     WHERE p.student_id = s.id
                       AND (LOWER_UA(COALESCE(p.series,'') || COALESCE(p.number,'')) LIKE ?
                            OR LOWER_UA(COALESCE(p.unique_number,'')) LIKE ?)
                     LIMIT 1) AS matched_passport
                FROM students s
                LEFT JOIN groups g ON g.id = s.group_id
                WHERE (
                    {fio_sql}
                    OR LOWER_UA(COALESCE(s.tax_id, '')) LIKE ?
                    OR LOWER_UA(COALESCE(s.edebo_code, '')) LIKE ?
                    OR LOWER_UA(COALESCE(s.phone, '')) LIKE ?
                    OR LOWER_UA(COALESCE(s.phone_backup, '')) LIKE ?
                    OR EXISTS (
                        SELECT 1 FROM student_enrollments se
                        WHERE se.student_id = s.id AND LOWER_UA(COALESCE(se.edebo_code, '')) LIKE ?
                    )
                    OR EXISTS (
                        SELECT 1 FROM passport_documents p
                        WHERE p.student_id = s.id
                          AND (LOWER_UA(COALESCE(p.series,'') || COALESCE(p.number,'')) LIKE ?
                               OR LOWER_UA(COALESCE(p.unique_number,'')) LIKE ?)
                    )
                    OR EXISTS (
                        SELECT 1 FROM military m
                        WHERE m.student_id = s.id
                          AND LOWER_UA(COALESCE(m.registry_number, '')) LIKE ?
                    )
                )
                {scope_sql}
                ORDER BY COALESCE(s.archived, 0) ASC, s.last_name_UA COLLATE UKRAINIAN, s.first_name_UA COLLATE UKRAINIAN
                LIMIT 10
            """
            params = (
                [q_compact, q_compact]
                + fio_params
                + [q_like, q_like, q_like, q_like, q_like, q_compact, q_compact, q_like]
                + scope_params
            )
            for r in conn.execute(sql_students, params).fetchall():
                full_name = f"{r['last_name_UA'] or ''} {r['first_name_UA'] or ''} {r['middle_name_UA'] or ''}".strip()
                hints = []
                if r['group_name']:
                    hints.append(f"{r['group_name']} ({r['group_year']})" if r['group_year'] else r['group_name'])
                if r['tax_id'] and q_low in (r['tax_id'] or '').lower():
                    hints.append(f"РНОКПП: {r['tax_id']}")
                if r['matched_passport']:
                    hints.append(f"Паспорт: {r['matched_passport']}")
                if r['registry_number'] and q_low in (r['registry_number'] or '').lower():
                    hints.append(f"Оберіг: {r['registry_number']}")
                if r['edebo_code'] and q_low in (r['edebo_code'] or '').lower():
                    hints.append(f"ЄДЕБО: {r['edebo_code']}")

                status = 'frozen' if r['is_frozen'] else ('archived' if r['archived'] else 'active')
                students_res.append({
                    'id': r['id'],
                    'title': full_name,
                    'subtitle': ' · '.join(hints) if hints else 'Картка студента',
                    'status': status,
                    'url': url_for('students.student_details', student_id=r['id']),
                })

        groups_res = []
        if can_groups and (is_admin or allowed_gids):
            g_scope_sql = ""
            g_params: list = [f"%{q_low}%", f"%{q_low}%", f"%{q_low}%"]
            if not is_admin:
                ph = ','.join('?' for _ in allowed_gids)
                g_scope_sql = f" AND g.id IN ({ph})"
                g_params.extend(allowed_gids)

            rows_g = conn.execute(
                f"""
                SELECT g.id, g.name, g.start_year, g.study_form, g.degree_level, COALESCE(g.archived, 0) AS archived,
                       (SELECT COUNT(*) FROM students st WHERE st.group_id = g.id AND COALESCE(st.archived, 0) = 0) AS cnt
                FROM groups g
                WHERE (
                    LOWER_UA(COALESCE(g.name, '')) LIKE ?
                    OR LOWER_UA(COALESCE(g.specialty, '')) LIKE ?
                    OR CAST(COALESCE(g.start_year, '') AS TEXT) LIKE ?
                )
                {g_scope_sql}
                ORDER BY COALESCE(g.archived, 0) ASC, g.start_year DESC, g.name COLLATE UKRAINIAN
                LIMIT 6
                """,
                g_params,
            ).fetchall()
            for g in rows_g:
                target_url = (
                    url_for('students.student_list', group_id=g['id'])
                    if (is_admin or has_any_permission('manage_students'))
                    else url_for('admin.manage_groups')
                )
                groups_res.append({
                    'id': g['id'],
                    'title': f"{g['name']} ({g['start_year']})",
                    'subtitle': f"{g['study_form'] or ''} · {g['degree_level'] or ''} · студентів: {g['cnt']}".strip(' ·'),
                    'archived': bool(g['archived']),
                    'url': target_url,
                })

        orders_res = []
        if can_orders:
            try:
                from routes.orders_service import ORDER_TYPE_LABELS
                q_like = f"%{q_low}%"
                rows_o = conn.execute(
                    """
                    SELECT o.id, o.order_type, o.order_number, o.order_date, o.title
                    FROM orders o
                    WHERE o.status = 'active'
                      AND (
                          LOWER_UA(COALESCE(o.order_number, '')) LIKE ?
                          OR LOWER_UA(COALESCE(o.title, '')) LIKE ?
                          OR LOWER_UA(COALESCE(o.notes, '')) LIKE ?
                      )
                    ORDER BY o.id DESC
                    LIMIT 6
                    """,
                    (q_like, q_like, q_like),
                ).fetchall()
                for o in rows_o:
                    t_lbl = ORDER_TYPE_LABELS.get(o['order_type'], o['order_type'])
                    orders_res.append({
                        'id': o['id'],
                        'title': f"Наказ № {o['order_number']} від {o['order_date']}",
                        'subtitle': f"{t_lbl}" + (f" · {o['title']}" if o['title'] else ''),
                        'url': url_for('admin.order_detail', order_id=o['id']),
                    })
            except Exception:
                pass

        return jsonify({
            'ok': True,
            'query': q,
            'students': students_res,
            'groups': groups_res,
            'orders': orders_res,
        })
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 7. Історія змін картки студента
# ═══════════════════════════════════════════════════════════════════════════════

@system_ops_bp.route('/students/<int:student_id>/history')
@login_required('')
def student_card_history_api(student_id: int):
    if not has_any_permission(
        'manage_students',
        'view_student_card',
        'manage_military_reports',
        'study_periods',
        'group_export',
        'manage_education_documents',
        'manage_passport_documents',
        'manage_orders',
        'edit_grades',
    ):
        abort(403)

    include_views = request.args.get('include_views') == '1'
    limit = max(10, min(200, request.args.get('limit', 80, type=int) or 80))

    conn = get_db()
    try:
        if not (has_any_permission('manage_military_reports') or user_can_access_student(conn, student_id)):
            abort(403)

        ensure_audit_logs_table(conn)
        view_filter = "" if include_views else " AND COALESCE(action_type, '') != 'view'"

        rows = conn.execute(
            f"""
            SELECT id, created_at, level, username, user_role, category, action_type,
                   section, action_text, entity_type, entity_id,
                   changes_json, details_text, route_path, ip_address
            FROM audit_logs
            WHERE (
                (entity_type = 'student' AND entity_id = ?)
                OR route_path = ?
                OR route_path LIKE ?
                OR details_json LIKE ?
                OR details_json LIKE ?
            )
            {view_filter}
            ORDER BY created_at DESC, id DESC
            LIMIT ?
            """,
            (
                str(student_id),
                f"/students/{student_id}",
                f"/students/{student_id}/%",
                f'%"student_id": {student_id}%',
                f'%"student_id":{student_id}%',
                limit,
            ),
        ).fetchall()

        entries = []
        for r in rows:
            changes = []
            if r['changes_json']:
                try:
                    parsed = json.loads(r['changes_json'])
                    if isinstance(parsed, list):
                        for ch in parsed:
                            if isinstance(ch, dict):
                                changes.append({
                                    'label': str(ch.get('label') or ch.get('field') or 'Поле'),
                                    'old': str(ch.get('old') if ch.get('old') is not None else '—'),
                                    'new': str(ch.get('new') if ch.get('new') is not None else '—'),
                                })
                except Exception:
                    pass

            entries.append({
                'id': r['id'],
                'created_at': r['created_at'],
                'username': r['username'] or 'система',
                'section': r['section'] or 'Студенти',
                'category': r['category'] or 'system',
                'action_type': r['action_type'] or 'other',
                'action_text': r['action_text'] or '',
                'details_text': r['details_text'] or '',
                'ip_address': r['ip_address'] or '',
                'changes': changes,
            })

        return jsonify({
            'ok': True,
            'student_id': student_id,
            'entries': entries,
        })
    finally:
        conn.close()
