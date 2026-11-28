"""
routes/military_telegram.py
===========================
Telegram-бот для збору військово-облікових документів (ВОД / PDF «Резерв+»),
цільових розсилок студентам за гнучкими фільтрами та модерації поданих документів.

Архітектура:
  - Працює всередині процесу Flask у керованому фоновому потоці (Long Polling через `requests`).
  - Не потребує змін у nginx.conf або додаткових сторонніх бібліотек.
  - Підтримує 3 способи прив'язки студента до Telegram (включно з випадком,
    коли студент не має Telegram і документ подає представник — батьки, родич, староста):
      1) Основний: кнопка «Поділитися контактом» (звірка за телефоном у картці).
      2) Персональне / групове посилання-запрошення (та QR-код) зі стартовим токеном.
      3) Автоприв'язка або режим представника при надсиланні PDF «Резерв+»
         (сувора ідентифікація за ПІБ + датою народження).
  - Виявляє:
      • відсутніх у БД студентів (одразу відповідає у Telegram + фіксує у черзі/сповіщеннях);
      • повторні завантаження за той самий місяць (за «Сформовано:» у PDF і датою подачі,
        перевіряє і чергу бота, і вже підтверджені файли в `attachments`).
"""

import calendar
from datetime import date, datetime, timedelta
import html as html_lib
import json
import os
import re
import secrets
import sqlite3
import tempfile
import threading
import time
import urllib.parse
import uuid

import requests
from flask import (
    Blueprint,
    Response,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

from routes.config import PROJECT_ROOT
from routes.db import get_db, is_test_mode_active
from routes.helpers import current_username
from routes.parse_military_pdf import extract_photo_from_pdf, parse_reserve_plus_pdf
from routes.photo import (
    auto_center_crop_box,
    crop_and_resize,
    load_and_validate_image,
    process_and_save_pending_photo_smart,
    save_final_image,
    save_raw_pending_photo,
)
from routes.security import is_safe_hex_token, resolve_under_dir
from routes.utils import (
    ensure_faculty_schema,
    log_action,
    logger,
    normalize_phone_ua,
    permission_required,
)

military_telegram_bp = Blueprint('military_telegram', __name__)

TELEGRAM_VOD_REL_DIR = os.path.join('uploads', 'telegram_vod')
TELEGRAM_VOD_ABS_DIR = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'telegram_vod')

_SCHEMA_LOCK = threading.Lock()
_SCHEMA_READY = False

_WORKER_LOCK = threading.Lock()
_WORKER_THREAD: threading.Thread | None = None
_WORKER_STOP_EVENT = threading.Event()
_WORKER_STATE = {
    'running': False,
    'started_at': None,
    'last_poll_at': None,
    'last_error': '',
    'poll_cycles': 0,
    'bot_username': '',
    'bot_first_name': '',
}

DEFAULT_WELCOME_TEXT = (
    "Вітаємо! Це офіційний бот військового обліку університету.\n\n"
    "Через цього бота ви можете:\n"
    "• швидко прив'язати свій профіль студента;\n"
    "• надіслати актуальний PDF-документ «Резерв+» (ВОД);\n"
    "• отримувати нагадування та підтвердження перевірки.\n\n"
    "Якщо у студента немає Telegram, подати документ може представник (батьки, родич або староста)."
)

DEFAULT_REMINDER_TEXT = (
    "Шановний(а) {Ім'я}!\n"
    "Просимо надіслати актуальний PDF-файл військово-облікового документа з застосунку «Резерв+» за {Місяць}.\n\n"
    "Для студента: {ПІБ} (група {Група}).\n"
    "Просто прикріпіть PDF-файл у відповідь на це повідомлення."
)

DEFAULT_APPROVAL_TEXT = (
    "✅ Ваш військово-обліковий документ «Резерв+» ({ПІБ}) перевірено та підтверджено відповідальним працівником.\n"
    "Дані в системі оновлено{ДатаСформовано}."
)

DEFAULT_REJECTION_TEXT = (
    "❌ Надісланий документ «Резерв+» ({ПІБ}) відхилено під час перевірки.\n"
    "Причина: {Причина}\n\n"
    "Будь ласка, сформуйте актуальний PDF у «Резерв+» та надішліть його повторно."
)


# ═══════════════════════════════════════════════════════════════════════════════
# 1. Схема БД (ідемпотентна міграція)
# ═══════════════════════════════════════════════════════════════════════════════

def ensure_telegram_vod_schema(conn) -> None:
    """Створює таблиці Telegram-бота ВОД та початкові налаштування / пресети."""
    ensure_faculty_schema(conn)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_bot_settings (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            bot_token TEXT NOT NULL DEFAULT '',
            bot_username TEXT NOT NULL DEFAULT '',
            bot_first_name TEXT NOT NULL DEFAULT '',
            enabled INTEGER NOT NULL DEFAULT 0,
            welcome_text TEXT NOT NULL DEFAULT '',
            reminder_default_text TEXT NOT NULL DEFAULT '',
            approval_text TEXT NOT NULL DEFAULT '',
            rejection_text TEXT NOT NULL DEFAULT '',
            auto_remind_enabled INTEGER NOT NULL DEFAULT 0,
            auto_remind_day INTEGER NOT NULL DEFAULT 5,
            auto_remind_hour INTEGER NOT NULL DEFAULT 10,
            auto_remind_filter_json TEXT NOT NULL DEFAULT '{}',
            last_auto_remind_ym TEXT NOT NULL DEFAULT '',
            last_update_id INTEGER NOT NULL DEFAULT 0,
            last_poll_at TEXT,
            last_error TEXT DEFAULT '',
            updated_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)

    existing_settings = conn.execute(
        "SELECT id FROM telegram_bot_settings WHERE id = 1"
    ).fetchone()
    if not existing_settings:
        env_token = (os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
        conn.execute(
            """
            INSERT INTO telegram_bot_settings (
                id, bot_token, enabled,
                welcome_text, reminder_default_text, approval_text, rejection_text,
                auto_remind_filter_json
            ) VALUES (1, ?, 0, ?, ?, ?, ?, ?)
            """,
            (
                env_token,
                DEFAULT_WELCOME_TEXT,
                DEFAULT_REMINDER_TEXT,
                DEFAULT_APPROVAL_TEXT,
                DEFAULT_REJECTION_TEXT,
                json.dumps({'vod_month_state': 'missing_this_month', 'age_mode': 'any'}, ensure_ascii=False),
            ),
        )

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_student_links (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            chat_id TEXT NOT NULL,
            telegram_user_id TEXT,
            tg_username TEXT,
            tg_first_name TEXT,
            tg_last_name TEXT,
            tg_phone TEXT,
            link_method TEXT NOT NULL DEFAULT 'manual',
            is_representative INTEGER NOT NULL DEFAULT 0,
            representative_note TEXT DEFAULT '',
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT DEFAULT (datetime('now', 'localtime')),
            UNIQUE(student_id, chat_id)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_links_student ON telegram_student_links(student_id, is_active)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_links_chat ON telegram_student_links(chat_id, is_active)")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_invite_tokens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            token TEXT NOT NULL UNIQUE,
            scope TEXT NOT NULL DEFAULT 'student',
            student_id INTEGER,
            group_id INTEGER,
            is_representative_default INTEGER NOT NULL DEFAULT 0,
            note TEXT DEFAULT '',
            created_by TEXT,
            expires_at TEXT,
            max_uses INTEGER,
            uses_count INTEGER NOT NULL DEFAULT 0,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_invites_token ON telegram_invite_tokens(token)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_invites_student ON telegram_invite_tokens(student_id)")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_pending_dialog (
            chat_id TEXT PRIMARY KEY,
            state TEXT NOT NULL,
            payload_json TEXT NOT NULL DEFAULT '{}',
            updated_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_vod_submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER,
            chat_id TEXT NOT NULL,
            telegram_user_id TEXT,
            tg_username TEXT,
            tg_first_name TEXT,
            tg_last_name TEXT,
            tg_phone TEXT,
            file_id TEXT,
            file_unique_id TEXT,
            original_filename TEXT,
            pdf_rel_path TEXT,
            photo_rel_path TEXT,
            photo_raw_rel_path TEXT,
            parsed_full_name TEXT,
            parsed_birth_date TEXT,
            parsed_formed_at TEXT,
            parsed_ym TEXT,
            upload_ym TEXT,
            parsed_json TEXT,
            new_military_json TEXT,
            new_student_json TEXT,
            military_diff_json TEXT,
            student_diff_json TEXT,
            status TEXT NOT NULL DEFAULT 'pending',
            submitted_as_representative INTEGER NOT NULL DEFAULT 0,
            is_monthly_duplicate INTEGER NOT NULL DEFAULT 0,
            duplicate_count INTEGER NOT NULL DEFAULT 0,
            duplicate_info TEXT DEFAULT '',
            is_outdated_pdf INTEGER NOT NULL DEFAULT 0,
            parse_error TEXT DEFAULT '',
            reject_reason TEXT DEFAULT '',
            reviewed_by TEXT,
            reviewed_at TEXT,
            approved_attachment_id INTEGER,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_sub_status ON telegram_vod_submissions(status, id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_sub_student ON telegram_vod_submissions(student_id, parsed_ym, upload_ym)")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_broadcast_presets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            description TEXT DEFAULT '',
            filter_json TEXT NOT NULL DEFAULT '{}',
            message_template TEXT NOT NULL DEFAULT '',
            created_by TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_auto_reminder_rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            preset_id INTEGER,
            filter_json TEXT NOT NULL DEFAULT '{}',
            message_template TEXT NOT NULL DEFAULT '',
            day_from INTEGER NOT NULL DEFAULT 15,
            day_to INTEGER NOT NULL DEFAULT 25,
            times_per_day INTEGER NOT NULL DEFAULT 1,
            start_hour INTEGER NOT NULL DEFAULT 9,
            end_hour INTEGER NOT NULL DEFAULT 20,
            stop_when_submitted INTEGER NOT NULL DEFAULT 1,
            is_enabled INTEGER NOT NULL DEFAULT 1,
            last_sent_date TEXT DEFAULT '',
            last_sent_slot_idx INTEGER NOT NULL DEFAULT -1,
            sent_today_count INTEGER NOT NULL DEFAULT 0,
            last_sent_at TEXT,
            last_run_summary TEXT DEFAULT '',
            total_runs INTEGER NOT NULL DEFAULT 0,
            created_by TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_auto_rules_enabled ON telegram_auto_reminder_rules(is_enabled, id)")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_broadcasts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            message_text TEXT NOT NULL,
            filter_json TEXT NOT NULL DEFAULT '{}',
            total_matched INTEGER NOT NULL DEFAULT 0,
            total_linked_students INTEGER NOT NULL DEFAULT 0,
            total_unlinked_students INTEGER NOT NULL DEFAULT 0,
            sent_chats_count INTEGER NOT NULL DEFAULT 0,
            failed_chats_count INTEGER NOT NULL DEFAULT 0,
            is_auto INTEGER NOT NULL DEFAULT 0,
            created_by TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_broadcast_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            broadcast_id INTEGER NOT NULL,
            student_id INTEGER,
            chat_id TEXT,
            status TEXT NOT NULL,
            error_text TEXT DEFAULT '',
            sent_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_bcast_logs_bid ON telegram_broadcast_logs(broadcast_id)")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_type TEXT NOT NULL,
            severity TEXT NOT NULL DEFAULT 'info',
            title TEXT NOT NULL,
            details TEXT DEFAULT '',
            student_id INTEGER,
            submission_id INTEGER,
            is_read INTEGER NOT NULL DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_events_read ON telegram_events(is_read, id)")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS telegram_group_contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id INTEGER NOT NULL UNIQUE,
            leader_name TEXT NOT NULL DEFAULT '',
            leader_chat_id TEXT NOT NULL DEFAULT '',
            curator_name TEXT NOT NULL DEFAULT '',
            curator_chat_id TEXT NOT NULL DEFAULT '',
            notes TEXT NOT NULL DEFAULT '',
            is_active INTEGER NOT NULL DEFAULT 1,
            updated_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_group_contacts_gid ON telegram_group_contacts(group_id)")

    # Ідемпотентні міграції колонок на існуючих таблицях
    def _ensure_cols(table_name: str, cols_spec: list[tuple[str, str]]) -> None:
        existing_cols = {
            r[1] for r in conn.execute(f"PRAGMA table_info({table_name})").fetchall()
        }
        for col_name, col_ddl in cols_spec:
            if col_name not in existing_cols:
                conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_ddl}")

    _ensure_cols('telegram_bot_settings', [
        ('auto_cleanup_enabled', 'INTEGER NOT NULL DEFAULT 1'),
        ('auto_cleanup_months', 'INTEGER NOT NULL DEFAULT 6'),
        ('auto_cleanup_include_not_found', 'INTEGER NOT NULL DEFAULT 1'),
        ('auto_cleanup_include_approved', 'INTEGER NOT NULL DEFAULT 0'),
        ('last_cleanup_date', "TEXT NOT NULL DEFAULT ''"),
        ('last_cleanup_summary', "TEXT NOT NULL DEFAULT ''"),
    ])

    _ensure_cols('telegram_vod_submissions', [
        ('risk_flags_json', "TEXT NOT NULL DEFAULT '[]'"),
        ('risk_level', "TEXT NOT NULL DEFAULT 'none'"),
        ('risk_count', 'INTEGER NOT NULL DEFAULT 0'),
        ('files_cleaned_at', "TEXT NOT NULL DEFAULT ''"),
    ])

    _ensure_cols('telegram_auto_reminder_rules', [
        ('escalation_enabled', 'INTEGER NOT NULL DEFAULT 0'),
        ('escalation_last_days', 'INTEGER NOT NULL DEFAULT 3'),
        ('escalation_send_to_leaders', 'INTEGER NOT NULL DEFAULT 1'),
        ('escalation_extra_chat_ids', "TEXT NOT NULL DEFAULT ''"),
        ('escalation_last_date', "TEXT NOT NULL DEFAULT ''"),
        ('escalation_last_summary', "TEXT NOT NULL DEFAULT ''"),
    ])

    try:
        from routes.military_reports import ensure_military_license_schema
        ensure_military_license_schema(conn)
    except Exception:
        pass

    # Початкові корисні пресети фільтрів (одноразово)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS _patch_migrations (
            key TEXT PRIMARY KEY,
            applied_at TEXT NOT NULL DEFAULT (datetime('now'))
        )
    """)
    if not conn.execute("SELECT 1 FROM _patch_migrations WHERE key = 'tg_vod_seed_presets_v1'").fetchone():
        starter_presets = [
            (
                '25+ без «Резерв+» за поточний місяць',
                'Студенти віком від 25 років, які ще не подали PDF за поточний місяць',
                json.dumps({'age_mode': '25_plus', 'vod_month_state': 'missing_this_month'}, ensure_ascii=False),
                DEFAULT_REMINDER_TEXT,
            ),
            (
                'Призовники (< 25) без «Резерв+» за місяць',
                'Призовники або студенти молодше 25 років без свіжого документа',
                json.dumps({'age_mode': 'under_25', 'vod_month_state': 'missing_this_month'}, ensure_ascii=False),
                DEFAULT_REMINDER_TEXT,
            ),
            (
                'Відстрочка закінчується найближчі 30 днів',
                'Студенти, у яких дата відстрочки спливає протягом 30 днів або вже минула',
                json.dumps({'deferral_expiring_days': 30}, ensure_ascii=False),
                "Шановний(а) {Ім'я}!\nУ вас спливає або вже закінчився термін дії відстрочки ({Група}). Просимо оновити військово-обліковий документ у «Резерв+» та надіслати новий PDF у цей бот.",
            ),
        ]
        for pname, pdesc, pfjson, pmsg in starter_presets:
            conn.execute(
                """
                INSERT OR IGNORE INTO telegram_broadcast_presets
                    (name, description, filter_json, message_template, created_by)
                VALUES (?, ?, ?, ?, 'system')
                """,
                (pname, pdesc, pfjson, pmsg),
            )
        conn.execute("INSERT OR IGNORE INTO _patch_migrations (key) VALUES ('tg_vod_seed_presets_v1')")


def ensure_telegram_vod_initialized() -> None:
    global _SCHEMA_READY
    if _SCHEMA_READY:
        return
    with _SCHEMA_LOCK:
        if _SCHEMA_READY:
            return
        conn = None
        try:
            conn = get_db()
            conn.execute("PRAGMA busy_timeout = 5000")
            ensure_telegram_vod_schema(conn)
            conn.commit()
            _SCHEMA_READY = True
        except Exception as e:
            logger.warning(f"Не вдалося ініціалізувати схему Telegram ВОД: {e}")
        finally:
            if conn is not None:
                conn.close()


def _record_event(conn, event_type: str, title: str, *, severity: str = 'info',
                  details: str = '', student_id=None, submission_id=None) -> None:
    conn.execute(
        """
        INSERT INTO telegram_events (event_type, severity, title, details, student_id, submission_id)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (event_type, severity, title, details or '', student_id, submission_id),
    )


def _get_bot_settings(conn) -> dict:
    ensure_telegram_vod_schema(conn)
    row = conn.execute("SELECT * FROM telegram_bot_settings WHERE id = 1").fetchone()
    return dict(row) if row else {}


def _mask_token(token: str) -> str:
    t = (token or '').strip()
    if not t:
        return ''
    if len(t) <= 12:
        return '*' * len(t)
    return f"{t[:8]}…{t[-6:]}"


def get_pending_telegram_badge_count(conn) -> int:
    """Кількість елементів, що потребують уваги у навбарі (нові PDF + не знайдені в БД + непрочитані тривоги)."""
    try:
        ensure_telegram_vod_schema(conn)
        cnt = conn.execute(
            "SELECT COUNT(*) FROM telegram_vod_submissions WHERE status IN ('pending', 'not_found')"
        ).fetchone()[0]
        return int(cnt or 0)
    except Exception:
        return 0


# ═══════════════════════════════════════════════════════════════════════════════
# 2. Допоміжні функції: дати, ПІБ, вік, місяці, QR SVG
# ═══════════════════════════════════════════════════════════════════════════════

def _norm_apos(s: str | None) -> str:
    if not s:
        return ''
    return (
        str(s)
        .replace('’', "'")
        .replace('ʼ', "'")
        .replace('‘', "'")
        .replace('`', "'")
    )


def _norm_name_token(s: str | None) -> str:
    clean = ' '.join(_norm_apos(s or '').split())
    return clean.casefold()


def _parse_date_obj(val) -> date | None:
    if not val:
        return None
    s = str(val).strip()
    if not s:
        return None
    m = re.match(r'^(\d{1,2})[./](\d{1,2})[./](\d{4})', s)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def _calc_age(birth_date_str: str | None, ref_date: date | None = None) -> int | None:
    b = _parse_date_obj(birth_date_str)
    if not b:
        return None
    ref = ref_date or date.today()
    return (ref.year - b.year) - ((ref.month, ref.day) < (b.month, b.day))


def _format_ym_ua(ym: str | None) -> str:
    if not ym or not re.match(r'^\d{4}-\d{2}$', ym):
        return ym or '—'
    names = {
        '01': 'Січень', '02': 'Лютий', '03': 'Березень', '04': 'Квітень',
        '05': 'Травень', '06': 'Червень', '07': 'Липень', '08': 'Серпень',
        '09': 'Вересень', '10': 'Жовтень', '11': 'Листопад', '12': 'Грудень',
    }
    y, m = ym.split('-')
    return f"{names.get(m, m)} {y}"


def _parse_ym(val: str | None) -> str | None:
    from routes.military_reports import _parse_ym_and_sort_key
    ym, _ = _parse_ym_and_sort_key(val)
    return ym


def _active_students_sql_where(alias: str = 's', include_frozen: bool = False) -> str:
    frozen_clause = "" if include_frozen else (
        f"AND NOT EXISTS (SELECT 1 FROM frozen_students fz WHERE fz.student_id = {alias}.id AND fz.resolved_at IS NULL)"
    )
    return f"""
        COALESCE({alias}.archived, 0) = 0
        {frozen_clause}
        AND NOT EXISTS (
            SELECT 1 FROM order_items oi
            JOIN orders o ON o.id = oi.order_id
            WHERE oi.entity_type = 'student'
              AND oi.entity_id = {alias}.id
              AND o.order_type = 'expulsion'
              AND COALESCE(o.status, 'active') = 'active'
        )
    """


def _find_active_student_by_fio_and_birth(conn, full_name: str, birth_date: str) -> tuple[dict | None, str]:
    """
    СУВОРИЙ пошук активного студента за ПІБ + датою народження.
    Повертає (student_dict, '') або (None, причина_відмови).
    """
    from routes.students.military import _normalize_date

    raw_name = ' '.join(_norm_apos(full_name or '').split())
    norm_birth = _normalize_date(birth_date or '')

    if not raw_name:
        return None, 'У PDF не розпізнано ПІБ'
    if not norm_birth:
        return None, 'У PDF не розпізнано дату народження (обов’язкова для звірки)'

    parts = raw_name.split(' ')
    if len(parts) < 2:
        return None, f'Неповне ПІБ у PDF: «{raw_name}»'

    pdf_last = _norm_name_token(parts[0])
    pdf_first = _norm_name_token(parts[1])
    pdf_middle = _norm_name_token(' '.join(parts[2:])) if len(parts) > 2 else ''
    pdf_full_norm = _norm_name_token(raw_name)

    rows = conn.execute(
        f"""
        SELECT s.*, g.name AS group_name, g.course AS group_course,
               g.study_form AS group_study_form
        FROM students s
        LEFT JOIN groups g ON g.id = s.group_id
        WHERE {_active_students_sql_where('s', include_frozen=True)}
        """
    ).fetchall()

    # Перший прохід: ПІБ (точний збіг П+І+ПБ) + точна дата народження
    exact_matches = []
    fio_only_matches = []
    for r in rows:
        db_last = _norm_name_token(r['last_name_UA'])
        db_first = _norm_name_token(r['first_name_UA'])
        db_middle = _norm_name_token(r['middle_name_UA'])
        db_full = ' '.join(x for x in (db_last, db_first, db_middle) if x)
        db_birth = _normalize_date(r['birth_date'])

        fio_eq = (
            (db_last == pdf_last and db_first == pdf_first and db_middle == pdf_middle)
            or (db_full == pdf_full_norm)
        )
        # Якщо у базі або у PDF не вказане по батькові, але Прізвище + Ім'я + Дата народження збігаються
        partial_fio_eq = (
            db_last == pdf_last
            and db_first == pdf_first
            and (not db_middle or not pdf_middle)
        )

        if fio_eq or partial_fio_eq:
            fio_only_matches.append(dict(r))
            if db_birth and db_birth == norm_birth:
                exact_matches.append(dict(r))

    if len(exact_matches) == 1:
        return exact_matches[0], ''
    if len(exact_matches) > 1:
        # Якщо один не заморожений, а інші заморожені — віддаємо активного, інакше першого
        for st in exact_matches:
            is_fz = conn.execute(
                "SELECT 1 FROM frozen_students WHERE student_id = ? AND resolved_at IS NULL LIMIT 1",
                (st['id'],),
            ).fetchone()
            if not is_fz:
                return st, ''
        return exact_matches[0], ''

    if fio_only_matches:
        db_dates = ', '.join(sorted({(_normalize_date(m.get('birth_date')) or 'не вказано') for m in fio_only_matches}))
        return None, (
            f"Знайдено збіг за ПІБ ({raw_name}), але дата народження у PDF ({norm_birth}) "
            f"не збігається з базою ({db_dates})"
        )

    return None, f"Активного студента «{raw_name}» ({norm_birth}) немає в базі даних"


def _compute_monthly_duplicate_info(conn, student_id: int, parsed_ym: str | None,
                                    upload_ym: str, exclude_sub_id: int | None = None) -> tuple[bool, int, str]:
    """
    Перевіряє повторні подачі за той самий місяць:
      1) У таблиці telegram_vod_submissions (статус pending або approved) за parsed_ym або upload_ym.
      2) У таблиці attachments (entity_type='military') за датою «Сформовано:» у PDF або uploaded_at.
    Повертає (is_duplicate, nth_attempt_for_month, human_description).
    """
    target_yms = {ym for ym in (parsed_ym, upload_ym) if ym}
    if not target_yms:
        return False, 1, ''

    prior_items = []
    linked_attachment_ids = set()

    sub_rows = conn.execute(
        """
        SELECT id, status, parsed_ym, upload_ym, parsed_formed_at, created_at,
               approved_attachment_id, original_filename
        FROM telegram_vod_submissions
        WHERE student_id = ?
          AND status IN ('pending', 'approved')
          AND (? IS NULL OR id != ?)
        ORDER BY id DESC
        """,
        (student_id, exclude_sub_id, exclude_sub_id),
    ).fetchall()

    for r in sub_rows:
        matched_ym = None
        if r['parsed_ym'] and r['parsed_ym'] in target_yms:
            matched_ym = r['parsed_ym']
        elif r['upload_ym'] and r['upload_ym'] in target_yms:
            matched_ym = r['upload_ym']
        if matched_ym:
            if r['approved_attachment_id']:
                linked_attachment_ids.add(int(r['approved_attachment_id']))
            st_ua = 'очікує перевірки' if r['status'] == 'pending' else 'підтверджено'
            prior_items.append(
                f"Telegram #{r['id']} ({st_ua}, сформовано: {r['parsed_formed_at'] or '—'}, подано: {(r['created_at'] or '')[:16]})"
            )

    mil_row = conn.execute("SELECT id FROM military WHERE student_id = ?", (student_id,)).fetchone()
    if mil_row:
        from routes.military_reports import _load_attachment_formed_map
        att_rows = conn.execute(
            """
            SELECT id, file_path, original_name, uploaded_at, uploaded_by
            FROM attachments
            WHERE entity_type = 'military' AND entity_id = ?
            ORDER BY id DESC
            """,
            (mil_row['id'],),
        ).fetchall()
        att_dicts = [dict(a) for a in att_rows]
        formed_map = _load_attachment_formed_map(conn, att_dicts)
        for a in att_dicts:
            if a['id'] in linked_attachment_ids:
                continue
            f_str = formed_map.get(a['id'], '')
            f_ym = _parse_ym(f_str)
            up_ym = _parse_ym(a.get('uploaded_at'))
            if (f_ym and f_ym in target_yms) or (up_ym and up_ym in target_yms):
                prior_items.append(
                    f"Файл у картці #{a['id']} ({a.get('original_name') or 'PDF'}, сформовано: {f_str or '—'}, завантажено: {(a.get('uploaded_at') or '')[:16]})"
                )

    if not prior_items:
        return False, 1, ''

    nth = len(prior_items) + 1
    summary = f"Повторна подача за місяць ({nth}-й раз): " + '; '.join(prior_items[:4])
    if len(prior_items) > 4:
        summary += f" (+ ще {len(prior_items) - 4})"
    return True, nth, summary


def _build_submission_diffs(conn, student_id: int, parsed: dict) -> tuple[dict, dict, list, list]:
    """Готує словники та diff для military і students, враховуючи ключ `address` у парсері."""
    from routes.students.military import (
        MILITARY_FIELDS,
        STUDENT_FIELDS,
        _build_diff,
        _current_military,
        _current_student,
        _normalize,
    )

    new_military = {}
    for key, _label in MILITARY_FIELDS:
        if key == 'data_verified_on_time':
            new_military[key] = 1 if parsed.get('data_verified_on_time') else 0
        else:
            new_military[key] = _normalize(parsed.get(key))

    new_student = {}
    for key, _label in STUDENT_FIELDS:
        if key == 'address_of_residence':
            new_student[key] = _normalize(parsed.get('address_of_residence') or parsed.get('address'))
        else:
            new_student[key] = _normalize(parsed.get(key))

    cur_mil = _current_military(conn, student_id)
    cur_stu = _current_student(conn, student_id)
    mil_diff = _build_diff(cur_mil, new_military, MILITARY_FIELDS)
    stu_diff = _build_diff(cur_stu, new_student, STUDENT_FIELDS)
    return new_military, new_student, mil_diff, stu_diff


def _norm_risk_text(val) -> str:
    if val is None:
        return ''
    return ' '.join(_norm_apos(str(val)).split())


def _analyze_submission_risks(
    conn,
    student_id: int | None,
    parsed: dict,
    *,
    cur_mil: dict | None = None,
    today_ref: date | None = None,
) -> tuple[list[dict], str]:
    """
    Автоматична звірка критичних розбіжностей та ризиків при розборі PDF «Резерв+»:
      - поява/зміна відмітки про розшук / звернення до поліції;
      - зняття/виключення з обліку або зміна військового статусу;
      - зміна ТЦК порівняно з БД;
      - зникнення активної відстрочки, її прострочення або швидке закінчення (≤30 дн.);
      - прострочення або швидке закінчення терміну дії ВОД (≤14 дн.).
    Повертає (список_ризиків, max_level), де max_level in ('none', 'warning', 'critical').
    """
    if not student_id or not isinstance(parsed, dict):
        return [], 'none'

    if cur_mil is None:
        row = conn.execute("SELECT * FROM military WHERE student_id = ?", (int(student_id),)).fetchone()
        cur_mil = dict(row) if row else {}

    today = today_ref or date.today()
    flags: list[dict] = []

    # 1. Поліція / розшук
    new_police = _norm_risk_text(parsed.get('police_reason'))
    old_police = _norm_risk_text(cur_mil.get('police_reason')) if cur_mil else ''
    new_police_date = _norm_risk_text(parsed.get('police_date'))
    new_sb = _norm_risk_text(parsed.get('status_badge'))
    old_sb = _norm_risk_text(cur_mil.get('status_badge')) if cur_mil else ''
    new_sb_low = new_sb.casefold()

    if new_police or 'розшук' in new_sb_low:
        if not old_police:
            lbl = 'Поліція / розшук (нове!)'
            det = f"У PDF з’явилась відмітка: {new_police or new_sb}"
        elif new_police and old_police and new_police.casefold() != old_police.casefold():
            lbl = 'Зміна запису поліції'
            det = f"Було: «{old_police}» → Стало: «{new_police}»"
        else:
            lbl = 'Відмітка поліції / розшуку'
            det = f"Активна відмітка: {new_police or new_sb}"
        if new_police_date:
            det += f" (від {new_police_date})"
        flags.append({'code': 'police_risk', 'level': 'critical', 'label': lbl, 'detail': det})

    # 2. Зняття / виключення з обліку або зміна статусу
    is_excl_status = any(k in new_sb_low for k in ('знят', 'виключ', 'невійськовозобов'))
    if is_excl_status:
        is_new_excl = old_sb.casefold() != new_sb_low
        flags.append({
            'code': 'status_removed',
            'level': 'critical' if is_new_excl else 'warning',
            'label': new_sb[:34] if len(new_sb) <= 34 else (new_sb[:31] + '…'),
            'detail': (
                f"Статус обліку змінився: «{old_sb or 'не вказано'}» → «{new_sb}»"
                if is_new_excl else
                f"Статус у PDF: «{new_sb}»"
            ),
        })
    elif old_sb and new_sb and old_sb.casefold() != new_sb_low:
        flags.append({
            'code': 'status_changed',
            'level': 'warning',
            'label': 'Зміна статусу ВО',
            'detail': f"Статус у БД «{old_sb}» → у PDF «{new_sb}»",
        })

    # 3. Зміна ТЦК
    old_tck = _norm_risk_text(cur_mil.get('tck_name')) if cur_mil else ''
    new_tck = _norm_risk_text(parsed.get('tck') or parsed.get('tck_name'))
    if old_tck and new_tck and old_tck.casefold() != new_tck.casefold():
        flags.append({
            'code': 'tck_changed',
            'level': 'critical',
            'label': 'Зміна ТЦК',
            'detail': f"ТЦК змінився: «{old_tck}» → «{new_tck}»",
        })

    # 4. Відстрочка (зникла / прострочена / спливає за <= 30 дн.)
    old_def_str = _norm_risk_text(cur_mil.get('deferral_until')) if cur_mil else ''
    new_def_str = _norm_risk_text(parsed.get('deferral_until'))
    old_def_dt = _parse_date_obj(old_def_str)
    new_def_dt = _parse_date_obj(new_def_str)

    if old_def_dt and old_def_dt >= today and not new_def_str:
        flags.append({
            'code': 'deferral_lost',
            'level': 'critical',
            'label': 'Відстрочка зникла!',
            'detail': f"У БД була відстрочка до {old_def_str}, у новому PDF її немає",
        })
    elif new_def_dt and new_def_dt < today:
        flags.append({
            'code': 'deferral_expired',
            'level': 'critical',
            'label': 'Відстрочка закінчилась',
            'detail': f"Термін відстрочки минув {new_def_str}",
        })
    elif new_def_dt and new_def_dt <= (today + timedelta(days=30)):
        days_left = (new_def_dt - today).days
        flags.append({
            'code': 'deferral_expiring',
            'level': 'warning',
            'label': f'Відстрочка ≤ {days_left} дн.',
            'detail': f"Діє до {new_def_str} (залишилось {days_left} дн.)",
        })

    # 5. Термін дії ВОД (valid_until)
    new_val_str = _norm_risk_text(parsed.get('valid_until'))
    new_val_dt = _parse_date_obj(new_val_str)
    if new_val_dt and new_val_dt < today:
        flags.append({
            'code': 'valid_until_expired',
            'level': 'warning',
            'label': 'ВОД прострочений',
            'detail': f"ВОД був дійсний до {new_val_str}",
        })
    elif new_val_dt and new_val_dt <= (today + timedelta(days=14)):
        days_left = (new_val_dt - today).days
        flags.append({
            'code': 'valid_until_expiring',
            'level': 'warning',
            'label': f'ВОД спливає ({days_left} дн.)',
            'detail': f"ВОД дійсний до {new_val_str}",
        })

    if any(f['level'] == 'critical' for f in flags):
        max_level = 'critical'
    elif flags:
        max_level = 'warning'
    else:
        max_level = 'none'
    return flags, max_level


def _hydrate_submission_risks(conn, sub_dict: dict, cur_mil_by_sid: dict | None = None) -> dict:
    """
    Читає збережені risk_flags або дораховує їх «на льоту» для активних/наявних подань у черзі,
    щоб критичні розбіжності підсвічувались навіть для раніше прийнятих у чергу PDF.
    """
    sid = sub_dict.get('student_id')
    flags = []
    level = sub_dict.get('risk_level') or 'none'
    need_compute = (sub_dict.get('status') == 'pending' and sid) or (
        sid and (sub_dict.get('risk_flags_json') in (None, '', '[]')) and level == 'none'
    )
    if need_compute:
        try:
            parsed = json.loads(sub_dict.get('parsed_json') or '{}')
        except Exception:
            parsed = {}
        cur_mil = cur_mil_by_sid.get(int(sid)) if cur_mil_by_sid is not None else None
        flags, level = _analyze_submission_risks(conn, int(sid), parsed, cur_mil=cur_mil)
    else:
        try:
            flags = json.loads(sub_dict.get('risk_flags_json') or '[]')
        except Exception:
            flags = []
    sub_dict['risk_flags'] = flags
    sub_dict['risk_level'] = level
    sub_dict['risk_count'] = len(flags)
    return sub_dict


# ═══════════════════════════════════════════════════════════════════════════════
# 3. Легковажний вбудований SVG QR-генератор (без зовнішніх залежностей)
# ═══════════════════════════════════════════════════════════════════════════════

def _make_qr_svg(data: str, size_px: int = 220) -> str:
    """
    Генерує SVG з QR-кодом. Якщо встановлено `qrcode` — використовує його;
    інакше використовує вбудований стандартний Byte-mode QR-кодер (ECC Level M).
    """
    try:
        import qrcode
        import qrcode.image.svg
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=8,
            border=2,
        )
        qr.add_data(data)
        qr.make(fit=True)
        matrix = qr.get_matrix()
        return _matrix_to_svg(matrix, size_px)
    except ImportError:
        matrix = _encode_qr_byte_m(data.encode('utf-8'))
        return _matrix_to_svg(matrix, size_px)


def _matrix_to_svg(matrix: list[list[bool]], size_px: int = 220) -> str:
    n = len(matrix)
    quiet = 2
    total = n + quiet * 2
    rects = []
    for r, row in enumerate(matrix):
        for c, val in enumerate(row):
            if val:
                rects.append(f'<rect x="{c + quiet}" y="{r + quiet}" width="1" height="1"/>')
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {total} {total}" '
        f'width="{size_px}" height="{size_px}" shape-rendering="crispEdges">'
        f'<rect width="{total}" height="{total}" fill="#ffffff"/>'
        f'<g fill="#000000">{"".join(rects)}</g>'
        f'</svg>'
    )


def _encode_qr_byte_m(payload: bytes) -> list[list[bool]]:
    """
    Чистий Python QR-кодер (Byte mode, ECC Level M, версії 1..10).
    Покриває URL довжиною до 213 байт (повністю достатньо для Telegram deep-link).
    """
    # (version, total_codewords, ec_codewords_per_block, num_blocks_g1, data_cw_g1, num_blocks_g2, data_cw_g2)
    VERSIONS_M = [
        (1, 26, 10, 1, 16, 0, 0),
        (2, 44, 16, 1, 28, 0, 0),
        (3, 70, 26, 1, 44, 0, 0),
        (4, 100, 18, 2, 32, 0, 0),
        (5, 134, 24, 2, 43, 0, 0),
        (6, 172, 16, 4, 27, 0, 0),
        (7, 196, 18, 4, 31, 0, 0),
        (8, 242, 22, 2, 38, 2, 39),
        (9, 292, 22, 3, 36, 2, 37),
        (10, 346, 26, 4, 43, 1, 44),
    ]
    ALIGN_POS = {
        1: [], 2: [6, 18], 3: [6, 22], 4: [6, 26], 5: [6, 30],
        6: [6, 34], 7: [6, 22, 38], 8: [6, 24, 42], 9: [6, 26, 46], 10: [6, 28, 50],
    }

    chosen = None
    for spec in VERSIONS_M:
        ver, _, _, b1, dc1, b2, dc2 = spec
        data_cap = b1 * dc1 + b2 * dc2
        cc_bits = 8 if ver < 10 else 16
        needed_bits = 4 + cc_bits + 8 * len(payload)
        if needed_bits <= data_cap * 8:
            chosen = spec
            break
    if chosen is None:
        raise ValueError("URL занадто довгий для QR")

    ver, total_cw, ec_cw_pb, b1, dc1, b2, dc2 = chosen
    data_cap = b1 * dc1 + b2 * dc2
    cc_bits = 8 if ver < 10 else 16

    bits = []
    def _push_bits(val: int, width: int):
        for i in range(width - 1, -1, -1):
            bits.append((val >> i) & 1)

    _push_bits(0b0100, 4)
    _push_bits(len(payload), cc_bits)
    for b in payload:
        _push_bits(b, 8)
    for _ in range(min(4, data_cap * 8 - len(bits))):
        bits.append(0)
    while len(bits) % 8 != 0:
        bits.append(0)
    data_bytes = []
    for i in range(0, len(bits), 8):
        val = 0
        for j in range(8):
            val = (val << 1) | bits[i + j]
        data_bytes.append(val)
    pad_cycle = (0xEC, 0x11)
    idx = 0
    while len(data_bytes) < data_cap:
        data_bytes.append(pad_cycle[idx & 1])
        idx += 1

    # GF(256) with primitive poly 0x11D
    exp = [0] * 512
    log = [0] * 256
    x = 1
    for i in range(255):
        exp[i] = x
        log[x] = i
        x <<= 1
        if x & 0x100:
            x ^= 0x11D
    for i in range(255, 512):
        exp[i] = exp[i - 255]

    def _gf_mul(a: int, b: int) -> int:
        if a == 0 or b == 0:
            return 0
        return exp[log[a] + log[b]]

    def _rs_gen(degree: int) -> list[int]:
        g = [1]
        for i in range(degree):
            ng = [0] * (len(g) + 1)
            root = exp[i]
            for j, coeff in enumerate(g):
                ng[j] ^= coeff
                ng[j + 1] ^= _gf_mul(coeff, root)
            g = ng
        return g

    def _rs_remainder(data_blk: list[int], gen: list[int]) -> list[int]:
        res = list(data_blk) + [0] * (len(gen) - 1)
        for i in range(len(data_blk)):
            factor = res[i]
            if factor != 0:
                for j in range(len(gen)):
                    res[i + j] ^= _gf_mul(gen[j], factor)
        return res[len(data_blk):]

    gen = _rs_gen(ec_cw_pb)
    blocks_data = []
    blocks_ec = []
    pos = 0
    for count, dlen in ((b1, dc1), (b2, dc2)):
        for _ in range(count):
            blk = data_bytes[pos:pos + dlen]
            pos += dlen
            blocks_data.append(blk)
            blocks_ec.append(_rs_remainder(blk, gen))

    interleaved = []
    max_d = max(len(b) for b in blocks_data)
    for i in range(max_d):
        for blk in blocks_data:
            if i < len(blk):
                interleaved.append(blk[i])
    for i in range(ec_cw_pb):
        for ec in blocks_ec:
            interleaved.append(ec[i])

    all_bits = []
    for b in interleaved:
        for i in range(7, -1, -1):
            all_bits.append((b >> i) & 1)
    rem_bits = 7 if 2 <= ver <= 6 else 0
    all_bits.extend([0] * rem_bits)

    size = ver * 4 + 17
    modules = [[False] * size for _ in range(size)]
    is_func = [[False] * size for _ in range(size)]

    def _set_func(r: int, c: int, val: bool):
        if 0 <= r < size and 0 <= c < size:
            modules[r][c] = val
            is_func[r][c] = True

    def _draw_finder(r0: int, c0: int):
        for dr in range(-1, 8):
            for dc in range(-1, 8):
                rr, cc = r0 + dr, c0 + dc
                if not (0 <= rr < size and 0 <= cc < size):
                    continue
                if 0 <= dr <= 6 and 0 <= dc <= 6:
                    dist = max(abs(dr - 3), abs(dc - 3))
                    _set_func(rr, cc, dist != 2)
                else:
                    _set_func(rr, cc, False)

    _draw_finder(0, 0)
    _draw_finder(0, size - 7)
    _draw_finder(size - 7, 0)

    for i in range(8, size - 8):
        val = (i % 2 == 0)
        _set_func(6, i, val)
        _set_func(i, 6, val)

    apos = ALIGN_POS[ver]
    for r in apos:
        for c in apos:
            if (r == 6 and c == 6) or (r == 6 and c == size - 7) or (r == size - 7 and c == 6):
                continue
            for dr in range(-2, 3):
                for dc in range(-2, 3):
                    _set_func(r + dr, c + dc, max(abs(dr), abs(dc)) != 1)

    _set_func(size - 8, 8, True)

    # Reserve format & version strips
    for i in range(9):
        if i != 6:
            is_func[8][i] = True
            is_func[i][8] = True
    for i in range(8):
        is_func[8][size - 1 - i] = True
        is_func[size - 1 - i][8] = True
    if ver >= 7:
        for i in range(6):
            for j in range(3):
                is_func[size - 11 + j][i] = True
                is_func[i][size - 11 + j] = True

    # Place data bits
    bit_idx = 0
    upward = True
    col = size - 1
    while col > 0:
        if col == 6:
            col -= 1
        rows_iter = range(size - 1, -1, -1) if upward else range(size)
        for r in rows_iter:
            for dc in (0, -1):
                cc = col + dc
                if not is_func[r][cc]:
                    modules[r][cc] = bool(all_bits[bit_idx]) if bit_idx < len(all_bits) else False
                    bit_idx += 1
        col -= 2
        upward = not upward

    # Format info (ECC M = 0b00, BCH(15,5) + mask 0x5412, bit 14=MSB .. bit 0=LSB)
    def _write_format(mask_id: int):
        fmt_data = (0b00 << 3) | mask_id
        rem = fmt_data
        for _ in range(10):
            rem = (rem << 1) ^ ((rem >> 9) * 0x537)
        bits15 = ((fmt_data << 10) | rem) ^ 0x5412

        def _fbit(idx: int) -> bool:
            return bool((bits15 >> idx) & 1)

        # Копія 1: навколо верхнього лівого finder
        for i in range(6):
            modules[8][i] = _fbit(14 - i)
        modules[8][7] = _fbit(8)
        modules[8][8] = _fbit(7)
        modules[7][8] = _fbit(6)
        for i in range(6):
            modules[5 - i][8] = _fbit(5 - i)

        # Копія 2: нижній лівий вертикально (біти 14..8) та верхній правий горизонтально (біти 7..0)
        for i in range(7):
            modules[size - 1 - i][8] = _fbit(14 - i)
        modules[size - 8][8] = True  # Dark module завжди темний
        for i in range(8):
            modules[8][size - 8 + i] = _fbit(7 - i)

    if ver >= 7:
        rem = ver
        for _ in range(12):
            rem = (rem << 1) ^ ((rem >> 11) * 0x1F25)
        vbits = (ver << 12) | rem
        for i in range(18):
            bit = bool((vbits >> i) & 1)
            a, b = size - 11 + (i % 3), i // 3
            modules[a][b] = bit
            modules[b][a] = bit

    MASK_FUNCS = [
        lambda r, c: (r + c) % 2 == 0,
        lambda r, c: r % 2 == 0,
        lambda r, c: c % 3 == 0,
        lambda r, c: (r + c) % 3 == 0,
        lambda r, c: (r // 2 + c // 3) % 2 == 0,
        lambda r, c: (r * c) % 2 + (r * c) % 3 == 0,
        lambda r, c: ((r * c) % 2 + (r * c) % 3) % 2 == 0,
        lambda r, c: ((r + c) % 2 + (r * c) % 3) % 2 == 0,
    ]

    def _apply_mask(mask_id: int):
        fn = MASK_FUNCS[mask_id]
        for r in range(size):
            for c in range(size):
                if not is_func[r][c] and fn(r, c):
                    modules[r][c] = not modules[r][c]

    def _penalty() -> int:
        pen = 0
        for r in range(size):
            run_col = modules[r][0]
            run_len = 1
            for c in range(1, size):
                if modules[r][c] == run_col:
                    run_len += 1
                else:
                    if run_len >= 5:
                        pen += 3 + (run_len - 5)
                    run_col = modules[r][c]
                    run_len = 1
            if run_len >= 5:
                pen += 3 + (run_len - 5)
        for c in range(size):
            run_col = modules[0][c]
            run_len = 1
            for r in range(1, size):
                if modules[r][c] == run_col:
                    run_len += 1
                else:
                    if run_len >= 5:
                        pen += 3 + (run_len - 5)
                    run_col = modules[r][c]
                    run_len = 1
            if run_len >= 5:
                pen += 3 + (run_len - 5)
        for r in range(size - 1):
            for c in range(size - 1):
                v = modules[r][c]
                if v == modules[r][c + 1] == modules[r + 1][c] == modules[r + 1][c + 1]:
                    pen += 3
        return pen

    best_mask = 0
    best_pen = 10**9
    for mid in range(8):
        _apply_mask(mid)
        _write_format(mid)
        p = _penalty()
        if p < best_pen:
            best_pen = p
            best_mask = mid
        _apply_mask(mid)

    _apply_mask(best_mask)
    _write_format(best_mask)
    return modules


# ═══════════════════════════════════════════════════════════════════════════════
# 4. Багатокритеріальна фільтрація студентів (без обмеження кількості фільтрів)
# ═══════════════════════════════════════════════════════════════════════════════

def _parse_filter_dict(raw) -> dict:
    """Нормалізує параметри фільтра (з query/form/JSON) до єдиної структури."""
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except Exception:
            raw = {}
    if not isinstance(raw, dict):
        raw = {}

    def _int_list(key):
        val = raw.get(key) or []
        if isinstance(val, (int, str)):
            val = [x for x in str(val).split(',') if x.strip()]
        out = []
        for x in val:
            try:
                out.append(int(x))
            except (TypeError, ValueError):
                pass
        return out

    def _str_list(key):
        val = raw.get(key) or []
        if isinstance(val, str):
            val = [x for x in val.split(',') if x.strip()]
        return [str(x).strip() for x in val if str(x).strip()]

    def _opt_int(key):
        v = raw.get(key)
        if v is None or str(v).strip() == '':
            return None
        try:
            return int(v)
        except (TypeError, ValueError):
            return None

    today_ym = date.today().strftime('%Y-%m')
    vod_month = str(raw.get('vod_month') or today_ym).strip()
    if not re.match(r'^\d{4}-\d{2}$', vod_month):
        vod_month = today_ym

    age_mode = str(raw.get('age_mode') or 'any').strip()
    if age_mode not in ('any', '25_plus', 'under_25', 'range'):
        age_mode = 'any'

    return {
        'search': str(raw.get('search') or '').strip(),
        'faculty_ids': _int_list('faculty_ids'),
        'group_ids': _int_list('group_ids'),
        'courses': _int_list('courses'),
        'study_forms': _str_list('study_forms'),
        'degree_levels': _str_list('degree_levels'),
        'license_ids': _int_list('license_ids'),
        'only_base_license': bool(raw.get('only_base_license') in (True, 1, '1', 'true', 'on')),
        'genders': _str_list('genders'),
        'age_mode': age_mode,
        'age_min': _opt_int('age_min'),
        'age_max': _opt_int('age_max'),
        'age_as_of': str(raw.get('age_as_of') or '').strip(),
        'military_statuses': _str_list('military_statuses'),
        'military_categories': _str_list('military_categories'),
        'exclude_women_no_mil': bool(raw.get('exclude_women_no_mil') in (True, 1, '1', 'true', 'on')),
        'vod_month': vod_month,
        'vod_month_state': str(raw.get('vod_month_state') or 'any').strip(),
        'formed_older_than_days': _opt_int('formed_older_than_days'),
        'deferral_expiring_days': _opt_int('deferral_expiring_days'),
        'valid_until_expiring_days': _opt_int('valid_until_expiring_days'),
        'has_police_reason': str(raw.get('has_police_reason') or 'any').strip(),
        'note_18': _str_list('note_18'),
        'exclude_note_18': bool(raw.get('exclude_note_18') in (True, 1, '1', 'true', 'on')),
        'tg_link_state': str(raw.get('tg_link_state') or 'any').strip(),
        'include_frozen': bool(raw.get('include_frozen') in (True, 1, '1', 'true', 'on')),
        'include_student_ids': _int_list('include_student_ids'),
        'exclude_student_ids': _int_list('exclude_student_ids'),
    }


def query_filtered_students(conn, raw_filters) -> dict:
    """
    Виконує гнучку багатокритеріальну фільтрацію студентів для Telegram-бота.
    Будь-які фільтри можна комбінувати одночасно.
    """
    from routes.military_reports import (
        CATEGORY_LABELS,
        NOTE_18_EXCLUDES_FROM_DODATOK_5,
        _detect_category,
        _get_military_visible_licenses,
        _has_base_license,
        _is_woman_without_military,
        _load_attachment_formed_map,
        _load_enrollment_licenses,
        _parse_ym_and_sort_key,
    )

    f = _parse_filter_dict(raw_filters)
    today = date.today()
    ref_age_date = _parse_date_obj(f['age_as_of']) or today
    vod_month = f['vod_month']

    allowed_license_ids, _ = _get_military_visible_licenses(conn)
    enrollment_licenses = _load_enrollment_licenses(conn)

    # Мапа факультетів (група -> факультет, або спеціальність -> факультет)
    fac_rows = conn.execute("SELECT id, name_ua, short_name FROM faculties").fetchall()
    fac_by_id = {int(r['id']): dict(r) for r in fac_rows}

    spec_fac_rows = conn.execute(
        "SELECT UPPER(TRIM(code)) AS code, faculty_id FROM specialties WHERE faculty_id IS NOT NULL"
    ).fetchall()
    spec_to_fac = {r['code']: int(r['faculty_id']) for r in spec_fac_rows}

    # Збираємо активні посилання Telegram на студентів
    link_rows = conn.execute(
        """
        SELECT *
        FROM telegram_student_links
        WHERE is_active = 1
        ORDER BY is_representative ASC, id DESC
        """
    ).fetchall()
    links_by_student: dict[int, list[dict]] = {}
    for lr in link_rows:
        links_by_student.setdefault(int(lr['student_id']), []).append(dict(lr))

    # Збираємо всі PDF / файли з attachments (entity_type='military') + кеш Сформовано
    att_rows_raw = conn.execute(
        """
        SELECT id, entity_id, file_path, original_name, uploaded_at
        FROM attachments
        WHERE entity_type = 'military'
        ORDER BY id DESC
        """
    ).fetchall()
    att_dicts = [dict(a) for a in att_rows_raw]
    pdf_formed_map = _load_attachment_formed_map(conn, att_dicts)
    atts_by_mil: dict[int, list[dict]] = {}
    for a in att_dicts:
        atts_by_mil.setdefault(int(a['entity_id']), []).append(a)

    # Збираємо подачі через Telegram за обраний місяць (pending/approved)
    tg_sub_rows = conn.execute(
        """
        SELECT student_id, status, parsed_ym, upload_ym, parsed_formed_at, created_at
        FROM telegram_vod_submissions
        WHERE student_id IS NOT NULL
          AND status IN ('pending', 'approved')
        ORDER BY id DESC
        """
    ).fetchall()
    subs_by_student: dict[int, list[dict]] = {}
    for sr in tg_sub_rows:
        subs_by_student.setdefault(int(sr['student_id']), []).append(dict(sr))

    students = conn.execute(
        f"""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.gender, s.tax_id, s.phone, s.phone_backup, s.email,
               s.license_id, s.group_id,
               g.name AS group_name, g.course AS group_course,
               g.study_form AS group_study_form, g.degree_level AS group_degree_level,
               g.faculty_id AS group_faculty_id, g.specialty_code AS group_specialty_code,
               g.faculty AS group_faculty_name,
               m.id AS military_id, m.status_badge, m.military_rank, m.category_override,
               m.tck_name, m.vos_code, m.note_18, m.formed_at, m.valid_until,
               m.deferral_type, m.deferral_until, m.police_reason,
               fz.id AS frozen_id
        FROM students s
        LEFT JOIN groups g ON g.id = s.group_id
        LEFT JOIN military m ON m.student_id = s.id
        LEFT JOIN frozen_students fz ON fz.student_id = s.id AND fz.resolved_at IS NULL
        WHERE {_active_students_sql_where('s', include_frozen=f['include_frozen'])}
        ORDER BY s.last_name_UA COLLATE UKRAINIAN,
                 s.first_name_UA COLLATE UKRAINIAN,
                 s.middle_name_UA COLLATE UKRAINIAN
        """
    ).fetchall()

    include_set = set(f['include_student_ids'])
    exclude_set = set(f['exclude_student_ids'])

    matched = []
    for s in students:
        sid = int(s['id'])
        if sid in exclude_set:
            continue

        # Якщо увімкнено явний whitelist (без інших критеріїв? Ні — include_student_ids ДОДАТКОВО пропускає або обмежує?
        # Зручніше: якщо користувач обрав конкретних студентів у фільтрі, або застосовувати як фільтр, якщо не порожній,
        # а якщо він обрав і фільтри — перетин; але в UI зробимо інтуїтивно: "Тільки обрані студенти" коли заповнено).
        if include_set and sid not in include_set:
            continue

        # Факультет групи (через faculty_id групи або спеціальності)
        eff_fac_id = s['group_faculty_id']
        if eff_fac_id is None and s['group_specialty_code']:
            eff_fac_id = spec_to_fac.get(str(s['group_specialty_code']).strip().upper())
        eff_fac = fac_by_id.get(int(eff_fac_id)) if eff_fac_id is not None else None
        eff_fac_name = eff_fac['name_ua'] if eff_fac else (s['group_faculty_name'] or '')

        # Ліцензія
        if f['only_base_license'] and not _has_base_license(sid, enrollment_licenses, allowed_license_ids, s['license_id']):
            continue
        eff_lids = enrollment_licenses.get(sid) or ({int(s['license_id'])} if s['license_id'] is not None else set())
        if f['license_ids'] and not (eff_lids & set(f['license_ids'])):
            continue

        # Група / факультет / курс / форма / ступінь / стать
        if f['faculty_ids'] and (eff_fac_id is None or int(eff_fac_id) not in f['faculty_ids']):
            continue
        if f['group_ids'] and (s['group_id'] is None or int(s['group_id']) not in f['group_ids']):
            continue
        if f['courses'] and (s['group_course'] is None or int(s['group_course']) not in f['courses']):
            continue
        if f['study_forms'] and (s['group_study_form'] or '') not in f['study_forms']:
            continue
        if f['degree_levels'] and (s['group_degree_level'] or '') not in f['degree_levels']:
            continue
        if f['genders'] and (s['gender'] or '') not in f['genders']:
            continue

        # Вік
        age = _calc_age(s['birth_date'], ref_age_date)
        if f['age_mode'] == '25_plus':
            if age is None or age < 25:
                continue
        elif f['age_mode'] == 'under_25':
            if age is None or age >= 25:
                continue
        elif f['age_mode'] == 'range':
            if age is None:
                continue
            if f['age_min'] is not None and age < f['age_min']:
                continue
            if f['age_max'] is not None and age > f['age_max']:
                continue

        # Жінки без ВО (опція)
        if f['exclude_women_no_mil'] and _is_woman_without_military(
            s['gender'], s['military_id'], s['status_badge'],
            s['military_rank'], s['vos_code'], s['tck_name'],
        ):
            continue

        # Військовий статус та категорія
        status_badge = (s['status_badge'] or '').strip()
        if f['military_statuses']:
            want_empty = '__empty__' in f['military_statuses']
            if not status_badge:
                if not want_empty:
                    continue
            elif status_badge not in f['military_statuses']:
                continue

        cat = _detect_category(
            s['military_rank'], status_badge, s['gender'],
            s['category_override'], s['note_18'],
        ) if s['military_id'] else None

        if f['military_categories']:
            cat_key = cat or 'other'
            if cat_key not in f['military_categories']:
                continue

        # Графа 18
        if f['exclude_note_18'] and s['note_18'] in NOTE_18_EXCLUDES_FROM_DODATOK_5:
            continue
        if f['note_18']:
            if '__empty__' in f['note_18'] and not s['note_18']:
                pass
            elif (s['note_18'] or '') not in f['note_18']:
                continue

        # Поліція
        has_police = bool((s['police_reason'] or '').strip())
        if f['has_police_reason'] == 'yes' and not has_police:
            continue
        if f['has_police_reason'] == 'no' and has_police:
            continue

        # Відстрочка / дійсний до
        if f['deferral_expiring_days'] is not None:
            d_def = _parse_date_obj(s['deferral_until'])
            if not d_def or d_def > (today + timedelta(days=f['deferral_expiring_days'])):
                continue

        if f['valid_until_expiring_days'] is not None:
            d_val = _parse_date_obj(s['valid_until'])
            if not d_val or d_val > (today + timedelta(days=f['valid_until_expiring_days'])):
                continue

        # Файли ВОД та подачі у вибраному місяці
        mil_atts = atts_by_mil.get(int(s['military_id']), []) if s['military_id'] else []
        att_months = set()
        latest_formed_dt: date | None = None
        latest_formed_str = (s['formed_at'] or '').strip()
        if latest_formed_str:
            d_f = _parse_date_obj(latest_formed_str)
            if d_f:
                latest_formed_dt = d_f

        for a in mil_atts:
            f_str = pdf_formed_map.get(a['id'], '')
            ym, _ = _parse_ym_and_sort_key(f_str)
            if ym:
                att_months.add(ym)
            d_a = _parse_date_obj(f_str)
            if d_a and (latest_formed_dt is None or d_a > latest_formed_dt):
                latest_formed_dt = d_a
                latest_formed_str = f_str

        tg_subs = subs_by_student.get(sid, [])
        sub_months_pending_or_ok = set()
        sub_months_pending = set()
        sub_months_approved = set()
        for sr in tg_subs:
            yms_sr = [y for y in (sr.get('parsed_ym'), sr.get('upload_ym')) if y]
            for y in yms_sr:
                sub_months_pending_or_ok.add(y)
                if sr.get('status') == 'pending':
                    sub_months_pending.add(y)
                elif sr.get('status') == 'approved':
                    sub_months_approved.add(y)

        has_att_in_month = vod_month in att_months
        has_approved_in_month = has_att_in_month or (vod_month in sub_months_approved)
        has_pending_in_month = (not has_approved_in_month) and (vod_month in sub_months_pending)
        has_any_in_month = has_approved_in_month or (vod_month in sub_months_pending_or_ok)

        vms = f['vod_month_state']
        if vms == 'missing_this_month' and has_any_in_month:
            continue
        elif vms == 'never_uploaded' and (mil_atts or tg_subs):
            continue
        elif vms == 'outdated_prev_months' and (not mil_atts or has_any_in_month):
            continue
        elif vms == 'already_submitted_or_pending' and not has_any_in_month:
            continue
        elif vms == 'has_approved_this_month' and not has_att_in_month:
            continue

        if f['formed_older_than_days'] is not None:
            if latest_formed_dt is not None:
                days_old = (today - latest_formed_dt).days
                if days_old <= f['formed_older_than_days']:
                    continue

        # Пошук по тексту
        full_name = ' '.join(
            x for x in (s['last_name_UA'], s['first_name_UA'], s['middle_name_UA'] or '') if x
        ).strip()
        if f['search']:
            hay = ' '.join([
                full_name,
                s['phone'] or '',
                s['phone_backup'] or '',
                s['email'] or '',
                s['tax_id'] or '',
                s['group_name'] or '',
                eff_fac_name,
            ]).lower()
            if not all(w in hay for w in f['search'].lower().split()):
                continue

        # Telegram прив'язка
        st_links = links_by_student.get(sid, [])
        is_linked = len(st_links) > 0
        has_rep = any(bool(l.get('is_representative')) for l in st_links)
        if f['tg_link_state'] == 'linked' and not is_linked:
            continue
        elif f['tg_link_state'] == 'unlinked' and is_linked:
            continue
        elif f['tg_link_state'] == 'representative' and not has_rep:
            continue

        cat_label = CATEGORY_LABELS[cat][0] if cat in CATEGORY_LABELS else 'Не визначена'
        matched.append({
            'id': sid,
            'full_name': full_name,
            'first_name': (s['first_name_UA'] or '').strip(),
            'birth_date': s['birth_date'] or '',
            'age': age,
            'gender': s['gender'] or '',
            'phone': s['phone'] or '',
            'phone_backup': s['phone_backup'] or '',
            'email': s['email'] or '',
            'group_id': s['group_id'],
            'group_name': s['group_name'] or '',
            'course': s['group_course'],
            'study_form': s['group_study_form'] or '',
            'degree_level': s['group_degree_level'] or '',
            'faculty_id': eff_fac_id,
            'faculty_name': eff_fac_name,
            'status_badge': status_badge,
            'category': cat,
            'category_label': cat_label,
            'formed_at': latest_formed_str,
            'deferral_until': s['deferral_until'] or '',
            'valid_until': s['valid_until'] or '',
            'has_att_in_month': has_att_in_month,
            'has_approved_in_month': has_approved_in_month,
            'has_pending_in_month': has_pending_in_month,
            'has_any_in_month': has_any_in_month,
            'links': st_links,
            'is_linked': is_linked,
            'has_representative': has_rep,
        })

    linked_students = [r for r in matched if r['is_linked']]
    unlinked_students = [r for r in matched if not r['is_linked']]
    unique_chats = set()
    for r in linked_students:
        for l in r['links']:
            unique_chats.add(str(l['chat_id']))

    return {
        'filters': f,
        'students': matched,
        'total_matched': len(matched),
        'linked_count': len(linked_students),
        'unlinked_count': len(unlinked_students),
        'unique_chats_count': len(unique_chats),
        'unlinked_students': unlinked_students,
    }


def _render_broadcast_message(template: str, student: dict, vod_month: str) -> str:
    """Підставляє динамічні плейсхолдери в текст розсилки."""
    text = template or DEFAULT_REMINDER_TEXT
    replacements = {
        '{ПІБ}': student.get('full_name') or '',
        "{Ім'я}": student.get('first_name') or student.get('full_name') or '',
        '{Група}': student.get('group_name') or '—',
        '{Курс}': str(student.get('course') or '—'),
        '{Факультет}': student.get('faculty_name') or '—',
        '{Місяць}': _format_ym_ua(vod_month),
        '{Вік}': str(student.get('age')) if student.get('age') is not None else '—',
        '{СтатусВО}': student.get('status_badge') or 'Не вказано',
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text


def _format_group_escalation_text(grp: dict, vod_month: str, *, rule_name: str = '') -> str:
    """Формує зрозумілий Telegram-звіт по групі (хто здав / хто НЕ здав «Резерв+»)."""
    ym_label = _format_ym_ua(vod_month)
    hdr = f"📋 Зведення ВОД «Резерв+» за {ym_label}"
    if rule_name:
        hdr += f"\n📌 Правило: {rule_name}"
    lines = [
        hdr,
        f"🎓 Група: {grp['group_name']} ({grp.get('faculty_name') or '—'})",
        f"✅ Підтверджено: {grp['approved_count']} з {grp['total']}",
    ]
    if grp['pending_count']:
        lines.append(f"⏳ На перевірці: {grp['pending_count']}")
    lines.append(f"❌ Не здали: {grp['missing_count']} ({grp['pct']}% виконання)")

    if grp['missing_students']:
        lines.append("\nБоржники за поточний місяць:")
        for idx, st in enumerate(grp['missing_students'], 1):
            tg_mark = '' if st.get('is_linked') else ' [без TG]'
            lines.append(f"{idx}. {st['full_name']}{tg_mark}")
        lines.append("\nПросимо нагадати зазначеним студентам надіслати PDF «Резерв+» у бота.")
    else:
        lines.append("\n🎉 Усі студенти групи подали актуальні документи!")

    return "\n".join(lines)


def _format_consolidated_escalation_text(summary: dict, vod_month: str, *, rule_name: str = '') -> str:
    """Формує зведений звіт по всіх групах з боржниками для куратора/відповідального."""
    ym_label = _format_ym_ua(vod_month)
    totals = summary['totals']
    groups_missing = [g for g in summary['groups'] if g['missing_count'] > 0]
    lines = [
        f"🚨 Ескалація ВОД «Резерв+» за {ym_label}" + (f" ({rule_name})" if rule_name else ""),
        f"Всього у вибірці: {totals['total_students']} | Здали: {totals['submitted_count']} | Не здали: {totals['missing_count']}",
        f"Груп з боржниками: {len(groups_missing)}",
        "",
    ]
    for g in groups_missing:
        names = ", ".join(st['full_name'] + ('' if st.get('is_linked') else ' [без TG]') for st in g['missing_students'])
        lines.append(f"• {g['group_name']} (не здали {g['missing_count']}/{g['total']}): {names}")
    return "\n".join(lines)


def build_group_vod_compliance_summary(
    conn,
    raw_filters=None,
    *,
    vod_month: str | None = None,
    only_with_missing: bool = False,
) -> dict:
    """
    Формує статистику здачі ВОД «Резерв+» по групах:
    хто здав (підтверджено / на перевірці), хто НЕ здав, і приєднує контакти старост/кураторів.
    """
    ensure_telegram_vod_schema(conn)
    raw_in = raw_filters if isinstance(raw_filters, dict) else {}
    flt = _parse_filter_dict(raw_in)
    if vod_month and re.match(r'^\d{4}-\d{2}$', str(vod_month).strip()):
        flt['vod_month'] = str(vod_month).strip()
    # Для повної картини по групах беремо всіх студентів обраної вибірки (без фільтра «лише боржники»)
    flt['vod_month_state'] = 'any'
    # Дефолтно показуємо лише налаштовані ліцензії ВО (Львівську тощо), якщо явно не задано інше
    if 'only_base_license' not in raw_in and not raw_in.get('show_all_licenses') and not flt.get('license_ids'):
        flt['only_base_license'] = True
    # Дефолтно не обмежуємо прив'язкою TG, щоб бачити реальну картину здачі в групі
    if 'tg_link_state' not in raw_in:
        flt['tg_link_state'] = 'any'
    if 'exclude_women_no_mil' not in raw_in:
        flt['exclude_women_no_mil'] = True
    if 'exclude_note_18' not in raw_in:
        flt['exclude_note_18'] = True

    res = query_filtered_students(conn, flt)
    eff_ym = res['filters']['vod_month']

    contacts_rows = conn.execute("SELECT * FROM telegram_group_contacts").fetchall()
    contacts_by_gid = {int(r['group_id']): dict(r) for r in contacts_rows}

    groups_map: dict[int, dict] = {}
    for st in res['students']:
        gid = int(st['group_id']) if st.get('group_id') is not None else 0
        if gid not in groups_map:
            groups_map[gid] = {
                'group_id': gid,
                'group_name': st.get('group_name') or 'Без групи',
                'course': st.get('course'),
                'study_form': st.get('study_form') or '',
                'faculty_name': st.get('faculty_name') or '',
                'approved_students': [],
                'pending_students': [],
                'missing_students': [],
                'contact': contacts_by_gid.get(gid) or {},
            }
        grp = groups_map[gid]
        if st.get('has_approved_in_month'):
            grp['approved_students'].append(st)
        elif st.get('has_pending_in_month'):
            grp['pending_students'].append(st)
        else:
            grp['missing_students'].append(st)

    groups_list = []
    tot_all = 0
    tot_app = 0
    tot_pen = 0
    tot_mis = 0
    for grp in groups_map.values():
        app_c = len(grp['approved_students'])
        pen_c = len(grp['pending_students'])
        mis_c = len(grp['missing_students'])
        total_c = app_c + pen_c + mis_c
        sub_c = app_c + pen_c
        pct = int(round((sub_c / total_c) * 100)) if total_c > 0 else 100
        grp['total'] = total_c
        grp['approved_count'] = app_c
        grp['pending_count'] = pen_c
        grp['submitted_count'] = sub_c
        grp['missing_count'] = mis_c
        grp['pct'] = pct
        grp['leader_text'] = _format_group_escalation_text(grp, eff_ym)

        tot_all += total_c
        tot_app += app_c
        tot_pen += pen_c
        tot_mis += mis_c

        if only_with_missing and mis_c == 0:
            continue
        groups_list.append(grp)

    groups_list.sort(key=lambda g: (-g['missing_count'], g['group_name'].casefold()))
    return {
        'vod_month': eff_ym,
        'vod_month_label': _format_ym_ua(eff_ym),
        'groups': groups_list,
        'totals': {
            'total_students': tot_all,
            'approved_count': tot_app,
            'pending_count': tot_pen,
            'submitted_count': tot_app + tot_pen,
            'missing_count': tot_mis,
            'groups_total': len(groups_map),
            'groups_with_missing': sum(1 for g in groups_map.values() if g['missing_count'] > 0),
        },
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 5. Telegram Bot API клієнт і фоновий Long Polling воркер
# ═══════════════════════════════════════════════════════════════════════════════

def _tg_call(token: str, method: str, payload: dict | None = None, timeout: float = 15.0) -> tuple[bool, dict | list | None, str]:
    if is_test_mode_active():
        return False, None, 'Увімкнено тестовий режим: виклики Telegram API заблоковано, щоб не турбувати реальних студентів'
    url = f"https://api.telegram.org/bot{token}/{method}"
    try:
        resp = requests.post(url, json=payload or {}, timeout=timeout)
        data = resp.json()
        if not data.get('ok'):
            desc = data.get('description') or f"HTTP {resp.status_code}"
            return False, None, desc
        return True, data.get('result'), ''
    except requests.RequestException as e:
        return False, None, str(e)
    except ValueError as e:
        return False, None, f"Некоректна відповідь Telegram: {e}"


def _tg_send_message(token: str, chat_id: str | int, text: str,
                     reply_markup: dict | None = None, parse_mode: str | None = None) -> tuple[bool, str]:
    payload: dict = {
        'chat_id': chat_id,
        'text': text[:4096],
        'disable_web_page_preview': True,
    }
    if reply_markup is not None:
        payload['reply_markup'] = reply_markup
    if parse_mode:
        payload['parse_mode'] = parse_mode
    ok, _, err = _tg_call(token, 'sendMessage', payload, timeout=15.0)
    return ok, err


def _tg_answer_callback(token: str, callback_id: str, text: str = '') -> None:
    _tg_call(token, 'answerCallbackQuery', {'callback_query_id': callback_id, 'text': text[:200]}, timeout=8.0)


def _tg_download_file(token: str, file_id: str) -> tuple[bool, bytes | None, str, str]:
    ok, res, err = _tg_call(token, 'getFile', {'file_id': file_id}, timeout=15.0)
    if not ok or not isinstance(res, dict) or not res.get('file_path'):
        return False, None, '', err or 'Не вдалося отримати шлях файлу в Telegram'
    file_path = res['file_path']
    dl_url = f"https://api.telegram.org/file/bot{token}/{file_path}"
    try:
        resp = requests.get(dl_url, timeout=45.0)
        resp.raise_for_status()
        return True, resp.content, os.path.basename(file_path), ''
    except requests.RequestException as e:
        return False, None, '', str(e)


def _main_keyboard(has_links: bool = False) -> dict:
    rows = [
        [{'text': '📱 Поділитися контактом (прив’язка за телефоном)', 'request_contact': True}],
        [{'text': '📄 Як подати «Резерв+» (PDF)'}, {'text': '📊 Мій статус ВОД'}],
        [{'text': '👥 Подати за іншого студента (представник)'}],
    ]
    if has_links:
        rows.append([{'text': '🔗 Мої прив’язані студенти'}])
    return {'keyboard': rows, 'resize_keyboard': True, 'one_time_keyboard': False}


def _set_dialog_state(conn, chat_id: str, state: str, payload: dict | None = None) -> None:
    conn.execute(
        """
        INSERT INTO telegram_pending_dialog (chat_id, state, payload_json, updated_at)
        VALUES (?, ?, ?, datetime('now', 'localtime'))
        ON CONFLICT(chat_id) DO UPDATE SET
            state = excluded.state,
            payload_json = excluded.payload_json,
            updated_at = datetime('now', 'localtime')
        """,
        (str(chat_id), state, json.dumps(payload or {}, ensure_ascii=False)),
    )


def _get_dialog_state(conn, chat_id: str) -> tuple[str | None, dict]:
    row = conn.execute(
        "SELECT state, payload_json FROM telegram_pending_dialog WHERE chat_id = ?",
        (str(chat_id),),
    ).fetchone()
    if not row:
        return None, {}
    try:
        payload = json.loads(row['payload_json'] or '{}')
    except Exception:
        payload = {}
    return row['state'], payload


def _clear_dialog_state(conn, chat_id: str) -> None:
    conn.execute("DELETE FROM telegram_pending_dialog WHERE chat_id = ?", (str(chat_id),))


def _upsert_student_link(conn, *, student_id: int, chat_id: str, from_user: dict,
                         link_method: str, is_representative: bool = False,
                         tg_phone: str = '', representative_note: str = '') -> int:
    tg_uid = str(from_user.get('id') or '')
    tg_username = (from_user.get('username') or '').strip()
    tg_first = (from_user.get('first_name') or '').strip()
    tg_last = (from_user.get('last_name') or '').strip()
    rep_int = 1 if is_representative else 0

    existing = conn.execute(
        "SELECT id FROM telegram_student_links WHERE student_id = ? AND chat_id = ?",
        (student_id, str(chat_id)),
    ).fetchone()

    if existing:
        conn.execute(
            """
            UPDATE telegram_student_links
            SET telegram_user_id = COALESCE(NULLIF(?, ''), telegram_user_id),
                tg_username = COALESCE(NULLIF(?, ''), tg_username),
                tg_first_name = COALESCE(NULLIF(?, ''), tg_first_name),
                tg_last_name = COALESCE(NULLIF(?, ''), tg_last_name),
                tg_phone = COALESCE(NULLIF(?, ''), tg_phone),
                link_method = ?,
                is_representative = ?,
                representative_note = CASE WHEN ? != '' THEN ? ELSE representative_note END,
                is_active = 1,
                updated_at = datetime('now', 'localtime')
            WHERE id = ?
            """,
            (
                tg_uid, tg_username, tg_first, tg_last, tg_phone,
                link_method, rep_int, representative_note, representative_note,
                existing['id'],
            ),
        )
        return int(existing['id'])

    cur = conn.execute(
        """
        INSERT INTO telegram_student_links (
            student_id, chat_id, telegram_user_id, tg_username,
            tg_first_name, tg_last_name, tg_phone,
            link_method, is_representative, representative_note, is_active
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
        """,
        (
            student_id, str(chat_id), tg_uid, tg_username,
            tg_first, tg_last, tg_phone,
            link_method, rep_int, representative_note or '',
        ),
    )
    return int(cur.lastrowid)


def _get_chat_linked_students(conn, chat_id: str) -> list[dict]:
    rows = conn.execute(
        """
        SELECT l.*, s.last_name_UA, s.first_name_UA, s.middle_name_UA, s.birth_date,
               g.name AS group_name, g.course AS group_course
        FROM telegram_student_links l
        JOIN students s ON s.id = l.student_id
        LEFT JOIN groups g ON g.id = s.group_id
        WHERE l.chat_id = ? AND l.is_active = 1 AND COALESCE(s.archived, 0) = 0
        ORDER BY l.is_representative ASC, l.id DESC
        """,
        (str(chat_id),),
    ).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        d['full_name'] = ' '.join(x for x in (d['last_name_UA'], d['first_name_UA'], d['middle_name_UA'] or '') if x).strip()
        out.append(d)
    return out


def _send_student_status_reply(conn, token: str, chat_id: str) -> None:
    linked = _get_chat_linked_students(conn, chat_id)
    if not linked:
        _tg_send_message(
            token, chat_id,
            "ℹ️ Ваш Telegram ще не прив’язаний до картки студента.\n\n"
            "Оберіть зручний спосіб:\n"
            "1️⃣ Натисніть «📱 Поділитися контактом» — якщо ваш номер вказано в картці студента.\n"
            "2️⃣ Перейдіть за персональним чи груповим посиланням-запрошенням від деканату.\n"
            "3️⃣ Або просто надішліть готовий PDF «Резерв+» — система знайде студента за ПІБ та датою народження.",
            reply_markup=_main_keyboard(False),
        )
        return

    today_ym = date.today().strftime('%Y-%m')
    lines = ["📊 Статус військово-облікових документів:\n"]
    for st in linked:
        sid = st['student_id']
        mil = conn.execute(
            "SELECT id, status_badge, formed_at, valid_until, deferral_until FROM military WHERE student_id = ?",
            (sid,),
        ).fetchone()
        pending_cnt = conn.execute(
            "SELECT COUNT(*) FROM telegram_vod_submissions WHERE student_id = ? AND status = 'pending'",
            (sid,),
        ).fetchone()[0]
        approved_month = conn.execute(
            """
            SELECT COUNT(*) FROM telegram_vod_submissions
            WHERE student_id = ? AND status = 'approved' AND (parsed_ym = ? OR upload_ym = ?)
            """,
            (sid, today_ym, today_ym),
        ).fetchone()[0]

        rep_tag = " (представник)" if st.get('is_representative') else ""
        lines.append(f"👤 {st['full_name']}{rep_tag}")
        lines.append(f"   Група: {st['group_name'] or '—'}")
        if mil:
            lines.append(f"   Статус ВО: {mil['status_badge'] or '—'}")
            lines.append(f"   Останній «Резерв+» у картці: {mil['formed_at'] or 'не вказано'}")
            if mil['deferral_until']:
                lines.append(f"   Відстрочка до: {mil['deferral_until']}")
        else:
            lines.append("   Картка військового обліку ще не заповнена.")
        if pending_cnt:
            lines.append(f"   ⏳ Документів на перевірці: {pending_cnt}")
        if approved_month:
            lines.append(f"   ✅ За {_format_ym_ua(today_ym)} подано та підтверджено.")
        lines.append("")

    _tg_send_message(token, chat_id, "\n".join(lines).strip(), reply_markup=_main_keyboard(True))


def _handle_invite_start(conn, token: str, chat_id: str, from_user: dict, arg: str) -> bool:
    """Обробляє /start s_<token> або /start g_<token>. Повертає True, якщо це було запрошення."""
    if not arg or not (arg.startswith('s_') or arg.startswith('g_')):
        return False

    inv_token = arg[2:].strip()
    inv = conn.execute(
        "SELECT * FROM telegram_invite_tokens WHERE token = ? AND is_active = 1",
        (inv_token,),
    ).fetchone()
    if not inv:
        _tg_send_message(
            token, chat_id,
            "⚠️ Посилання-запрошення недійсне або вже відкликане.\n"
            "Ви можете скористатися кнопкою «📱 Поділитися контактом» або надіслати PDF «Резерв+» напряму.",
            reply_markup=_main_keyboard(bool(_get_chat_linked_students(conn, chat_id))),
        )
        return True

    if inv['expires_at']:
        try:
            exp_dt = datetime.strptime(inv['expires_at'][:19], '%Y-%m-%d %H:%M:%S')
            if datetime.now() > exp_dt:
                _tg_send_message(token, chat_id, "⚠️ Термін дії цього посилання-запрошення закінчився.")
                return True
        except Exception:
            pass

    if inv['max_uses'] is not None and int(inv['uses_count'] or 0) >= int(inv['max_uses']):
        _tg_send_message(token, chat_id, "⚠️ Це посилання-запрошення вже досягло ліміту використань.")
        return True

    if inv['scope'] == 'student' and inv['student_id']:
        st = conn.execute(
            """
            SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA, g.name AS group_name
            FROM students s LEFT JOIN groups g ON g.id = s.group_id
            WHERE s.id = ?
            """,
            (inv['student_id'],),
        ).fetchone()
        if not st:
            _tg_send_message(token, chat_id, "⚠️ Студента за цим запрошенням не знайдено.")
            return True
        full_name = ' '.join(x for x in (st['last_name_UA'], st['first_name_UA'], st['middle_name_UA'] or '') if x).strip()
        is_rep = bool(inv['is_representative_default'])
        note = (inv['note'] or '').strip() or ('Представник (за запрошенням)' if is_rep else '')
        link_id = _upsert_student_link(
            conn,
            student_id=int(st['id']),
            chat_id=chat_id,
            from_user=from_user,
            link_method='invite_student',
            is_representative=is_rep,
            representative_note=note,
        )
        conn.execute(
            "UPDATE telegram_invite_tokens SET uses_count = uses_count + 1 WHERE id = ?",
            (inv['id'],),
        )
        _record_event(
            conn, 'student_linked',
            f"Прив’язано за запрошенням: {full_name}",
            severity='info',
            details=f"chat_id={chat_id}, @{from_user.get('username') or '—'}, представник={'так' if is_rep else 'ні'}",
            student_id=int(st['id']),
        )
        conn.commit()

        inline = {
            'inline_keyboard': [[
                {'text': '✅ Це мій власний акаунт', 'callback_data': f'mark_rep:{link_id}:0'},
                {'text': '👥 Я подаю за студента (представник)', 'callback_data': f'mark_rep:{link_id}:1'},
            ]]
        }
        _tg_send_message(
            token, chat_id,
            f"✅ Telegram успішно прив’язано до картки студента:\n"
            f"👤 {full_name}\n"
            f"🎓 Група: {st['group_name'] or '—'}\n\n"
            f"Тепер ви можете просто надіслати PDF-файл із «Резерв+» у цей чат.\n"
            f"(За потреби вкажіть нижче, чи ви сам студент, чи представник без Telegram у студента):",
            reply_markup=inline,
        )
        _tg_send_message(token, chat_id, "Головне меню:", reply_markup=_main_keyboard(True))
        return True

    if inv['scope'] == 'group' and inv['group_id']:
        grp = conn.execute("SELECT id, name FROM groups WHERE id = ?", (inv['group_id'],)).fetchone()
        if not grp:
            _tg_send_message(token, chat_id, "⚠️ Групу за цим запрошенням не знайдено.")
            return True
        conn.execute(
            "UPDATE telegram_invite_tokens SET uses_count = uses_count + 1 WHERE id = ?",
            (inv['id'],),
        )
        _set_dialog_state(
            conn, chat_id, 'awaiting_group_student_search',
            {'group_id': int(grp['id']), 'group_name': grp['name']},
        )
        conn.commit()
        _tg_send_message(
            token, chat_id,
            f"🎓 Запрошення для групи «{grp['name']}».\n\n"
            f"Щоб прив’язатися до своєї картки:\n"
            f"• натисніть «📱 Поділитися контактом», АБО\n"
            f"• напишіть своє Прізвище та Ім’я (наприклад: Іваненко Іван), АБО\n"
            f"• одразу надішліть PDF «Резерв+».\n\n"
            f"(Для скасування введіть /cancel)",
            reply_markup=_main_keyboard(bool(_get_chat_linked_students(conn, chat_id))),
        )
        return True

    return False


def _handle_contact_share(conn, token: str, chat_id: str, from_user: dict, contact: dict) -> None:
    raw_phone = contact.get('phone_number') or ''
    contact_uid = contact.get('user_id')
    sender_uid = from_user.get('id')
    if contact_uid and sender_uid and str(contact_uid) != str(sender_uid):
        _tg_send_message(
            token, chat_id,
            "⚠️ Будь ласка, поділіться саме власним контактом кнопкою «📱 Поділитися контактом» (а не пересланим контактом іншої людини).",
            reply_markup=_main_keyboard(bool(_get_chat_linked_students(conn, chat_id))),
        )
        return

    norm_phone = normalize_phone_ua(raw_phone)
    if not norm_phone:
        _tg_send_message(token, chat_id, "⚠️ Не вдалося розпізнати номер телефону. Спробуйте ще раз або надішліть PDF «Резерв+».")
        return

    rows = conn.execute(
        f"""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.phone, s.phone_backup, g.name AS group_name
        FROM students s
        LEFT JOIN groups g ON g.id = s.group_id
        WHERE {_active_students_sql_where('s', include_frozen=True)}
          AND (COALESCE(s.phone, '') != '' OR COALESCE(s.phone_backup, '') != '')
        """
    ).fetchall()

    matches = []
    for r in rows:
        p1 = normalize_phone_ua(r['phone'])
        p2 = normalize_phone_ua(r['phone_backup'])
        if norm_phone in (p1, p2):
            full_name = ' '.join(x for x in (r['last_name_UA'], r['first_name_UA'], r['middle_name_UA'] or '') if x).strip()
            matches.append({
                'id': int(r['id']),
                'full_name': full_name,
                'group_name': r['group_name'] or '—',
                'is_backup': (norm_phone == p2 and norm_phone != p1),
            })

    if not matches:
        _tg_send_message(
            token, chat_id,
            f"ℹ️ Номер {norm_phone} не знайдено серед активних студентів у базі даних.\n\n"
            "Що можна зробити:\n"
            "• Надішліть свій PDF-файл із «Резерв+» прямо сюди — бот знайде вас за ПІБ та датою народження і прив’яже акаунт;\n"
            "• Або попросіть у відповідального за військовий облік / деканату персональне посилання-запрошення (особливо якщо ви подаєте документ за студента як представник).",
            reply_markup=_main_keyboard(bool(_get_chat_linked_students(conn, chat_id))),
        )
        return

    if len(matches) == 1:
        m = matches[0]
        is_rep = bool(m['is_backup'])
        link_id = _upsert_student_link(
            conn,
            student_id=m['id'],
            chat_id=chat_id,
            from_user=from_user,
            link_method='contact',
            is_representative=is_rep,
            tg_phone=norm_phone,
            representative_note=('Збіг за резервним телефоном' if is_rep else ''),
        )
        _clear_dialog_state(conn, chat_id)
        _record_event(
            conn, 'student_linked',
            f"Прив’язано за телефоном: {m['full_name']}",
            severity='info',
            details=f"тел. {norm_phone}, chat_id={chat_id}",
            student_id=m['id'],
        )
        conn.commit()
        _tg_send_message(
            token, chat_id,
            f"✅ Контакт підтверджено! Профіль прив’язано:\n"
            f"👤 {m['full_name']}\n"
            f"🎓 Група: {m['group_name']}\n\n"
            f"Тепер надсилайте PDF-файл із застосунку «Резерв+» у цей чат.",
            reply_markup=_main_keyboard(True),
        )
        if is_rep:
            _tg_send_message(
                token, chat_id,
                "Примітка: номер збігся з додатковим телефоном у картці. Якщо ви є власником картки, а не представником — натисніть нижче:",
                reply_markup={'inline_keyboard': [[
                    {'text': '✅ Це моя картка', 'callback_data': f'mark_rep:{link_id}:0'},
                    {'text': '👥 Я представник', 'callback_data': f'mark_rep:{link_id}:1'},
                ]]},
            )
        return

    # Кілька студентів з тим самим номером (напр., телефон батьків вказано у кількох дітей)
    kb = []
    for m in matches[:10]:
        kb.append([{'text': f"👤 {m['full_name']} ({m['group_name']})", 'callback_data': f"bind_sel:{m['id']}:1"}])
    _tg_send_message(
        token, chat_id,
        f"За номером {norm_phone} знайдено кілька студентів. Оберіть, кого прив’язати до цього чату (можна по черзі кількох):",
        reply_markup={'inline_keyboard': kb},
    )


def _handle_pdf_document(conn, token: str, chat_id: str, from_user: dict, document: dict) -> None:
    file_name = (document.get('file_name') or 'reserve_plus.pdf').strip()
    mime = (document.get('mime_type') or '').lower()
    file_size = int(document.get('file_size') or 0)
    file_id = document.get('file_id') or ''
    file_unique_id = document.get('file_unique_id') or ''

    if not file_name.lower().endswith('.pdf') and 'pdf' not in mime:
        _tg_send_message(
            token, chat_id,
            "⚠️ Будь ласка, надішліть саме PDF-файл, вивантажений із застосунку «Резерв+» (розширення .pdf). Фото та скріншоти не приймаються.",
        )
        return

    if file_size > 20 * 1024 * 1024:
        _tg_send_message(token, chat_id, "⚠️ Файл занадто великий (максимум 20 МБ).")
        return

    ok_dl, pdf_bytes, _, dl_err = _tg_download_file(token, file_id)
    if not ok_dl or not pdf_bytes:
        _tg_send_message(token, chat_id, f"⚠️ Не вдалося завантажити файл із Telegram: {dl_err}")
        return

    if b'%PDF-' not in pdf_bytes[:32]:
        _tg_send_message(token, chat_id, "⚠️ Надісланий файл не є коректним PDF-документом.")
        return

    os.makedirs(TELEGRAM_VOD_ABS_DIR, exist_ok=True)
    safe_pdf_name = f"vod_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{secrets.token_hex(6)}.pdf"
    pdf_abs = os.path.join(TELEGRAM_VOD_ABS_DIR, safe_pdf_name)
    pdf_rel = f"uploads/telegram_vod/{safe_pdf_name}"
    with open(pdf_abs, 'wb') as f:
        f.write(pdf_bytes)

    try:
        parsed = parse_reserve_plus_pdf(pdf_abs)
    except Exception as e:
        logger.warning(f"Telegram PDF parse failed (chat_id={chat_id}, file={file_name}): {e}")
        try:
            os.remove(pdf_abs)
        except OSError:
            pass
        _tg_send_message(
            token, chat_id,
            "❌ Не вдалося розпізнати документ як витяг із «Резерв+».\n"
            "Переконайтеся, що це оригінальний PDF-файл «Резерв+» (розширений витяг із застосунку, без перезбереження як фото).",
        )
        return

    parsed_full_name = (parsed.get('full_name') or '').strip()
    from routes.students.military import _normalize_date
    parsed_birth_date = _normalize_date(parsed.get('birth_date') or '')
    parsed_formed_at = (parsed.get('formed_at') or '').strip()
    parsed_ym = _parse_ym(parsed_formed_at)
    upload_ym = date.today().strftime('%Y-%m')

    # Перевірка фото 3x4
    photo_rel = ''
    photo_raw_rel = ''
    try:
        raw_photo_bytes = extract_photo_from_pdf(pdf_abs)
        if raw_photo_bytes:
            raw_p = save_raw_pending_photo(raw_photo_bytes)
            crop_p = process_and_save_pending_photo_smart(raw_photo_bytes)
            # Копіюємо у telegram_vod, щоб права доступу в security.py чітко відповідали manage_military_telegram
            if crop_p:
                src_crop = os.path.join(PROJECT_ROOT, 'static', crop_p)
                if os.path.exists(src_crop):
                    pname = f"photo_{secrets.token_hex(8)}.jpg"
                    with open(src_crop, 'rb') as sf, open(os.path.join(TELEGRAM_VOD_ABS_DIR, pname), 'wb') as df:
                        df.write(sf.read())
                    photo_rel = f"uploads/telegram_vod/{pname}"
                    try:
                        os.remove(src_crop)
                    except OSError:
                        pass
            if raw_p:
                src_raw = os.path.join(PROJECT_ROOT, 'static', raw_p)
                if os.path.exists(src_raw):
                    rname = f"raw_{secrets.token_hex(8)}.jpg"
                    with open(src_raw, 'rb') as sf, open(os.path.join(TELEGRAM_VOD_ABS_DIR, rname), 'wb') as df:
                        df.write(sf.read())
                    photo_raw_rel = f"uploads/telegram_vod/{rname}"
                    try:
                        os.remove(src_raw)
                    except OSError:
                        pass
    except Exception as e:
        logger.warning(f"Telegram PDF photo extract warning: {e}")

    # Суворий пошук студента за ПІБ + Дата народження
    student, not_found_reason = _find_active_student_by_fio_and_birth(
        conn, parsed_full_name, parsed_birth_date
    )

    tg_uid = str(from_user.get('id') or '')
    tg_username = (from_user.get('username') or '').strip()
    tg_first = (from_user.get('first_name') or '').strip()
    tg_last = (from_user.get('last_name') or '').strip()

    linked_rows = _get_chat_linked_students(conn, chat_id)
    linked_phone = next((l.get('tg_phone') for l in linked_rows if l.get('tg_phone')), '')

    if not student:
        cur = conn.execute(
            """
            INSERT INTO telegram_vod_submissions (
                student_id, chat_id, telegram_user_id, tg_username, tg_first_name, tg_last_name, tg_phone,
                file_id, file_unique_id, original_filename, pdf_rel_path, photo_rel_path, photo_raw_rel_path,
                parsed_full_name, parsed_birth_date, parsed_formed_at, parsed_ym, upload_ym,
                parsed_json, status, parse_error
            ) VALUES (NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'not_found', ?)
            """,
            (
                str(chat_id), tg_uid, tg_username, tg_first, tg_last, linked_phone,
                file_id, file_unique_id, file_name, pdf_rel, photo_rel, photo_raw_rel,
                parsed_full_name, parsed_birth_date, parsed_formed_at, parsed_ym, upload_ym,
                json.dumps(parsed, ensure_ascii=False, default=str),
                not_found_reason,
            ),
        )
        sub_id = int(cur.lastrowid)
        _record_event(
            conn, 'not_found_pdf',
            f"Студента не знайдено в БД: {parsed_full_name or 'невідомо'} ({parsed_birth_date or '—'})",
            severity='danger',
            details=f"{not_found_reason}. Відправник Telegram: @{tg_username or chat_id}",
            submission_id=sub_id,
        )
        conn.commit()
        _tg_send_message(
            token, chat_id,
            f"❌ Студента з даними з PDF:\n"
            f"• ПІБ: {parsed_full_name or '—'}\n"
            f"• Дата народження: {parsed_birth_date or '—'}\n"
            f"НЕ знайдено в базі активних студентів університету.\n\n"
            f"Причина: {not_found_reason}\n"
            f"Повідомлення передано відповідальному працівнику. Якщо сталася помилка в ПІБ чи даті народження — зверніться до деканату.",
            reply_markup=_main_keyboard(bool(linked_rows)),
        )
        return

    sid = int(student['id'])
    full_student_name = ' '.join(
        x for x in (student['last_name_UA'], student['first_name_UA'], student['middle_name_UA'] or '') if x
    ).strip()

    # Перевіряємо, чи цей chat_id подає як представник (або чи це 3-й спосіб автоприв'язки)
    dialog_state, _ = _get_dialog_state(conn, chat_id)
    in_proxy_mode = (dialog_state == 'awaiting_proxy_pdf')
    existing_link = next((l for l in linked_rows if int(l['student_id']) == sid), None)
    auto_bound_link_id = None
    submitted_as_rep = False

    if existing_link:
        submitted_as_rep = bool(existing_link.get('is_representative')) or in_proxy_mode
    else:
        # Спосіб 3: Автоприв'язка / прив'язка представника при надсиланні PDF
        is_rep_auto = bool(in_proxy_mode or len(linked_rows) > 0)
        rep_note = 'Представник (подав PDF за студента)' if is_rep_auto else 'Автоприв’язка за PDF'
        auto_bound_link_id = _upsert_student_link(
            conn,
            student_id=sid,
            chat_id=chat_id,
            from_user=from_user,
            link_method='pdf_auto',
            is_representative=is_rep_auto,
            representative_note=rep_note,
        )
        submitted_as_rep = is_rep_auto
        _record_event(
            conn, 'student_linked',
            f"Автоприв’язка через PDF: {full_student_name}",
            severity='info',
            details=f"chat_id={chat_id}, @{tg_username or '—'}, представник={'так' if is_rep_auto else 'ні'}",
            student_id=sid,
        )

    _clear_dialog_state(conn, chat_id)

    # Перевірка повторного завантаження за місяць
    is_dup, dup_cnt, dup_info = _compute_monthly_duplicate_info(conn, sid, parsed_ym, upload_ym)
    is_outdated = bool(parsed_ym and parsed_ym < upload_ym)

    new_military, new_student, mil_diff, stu_diff = _build_submission_diffs(conn, sid, parsed)
    risk_flags, risk_level = _analyze_submission_risks(conn, sid, parsed)

    cur = conn.execute(
        """
        INSERT INTO telegram_vod_submissions (
            student_id, chat_id, telegram_user_id, tg_username, tg_first_name, tg_last_name, tg_phone,
            file_id, file_unique_id, original_filename, pdf_rel_path, photo_rel_path, photo_raw_rel_path,
            parsed_full_name, parsed_birth_date, parsed_formed_at, parsed_ym, upload_ym,
            parsed_json, new_military_json, new_student_json, military_diff_json, student_diff_json,
            status, submitted_as_representative,
            is_monthly_duplicate, duplicate_count, duplicate_info, is_outdated_pdf,
            risk_flags_json, risk_level, risk_count
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            sid, str(chat_id), tg_uid, tg_username, tg_first, tg_last, linked_phone,
            file_id, file_unique_id, file_name, pdf_rel, photo_rel, photo_raw_rel,
            parsed_full_name, parsed_birth_date, parsed_formed_at, parsed_ym, upload_ym,
            json.dumps(parsed, ensure_ascii=False, default=str),
            json.dumps(new_military, ensure_ascii=False),
            json.dumps(new_student, ensure_ascii=False),
            json.dumps(mil_diff, ensure_ascii=False),
            json.dumps(stu_diff, ensure_ascii=False),
            1 if submitted_as_rep else 0,
            1 if is_dup else 0,
            dup_cnt if is_dup else 1,
            dup_info,
            1 if is_outdated else 0,
            json.dumps(risk_flags, ensure_ascii=False),
            risk_level,
            len(risk_flags),
        ),
    )
    sub_id = int(cur.lastrowid)

    if is_dup:
        _record_event(
            conn, 'monthly_duplicate',
            f"Повторне завантаження ВОД за місяць ({dup_cnt}-й раз): {full_student_name}",
            severity='warning',
            details=dup_info,
            student_id=sid,
            submission_id=sub_id,
        )

    if risk_flags:
        risk_summary = '; '.join(rf['detail'] for rf in risk_flags[:4])
        _record_event(
            conn, 'critical_discrepancy',
            f"{'Критичні розбіжності' if risk_level == 'critical' else 'Попередження'} у ВОД: {full_student_name}",
            severity='danger' if risk_level == 'critical' else 'warning',
            details=risk_summary,
            student_id=sid,
            submission_id=sub_id,
        )

    conn.commit()

    reply_lines = [
        f"📥 Документ «Резерв+» прийнято на перевірку (#{sub_id})!",
        f"👤 Студент: {full_student_name}",
        f"🎓 Група: {student.get('group_name') or '—'}",
        f"🗓 Сформовано в PDF: {parsed_formed_at or 'не вказано'}",
        f"🔎 Виявлено змін: військових полів — {len(mil_diff)}, картки — {len(stu_diff)}.",
    ]
    if is_dup:
        reply_lines.append(
            f"\nℹ️ Увага: це вже {dup_cnt}-й документ за цей місяць. Повідомлення про повторну подачу надіслано відповідальному працівнику."
        )
    if is_outdated:
        reply_lines.append(
            f"\n⚠️ Примітка: у PDF вказано дату формування за минулий місяць ({_format_ym_ua(parsed_ym)}). Якщо потрібен свіжий витяг — сформуйте новий документ у «Резерв+»."
        )
    reply_lines.append("\nПісля перевірки відповідальним працівником ви отримаєте підтвердження у цьому чаті.")

    _tg_send_message(token, chat_id, "\n".join(reply_lines), reply_markup=_main_keyboard(True))

    if auto_bound_link_id is not None and not in_proxy_mode:
        _tg_send_message(
            token, chat_id,
            f"🔗 Ми автоматично прив’язали ваш Telegram до картки студента {full_student_name}. Підкажіть, будь ласка:",
            reply_markup={'inline_keyboard': [[
                {'text': '✅ Це мій власний акаунт', 'callback_data': f'mark_rep:{auto_bound_link_id}:0'},
                {'text': '👥 Я подаю за студента (представник)', 'callback_data': f'mark_rep:{auto_bound_link_id}:1'},
            ]]},
        )


def _handle_callback_query(conn, token: str, cq: dict) -> None:
    cq_id = cq.get('id')
    from_user = cq.get('from') or {}
    msg = cq.get('message') or {}
    chat_id = str((msg.get('chat') or {}).get('id') or '')
    data = (cq.get('data') or '').strip()

    if not chat_id or not data:
        if cq_id:
            _tg_answer_callback(token, cq_id)
        return

    parts = data.split(':')
    action = parts[0]

    if action == 'mark_rep' and len(parts) == 3:
        link_id = int(parts[1])
        rep_val = 1 if parts[2] == '1' else 0
        row = conn.execute(
            """
            SELECT l.id, l.student_id, s.last_name_UA, s.first_name_UA, s.middle_name_UA
            FROM telegram_student_links l
            JOIN students s ON s.id = l.student_id
            WHERE l.id = ? AND l.chat_id = ?
            """,
            (link_id, chat_id),
        ).fetchone()
        if row:
            note = 'Представник (позначено в боті)' if rep_val else ''
            conn.execute(
                "UPDATE telegram_student_links SET is_representative = ?, representative_note = ?, updated_at = datetime('now', 'localtime') WHERE id = ?",
                (rep_val, note, link_id),
            )
            conn.commit()
            fname = ' '.join(x for x in (row['last_name_UA'], row['first_name_UA'], row['middle_name_UA'] or '') if x).strip()
            msg_txt = (
                f"👥 Позначено як представника для студента {fname}. Тепер ви можете також надсилати документи й за інших студентів через кнопку «👥 Подати за іншого студента»."
                if rep_val else
                f"✅ Підтверджено: це ваш особистий акаунт ({fname})."
            )
            _tg_answer_callback(token, cq_id, "Збережено")
            _tg_send_message(token, chat_id, msg_txt, reply_markup=_main_keyboard(True))
        else:
            _tg_answer_callback(token, cq_id, "Запис не знайдено")
        return

    if action == 'bind_sel' and len(parts) == 3:
        sid = int(parts[1])
        is_rep = (parts[2] == '1')
        st = conn.execute(
            "SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA, g.name AS group_name "
            "FROM students s LEFT JOIN groups g ON g.id = s.group_id WHERE s.id = ?",
            (sid,),
        ).fetchone()
        if st:
            fname = ' '.join(x for x in (st['last_name_UA'], st['first_name_UA'], st['middle_name_UA'] or '') if x).strip()
            link_id = _upsert_student_link(
                conn,
                student_id=sid,
                chat_id=chat_id,
                from_user=from_user,
                link_method='invite_group' if not is_rep else 'contact',
                is_representative=is_rep,
                representative_note=('Обрано зі списку в боті' if is_rep else ''),
            )
            _clear_dialog_state(conn, chat_id)
            conn.commit()
            _tg_answer_callback(token, cq_id, f"Прив’язано: {fname}")
            _tg_send_message(
                token, chat_id,
                f"✅ Прив’язано студента: {fname} ({st['group_name'] or '—'}).\nМожете надсилати PDF «Резерв+».",
                reply_markup={'inline_keyboard': [[
                    {'text': '✅ Це мій акаунт', 'callback_data': f'mark_rep:{link_id}:0'},
                    {'text': '👥 Я представник', 'callback_data': f'mark_rep:{link_id}:1'},
                ]]},
            )
        else:
            _tg_answer_callback(token, cq_id, "Студента не знайдено")
        return

    if cq_id:
        _tg_answer_callback(token, cq_id)


def _process_telegram_update(conn, token: str, settings: dict, update: dict) -> None:
    if 'callback_query' in update:
        _handle_callback_query(conn, token, update['callback_query'])
        return

    msg = update.get('message')
    if not isinstance(msg, dict):
        return
    chat = msg.get('chat') or {}
    if chat.get('type') != 'private':
        return
    chat_id = str(chat.get('id') or '')
    if not chat_id:
        return
    from_user = msg.get('from') or {}

    if msg.get('contact'):
        _handle_contact_share(conn, token, chat_id, from_user, msg['contact'])
        return

    if msg.get('document'):
        _handle_pdf_document(conn, token, chat_id, from_user, msg['document'])
        return

    text = (msg.get('text') or '').strip()
    if not text:
        return

    linked = _get_chat_linked_students(conn, chat_id)
    has_links = len(linked) > 0

    if text.startswith('/start'):
        _clear_dialog_state(conn, chat_id)
        parts = text.split(maxsplit=1)
        arg = parts[1].strip() if len(parts) > 1 else ''
        if arg and _handle_invite_start(conn, token, chat_id, from_user, arg):
            return
        welcome = (settings.get('welcome_text') or DEFAULT_WELCOME_TEXT).strip()
        _tg_send_message(
            token, chat_id,
            f"{welcome}\n\n"
            "Для прив’язки натисніть кнопку «📱 Поділитися контактом» нижче або просто надішліть PDF-файл із «Резерв+».",
            reply_markup=_main_keyboard(has_links),
        )
        return

    if text == '/cancel':
        _clear_dialog_state(conn, chat_id)
        conn.commit()
        _tg_send_message(token, chat_id, "Дію скасовано.", reply_markup=_main_keyboard(has_links))
        return

    if text in ('/status', '📊 Мій статус ВОД'):
        _send_student_status_reply(conn, token, chat_id)
        return

    if text in ('/help', '📄 Як подати «Резерв+» (PDF)'):
        _tg_send_message(
            token, chat_id,
            "📄 Як подати документ «Резерв+»:\n"
            "1. Відкрийте застосунок «Резерв+» на смартфоні.\n"
            "2. Натисніть «⋮» → «Завантажити PDF» (або «Згенерувати PDF»).\n"
            "3. Надішліть отриманий .pdf файл прямо у цей чат як документ.\n\n"
            "Система автоматично перевірить ПІБ і дату народження у PDF та передасть документ відповідальному працівнику.",
            reply_markup=_main_keyboard(has_links),
        )
        return

    if text == '🔗 Мої прив’язані студенти':
        if not linked:
            _tg_send_message(token, chat_id, "Наразі немає прив’язаних студентів.", reply_markup=_main_keyboard(False))
            return
        lines = ["🔗 Студенти, прив’язані до цього Telegram-акаунта:\n"]
        for st in linked:
            role = "представник" if st.get('is_representative') else "особистий"
            lines.append(f"• {st['full_name']} ({st['group_name'] or '—'}) — {role}")
        _tg_send_message(token, chat_id, "\n".join(lines), reply_markup=_main_keyboard(True))
        return

    if text == '👥 Подати за іншого студента (представник)':
        _set_dialog_state(conn, chat_id, 'awaiting_proxy_pdf', {})
        conn.commit()
        _tg_send_message(
            token, chat_id,
            "👥 Режим представника активовано.\n\n"
            "Якщо у студента немає Telegram (або ви батьки/родич/староста):\n"
            "Просто надішліть у цей чат PDF-файл «Резерв+» цього студента — система знайде його за ПІБ та датою народження, "
            "позначить вас як представника і дозволить далі отримувати нагадування для нього.\n\n"
            "(Для виходу введіть /cancel)",
            reply_markup=_main_keyboard(has_links),
        )
        return

    # Стан пошуку в групі (після переходу за груповим QR / посиланням g_<token>)
    dialog_state, dialog_payload = _get_dialog_state(conn, chat_id)
    if dialog_state == 'awaiting_group_student_search' and dialog_payload.get('group_id'):
        gid = int(dialog_payload['group_id'])
        q = ' '.join(_norm_apos(text).split()).casefold()
        if len(q) < 2:
            _tg_send_message(token, chat_id, "Введіть принаймні 2 символи прізвища або надішліть PDF «Резерв+».")
            return
        rows = conn.execute(
            f"""
            SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA, g.name AS group_name
            FROM students s LEFT JOIN groups g ON g.id = s.group_id
            WHERE s.group_id = ? AND {_active_students_sql_where('s', include_frozen=True)}
            ORDER BY s.last_name_UA COLLATE UKRAINIAN
            """,
            (gid,),
        ).fetchall()
        found = []
        tokens = q.split()
        for r in rows:
            fn = ' '.join(x for x in (r['last_name_UA'], r['first_name_UA'], r['middle_name_UA'] or '') if x).strip()
            fn_norm = _norm_apos(fn).casefold()
            if all(t in fn_norm for t in tokens):
                found.append({'id': int(r['id']), 'full_name': fn, 'group_name': r['group_name'] or ''})
        if not found:
            _tg_send_message(
                token, chat_id,
                f"У групі «{dialog_payload.get('group_name')}» не знайдено студента за запитом «{text}». Спробуйте інше написання або надішліть PDF «Резерв+».",
            )
            return
        kb = [
            [{'text': f"👤 {m['full_name']}", 'callback_data': f"bind_sel:{m['id']}:0"}]
            for m in found[:10]
        ]
        _tg_send_message(
            token, chat_id,
            "Оберіть студента зі списку для прив’язки:",
            reply_markup={'inline_keyboard': kb},
        )
        return

    _tg_send_message(
        token, chat_id,
        "Надішліть PDF-файл із застосунку «Резерв+» або скористайтеся кнопками меню нижче.",
        reply_markup=_main_keyboard(has_links),
    )


def _is_day_in_window(day: int, day_from: int, day_to: int) -> bool:
    """Перевіряє, чи входить день місяця у діапазон [day_from..day_to] (підтримує перехід через кінець місяця)."""
    d_from = max(1, min(31, int(day_from or 1)))
    d_to = max(1, min(31, int(day_to or 31)))
    if d_from <= d_to:
        return d_from <= day <= d_to
    return day >= d_from or day <= d_to


def _compute_daily_reminder_slots(times_per_day: int, start_hour: int, end_hour: int) -> list[int]:
    """
    Повертає відсортований список хвилин від початку доби (0..1439) для `times_per_day` запусків
    рівномірно в межах [start_hour:00 .. end_hour:00].
    """
    n = max(1, min(24, int(times_per_day or 1)))
    sh = max(0, min(23, int(start_hour if start_hour is not None else 9)))
    eh = max(0, min(23, int(end_hour if end_hour is not None else 20)))
    if eh < sh:
        eh = sh
    s_min = sh * 60
    e_min = eh * 60
    if n == 1:
        return [s_min]
    span = e_min - s_min
    if span <= 0:
        return [min(23 * 60 + 55, s_min + i * 30) for i in range(n)]
    return [int(round(s_min + i * span / (n - 1))) for i in range(n)]


def _format_daily_slots(times_per_day: int, start_hour: int, end_hour: int) -> list[str]:
    return [f"{m // 60:02d}:{m % 60:02d}" for m in _compute_daily_reminder_slots(times_per_day, start_hour, end_hour)]


def _prepare_rule_filter(raw_filter_json: str, ym_now: str, stop_when_submitted: bool) -> dict:
    try:
        flt = json.loads(raw_filter_json or '{}')
    except Exception:
        flt = {}
    if not isinstance(flt, dict):
        flt = {}
    flt['vod_month'] = ym_now
    if 'only_base_license' not in flt and not flt.get('license_ids'):
        flt['only_base_license'] = True
    if stop_when_submitted:
        vms = flt.get('vod_month_state')
        if not vms or vms in ('any', 'already_submitted_or_pending', 'has_approved_this_month'):
            flt['vod_month_state'] = 'missing_this_month'
    return flt


def _execute_auto_rule_dispatch(conn, token: str, rule: dict, *, now: datetime | None = None,
                                slot_idx: int | None = None, created_by: str = 'auto_rule') -> dict:
    now_dt = now or datetime.now()
    ym_now = now_dt.strftime('%Y-%m')
    today_str = now_dt.strftime('%Y-%m-%d')
    now_str = now_dt.strftime('%Y-%m-%d %H:%M:%S')

    stop_sub = bool(int(rule.get('stop_when_submitted') if rule.get('stop_when_submitted') is not None else 1))
    flt = _prepare_rule_filter(rule.get('filter_json') or '{}', ym_now, stop_sub)
    template = (rule.get('message_template') or DEFAULT_REMINDER_TEXT).strip()
    rule_id = int(rule['id'])

    prev_date = (rule.get('last_sent_date') or '').strip()
    new_today_cnt = (int(rule.get('sent_today_count') or 0) + 1) if prev_date == today_str else 1
    new_slot_idx = int(slot_idx) if slot_idx is not None else int(rule.get('last_sent_slot_idx') if prev_date == today_str else -1)

    # Попередньо фіксуємо запуск у БД до мережевої відправки, щоб виключити дублі при збоях
    conn.execute(
        """
        UPDATE telegram_auto_reminder_rules
        SET last_sent_date = ?,
            last_sent_slot_idx = ?,
            sent_today_count = ?,
            last_sent_at = ?,
            total_runs = COALESCE(total_runs, 0) + 1,
            updated_at = datetime('now', 'localtime')
        WHERE id = ?
        """,
        (today_str, new_slot_idx, new_today_cnt, now_str, rule_id),
    )
    conn.commit()

    summary = execute_broadcast(
        conn,
        token=token,
        title=f"Авто-нагадування «{rule.get('name') or ('Правило #' + str(rule_id))}» ({_format_ym_ua(ym_now)})",
        message_template=template,
        raw_filters=flt,
        created_by=created_by,
        is_auto=True,
        skip_if_empty=True,
    )

    if summary.get('skipped_empty'):
        if summary['total_matched'] == 0:
            run_msg = f"{now_str[:16]} — усі за фільтром уже надали PDF (збігів: 0)"
        else:
            run_msg = f"{now_str[:16]} — збігів: {summary['total_matched']}, але 0 з прив’язаним Telegram"
    else:
        run_msg = (
            f"{now_str[:16]} — підійшло {summary['total_matched']}, "
            f"надіслано {summary['sent_chats']}, помилок {summary['failed_chats']}"
        )

    conn.execute(
        "UPDATE telegram_auto_reminder_rules SET last_run_summary = ? WHERE id = ?",
        (run_msg[:300], rule_id),
    )
    conn.commit()
    return summary


def _is_day_in_escalation_window(day: int, day_from: int, day_to: int, last_days: int,
                                 ref_date: date | None = None) -> bool:
    """
    Перевіряє, чи входить день місяця в останні `last_days` днів вікна розсилки [day_from..day_to]
    (наприклад, для вікна 15..25 і last_days=3 повертає True для 23, 24, 25 чисел).
    """
    if not _is_day_in_window(day, day_from, day_to):
        return False
    d_to = max(1, min(31, int(day_to or 25)))
    n_last = max(1, min(15, int(last_days or 3)))
    ref = ref_date or date.today()
    if day <= d_to:
        dist = d_to - day
    else:
        days_in_month = calendar.monthrange(ref.year, ref.month)[1]
        dist = (days_in_month - day) + d_to
    return 0 <= dist < n_last


def _send_chunked_tg_message(token: str, chat_id: str, text: str) -> tuple[bool, str]:
    """Надсилає довгий звіт у Telegram, розбиваючи по рядках до 3900 символів."""
    txt = (text or '').strip()
    if not txt:
        return True, ''
    if len(txt) <= 3900:
        return _tg_send_message(token, chat_id, txt)

    chunks = []
    cur_lines = []
    cur_len = 0
    for line in txt.splitlines():
        add_len = len(line) + 1
        if cur_lines and cur_len + add_len > 3900:
            chunks.append('\n'.join(cur_lines))
            cur_lines = [line]
            cur_len = len(line)
        else:
            cur_lines.append(line)
            cur_len += add_len
    if cur_lines:
        chunks.append('\n'.join(cur_lines))

    last_err = ''
    all_ok = True
    for ch in chunks:
        ok, err = _tg_send_message(token, chat_id, ch)
        if not ok:
            all_ok = False
            last_err = err
    return all_ok, last_err


def _execute_rule_escalation(conn, token: str, rule: dict, *, now: datetime | None = None,
                             force: bool = False) -> dict:
    """
    Формує та розсилає ескалаційне зведення по групах (хто здав / не здав «Резерв+»)
    старостам, кураторам груп та/або додатковим координаторам.
    """
    now_dt = now or datetime.now()
    ym_now = now_dt.strftime('%Y-%m')
    today_str = now_dt.strftime('%Y-%m-%d')
    rule_id = int(rule['id'])
    rule_name = rule.get('name') or f"Правило #{rule_id}"

    try:
        flt = json.loads(rule.get('filter_json') or '{}')
    except Exception:
        flt = {}
    if not isinstance(flt, dict):
        flt = {}

    summary = build_group_vod_compliance_summary(
        conn, flt, vod_month=ym_now, only_with_missing=True
    )

    if not force:
        conn.execute(
            """
            UPDATE telegram_auto_reminder_rules
            SET escalation_last_date = ?, updated_at = datetime('now', 'localtime')
            WHERE id = ?
            """,
            (today_str, rule_id),
        )
        conn.commit()

    send_leaders = bool(int(rule.get('escalation_send_to_leaders') if rule.get('escalation_send_to_leaders') is not None else 1))
    extra_raw = (rule.get('escalation_extra_chat_ids') or '').strip()
    extra_cids = [c.strip() for c in re.split(r'[\s,;]+', extra_raw) if c.strip()]

    sent_leaders = 0
    sent_extra = 0
    failed = 0
    groups_missing = summary['groups']

    if not groups_missing:
        esc_msg = f"{today_str} — боржників за {_format_ym_ua(ym_now)} немає"
        conn.execute(
            "UPDATE telegram_auto_reminder_rules SET escalation_last_summary = ? WHERE id = ?",
            (esc_msg[:300], rule_id),
        )
        conn.commit()
        return {'groups_missing': 0, 'sent_leaders': 0, 'sent_extra': 0, 'failed': 0, 'missing_students': 0}

    if send_leaders:
        for grp in groups_missing:
            cinfo = grp.get('contact') or {}
            if cinfo and not int(cinfo.get('is_active') if cinfo.get('is_active') is not None else 1):
                continue
            targets = []
            for key in ('leader_chat_id', 'curator_chat_id'):
                cid = (cinfo.get(key) or '').strip()
                if cid and cid not in targets:
                    targets.append(cid)
            if not targets:
                continue
            msg_txt = _format_group_escalation_text(grp, ym_now, rule_name=rule_name)
            for cid in targets:
                ok, _err = _send_chunked_tg_message(token, cid, msg_txt)
                if ok:
                    sent_leaders += 1
                else:
                    failed += 1

    if extra_cids:
        cons_txt = _format_consolidated_escalation_text(summary, ym_now, rule_name=rule_name)
        seen_extra = set()
        for cid in extra_cids:
            if cid in seen_extra:
                continue
            seen_extra.add(cid)
            ok, _err = _send_chunked_tg_message(token, cid, cons_txt)
            if ok:
                sent_extra += 1
            else:
                failed += 1

    esc_msg = (
        f"{today_str} — груп з боржниками: {len(groups_missing)} "
        f"(студентів: {summary['totals']['missing_count']}), "
        f"старостам/кураторам: {sent_leaders}, координаторам: {sent_extra}, помилок: {failed}"
    )
    conn.execute(
        "UPDATE telegram_auto_reminder_rules SET escalation_last_date = ?, escalation_last_summary = ? WHERE id = ?",
        (today_str, esc_msg[:300], rule_id),
    )
    _record_event(
        conn, 'escalation_sent',
        f"Ескалація по «{rule_name}»: {summary['totals']['missing_count']} боржн. у {len(groups_missing)} гр.",
        severity='warning' if failed else 'info',
        details=esc_msg,
    )
    conn.commit()
    return {
        'groups_missing': len(groups_missing),
        'missing_students': summary['totals']['missing_count'],
        'sent_leaders': sent_leaders,
        'sent_extra': sent_extra,
        'failed': failed,
    }


def _cleanup_target_where_sql(*, include_not_found: bool, include_approved: bool) -> str:
    clauses = [
        "sub.status = 'rejected'",
        "(sub.is_monthly_duplicate = 1 AND sub.status IN ('approved', 'rejected'))",
    ]
    if include_not_found:
        clauses.append("sub.status = 'not_found'")
    if include_approved:
        clauses.append(
            "(sub.status = 'approved' AND sub.approved_attachment_id IS NOT NULL "
            "AND EXISTS (SELECT 1 FROM attachments a WHERE a.id = sub.approved_attachment_id))"
        )
    return '(' + ' OR '.join(clauses) + ')'


def _safe_resolve_vod_file(rel_path: str) -> str | None:
    if not rel_path or not isinstance(rel_path, str):
        return None
    static_root = os.path.join(PROJECT_ROOT, 'static')
    abs_p = resolve_under_dir(static_root, rel_path)
    if not abs_p:
        return None
    vod_real = os.path.realpath(TELEGRAM_VOD_ABS_DIR)
    if abs_p != vod_real and not abs_p.startswith(vod_real + os.sep):
        return None
    return abs_p


def get_telegram_vod_storage_stats(
    conn,
    *,
    months: int = 6,
    include_not_found: bool = True,
    include_approved: bool = False,
) -> dict:
    """Рахує поточний обсяг папки telegram_vod/ та кандидатів на очищення."""
    months = max(1, min(60, int(months or 6)))
    cutoff_str = (datetime.now() - timedelta(days=months * 30)).strftime('%Y-%m-%d %H:%M:%S')
    disk_files = 0
    disk_bytes = 0
    if os.path.isdir(TELEGRAM_VOD_ABS_DIR):
        for root_d, _dirs, files in os.walk(TELEGRAM_VOD_ABS_DIR):
            for fn in files:
                fp = os.path.join(root_d, fn)
                try:
                    if os.path.isfile(fp):
                        disk_files += 1
                        disk_bytes += os.path.getsize(fp)
                except OSError:
                    pass

    where_status = _cleanup_target_where_sql(
        include_not_found=include_not_found,
        include_approved=include_approved,
    )
    cand_rows = conn.execute(
        f"""
        SELECT sub.id, sub.pdf_rel_path, sub.photo_rel_path, sub.photo_raw_rel_path
        FROM telegram_vod_submissions sub
        WHERE COALESCE(NULLIF(sub.reviewed_at, ''), sub.created_at) < ?
          AND (COALESCE(sub.pdf_rel_path, '') != ''
               OR COALESCE(sub.photo_rel_path, '') != ''
               OR COALESCE(sub.photo_raw_rel_path, '') != '')
          AND {where_status}
        """,
        (cutoff_str,),
    ).fetchall()

    cand_bytes = 0
    cand_files = 0
    for r in cand_rows:
        for col in ('pdf_rel_path', 'photo_rel_path', 'photo_raw_rel_path'):
            rel = r[col] or ''
            if rel:
                p = _safe_resolve_vod_file(rel)
                if p and os.path.isfile(p):
                    cand_files += 1
                    try:
                        cand_bytes += os.path.getsize(p)
                    except OSError:
                        pass

    return {
        'disk_files': disk_files,
        'disk_bytes': disk_bytes,
        'disk_mb': round(disk_bytes / (1024 * 1024), 2),
        'candidate_rows': len(cand_rows),
        'candidate_files': cand_files,
        'candidate_bytes': cand_bytes,
        'candidate_mb': round(cand_bytes / (1024 * 1024), 2),
        'cutoff_date': cutoff_str[:10],
    }


def cleanup_old_telegram_vod_files(
    conn,
    *,
    months: int = 6,
    include_not_found: bool = True,
    include_approved: bool = False,
    dry_run: bool = False,
) -> dict:
    """
    Регламентне очищення застарілих тимчасових файлів черги Telegram-бота (static/uploads/telegram_vod/)
    для відхилених/дублюючих (і за опцією старих not_found / вже перенесених approved) подань.
    Самі записи в БД та історія diff зберігаються, очищаються лише тимчасові файли на диску.
    """
    ensure_telegram_vod_schema(conn)
    months = max(1, min(60, int(months or 6)))
    cutoff_str = (datetime.now() - timedelta(days=months * 30)).strftime('%Y-%m-%d %H:%M:%S')
    where_status = _cleanup_target_where_sql(
        include_not_found=include_not_found,
        include_approved=include_approved,
    )

    cand_rows = conn.execute(
        f"""
        SELECT sub.id, sub.pdf_rel_path, sub.photo_rel_path, sub.photo_raw_rel_path
        FROM telegram_vod_submissions sub
        WHERE COALESCE(NULLIF(sub.reviewed_at, ''), sub.created_at) < ?
          AND (COALESCE(sub.pdf_rel_path, '') != ''
               OR COALESCE(sub.photo_rel_path, '') != ''
               OR COALESCE(sub.photo_raw_rel_path, '') != '')
          AND {where_status}
        """,
        (cutoff_str,),
    ).fetchall()

    if not cand_rows:
        return {'cleaned_rows': 0, 'deleted_files': 0, 'freed_bytes': 0, 'freed_mb': 0.0, 'cutoff_date': cutoff_str[:10]}

    cand_ids = {int(r['id']) for r in cand_rows}

    # Захищені шляхи: усе, що згадується в attachments або в інших записах черги (особливо pending)
    protected_rels: set[str] = set()
    for a in conn.execute("SELECT file_path FROM attachments WHERE COALESCE(file_path, '') != ''").fetchall():
        protected_rels.add(str(a['file_path']).replace('\\', '/').lstrip('/'))

    for other in conn.execute(
        "SELECT id, pdf_rel_path, photo_rel_path, photo_raw_rel_path FROM telegram_vod_submissions"
    ).fetchall():
        if int(other['id']) in cand_ids:
            continue
        for col in ('pdf_rel_path', 'photo_rel_path', 'photo_raw_rel_path'):
            if other[col]:
                protected_rels.add(str(other[col]).replace('\\', '/').lstrip('/'))

    deleted_files = 0
    freed_bytes = 0
    cleaned_rows = 0
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    for r in cand_rows:
        rid = int(r['id'])
        if dry_run:
            cleaned_rows += 1
        for col in ('pdf_rel_path', 'photo_rel_path', 'photo_raw_rel_path'):
            rel = (r[col] or '').replace('\\', '/').lstrip('/')
            if not rel:
                continue
            if rel in protected_rels:
                continue
            abs_p = _safe_resolve_vod_file(rel)
            if not abs_p:
                continue
            if os.path.isfile(abs_p):
                try:
                    sz = os.path.getsize(abs_p)
                except OSError:
                    sz = 0
                if not dry_run:
                    try:
                        os.remove(abs_p)
                        deleted_files += 1
                        freed_bytes += sz
                    except OSError as e:
                        logger.warning(f"VOD cleanup remove warning ({abs_p}): {e}")
                else:
                    deleted_files += 1
                    freed_bytes += sz
        if not dry_run:
            conn.execute(
                """
                UPDATE telegram_vod_submissions
                SET pdf_rel_path = '',
                    photo_rel_path = '',
                    photo_raw_rel_path = '',
                    files_cleaned_at = ?
                WHERE id = ?
                """,
                (now_str, rid),
            )
            cleaned_rows += 1

    if not dry_run:
        today_str = date.today().strftime('%Y-%m-%d')
        freed_mb = round(freed_bytes / (1024 * 1024), 2)
        summary_txt = f"{today_str} — очищено записів: {cleaned_rows}, файлів: {deleted_files} ({freed_mb} МБ)"
        conn.execute(
            """
            UPDATE telegram_bot_settings
            SET last_cleanup_date = ?, last_cleanup_summary = ?, updated_at = datetime('now', 'localtime')
            WHERE id = 1
            """,
            (today_str, summary_txt[:300]),
        )
        if deleted_files > 0 or cleaned_rows > 0:
            _record_event(
                conn, 'vod_files_cleanup',
                f"Автоочищення тимчасових файлів черги: видалено {deleted_files} файлів ({freed_mb} МБ)",
                severity='info',
                details=summary_txt,
            )
        conn.commit()

    return {
        'cleaned_rows': cleaned_rows,
        'deleted_files': deleted_files,
        'freed_bytes': freed_bytes,
        'freed_mb': round(freed_bytes / (1024 * 1024), 2),
        'cutoff_date': cutoff_str[:10],
    }


def _run_scheduled_vod_cleanup_if_due(conn, settings: dict) -> None:
    if not int(settings.get('auto_cleanup_enabled') if settings.get('auto_cleanup_enabled') is not None else 1):
        return
    today_str = date.today().strftime('%Y-%m-%d')
    if (settings.get('last_cleanup_date') or '') == today_str:
        return
    months = int(settings.get('auto_cleanup_months') or 6)
    inc_nf = bool(int(settings.get('auto_cleanup_include_not_found') if settings.get('auto_cleanup_include_not_found') is not None else 1))
    inc_app = bool(int(settings.get('auto_cleanup_include_approved') or 0))
    cleanup_old_telegram_vod_files(
        conn,
        months=months,
        include_not_found=inc_nf,
        include_approved=inc_app,
        dry_run=False,
    )


def _run_due_filter_reminder_rules(conn, token: str, now: datetime | None = None) -> int:
    """Перевіряє всі активні правила нагадувань за фільтрами (період днів місяця + частота + ескалація)."""
    now_dt = now or datetime.now()
    today_str = now_dt.strftime('%Y-%m-%d')
    now_min = now_dt.hour * 60 + now_dt.minute

    rows = conn.execute(
        "SELECT * FROM telegram_auto_reminder_rules WHERE is_enabled = 1 ORDER BY id ASC"
    ).fetchall()
    triggered = 0
    for r in rows:
        rule = dict(r)
        if not _is_day_in_window(now_dt.day, rule.get('day_from', 15), rule.get('day_to', 25)):
            continue
        slots = _compute_daily_reminder_slots(
            rule.get('times_per_day', 1),
            rule.get('start_hour', 9),
            rule.get('end_hour', 20),
        )
        due_idx = -1
        for idx, sm in enumerate(slots):
            if now_min >= sm:
                due_idx = idx

        raw_prev_slot = rule.get('last_sent_slot_idx')
        prev_slot = int(raw_prev_slot) if ((rule.get('last_sent_date') or '') == today_str and raw_prev_slot is not None) else -1
        if due_idx >= 0 and due_idx > prev_slot:
            try:
                _execute_auto_rule_dispatch(conn, token, rule, now=now_dt, slot_idx=due_idx, created_by='auto_rule')
                triggered += 1
            except Exception as e:
                logger.error(f"Auto reminder rule #{rule['id']} error: {e}", exc_info=True)

        # Перевірка ескалації наприкінці вікна (раз на день після start_hour)
        if (
            int(rule.get('escalation_enabled') or 0) == 1
            and (rule.get('escalation_last_date') or '') != today_str
            and now_dt.hour >= int(rule.get('start_hour') if rule.get('start_hour') is not None else 9)
            and _is_day_in_escalation_window(
                now_dt.day,
                rule.get('day_from', 15),
                rule.get('day_to', 25),
                rule.get('escalation_last_days', 3),
                ref_date=now_dt.date(),
            )
        ):
            try:
                _execute_rule_escalation(conn, token, rule, now=now_dt, force=False)
            except Exception as e:
                logger.error(f"Auto reminder escalation rule #{rule['id']} error: {e}", exc_info=True)

    return triggered


def _run_scheduled_auto_reminder_if_due(conn, token: str, settings: dict) -> None:
    """Перевіряє та виконує: 1) загальне щомісячне нагадування; 2) правила за фільтрами; 3) щоденне очищення файлів."""
    _run_due_filter_reminder_rules(conn, token)
    try:
        _run_scheduled_vod_cleanup_if_due(conn, settings)
    except Exception as e:
        logger.error(f"Telegram VOD scheduled cleanup error: {e}", exc_info=True)

    if not int(settings.get('auto_remind_enabled') or 0):
        return
    now = datetime.now()
    ym_now = now.strftime('%Y-%m')
    if (settings.get('last_auto_remind_ym') or '') == ym_now:
        return
    target_day = max(1, min(28, int(settings.get('auto_remind_day') or 5)))
    target_hour = max(0, min(23, int(settings.get('auto_remind_hour') or 10)))
    if now.day < target_day or (now.day == target_day and now.hour < target_hour):
        return

    try:
        flt = json.loads(settings.get('auto_remind_filter_json') or '{}')
    except Exception:
        flt = {}
    if not isinstance(flt, dict):
        flt = {}
    flt['vod_month'] = ym_now
    if 'only_base_license' not in flt and not flt.get('license_ids'):
        flt['only_base_license'] = True
    if not flt.get('vod_month_state') or flt.get('vod_month_state') == 'any':
        flt['vod_month_state'] = 'missing_this_month'

    # Фіксуємо місяць до відправки, щоб уникнути дублю при збоях мережі
    conn.execute(
        "UPDATE telegram_bot_settings SET last_auto_remind_ym = ?, updated_at = datetime('now', 'localtime') WHERE id = 1",
        (ym_now,),
    )
    conn.commit()

    template = (settings.get('reminder_default_text') or DEFAULT_REMINDER_TEXT).strip()
    execute_broadcast(
        conn,
        token=token,
        title=f"Автоматичне нагадування за {_format_ym_ua(ym_now)}",
        message_template=template,
        raw_filters=flt,
        created_by='auto_scheduler',
        is_auto=True,
    )


def _worker_loop() -> None:
    logger.info("Telegram ВОД: фоновий потік запущено")
    last_auto_check = 0.0

    while not _WORKER_STOP_EVENT.is_set():
        if is_test_mode_active():
            with _WORKER_LOCK:
                _WORKER_STATE['running'] = False
            break
        conn = None
        try:
            conn = get_db()
            conn.execute("PRAGMA busy_timeout = 5000")
            settings = _get_bot_settings(conn)
            token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
            enabled = bool(int(settings.get('enabled') or 0))
            offset = int(settings.get('last_update_id') or 0) + 1

            if not token or not enabled:
                with _WORKER_LOCK:
                    _WORKER_STATE['running'] = False
                conn.close()
                break

            with _WORKER_LOCK:
                _WORKER_STATE['bot_username'] = settings.get('bot_username') or ''
                _WORKER_STATE['bot_first_name'] = settings.get('bot_first_name') or ''
            conn.close()
            conn = None

            ok, updates, err = _tg_call(
                token,
                'getUpdates',
                {'offset': offset, 'timeout': 20, 'allowed_updates': ['message', 'callback_query']},
                timeout=26.0,
            )

            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            conn = get_db()
            conn.execute("PRAGMA busy_timeout = 5000")

            if not ok:
                with _WORKER_LOCK:
                    _WORKER_STATE['last_poll_at'] = now_str
                    _WORKER_STATE['last_error'] = err
                conn.execute(
                    "UPDATE telegram_bot_settings SET last_poll_at = ?, last_error = ? WHERE id = 1",
                    (now_str, err[:500]),
                )
                conn.commit()
                conn.close()
                conn = None
                _WORKER_STOP_EVENT.wait(5.0)
                continue

            max_uid = offset - 1
            if isinstance(updates, list) and updates:
                settings_fresh = _get_bot_settings(conn)
                for upd in updates:
                    uid = int(upd.get('update_id') or 0)
                    if uid > max_uid:
                        max_uid = uid
                    try:
                        _process_telegram_update(conn, token, settings_fresh, upd)
                    except Exception as e:
                        logger.error(f"Telegram update error (uid={uid}): {e}", exc_info=True)
                        conn.rollback()

            conn.execute(
                """
                UPDATE telegram_bot_settings
                SET last_update_id = ?, last_poll_at = ?, last_error = ''
                WHERE id = 1
                """,
                (max_uid, now_str),
            )
            conn.commit()

            with _WORKER_LOCK:
                _WORKER_STATE['last_poll_at'] = now_str
                _WORKER_STATE['last_error'] = ''
                _WORKER_STATE['poll_cycles'] += 1

            # Періодична перевірка авто-нагадувань і правил за розкладом (раз на хвилину)
            if time.monotonic() - last_auto_check >= 60:
                last_auto_check = time.monotonic()
                try:
                    settings_fresh = _get_bot_settings(conn)
                    _run_scheduled_auto_reminder_if_due(conn, token, settings_fresh)
                except Exception as e:
                    logger.error(f"Telegram auto-reminder error: {e}", exc_info=True)

            conn.close()
            conn = None

        except Exception as e:
            logger.error(f"Telegram ВОД worker loop error: {e}", exc_info=True)
            with _WORKER_LOCK:
                _WORKER_STATE['last_error'] = str(e)
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass
            _WORKER_STOP_EVENT.wait(5.0)

    with _WORKER_LOCK:
        _WORKER_STATE['running'] = False
    logger.info("Telegram ВОД: фоновий потік зупинено")


def start_bot_worker() -> tuple[bool, str]:
    global _WORKER_THREAD
    if is_test_mode_active():
        return False, 'Увімкнено тестовий режим: запуск Telegram-бота тимчасово заблоковано'
    with _WORKER_LOCK:
        if _WORKER_THREAD and _WORKER_THREAD.is_alive():
            return True, 'Бот уже працює'
        conn = get_db()
        try:
            settings = _get_bot_settings(conn)
            token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
            if not token:
                return False, 'Не вказано токен Telegram-бота'
            conn.execute("UPDATE telegram_bot_settings SET enabled = 1 WHERE id = 1")
            conn.commit()
        finally:
            conn.close()

        _WORKER_STOP_EVENT.clear()
        _WORKER_STATE['running'] = True
        _WORKER_STATE['started_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        _WORKER_STATE['last_error'] = ''
        _WORKER_THREAD = threading.Thread(
            target=_worker_loop, name='telegram-vod-poller', daemon=True
        )
        _WORKER_THREAD.start()
        return True, 'Telegram-бота запущено'


def stop_bot_worker(disable_in_db: bool = True) -> tuple[bool, str]:
    global _WORKER_THREAD
    with _WORKER_LOCK:
        _WORKER_STOP_EVENT.set()
        th = _WORKER_THREAD
        _WORKER_STATE['running'] = False
    if disable_in_db:
        conn = get_db()
        try:
            ensure_telegram_vod_schema(conn)
            conn.execute("UPDATE telegram_bot_settings SET enabled = 0 WHERE id = 1")
            conn.commit()
        finally:
            conn.close()
    if th and th.is_alive():
        th.join(timeout=2.0)
    return True, 'Telegram-бота зупинено'


def ensure_bot_started_if_enabled() -> None:
    """Викликається при старті застосунку: автозапускає воркер, якщо enabled=1 і є токен."""
    if is_test_mode_active():
        return
    ensure_telegram_vod_initialized()
    with _WORKER_LOCK:
        if _WORKER_THREAD and _WORKER_THREAD.is_alive():
            return
    conn = None
    try:
        conn = get_db()
        settings = _get_bot_settings(conn)
        token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
        if int(settings.get('enabled') or 0) == 1 and token:
            conn.close()
            conn = None
            start_bot_worker()
    except Exception as e:
        logger.warning(f"Не вдалося автозапустити Telegram-бота ВОД: {e}")
    finally:
        if conn is not None:
            conn.close()


def get_worker_status(conn=None) -> dict:
    with _WORKER_LOCK:
        alive = bool(_WORKER_THREAD and _WORKER_THREAD.is_alive())
        state = dict(_WORKER_STATE)
        state['running'] = alive
    state['test_mode_active'] = is_test_mode_active()
    if conn is not None:
        s = _get_bot_settings(conn)
        state['enabled_in_db'] = bool(int(s.get('enabled') or 0))
        state['has_token'] = bool((s.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip())
        state['bot_username'] = s.get('bot_username') or state.get('bot_username') or ''
        state['bot_first_name'] = s.get('bot_first_name') or state.get('bot_first_name') or ''
        if not state.get('last_poll_at'):
            state['last_poll_at'] = s.get('last_poll_at')
        if not state.get('last_error'):
            state['last_error'] = s.get('last_error') or ''
    return state


# ═══════════════════════════════════════════════════════════════════════════════
# 6. Розсилка повідомлень за фільтрами
# ═══════════════════════════════════════════════════════════════════════════════

def execute_broadcast(conn, *, token: str, title: str, message_template: str,
                      raw_filters, created_by: str, is_auto: bool = False,
                      skip_if_empty: bool = False) -> dict:
    res = query_filtered_students(conn, raw_filters)
    f = res['filters']
    vod_month = f['vod_month']

    if skip_if_empty and res['linked_count'] == 0:
        return {
            'broadcast_id': None,
            'total_matched': res['total_matched'],
            'linked_count': 0,
            'unlinked_count': res['unlinked_count'],
            'sent_chats': 0,
            'failed_chats': 0,
            'skipped_empty': True,
        }

    cur = conn.execute(
        """
        INSERT INTO telegram_broadcasts (
            title, message_text, filter_json, total_matched,
            total_linked_students, total_unlinked_students,
            sent_chats_count, failed_chats_count, is_auto, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, 0, 0, ?, ?)
        """,
        (
            title[:200],
            message_template,
            json.dumps(f, ensure_ascii=False),
            res['total_matched'],
            res['linked_count'],
            res['unlinked_count'],
            1 if is_auto else 0,
            created_by,
        ),
    )
    bcast_id = int(cur.lastrowid)

    sent_chats = 0
    failed_chats = 0
    logs_to_insert = []

    for st in res['students']:
        sid = st['id']
        links = st['links']
        if not links:
            logs_to_insert.append((bcast_id, sid, '', 'no_telegram', 'Студент не прив’язаний до Telegram'))
            continue

        msg_body = _render_broadcast_message(message_template, st, vod_month)
        for lnk in links:
            cid = str(lnk['chat_id'])
            final_text = msg_body
            if lnk.get('is_representative') and st['full_name'] not in msg_body:
                final_text = f"👥 Для студента: {st['full_name']} ({st['group_name'] or '—'})\n\n{msg_body}"
            ok, err = _tg_send_message(token, cid, final_text)
            if ok:
                sent_chats += 1
                logs_to_insert.append((bcast_id, sid, cid, 'sent', ''))
            else:
                failed_chats += 1
                logs_to_insert.append((bcast_id, sid, cid, 'failed', err[:300]))

    if logs_to_insert:
        conn.executemany(
            """
            INSERT INTO telegram_broadcast_logs (broadcast_id, student_id, chat_id, status, error_text)
            VALUES (?, ?, ?, ?, ?)
            """,
            logs_to_insert,
        )

    conn.execute(
        """
        UPDATE telegram_broadcasts
        SET sent_chats_count = ?, failed_chats_count = ?
        WHERE id = ?
        """,
        (sent_chats, failed_chats, bcast_id),
    )
    _record_event(
        conn, 'broadcast_sent',
        f"Розсилка «{title}»: надіслано {sent_chats}, помилок {failed_chats}",
        severity='warning' if failed_chats else 'info',
        details=f"Збіглося студентів: {res['total_matched']}, з Telegram: {res['linked_count']}, без Telegram: {res['unlinked_count']}",
    )
    conn.commit()
    return {
        'broadcast_id': bcast_id,
        'total_matched': res['total_matched'],
        'linked_count': res['linked_count'],
        'unlinked_count': res['unlinked_count'],
        'sent_chats': sent_chats,
        'failed_chats': failed_chats,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 7. Веб-маршрути керування Telegram-ботом (тільки з правом manage_military_telegram)
# ═══════════════════════════════════════════════════════════════════════════════

def _load_filter_catalogs(conn) -> dict:
    from routes.military_reports import CATEGORY_LABELS, NOTE_18_LABELS, STUDY_FORM_OPTIONS

    faculties = [
        dict(r) for r in conn.execute(
            "SELECT id, name_ua, short_name FROM faculties WHERE COALESCE(is_active, 1) = 1 ORDER BY sort_order, name_ua"
        ).fetchall()
    ]
    groups = [
        dict(r) for r in conn.execute(
            """
            SELECT id, name, start_year, course, study_form, degree_level, faculty_id
            FROM groups
            WHERE COALESCE(archived, 0) = 0
            ORDER BY course, name COLLATE UKRAINIAN
            """
        ).fetchall()
    ]
    degree_levels = [
        r['name_ua'] for r in conn.execute(
            "SELECT name_ua FROM degree_levels WHERE COALESCE(is_active, 1) = 1 ORDER BY id"
        ).fetchall()
    ]
    from routes.military_reports import _get_military_visible_licenses, ensure_military_license_schema
    ensure_military_license_schema(conn)
    _, mil_lic_label = _get_military_visible_licenses(conn)
    licenses = [
        dict(r) for r in conn.execute(
            "SELECT id, short_name_ua, name_ua, is_base, COALESCE(show_in_military, 0) AS show_in_military "
            "FROM institution_licenses WHERE COALESCE(is_active, 1) = 1 ORDER BY is_base DESC, id"
        ).fetchall()
    ]
    status_badges = [
        r['status_badge'] for r in conn.execute(
            "SELECT DISTINCT status_badge FROM military WHERE COALESCE(status_badge, '') != '' ORDER BY status_badge"
        ).fetchall()
    ]
    courses = sorted({int(g['course']) for g in groups if g.get('course') is not None})
    return {
        'faculties': faculties,
        'groups': groups,
        'degree_levels': degree_levels,
        'licenses': licenses,
        'military_license_label': mil_lic_label,
        'study_forms': STUDY_FORM_OPTIONS,
        'courses': courses,
        'status_badges': status_badges,
        'category_labels': CATEGORY_LABELS,
        'note_18_labels': NOTE_18_LABELS,
    }


def _extract_filters_from_request(src) -> dict:
    """Читає параметри фільтра з request.form або request.json."""
    if request.is_json:
        return _parse_filter_dict(request.get_json(silent=True) or {})
    return _parse_filter_dict({
        'search': src.get('search', ''),
        'faculty_ids': src.getlist('faculty_ids') if hasattr(src, 'getlist') else src.get('faculty_ids', []),
        'group_ids': src.getlist('group_ids') if hasattr(src, 'getlist') else src.get('group_ids', []),
        'courses': src.getlist('courses') if hasattr(src, 'getlist') else src.get('courses', []),
        'study_forms': src.getlist('study_forms') if hasattr(src, 'getlist') else src.get('study_forms', []),
        'degree_levels': src.getlist('degree_levels') if hasattr(src, 'getlist') else src.get('degree_levels', []),
        'license_ids': src.getlist('license_ids') if hasattr(src, 'getlist') else src.get('license_ids', []),
        'only_base_license': src.get('only_base_license'),
        'genders': src.getlist('genders') if hasattr(src, 'getlist') else src.get('genders', []),
        'age_mode': src.get('age_mode', 'any'),
        'age_min': src.get('age_min'),
        'age_max': src.get('age_max'),
        'age_as_of': src.get('age_as_of', ''),
        'military_statuses': src.getlist('military_statuses') if hasattr(src, 'getlist') else src.get('military_statuses', []),
        'military_categories': src.getlist('military_categories') if hasattr(src, 'getlist') else src.get('military_categories', []),
        'exclude_women_no_mil': src.get('exclude_women_no_mil'),
        'vod_month': src.get('vod_month', ''),
        'vod_month_state': src.get('vod_month_state', 'any'),
        'formed_older_than_days': src.get('formed_older_than_days'),
        'deferral_expiring_days': src.get('deferral_expiring_days'),
        'valid_until_expiring_days': src.get('valid_until_expiring_days'),
        'has_police_reason': src.get('has_police_reason', 'any'),
        'note_18': src.getlist('note_18') if hasattr(src, 'getlist') else src.get('note_18', []),
        'exclude_note_18': src.get('exclude_note_18'),
        'tg_link_state': src.get('tg_link_state', 'any'),
        'include_frozen': src.get('include_frozen'),
        'include_student_ids': src.getlist('include_student_ids') if hasattr(src, 'getlist') else src.get('include_student_ids', []),
        'exclude_student_ids': src.getlist('exclude_student_ids') if hasattr(src, 'getlist') else src.get('exclude_student_ids', []),
    })


@military_telegram_bp.route('/admin/military/telegram')
@permission_required('manage_military_telegram')
def telegram_dashboard():
    active_tab = (request.args.get('tab') or 'queue').strip()
    if active_tab not in ('queue', 'groups', 'broadcast', 'links', 'settings'):
        active_tab = 'queue'
    queue_filter = (request.args.get('qstatus') or 'attention').strip()

    conn = get_db()
    ensure_telegram_vod_schema(conn)
    settings = _get_bot_settings(conn)
    worker_status = get_worker_status(conn)
    catalogs = _load_filter_catalogs(conn)

    # Завантажуємо поточні військові записи для актуального підрахунку ризиків у черзі
    cur_mil_rows = conn.execute("SELECT * FROM military").fetchall()
    cur_mil_by_sid = {int(m['student_id']): dict(m) for m in cur_mil_rows if m['student_id'] is not None}

    # Лічильники черги
    counts = {
        'pending': conn.execute("SELECT COUNT(*) FROM telegram_vod_submissions WHERE status='pending'").fetchone()[0],
        'not_found': conn.execute("SELECT COUNT(*) FROM telegram_vod_submissions WHERE status='not_found'").fetchone()[0],
        'duplicates': conn.execute(
            "SELECT COUNT(*) FROM telegram_vod_submissions WHERE is_monthly_duplicate=1 AND status IN ('pending', 'approved')"
        ).fetchone()[0],
        'approved': conn.execute("SELECT COUNT(*) FROM telegram_vod_submissions WHERE status='approved'").fetchone()[0],
        'rejected': conn.execute("SELECT COUNT(*) FROM telegram_vod_submissions WHERE status='rejected'").fetchone()[0],
        'risks': 0,
    }
    counts['attention'] = counts['pending'] + counts['not_found']

    pending_for_risk_cnt = conn.execute(
        "SELECT id, student_id, status, parsed_json, risk_flags_json, risk_level FROM telegram_vod_submissions WHERE status='pending'"
    ).fetchall()
    risk_pending_ids = set()
    for pr in pending_for_risk_cnt:
        pd = _hydrate_submission_risks(conn, dict(pr), cur_mil_by_sid=cur_mil_by_sid)
        if pd['risk_level'] in ('critical', 'warning'):
            risk_pending_ids.add(int(pd['id']))
    counts['risks'] = len(risk_pending_ids)

    where_sub = "1=1"
    if queue_filter == 'attention':
        where_sub = "sub.status IN ('pending', 'not_found')"
    elif queue_filter == 'duplicates':
        where_sub = "sub.is_monthly_duplicate = 1"
    elif queue_filter == 'risks':
        where_sub = "sub.status = 'pending'"
    elif queue_filter in ('pending', 'not_found', 'approved', 'rejected'):
        where_sub = f"sub.status = '{queue_filter}'"

    submissions_raw = conn.execute(
        f"""
        SELECT sub.*,
               s.last_name_UA, s.first_name_UA, s.middle_name_UA, s.birth_date AS db_birth_date,
               g.name AS group_name
        FROM telegram_vod_submissions sub
        LEFT JOIN students s ON s.id = sub.student_id
        LEFT JOIN groups g ON g.id = s.group_id
        WHERE {where_sub}
        ORDER BY
            CASE sub.status WHEN 'not_found' THEN 0 WHEN 'pending' THEN 1 ELSE 2 END,
            sub.is_monthly_duplicate DESC,
            sub.id DESC
        LIMIT 250
        """
    ).fetchall()

    submissions = []
    for r in submissions_raw:
        d = dict(r)
        if d.get('student_id'):
            d['student_full_name'] = ' '.join(
                x for x in (d.get('last_name_UA'), d.get('first_name_UA'), d.get('middle_name_UA') or '') if x
            ).strip()
        else:
            d['student_full_name'] = ''
        try:
            d['mil_diff'] = json.loads(d.get('military_diff_json') or '[]')
        except Exception:
            d['mil_diff'] = []
        try:
            d['stu_diff'] = json.loads(d.get('student_diff_json') or '[]')
        except Exception:
            d['stu_diff'] = []
        _hydrate_submission_risks(conn, d, cur_mil_by_sid=cur_mil_by_sid)
        if queue_filter == 'risks' and d['risk_level'] not in ('critical', 'warning'):
            continue
        submissions.append(d)

    if queue_filter in ('attention', 'pending', 'risks'):
        risk_rank = {'critical': 0, 'warning': 1, 'none': 2}
        submissions.sort(key=lambda x: (
            0 if x['status'] == 'not_found' else (1 if x['status'] == 'pending' else 2),
            risk_rank.get(x.get('risk_level') or 'none', 2),
            0 if x.get('is_monthly_duplicate') else 1,
            -int(x['id']),
        ))

    # Прив'язки студентів
    links = [
        dict(r) for r in conn.execute(
            """
            SELECT l.*, s.last_name_UA, s.first_name_UA, s.middle_name_UA, s.phone AS db_phone,
                   s.group_id AS group_id, g.name AS group_name
            FROM telegram_student_links l
            JOIN students s ON s.id = l.student_id
            LEFT JOIN groups g ON g.id = s.group_id
            WHERE l.is_active = 1
            ORDER BY l.updated_at DESC, l.id DESC
            """
        ).fetchall()
    ]
    linked_by_group: dict[int, list[dict]] = {}
    for l in links:
        l['student_full_name'] = ' '.join(
            x for x in (l['last_name_UA'], l['first_name_UA'], l['middle_name_UA'] or '') if x
        ).strip()
        if l.get('group_id') is not None:
            linked_by_group.setdefault(int(l['group_id']), []).append({
                'student_full_name': l['student_full_name'],
                'chat_id': l['chat_id'],
                'tg_username': l.get('tg_username') or '',
                'is_representative': bool(l.get('is_representative')),
            })

    # Контакти старост та кураторів груп
    group_contacts = [
        dict(r) for r in conn.execute(
            """
            SELECT gc.*, g.name AS group_name, g.course AS group_course, g.study_form AS group_study_form
            FROM telegram_group_contacts gc
            LEFT JOIN groups g ON g.id = gc.group_id
            ORDER BY g.course, g.name COLLATE UKRAINIAN
            """
        ).fetchall()
    ]
    group_contacts_by_gid = {int(c['group_id']): c for c in group_contacts}

    # Активні запрошення
    invites = [
        dict(r) for r in conn.execute(
            """
            SELECT i.*,
                   s.last_name_UA, s.first_name_UA, s.middle_name_UA,
                   g.name AS group_name
            FROM telegram_invite_tokens i
            LEFT JOIN students s ON s.id = i.student_id
            LEFT JOIN groups g ON g.id = i.group_id
            WHERE i.is_active = 1
            ORDER BY i.id DESC
            LIMIT 100
            """
        ).fetchall()
    ]
    for inv in invites:
        if inv.get('student_id'):
            inv['target_label'] = ' '.join(
                x for x in (inv.get('last_name_UA'), inv.get('first_name_UA'), inv.get('middle_name_UA') or '') if x
            ).strip()
        elif inv.get('group_id'):
            inv['target_label'] = f"Група: {inv.get('group_name') or inv['group_id']}"
        else:
            inv['target_label'] = '—'

    # Активні студенти для ручного вибору (прив'язка, інвайт або rematch)
    active_students = [
        {
            'id': r['id'],
            'full_name': ' '.join(x for x in (r['last_name_UA'], r['first_name_UA'], r['middle_name_UA'] or '') if x).strip(),
            'birth_date': r['birth_date'] or '',
            'group_name': r['group_name'] or '',
            'phone': r['phone'] or '',
        }
        for r in conn.execute(
            f"""
            SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA, s.birth_date, s.phone, g.name AS group_name
            FROM students s LEFT JOIN groups g ON g.id = s.group_id
            WHERE {_active_students_sql_where('s', include_frozen=True)}
            ORDER BY s.last_name_UA COLLATE UKRAINIAN, s.first_name_UA COLLATE UKRAINIAN
            """
        ).fetchall()
    ]

    presets = [
        dict(r) for r in conn.execute(
            "SELECT * FROM telegram_broadcast_presets ORDER BY id ASC"
        ).fetchall()
    ]
    preset_name_by_id = {int(p['id']): p['name'] for p in presets}
    today_str = date.today().strftime('%Y-%m-%d')
    auto_rules = []
    for r in conn.execute(
        "SELECT * FROM telegram_auto_reminder_rules ORDER BY is_enabled DESC, id ASC"
    ).fetchall():
        rd = dict(r)
        rd['slots_preview'] = _format_daily_slots(
            rd.get('times_per_day', 1), rd.get('start_hour', 9), rd.get('end_hour', 20)
        )
        rd['effective_sent_today'] = int(rd.get('sent_today_count') or 0) if (rd.get('last_sent_date') or '') == today_str else 0
        rd['preset_name'] = preset_name_by_id.get(int(rd['preset_id'])) if rd.get('preset_id') else ''
        auto_rules.append(rd)

    broadcasts = [
        dict(r) for r in conn.execute(
            "SELECT * FROM telegram_broadcasts ORDER BY id DESC LIMIT 40"
        ).fetchall()
    ]
    events = [
        dict(r) for r in conn.execute(
            "SELECT * FROM telegram_events ORDER BY id DESC LIMIT 60"
        ).fetchall()
    ]
    unread_events = sum(1 for e in events if not e['is_read'])

    from routes.military_reports import _recent_months_list
    month_options = [{'ym': ym, 'label': _format_ym_ua(ym)} for ym in _recent_months_list(12)]
    today_ym = date.today().strftime('%Y-%m')

    # Зведення по групах (для вкладки groups)
    grp_month = (request.args.get('gmonth') or today_ym).strip()
    grp_faculty_id = request.args.get('gfaculty', type=int)
    grp_course = request.args.get('gcourse', type=int)
    grp_group_id = request.args.get('ggroup', type=int)
    grp_only_missing = bool(request.args.get('gmissing') == '1')
    grp_all_licenses = bool(request.args.get('gall_lic') == '1')
    grp_age_mode = (request.args.get('gage') or 'any').strip()
    group_summary = None
    if active_tab == 'groups':
        g_flt = {
            'vod_month': grp_month,
            'faculty_ids': [grp_faculty_id] if grp_faculty_id else [],
            'courses': [grp_course] if grp_course else [],
            'group_ids': [grp_group_id] if grp_group_id else [],
            'age_mode': grp_age_mode if grp_age_mode in ('any', '25_plus', 'under_25') else 'any',
            'only_base_license': not grp_all_licenses,
            'exclude_women_no_mil': True,
            'exclude_note_18': True,
        }
        group_summary = build_group_vod_compliance_summary(
            conn, g_flt, vod_month=grp_month, only_with_missing=grp_only_missing
        )

    vod_storage_stats = None
    if active_tab == 'settings':
        vod_storage_stats = get_telegram_vod_storage_stats(
            conn,
            months=int(settings.get('auto_cleanup_months') or 6),
            include_not_found=bool(int(settings.get('auto_cleanup_include_not_found') if settings.get('auto_cleanup_include_not_found') is not None else 1)),
            include_approved=bool(int(settings.get('auto_cleanup_include_approved') or 0)),
        )

    conn.close()

    return render_template(
        'military_reports/telegram_bot.html',
        active_tab=active_tab,
        queue_filter=queue_filter,
        settings=settings,
        masked_token=_mask_token(settings.get('bot_token') or ''),
        worker_status=worker_status,
        catalogs=catalogs,
        military_license_label=catalogs.get('military_license_label') or '—',
        counts=counts,
        submissions=submissions,
        links=links,
        linked_by_group=linked_by_group,
        group_contacts=group_contacts,
        group_contacts_by_gid=group_contacts_by_gid,
        invites=invites,
        active_students=active_students,
        presets=presets,
        auto_rules=auto_rules,
        broadcasts=broadcasts,
        events=events,
        unread_events=unread_events,
        month_options=month_options,
        today_ym=today_ym,
        group_summary=group_summary,
        grp_filters={
            'gmonth': grp_month,
            'gfaculty': grp_faculty_id,
            'gcourse': grp_course,
            'ggroup': grp_group_id,
            'gmissing': grp_only_missing,
            'gall_lic': grp_all_licenses,
            'gage': grp_age_mode,
        },
        vod_storage_stats=vod_storage_stats,
    )


@military_telegram_bp.route('/admin/military/telegram/filter-preview', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_filter_preview():
    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        filters = _extract_filters_from_request(request.form)
        res = query_filtered_students(conn, filters)
        return jsonify({
            'ok': True,
            'total_matched': res['total_matched'],
            'linked_count': res['linked_count'],
            'unlinked_count': res['unlinked_count'],
            'unique_chats_count': res['unique_chats_count'],
            'students': [
                {
                    'id': s['id'],
                    'full_name': s['full_name'],
                    'group_name': s['group_name'],
                    'faculty_name': s['faculty_name'],
                    'age': s['age'],
                    'status_badge': s['status_badge'],
                    'formed_at': s['formed_at'],
                    'deferral_until': s['deferral_until'],
                    'phone': s['phone'] or s['phone_backup'],
                    'is_linked': s['is_linked'],
                    'has_representative': s['has_representative'],
                    'has_att_in_month': s['has_att_in_month'],
                    'has_any_in_month': s['has_any_in_month'],
                }
                for s in res['students'][:300]
            ],
            'truncated': res['total_matched'] > 300,
        })
    finally:
        conn.close()


@military_telegram_bp.route('/admin/military/telegram/presets/save', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_preset_save():
    name = (request.form.get('preset_name') or '').strip()
    desc = (request.form.get('preset_description') or '').strip()
    msg_tpl = (request.form.get('message_template') or '').strip()
    raw_flt = request.form.get('filter_json') or '{}'
    if not name:
        flash('Вкажіть назву збереженого фільтра (пресета).', 'warning')
        return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))

    filters = _parse_filter_dict(raw_flt)
    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        conn.execute(
            """
            INSERT INTO telegram_broadcast_presets (name, description, filter_json, message_template, created_by)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET
                description = excluded.description,
                filter_json = excluded.filter_json,
                message_template = excluded.message_template,
                updated_at = datetime('now', 'localtime')
            """,
            (name, desc, json.dumps(filters, ensure_ascii=False), msg_tpl, current_username()),
        )
        conn.commit()
        flash(f'Пресет фільтрів «{name}» збережено.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))


@military_telegram_bp.route('/admin/military/telegram/presets/<int:preset_id>/delete', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_preset_delete(preset_id):
    conn = get_db()
    try:
        conn.execute("DELETE FROM telegram_broadcast_presets WHERE id = ?", (preset_id,))
        conn.commit()
        flash('Пресет фільтрів видалено.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))


@military_telegram_bp.route('/admin/military/telegram/auto-rules/save', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_auto_rule_save():
    rule_id = request.form.get('rule_id', type=int)
    name = (request.form.get('rule_name') or '').strip()
    preset_id = request.form.get('preset_id', type=int)
    filter_source = (request.form.get('filter_source') or 'constructor').strip()
    raw_flt = request.form.get('filter_json') or '{}'
    message_template = (request.form.get('message_template') or '').strip() or DEFAULT_REMINDER_TEXT

    day_from = max(1, min(31, request.form.get('day_from', type=int) or 15))
    day_to = max(1, min(31, request.form.get('day_to', type=int) or 25))
    times_per_day = max(1, min(24, request.form.get('times_per_day', type=int) or 1))
    start_hour = request.form.get('start_hour', type=int)
    end_hour = request.form.get('end_hour', type=int)
    start_hour = max(0, min(23, 9 if start_hour is None else start_hour))
    end_hour = max(0, min(23, 20 if end_hour is None else end_hour))
    if end_hour < start_hour:
        end_hour = start_hour

    stop_when_submitted = 1 if request.form.get('stop_when_submitted') else 0
    is_enabled = 1 if request.form.get('is_enabled') else 0

    escalation_enabled = 1 if request.form.get('escalation_enabled') else 0
    escalation_last_days = max(1, min(15, request.form.get('escalation_last_days', type=int) or 3))
    escalation_send_to_leaders = 1 if request.form.get('escalation_send_to_leaders') else 0
    escalation_extra_raw = (request.form.get('escalation_extra_chat_ids') or '').strip()
    escalation_extra_clean = ', '.join(
        c.strip() for c in re.split(r'[\s,;]+', escalation_extra_raw) if c.strip()
    )

    if not name:
        flash('Вкажіть назву правила авто-нагадування.', 'warning')
        return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))

    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        if filter_source == 'preset' and preset_id:
            pr = conn.execute(
                "SELECT id, filter_json, message_template FROM telegram_broadcast_presets WHERE id = ?",
                (preset_id,),
            ).fetchone()
            if pr:
                filters = _parse_filter_dict(pr['filter_json'])
                if not (request.form.get('message_template') or '').strip() and pr['message_template']:
                    message_template = pr['message_template']
            else:
                filters = _parse_filter_dict(raw_flt)
                preset_id = None
        else:
            filters = _parse_filter_dict(raw_flt)
            if filter_source != 'preset':
                preset_id = None

        flt_json_clean = json.dumps(filters, ensure_ascii=False)

        if rule_id:
            conn.execute(
                """
                UPDATE telegram_auto_reminder_rules
                SET name = ?, preset_id = ?, filter_json = ?, message_template = ?,
                    day_from = ?, day_to = ?, times_per_day = ?, start_hour = ?, end_hour = ?,
                    stop_when_submitted = ?, is_enabled = ?,
                    escalation_enabled = ?, escalation_last_days = ?,
                    escalation_send_to_leaders = ?, escalation_extra_chat_ids = ?,
                    updated_at = datetime('now', 'localtime')
                WHERE id = ?
                """,
                (
                    name, preset_id, flt_json_clean, message_template,
                    day_from, day_to, times_per_day, start_hour, end_hour,
                    stop_when_submitted, is_enabled,
                    escalation_enabled, escalation_last_days,
                    escalation_send_to_leaders, escalation_extra_clean,
                    rule_id,
                ),
            )
            flash(f'Правило авто-нагадування «{name}» оновлено.', 'success')
        else:
            conn.execute(
                """
                INSERT INTO telegram_auto_reminder_rules (
                    name, preset_id, filter_json, message_template,
                    day_from, day_to, times_per_day, start_hour, end_hour,
                    stop_when_submitted, is_enabled,
                    escalation_enabled, escalation_last_days,
                    escalation_send_to_leaders, escalation_extra_chat_ids,
                    created_by
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    name, preset_id, flt_json_clean, message_template,
                    day_from, day_to, times_per_day, start_hour, end_hour,
                    stop_when_submitted, is_enabled,
                    escalation_enabled, escalation_last_days,
                    escalation_send_to_leaders, escalation_extra_clean,
                    current_username(),
                ),
            )
            flash(f'Правило авто-нагадування «{name}» створено ({day_from}–{day_to} числа, {times_per_day} р./день).', 'success')
        conn.commit()
    finally:
        conn.close()

    log_action(
        current_username(),
        f"{'оновив' if rule_id else 'створив'} правило авто-нагадування ВОД: «{name}»",
        category='update',
        section='Військовий облік',
        details=f"дні: {day_from}..{day_to}, повторів/день: {times_per_day} ({start_hour:02d}:00-{end_hour:02d}:00), ескалація: {'так' if escalation_enabled else 'ні'}",
    )
    return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))


@military_telegram_bp.route('/admin/military/telegram/auto-rules/<int:rule_id>/escalate-now', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_auto_rule_escalate_now(rule_id):
    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        settings = _get_bot_settings(conn)
        token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
        if not token:
            flash('Спочатку вкажіть токен Telegram-бота в налаштуваннях.', 'error')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='settings'))
        row = conn.execute("SELECT * FROM telegram_auto_reminder_rules WHERE id = ?", (rule_id,)).fetchone()
        if not row:
            flash('Правило не знайдено.', 'warning')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))
        rule = dict(row)
        res = _execute_rule_escalation(conn, token, rule, force=True)
        if res['groups_missing'] == 0:
            flash(f"По правилу «{rule['name']}» боржників за поточний місяць немає — ескалація не потрібна.", 'info')
        elif res['sent_leaders'] == 0 and res['sent_extra'] == 0 and res['failed'] == 0:
            flash(
                f"Знайдено {res['missing_students']} боржників у {res['groups_missing']} групах, "
                "але не вказано жодного Chat ID старости/куратора або координатора.",
                'warning',
            )
        else:
            flash(
                f"Ескалацію надіслано! Груп з боржниками: {res['groups_missing']} "
                f"(студентів: {res['missing_students']}). Старостам/кураторам: {res['sent_leaders']}, "
                f"координаторам: {res['sent_extra']}, помилок: {res['failed']}.",
                'success' if res['failed'] == 0 else 'warning',
            )
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))


@military_telegram_bp.route('/admin/military/telegram/auto-rules/<int:rule_id>/toggle', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_auto_rule_toggle(rule_id):
    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        row = conn.execute("SELECT name, is_enabled FROM telegram_auto_reminder_rules WHERE id = ?", (rule_id,)).fetchone()
        if row:
            new_val = 0 if int(row['is_enabled'] or 0) else 1
            conn.execute(
                "UPDATE telegram_auto_reminder_rules SET is_enabled = ?, updated_at = datetime('now', 'localtime') WHERE id = ?",
                (new_val, rule_id),
            )
            conn.commit()
            flash(
                f"Правило «{row['name']}» {'увімкнено' if new_val else 'призупинено'}.",
                'success' if new_val else 'info',
            )
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))


@military_telegram_bp.route('/admin/military/telegram/auto-rules/<int:rule_id>/run-now', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_auto_rule_run_now(rule_id):
    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        settings = _get_bot_settings(conn)
        token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
        if not token:
            flash('Спочатку вкажіть токен Telegram-бота в налаштуваннях.', 'error')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='settings'))
        row = conn.execute("SELECT * FROM telegram_auto_reminder_rules WHERE id = ?", (rule_id,)).fetchone()
        if not row:
            flash('Правило не знайдено.', 'warning')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))
        rule = dict(row)
        summary = _execute_auto_rule_dispatch(conn, token, rule, created_by=current_username())
        if summary.get('skipped_empty'):
            flash(
                f"Правило «{rule['name']}»: збігів за фільтром {summary['total_matched']}, "
                "прив’язаних до Telegram чатів для відправки немає (або всі вже подали PDF).",
                'info',
            )
        else:
            flash(
                f"Правило «{rule['name']}» виконано зараз: підійшло {summary['total_matched']}, "
                f"надіслано чатів {summary['sent_chats']}, помилок {summary['failed_chats']}.",
                'success' if summary['failed_chats'] == 0 else 'warning',
            )
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))


@military_telegram_bp.route('/admin/military/telegram/auto-rules/<int:rule_id>/delete', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_auto_rule_delete(rule_id):
    conn = get_db()
    try:
        conn.execute("DELETE FROM telegram_auto_reminder_rules WHERE id = ?", (rule_id,))
        conn.commit()
        flash('Правило авто-нагадування видалено.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))


@military_telegram_bp.route('/admin/military/telegram/broadcast', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_send_broadcast():
    title = (request.form.get('broadcast_title') or 'Нагадування про ВОД «Резерв+»').strip()
    message_template = (request.form.get('message_text') or '').strip()
    if not message_template:
        flash('Введіть текст повідомлення для розсилки.', 'warning')
        return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))

    raw_flt = request.form.get('filter_json')
    filters = _parse_filter_dict(raw_flt) if raw_flt else _extract_filters_from_request(request.form)

    conn = get_db()
    try:
        settings = _get_bot_settings(conn)
        token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
        if not token:
            flash('Спочатку вкажіть токен Telegram-бота в налаштуваннях.', 'error')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='settings'))

        summary = execute_broadcast(
            conn,
            token=token,
            title=title,
            message_template=message_template,
            raw_filters=filters,
            created_by=current_username(),
            is_auto=False,
        )
    finally:
        conn.close()

    log_action(
        current_username(),
        f"надіслав розсилку в Telegram-бот ВОД: «{title}»",
        category='update',
        section='Військовий облік',
        details=f"збіг: {summary['total_matched']}, надіслано чатів: {summary['sent_chats']}, помилок: {summary['failed_chats']}",
    )
    flash(
        f"Розсилку виконано! Підійшло студентів: {summary['total_matched']} "
        f"(з Telegram: {summary['linked_count']}, без Telegram: {summary['unlinked_count']}). "
        f"Надіслано чатів: {summary['sent_chats']}, помилок: {summary['failed_chats']}.",
        'success' if summary['failed_chats'] == 0 else 'warning',
    )
    return redirect(url_for('military_telegram.telegram_dashboard', tab='broadcast'))


@military_telegram_bp.route('/admin/military/telegram/submissions/<int:sub_id>')
@permission_required('manage_military_telegram')
def telegram_submission_review(sub_id):
    from routes.students.military import MILITARY_FIELDS, STUDENT_FIELDS, _current_military, _current_student, _field_label

    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        sub_row = conn.execute("SELECT * FROM telegram_vod_submissions WHERE id = ?", (sub_id,)).fetchone()
        if not sub_row:
            flash('Подання не знайдено.', 'warning')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))
        sub = dict(sub_row)

        student = _current_student(conn, sub['student_id']) if sub.get('student_id') else None
        military = _current_military(conn, sub['student_id']) if sub.get('student_id') else None

        try:
            parsed = json.loads(sub.get('parsed_json') or '{}')
        except Exception:
            parsed = {}
        try:
            new_military = json.loads(sub.get('new_military_json') or '{}')
        except Exception:
            new_military = {}
        try:
            new_student = json.loads(sub.get('new_student_json') or '{}')
        except Exception:
            new_student = {}
        try:
            mil_diff = json.loads(sub.get('military_diff_json') or '[]')
        except Exception:
            mil_diff = []
        try:
            stu_diff = json.loads(sub.get('student_diff_json') or '[]')
        except Exception:
            stu_diff = []

        # Перераховуємо актуальний diff і критичні ризики на момент відкриття (якщо студент відомий та ще pending)
        risk_flags = []
        risk_level = 'none'
        if sub.get('student_id') and sub['status'] == 'pending':
            new_military, new_student, mil_diff, stu_diff = _build_submission_diffs(conn, int(sub['student_id']), parsed)
            risk_flags, risk_level = _analyze_submission_risks(conn, int(sub['student_id']), parsed, cur_mil=military)
        elif sub.get('student_id'):
            _hydrate_submission_risks(conn, sub, cur_mil_by_sid={int(sub['student_id']): (military or {})})
            risk_flags = sub.get('risk_flags') or []
            risk_level = sub.get('risk_level') or 'none'

        active_students = []
        if sub['status'] in ('not_found', 'pending'):
            active_students = [
                {
                    'id': r['id'],
                    'full_name': ' '.join(x for x in (r['last_name_UA'], r['first_name_UA'], r['middle_name_UA'] or '') if x).strip(),
                    'birth_date': r['birth_date'] or '',
                    'group_name': r['group_name'] or '',
                }
                for r in conn.execute(
                    f"""
                    SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA, s.birth_date, g.name AS group_name
                    FROM students s LEFT JOIN groups g ON g.id = s.group_id
                    WHERE {_active_students_sql_where('s', include_frozen=True)}
                    ORDER BY s.last_name_UA COLLATE UKRAINIAN
                    """
                ).fetchall()
            ]
    finally:
        conn.close()

    return render_template(
        'military_reports/telegram_submission_review.html',
        sub=sub,
        student=student,
        military=military,
        parsed=parsed,
        new_military=new_military,
        new_student=new_student,
        mil_diff=mil_diff,
        stu_diff=stu_diff,
        risk_flags=risk_flags,
        risk_level=risk_level,
        military_fields=MILITARY_FIELDS,
        student_fields=STUDENT_FIELDS,
        field_label=_field_label,
        active_students=active_students,
    )


def _apply_submission_approval(conn, sub: dict, *, apply_mil_fields: set | None,
                               apply_stu_fields: set | None, apply_photo: bool,
                               username: str, notify_telegram: bool = True) -> dict:
    from routes.students.military import (
        MILITARY_FIELDS,
        _ensure_military_record,
        _log_history,
        _save_military,
        _save_pdf_as_attachment,
        _save_student_fields,
    )

    sid = int(sub['student_id'])
    student_row = conn.execute(
        "SELECT id, last_name_UA, first_name_UA, middle_name_UA, group_id, photo FROM students WHERE id = ?",
        (sid,),
    ).fetchone()
    if not student_row:
        raise ValueError("Студента не знайдено в БД")

    new_military = json.loads(sub.get('new_military_json') or '{}')
    new_student = json.loads(sub.get('new_student_json') or '{}')

    if apply_mil_fields is None:
        mil_payload = {k: v for k, v in new_military.items() if v not in ('', None)}
    else:
        mil_payload = {k: v for k, v in new_military.items() if k in apply_mil_fields}

    mil_diff = []
    if mil_payload:
        mil_id, mil_diff = _save_military(conn, sid, mil_payload, source='telegram_bot', username=username)
    else:
        mil_id = _ensure_military_record(conn, sid)

    stu_diff = []
    if apply_stu_fields:
        stu_payload = {k: v for k, v in new_student.items() if k in apply_stu_fields}
        if stu_payload:
            stu_diff = _save_student_fields(conn, sid, stu_payload, source='telegram_bot', username=username)

    photo_saved = False
    if apply_photo and sub.get('photo_rel_path'):
        static_root = os.path.join(PROJECT_ROOT, 'static')
        p_abs = resolve_under_dir(static_root, sub['photo_rel_path'])
        if p_abs and os.path.exists(p_abs):
            try:
                with open(p_abs, 'rb') as pf:
                    raw = pf.read()
                img = load_and_validate_image(raw)
                crop = auto_center_crop_box(img.width, img.height)
                final_img = crop_and_resize(img, crop)
                new_rel = save_final_image(final_img, sid)
                conn.execute("UPDATE students SET photo = ? WHERE id = ?", (new_rel, sid))
                _log_history(
                    conn, sid, username, 'telegram_bot', 'student',
                    [{'field': 'photo', 'label': 'Фото 3×4', 'old': student_row['photo'] or '', 'new': new_rel}],
                )
                photo_saved = True
            except Exception as e:
                logger.warning(f"Telegram approve photo warning (sub_id={sub['id']}): {e}")

    approved_att_id = None
    if sub.get('pdf_rel_path'):
        static_root = os.path.join(PROJECT_ROOT, 'static')
        pdf_abs = resolve_under_dir(static_root, sub['pdf_rel_path'])
        if pdf_abs and os.path.exists(pdf_abs):
            with open(pdf_abs, 'rb') as pf:
                pdf_bytes = pf.read()
            orig_name = sub.get('original_filename') or 'reserve_plus_telegram.pdf'
            _save_pdf_as_attachment(conn, sid, pdf_bytes, orig_name, username)
            row_att = conn.execute(
                "SELECT id FROM attachments WHERE entity_type='military' AND entity_id=? ORDER BY id DESC LIMIT 1",
                (mil_id,),
            ).fetchone()
            if row_att:
                approved_att_id = int(row_att['id'])

    conn.execute(
        """
        UPDATE telegram_vod_submissions
        SET status = 'approved',
            reviewed_by = ?,
            reviewed_at = datetime('now', 'localtime'),
            approved_attachment_id = ?
        WHERE id = ?
        """,
        (username, approved_att_id, sub['id']),
    )

    full_fio = ' '.join(
        x for x in (student_row['last_name_UA'], student_row['first_name_UA'], student_row['middle_name_UA'] or '') if x
    ).strip()

    _record_event(
        conn, 'submission_approved',
        f"Підтверджено ВОД з Telegram: {full_fio}",
        severity='info',
        details=f"Подання #{sub['id']}, змін військових: {len(mil_diff)}, картки: {len(stu_diff)}",
        student_id=sid,
        submission_id=sub['id'],
    )
    conn.commit()

    log_action(
        username,
        f"підтвердив ВОД із Telegram: {full_fio} (ID {sid})",
        group_ids=[student_row['group_id']],
        category='import',
        section='Військовий облік',
        entity_type='student',
        entity_id=sid,
        entity_name=full_fio,
        changes=list(mil_diff) + list(stu_diff),
        details=f"подання #{sub['id']}, фото: {'так' if photo_saved else 'ні'}",
    )

    if notify_telegram and sub.get('chat_id'):
        settings = _get_bot_settings(conn)
        token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
        if token:
            tpl = (settings.get('approval_text') or DEFAULT_APPROVAL_TEXT).strip()
            formed_txt = f" (дата формування PDF: {sub['parsed_formed_at']})" if sub.get('parsed_formed_at') else ''
            msg = tpl.replace('{ПІБ}', full_fio).replace('{ДатаСформовано}', formed_txt)
            _tg_send_message(token, sub['chat_id'], msg)

    return {'mil_diff': len(mil_diff), 'stu_diff': len(stu_diff), 'photo_saved': photo_saved}


@military_telegram_bp.route('/admin/military/telegram/submissions/<int:sub_id>/approve', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_submission_approve(sub_id):
    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        sub_row = conn.execute("SELECT * FROM telegram_vod_submissions WHERE id = ?", (sub_id,)).fetchone()
        if not sub_row:
            flash('Подання не знайдено.', 'warning')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))
        sub = dict(sub_row)
        if not sub.get('student_id'):
            flash('Спочатку прив’яжіть подання до активного студента в БД.', 'error')
            return redirect(url_for('military_telegram.telegram_submission_review', sub_id=sub_id))
        if sub['status'] != 'pending':
            flash('Це подання вже було опрацьоване.', 'info')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))

        if request.form.get('explicit_fields') == '1':
            apply_mil = set(request.form.getlist('apply_mil_field'))
            apply_stu = set(request.form.getlist('apply_stu_field'))
            apply_photo = bool(request.form.get('apply_photo'))
        else:
            apply_mil = None
            apply_stu = set()
            apply_photo = False

        res = _apply_submission_approval(
            conn, sub,
            apply_mil_fields=apply_mil,
            apply_stu_fields=apply_stu,
            apply_photo=apply_photo,
            username=current_username(),
            notify_telegram=True,
        )
        flash(
            f"Підтверджено! Оновлено військових полів: {res['mil_diff']}, картки: {res['stu_diff']}, "
            f"фото: {'оновлено' if res['photo_saved'] else 'без змін'}. Студента повідомлено в Telegram.",
            'success',
        )
    except Exception as e:
        conn.rollback()
        logger.error(f"Помилка підтвердження подання #{sub_id}: {e}", exc_info=True)
        flash(f'Помилка підтвердження: {e}', 'error')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))


@military_telegram_bp.route('/admin/military/telegram/submissions/<int:sub_id>/reject', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_submission_reject(sub_id):
    reason = (request.form.get('reject_reason') or 'Документ не відповідає вимогам або застарілий').strip()
    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        sub_row = conn.execute("SELECT * FROM telegram_vod_submissions WHERE id = ?", (sub_id,)).fetchone()
        if not sub_row:
            flash('Подання не знайдено.', 'warning')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))
        sub = dict(sub_row)
        conn.execute(
            """
            UPDATE telegram_vod_submissions
            SET status = 'rejected', reject_reason = ?,
                reviewed_by = ?, reviewed_at = datetime('now', 'localtime')
            WHERE id = ?
            """,
            (reason, current_username(), sub_id),
        )
        fio = sub.get('parsed_full_name') or 'Студент'
        _record_event(
            conn, 'submission_rejected',
            f"Відхилено подання #{sub_id} ({fio})",
            severity='info',
            details=reason,
            student_id=sub.get('student_id'),
            submission_id=sub_id,
        )
        conn.commit()

        settings = _get_bot_settings(conn)
        token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
        if token and sub.get('chat_id'):
            tpl = (settings.get('rejection_text') or DEFAULT_REJECTION_TEXT).strip()
            msg = tpl.replace('{ПІБ}', fio).replace('{Причина}', reason)
            _tg_send_message(token, sub['chat_id'], msg)

        log_action(
            current_username(),
            f"відхилив ВОД із Telegram #{sub_id} ({fio})",
            category='update',
            section='Військовий облік',
            details=reason,
        )
        flash('Подання відхилено, повідомлення надіслано в Telegram.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))


@military_telegram_bp.route('/admin/military/telegram/submissions/<int:sub_id>/rematch', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_submission_rematch(sub_id):
    """Ручне співставлення не знайденого в БД PDF із конкретним студентом."""
    student_id = request.form.get('student_id', type=int)
    also_link_tg = bool(request.form.get('also_link_tg'))
    is_rep = bool(request.form.get('is_representative'))
    if not student_id:
        flash('Оберіть студента зі списку.', 'warning')
        return redirect(url_for('military_telegram.telegram_submission_review', sub_id=sub_id))

    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        sub_row = conn.execute("SELECT * FROM telegram_vod_submissions WHERE id = ?", (sub_id,)).fetchone()
        st = conn.execute("SELECT * FROM students WHERE id = ?", (student_id,)).fetchone()
        if not sub_row or not st:
            flash('Запис або студента не знайдено.', 'error')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))
        sub = dict(sub_row)
        try:
            parsed = json.loads(sub.get('parsed_json') or '{}')
        except Exception:
            parsed = {}

        new_mil, new_stu, mil_diff, stu_diff = _build_submission_diffs(conn, student_id, parsed)
        risk_flags, risk_level = _analyze_submission_risks(conn, student_id, parsed)
        is_dup, dup_cnt, dup_info = _compute_monthly_duplicate_info(
            conn, student_id, sub.get('parsed_ym'), sub.get('upload_ym') or date.today().strftime('%Y-%m'),
            exclude_sub_id=sub_id,
        )
        conn.execute(
            """
            UPDATE telegram_vod_submissions
            SET student_id = ?,
                status = 'pending',
                parse_error = '',
                new_military_json = ?,
                new_student_json = ?,
                military_diff_json = ?,
                student_diff_json = ?,
                is_monthly_duplicate = ?,
                duplicate_count = ?,
                duplicate_info = ?,
                risk_flags_json = ?,
                risk_level = ?,
                risk_count = ?
            WHERE id = ?
            """,
            (
                student_id,
                json.dumps(new_mil, ensure_ascii=False),
                json.dumps(new_stu, ensure_ascii=False),
                json.dumps(mil_diff, ensure_ascii=False),
                json.dumps(stu_diff, ensure_ascii=False),
                1 if is_dup else 0,
                dup_cnt if is_dup else 1,
                dup_info,
                json.dumps(risk_flags, ensure_ascii=False),
                risk_level,
                len(risk_flags),
                sub_id,
            ),
        )
        if also_link_tg and sub.get('chat_id'):
            _upsert_student_link(
                conn,
                student_id=student_id,
                chat_id=sub['chat_id'],
                from_user={
                    'id': sub.get('telegram_user_id') or sub['chat_id'],
                    'username': sub.get('tg_username') or '',
                    'first_name': sub.get('tg_first_name') or '',
                    'last_name': sub.get('tg_last_name') or '',
                },
                link_method='manual',
                is_representative=is_rep,
                tg_phone=sub.get('tg_phone') or '',
                representative_note='Прив’язано вручну при розборі PDF',
            )
        conn.commit()
        flash('Документ прив’язано до студента та перенесено в чергу перевірки.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_submission_review', sub_id=sub_id))


@military_telegram_bp.route('/admin/military/telegram/submissions/bulk-approve', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_bulk_approve():
    ids = [int(x) for x in request.form.getlist('submission_ids') if str(x).isdigit()]
    if not ids:
        flash('Оберіть хоча б одне подання для масового підтвердження.', 'warning')
        return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))

    conn = get_db()
    approved_cnt = 0
    skipped_cnt = 0
    try:
        ensure_telegram_vod_schema(conn)
        for sid_sub in ids:
            sub_row = conn.execute("SELECT * FROM telegram_vod_submissions WHERE id = ?", (sid_sub,)).fetchone()
            if not sub_row:
                skipped_cnt += 1
                continue
            sub = dict(sub_row)
            if sub['status'] != 'pending' or not sub.get('student_id'):
                skipped_cnt += 1
                continue
            try:
                _apply_submission_approval(
                    conn, sub,
                    apply_mil_fields=None,
                    apply_stu_fields=set(),
                    apply_photo=False,
                    username=current_username(),
                    notify_telegram=True,
                )
                approved_cnt += 1
            except Exception as e:
                conn.rollback()
                logger.error(f"Bulk approve error (sub_id={sid_sub}): {e}")
                skipped_cnt += 1
    finally:
        conn.close()

    flash(f"Масове підтвердження: підтверджено {approved_cnt}, пропущено {skipped_cnt}.", 'success')
    return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))


@military_telegram_bp.route('/admin/military/telegram/submissions/<int:sub_id>/delete', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_submission_delete(sub_id):
    conn = get_db()
    try:
        sub = conn.execute("SELECT * FROM telegram_vod_submissions WHERE id = ?", (sub_id,)).fetchone()
        if sub:
            static_root = os.path.join(PROJECT_ROOT, 'static')
            for col in ('pdf_rel_path', 'photo_rel_path', 'photo_raw_rel_path'):
                if sub[col]:
                    p = resolve_under_dir(static_root, sub[col])
                    if p and os.path.exists(p):
                        try:
                            os.remove(p)
                        except OSError:
                            pass
            conn.execute("DELETE FROM telegram_vod_submissions WHERE id = ?", (sub_id,))
            conn.commit()
            flash('Запис подання видалено.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='queue'))


# ═══════════════════════════════════════════════════════════════════════════════
# 8. Запрошення (QR / Deep Links) і ручна прив'язка
# ═══════════════════════════════════════════════════════════════════════════════

def _create_invite_token_row(conn, *, scope: str, student_id: int | None, group_id: int | None,
                             is_rep_default: bool, note: str, expires_days: int | None,
                             max_uses: int | None, created_by: str) -> dict:
    ensure_telegram_vod_schema(conn)
    token_str = secrets.token_hex(10)
    exp_str = None
    if expires_days and expires_days > 0:
        exp_str = (datetime.now() + timedelta(days=expires_days)).strftime('%Y-%m-%d %H:%M:%S')
    cur = conn.execute(
        """
        INSERT INTO telegram_invite_tokens (
            token, scope, student_id, group_id, is_representative_default,
            note, created_by, expires_at, max_uses, is_active
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
        """,
        (token_str, scope, student_id, group_id, 1 if is_rep_default else 0, note, created_by, exp_str, max_uses),
    )
    conn.commit()
    settings = _get_bot_settings(conn)
    bot_user = (settings.get('bot_username') or '').strip().lstrip('@')
    prefix = 's_' if scope == 'student' else 'g_'
    tg_url = f"https://t.me/{bot_user}?start={prefix}{token_str}" if bot_user else f"?start={prefix}{token_str}"
    return {'id': int(cur.lastrowid), 'token': token_str, 'tg_url': tg_url, 'bot_username': bot_user}


@military_telegram_bp.route('/admin/military/telegram/invites/create', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_invite_create():
    scope = (request.form.get('scope') or 'student').strip()
    if scope not in ('student', 'group'):
        scope = 'student'
    student_id = request.form.get('student_id', type=int) if scope == 'student' else None
    group_id = request.form.get('group_id', type=int) if scope == 'group' else None
    is_rep = bool(request.form.get('is_representative_default'))
    note = (request.form.get('note') or '').strip()
    expires_days = request.form.get('expires_days', type=int)
    max_uses = request.form.get('max_uses', type=int)

    if scope == 'student' and not student_id:
        flash('Оберіть студента для створення посилання.', 'warning')
        return redirect(url_for('military_telegram.telegram_dashboard', tab='links'))
    if scope == 'group' and not group_id:
        flash('Оберіть групу для створення групового посилання.', 'warning')
        return redirect(url_for('military_telegram.telegram_dashboard', tab='links'))

    conn = get_db()
    try:
        res = _create_invite_token_row(
            conn,
            scope=scope,
            student_id=student_id,
            group_id=group_id,
            is_rep_default=is_rep,
            note=note,
            expires_days=expires_days,
            max_uses=max_uses,
            created_by=current_username(),
        )
        flash(f"Запрошення створено: {res['tg_url']}", 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='links'))


@military_telegram_bp.route('/admin/military/telegram/student/<int:student_id>/quick-invite', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_student_quick_invite(student_id):
    is_rep = bool(request.form.get('is_representative_default') in ('1', 'true', 'on'))
    note = (request.form.get('note') or '').strip()
    conn = get_db()
    try:
        st = conn.execute("SELECT id FROM students WHERE id = ?", (student_id,)).fetchone()
        if not st:
            return jsonify({'ok': False, 'error': 'Студента не знайдено'}), 404
        res = _create_invite_token_row(
            conn,
            scope='student',
            student_id=student_id,
            group_id=None,
            is_rep_default=is_rep,
            note=note,
            expires_days=30,
            max_uses=None,
            created_by=current_username(),
        )
        qr_url = url_for('military_telegram.telegram_qr_svg', data=res['tg_url'])
        return jsonify({'ok': True, 'tg_url': res['tg_url'], 'qr_url': qr_url, 'bot_username': res['bot_username']})
    finally:
        conn.close()


@military_telegram_bp.route('/admin/military/telegram/invites/<int:inv_id>/revoke', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_invite_revoke(inv_id):
    conn = get_db()
    try:
        conn.execute("UPDATE telegram_invite_tokens SET is_active = 0 WHERE id = ?", (inv_id,))
        conn.commit()
        flash('Посилання-запрошення відкликано.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='links'))


@military_telegram_bp.route('/admin/military/telegram/links/manual', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_link_manual():
    student_id = request.form.get('student_id', type=int)
    chat_id = (request.form.get('chat_id') or '').strip()
    tg_username = (request.form.get('tg_username') or '').strip().lstrip('@')
    tg_phone = normalize_phone_ua(request.form.get('tg_phone') or '') or ''
    is_rep = bool(request.form.get('is_representative'))
    rep_note = (request.form.get('representative_note') or '').strip()

    if not student_id or not chat_id:
        flash('Вкажіть студента та Telegram Chat ID.', 'warning')
        return redirect(url_for('military_telegram.telegram_dashboard', tab='links'))

    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        _upsert_student_link(
            conn,
            student_id=student_id,
            chat_id=chat_id,
            from_user={'id': chat_id, 'username': tg_username},
            link_method='manual',
            is_representative=is_rep,
            tg_phone=tg_phone,
            representative_note=rep_note,
        )
        conn.commit()
        flash('Прив’язку Telegram збережено.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='links'))


@military_telegram_bp.route('/admin/military/telegram/links/<int:link_id>/toggle-rep', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_link_toggle_rep(link_id):
    conn = get_db()
    try:
        row = conn.execute("SELECT is_representative FROM telegram_student_links WHERE id = ?", (link_id,)).fetchone()
        if row:
            new_val = 0 if int(row['is_representative'] or 0) else 1
            conn.execute(
                "UPDATE telegram_student_links SET is_representative = ?, updated_at = datetime('now', 'localtime') WHERE id = ?",
                (new_val, link_id),
            )
            conn.commit()
            flash('Ознаку представника оновлено.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='links'))


@military_telegram_bp.route('/admin/military/telegram/links/<int:link_id>/delete', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_link_delete(link_id):
    conn = get_db()
    try:
        conn.execute("DELETE FROM telegram_student_links WHERE id = ?", (link_id,))
        conn.commit()
        flash('Прив’язку Telegram видалено.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab='links'))


@military_telegram_bp.route('/admin/military/telegram/qr.svg')
@permission_required('manage_military_telegram')
def telegram_qr_svg():
    data = (request.args.get('data') or '').strip()
    if not data or len(data) > 210:
        return Response('Некоректні дані для QR', status=400)
    try:
        svg = _make_qr_svg(data, size_px=240)
        return Response(svg, mimetype='image/svg+xml', headers={'Cache-Control': 'private, max-age=300'})
    except Exception as e:
        return Response(f'Помилка генерації QR: {e}', status=400)


# ═══════════════════════════════════════════════════════════════════════════════
# 9. Налаштування та управління фоновим ботом
# ═══════════════════════════════════════════════════════════════════════════════

@military_telegram_bp.route('/admin/military/telegram/settings', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_save_settings():
    new_token = (request.form.get('bot_token') or '').strip()
    welcome_text = (request.form.get('welcome_text') or '').strip() or DEFAULT_WELCOME_TEXT
    reminder_text = (request.form.get('reminder_default_text') or '').strip() or DEFAULT_REMINDER_TEXT
    approval_text = (request.form.get('approval_text') or '').strip() or DEFAULT_APPROVAL_TEXT
    rejection_text = (request.form.get('rejection_text') or '').strip() or DEFAULT_REJECTION_TEXT

    auto_remind_enabled = 1 if request.form.get('auto_remind_enabled') else 0
    auto_remind_day = max(1, min(28, request.form.get('auto_remind_day', type=int) or 5))
    auto_remind_hour = max(0, min(23, request.form.get('auto_remind_hour', type=int) or 10))
    auto_remind_filter_raw = (request.form.get('auto_remind_filter_json') or '{}').strip()
    auto_remind_filter = json.dumps(_parse_filter_dict(auto_remind_filter_raw), ensure_ascii=False)

    auto_cleanup_enabled = 1 if request.form.get('auto_cleanup_enabled') else 0
    auto_cleanup_months = max(1, min(60, request.form.get('auto_cleanup_months', type=int) or 6))
    auto_cleanup_include_not_found = 1 if request.form.get('auto_cleanup_include_not_found') else 0
    auto_cleanup_include_approved = 1 if request.form.get('auto_cleanup_include_approved') else 0

    licenses_submitted = request.form.get('military_licenses_submitted') == '1'
    selected_license_ids = [
        int(x) for x in request.form.getlist('military_license_ids')
        if str(x).strip().isdigit()
    ]
    mil_lic_details = ''

    conn = get_db()
    try:
        cur_settings = _get_bot_settings(conn)
        effective_token = new_token if new_token else (cur_settings.get('bot_token') or '')
        bot_username = cur_settings.get('bot_username') or ''
        bot_first_name = cur_settings.get('bot_first_name') or ''

        if effective_token and (new_token or not bot_username):
            ok, me, err = _tg_call(effective_token, 'getMe', timeout=10.0)
            if ok and isinstance(me, dict):
                bot_username = me.get('username') or ''
                bot_first_name = me.get('first_name') or ''
            elif new_token:
                flash(f'Увага: токен збережено, але Telegram повернув помилку перевірки: {err}', 'warning')

        conn.execute(
            """
            UPDATE telegram_bot_settings
            SET bot_token = ?,
                bot_username = ?,
                bot_first_name = ?,
                welcome_text = ?,
                reminder_default_text = ?,
                approval_text = ?,
                rejection_text = ?,
                auto_remind_enabled = ?,
                auto_remind_day = ?,
                auto_remind_hour = ?,
                auto_remind_filter_json = ?,
                auto_cleanup_enabled = ?,
                auto_cleanup_months = ?,
                auto_cleanup_include_not_found = ?,
                auto_cleanup_include_approved = ?,
                updated_at = datetime('now', 'localtime')
            WHERE id = 1
            """,
            (
                effective_token, bot_username, bot_first_name,
                welcome_text, reminder_text, approval_text, rejection_text,
                auto_remind_enabled, auto_remind_day, auto_remind_hour, auto_remind_filter,
                auto_cleanup_enabled, auto_cleanup_months,
                auto_cleanup_include_not_found, auto_cleanup_include_approved,
            ),
        )
        if licenses_submitted:
            from routes.military_reports import _set_military_visible_licenses
            lic_ok, mil_label = _set_military_visible_licenses(conn, selected_license_ids)
            if not lic_ok:
                flash('Потрібно залишити обраною хоча б одну ліцензію для списків військового обліку.', 'warning')
            else:
                mil_lic_details = f"ліцензії ВО: {mil_label}"
        conn.commit()
    finally:
        conn.close()

    if new_token and get_worker_status().get('running'):
        stop_bot_worker(disable_in_db=False)
        start_bot_worker()

    log_action(
        current_username(),
        "оновив налаштування Telegram-бота ВОД",
        category='update',
        section='Військовий облік',
        details=mil_lic_details,
    )
    flash('Налаштування Telegram-бота збережено.', 'success')
    return redirect(url_for('military_telegram.telegram_dashboard', tab='settings'))


@military_telegram_bp.route('/admin/military/telegram/cleanup', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_run_cleanup_now():
    months = request.form.get('cleanup_months', type=int)
    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        settings = _get_bot_settings(conn)
        eff_months = max(1, min(60, months or int(settings.get('auto_cleanup_months') or 6)))
        inc_nf = bool(int(settings.get('auto_cleanup_include_not_found') if settings.get('auto_cleanup_include_not_found') is not None else 1))
        inc_app = bool(int(settings.get('auto_cleanup_include_approved') or 0))
        res = cleanup_old_telegram_vod_files(
            conn,
            months=eff_months,
            include_not_found=inc_nf,
            include_approved=inc_app,
            dry_run=False,
        )
    finally:
        conn.close()

    log_action(
        current_username(),
        "виконав очищення тимчасових файлів черги Telegram ВОД",
        category='delete',
        section='Військовий облік',
        details=f"старше {eff_months} міс.: записів {res['cleaned_rows']}, файлів {res['deleted_files']} ({res['freed_mb']} МБ)",
    )
    if res['deleted_files'] == 0 and res['cleaned_rows'] == 0:
        flash(f"Застарілих файлів старше {eff_months} міс. (до {res['cutoff_date']}) не знайдено.", 'info')
    else:
        flash(
            f"Очищення завершено! Опрацьовано подань: {res['cleaned_rows']}, видалено файлів: {res['deleted_files']} "
            f"(звільнено {res['freed_mb']} МБ).",
            'success',
        )
    return redirect(url_for('military_telegram.telegram_dashboard', tab='settings'))


@military_telegram_bp.route('/admin/military/telegram/group-contacts/save', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_group_contact_save():
    group_id = request.form.get('group_id', type=int)
    leader_name = (request.form.get('leader_name') or '').strip()
    leader_chat_id = (request.form.get('leader_chat_id') or '').strip()
    curator_name = (request.form.get('curator_name') or '').strip()
    curator_chat_id = (request.form.get('curator_chat_id') or '').strip()
    notes = (request.form.get('notes') or '').strip()
    is_active = 1 if request.form.get('is_active') else 0
    return_tab = (request.form.get('return_tab') or 'groups').strip()
    if return_tab not in ('groups', 'links'):
        return_tab = 'groups'

    if not group_id:
        flash('Оберіть академічну групу.', 'warning')
        return redirect(url_for('military_telegram.telegram_dashboard', tab=return_tab))
    if not leader_chat_id and not curator_chat_id:
        flash('Вкажіть хоча б один Telegram Chat ID (старости або куратора).', 'warning')
        return redirect(url_for('military_telegram.telegram_dashboard', tab=return_tab))

    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        grp = conn.execute("SELECT id, name FROM groups WHERE id = ?", (group_id,)).fetchone()
        if not grp:
            flash('Групу не знайдено.', 'error')
            return redirect(url_for('military_telegram.telegram_dashboard', tab=return_tab))
        conn.execute(
            """
            INSERT INTO telegram_group_contacts (
                group_id, leader_name, leader_chat_id, curator_name, curator_chat_id, notes, is_active
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(group_id) DO UPDATE SET
                leader_name = excluded.leader_name,
                leader_chat_id = excluded.leader_chat_id,
                curator_name = excluded.curator_name,
                curator_chat_id = excluded.curator_chat_id,
                notes = excluded.notes,
                is_active = excluded.is_active,
                updated_at = datetime('now', 'localtime')
            """,
            (group_id, leader_name, leader_chat_id, curator_name, curator_chat_id, notes, is_active),
        )
        conn.commit()
        flash(f"Контакти для звітності по групі «{grp['name']}» збережено.", 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab=return_tab))


@military_telegram_bp.route('/admin/military/telegram/group-contacts/<int:contact_id>/delete', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_group_contact_delete(contact_id):
    return_tab = (request.form.get('return_tab') or 'links').strip()
    if return_tab not in ('groups', 'links'):
        return_tab = 'links'
    conn = get_db()
    try:
        conn.execute("DELETE FROM telegram_group_contacts WHERE id = ?", (contact_id,))
        conn.commit()
        flash('Контакт старости/куратора видалено.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab=return_tab))


@military_telegram_bp.route('/admin/military/telegram/group-summary/send', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_group_summary_send():
    vod_month = (request.form.get('vod_month') or date.today().strftime('%Y-%m')).strip()
    group_id = request.form.get('group_id', type=int)
    age_mode = (request.form.get('gage') or 'any').strip()
    gall_lic = request.form.get('gall_lic') == '1'
    send_all_missing = bool(request.form.get('send_all_missing') == '1')
    custom_chat_id = (request.form.get('custom_chat_id') or '').strip()

    conn = get_db()
    try:
        ensure_telegram_vod_schema(conn)
        settings = _get_bot_settings(conn)
        token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
        if not token:
            flash('Спочатку вкажіть токен Telegram-бота в налаштуваннях.', 'error')
            return redirect(url_for('military_telegram.telegram_dashboard', tab='settings'))

        flt = {
            'vod_month': vod_month,
            'group_ids': [group_id] if (group_id and not send_all_missing) else [],
            'age_mode': age_mode if age_mode in ('any', '25_plus', 'under_25') else 'any',
            'only_base_license': not gall_lic,
            'exclude_women_no_mil': True,
            'exclude_note_18': True,
        }
        summary = build_group_vod_compliance_summary(
            conn, flt, vod_month=vod_month, only_with_missing=send_all_missing
        )
        sent = 0
        failed = 0
        groups_touched = 0
        for grp in summary['groups']:
            cinfo = grp.get('contact') or {}
            targets = []
            if custom_chat_id and not send_all_missing:
                targets.append(custom_chat_id)
            else:
                if cinfo and not int(cinfo.get('is_active') if cinfo.get('is_active') is not None else 1):
                    continue
                for k in ('leader_chat_id', 'curator_chat_id'):
                    cid = (cinfo.get(k) or '').strip()
                    if cid and cid not in targets:
                        targets.append(cid)
            if not targets:
                continue
            groups_touched += 1
            msg_txt = _format_group_escalation_text(grp, vod_month)
            for cid in targets:
                ok, _err = _send_chunked_tg_message(token, cid, msg_txt)
                if ok:
                    sent += 1
                else:
                    failed += 1

        if groups_touched == 0:
            flash('Не знайдено жодної групи з налаштованим Chat ID старости/куратора (або всі студенти вже подали PDF).', 'warning')
        else:
            _record_event(
                conn, 'group_summary_sent',
                f"Надіслано зведення по групах ({groups_touched} гр.) за {_format_ym_ua(vod_month)}",
                severity='warning' if failed else 'info',
                details=f"Успішно чатів: {sent}, помилок: {failed}",
            )
            conn.commit()
            flash(
                f"Зведення надіслано для {groups_touched} груп(и)! Успішно: {sent}, помилок: {failed}.",
                'success' if failed == 0 else 'warning',
            )
    finally:
        conn.close()

    return redirect(url_for(
        'military_telegram.telegram_dashboard',
        tab='groups',
        gmonth=vod_month,
        gall_lic='1' if gall_lic else None,
    ))


@military_telegram_bp.route('/admin/military/telegram/bot/action', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_bot_action():
    act = (request.form.get('bot_action') or '').strip()
    conn = get_db()
    try:
        settings = _get_bot_settings(conn)
        token = (settings.get('bot_token') or os.environ.get('TELEGRAM_VOD_BOT_TOKEN') or '').strip()
    finally:
        conn.close()

    if act == 'test':
        if not token:
            flash('Токен бота ще не вказано.', 'warning')
        else:
            ok, me, err = _tg_call(token, 'getMe', timeout=10.0)
            if ok and isinstance(me, dict):
                c2 = get_db()
                try:
                    c2.execute(
                        "UPDATE telegram_bot_settings SET bot_username = ?, bot_first_name = ?, last_error = '' WHERE id = 1",
                        (me.get('username') or '', me.get('first_name') or ''),
                    )
                    c2.commit()
                finally:
                    c2.close()
                flash(f"З’єднання успішне! Бот: {me.get('first_name')} (@{me.get('username')}).", 'success')
            else:
                flash(f"Помилка підключення до Telegram API: {err}", 'error')
    elif act == 'start':
        ok, msg = start_bot_worker()
        flash(msg, 'success' if ok else 'error')
    elif act == 'stop':
        ok, msg = stop_bot_worker(disable_in_db=True)
        flash(msg, 'info')
    elif act == 'restart':
        stop_bot_worker(disable_in_db=False)
        ok, msg = start_bot_worker()
        flash(msg, 'success' if ok else 'error')

    return redirect(url_for('military_telegram.telegram_dashboard', tab='settings'))


@military_telegram_bp.route('/admin/military/telegram/events/mark-read', methods=['POST'])
@permission_required('manage_military_telegram')
def telegram_events_mark_read():
    conn = get_db()
    try:
        conn.execute("UPDATE telegram_events SET is_read = 1 WHERE is_read = 0")
        conn.commit()
        flash('Усі сповіщення позначено як прочитані.', 'success')
    finally:
        conn.close()
    return redirect(url_for('military_telegram.telegram_dashboard', tab=request.form.get('tab') or 'settings'))
