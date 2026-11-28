"""
routes/diploma_supplements.py
=============================
Центр підготовки, перевірки, імпорту та друку додатків до диплома:
  1. Preflight-аудит готовності групи та кожного студента перед друком;
  2. Реєстр претендентів на «Диплом з відзнакою» (з «пограничними» студентами,
     урахуванням/ігноруванням курсових і ручним перевизначенням);
  3. Масове введення/редагування оцінок за практики/курсові/атестації та
     двомовних тем кваліфікаційних робіт (UA + EN із збереженням у БД і
     пакетним автоперекладом через DeepL);
  4. Імпорт номерів диплома/додатка з Excel (ЄДЕБО) із перевіркою дублів і
     розбіжностей, а також швидка пакетна нумерація діапазоном;
  5. Конвертація .docx -> .pdf (LibreOffice / ONLYOFFICE ConvertService) та
     збірка єдиного PDF для друку (з опцією додавання порожнього звороту);
  6. Журнал генерації/друку/видачі з автодетектом «Потребує передруку» за
     SHA-256 підписом даних;
  7. Діагностика та тестовий рендер .docx-шаблонів на реальному студенту.
"""
from __future__ import annotations

import difflib
import hashlib
import io
import json
import os
import re
import shutil
import sqlite3
import subprocess
import tempfile
import threading
import time
import uuid
import zipfile
from datetime import datetime
from urllib.parse import urlparse

import jwt
import requests
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
from openpyxl import load_workbook
from werkzeug.utils import secure_filename

from routes.config import (
    DEEPL_API_KEY,
    ONLYOFFICE_CALLBACK_BASE_URL,
    ONLYOFFICE_COMMAND_URL,
    ONLYOFFICE_JWT_SECRET,
    ONLYOFFICE_PUBLIC_URL,
    ONLYOFFICE_VERIFY_SSL,
    PROJECT_ROOT,
)
from routes.db import get_db
from routes.helpers import current_username, sort_ukrainian
from routes.security import (
    is_safe_hex_token,
    resolve_under_dir,
    validate_upload_file,
)
from routes.user_scope import get_allowed_group_ids
from routes.utils import (
    TEMPLATE_FOLDER,
    apply_track_overrides,
    get_templates_with_metadata,
    is_student_on_reduced_program,
    log_action,
    logger,
    permission_required,
    program_track_condition,
)

diploma_bp = Blueprint('diploma', __name__)

_SCHEMA_LOCK = threading.Lock()
_SCHEMA_READY = False

EDEBO_PREVIEW_DIR = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'tmp', 'edebo_preview')
EDEBO_PREVIEW_TTL = 6 * 3600

DEFAULT_DIPLOMA_SETTINGS = {
    'filename_mask': '{last_name_UA}_{first_name_UA}',
    'honours_include_courseworks': '0',
    'pdf_duplex_pad': '0',
}

ALLOWED_MASK_KEYS = {
    'last_name_UA',
    'first_name_UA',
    'middle_name_UA',
    'last_name_ENG',
    'first_name_ENG',
    'group_name',
    'diploma_number',
    'appendix_number',
    'edebo_code',
    'id',
}


def _table_columns(conn, table_name: str) -> set[str]:
    try:
        return {r[1] for r in conn.execute(f"PRAGMA table_info({table_name})").fetchall()}
    except Exception:
        return set()


def ensure_diploma_supplement_schema(conn) -> None:
    """Ідемпотентно додає колонки та таблиці для центру додатків до диплома."""
    global _SCHEMA_READY
    with _SCHEMA_LOCK:
        # 1. diplomas: ручне перевизначення відзнаки
        dip_cols = _table_columns(conn, 'diplomas')
        if dip_cols:
            if 'honours_override' not in dip_cols:
                conn.execute("ALTER TABLE diplomas ADD COLUMN honours_override TEXT")
            if 'honours_note' not in dip_cols:
                conn.execute("ALTER TABLE diplomas ADD COLUMN honours_note TEXT")

        # 2. activity_grades: збережений EN-переклад теми
        ag_cols = _table_columns(conn, 'activity_grades')
        if ag_cols and 'name_en' not in ag_cols:
            conn.execute("ALTER TABLE activity_grades ADD COLUMN name_en TEXT")

        # 3. subjects / practices / courseworks / attestations: опційне name_en
        for tbl in ('subjects', 'practices', 'courseworks', 'attestations'):
            cols = _table_columns(conn, tbl)
            if cols and 'name_en' not in cols:
                conn.execute(f"ALTER TABLE {tbl} ADD COLUMN name_en TEXT")

        # 4. Налаштування додатків (ключ-значення)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS diploma_settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        for k, v in DEFAULT_DIPLOMA_SETTINGS.items():
            conn.execute(
                "INSERT OR IGNORE INTO diploma_settings (key, value) VALUES (?, ?)",
                (k, v),
            )

        # 5. Журнал генерації, друку та видачі додатків до диплома
        conn.execute("""
            CREATE TABLE IF NOT EXISTS diploma_generation_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                group_id INTEGER,
                template_name TEXT,
                diploma_number TEXT,
                appendix_number TEXT,
                honours_flag INTEGER NOT NULL DEFAULT 0,
                weighted_gpa REAL,
                total_credits REAL,
                data_signature TEXT,
                status TEXT NOT NULL DEFAULT 'generated',
                generated_by TEXT,
                generated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
                printed_at TEXT,
                printed_by TEXT,
                issued_at TEXT,
                issued_by TEXT,
                notes TEXT
            )
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_dgl_student_group
            ON diploma_generation_logs(student_id, group_id, id DESC)
        """)
        conn.commit()
        _SCHEMA_READY = True


def get_diploma_settings(conn) -> dict[str, str]:
    ensure_diploma_supplement_schema(conn)
    rows = conn.execute("SELECT key, value FROM diploma_settings").fetchall()
    out = dict(DEFAULT_DIPLOMA_SETTINGS)
    for r in rows:
        out[r['key']] = r['value'] if r['value'] is not None else DEFAULT_DIPLOMA_SETTINGS.get(r['key'], '')
    return out


def save_diploma_settings(conn, updates: dict[str, str]) -> None:
    ensure_diploma_supplement_schema(conn)
    for k, v in updates.items():
        if k not in DEFAULT_DIPLOMA_SETTINGS:
            continue
        conn.execute(
            """
            INSERT INTO diploma_settings (key, value) VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (k, str(v) if v is not None else ''),
        )
    conn.commit()


def format_custom_filename(student_dict: dict, diploma_row: dict | None, mask: str | None, ext: str = '.docx') -> str:
    """Формує безпечне ім'я файлу за маскою з налаштувань."""
    mask = (mask or '').strip() or DEFAULT_DIPLOMA_SETTINGS['filename_mask']
    dip_num = ''
    app_num = ''
    if diploma_row:
        dip_num = (diploma_row.get('diploma_number') or '').strip()
        app_num = (diploma_row.get('appendix_number') or '').strip()
        if dip_num:
            dip_num = dip_num.zfill(6)
        if app_num:
            app_num = app_num.zfill(6)

    raw_group = (student_dict.get('group_short_name') or student_dict.get('group_name') or '').strip()
    if '(' in raw_group:
        raw_group = raw_group.split('(', 1)[0].strip()

    mapping = {
        'id': str(student_dict.get('id') or ''),
        'last_name_UA': (student_dict.get('last_name_UA') or '').strip(),
        'first_name_UA': (student_dict.get('first_name_UA') or '').strip(),
        'middle_name_UA': (student_dict.get('middle_name_UA') or '').strip(),
        'last_name_ENG': (student_dict.get('last_name_ENG') or '').strip(),
        'first_name_ENG': (student_dict.get('first_name_ENG') or '').strip(),
        'group_name': raw_group,
        'diploma_number': dip_num,
        'appendix_number': app_num,
        'edebo_code': (student_dict.get('edebo_code') or '').strip(),
    }

    def _repl(match):
        key = match.group(1).strip()
        if key in ALLOWED_MASK_KEYS:
            return mapping.get(key, '')
        return ''

    rendered = re.sub(r'\{([a-zA-Z0-9_]+)\}', _repl, mask)
    rendered = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', '_', rendered)
    rendered = re.sub(r'\s+', '_', rendered)
    rendered = re.sub(r'_+', '_', rendered).strip('_. ')
    if not rendered:
        fallback = f"{mapping['last_name_UA']}_{mapping['first_name_UA']}".strip('_')
        rendered = fallback or f"student_{mapping['id'] or 'doc'}"
    if ext and not rendered.lower().endswith(ext.lower()):
        rendered += ext
    return rendered


def _grade_letter(grade) -> str | None:
    try:
        g = int(grade)
    except (TypeError, ValueError):
        return None
    if not 0 <= g <= 100:
        return None
    if g < 60:
        return 'Fx' if g >= 35 else 'F'
    if g <= 63:
        return 'E'
    if g <= 73:
        return 'D'
    if g <= 81:
        return 'C'
    if g <= 89:
        return 'B'
    return 'A'


def _to_slash_date(date_str: str | None) -> str:
    if not date_str:
        return ''
    s = str(date_str).strip()
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%d/%m/%Y')
        except ValueError:
            continue
    return s


def get_graduation_order_for_student(conn, student_id: int, group_id: int | None = None) -> dict[str, str]:
    """Шукає активний наказ про випуск ('graduation') для студента або групи."""
    try:
        row = conn.execute(
            """
            SELECT o.order_number, o.order_date
            FROM orders o
            JOIN order_items oi ON oi.order_id = o.id
            WHERE o.order_type = 'graduation'
              AND COALESCE(o.status, 'active') = 'active'
              AND (
                  (oi.entity_type = 'student' AND oi.entity_id = ?)
                  OR (oi.entity_type = 'group' AND oi.entity_id = ?)
              )
            ORDER BY CASE WHEN oi.entity_type = 'student' THEN 0 ELSE 1 END, o.id DESC
            LIMIT 1
            """,
            (student_id, group_id or -1),
        ).fetchone()
        if row:
            order_date = (row['order_date'] or '').strip()
            return {
                'graduation_order_number': (row['order_number'] or '').strip(),
                'graduation_order_date': order_date,
                'graduation_order_date_slash': _to_slash_date(order_date),
            }
    except Exception:
        pass
    return {
        'graduation_order_number': '',
        'graduation_order_date': '',
        'graduation_order_date_slash': '',
    }


def compute_student_diploma_metrics(
    conn,
    student_id: int,
    group_id: int | None,
    include_courseworks: bool | None = None,
    *,
    accreditation_pairs: set | None = None,
    dup_diploma_nums: set | None = None,
    dup_appendix_nums: set | None = None,
) -> dict:
    """
    Повний розрахунок даних для Додатка до диплома (Preflight-аудит, відзнака,
    середньозважений бал, кредити, підпис даних і статус журналу друку).
    """
    ensure_diploma_supplement_schema(conn)
    if include_courseworks is None:
        settings = get_diploma_settings(conn)
        include_courseworks = settings.get('honours_include_courseworks') == '1'

    student = conn.execute(
        """
        SELECT s.*,
               COALESCE(e.edebo_code, s.edebo_code) AS effective_edebo,
               COALESCE(e.license_id, s.license_id) AS effective_license_id,
               COALESCE(e.program_credits_override, s.program_credits_override) AS effective_credits_override,
               g.name AS group_short_name,
               g.start_year, g.study_form, g.program_credits,
               g.degree_level, g.degree_level_en,
               g.specialty, g.specialty_en, g.specialty_code,
               g.educational_program, g.educational_program_en,
               g.qualification_name, g.qualification_name_en,
               g.knowledge_area, g.knowledge_area_en,
               COALESCE(NULLIF(TRIM(g.faculty), ''), fl.name_ua, '') AS faculty,
               COALESCE(NULLIF(TRIM(g.faculty_en), ''), fl.name_en, '') AS faculty_en
        FROM students s
        LEFT JOIN groups g ON g.id = ?
        LEFT JOIN specialties sp ON sp.code = g.specialty_code
        LEFT JOIN faculties fl ON fl.id = COALESCE(g.faculty_id, sp.faculty_id)
        LEFT JOIN student_enrollments e ON e.student_id = s.id AND e.group_id = g.id
        WHERE s.id = ?
        """,
        (group_id, student_id),
    ).fetchone()

    if not student:
        return {}

    st = dict(student)
    is_reduced = is_student_on_reduced_program(conn, student_id, group_id) if group_id else False
    track_cond = program_track_condition(is_reduced)

    base_credits = st.get('program_credits') or 0
    ovr_credits = st.get('effective_credits_override')
    try:
        expected_credits = int(ovr_credits) if (is_reduced and ovr_credits) else int(base_credits or 0)
    except (TypeError, ValueError):
        expected_credits = 0

    # Освітні компоненти плану
    subjects = []
    practices = []
    courseworks = []
    attestations = []

    if group_id:
        subjects = [dict(r) for r in conn.execute(
            f"""
            SELECT s.id, s.code, s.name, s.name_en, s.credits, s.reduced_credits,
                   s.type, s.reduced_type, s.position, g.grade
            FROM subjects s
            LEFT JOIN grades g ON g.subject_id = s.id AND g.student_id = ?
            WHERE s.group_id = ?{track_cond}
            ORDER BY s.position
            """,
            (student_id, group_id),
        ).fetchall()]
        for item in subjects:
            apply_track_overrides(item, is_reduced)

        practices = [dict(r) for r in conn.execute(
            f"""
            SELECT p.id, p.code, p.name, p.name_en, p.credits, p.reduced_credits,
                   p.type, p.reduced_type, p.position, ag.grade
            FROM practices p
            LEFT JOIN activity_grades ag
                   ON ag.entity_id = p.id AND ag.entity_type = 'practice' AND ag.student_id = ?
            WHERE p.group_id = ?{track_cond}
            ORDER BY p.position
            """,
            (student_id, group_id),
        ).fetchall()]
        for item in practices:
            apply_track_overrides(item, is_reduced)

        courseworks = [dict(r) for r in conn.execute(
            f"""
            SELECT c.id, c.code, c.name, c.name_en, c.credits, c.reduced_credits,
                   c.type, c.reduced_type, c.position, ag.grade
            FROM courseworks c
            LEFT JOIN activity_grades ag
                   ON ag.entity_id = c.id AND ag.entity_type = 'coursework' AND ag.student_id = ?
            WHERE c.group_id = ?{track_cond}
            ORDER BY c.position
            """,
            (student_id, group_id),
        ).fetchall()]
        for item in courseworks:
            apply_track_overrides(item, is_reduced)

        attestations = [dict(r) for r in conn.execute(
            f"""
            SELECT a.id, a.code, a.name, a.name_en AS att_name_en, a.credits, a.reduced_credits,
                   a.type, a.reduced_type, a.position,
                   ag.grade, ag.name AS student_name, ag.name_en AS topic_en
            FROM attestations a
            LEFT JOIN activity_grades ag
                   ON ag.entity_id = a.id AND ag.entity_type = 'attestation' AND ag.student_id = ?
            WHERE a.group_id = ?{track_cond}
            ORDER BY a.position
            """,
            (student_id, group_id),
        ).fetchall()]
        for item in attestations:
            apply_track_overrides(item, is_reduced)

    def _safe_float(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    def _safe_grade_int(v):
        if v is None or str(v).strip() == '':
            return None
        try:
            g = int(str(v).strip())
            return g if 0 <= g <= 100 else None
        except (TypeError, ValueError):
            return None

    actual_credits_raw = sum(
        _safe_float(x.get('credits'))
        for grp in (subjects, practices, courseworks, attestations)
        for x in grp
    )
    actual_credits = round(actual_credits_raw, 2)
    if actual_credits.is_integer():
        actual_credits = int(actual_credits)

    credits_match = (not expected_credits) or abs(float(actual_credits) - float(expected_credits)) < 0.01
    credits_diff = round(float(actual_credits) - float(expected_credits), 2) if expected_credits else 0

    # Оцінки, середній бал, середньозважений бал
    numeric_all = []
    weighted_num = 0.0
    weighted_den = 0.0

    for item in subjects + practices + courseworks + attestations:
        g_int = _safe_grade_int(item.get('grade'))
        cr = _safe_float(item.get('credits'))
        if g_int is not None:
            numeric_all.append(g_int)
            if cr > 0:
                weighted_num += g_int * cr
                weighted_den += cr

    gpa_100 = round(sum(numeric_all) / len(numeric_all), 2) if numeric_all else None
    weighted_gpa_100 = round(weighted_num / weighted_den, 2) if weighted_den > 0 else gpa_100
    gpa_5 = None
    if weighted_gpa_100 is not None:
        # Стандартне перетворення за 100-бальною у 5-бальну шкалу
        gpa_5 = round(weighted_gpa_100 / 20.0, 2)

    # Розрахунок права на диплом з відзнакою
    honours_pool = list(subjects) + list(practices)
    if include_courseworks:
        honours_pool += list(courseworks)

    graded_count = 0
    excellent_count = 0
    de_count = 0
    fail_count = 0
    for item in honours_pool:
        g_int = _safe_grade_int(item.get('grade'))
        if g_int is None:
            continue
        graded_count += 1
        letter = _grade_letter(g_int)
        if letter == 'A':
            excellent_count += 1
        elif letter in ('D', 'E'):
            de_count += 1
        elif letter in ('Fx', 'F'):
            fail_count += 1

    excellent_pct = round(excellent_count / graded_count * 100, 1) if graded_count else 0.0

    att_grades = [_safe_grade_int(a.get('grade')) for a in attestations]
    att_grades_present = [g for g in att_grades if g is not None]
    attestation_is_a = (
        bool(att_grades_present) and all(_grade_letter(g) == 'A' for g in att_grades_present)
    ) if attestations else True

    auto_honours = (
        graded_count > 0
        and (excellent_count / graded_count) >= 0.75
        and de_count == 0
        and fail_count == 0
        and attestation_is_a
    )

    # Запис диплома
    dip_row = conn.execute(
        """
        SELECT diploma_number, appendix_number, honours_override, honours_note
        FROM diplomas
        WHERE student_id = ? AND group_id = ?
        ORDER BY id DESC LIMIT 1
        """,
        (student_id, group_id),
    ).fetchone()
    diploma_number = (dip_row['diploma_number'] or '').strip() if dip_row else ''
    appendix_number = (dip_row['appendix_number'] or '').strip() if dip_row else ''
    honours_override = (dip_row['honours_override'] or '').strip() if dip_row else ''
    honours_note = (dip_row['honours_note'] or '').strip() if dip_row else ''
    if diploma_number:
        diploma_number = diploma_number.zfill(6)
    if appendix_number:
        appendix_number = appendix_number.zfill(6)

    if honours_override == 'force_yes':
        final_honours = True
        honours_category = 'override_yes'
        honours_reason = 'Підтверджено вручну' + (f' ({honours_note})' if honours_note else '')
    elif honours_override == 'force_no':
        final_honours = False
        honours_category = 'override_no'
        honours_reason = 'Знято вручну' + (f' ({honours_note})' if honours_note else '')
    elif auto_honours:
        final_honours = True
        honours_category = 'eligible'
        honours_reason = f"{excellent_pct}% «A», без D/E/Fx" + (", атестація «A»" if attestations else "")
    else:
        final_honours = False
        # Перевіряємо «на межі» (пограничний випадок)
        reasons = []
        if graded_count == 0:
            honours_category = 'blocked'
            honours_reason = 'Немає оцінок'
        else:
            if fail_count > 0:
                reasons.append(f'незарах./Fx/F: {fail_count}')
            if de_count > 0:
                reasons.append(f'оцінок D/E: {de_count}')
            if attestations and att_grades_present and not attestation_is_a:
                reasons.append('атестація не «A»')
            if excellent_pct < 75.0:
                needed = max(1, int((0.75 * graded_count) - excellent_count + 0.9999))
                reasons.append(f'{excellent_pct}% «A» (до 75% не вистачає {needed})')
            if fail_count == 0 and de_count <= 1 and excellent_pct >= 68.0:
                honours_category = 'borderline'
            else:
                honours_category = 'blocked'
            honours_reason = '; '.join(reasons) if reasons else 'Недостатньо «A»'

    # Перевірка комплектності та прогалин для Додатка до диплома
    preflight_issues = []

    if not (st.get('last_name_ENG') or '').strip() or not (st.get('first_name_ENG') or '').strip():
        preflight_issues.append({'key': 'name_en', 'severity': 'error', 'label': 'Немає ПІБ латиницею (ENG)'})
    if not (st.get('birth_date') or '').strip():
        preflight_issues.append({'key': 'birth_date', 'severity': 'error', 'label': 'Немає дати народження'})
    if not (st.get('effective_edebo') or '').strip():
        preflight_issues.append({'key': 'edebo', 'severity': 'warning', 'label': 'Немає коду ЄДЕБО'})

    if group_id:
        missing_group_en = []
        if not (st.get('educational_program_en') or '').strip():
            missing_group_en.append('ОП (EN)')
        if not (st.get('qualification_name_en') or '').strip():
            missing_group_en.append('Кваліфікація (EN)')
        if not (st.get('faculty_en') or '').strip():
            missing_group_en.append('Факультет (EN)')
        if missing_group_en:
            preflight_issues.append({
                'key': 'group_en',
                'severity': 'error',
                'label': 'У групи не заповнено: ' + ', '.join(missing_group_en),
            })

        if accreditation_pairs is None:
            acc = conn.execute(
                "SELECT 1 FROM accreditations WHERE degree = ? AND specialty = ? LIMIT 1",
                (st.get('degree_level') or '', st.get('specialty') or ''),
            ).fetchone()
            has_acc = acc is not None
        else:
            has_acc = (st.get('degree_level') or '', st.get('specialty') or '') in accreditation_pairs
        if not has_acc:
            preflight_issues.append({
                'key': 'accreditation',
                'severity': 'error',
                'label': f"Немає акредитації ({st.get('degree_level') or '?'} / {st.get('specialty') or '?'})",
            })

    # Попередня освіта
    edu_doc = conn.execute(
        """
        SELECT document_type, document_number, institution_name, country,
               document_type_en, institution_name_en, country_en
        FROM education_documents
        WHERE student_id = ? AND (group_id = ? OR group_id IS NULL)
        ORDER BY (group_id = ?) DESC, id DESC LIMIT 1
        """,
        (student_id, group_id, group_id),
    ).fetchone()
    if not edu_doc:
        preflight_issues.append({'key': 'edu_doc', 'severity': 'error', 'label': 'Немає документа про попередню освіту'})
    else:
        missing_edu_en = []
        if not (edu_doc['document_number'] or '').strip():
            missing_edu_en.append('номер')
        if not (edu_doc['document_type_en'] or '').strip():
            missing_edu_en.append('тип (EN)')
        if not (edu_doc['institution_name_en'] or '').strip():
            missing_edu_en.append('заклад (EN)')
        if not (edu_doc['country_en'] or '').strip():
            missing_edu_en.append('країна (EN)')
        if missing_edu_en:
            preflight_issues.append({
                'key': 'edu_doc_en',
                'severity': 'error',
                'label': 'Попередня освіта — не заповнено: ' + ', '.join(missing_edu_en),
            })

    # Періоди навчання / ліцензія
    has_periods = conn.execute(
        "SELECT 1 FROM student_study_periods WHERE student_id = ? AND (group_id = ? OR group_id IS NULL) LIMIT 1",
        (student_id, group_id),
    ).fetchone() is not None
    if not has_periods and not st.get('effective_license_id'):
        preflight_issues.append({
            'key': 'study_period',
            'severity': 'warning',
            'label': 'Не задано період навчання та ліцензію',
        })

    # Номери диплома/додатка
    if not diploma_number:
        preflight_issues.append({'key': 'diploma_number', 'severity': 'error', 'label': 'Не вказано номер диплома'})
    elif dup_diploma_nums and diploma_number in dup_diploma_nums:
        preflight_issues.append({'key': 'diploma_dup', 'severity': 'error', 'label': f'Дублікат № диплома ({diploma_number})'})

    if not appendix_number:
        preflight_issues.append({'key': 'appendix_number', 'severity': 'error', 'label': 'Не вказано номер додатка'})
    elif dup_appendix_nums and appendix_number in dup_appendix_nums:
        preflight_issues.append({'key': 'appendix_dup', 'severity': 'error', 'label': f'Дублікат № додатка ({appendix_number})'})

    # Оцінки предметів і активностей
    subjects_filled = sum(1 for x in subjects if _safe_grade_int(x.get('grade')) is not None)
    practices_filled = sum(1 for x in practices if _safe_grade_int(x.get('grade')) is not None)
    courseworks_filled = sum(1 for x in courseworks if _safe_grade_int(x.get('grade')) is not None)
    attestations_filled = sum(1 for x in attestations if _safe_grade_int(x.get('grade')) is not None)

    if subjects and subjects_filled < len(subjects):
        preflight_issues.append({
            'key': 'subjects_grades',
            'severity': 'error',
            'label': f'Оцінки предметів: {subjects_filled}/{len(subjects)}',
        })
    if practices and practices_filled < len(practices):
        preflight_issues.append({
            'key': 'practices_grades',
            'severity': 'error',
            'label': f'Оцінки практик: {practices_filled}/{len(practices)}',
        })
    if courseworks and courseworks_filled < len(courseworks):
        preflight_issues.append({
            'key': 'courseworks_grades',
            'severity': 'error',
            'label': f'Оцінки курсових: {courseworks_filled}/{len(courseworks)}',
        })
    if attestations and attestations_filled < len(attestations):
        preflight_issues.append({
            'key': 'attestations_grades',
            'severity': 'error',
            'label': f'Оцінки атестації: {attestations_filled}/{len(attestations)}',
        })

    # Теми атестацій (UA і EN)
    missing_topics_ua = sum(
        1 for a in attestations
        if 'робот' in (a.get('name') or '').lower() and not (a.get('student_name') or '').strip()
    )
    missing_topics_en = sum(
        1 for a in attestations
        if (a.get('student_name') or '').strip() and not (a.get('topic_en') or '').strip()
    )
    if missing_topics_ua:
        preflight_issues.append({
            'key': 'topic_ua',
            'severity': 'warning',
            'label': 'Не вказано тему кваліфікаційної роботи',
        })
    if missing_topics_en:
        preflight_issues.append({
            'key': 'topic_en',
            'severity': 'warning',
            'label': 'Немає збереженого EN-перекладу теми (працюватиме DeepL на льоту)',
        })

    if not credits_match:
        preflight_issues.append({
            'key': 'credits_mismatch',
            'severity': 'warning',
            'label': f'Кредити плану: {actual_credits} з {expected_credits} ECTS',
        })

    error_count = sum(1 for i in preflight_issues if i['severity'] == 'error')
    warning_count = sum(1 for i in preflight_issues if i['severity'] == 'warning')
    ready_status = 'ready' if error_count == 0 and warning_count == 0 else ('warn' if error_count == 0 else 'blocked')

    grad_order = get_graduation_order_for_student(conn, student_id, group_id)

    # Підпис даних (для виявлення «Потребує передруку»)
    sig_payload = {
        'fio_ua': [st.get('last_name_UA') or '', st.get('first_name_UA') or '', st.get('middle_name_UA') or ''],
        'fio_en': [st.get('last_name_ENG') or '', st.get('first_name_ENG') or ''],
        'birth': st.get('birth_date') or '',
        'edebo': st.get('effective_edebo') or '',
        'dip': diploma_number,
        'app': appendix_number,
        'honours': final_honours,
        'subjects': [(x['id'], x.get('grade'), x.get('credits')) for x in subjects],
        'practices': [(x['id'], x.get('grade'), x.get('credits')) for x in practices],
        'courseworks': [(x['id'], x.get('grade'), x.get('credits')) for x in courseworks],
        'attestations': [
            (x['id'], x.get('grade'), x.get('student_name') or '', x.get('topic_en') or '')
            for x in attestations
        ],
    }
    data_signature = hashlib.sha256(
        json.dumps(sig_payload, ensure_ascii=False, sort_keys=True).encode('utf-8')
    ).hexdigest()[:24]

    last_log_row = conn.execute(
        """
        SELECT id, status, generated_by, generated_at, printed_at, issued_at, data_signature, template_name
        FROM diploma_generation_logs
        WHERE student_id = ? AND COALESCE(group_id, -1) = COALESCE(?, -1)
        ORDER BY id DESC LIMIT 1
        """,
        (student_id, group_id),
    ).fetchone()
    last_log = dict(last_log_row) if last_log_row else None
    needs_reprint = bool(last_log and last_log.get('data_signature') and last_log['data_signature'] != data_signature)

    total_ects_credits = actual_credits or expected_credits
    try:
        total_hours = int(round(float(total_ects_credits) * 30)) if total_ects_credits else 0
    except (TypeError, ValueError):
        total_hours = 0

    return {
        'student_id': student_id,
        'group_id': group_id,
        'full_name_ua': f"{st.get('last_name_UA') or ''} {st.get('first_name_UA') or ''} {st.get('middle_name_UA') or ''}".strip(),
        'full_name_en': f"{st.get('last_name_ENG') or ''} {st.get('first_name_ENG') or ''}".strip(),
        'edebo_code': st.get('effective_edebo') or '',
        'is_reduced': is_reduced,
        'expected_credits': expected_credits,
        'actual_credits': actual_credits,
        'credits_match': credits_match,
        'credits_diff': credits_diff,
        'total_ects_credits': total_ects_credits,
        'total_hours': total_hours,
        'subjects_total': len(subjects),
        'subjects_filled': subjects_filled,
        'practices_total': len(practices),
        'practices_filled': practices_filled,
        'courseworks_total': len(courseworks),
        'courseworks_filled': courseworks_filled,
        'attestations_total': len(attestations),
        'attestations_filled': attestations_filled,
        'gpa_100': gpa_100,
        'weighted_gpa_100': weighted_gpa_100,
        'gpa_5': gpa_5,
        'graded_count': graded_count,
        'excellent_count': excellent_count,
        'excellent_pct': excellent_pct,
        'de_count': de_count,
        'fail_count': fail_count,
        'attestation_is_a': attestation_is_a,
        'auto_honours': auto_honours,
        'honours_override': honours_override,
        'honours_note': honours_note,
        'final_honours': final_honours,
        'honours_category': honours_category,
        'honours_reason': honours_reason,
        'diploma_number': diploma_number,
        'appendix_number': appendix_number,
        'preflight_issues': preflight_issues,
        'error_count': error_count,
        'warning_count': warning_count,
        'ready_status': ready_status,
        'data_signature': data_signature,
        'last_log': last_log,
        'needs_reprint': needs_reprint,
        **grad_order,
    }


def record_diploma_generation(
    conn,
    student_id: int,
    group_id: int | None,
    template_name: str,
    username: str,
    metrics: dict | None = None,
) -> int | None:
    """Фіксує факт генерації додатка до диплома у журналі."""
    ensure_diploma_supplement_schema(conn)
    if metrics is None:
        metrics = compute_student_diploma_metrics(conn, student_id, group_id)
    if not metrics:
        return None
    cur = conn.execute(
        """
        INSERT INTO diploma_generation_logs (
            student_id, group_id, template_name, diploma_number, appendix_number,
            honours_flag, weighted_gpa, total_credits, data_signature, status, generated_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'generated', ?)
        """,
        (
            student_id,
            group_id,
            os.path.basename(template_name or ''),
            metrics.get('diploma_number') or '',
            metrics.get('appendix_number') or '',
            1 if metrics.get('final_honours') else 0,
            metrics.get('weighted_gpa_100'),
            metrics.get('total_ects_credits'),
            metrics.get('data_signature') or '',
            username or 'Система',
        ),
    )
    conn.commit()
    return cur.lastrowid


# ============================================================
# Конвертація DOCX -> PDF та об'єднання PDF (PyMuPDF)
# ============================================================

def _find_soffice_executable() -> str | None:
    for name in ('soffice', 'libreoffice'):
        found = shutil.which(name)
        if found:
            return found
    win_candidates = [
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    ]
    for p in win_candidates:
        if os.path.isfile(p):
            return p
    return None


def _derive_convert_service_url() -> str | None:
    if ONLYOFFICE_COMMAND_URL:
        parsed = urlparse(ONLYOFFICE_COMMAND_URL.strip())
        if parsed.scheme in ('http', 'https') and parsed.netloc:
            return f"{parsed.scheme}://{parsed.netloc}/ConvertService.ashx"
    if ONLYOFFICE_PUBLIC_URL:
        base = ONLYOFFICE_PUBLIC_URL.strip().rstrip('/')
        if base.startswith(('http://', 'https://')):
            return f"{base}/ConvertService.ashx"
    return None


def _is_safe_onlyoffice_result_url(url: str) -> bool:
    if not url or not isinstance(url, str):
        return False
    parsed = urlparse(url.strip())
    if parsed.scheme not in ('http', 'https') or not parsed.hostname:
        return False
    host = parsed.hostname.lower()
    port = parsed.port or (443 if parsed.scheme == 'https' else 80)
    loopback = {'localhost', '127.0.0.1', '::1'}
    for cand in (ONLYOFFICE_PUBLIC_URL, ONLYOFFICE_COMMAND_URL):
        if not cand:
            continue
        cp = urlparse(cand.strip())
        if cp.scheme not in ('http', 'https') or not cp.hostname:
            continue
        cport = cp.port or (443 if cp.scheme == 'https' else 80)
        if parsed.scheme == cp.scheme and port == cport:
            if host == cp.hostname.lower() or (host in loopback and cp.hostname.lower() in loopback):
                return True
    return False


def convert_docx_to_pdf(docx_path: str, out_pdf_path: str, doc_id_for_onlyoffice: str | None = None) -> tuple[bool, str]:
    """
    Конвертує .docx у .pdf:
      1) Через локальний LibreOffice (`soffice --headless --convert-to pdf`), якщо встановлений.
      2) Інакше (якщо передано doc_id у ONLYOFFICE) — через ONLYOFFICE ConvertService.ashx.
    """
    if not os.path.isfile(docx_path):
        return False, 'Вихідний файл .docx не знайдено'

    soffice = _find_soffice_executable()
    if soffice:
        tmp_dir = tempfile.mkdtemp(prefix='lo_pdf_')
        try:
            proc = subprocess.run(
                [soffice, '--headless', '--norestore', '--convert-to', 'pdf', '--outdir', tmp_dir, docx_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=90,
                check=False,
            )
            stem = os.path.splitext(os.path.basename(docx_path))[0]
            produced = os.path.join(tmp_dir, f'{stem}.pdf')
            if proc.returncode == 0 and os.path.isfile(produced):
                shutil.copyfile(produced, out_pdf_path)
                return True, ''
            err_msg = (proc.stderr or proc.stdout or b'').decode('utf-8', errors='replace').strip()
            logger.warning(f"LibreOffice не зміг конвертувати {docx_path}: {err_msg}")
        except Exception as e:
            logger.warning(f"Помилка LibreOffice при конвертації {docx_path}: {e}")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    # Fallback: ONLYOFFICE Document Server ConvertService
    convert_url = _derive_convert_service_url()
    if convert_url and doc_id_for_onlyoffice and ONLYOFFICE_CALLBACK_BASE_URL:
        file_url = f"{ONLYOFFICE_CALLBACK_BASE_URL.rstrip('/')}/office/file/{doc_id_for_onlyoffice}"
        payload = {
            'async': False,
            'filetype': 'docx',
            'key': uuid.uuid4().hex,
            'outputtype': 'pdf',
            'title': os.path.basename(docx_path),
            'url': file_url,
        }
        if ONLYOFFICE_JWT_SECRET:
            payload['token'] = jwt.encode(payload, ONLYOFFICE_JWT_SECRET, algorithm='HS256')
        try:
            resp = requests.post(
                convert_url,
                json=payload,
                headers={'Accept': 'application/json'},
                timeout=60,
                verify=ONLYOFFICE_VERIFY_SSL,
            )
            resp.raise_for_status()
            data = resp.json()
            if data.get('endConvert') and data.get('fileUrl'):
                pdf_url = data['fileUrl']
                if _is_safe_onlyoffice_result_url(pdf_url):
                    pdf_resp = requests.get(
                        pdf_url, timeout=60, verify=ONLYOFFICE_VERIFY_SSL, allow_redirects=False
                    )
                    pdf_resp.raise_for_status()
                    with open(out_pdf_path, 'wb') as f:
                        f.write(pdf_resp.content)
                    return True, ''
            if data.get('error'):
                return False, f"ONLYOFFICE ConvertService error={data.get('error')}"
        except Exception as e:
            logger.warning(f"ONLYOFFICE ConvertService недоступний ({convert_url}): {e}")

    return False, (
        'Для експорту в PDF потрібен LibreOffice (soffice) на сервері або запущений ONLYOFFICE ConvertService.'
    )


def merge_pdfs_with_fitz(pdf_paths: list[str], out_pdf_path: str, duplex_pad: bool = False) -> tuple[bool, str]:
    """Об'єднує список PDF-файлів в один через PyMuPDF (fitz)."""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return False, 'На сервері не встановлено PyMuPDF (fitz) для об’єднання PDF.'

    merged = fitz.open()
    try:
        for idx, p in enumerate(pdf_paths):
            if not os.path.isfile(p):
                continue
            doc = fitz.open(p)
            try:
                page_count = doc.page_count
                if page_count == 0:
                    continue
                merged.insert_pdf(doc)
                if duplex_pad and (page_count % 2 == 1) and idx < len(pdf_paths) - 1:
                    last_rect = doc[-1].rect
                    merged.new_page(width=last_rect.width, height=last_rect.height)
            finally:
                doc.close()
        if merged.page_count == 0:
            return False, 'Не знайдено жодної сторінки для збірки PDF.'
        merged.save(out_pdf_path)
        return True, ''
    finally:
        merged.close()


# ============================================================
# Діагностика та тест-рендер шаблонів (.docx)
# ============================================================

def inspect_docx_template(template_path: str, context_dict: dict) -> dict:
    """
    Аналізує .docx-шаблон на наявність розірваних або невідомих Jinja-тегів
    і порівнює знайдені змінні з реальним контекстом студента.
    """
    from docxtpl import DocxTemplate
    import jinja2
    import xml.etree.ElementTree as ET

    result = {
        'syntax_ok': True,
        'syntax_error': '',
        'undeclared_vars': [],
        'unknown_vars': [],
        'empty_vars': [],
        'filled_vars': [],
        'broken_tag_warnings': [],
    }

    # 1. Перевірка XML всередині .docx на незбалансовані дужки / розірвані runs
    try:
        with zipfile.ZipFile(template_path, 'r') as zf:
            xml_names = [
                n for n in zf.namelist()
                if n.startswith('word/') and n.endswith('.xml')
                and ('document' in n or 'header' in n or 'footer' in n)
            ]
            for xname in xml_names:
                raw_xml = zf.read(xname)
                try:
                    root = ET.fromstring(raw_xml)
                except ET.ParseError as pe:
                    result['broken_tag_warnings'].append(f"{xname}: пошкоджений XML ({pe})")
                    continue
                for p_idx, p in enumerate(root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'), 1):
                    texts = [t.text or '' for t in p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')]
                    joined = ''.join(texts)
                    if '{' not in joined and '}' not in joined:
                        continue
                    if joined.count('{{') != joined.count('}}') or joined.count('{%') != joined.count('%}'):
                        snippet = joined.strip()[:90]
                        result['broken_tag_warnings'].append(
                            f"{os.path.basename(xname)} (абзац ~{p_idx}): незбалансовані теги «{snippet}»"
                        )
    except Exception as e:
        result['broken_tag_warnings'].append(f"Попередження читання XML: {e}")

    # 2. Витягування змінних через DocxTemplate
    try:
        tpl = DocxTemplate(template_path)
        jinja_env = jinja2.Environment()
        undeclared = sorted(tpl.get_undeclared_template_variables(jinja_env=jinja_env))
        result['undeclared_vars'] = undeclared
    except Exception as e:
        result['syntax_ok'] = False
        result['syntax_error'] = str(e)
        return result

    for var in undeclared:
        if var not in context_dict:
            result['unknown_vars'].append(var)
        else:
            val = context_dict.get(var)
            is_empty = (val is None or val == '' or val == [] or val == {})
            if is_empty:
                result['empty_vars'].append(var)
            else:
                if isinstance(val, list):
                    preview = f"список ({len(val)} елемент.)"
                else:
                    s = str(val).replace('\n', ' ').strip()
                    preview = s if len(s) <= 80 else s[:77] + '...'
                result['filled_vars'].append({'name': var, 'preview': preview})

    return result


# ============================================================
# Автопереклад (DeepL) з кешем у БД
# ============================================================

def _translate_text_deepl(text: str) -> str:
    from routes.admin.documents import translate_to_en
    return translate_to_en(text) or ''


# ============================================================
# Імпорт номерів диплома/додатка з Excel (ЄДЕБО)
# ============================================================

def _cleanup_old_edebo_previews() -> None:
    if not os.path.isdir(EDEBO_PREVIEW_DIR):
        return
    now = time.time()
    for f in os.listdir(EDEBO_PREVIEW_DIR):
        p = os.path.join(EDEBO_PREVIEW_DIR, f)
        try:
            if os.path.isfile(p) and now - os.path.getmtime(p) > EDEBO_PREVIEW_TTL:
                os.remove(p)
        except OSError:
            pass


def _norm_header(h) -> str:
    s = str(h or '').strip().lower()
    s = re.sub(r'\s+', ' ', s)
    return s


def _norm_fio(s) -> str:
    s = str(s or '').lower()
    s = s.replace('ʼ', "'").replace('’', "'")
    s = re.sub(r'\(.*?\)', ' ', s)
    s = re.sub(r"[^\w\s'-]", ' ', s, flags=re.UNICODE)
    return ' '.join(s.split())


def _detect_edebo_columns(headers: list[str]) -> dict[str, int | None]:
    """Автоматично визначає індекси колонок у таблиці ЄДЕБО або довільному Excel."""
    mapping = {
        'edebo_code': None,
        'fio_ua': None,
        'last_name': None,
        'first_name': None,
        'middle_name': None,
        'fio_en': None,
        'birth_date': None,
        'diploma_number': None,
        'appendix_number': None,
    }
    for idx, raw in enumerate(headers):
        h = _norm_header(raw)
        if not h:
            continue
        if mapping['edebo_code'] is None and any(k in h for k in ('єдебо', 'edebo', 'код картки', 'ідентифікатор')):
            mapping['edebo_code'] = idx
        elif mapping['appendix_number'] is None and any(k in h for k in ('додат', 'appendix', 'supplement')):
            mapping['appendix_number'] = idx
        elif mapping['diploma_number'] is None and ('диплом' in h or 'diploma' in h or 'реєстраційний номер' in h):
            mapping['diploma_number'] = idx
        elif mapping['fio_en'] is None and any(k in h for k in ('англ', 'latin', 'eng', 'english')):
            mapping['fio_en'] = idx
        elif mapping['birth_date'] is None and any(k in h for k in ('народж', 'birth')):
            mapping['birth_date'] = idx
        elif mapping['last_name'] is None and h in ('прізвище', 'призвище'):
            mapping['last_name'] = idx
        elif mapping['first_name'] is None and h in ("ім'я", 'імя', 'ім’я'):
            mapping['first_name'] = idx
        elif mapping['middle_name'] is None and ('по батькові' in h or 'по-батькові' in h):
            mapping['middle_name'] = idx
        elif mapping['fio_ua'] is None and any(k in h for k in ('піб', 'фіо', 'здобувач', 'студент', 'прізвище')):
            mapping['fio_ua'] = idx
    return mapping


def _clean_doc_num(val) -> str:
    if val is None:
        return ''
    if isinstance(val, float) and val.is_integer():
        s = str(int(val))
    else:
        s = str(val).strip()
    s = re.sub(r'\s+', '', s)
    if s.isdigit():
        s = s.zfill(6)
    return s


# ============================================================
# Маршрути Центру додатків до диплома
# ============================================================

def _load_accessible_groups(conn, include_archived: bool = False):
    is_admin = bool(session.get('is_admin', False)) or session.get('role') == 'admin'
    archived_cond = "" if include_archived else "WHERE COALESCE(g.archived, 0) = 0"
    if not is_admin:
        gids = get_allowed_group_ids(conn, session.get('user_id')) or []
        if not gids:
            return []
        ph = ','.join('?' for _ in gids)
        extra = f"g.id IN ({ph})"
        if archived_cond:
            archived_cond += f" AND {extra}"
        else:
            archived_cond = f"WHERE {extra}"
        rows = conn.execute(
            f"""
            SELECT g.id, g.name, g.start_year, g.study_form, g.program_credits,
                   COALESCE(g.archived, 0) AS archived,
                   g.degree_level, g.specialty
            FROM groups g
            {archived_cond}
            ORDER BY COALESCE(g.archived, 0), g.name COLLATE UKRAINIAN, g.start_year DESC
            """,
            gids,
        ).fetchall()
    else:
        rows = conn.execute(
            f"""
            SELECT g.id, g.name, g.start_year, g.study_form, g.program_credits,
                   COALESCE(g.archived, 0) AS archived,
                   g.degree_level, g.specialty
            FROM groups g
            {archived_cond}
            ORDER BY COALESCE(g.archived, 0), g.name COLLATE UKRAINIAN, g.start_year DESC
            """
        ).fetchall()
    return [dict(r) for r in rows]


def _get_group_students_for_diploma(conn, group_id: int):
    rows = conn.execute(
        """
        SELECT DISTINCT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.last_name_ENG, s.first_name_ENG, s.birth_date,
               COALESCE(e.edebo_code, s.edebo_code) AS edebo_code,
               e.status AS enrollment_status
        FROM students s
        LEFT JOIN student_enrollments e
               ON e.student_id = s.id AND e.group_id = ?
        WHERE (e.id IS NOT NULL OR s.group_id = ?)
          AND COALESCE(e.status, 'active') IN ('active', 'graduated')
        """,
        (group_id, group_id),
    ).fetchall()
    items = [dict(r) for r in rows]
    return sort_ukrainian(
        items,
        key_func=lambda s: f"{s['last_name_UA']} {s['first_name_UA']} {s['middle_name_UA'] or ''}",
    )


@diploma_bp.route('/admin/diploma-supplements', methods=['GET'])
@permission_required('manage_diplomas')
def hub():
    """Головна сторінка Центру підготовки додатків до диплома."""
    conn = get_db()
    ensure_diploma_supplement_schema(conn)

    include_archived = request.args.get('include_archived') == '1'
    groups = _load_accessible_groups(conn, include_archived=include_archived)
    selected_group_id = request.args.get('group_id', type=int)
    active_tab = (request.args.get('tab') or 'audit').strip()
    if active_tab not in ('audit', 'activities', 'import', 'logs'):
        active_tab = 'audit'

    if selected_group_id and not any(g['id'] == selected_group_id for g in groups):
        # Перевіримо, чи це архівна група і include_archived був вимкнений
        if not include_archived:
            groups_all = _load_accessible_groups(conn, include_archived=True)
            if any(g['id'] == selected_group_id for g in groups_all):
                groups = groups_all
                include_archived = True
            else:
                selected_group_id = None
        else:
            selected_group_id = None

    selected_group = next((g for g in groups if g['id'] == selected_group_id), None)
    settings = get_diploma_settings(conn)
    include_cw = settings.get('honours_include_courseworks') == '1'

    student_rows = []
    kpi = {
        'total': 0,
        'ready': 0,
        'warn': 0,
        'blocked': 0,
        'honours': 0,
        'borderline': 0,
        'needs_reprint': 0,
    }
    group_practices = []
    group_courseworks = []
    group_attestations = []
    group_subjects = []
    activity_values = {}  # (student_id, entity_type, entity_id) -> {'grade', 'name', 'name_en'}
    logs_list = []

    if selected_group_id and selected_group:
        accreditation_pairs = {
            (r['degree'], r['specialty'])
            for r in conn.execute("SELECT degree, specialty FROM accreditations").fetchall()
        }
        dup_diploma_nums = {
            r['diploma_number']
            for r in conn.execute(
                "SELECT diploma_number FROM diplomas WHERE diploma_number IS NOT NULL AND diploma_number != '' GROUP BY diploma_number HAVING COUNT(*) > 1"
            ).fetchall()
        }
        dup_appendix_nums = {
            r['appendix_number']
            for r in conn.execute(
                "SELECT appendix_number FROM diplomas WHERE appendix_number IS NOT NULL AND appendix_number != '' GROUP BY appendix_number HAVING COUNT(*) > 1"
            ).fetchall()
        }

        students = _get_group_students_for_diploma(conn, selected_group_id)
        for st in students:
            m = compute_student_diploma_metrics(
                conn,
                st['id'],
                selected_group_id,
                include_courseworks=include_cw,
                accreditation_pairs=accreditation_pairs,
                dup_diploma_nums=dup_diploma_nums,
                dup_appendix_nums=dup_appendix_nums,
            )
            if not m:
                continue
            student_rows.append(m)
            kpi['total'] += 1
            if m['ready_status'] == 'ready':
                kpi['ready'] += 1
            elif m['ready_status'] == 'warn':
                kpi['warn'] += 1
            else:
                kpi['blocked'] += 1
            if m['final_honours']:
                kpi['honours'] += 1
            if m['honours_category'] == 'borderline':
                kpi['borderline'] += 1
            if m['needs_reprint']:
                kpi['needs_reprint'] += 1

        group_subjects = [
            dict(r) for r in conn.execute(
                "SELECT id, code, name, name_en, credits, type FROM subjects WHERE group_id = ? ORDER BY position",
                (selected_group_id,),
            ).fetchall()
        ]
        group_practices = [
            dict(r) for r in conn.execute(
                "SELECT id, code, name, name_en, credits, type FROM practices WHERE group_id = ? ORDER BY position",
                (selected_group_id,),
            ).fetchall()
        ]
        group_courseworks = [
            dict(r) for r in conn.execute(
                "SELECT id, code, name, name_en, credits, type FROM courseworks WHERE group_id = ? ORDER BY position",
                (selected_group_id,),
            ).fetchall()
        ]
        group_attestations = [
            dict(r) for r in conn.execute(
                "SELECT id, code, name, name_en, credits, type FROM attestations WHERE group_id = ? ORDER BY position",
                (selected_group_id,),
            ).fetchall()
        ]

        sids = [s['id'] for s in students]
        if sids:
            ph = ','.join('?' for _ in sids)
            for ag in conn.execute(
                f"SELECT student_id, entity_id, entity_type, grade, name, name_en FROM activity_grades WHERE student_id IN ({ph})",
                sids,
            ).fetchall():
                activity_values[(ag['student_id'], ag['entity_type'], ag['entity_id'])] = {
                    'grade': '' if ag['grade'] is None else str(ag['grade']),
                    'name': ag['name'] or '',
                    'name_en': ag['name_en'] or '',
                }

        logs_list = [
            dict(r) for r in conn.execute(
                """
                SELECT l.*, s.last_name_UA || ' ' || s.first_name_UA || ' ' || COALESCE(s.middle_name_UA, '') AS full_name
                FROM diploma_generation_logs l
                JOIN students s ON s.id = l.student_id
                WHERE l.group_id = ?
                ORDER BY l.id DESC
                LIMIT 200
                """,
                (selected_group_id,),
            ).fetchall()
        ]

    available_templates = get_templates_with_metadata(is_admin=session.get('is_admin', False))
    conn.close()

    return render_template(
        'diploma_supplements.html',
        groups=groups,
        selected_group_id=selected_group_id,
        selected_group=selected_group,
        include_archived=include_archived,
        active_tab=active_tab,
        student_rows=student_rows,
        kpi=kpi,
        settings=settings,
        deepl_enabled=bool(DEEPL_API_KEY),
        group_subjects=group_subjects,
        group_practices=group_practices,
        group_courseworks=group_courseworks,
        group_attestations=group_attestations,
        activity_values=activity_values,
        logs_list=logs_list,
        available_templates=available_templates,
    )


@diploma_bp.route('/admin/diploma-supplements/honours-override', methods=['POST'])
@permission_required('manage_diplomas')
def honours_override():
    """AJAX-збереження ручного рішення щодо диплома з відзнакою."""
    data = request.get_json(silent=True) or {}
    student_id = data.get('student_id')
    group_id = data.get('group_id')
    override_val = (data.get('override') or '').strip()
    note = (data.get('note') or '').strip()[:240]

    if override_val not in ('', 'force_yes', 'force_no'):
        return jsonify({'ok': False, 'error': 'Некоректне значення перевизначення'}), 400
    try:
        student_id = int(student_id)
        group_id = int(group_id)
    except (TypeError, ValueError):
        return jsonify({'ok': False, 'error': 'Не вказано студента або групу'}), 400

    conn = get_db()
    ensure_diploma_supplement_schema(conn)
    existing = conn.execute(
        "SELECT id FROM diplomas WHERE student_id = ? AND group_id = ?",
        (student_id, group_id),
    ).fetchone()
    db_override = override_val or None
    if existing:
        conn.execute(
            "UPDATE diplomas SET honours_override = ?, honours_note = ? WHERE id = ?",
            (db_override, note or None, existing['id']),
        )
    else:
        conn.execute(
            "INSERT INTO diplomas (student_id, group_id, diploma_number, appendix_number, honours_override, honours_note) "
            "VALUES (?, ?, '', '', ?, ?)",
            (student_id, group_id, db_override, note or None),
        )
    conn.commit()
    metrics = compute_student_diploma_metrics(conn, student_id, group_id)
    conn.close()

    log_action(
        current_username(),
        f"змінив статус «Диплом з відзнакою» для студента ID {student_id} (група {group_id}): {override_val or 'авто'}",
        group_ids=[group_id],
        category='update',
        section='Дипломи та додатки',
    )
    return jsonify({
        'ok': True,
        'final_honours': metrics.get('final_honours', False),
        'honours_category': metrics.get('honours_category', ''),
        'honours_reason': metrics.get('honours_reason', ''),
    })


@diploma_bp.route('/admin/diploma-supplements/activities-bulk', methods=['POST'])
@permission_required('manage_diplomas')
def activities_bulk_save():
    """Масове збереження оцінок практик/курсових/атестацій і тем робіт (UA + EN) для групи."""
    group_id = request.form.get('group_id', type=int)
    if not group_id:
        flash('Не обрано групу', 'danger')
        return redirect(url_for('diploma.hub'))

    conn = get_db()
    ensure_diploma_supplement_schema(conn)
    students = _get_group_students_for_diploma(conn, group_id)
    practices = conn.execute("SELECT id FROM practices WHERE group_id = ?", (group_id,)).fetchall()
    courseworks = conn.execute("SELECT id FROM courseworks WHERE group_id = ?", (group_id,)).fetchall()
    attestations = conn.execute("SELECT id FROM attestations WHERE group_id = ?", (group_id,)).fetchall()

    updated = 0
    try:
        for st in students:
            sid = st['id']
            for etype, entities in (
                ('practice', practices),
                ('coursework', courseworks),
                ('attestation', attestations),
            ):
                for ent in entities:
                    eid = ent['id']
                    raw_grade = (request.form.get(f"grade_{sid}_{etype}_{eid}") or '').strip()
                    topic_ua = (request.form.get(f"topic_ua_{sid}_{eid}") or '').strip() if etype == 'attestation' else ''
                    topic_en = (request.form.get(f"topic_en_{sid}_{eid}") or '').strip() if etype == 'attestation' else ''

                    existing = conn.execute(
                        "SELECT id, grade, name, name_en FROM activity_grades WHERE student_id = ? AND entity_id = ? AND entity_type = ?",
                        (sid, eid, etype),
                    ).fetchone()

                    if not raw_grade and not topic_ua and not topic_en:
                        if existing:
                            conn.execute("DELETE FROM activity_grades WHERE id = ?", (existing['id'],))
                            updated += 1
                        continue

                    grade_val = None
                    if raw_grade:
                        try:
                            g_int = int(raw_grade)
                            if 0 <= g_int <= 100:
                                grade_val = g_int
                            else:
                                continue
                        except ValueError:
                            continue

                    if existing:
                        conn.execute(
                            "UPDATE activity_grades SET grade = ?, name = ?, name_en = ? WHERE id = ?",
                            (grade_val, topic_ua, topic_en, existing['id']),
                        )
                    else:
                        conn.execute(
                            "INSERT INTO activity_grades (student_id, entity_id, entity_type, grade, name, name_en) "
                            "VALUES (?, ?, ?, ?, ?, ?)",
                            (sid, eid, etype, grade_val, topic_ua, topic_en),
                        )
                    updated += 1
        conn.commit()
        log_action(
            current_username(),
            f"масово оновив активності та теми робіт (група ID {group_id})",
            group_ids=[group_id],
            category='update',
            section='Дипломи та додатки',
            details=f"оновлено записів: {updated}",
        )
        flash(f"Збережено зміни ({updated} позицій)", 'success')
    except Exception as e:
        conn.rollback()
        logger.error(f"activities_bulk_save error: {e}", exc_info=True)
        flash('Помилка збереження оцінок та тем', 'danger')
    finally:
        conn.close()

    return redirect(url_for('diploma.hub', group_id=group_id, tab='activities', include_archived=request.form.get('include_archived') or None))


@diploma_bp.route('/admin/diploma-supplements/subjects-en-bulk', methods=['POST'])
@permission_required('manage_diplomas')
def subjects_en_bulk_save():
    """Зберігає англійські переклади назв предметів/практик/курсових/атестацій групи."""
    group_id = request.form.get('group_id', type=int)
    if not group_id:
        flash('Не обрано групу', 'danger')
        return redirect(url_for('diploma.hub'))

    conn = get_db()
    ensure_diploma_supplement_schema(conn)
    saved = 0
    try:
        for prefix, tbl in (
            ('subj', 'subjects'),
            ('prac', 'practices'),
            ('cw', 'courseworks'),
            ('att', 'attestations'),
        ):
            rows = conn.execute(f"SELECT id FROM {tbl} WHERE group_id = ?", (group_id,)).fetchall()
            for r in rows:
                val = (request.form.get(f"{prefix}_en_{r['id']}") or '').strip()
                conn.execute(
                    f"UPDATE {tbl} SET name_en = ? WHERE id = ? AND group_id = ?",
                    (val or None, r['id'], group_id),
                )
                if val:
                    saved += 1
        conn.commit()
        log_action(
            current_username(),
            f"зберіг EN-переклади компонентів програми (група ID {group_id})",
            group_ids=[group_id],
            category='update',
            section='Дипломи та додатки',
        )
        flash(f"Переклади компонентів збережено (заповнено EN: {saved})", 'success')
    except Exception as e:
        conn.rollback()
        logger.error(f"subjects_en_bulk_save error: {e}", exc_info=True)
        flash('Помилка збереження перекладів', 'danger')
    finally:
        conn.close()

    return redirect(url_for('diploma.hub', group_id=group_id, tab='activities', include_archived=request.form.get('include_archived') or None))


@diploma_bp.route('/admin/diploma-supplements/translate-missing', methods=['POST'])
@permission_required('manage_diplomas')
def translate_missing():
    """Пакетний або точковий автопереклад через DeepL для тем робіт і предметів."""
    data = request.get_json(silent=True) or {}
    mode = data.get('mode') or 'topics'
    group_id = data.get('group_id')
    text = (data.get('text') or '').strip()

    if mode == 'single':
        if not text:
            return jsonify({'ok': True, 'translation': ''})
        tr = _translate_text_deepl(text)
        return jsonify({'ok': bool(tr), 'translation': tr, 'error': '' if tr else 'DeepL не повернув переклад'})

    try:
        group_id = int(group_id)
    except (TypeError, ValueError):
        return jsonify({'ok': False, 'error': 'Не вказано групу'}), 400

    if not DEEPL_API_KEY:
        return jsonify({'ok': False, 'error': 'DEEPL_API_KEY не налаштовано в .env'}), 400

    conn = get_db()
    ensure_diploma_supplement_schema(conn)
    translated_count = 0

    try:
        if mode == 'topics':
            students = _get_group_students_for_diploma(conn, group_id)
            sids = [s['id'] for s in students]
            att_ids = [r['id'] for r in conn.execute("SELECT id FROM attestations WHERE group_id = ?", (group_id,)).fetchall()]
            if sids and att_ids:
                ph_s = ','.join('?' for _ in sids)
                ph_a = ','.join('?' for _ in att_ids)
                rows = conn.execute(
                    f"""
                    SELECT id, name, name_en FROM activity_grades
                    WHERE entity_type = 'attestation'
                      AND student_id IN ({ph_s})
                      AND entity_id IN ({ph_a})
                      AND COALESCE(TRIM(name), '') != ''
                      AND COALESCE(TRIM(name_en), '') = ''
                    """,
                    sids + att_ids,
                ).fetchall()
                for r in rows:
                    tr = _translate_text_deepl(r['name'].strip())
                    if tr:
                        conn.execute("UPDATE activity_grades SET name_en = ? WHERE id = ?", (tr, r['id']))
                        translated_count += 1
        elif mode == 'components':
            for tbl in ('subjects', 'practices', 'courseworks', 'attestations'):
                rows = conn.execute(
                    f"SELECT id, name FROM {tbl} WHERE group_id = ? AND COALESCE(TRIM(name), '') != '' AND COALESCE(TRIM(name_en), '') = ''",
                    (group_id,),
                ).fetchall()
                for r in rows:
                    tr = _translate_text_deepl(r['name'].strip())
                    if tr:
                        conn.execute(f"UPDATE {tbl} SET name_en = ? WHERE id = ?", (tr, r['id']))
                        translated_count += 1
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        logger.error(f"translate_missing error: {e}", exc_info=True)
        return jsonify({'ok': False, 'error': 'Помилка перекладу'}), 500

    conn.close()
    return jsonify({'ok': True, 'translated_count': translated_count})


@diploma_bp.route('/admin/diploma-supplements/auto-number', methods=['POST'])
@permission_required('manage_diplomas')
def auto_number():
    """Пакетна послідовна нумерація дипломів/додатків діапазоном для групи."""
    group_id = request.form.get('group_id', type=int)
    if not group_id:
        flash('Не обрано групу', 'danger')
        return redirect(url_for('diploma.hub'))

    dip_start_raw = (request.form.get('diploma_start') or '').strip()
    app_start_raw = (request.form.get('appendix_start') or '').strip()
    dip_prefix = (request.form.get('diploma_prefix') or '').strip()
    app_prefix = (request.form.get('appendix_prefix') or '').strip()
    only_empty = request.form.get('only_empty') == '1'
    pad_len = request.form.get('pad_len', type=int) or 6
    pad_len = max(1, min(12, pad_len))

    if not dip_start_raw and not app_start_raw:
        flash('Вкажіть початковий номер диплома або додатка', 'warning')
        return redirect(url_for('diploma.hub', group_id=group_id, tab='import'))

    try:
        dip_num = int(dip_start_raw) if dip_start_raw else None
        app_num = int(app_start_raw) if app_start_raw else None
    except ValueError:
        flash('Початковий номер має бути цілим числом', 'danger')
        return redirect(url_for('diploma.hub', group_id=group_id, tab='import'))

    conn = get_db()
    ensure_diploma_supplement_schema(conn)
    students = _get_group_students_for_diploma(conn, group_id)
    updated = 0

    try:
        for st in students:
            sid = st['id']
            existing = conn.execute(
                "SELECT id, diploma_number, appendix_number FROM diplomas WHERE student_id = ? AND group_id = ?",
                (sid, group_id),
            ).fetchone()
            cur_dip = (existing['diploma_number'] or '').strip() if existing else ''
            cur_app = (existing['appendix_number'] or '').strip() if existing else ''

            new_dip = cur_dip
            new_app = cur_app
            changed = False

            if dip_num is not None and (not only_empty or not cur_dip):
                new_dip = f"{dip_prefix}{str(dip_num).zfill(pad_len)}"
                dip_num += 1
                changed = True
            if app_num is not None and (not only_empty or not cur_app):
                new_app = f"{app_prefix}{str(app_num).zfill(pad_len)}"
                app_num += 1
                changed = True

            if not changed:
                continue

            if existing:
                conn.execute(
                    "UPDATE diplomas SET diploma_number = ?, appendix_number = ? WHERE id = ?",
                    (new_dip, new_app, existing['id']),
                )
            else:
                conn.execute(
                    "INSERT INTO diplomas (student_id, group_id, diploma_number, appendix_number) VALUES (?, ?, ?, ?)",
                    (sid, group_id, new_dip, new_app),
                )
            updated += 1
        conn.commit()
        log_action(
            current_username(),
            f"пакетно пронумерував дипломи/додатки для групи ID {group_id}",
            group_ids=[group_id],
            category='update',
            section='Дипломи та додатки',
            details=f"оновлено студентів: {updated}",
        )
        flash(f"Пронумеровано студентів: {updated}", 'success')
    except Exception as e:
        conn.rollback()
        logger.error(f"auto_number error: {e}", exc_info=True)
        flash('Помилка пакетної нумерації', 'danger')
    finally:
        conn.close()

    return redirect(url_for('diploma.hub', group_id=group_id, tab='audit', include_archived=request.form.get('include_archived') or None))


@diploma_bp.route('/admin/diploma-supplements/edebo-import/preview', methods=['POST'])
@permission_required('manage_diplomas')
def edebo_import_preview():
    """Розбирає завантажений Excel (ЄДЕБО або таблицю номерів) і показує прев'ю зіставлення."""
    _cleanup_old_edebo_previews()
    group_id = request.form.get('group_id', type=int)
    if not group_id:
        flash('Спочатку оберіть групу', 'danger')
        return redirect(url_for('diploma.hub'))

    file = request.files.get('excel_file')
    if not file or not file.filename:
        flash('Оберіть файл Excel (.xlsx)', 'danger')
        return redirect(url_for('diploma.hub', group_id=group_id, tab='import'))

    ok_file, err_file = validate_upload_file(file, {'.xlsx'})
    if not ok_file:
        flash(err_file, 'danger')
        return redirect(url_for('diploma.hub', group_id=group_id, tab='import'))

    tmp_xlsx = os.path.join(tempfile.gettempdir(), f"edebo_{uuid.uuid4().hex}.xlsx")
    file.save(tmp_xlsx)

    conn = get_db()
    ensure_diploma_supplement_schema(conn)
    try:
        wb = load_workbook(tmp_xlsx, data_only=True, read_only=True)
        ws = wb.active
        all_rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    except Exception as e:
        conn.close()
        try:
            os.remove(tmp_xlsx)
        except OSError:
            pass
        flash(f"Не вдалося прочитати Excel: {e}", 'danger')
        return redirect(url_for('diploma.hub', group_id=group_id, tab='import'))
    finally:
        try:
            os.remove(tmp_xlsx)
        except OSError:
            pass

    if len(all_rows) < 2:
        conn.close()
        flash('У файлі замало рядків (потрібен рядок заголовків і хоча б один рядок даних)', 'warning')
        return redirect(url_for('diploma.hub', group_id=group_id, tab='import'))

    # Знаходимо рядок заголовків серед перших 10 рядків
    header_idx = 0
    col_map = _detect_edebo_columns(all_rows[0])
    for idx in range(min(10, len(all_rows))):
        cand_map = _detect_edebo_columns(all_rows[idx])
        if (cand_map['diploma_number'] is not None or cand_map['appendix_number'] is not None) and (
            cand_map['edebo_code'] is not None or cand_map['fio_ua'] is not None or cand_map['last_name'] is not None
        ):
            header_idx = idx
            col_map = cand_map
            break

    if col_map['diploma_number'] is None and col_map['appendix_number'] is None:
        conn.close()
        flash('Не знайдено колонок з номером диплома або додатка (очікуються назви типу «Номер диплома», «Номер додатка»).', 'danger')
        return redirect(url_for('diploma.hub', group_id=group_id, tab='import'))

    group_row = conn.execute("SELECT id, name, start_year FROM groups WHERE id = ?", (group_id,)).fetchone()
    students = _get_group_students_for_diploma(conn, group_id)

    # Поточні номери в групі та інші дублі в БД
    existing_by_student = {}
    for d in conn.execute(
        "SELECT student_id, diploma_number, appendix_number FROM diplomas WHERE group_id = ?",
        (group_id,),
    ).fetchall():
        existing_by_student[d['student_id']] = dict(d)

    all_dip_owners = {}
    all_app_owners = {}
    for d in conn.execute("SELECT student_id, group_id, diploma_number, appendix_number FROM diplomas").fetchall():
        if d['diploma_number']:
            all_dip_owners.setdefault(d['diploma_number'].zfill(6), []).append((d['student_id'], d['group_id']))
        if d['appendix_number']:
            all_app_owners.setdefault(d['appendix_number'].zfill(6), []).append((d['student_id'], d['group_id']))

    by_edebo = {
        str(s['edebo_code']).strip(): s for s in students if s.get('edebo_code') and str(s['edebo_code']).strip()
    }
    student_candidates = [
        (s, _norm_fio(f"{s['last_name_UA']} {s['first_name_UA']} {s['middle_name_UA'] or ''}"))
        for s in students
    ]

    preview_rows = []
    seen_dip_in_file = set()
    seen_app_in_file = set()

    for r_num, row in enumerate(all_rows[header_idx + 1:], start=header_idx + 2):
        if not row or all(c is None or str(c).strip() == '' for c in row):
            continue

        def _cell(key):
            ci = col_map.get(key)
            if ci is None or ci >= len(row):
                return ''
            return '' if row[ci] is None else str(row[ci]).strip()

        edebo_val = _cell('edebo_code')
        if col_map['last_name'] is not None:
            fio_excel = ' '.join(x for x in (_cell('last_name'), _cell('first_name'), _cell('middle_name')) if x)
        else:
            fio_excel = _cell('fio_ua')

        fio_en_excel = _cell('fio_en')
        birth_excel = _cell('birth_date')
        dip_val = _clean_doc_num(row[col_map['diploma_number']]) if col_map['diploma_number'] is not None and col_map['diploma_number'] < len(row) else ''
        app_val = _clean_doc_num(row[col_map['appendix_number']]) if col_map['appendix_number'] is not None and col_map['appendix_number'] < len(row) else ''

        if not dip_val and not app_val:
            continue

        matched_st = None
        match_method = ''
        match_score = 0.0

        if edebo_val and edebo_val in by_edebo:
            matched_st = by_edebo[edebo_val]
            match_method = 'ЄДЕБО'
            match_score = 100.0
        elif fio_excel:
            norm_ex = _norm_fio(fio_excel)
            best_st, best_sc = None, 0.0
            for st_obj, st_norm in student_candidates:
                sc = difflib.SequenceMatcher(None, norm_ex, st_norm).ratio()
                if sc > best_sc:
                    best_st, best_sc = st_obj, sc
            if best_st and best_sc >= 0.72:
                matched_st = best_st
                match_method = 'ПІБ'
                match_score = round(best_sc * 100, 1)

        warnings = []
        status = 'ok'
        if not matched_st:
            status = 'error'
            warnings.append('Студента не знайдено у групі')
        else:
            if fio_en_excel:
                db_en = f"{matched_st.get('last_name_ENG') or ''} {matched_st.get('first_name_ENG') or ''}".strip().lower()
                if db_en and db_en != fio_en_excel.lower():
                    warnings.append(f"EN у базі «{matched_st.get('last_name_ENG')} {matched_st.get('first_name_ENG')}» ≠ «{fio_en_excel}»")
            if birth_excel and matched_st.get('birth_date'):
                if _to_slash_date(birth_excel) != _to_slash_date(matched_st['birth_date']):
                    warnings.append(f"Д/н у базі {matched_st['birth_date']} ≠ {birth_excel}")

        if dip_val:
            if dip_val in seen_dip_in_file:
                warnings.append(f'Дубль диплома {dip_val} у файлі')
                status = 'warn' if status == 'ok' else status
            seen_dip_in_file.add(dip_val)
            owners = [o for o in all_dip_owners.get(dip_val, []) if not matched_st or o != (matched_st['id'], group_id)]
            if owners:
                warnings.append(f'Номер диплома {dip_val} вже зайнятий у БД')
                status = 'warn' if status == 'ok' else status

        if app_val:
            if app_val in seen_app_in_file:
                warnings.append(f'Дубль додатка {app_val} у файлі')
                status = 'warn' if status == 'ok' else status
            seen_app_in_file.add(app_val)
            owners = [o for o in all_app_owners.get(app_val, []) if not matched_st or o != (matched_st['id'], group_id)]
            if owners:
                warnings.append(f'Номер додатка {app_val} вже зайнятий у БД')
                status = 'warn' if status == 'ok' else status

        ex_rec = existing_by_student.get(matched_st['id'], {}) if matched_st else {}
        preview_rows.append({
            'row': r_num,
            'edebo_excel': edebo_val,
            'fio_excel': fio_excel,
            'student_id': matched_st['id'] if matched_st else None,
            'student_name': f"{matched_st['last_name_UA']} {matched_st['first_name_UA']} {matched_st['middle_name_UA'] or ''}".strip() if matched_st else '',
            'match_method': match_method,
            'match_score': match_score,
            'diploma_number': dip_val,
            'appendix_number': app_val,
            'existing_diploma': (ex_rec.get('diploma_number') or '').strip(),
            'existing_appendix': (ex_rec.get('appendix_number') or '').strip(),
            'warnings': warnings,
            'status': status,
        })

    conn.close()
    if not preview_rows:
        flash('У файлі не знайдено рядків із номерами дипломів чи додатків', 'warning')
        return redirect(url_for('diploma.hub', group_id=group_id, tab='import'))

    os.makedirs(EDEBO_PREVIEW_DIR, exist_ok=True)
    token = uuid.uuid4().hex
    state_path = resolve_under_dir(EDEBO_PREVIEW_DIR, f"{token}.json")
    with open(state_path, 'w', encoding='utf-8') as f:
        json.dump({
            'user_id': session.get('user_id'),
            'group_id': group_id,
            'group_name': group_row['name'] if group_row else f'#{group_id}',
            'filename': secure_filename(file.filename),
            'rows': preview_rows,
            'created_at': time.time(),
        }, f, ensure_ascii=False)

    return render_template(
        'diploma_edebo_preview.html',
        token=token,
        group_id=group_id,
        group_name=group_row['name'] if group_row else f'#{group_id}',
        filename=secure_filename(file.filename),
        preview_rows=preview_rows,
        students=students,
    )


@diploma_bp.route('/admin/diploma-supplements/edebo-import/commit/<token>', methods=['POST'])
@permission_required('manage_diplomas')
def edebo_import_commit(token):
    """Застосовує підтверджений імпорт номерів диплома/додатка з Excel."""
    if not is_safe_hex_token(token):
        flash('Некоректний токен сесії імпорту', 'danger')
        return redirect(url_for('diploma.hub'))
    state_path = resolve_under_dir(EDEBO_PREVIEW_DIR, f"{token}.json")
    if not state_path or not os.path.isfile(state_path):
        flash('Сесія імпорту застаріла або не знайдена', 'warning')
        return redirect(url_for('diploma.hub'))

    with open(state_path, 'r', encoding='utf-8') as f:
        state = json.load(f)
    if state.get('user_id') != session.get('user_id') and not session.get('is_admin'):
        flash('Доступ до цієї сесії імпорту заборонено', 'danger')
        return redirect(url_for('diploma.hub'))

    group_id = state['group_id']
    conn = get_db()
    ensure_diploma_supplement_schema(conn)
    saved = 0

    try:
        for item in state.get('rows', []):
            r_num = item['row']
            if not request.form.get(f"apply_{r_num}"):
                continue
            sid = request.form.get(f"student_id_{r_num}", type=int) or item.get('student_id')
            if not sid:
                continue
            dip_num = (request.form.get(f"diploma_{r_num}") or item.get('diploma_number') or '').strip()
            app_num = (request.form.get(f"appendix_{r_num}") or item.get('appendix_number') or '').strip()
            if dip_num.isdigit():
                dip_num = dip_num.zfill(6)
            if app_num.isdigit():
                app_num = app_num.zfill(6)

            existing = conn.execute(
                "SELECT id, diploma_number, appendix_number FROM diplomas WHERE student_id = ? AND group_id = ?",
                (sid, group_id),
            ).fetchone()
            if existing:
                conn.execute(
                    "UPDATE diplomas SET diploma_number = ?, appendix_number = ? WHERE id = ?",
                    (
                        dip_num or existing['diploma_number'] or '',
                        app_num or existing['appendix_number'] or '',
                        existing['id'],
                    ),
                )
            else:
                conn.execute(
                    "INSERT INTO diplomas (student_id, group_id, diploma_number, appendix_number) VALUES (?, ?, ?, ?)",
                    (sid, group_id, dip_num, app_num),
                )
            saved += 1
        conn.commit()
        try:
            os.remove(state_path)
        except OSError:
            pass
        log_action(
            current_username(),
            f"імпортував номери диплома/додатка з Excel ({state.get('filename')}) для групи {state.get('group_name')}",
            group_ids=[group_id],
            category='import',
            section='Дипломи та додатки',
            details=f"оновлено записів: {saved}",
        )
        flash(f"Успішно оновлено номерів диплома/додатка: {saved}", 'success')
    except Exception as e:
        conn.rollback()
        logger.error(f"edebo_import_commit error: {e}", exc_info=True)
        flash('Помилка збереження імпортованих номерів', 'danger')
    finally:
        conn.close()

    return redirect(url_for('diploma.hub', group_id=group_id, tab='audit'))


@diploma_bp.route('/admin/diploma-supplements/logs/status', methods=['POST'])
@permission_required('manage_diplomas')
def update_log_status():
    """Оновлює статус друку/видачі в журналі додатків до диплома."""
    group_id = request.form.get('group_id', type=int)
    new_status = (request.form.get('status') or '').strip()
    log_ids = [int(x) for x in request.form.getlist('log_ids') if str(x).isdigit()]
    single_id = request.form.get('log_id', type=int)
    if single_id:
        log_ids.append(single_id)

    if new_status not in ('generated', 'printed', 'issued') or not log_ids:
        flash('Оберіть записи та коректний статус', 'warning')
        return redirect(url_for('diploma.hub', group_id=group_id, tab='logs'))

    conn = get_db()
    ensure_diploma_supplement_schema(conn)
    user = current_username()
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        ph = ','.join('?' for _ in log_ids)
        if new_status == 'printed':
            conn.execute(
                f"UPDATE diploma_generation_logs SET status = 'printed', printed_at = COALESCE(printed_at, ?), printed_by = COALESCE(printed_by, ?) WHERE id IN ({ph})",
                [now_str, user] + log_ids,
            )
        elif new_status == 'issued':
            conn.execute(
                f"UPDATE diploma_generation_logs SET status = 'issued', printed_at = COALESCE(printed_at, ?), printed_by = COALESCE(printed_by, ?), issued_at = ?, issued_by = ? WHERE id IN ({ph})",
                [now_str, user, now_str, user] + log_ids,
            )
        else:
            conn.execute(
                f"UPDATE diploma_generation_logs SET status = 'generated' WHERE id IN ({ph})",
                log_ids,
            )
        conn.commit()
        flash(f"Статус оновлено для {len(log_ids)} запис(ів)", 'success')
    except Exception as e:
        conn.rollback()
        logger.error(f"update_log_status error: {e}", exc_info=True)
        flash('Помилка оновлення статусу журналу', 'danger')
    finally:
        conn.close()

    return redirect(url_for('diploma.hub', group_id=group_id, tab='logs'))


@diploma_bp.route('/admin/diploma-supplements/settings', methods=['POST'])
@permission_required('manage_diplomas')
def save_settings_route():
    """Зберігає налаштування маски файлу, правил відзнаки та двостороннього PDF."""
    group_id = request.form.get('group_id', type=int)
    mask = (request.form.get('filename_mask') or '').strip() or DEFAULT_DIPLOMA_SETTINGS['filename_mask']
    include_cw = '1' if request.form.get('honours_include_courseworks') == '1' else '0'
    duplex = '1' if request.form.get('pdf_duplex_pad') == '1' else '0'

    conn = get_db()
    save_diploma_settings(conn, {
        'filename_mask': mask,
        'honours_include_courseworks': include_cw,
        'pdf_duplex_pad': duplex,
    })
    conn.close()
    log_action(
        current_username(),
        "оновив налаштування додатків до диплома",
        category='update',
        section='Дипломи та додатки',
        details=f"маска='{mask}', курсові_у_відзнаці={include_cw}, duplex_pdf={duplex}",
    )
    flash('Налаштування збережено', 'success')
    return redirect(url_for('diploma.hub', group_id=group_id, tab='logs'))


@diploma_bp.route('/admin/diploma-supplements/api/students', methods=['GET'])
@permission_required('manage_templates')
def api_search_students():
    """Швидкий пошук студентів + їхніх груп для тестового рендера шаблону."""
    q = (request.args.get('q') or '').strip().lower()
    conn = get_db()
    rows = conn.execute(
        """
        SELECT s.id,
               s.last_name_UA || ' ' || s.first_name_UA || ' ' || COALESCE(s.middle_name_UA, '') AS full_name,
               COALESCE(e.group_id, s.group_id) AS group_id,
               g.name AS group_name,
               g.start_year,
               g.degree_level
        FROM students s
        LEFT JOIN student_enrollments e ON e.student_id = s.id AND e.status IN ('active', 'graduated')
        LEFT JOIN groups g ON g.id = COALESCE(e.group_id, s.group_id)
        WHERE COALESCE(s.archived, 0) = 0 OR e.status = 'graduated'
        ORDER BY s.last_name_UA COLLATE UKRAINIAN, s.first_name_UA COLLATE UKRAINIAN
        LIMIT 400
        """
    ).fetchall()
    conn.close()

    items = []
    for r in rows:
        name = (r['full_name'] or '').strip()
        gname = r['group_name'] or 'Без групи'
        hay = f"{name} {gname} {r['degree_level'] or ''}".lower()
        if q and q not in hay:
            continue
        items.append({
            'student_id': r['id'],
            'group_id': r['group_id'],
            'label': f"{name} — {gname} ({r['start_year'] or ''})",
        })
        if len(items) >= 40:
            break
    return jsonify({'ok': True, 'students': items})


@diploma_bp.route('/admin/templates/<filename>/test-render', methods=['POST'])
@permission_required('manage_templates')
def test_render_template(filename):
    """Тестова перевірка .docx-шаблону на реальному студенту без запису у журнал."""
    from routes.gen_docx import gen_doc
    from routes import office_editor

    safe_name = secure_filename(filename)
    tpl_path = resolve_under_dir(TEMPLATE_FOLDER, safe_name)
    if not tpl_path or not os.path.isfile(tpl_path):
        return jsonify({'ok': False, 'error': f"Шаблон «{safe_name}» не знайдено"}), 404

    data = request.get_json(silent=True) or {}
    student_id = data.get('student_id')
    group_id = data.get('group_id')
    try:
        student_id = int(student_id)
    except (TypeError, ValueError):
        return jsonify({'ok': False, 'error': 'Оберіть студента для тестування'}), 400
    try:
        group_id = int(group_id) if group_id else None
    except (TypeError, ValueError):
        group_id = None

    conn = get_db()
    ensure_diploma_supplement_schema(conn)
    student = conn.execute(
        """
        SELECT s.*,
               COALESCE(e.edebo_code, s.edebo_code) AS enrollment_edebo_code,
               COALESCE(?, s.group_id) AS effective_group_id,
               g.name || ' (' || g.start_year || ', ' || g.study_form || ', ' || g.program_credits || ' кредитів)' AS group_name,
               g.study_form, g.start_year, g.program_credits,
               g.qualification_name, g.degree_level, g.specialty, g.educational_program, g.knowledge_area,
               g.qualification_name_en, g.degree_level_en, g.specialty_en, g.isced_f_ua, g.isced_f_en,
               g.educational_program_en, g.knowledge_area_en,
               COALESCE(NULLIF(TRIM(g.faculty), ''), fl.name_ua, '') AS faculty,
               COALESCE(NULLIF(TRIM(g.faculty_en), ''), fl.name_en, '') AS faculty_en,
               il.name_ua AS institution_name_and_status, il.name_en AS institution_name_and_status_en,
               il.short_name_ua AS license_short_name_ua, il.short_name_en AS license_short_name_en,
               g.entry_requirements, g.entry_requirements_en,
               g.learning_outcomes, g.learning_outcomes_en, g.program_includes, g.program_includes_en,
               g.entry_requirements_reduced, g.entry_requirements_reduced_en,
               g.learning_outcomes_reduced, g.learning_outcomes_reduced_en,
               g.program_includes_reduced, g.program_includes_reduced_en
        FROM students s
        LEFT JOIN groups g ON g.id = COALESCE(?, s.group_id)
        LEFT JOIN specialties sp ON sp.code = g.specialty_code
        LEFT JOIN faculties fl ON fl.id = COALESCE(g.faculty_id, sp.faculty_id)
        LEFT JOIN student_enrollments e ON e.student_id = s.id AND e.group_id = g.id
        LEFT JOIN institution_licenses il ON il.id = COALESCE(e.license_id, s.license_id)
        WHERE s.id = ?
        """,
        (group_id, group_id, student_id),
    ).fetchone()
    if not student:
        conn.close()
        return jsonify({'ok': False, 'error': 'Студента не знайдено'}), 404

    student_dict = dict(student)
    if student_dict.get('effective_group_id'):
        student_dict['group_id'] = student_dict['effective_group_id']
    if student_dict.get('enrollment_edebo_code'):
        student_dict['edebo_code'] = student_dict['enrollment_edebo_code']
    military = conn.execute("SELECT * FROM military WHERE student_id = ?", (student_id,)).fetchone()
    military_dict = dict(military) if military else {}
    conn.close()

    out_path = os.path.join(office_editor.SESSIONS_DIR, f"test_{uuid.uuid4().hex}.docx")
    captured_context = {}
    render_ok = True
    render_error = ''
    try:
        gen_doc(
            student_dict,
            military_dict,
            template=tpl_path,
            out=out_path,
            user_name=current_username(),
            record_log=False,
            capture_context=captured_context,
        )
    except Exception as e:
        render_ok = False
        render_error = str(e)

    diag = inspect_docx_template(tpl_path, captured_context)
    if not render_ok and not diag['syntax_error']:
        diag['syntax_ok'] = False
        diag['syntax_error'] = render_error

    doc_id = None
    edit_url = ''
    download_url = ''
    if render_ok and os.path.isfile(out_path):
        disp_name = f"TEST_{safe_name}"
        doc_id = office_editor.create_editing_session(out_path, disp_name, session.get('user_id'))
        edit_url = url_for('office.edit', doc_id=doc_id)
        download_url = url_for('office.finalize', doc_id=doc_id)

    return jsonify({
        'ok': render_ok and diag['syntax_ok'],
        'render_ok': render_ok,
        'diagnostics': diag,
        'doc_id': doc_id,
        'edit_url': edit_url,
        'download_url': download_url,
    })
