"""
routes/students/import_export.py
==================================
Масовий імпорт студентів з Excel-файлу з ОБОВ'ЯЗКОВИМ підтвердженням
усіх змін через проміжну сторінку прев'ю.

Архітектура (Вариант Б, повне прев'ю):
  POST /import_from_excel              → парсинг, БЕЗ запису в БД, redirect
  GET  /import_from_excel/preview/<t>  → сторінка прев'ю з чекбоксами
  POST /import_from_excel/commit/<t>   → запис лише обраного
  GET  /import_from_excel/report/<id>  → підсумковий звіт (skipped + stats)

Стан між кроками зберігається в static/uploads/import_sessions/students/<token>.json
(не в сесії — Excel може дати 500+ рядків, cookie не витримає).

Правила вибору групи для enrollment-полів (edebo_code, program_credits_override):
  - Колонка A заповнена → довіряємо Excel, пишемо в цю групу.
  - Колонка A порожня + 1 активний enrollment → пишемо туди.
  - Колонка A порожня + >1 активних enrollment → ОБОВ'ЯЗКОВИЙ вибір radio
    у прев'ю (без вибору поле не записується, warning у звіті).
"""
from flask import render_template, request, redirect, url_for, session, flash
from routes.db import get_db
from routes.utils import log_action, login_required, permission_required, logger
from routes.helpers import current_username
from routes.students import students_bp
from routes.gender import detect_gender_from_middle_name
from routes.enrollments_service import (
    add_enrollment, ensure_enrollments_table,
    update_enrollment_edebo_code, update_enrollment_program_credits,
)
import sqlite3
from routes.utils import generate_english_name
from werkzeug.utils import secure_filename
from datetime import datetime
import openpyxl
import os
import re
import json
import time
import uuid

# ═══════════════════════════════════════════════════════════════════════
# Каталог для тимчасових станів імпорту (між кроками)
# ═══════════════════════════════════════════════════════════════════════
STUDENTS_IMPORT_DIR = os.path.join(
    os.getcwd(), 'static', 'uploads', 'import_sessions', 'students'
)
STUDENTS_IMPORT_TTL = 6 * 3600  # 6 годин

# Каталог для звітів (використовується і старим import_report)
REPORTS_DIR = os.path.join(os.getcwd(), 'static', 'uploads', 'import_reports')
REPORTS_TTL_SECONDS = 24 * 3600


# ═══════════════════════════════════════════════════════════════════════
# Допоміжні функції — перенесені БЕЗ ЗМІН з поточного файлу
# (використовуються в _build_import_preview)
# ═══════════════════════════════════════════════════════════════════════

def _normalize_date_for_compare(s):
    """Нормалізує дату ДД.ММ.РРРР для порівняння (ігнорує роздільники)."""
    if not s:
        return ''
    s = str(s).strip()
    for sep in ('.', '/', '-'):
        s = s.replace(sep, '')
    return s


def _normalize_excel_code(raw):
    """Нормалізує «кодоподібне» значення з Excel в рядок."""
    if raw is None:
        return ''
    if isinstance(raw, bool):
        return str(raw)
    if isinstance(raw, int):
        return str(raw)
    if isinstance(raw, float):
        if raw.is_integer():
            return str(int(raw))
        return str(raw)
    return str(raw).strip()


def _normalize_gender(raw):
    """Нормалізує значення колонки «Стать» з Excel до 'M'/'F' або None."""
    if raw is None:
        return None
    v = str(raw).strip().lower()
    if not v:
        return None
    if v in ('m', 'ч', 'чоловіча', 'чоловік', 'male', 'man'):
        return 'M'
    if v in ('f', 'ж', 'жіноча', 'жінка', 'female', 'woman'):
        return 'F'
    return None


FIELD_LABELS = {
    'last_name_UA': 'Прізвище (укр.)',
    'first_name_UA': "Ім'я (укр.)",
    'middle_name_UA': 'По батькові (укр.)',
    'last_name_ENG': 'Прізвище (лат.)',
    'first_name_ENG': "Ім'я (лат.)",
    'birth_date': 'Дата народження',
    'phone': 'Телефон',
    'phone_backup': 'Резервний телефон',
    'email': 'Email',
    'tax_id': 'ІПН',
    'edebo_code': 'ЄДЕБО',
    'address_of_residence': 'Адреса',
    'gender': 'Стать',
}

ENROLLMENT_FIELD_LABELS = {
    'edebo_code': 'ЄДЕБО',
    'program_credits_override': 'Кредити скороченої програми',
}


def _collect_field_diffs(existing, new_values):
    """
    Порівнює значення з Excel з БД. Повертає список полів у форматі прев'ю:
        [
            {
                'field': 'phone',
                'label': 'Телефон',
                'old': '...',
                'new': '...',
                'type': 'fill_empty' | 'conflict',
                'enabled': True|False,   # за замовчуванням: fill_empty → True, conflict → False
            },
            ...
        ]

    Пусті в БД поля → fill_empty (увімкнено).
    Заповнені, але різні → conflict (вимкнено — користувач має свідомо вирішити).
    Однакові → не включаються (нічого не змінюється).
    """
    from routes.db import normalize_name_for_match

    diffs = []
    for field, new_val in new_values.items():
        if not new_val:
            continue

        cur_raw = existing[field] if existing[field] is not None else ''

        # Нормалізація для порівняння
        if field == 'birth_date':
            cur_norm = _normalize_date_for_compare(cur_raw)
            new_norm = _normalize_date_for_compare(new_val)
        elif field == 'gender':
            cur_norm = str(cur_raw).strip().upper()
            new_norm = str(new_val).strip().upper()
        elif field in ('phone', 'phone_backup'):
            from routes.utils import normalize_phone_ua
            cur_norm = normalize_phone_ua(cur_raw) or ''
            new_norm = normalize_phone_ua(new_val) or ''
        elif field in ('last_name_UA', 'first_name_UA', 'middle_name_UA'):
            cur_norm = normalize_name_for_match(cur_raw)
            new_norm = normalize_name_for_match(new_val)
        else:
            cur_norm = str(cur_raw).strip()
            new_norm = str(new_val).strip()

        if cur_norm == new_norm:
            continue  # однаково — нічого не змінюємо

        diffs.append({
            'field': field,
            'label': FIELD_LABELS.get(field, field),
            'old': str(cur_raw or ''),
            'new': str(new_val),
            'type': 'fill_empty' if not cur_norm else 'conflict',
            'enabled': (not cur_norm),  # fill_empty → True, conflict → False
        })

    return diffs


# ═══════════════════════════════════════════════════════════════════════
# Стан між кроками імпорту (JSON-файли)
# ═══════════════════════════════════════════════════════════════════════

def _cleanup_old_states():
    if not os.path.isdir(STUDENTS_IMPORT_DIR):
        return
    now = time.time()
    for name in os.listdir(STUDENTS_IMPORT_DIR):
        path = os.path.join(STUDENTS_IMPORT_DIR, name)
        try:
            if os.path.isfile(path) and now - os.path.getmtime(path) > STUDENTS_IMPORT_TTL:
                os.remove(path)
        except OSError:
            pass


def _state_path(token):
    token = re.sub(r'[^0-9a-f]', '', token or '')
    return os.path.join(STUDENTS_IMPORT_DIR, f'{token}.json')


def _save_state(token, state):
    os.makedirs(STUDENTS_IMPORT_DIR, exist_ok=True)
    with open(_state_path(token), 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, default=str)


def _load_state(token):
    token = re.sub(r'[^0-9a-f]', '', token or '')
    if not token:
        return None
    path = _state_path(token)
    if not os.path.isfile(path):
        return None
    with open(path, encoding='utf-8') as f:
        return json.load(f)


# ═══════════════════════════════════════════════════════════════════════
# Визначення цільового enrollment'а для enrollment-полів
# ═══════════════════════════════════════════════════════════════════════

def _determine_target_enrollment(conn, student_id, group_from_excel):
    """
    Повертає dict:
        {
            'target_group_id': int | None,
            'groups_available': [{'group_id', 'group_name', 'is_primary'}],
            'needs_choice': bool,
            'reason': 'excel_column_a' | 'single_enrollment' | 'multiple_enrollments' | 'no_active_enrollments',
        }

    Правила:
      - Колонка A заповнена → target = A, needs_choice=False (довіряємо Excel).
      - Колонка A порожня + 1 активний enrollment → target = він, needs_choice=False.
      - Колонка A порожня + >1 → target=None, needs_choice=True (radio у прев'ю).
      - Немає активних enrollment'ів → target=None, needs_choice=False.
    """
    if group_from_excel:
        return {
            'target_group_id': int(group_from_excel),
            'groups_available': [],
            'needs_choice': False,
            'reason': 'excel_column_a',
        }

    rows = conn.execute("""
        SELECT e.group_id, e.is_primary, g.name AS group_name
        FROM student_enrollments e
        JOIN groups g ON g.id = e.group_id
        WHERE e.student_id = ? AND e.status = 'active'
        ORDER BY e.is_primary DESC, e.id
    """, (student_id,)).fetchall()

    if not rows:
        return {
            'target_group_id': None,
            'groups_available': [],
            'needs_choice': False,
            'reason': 'no_active_enrollments',
        }

    if len(rows) == 1:
        r = rows[0]
        return {
            'target_group_id': r['group_id'],
            'groups_available': [{
                'group_id': r['group_id'],
                'group_name': r['group_name'] or f"#{r['group_id']}",
                'is_primary': bool(r['is_primary']),
            }],
            'needs_choice': False,
            'reason': 'single_enrollment',
        }

    return {
        'target_group_id': None,
        'groups_available': [{
            'group_id': r['group_id'],
            'group_name': r['group_name'] or f"#{r['group_id']}",
            'is_primary': bool(r['is_primary']),
        } for r in rows],
        'needs_choice': True,
        'reason': 'multiple_enrollments',
    }


# ═══════════════════════════════════════════════════════════════════════
# Головний парсер: Excel → preview (БЕЗ ЗАПИСУ В БД)
# ═══════════════════════════════════════════════════════════════════════

def _build_import_preview(filepath, conn, allowed_group_ids):
    """
    Читає Excel, для кожного рядка будує preview-структуру.
    НЕ ПИШЕ В БД — тільки збирає дані.

    Повертає (preview, global_errors):
      preview       — список dict (по одному на рядок Excel)
      global_errors — список dict (сканування файлу цілком, напр. «порожній файл»)
    """
    preview = []
    global_errors = []

    license_lookup = {}
    for r in conn.execute("SELECT id, short_name_ua FROM institution_licenses").fetchall():
        if r['short_name_ua']:
            license_lookup[r['short_name_ua'].strip().lower()] = r['id']

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        sheet = wb.active
    except Exception as e:
        logger.error(f'[import_preview] open xlsx: {e}', exc_info=True)
        global_errors.append({'reason': f'Не вдалося відкрити файл: {e}'})
        return preview, global_errors

    row_counter = 0
    try:
        for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            row_counter += 1
            if not row or all(cell is None for cell in row):
                continue

            item = _build_row_preview(i, row, conn, allowed_group_ids, license_lookup)
            if item:
                preview.append(item)
    finally:
        try:
            wb.close()
        except Exception:
            pass

    if row_counter == 0:
        global_errors.append({'reason': 'У файлі немає рядків з даними'})

    return preview, global_errors


def _build_row_preview(i, row, conn, allowed_group_ids, license_lookup):
    """
    Парсить один рядок Excel → повертає preview_item або None (якщо рядок
    повністю невалідний і його треба пропустити — тоді повертає item з error).
    """
    from routes.utils import normalize_phone_pair

    # ── Базові колонки A-K ──
    group_id_raw = row[0] if len(row) > 0 else None
    full_name = row[1] if len(row) > 1 else None
    birth_date_raw = row[2] if len(row) > 2 else None
    edebo_code = (
        _normalize_excel_code(row[3])
        if len(row) > 3 and row[3] not in (None, '') else ''
    )

    # Колонка E: ліцензія
    license_id = None
    license_name = None
    if len(row) > 4 and row[4] not in (None, ''):
        lk = str(row[4]).strip().lower()
        license_id = license_lookup.get(lk)
        license_name = str(row[4]).strip() if license_id is None else str(row[4]).strip()

    # Колонка F: кредити override
    program_credits_override = None
    if len(row) > 5 and row[5] not in (None, ''):
        try:
            program_credits_override = int(row[5])
        except (ValueError, TypeError):
            program_credits_override = None

    # Колонка G: телефони
    phone = None
    phone_backup = None
    if len(row) > 6 and row[6] not in (None, ''):
        parts = [p.strip() for p in str(row[6]).split(',') if p.strip()]
        if parts:
            phone, phone_backup = normalize_phone_pair(
                parts[0], parts[1] if len(parts) > 1 else None,
            )

    email = (str(row[7]).strip()
             if len(row) > 7 and row[7] not in (None, '') else None)
    tax_id = (
        _normalize_excel_code(row[8]) or None
        if len(row) > 8 and row[8] not in (None, '') else None
    )
    address = (str(row[9]).strip()
               if len(row) > 9 and row[9] not in (None, '') else None)
    gender = (_normalize_gender(row[10]) if len(row) > 10 else None)

    # ── Опціональні L-N ──
    force_student_id = None
    if len(row) > 11 and row[11] not in (None, ''):
        try:
            force_student_id = int(row[11])
        except (ValueError, TypeError):
            pass

    force_is_primary = None
    if len(row) > 12 and row[12] not in (None, ''):
        raw = str(row[12]).strip().lower()
        if raw in ('1', 'true', 'yes', 'так'):
            force_is_primary = True
        elif raw in ('0', 'false', 'no', 'ні'):
            force_is_primary = False

    update_existing_data = True
    if len(row) > 13 and row[13] not in (None, ''):
        raw = str(row[13]).strip().lower()
        if raw in ('0', 'false', 'no', 'ні'):
            update_existing_data = False

    # ── Порожнє ПІБ — обов'язковий пропуск ──
    if not full_name:
        return {
            'row': i, 'status': 'error',
            'fio_excel': None, 'student_id': None, 'student_fio': None,
            'error': 'Порожнє ПІБ (колонка B)',
            'warnings': [],
            'new_student': None, 'enrollment': None,
            'card_fields': [], 'enrollment_fields': [],
        }

    # ── Розбір ПІБ ──
    name_parts = str(full_name).strip().split()
    if len(name_parts) < 3:
        return {
            'row': i, 'status': 'error',
            'fio_excel': str(full_name).strip(),
            'student_id': None, 'student_fio': None,
            'error': f'Невірний формат ПІБ: «{full_name}» (мін. 3 слова)',
            'warnings': [], 'new_student': None, 'enrollment': None,
            'card_fields': [], 'enrollment_fields': [],
        }
    last_name = name_parts[0]
    middle_name = name_parts[-1]
    first_name = (name_parts[1] if len(name_parts) == 3
                  else ' '.join(name_parts[1:-1]))

    if gender is None:
        gender = detect_gender_from_middle_name(middle_name)

    # ── Дата народження ──
    if isinstance(birth_date_raw, datetime):
        birth_date = birth_date_raw.strftime('%d.%m.%Y')
    elif birth_date_raw in (None, ''):
        birth_date = ''
    else:
        birth_date = str(birth_date_raw).strip()

    # ── Група A: update-only режим або валідація ──
    update_only_mode = False
    group_id = None
    group_name = None
    group_form = None

    if group_id_raw in (None, ''):
        update_only_mode = True
    else:
        try:
            group_id = int(group_id_raw)
        except (ValueError, TypeError):
            return {
                'row': i, 'status': 'error',
                'fio_excel': str(full_name).strip(),
                'student_id': None, 'student_fio': None,
                'error': f'Некоректний ID групи в колонці A: «{group_id_raw}»',
                'warnings': [], 'new_student': None, 'enrollment': None,
                'card_fields': [], 'enrollment_fields': [],
            }
        if group_id not in allowed_group_ids:
            return {
                'row': i, 'status': 'error',
                'fio_excel': str(full_name).strip(),
                'student_id': None, 'student_fio': None,
                'error': f'Група ID {group_id} не існує або недоступна',
                'warnings': [], 'new_student': None, 'enrollment': None,
                'card_fields': [], 'enrollment_fields': [],
            }
        g_row = conn.execute(
            "SELECT name, study_form FROM groups WHERE id = ?", (group_id,)
        ).fetchone()
        if g_row:
            group_name = g_row['name']
            group_form = g_row['study_form']

    # ── Пошук існуючого студента ──
    SELECT_EXISTING = """
        SELECT id, group_id, archived,
               phone, phone_backup, email, tax_id,
               edebo_code, address_of_residence, gender,
               program_credits_override, license_id,
               last_name_UA, first_name_UA, middle_name_UA,
               birth_date
        FROM students
    """
    existing = None
    warnings = []

    if force_student_id is not None:
        existing = conn.execute(
            SELECT_EXISTING + " WHERE id = ?", (force_student_id,)
        ).fetchone()
        if not existing:
            return {
                'row': i, 'status': 'error',
                'fio_excel': str(full_name).strip(),
                'student_id': None, 'student_fio': None,
                'error': f'student_id={force_student_id} не існує в базі',
                'warnings': [], 'new_student': None, 'enrollment': None,
                'card_fields': [], 'enrollment_fields': [],
            }
    else:
        if tax_id:
            existing = conn.execute(
                SELECT_EXISTING + " WHERE tax_id = ?", (tax_id,)
            ).fetchone()

        if not existing and birth_date:
            existing = conn.execute(
                SELECT_EXISTING + """
                WHERE NORM_NAME(
                          last_name_UA || ' ' || first_name_UA
                          || ' ' || COALESCE(middle_name_UA, '')
                      ) = NORM_NAME(?)
                  AND REPLACE(
                          REPLACE(
                              REPLACE(COALESCE(birth_date, ''), '-', ''),
                              '/', ''),
                          '.', '')
                      = REPLACE(
                          REPLACE(
                              REPLACE(?, '-', ''),
                              '/', ''),
                          '.', '')
                """,
                (str(full_name).strip(), birth_date),
            ).fetchone()

        if not existing and update_only_mode:
            candidates = conn.execute(
                SELECT_EXISTING + """
                WHERE NORM_NAME(
                          last_name_UA || ' ' || first_name_UA
                          || ' ' || COALESCE(middle_name_UA, '')
                      ) = NORM_NAME(?)
                """,
                (str(full_name).strip(),),
            ).fetchall()

            if len(candidates) == 1:
                existing = candidates[0]
                if birth_date:
                    warnings.append(
                        f"Точного збігу за ФІО+датою не знайдено, "
                        f"але за ПІБ — рівно один студент (ID {existing['id']}). "
                        f"Перевірте, чи це та людина."
                    )
            elif len(candidates) > 1:
                return {
                    'row': i, 'status': 'error',
                    'fio_excel': str(full_name).strip(),
                    'student_id': None, 'student_fio': None,
                    'error': (f"Знайдено {len(candidates)} студентів з таким ПІБ — "
                              f"уточніть дату народження, ІПН (кол. I) або student_id (кол. L)"),
                    'warnings': [], 'new_student': None, 'enrollment': None,
                    'card_fields': [], 'enrollment_fields': [],
                }

    # ═════════════════════════════════════════════════════════════════
    # СЦЕНАРІЙ 1: Новий студент (не знайдено в БД)
    # ═════════════════════════════════════════════════════════════════
    if not existing:
        if update_only_mode:
            return {
                'row': i, 'status': 'error',
                'fio_excel': f"{last_name} {first_name} {middle_name}",
                'student_id': None, 'student_fio': None,
                'error': ('Студента не знайдено в базі, а група (колонка A) '
                          'не вказана — для нового студента група обов\'язкова'),
                'warnings': [], 'new_student': None, 'enrollment': None,
                'card_fields': [], 'enrollment_fields': [],
            }

        last_name_eng, first_name_eng = generate_english_name(last_name, first_name)

        new_student_fields = {
            'last_name_UA': last_name,
            'first_name_UA': first_name,
            'middle_name_UA': middle_name,
            'last_name_ENG': last_name_eng,
            'first_name_ENG': first_name_eng,
            'birth_date': birth_date,
            'edebo_code': edebo_code or None,
            'license_id': license_id,
            'program_credits_override': program_credits_override,
            'phone': phone,
            'phone_backup': phone_backup,
            'email': email,
            'tax_id': tax_id,
            'address_of_residence': address,
            'gender': gender,
        }

        return {
            'row': i, 'status': 'new',
            'fio_excel': f"{last_name} {first_name} {middle_name}",
            'student_id': None, 'student_fio': None,
            'error': None,
            'warnings': warnings,
            'new_student': {
                'enabled': True,
                'fields': new_student_fields,
                'group_id': group_id,
                'group_name': group_name,
                'group_form': group_form,
                'license_id': license_id,
                'license_name': license_name,
            },
            'enrollment': {
                'enabled': True,
                'action': 'create',
                'group_id': group_id,
                'group_name': group_name,
                'group_form': group_form,
                'is_primary': True,  # новий студент → перша група = основна
                'license_id': license_id,
                'license_name': license_name,
                'edebo_code': edebo_code or None,
                'program_credits_override': program_credits_override,
            },
            'card_fields': [],
            'enrollment_fields': [],
        }

    # ═════════════════════════════════════════════════════════════════
    # СЦЕНАРІЙ 2: Студент вже є в БД
    # ═════════════════════════════════════════════════════════════════
    sid = existing['id']
    student_fio = (
        f"{existing['last_name_UA']} {existing['first_name_UA']} "
        f"{existing['middle_name_UA'] or ''}".strip()
    )

    # ── Enrollment ──
    enrollment = None
    if not update_only_mode:
        enroll_row = conn.execute(
            "SELECT status FROM student_enrollments "
            "WHERE student_id = ? AND group_id = ?",
            (sid, group_id),
        ).fetchone()

        if enroll_row and enroll_row['status'] == 'active':
            # Сценарій A: вже активний → пропуск рядка
            return {
                'row': i, 'status': 'error',
                'fio_excel': f"{last_name} {first_name} {middle_name}",
                'student_id': sid, 'student_fio': student_fio,
                'error': f'Вже активно навчається в цій групі (студент ID {sid})',
                'warnings': [], 'new_student': None, 'enrollment': None,
                'card_fields': [], 'enrollment_fields': [],
            }

        was_reactivation = bool(
            enroll_row and enroll_row['status'] != 'active'
        )
        is_primary_new = (
            force_is_primary if force_is_primary is not None else False
        )

        enrollment = {
            'enabled': True,
            'action': 'reactivate' if was_reactivation else 'create',
            'group_id': group_id,
            'group_name': group_name,
            'group_form': group_form,
            'is_primary': is_primary_new,
            'license_id': license_id,
            'license_name': license_name,
            'edebo_code': edebo_code or None,
            'program_credits_override': program_credits_override,
        }
    else:
        # update-only: enrollment не створюємо, але користувачу треба
        # показати, куди писатимуться enrollment-поля (edebo, credits).
        enrollment = None

    # ── Поля картки студента (students.*) ──
    card_fields = []
    if update_existing_data:
        field_candidates = {
            'last_name_UA': last_name,
            'first_name_UA': first_name,
            'middle_name_UA': middle_name,
            'birth_date': birth_date,
            'phone': phone,
            'phone_backup': phone_backup,
            'email': email,
            'tax_id': tax_id,
            'address_of_residence': address,
            'gender': gender,
        }
        card_fields = _collect_field_diffs(existing, field_candidates)

    # ── Поля enrollment'а (edebo_code, program_credits_override) ──
    enrollment_fields = []
    target_info = _determine_target_enrollment(conn, sid, group_id)

    if update_existing_data:
        # edebo_code
        if edebo_code:
            cur_edc = (
                str(existing['edebo_code'] or '').strip()
                if not update_only_mode
                else None
            )

            # Визначаємо поточне значення у цільовому enrollment'ї
            target_gid_for_check = (
                target_info['target_group_id']
                or (group_id if group_id else None)
            )
            if target_gid_for_check:
                row_edc = conn.execute(
                    "SELECT edebo_code FROM student_enrollments "
                    "WHERE student_id = ? AND group_id = ?",
                    (sid, target_gid_for_check),
                ).fetchone()
                cur_edc = str(row_edc['edebo_code'] or '').strip() if row_edc else ''

            new_edc = str(edebo_code).strip()
            if cur_edc != new_edc:
                enrollment_fields.append({
                    'field': 'edebo_code',
                    'label': ENROLLMENT_FIELD_LABELS['edebo_code'],
                    'old': cur_edc,
                    'new': new_edc,
                    'type': 'fill_empty' if not cur_edc else 'conflict',
                    'enabled': (not cur_edc),
                    'target_group_id': target_info['target_group_id'],
                    'groups_available': target_info['groups_available'],
                    'needs_choice': target_info['needs_choice'],
                })

        # program_credits_override — тільки якщо буде enrollment
        if program_credits_override and not update_only_mode:
            row_cred = conn.execute(
                "SELECT program_credits_override FROM student_enrollments "
                "WHERE student_id = ? AND group_id = ?",
                (sid, group_id),
            ).fetchone()
            cur_cred = (
                row_cred['program_credits_override']
                if row_cred else None
            )

            if cur_cred != program_credits_override:
                enrollment_fields.append({
                    'field': 'program_credits_override',
                    'label': ENROLLMENT_FIELD_LABELS['program_credits_override'],
                    'old': str(cur_cred) if cur_cred else '',
                    'new': str(program_credits_override),
                    'type': 'fill_empty' if not cur_cred else 'conflict',
                    'enabled': (not cur_cred),
                    'target_group_id': group_id,  # credits пишуться ТІЛЬКИ в group_id з A
                    'groups_available': [],
                    'needs_choice': False,
                })

    status = 'existing'

    return {
        'row': i, 'status': status,
        'fio_excel': f"{last_name} {first_name} {middle_name}",
        'student_id': sid, 'student_fio': student_fio,
        'error': None,
        'warnings': warnings,
        'new_student': None,
        'enrollment': enrollment,
        'card_fields': card_fields,
        'enrollment_fields': enrollment_fields,
    }


# ═══════════════════════════════════════════════════════════════════════
# Маршрут 1: POST /import_from_excel — парсинг + редирект на preview
# ═══════════════════════════════════════════════════════════════════════

@students_bp.route('/import_from_excel', methods=['GET', 'POST'])
@permission_required('import_from_excel')
def import_from_excel():
    """
    GET  — показує форму завантаження файлу.
    POST — приймає файл, будує preview (БЕЗ запису), редіректить на
           /import_from_excel/preview/<token>.
    """
    _cleanup_old_states()

    if request.method == 'GET':
        return render_template('import_excel.html')

    file = request.files.get('excel_file')
    if not file or not file.filename.lower().endswith('.xlsx'):
        flash('Оберіть, будь ласка, файл формату .xlsx', 'error')
        return redirect(url_for('students.import_from_excel'))

    filename = secure_filename(file.filename)
    filepath = os.path.join('uploads', f'{uuid.uuid4().hex}_{filename}')
    os.makedirs('uploads', exist_ok=True)
    file.save(filepath)

    conn = get_db()
    ensure_enrollments_table(conn)
    try:
        role = session.get('role')
        if role == 'admin':
            allowed_group_ids = {
                r['id'] for r in conn.execute(
                    "SELECT id FROM groups WHERE archived = FALSE"
                ).fetchall()
            }
        else:
            user_group_ids = session.get('group_ids', []) or []
            if user_group_ids:
                ph = ','.join('?' * len(user_group_ids))
                allowed_group_ids = {
                    r['id'] for r in conn.execute(
                        f"SELECT id FROM groups WHERE id IN ({ph}) "
                        f"AND archived = FALSE", user_group_ids,
                    ).fetchall()
                }
            else:
                allowed_group_ids = set()

        preview, global_errors = _build_import_preview(
            filepath, conn, allowed_group_ids,
        )
    finally:
        try:
            os.remove(filepath)
        except OSError:
            pass
        conn.close()

    if global_errors:
        for ge in global_errors:
            flash(ge['reason'], 'error')
        return redirect(url_for('students.import_from_excel'))

    if not preview:
        flash('У файлі не знайдено жодного рядка для обробки', 'warning')
        return redirect(url_for('students.import_from_excel'))

    token = uuid.uuid4().hex[:32]
    _save_state(token, {
        'preview': preview,
        'filename': filename,
        'user_id': session.get('user_id'),
        'allowed_group_ids': list(allowed_group_ids),
        'created_at': time.time(),
    })

    return redirect(url_for('students.import_preview', token=token))


# ═══════════════════════════════════════════════════════════════════════
# Маршрут 2: GET /import_from_excel/preview/<token>
# ═══════════════════════════════════════════════════════════════════════

@students_bp.route('/import_from_excel/preview/<token>')
@permission_required('import_from_excel')
def import_preview(token):
    """Сторінка прев'ю: усі зміни з чекбоксами/radio."""
    state = _load_state(token)
    if not state:
        flash('Сесію імпорту не знайдено або вона застаріла', 'warning')
        return redirect(url_for('students.import_from_excel'))

    # Захист: чужа сесія (не адмін)
    if state.get('user_id') != session.get('user_id') and not session.get('is_admin'):
        flash('Сесію імпорту створено іншим користувачем', 'error')
        return redirect(url_for('students.import_from_excel'))

    preview = state['preview']

    # Статистика для верхньої панелі
    stats = {
        'new_students': sum(1 for p in preview if p['new_student']),
        'enrollments': sum(1 for p in preview if p['enrollment']),
        'updates': sum(1 for p in preview
                       if p['card_fields'] or p['enrollment_fields']),
        'errors': sum(1 for p in preview if p['error']),
    }

    return render_template(
        'import_excel_preview.html',
        token=token,
        preview=preview,
        stats=stats,
        filename=state.get('filename', ''),
    )


# ═══════════════════════════════════════════════════════════════════════
# Маршрут 3: POST /import_from_excel/commit/<token>
# ═══════════════════════════════════════════════════════════════════════

@students_bp.route('/import_from_excel/commit/<token>', methods=['POST'])
@permission_required('import_from_excel')
def import_commit(token):
    """Застосовує лише те, що користувач позначив чекбоксами у прев'ю."""
    state = _load_state(token)
    if not state:
        flash('Сесію імпорту не знайдено або вона застаріла', 'warning')
        return redirect(url_for('students.import_from_excel'))

    if state.get('user_id') != session.get('user_id') and not session.get('is_admin'):
        flash('Сесію імпорту створено іншим користувачем', 'error')
        return redirect(url_for('students.import_from_excel'))

    allowed_group_ids = set(state.get('allowed_group_ids') or [])

    # ── Валідація: чи всі "needs_choice" поля мають обрану групу ──
    missing_choices = []
    for item in state['preview']:
        if item.get('error'):
            continue
        for ef in (item.get('enrollment_fields') or []):
            if not ef.get('needs_choice'):
                continue
            field_key = f'apply_enroll_{item["row"]}_{ef["field"]}'
            target_key = f'target_group_{item["row"]}_{ef["field"]}'
            if request.form.get(field_key) and not request.form.get(target_key):
                missing_choices.append(
                    f"рядок {item['row']}, поле «{ef['label']}» — не обрано групу"
                )

    if missing_choices:
        flash(
            'Не обрано групу для деяких полів: ' + '; '.join(missing_choices),
            'error',
        )
        return redirect(url_for('students.import_preview', token=token))

    conn = get_db()
    ensure_enrollments_table(conn)

    stats = {
        'students_created': 0,
        'enrollments_created': 0,
        'enrollments_reactivated': 0,
        'card_fields_updated': 0,
        'enrollment_fields_updated': 0,
        'skipped_by_user': 0,
        'skipped_errors': 0,
        'warnings': [],
    }

    try:
        for item in state['preview']:
            row = item['row']

            if item.get('error'):
                stats['skipped_errors'] += 1
                continue

            student_id = item.get('student_id')

            # ── 1. Новий студент ──
            if item.get('new_student'):
                if not request.form.get(f'apply_new_student_{row}'):
                    stats['skipped_by_user'] += 1
                    # Якщо новий не створюємо — enrollment теж пропускаємо
                    continue
                ns = item['new_student']
                f = ns['fields']
                try:
                    cur = conn.execute("""
                        INSERT INTO students (
                            last_name_UA, first_name_UA, middle_name_UA,
                            last_name_ENG, first_name_ENG, birth_date,
                            group_id, edebo_code, license_id,
                            program_credits_override, phone, phone_backup,
                            email, tax_id, address_of_residence, gender
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        f['last_name_UA'], f['first_name_UA'],
                        f['middle_name_UA'], f['last_name_ENG'],
                        f['first_name_ENG'], f['birth_date'],
                        ns['group_id'], f['edebo_code'], f['license_id'],
                        f['program_credits_override'], f['phone'],
                        f['phone_backup'], f['email'], f['tax_id'],
                        f['address_of_residence'], f['gender'],
                    ))
                    student_id = cur.lastrowid
                    stats['students_created'] += 1
                except Exception as e:
                    logger.error(f'[import_commit] INSERT student row {row}: {e}', exc_info=True)
                    stats['warnings'].append({
                        'row': row, 'reason': f'Не вдалося створити студента: {e}',
                    })
                    continue

            # ── 2. Enrollment ──
            if item.get('enrollment'):
                if not request.form.get(f'apply_enrollment_{row}'):
                    stats['skipped_by_user'] += 1
                    continue
                enr = item['enrollment']
                ok, msg = add_enrollment(
                    student_id, enr['group_id'],
                    is_primary=enr['is_primary'],
                    license_id=enr['license_id'],
                    edebo_code=enr['edebo_code'],
                    program_credits_override=enr['program_credits_override'],
                    conn=conn,
                )
                if not ok:
                    stats['warnings'].append({
                        'row': row, 'reason': f'Enrollment: {msg}',
                    })
                    continue
                if enr['action'] == 'create':
                    stats['enrollments_created'] += 1
                else:
                    stats['enrollments_reactivated'] += 1

            # ── 3. Поля картки студента ──
            if item.get('card_fields') and student_id:
                updates = {}
                for cf in item['card_fields']:
                    if request.form.get(f'apply_card_{row}_{cf["field"]}'):
                        updates[cf['field']] = cf['new']
                if updates:
                    try:
                        set_clause = ', '.join(f"{k} = ?" for k in updates)
                        conn.execute(
                            f"UPDATE students SET {set_clause} WHERE id = ?",
                            list(updates.values()) + [student_id],
                        )
                        stats['card_fields_updated'] += len(updates)
                    except Exception as e:
                        logger.error(f'[import_commit] UPDATE students row {row}: {e}', exc_info=True)
                        stats['warnings'].append({
                            'row': row, 'reason': f'Не вдалося оновити картку: {e}',
                        })

            # ── 4. Поля enrollment'а ──
            if item.get('enrollment_fields') and student_id:
                for ef in item['enrollment_fields']:
                    if not request.form.get(f'apply_enroll_{row}_{ef["field"]}'):
                        continue

                    # Куди писати: якщо needs_choice — з форми, інакше target_group_id з preview
                    if ef.get('needs_choice'):
                        target_gid = request.form.get(f'target_group_{row}_{ef["field"]}')
                        try:
                            target_gid = int(target_gid) if target_gid else None
                        except (ValueError, TypeError):
                            target_gid = None
                    else:
                        target_gid = ef.get('target_group_id')

                    if not target_gid:
                        stats['warnings'].append({
                            'row': row,
                            'field': ef['label'],
                            'reason': 'не визначено цільову групу — поле не записано',
                        })
                        continue

                    if int(target_gid) not in allowed_group_ids:
                        stats['warnings'].append({
                            'row': row,
                            'field': ef['label'],
                            'reason': f'група ID {target_gid} недоступна — пропущено',
                        })
                        continue

                    try:
                        if ef['field'] == 'edebo_code':
                            update_enrollment_edebo_code(
                                student_id, target_gid, ef['new'], conn=conn,
                            )
                        elif ef['field'] == 'program_credits_override':
                            try:
                                new_val = int(ef['new'])
                            except (ValueError, TypeError):
                                new_val = None
                            update_enrollment_program_credits(
                                student_id, target_gid, new_val, conn=conn,
                            )
                        stats['enrollment_fields_updated'] += 1
                    except Exception as e:
                        logger.error(f'[import_commit] enrollment field row {row}: {e}', exc_info=True)
                        stats['warnings'].append({
                            'row': row, 'field': ef['label'],
                            'reason': f'Помилка: {e}',
                        })

        conn.commit()
    except Exception as e:
        conn.rollback()
        logger.error(f'[import_commit] global: {e}', exc_info=True)
        flash(f'Помилка імпорту: {e}', 'error')
        conn.close()
        return redirect(url_for('students.import_preview', token=token))
    finally:
        conn.close()

    # Логуємо
    log_action(
        current_username(),
        f"імпорт студентів з Excel (підтверджено): "
        f"нових {stats['students_created']}, "
        f"enrollments {stats['enrollments_created'] + stats['enrollments_reactivated']}, "
        f"оновлено полів {stats['card_fields_updated'] + stats['enrollment_fields_updated']}, "
        f"пропущено {stats['skipped_by_user'] + stats['skipped_errors']}",
        details=f"файл: {state.get('filename')}, "
                f"попереджень: {len(stats['warnings'])}",
    )

    # Видаляємо стан (успіх чи ні — повторно не застосувати)
    try:
        os.remove(_state_path(token))
    except OSError:
        pass

    # Зберігаємо звіт для сторінки /report
    report_id = None
    if stats['warnings'] or stats['skipped_errors'] or stats['skipped_by_user']:
        os.makedirs(REPORTS_DIR, exist_ok=True)
        report_id = uuid.uuid4().hex
        with open(os.path.join(REPORTS_DIR, f'{report_id}.json'), 'w', encoding='utf-8') as f:
            json.dump({
                'filename': state.get('filename'),
                'stats': stats,
                'when': datetime.now().strftime('%d.%m.%Y %H:%M'),
                'skipped_rows': [
                    {'row': p['row'], 'fio': p.get('fio_excel'),
                     'reason': p.get('error')}
                    for p in state['preview'] if p.get('error')
                ],
            }, f, ensure_ascii=False)
        return redirect(url_for('students.import_report', report_id=report_id))

    # Інакше — коротке flash + редірект на список
    flash(
        f"✅ Імпорт завершено. "
        f"Нових студентів: {stats['students_created']}. "
        f"Навчань: {stats['enrollments_created'] + stats['enrollments_reactivated']}. "
        f"Оновлено полів: {stats['card_fields_updated'] + stats['enrollment_fields_updated']}.",
        'success',
    )
    return redirect(url_for('students.student_list'))


# ═══════════════════════════════════════════════════════════════════════
# Маршрут 4: GET /import_from_excel/report/<report_id>
# (спрощений — тільки статистика + skipped + warnings)
# ═══════════════════════════════════════════════════════════════════════

@students_bp.route('/import_from_excel/report/<report_id>')
@permission_required('import_from_excel')
def import_report(report_id):
    """Підсумковий звіт після commit."""
    report_id = re.sub(r'[^0-9a-f]', '', report_id or '')
    report_path = os.path.join(REPORTS_DIR, f'{report_id}.json')
    if not report_id or not os.path.isfile(report_path):
        flash('Звіт про імпорт не знайдено або він застарів', 'warning')
        return redirect(url_for('students.import_from_excel'))

    with open(report_path, encoding='utf-8') as f:
        report = json.load(f)

    return render_template(
        'import_excel_report.html',
        report=report,
        report_id=report_id,
    )