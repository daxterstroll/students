"""
routes/students/import_export.py
==================================
Масовий імпорт студентів з Excel-файлу (з нечітким/нормальним
зіставленням груп, генерацією англійського написання ПІБ тощо).
"""
from flask import render_template, request, redirect, url_for, session, flash
from routes.db import get_db
from routes.utils import log_action, login_required, permission_required, logger
from routes.helpers import current_username
from routes.students import students_bp
from routes.gender import detect_gender_from_middle_name
from routes.enrollments_service import add_enrollment, ensure_enrollments_table
import sqlite3
from routes.utils import generate_english_name
from werkzeug.utils import secure_filename
from datetime import datetime
import openpyxl
import os
import re
import json
import time


def _normalize_date_for_compare(s):
    """Нормалізує дату ДД.ММ.РРРР для порівняння (ігнорує роздільники)."""
    if not s:
        return ''
    s = str(s).strip()
    for sep in ('.', '/', '-'):
        s = s.replace(sep, '')
    return s


def _collect_field_conflicts(existing, new_values):
    """Порівнює значення з Excel з БД.

    Повертає (updates, conflicts):
      updates   — dict field → new_value (для порожніх полів БД — 
                  оновлюються автоматично, як і раніше);
      conflicts — список {'field', 'label', 'current', 'new'} —
                  поле заповнено в БД і ВІДРІЗНЯЄТЬСЯ від Excel.
                  Ці поля не змінюються автоматично — тільки через
                  кнопку «Замінити» в звіті імпорту.
    """
    from routes.db import normalize_name_for_match

    FIELD_LABELS = {
        'last_name_UA': 'Прізвище (укр.)',
        'first_name_UA': "Ім'я (укр.)",
        'middle_name_UA': 'По батькові (укр.)',
        'birth_date': 'Дата народження',
        'phone': 'Телефон',
        'phone_backup': 'Резервний телефон',
        'email': 'Email',
        'tax_id': 'ІПН',
        'edebo_code': 'ЄДЕБО',
        'address_of_residence': 'Адреса',
        'gender': 'Стать',
    }

    updates = {}
    conflicts = []

    for field, new_val in new_values.items():
        if not new_val:
            continue

        cur_raw = existing[field] if existing[field] is not None else ''

        # Нормалізація для порівняння (щоб не було «помилкових» конфліктів)
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
            # ФІО — через normalize_name_for_match: ігнорує апострофи
            # різних Unicode-варіантів, дефіс↔пробіл, регістр.
            cur_norm = normalize_name_for_match(cur_raw)
            new_norm = normalize_name_for_match(new_val)
        else:
            cur_norm = str(cur_raw).strip()
            new_norm = str(new_val).strip()

        if not cur_norm:
            # Пусто в БД → обновити автоматично
            updates[field] = new_val
        elif cur_norm != new_norm:
            # Заповнено і ВІДРІЗНЯЄТЬСЯ → конфлікт
            conflicts.append({
                'field': field,
                'label': FIELD_LABELS.get(field, field),
                'current': str(cur_raw or '—'),
                'new': str(new_val),
            })
        # cur_norm == new_norm → нічого, все співпадає

    return updates, conflicts

def _normalize_excel_code(raw):
    """Нормалізує «кодоподібне» значення з Excel (ЄДЕБО, ІПН, номер
    документа тощо) в рядок.

    openpyxl повертає числа як int/float, а не як рядок — а для кодів
    важливі саме цифри без зайвих «.0». Наприклад, якщо в Excel клітинка
    має числовий формат, значення 12345678 стане int 12345678, а
    12345678.5 — float. Без цієї функції int падає на .strip(), а float
    дає "12345678.0" — зіпсоване значення в БД.

    Правила:
      - None / '' → ''
      - int → str(int)                 (12345678 → "12345678")
      - float, що ціле → str(int)      (12345678.0 → "12345678")
      - float, неціле → str(float)     (рідкісний випадок, зберігаємо як є)
      - bool → str                     (рідкісний випадок)
      - інше → str(...).strip()
    """
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
    """
    Нормалізує значення колонки «Стать» з Excel до 'M'/'F' або None.

    Приймає: 'M', 'F', 'Ч', 'Ж', 'ч', 'ж', 'чоловіча', 'жіноча',
             'male', 'female', 'm', 'f' (без урахування регістру).
    Порожнє/невідоме → None (тоді визначається за по батькові).
    """
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


REPORTS_DIR = os.path.join(os.getcwd(), 'import_reports')
REPORTS_TTL_SECONDS = 24 * 3600  # звіти живуть 1 добу


def _cleanup_old_reports():
    """Видаляє звіти про імпорт, старші за добу (щоб папка не росла)."""
    if not os.path.isdir(REPORTS_DIR):
        return
    now = time.time()
    for f in os.listdir(REPORTS_DIR):
        path = os.path.join(REPORTS_DIR, f)
        try:
            if os.path.isfile(path) and now - os.path.getmtime(path) > REPORTS_TTL_SECONDS:
                os.remove(path)
        except OSError:
            pass


@students_bp.route('/import_from_excel', methods=['GET', 'POST'])
@permission_required('import_from_excel')
def import_from_excel():
    """
    Імпорт студентів з Excel-файлу.

    Колонки A-K — базові (обов'язкові/необов'язкові як раніше).
    Колонки L-N — нові, ОПЦІОНАЛЬНІ, для безпечної роботи з повторним
    навчанням / паралельними enrollments:

        L  student_id          — якщо вказано, студент шукається ТІЛЬКИ
                                 за цим ID (найнадійніше). Порожньо —
                                 пошук за tax_id, потім за ПІБ+датою.
        M  is_primary          — 1/0/порожньо. Якщо "1" — ця група стане
                                 основною для студента. Порожньо =
                                 автоопределение: для нового студента
                                 група стає основною, для існуючого —
                                 НЕ змінюється.
        N  update_existing_data — 1/0/порожньо. За замовчуванням "1":
                                 оновлює ПОРОЖНІ поля картки (телефон,
                                 email, ІПН, ЄДЕБО, адресу, кредити
                                 override, стать). НЕ перезаписує
                                 заповнені — щоб не затерти ручні правки.

    Підтримується 4 сценарії для вже наявного в базі студента:
        A. Активний enrollment у цій групі → пропуск (дублікат).
        B. Enrollment був закритий (expelled/graduated/frozen)
           → перевідкривається як active.
        C. Enrollment в іншій групі, у цій немає → паралельне навчання.
        D. Студент архівний (archived=TRUE) → enrollment створюється,
           студент разархівується (add_enrollment це робить сам).

    Для всіх B/C/D додатково, якщо update_existing_data != 0 —
    оновлюються порожні поля картки з файлу.
    """
    _cleanup_old_reports()
    if request.method == 'POST':
        file = request.files.get('excel_file')
        if not file or not file.filename.endswith('.xlsx'):
            flash("Будь ласка, виберіть файл формату .xlsx")
            return render_template('import_excel.html')

        filename = secure_filename(file.filename)
        filepath = os.path.join('uploads', filename)
        os.makedirs('uploads', exist_ok=True)
        file.save(filepath)

        conn = get_db()
        ensure_enrollments_table(conn)

        # ── Лічильники ──
        inserted = 0             # створено нових студентів
        enrollments_added = 0    # паралельних навчань у наявних
        enrollments_reactivated = 0  # перевідкрито закритих (B/D)
        data_updated = 0         # студентів, яким оновили дані картки
        skipped_rows = []        # [{row, fio, birth_date, reason}]
        warnings = []            # [{row, fio, reason}] — некритичні
        gender_auto = 0
        gender_missing = 0
        
        # ── Детальні списки успішних операцій (щоб бачити КОГО саме
        #    обробило, а не тільки сумарне число в flash-повідомленні) ──
        inserted_list = []               # [{row, fio, student_id, group_id}]
        enrollments_added_list = []      # [{row, fio, student_id, group_id, group_name}]
        enrollments_reactivated_list = []  # [{row, fio, student_id, group_id, group_name}]
        data_updated_list = []           # [{row, fio, student_id, fields}]

        role = session.get('role')
        user_group_ids = session.get('group_ids', [])

        if role == 'admin':
            allowed_group_ids = {
                r['id'] for r in conn.execute(
                    "SELECT id FROM groups WHERE archived = FALSE"
                ).fetchall()
            }
        else:
            if not user_group_ids:
                allowed_group_ids = set()
            else:
                ph = ','.join('?' * len(user_group_ids))
                allowed_group_ids = {
                    r['id'] for r in conn.execute(
                        f"SELECT id FROM groups WHERE id IN ({ph}) "
                        f"AND archived = FALSE", user_group_ids
                    ).fetchall()
                }

        license_lookup = {}
        for r in conn.execute(
                "SELECT id, short_name_ua FROM institution_licenses").fetchall():
            if r['short_name_ua']:
                license_lookup[r['short_name_ua'].strip().lower()] = r['id']

        try:
            wb = openpyxl.load_workbook(filepath)
            sheet = wb.active

            for i, row in enumerate(
                    sheet.iter_rows(min_row=2, values_only=True), start=2):
                try:
                    if not row or len(row) < 4:
                        skipped_rows.append({
                            'row': i, 'fio': '—', 'birth_date': '—',
                            'reason': 'Порожній рядок або менше 4 колонок',
                        })
                        continue

                    # ── A: ID групи ──
                    group_id = row[0]
                    # ── B: ПІБ ──
                    full_name = row[1]
                    # ── C: Дата народження ──
                    birth_date_raw = row[2]
                    # ── D: ЄДЕБО ──
                    # Через _normalize_excel_code, бо openpyxl може
                    # повернути int/float — і тоді .strip() на ньому
                    # падає, а float дає "12345678.0".
                    edebo_code = (
                        _normalize_excel_code(row[3])
                        if len(row) > 3 and row[3] not in (None, '')
                        else ''
                    )

                    # ── E: Ліцензія ──
                    license_id = None
                    if len(row) > 4 and row[4] not in (None, ''):
                        lk = str(row[4]).strip().lower()
                        license_id = license_lookup.get(lk)
                        if license_id is None:
                            warnings.append({
                                'row': i,
                                'fio': (str(full_name).strip()
                                        if full_name else '—'),
                                'reason': (f"Ліцензію '{row[4]}' не знайдено "
                                           f"в каталозі — enrollment створено "
                                           f"без ліцензії"),
                            })

                    # ── F: Кредити скороченої ──
                    program_credits_override = None
                    if len(row) > 5 and row[5] not in (None, ''):
                        try:
                            program_credits_override = int(row[5])
                        except (ValueError, TypeError):
                            warnings.append({
                                'row': i,
                                'fio': (str(full_name).strip()
                                        if full_name else '—'),
                                'reason': (f"Некоректне значення кредитів "
                                           f"скороченої програми '{row[5]}' "
                                           f"— поле пропущено"),
                            })

                    # ── G: Телефон ──
                    from routes.utils import normalize_phone_pair
                    phone = None
                    phone_backup = None
                    if len(row) > 6 and row[6] not in (None, ''):
                        parts = [p.strip()
                                 for p in str(row[6]).split(',') if p.strip()]
                        if parts:
                            phone, phone_backup = normalize_phone_pair(
                                parts[0],
                                parts[1] if len(parts) > 1 else None,
                            )

                    # ── H: Email ──
                    email = (str(row[7]).strip()
                             if len(row) > 7 and row[7] not in (None, '')
                             else None)

                    # ── I: ІПН ──
                    tax_id = (
                        _normalize_excel_code(row[8]) or None
                        if len(row) > 8 and row[8] not in (None, '')
                        else None
                    )

                    # ── J: Адреса ──
                    address_of_residence = (
                        str(row[9]).strip()
                        if len(row) > 9 and row[9] not in (None, '')
                        else None)

                    # ── K: Стать ──
                    gender = (_normalize_gender(row[10])
                              if len(row) > 10 else None)

                    if not full_name:
                        skipped_rows.append({
                            'row': i,
                            'fio': '—',
                            'birth_date': str(birth_date_raw or '').strip() or '—',
                            'reason': 'Порожнє ПІБ (колонка B) — рядок неможливо обробити',
                        })
                        continue

                    # ── Колонка A (ID групи) може бути порожньою — це
                    #    режим "лише оновити картку наявного студента"
                    #    (без створення/перевідкриття enrollment'а).
                    #    Корисно для масової правки телефонів/ІПН/адрес
                    #    по вже наявній базі без ризику зачепити навчання.
                    update_only_mode = False
                    if group_id in (None, ''):
                        update_only_mode = True
                        group_id = None
                    else:
                        try:
                            group_id = int(group_id)
                        except (ValueError, TypeError):
                            skipped_rows.append({
                                'row': i,
                                'fio': (str(full_name).strip()
                                        if full_name else '—'),
                                'birth_date': str(birth_date_raw or '').strip(),
                                'reason': f"Некоректний ID групи: '{row[0]}'",
                            })
                            continue

                        if group_id not in allowed_group_ids:
                            skipped_rows.append({
                                'row': i,
                                'fio': (str(full_name).strip()
                                        if full_name else '—'),
                                'birth_date': str(birth_date_raw or '').strip(),
                                'reason': f"Група ID {group_id} не існує "
                                          f"або недоступна",
                            })
                            continue

                    # ПІБ розбираємо з підтримкою 4+ слів — для випадків,
                    # коли в людини подвійне ім'я через пробіл, напр.
                    # "Шевченко Анна Марія Петрівна":
                    #   - перше слово — прізвище;
                    #   - останнє — по батькові;
                    #   - все між ними — ім'я.
                    # Це відповідає тому, як при пошуку NORM_NAME()
                    # з'єднує поля БД в один рядок — порядок збережено.
                    name_parts = str(full_name).strip().split()
                    if len(name_parts) < 3:
                        skipped_rows.append({
                            'row': i,
                            'fio': str(full_name).strip(),
                            'birth_date': str(birth_date_raw or '').strip(),
                            'reason': (f"Невірний формат ПІБ: «{full_name}» "
                                       f"(очікується мінімум 3 слова через пробіл)"),
                        })
                        continue
                    last_name = name_parts[0]
                    middle_name = name_parts[-1]
                    if len(name_parts) == 3:
                        first_name = name_parts[1]
                    else:
                        # 4+ слів: все між прізвищем і по батькові — ім'я
                        first_name = ' '.join(name_parts[1:-1])

                    if gender is None:
                        gender = detect_gender_from_middle_name(middle_name)
                        if gender is None:
                            gender_missing += 1
                        else:
                            gender_auto += 1

                    if isinstance(birth_date_raw, datetime):
                        birth_date = birth_date_raw.strftime("%d.%m.%Y")
                    elif birth_date_raw in (None, ''):
                        # Порожня клітинка — саме порожній рядок, а не
                        # str(None) = 'None'. Інакше пошук по ФІО без дати
                        # ніколи не спрацює: 'None' truthy, і код піде
                        # шукати студента з буквальною датою народження
                        # "None", не знайде і не запустить фолбек.
                        birth_date = ''
                    else:
                        birth_date = str(birth_date_raw).strip()

                    # ══════════════════════════════════════════════════════
                    # ── НОВІ опціональні колонки L-N ──
                    # ══════════════════════════════════════════════════════

                    # L: student_id — прямий ID (найнадійніше)
                    force_student_id = None
                    if len(row) > 11 and row[11] not in (None, ''):
                        try:
                            force_student_id = int(row[11])
                        except (ValueError, TypeError):
                            skipped_rows.append({
                                'row': i,
                                'fio': (str(full_name).strip()
                                        if full_name else '—'),
                                'birth_date': birth_date,
                                'reason': (f"Некоректний student_id у колонці L: "
                                           f"'{row[11]}' — очікується ціле число"),
                            })
                            continue

                    # M: is_primary — чи робити цю групу основною
                    force_is_primary = None
                    if len(row) > 12 and row[12] not in (None, ''):
                        raw = str(row[12]).strip().lower()
                        if raw in ('1', 'true', 'yes', 'так'):
                            force_is_primary = True
                        elif raw in ('0', 'false', 'no', 'ні'):
                            force_is_primary = False

                    # N: update_existing_data — оновлювати чи ні дані картки
                    update_existing_data = True
                    if len(row) > 13 and row[13] not in (None, ''):
                        raw = str(row[13]).strip().lower()
                        if raw in ('0', 'false', 'no', 'ні'):
                            update_existing_data = False

                    # ══════════════════════════════════════════════════════
                    # ── Пошук існуючого студента ──
                    # ══════════════════════════════════════════════════════
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

                    # L задано → шукаємо тільки по ID
                    if force_student_id is not None:
                        existing = conn.execute(
                            SELECT_EXISTING + " WHERE id = ?",
                            (force_student_id,),
                        ).fetchone()
                        if not existing:
                            skipped_rows.append({
                                'row': i,
                                'fio': (str(full_name).strip()
                                        if full_name else '—'),
                                'birth_date': birth_date,
                                'reason': (f"student_id={force_student_id} "
                                           f"не існує в базі"),
                            })
                            continue
                    else:
                        # ──────────────────────────────────────────────
                        # Пошук студента. Усі порівняння — нормалізовані:
                        #   - регістр (LOWER_UA обробляє кирилицю);
                        #   - пробіли по краях (TRIM);
                        #   - NULL у по батькові (COALESCE(..., ''));
                        #   - формат дати (. / - / пробіли ігноруються).
                        # Без цього: "Шевченко" ≠ "шевченко",
                        # "15.03.2001" ≠ "2001-03-15", NULL ≠ ''.
                        # ──────────────────────────────────────────────

                        # 1) tax_id → точний збіг
                        if tax_id:
                            existing = conn.execute(
                                SELECT_EXISTING + " WHERE tax_id = ?",
                                (tax_id,),
                            ).fetchone()

                        # 2) ФІО + дата (якщо дата вказана) — з
                        #    нормалізацією регістру, пробілів і формату
                        #    дати з обох боків.
                        if not existing and birth_date:
                            existing = conn.execute(
                                SELECT_EXISTING + """
                                WHERE NORM_NAME(
                                          last_name_UA || ' ' || first_name_UA
                                          || ' ' || COALESCE(middle_name_UA, '')
                                      )
                                      = NORM_NAME(?)
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

                        # 3) ФІО без дати (fallback). Два підсценарії:
                        #    a) дата в Excel порожня — звичайний фолбек;
                        #    b) дата в Excel заповнена, але НЕ співпала
                        #       з БД — пробуємо по ФІО, але з попередженням
                        #       «дата не співпала, знайдено за ПІБ».
                        #    Працює ТІЛЬКИ в update-only режимі (без
                        #    групи), інакше є ризик оновити не того
                        #    студента при ФІО-дублікаті.
                        if not existing and update_only_mode:
                            candidates = conn.execute(
                                SELECT_EXISTING + """
                                WHERE NORM_NAME(
                                          last_name_UA || ' ' || first_name_UA
                                          || ' ' || COALESCE(middle_name_UA, '')
                                      )
                                      = NORM_NAME(?)
                                """,
                                (str(full_name).strip(),),
                            ).fetchall()

                            if len(candidates) == 1:
                                existing = candidates[0]
                                # Якщо дата в Excel була, але ми знайшли
                                # по ФІО — це може бути:
                                # (а) в БД дата в іншому форматі (рідко);
                                # (б) в БД дата не заповнена;
                                # (в) в Excel дата не та (помилка).
                                # У всіх трьох випадках — попередження
                                # в звіті, щоб адмін перевірив.
                                if birth_date:
                                    warnings.append({
                                        'row': i,
                                        'fio': f"{last_name} {first_name} {middle_name}",
                                        'reason': (
                                            f"Точного збігу за ФІО+датою не знайдено, "
                                            f"але за ПІБ — рівно один студент "
                                            f"(ID {existing['id']}). "
                                            f"В Excel дата: «{birth_date}», "
                                            f"у БД: «{existing['birth_date'] or '—'}». "
                                            f"Оновлено за ПІБ — перевірте, чи це та людина."
                                        ),
                                        # ── Поля для кнопки "Оновити дату в БД"
                                        #    у звіті імпорту ──
                                        'action': 'update_birth_date',
                                        'student_id': existing['id'],
                                        'new_birth_date': birth_date,
                                        'current_birth_date': existing['birth_date'] or '',
                                        'resolved': False,
                                    })
                            elif len(candidates) > 1:
                                skipped_rows.append({
                                    'row': i,
                                    'fio': (str(full_name).strip()
                                            if full_name else '—'),
                                    'birth_date': birth_date or '—',
                                    'reason': (f"Знайдено {len(candidates)} "
                                               f"студентів з таким ПІБ — "
                                               f"уточніть дату народження, "
                                               f"ІПН (кол. I) або student_id (кол. L)"),
                                })
                                continue

                    # ══════════════════════════════════════════════════════
                    # ── СЦЕНАРІЙ 1: новий студент ──
                    # ══════════════════════════════════════════════════════
                    if not existing:
                        if update_only_mode:
                            # Не знаємо, в яку групу зарахувати нового
                            # студента — група для нових обов'язкова.
                            skipped_rows.append({
                                'row': i,
                                'fio': f"{last_name} {first_name} {middle_name}",
                                'birth_date': birth_date,
                                'reason': ("Студента не знайдено в базі, "
                                           "а група (колонка A) не вказана — "
                                           "для нового студента група обов'язкова"),
                            })
                            continue
                        last_name_eng, first_name_eng = generate_english_name(
                            last_name, first_name)
                        conn.execute("""
                            INSERT INTO students (
                                last_name_UA, first_name_UA, middle_name_UA,
                                last_name_ENG, first_name_ENG, birth_date,
                                group_id, edebo_code, license_id,
                                program_credits_override, phone, phone_backup,
                                email, tax_id, address_of_residence, gender
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            last_name, first_name, middle_name,
                            last_name_eng, first_name_eng, birth_date,
                            group_id, edebo_code, license_id,
                            program_credits_override, phone, phone_backup,
                            email, tax_id, address_of_residence, gender,
                        ))
                        new_id = conn.execute(
                            "SELECT last_insert_rowid()").fetchone()[0]
                        # Для нового студента ця група — завжди основна
                        add_enrollment(
                            new_id, group_id,
                            is_primary=True,
                            license_id=license_id,
                            program_credits_override=program_credits_override,
                            edebo_code=edebo_code, 
                            conn=conn,
                        )
                        inserted += 1
                        inserted_list.append({
                            'row': i,
                            'fio': f"{last_name} {first_name} {middle_name}",
                            'student_id': new_id,
                            'group_id': group_id,
                        })
                        continue

                    # ══════════════════════════════════════════════════════
                    # ── СЦЕНАРІЙ 2: людина вже в базі ──
                    # ══════════════════════════════════════════════════════
                    sid = existing['id']
                    was_archived = bool(existing['archived'])

                    if update_only_mode:
                        # Оновлюємо лише поля картки — enrollment не
                        # чіпаємо, наказ/групу/ліцензію не змінюємо.
                        enroll_row = None
                        was_reactivation = False
                    else:
                        # Перевіряємо поточний enrollment у цій групі
                        enroll_row = conn.execute("""
                            SELECT status FROM student_enrollments
                            WHERE student_id = ? AND group_id = ?
                        """, (sid, group_id)).fetchone()

                        # Сценарій A: активний enrollment → дублікат
                        if enroll_row and enroll_row['status'] == 'active':
                            skipped_rows.append({
                                'row': i,
                                'fio': f"{last_name} {first_name} {middle_name}",
                                'birth_date': birth_date,
                                'reason': (f"Вже активно навчається в цій групі "
                                           f"(студент ID {sid})"),
                            })
                            continue

                        # Сценарії B/C/D: створюємо / перевідкриваємо enrollment
                        was_reactivation = bool(
                            enroll_row and enroll_row['status'] != 'active'
                        )
                        is_primary = (
                            force_is_primary if force_is_primary is not None
                            else False  # наявного — primary не змінюємо автоматично
                        )

                        ok, msg = add_enrollment(
                            sid, group_id,
                            is_primary=is_primary,
                            license_id=license_id,
                            edebo_code=edebo_code,
                            conn=conn,
                        )
                        if not ok:
                            skipped_rows.append({
                                'row': i,
                                'fio': f"{last_name} {first_name} {middle_name}",
                                'birth_date': birth_date,
                                'reason': msg,
                            })
                            continue

                        # group_name для відображення в звіті
                        try:
                            gr_row = conn.execute(
                                "SELECT name FROM groups WHERE id = ?",
                                (group_id,),
                            ).fetchone()
                            gr_name = gr_row['name'] if gr_row else f'#{group_id}'
                        except Exception:
                            gr_name = f'#{group_id}'

                        entry = {
                            'row': i,
                            'fio': f"{last_name} {first_name} {middle_name}",
                            'student_id': sid,
                            'group_id': group_id,
                            'group_name': gr_name,
                        }
                        if was_reactivation:
                            enrollments_reactivated += 1
                            enrollments_reactivated_list.append(entry)
                        else:
                            enrollments_added += 1
                            enrollments_added_list.append(entry)

                    # ── Оновлення картки: пусті поля — автоматично,
                    #    заповнені-але-різні — в конфлікти (з кнопкою
                    #    «Замінити» в звіті). ──
                    if update_existing_data:
                        # Усі поля, які можна оновити з Excel. ФІО і
                        # дата народження — теж, бо якщо студент
                        # знайдено за ФІО+датою, то вони збіглися; якщо
                        # за student_id/tax_id — можуть відрізнятися
                        # (напр. виправляємо помилку в даті).
                        field_candidates = {
                            'last_name_UA': last_name,
                            'first_name_UA': first_name,
                            'middle_name_UA': middle_name,
                            'birth_date': birth_date,
                            'phone': phone,
                            'phone_backup': phone_backup,
                            'email': email,
                            'tax_id': tax_id,
                            # 'edebo_code' — ОКРЕМО, нижче:
                            # per-enrollment, не через UPDATE students
                            'address_of_residence': address_of_residence,
                            'gender': gender,
                        }

                        updates, conflicts = _collect_field_conflicts(
                            existing, field_candidates,
                        )

                        # program_credits_override — per-enrollment,
                        # оновлюємо окремо (не через UPDATE students).
                        if program_credits_override and not update_only_mode:
                            from routes.enrollments_service import (
                                update_enrollment_program_credits,
                            )
                            ok_pc, msg_pc = update_enrollment_program_credits(
                                sid, group_id, program_credits_override,
                                conn=conn,
                            )
                            if not ok_pc:
                                warnings.append({
                                    'row': i,
                                    'fio': f"{last_name} {first_name} {middle_name}",
                                    'reason': (f"Кредити скороченої програми "
                                               f"не оновлено: {msg_pc}"),
                                    'action': 'other',
                                })

                        # ── edebo_code — ОКРЕМО: per-enrollment, ──
                        # НЕ через UPDATE students (там лише кеш primary).
                        # Записуємо через update_enrollment_edebo_code() —
                        # ця функція обновлює enrollment.edebo_code і
                        # автоматично синхронізує students.edebo_code,
                        # якщо enrollment primary.
                        edebo_updated = False
                        if edebo_code:
                            from routes.enrollments_service import (
                                update_enrollment_edebo_code,
                            )

                            if update_only_mode:
                                # Цільова група — primary enrollment
                                target_enroll_row = conn.execute("""
                                    SELECT group_id, edebo_code
                                    FROM student_enrollments
                                    WHERE student_id = ?
                                      AND is_primary = 1
                                      AND status = 'active'
                                    LIMIT 1
                                """, (sid,)).fetchone()
                            else:
                                # Цільова група — конкретна з колонки A
                                target_enroll_row = conn.execute("""
                                    SELECT group_id, edebo_code
                                    FROM student_enrollments
                                    WHERE student_id = ?
                                      AND group_id = ?
                                      AND status = 'active'
                                    LIMIT 1
                                """, (sid, group_id)).fetchone()

                            if target_enroll_row:
                                tgt_gid = target_enroll_row['group_id']
                                # str() — на випадок, якщо в БД збережено
                                # як int (старі дані) або None.
                                cur_edc = str(
                                    target_enroll_row['edebo_code'] or ''
                                ).strip()
                                new_edc = str(edebo_code or '').strip()

                                if not cur_edc:
                                    # Порожньо в enrollment → оновити
                                    # (students.edebo_code синхронізується
                                    # автоматично, якщо primary)
                                    update_enrollment_edebo_code(
                                        sid, tgt_gid, edebo_code, conn=conn,
                                    )
                                    edebo_updated = True
                                elif cur_edc != new_edc:
                                    # Різні значення → конфлікт
                                    conflicts.append({
                                        'field': 'edebo_code',
                                        'label': 'ЄДЕБО',
                                        'current': cur_edc,
                                        'new': new_edc,
                                        'group_id': tgt_gid,
                                    })

                        # 1) Пусті поля БД — оновлюємо одразу
                        if updates:
                            set_clause = ', '.join(f"{k} = ?"
                                                   for k in updates)
                            conn.execute(
                                f"UPDATE students SET {set_clause} "
                                f"WHERE id = ?",
                                list(updates.values()) + [sid],
                            )

                        # data_updated +1 якщо хоч щось оновили
                        if updates or edebo_updated:
                            data_updated += 1
                            data_updated_list.append({
                                'row': i,
                                'fio': f"{last_name} {first_name} {middle_name}",
                                'student_id': sid,
                                'fields': list(updates.keys()) + (
                                    ['edebo_code'] if edebo_updated else []
                                ),
                            })

                        # 2) Заповнені-але-різні — conflict-warning з
                        #    кнопками «Замінити вибірково» в звіті.
                        if conflicts:
                            warnings.append({
                                'row': i,
                                'fio': f"{last_name} {first_name} {middle_name}",
                                'reason': (
                                    f"Студент ID {sid} — {len(conflicts)} "
                                    f"розбіжностей з БД. Можна замінити "
                                    f"вибірково."
                                ),
                                'action': 'replace_fields',
                                'student_id': sid,
                                'conflicts': conflicts,
                                'resolved': False,
                            })

                    # ── Попередження про архівного студента ──
                    if was_archived:
                        if update_only_mode:
                            warnings.append({
                                'row': i,
                                'fio': f"{last_name} {first_name} {middle_name}",
                                'reason': (f"Студент в архіві (ID {sid}) — "
                                           f"оновлено лише поля картки; "
                                           f"enrollment не створювався "
                                           f"(update-only режим). Для "
                                           f"разархівації вкажіть групу "
                                           f"в колонці A."),
                            })
                        else:
                            warnings.append({
                                'row': i,
                                'fio': f"{last_name} {first_name} {middle_name}",
                                'reason': (f"Студент був в архіві (ID {sid}) — "
                                           f"enrollment створено, студент "
                                           f"разархівовано"),
                            })

                except Exception as e:
                    logger.debug(
                        f"Пропущено рядок {i} при імпорті: {e}")
                    skipped_rows.append({
                        'row': i,
                        'fio': (full_name.strip()
                                if full_name else '—'),
                        'birth_date': str(birth_date_raw or '').strip(),
                        'reason': f"Помилка обробки: {e}",
                    })
                    continue

            wb.close()
            conn.commit()

        except Exception as e:
            conn.rollback()
            logger.error(
                f"Помилка при імпорті студентів з Excel "
                f"(файл: {filename}): {e}", exc_info=True)
            flash(f"⚠️ Помилка при читанні файлу: {e}")
        finally:
            conn.close()

        log_action(
            current_username(),
            f"імпорт студентів з Excel: нових {inserted}, "
            f"паралельних {enrollments_added}, "
            f"реактивовано {enrollments_reactivated}, "
            f"оновлено {data_updated}, пропущено {len(skipped_rows)}",
            details=f"файл: {filename}, стать авто: {gender_auto}, "
                    f"без статі: {gender_missing}, "
                    f"попереджень: {len(warnings)}",
        )

        # ── Фінальне повідомлення ──
        summary_parts = [
            f"нових: {inserted}",
            f"паралельних навчань: {enrollments_added}",
        ]
        if enrollments_reactivated:
            summary_parts.append(f"реактивовано: {enrollments_reactivated}")
        if data_updated:
            summary_parts.append(f"оновлено карток: {data_updated}")
        summary_parts.append(f"пропущено: {len(skipped_rows)}")
        flash("✅ Імпорт завершено. " + ", ".join(summary_parts), "success")

        if gender_auto:
            flash(f"Стать визначено автоматично: {gender_auto}", "info")
        if gender_missing:
            flash(f"⚠️ {gender_missing} студентів без по батькові — стать "
                  f"не визначена. Знайдіть їх у картці та заповніть вручну.",
                  "warning")
        if warnings:
            flash(f"⚠️ {len(warnings)} попереджень — див. детальний звіт.",
                  "warning")

        # ── Детальний звіт, якщо є пропуски або попередження ──
        if skipped_rows or warnings:
            import uuid as _uuid
            os.makedirs(REPORTS_DIR, exist_ok=True)
            report_id = _uuid.uuid4().hex
            report_path = os.path.join(REPORTS_DIR, f"{report_id}.json")
            with open(report_path, 'w', encoding='utf-8') as f:
                json.dump({
                    'filename': filename,
                    'inserted': inserted,
                    'enrollments_added': enrollments_added,
                    'enrollments_reactivated': enrollments_reactivated,
                    'data_updated': data_updated,
                    'skipped_rows': skipped_rows,
                    'warnings': warnings,
                    'gender_auto': gender_auto,
                    'gender_missing': gender_missing,
                    'when': datetime.now().strftime('%d.%m.%Y %H:%M'),
                    # ── Нове: детальні списки замість голих счётчиків ──
                    'inserted_list': inserted_list,
                    'enrollments_added_list': enrollments_added_list,
                    'enrollments_reactivated_list': enrollments_reactivated_list,
                    'data_updated_list': data_updated_list,
                }, f, ensure_ascii=False)
            return redirect(url_for(
                'students.import_report', report_id=report_id))

        return redirect(url_for('students.student_list'))

    return render_template('import_excel.html')
    
@students_bp.route('/import_from_excel/report/<report_id>')
@permission_required('import_from_excel')
def import_report(report_id):
    """Детальний звіт про імпорт студентів з Excel (хто пропущений і чому).

    Звіт зберігається в окремому JSON-файлі (а не в cookie-сесії), бо при
    196 пропущених рядках сесія Flask перевищує ліміт cookie (4 КБ) —
    браузер мовчки відкидає таку cookie, і звіт губиться.
    """
    # Захист від path traversal: лише hex-символи
    report_id = re.sub(r'[^0-9a-f]', '', report_id or '')
    report_path = os.path.join(REPORTS_DIR, f"{report_id}.json")
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


@students_bp.route('/import_from_excel/report/<report_id>/apply-birth-date', methods=['POST'],)
@permission_required('import_from_excel')
def import_report_apply_birth_date(report_id):
    """Підтвердити оновлення дати народження в БД за звітом імпорту.

    Використовується для тих рядків, де при імпорті не співпала дата
    (напр. у БД була порожня, а в Excel — заповнена). Замість
    повторного імпорту Excel адмін натискає кнопку в звіті — і ми
    записуємо конкретну дату конкретному студенту.

    Захист: тільки для warnings з action='update_birth_date';
    resolved-записи не перезаписуються (ідемпотентно).
    """
    report_id = re.sub(r'[^0-9a-f]', '', report_id or '')
    report_path = os.path.join(REPORTS_DIR, f"{report_id}.json")
    if not report_id or not os.path.isfile(report_path):
        flash('Звіт про імпорт не знайдено або він застарів', 'warning')
        return redirect(url_for('students.import_from_excel'))

    row_num = request.form.get('row', type=int)
    if not row_num:
        flash('Не вказано рядок звіту', 'error')
        return redirect(url_for('students.import_report', report_id=report_id))

    with open(report_path, encoding='utf-8') as f:
        report = json.load(f)

    # Знайти саме той warning
    target = None
    for w in report.get('warnings', []):
        if w.get('row') == row_num and w.get('action') == 'update_birth_date':
            target = w
            break

    if not target:
        flash('Запис для цього рядка не знайдено або він не підтримує '
              'оновлення дати', 'warning')
        return redirect(url_for('students.import_report', report_id=report_id))

    if target.get('resolved'):
        flash('Дату вже було оновлено раніше — повторно не змінюємо', 'info')
        return redirect(url_for('students.import_report', report_id=report_id))

    sid = target.get('student_id')
    new_bd = target.get('new_birth_date')
    if not sid or not new_bd:
        flash('Некоректні дані запису (немає student_id або new_birth_date)',
              'error')
        return redirect(url_for('students.import_report', report_id=report_id))

    conn = get_db()
    try:
        student = conn.execute(
            "SELECT last_name_UA, first_name_UA, birth_date "
            "FROM students WHERE id = ?",
            (sid,),
        ).fetchone()
        if not student:
            flash(f'Студента ID {sid} не знайдено (можливо, видалений)', 'error')
            return redirect(url_for('students.import_report', report_id=report_id))

        # Перечитуємо СВІЖЕ значення з БД (а не з звіту) — на випадок,
        # якщо хтось уже виправив дату вручну після імпорту.
        old_bd = (student['birth_date'] or '').strip()

        if old_bd == new_bd:
            # Уже актуально — просто позначаємо resolved, не UPDATE.
            target['resolved'] = True
            target['resolved_at'] = datetime.now().strftime('%d.%m.%Y %H:%M')
            target['resolved_by'] = current_username()
            target['resolution_note'] = 'Дата в БД вже така сама'
            with open(report_path, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False)
            flash('Дата в БД вже така сама, як у Excel — нічого не змінено',
                  'info')
            return redirect(url_for('students.import_report', report_id=report_id))

        conn.execute(
            "UPDATE students SET birth_date = ? WHERE id = ?",
            (new_bd, sid),
        )
        conn.commit()

        target['resolved'] = True
        target['resolved_at'] = datetime.now().strftime('%d.%m.%Y %H:%M')
        target['resolved_by'] = current_username()
        target['old_birth_date_at_apply'] = old_bd
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False)

        fio = f"{student['last_name_UA']} {student['first_name_UA']}".strip()
        log_action(
            current_username(),
            f'оновив дату народження студента {fio} (ID {sid}) '
            f'за звітом імпорту',
            details=f'{old_bd or "—"} → {new_bd}',
        )
        flash(f'Дату народження оновлено: «{old_bd or "—"}» → «{new_bd}»',
              'success')

    except Exception as e:
        conn.rollback()
        logger.error(
            f'import_report_apply_birth_date failed '
            f'(report={report_id}, student={sid}): {e}',
            exc_info=True,
        )
        flash(f'Помилка оновлення: {e}', 'error')
    finally:
        conn.close()

    return redirect(url_for('students.import_report', report_id=report_id))
    
@students_bp.route(
    '/import_from_excel/report/<report_id>/apply-fields',
    methods=['POST'],
)
@permission_required('import_from_excel')
def import_report_apply_fields(report_id):
    """Замінити вибрані поля з конфліктів (action='replace_fields').

    Адмін у звіті бачить таблицю конфліктів для кожного студента:
        Поле | Було в БД | Стане (з Excel) | ☑
    Обирає чекбоксами, що замінити, і натискає «Замінити обрані».

    ФІО оновлюється разом з ENG-варіантами (транслітерація).
    """
    report_id = re.sub(r'[^0-9a-f]', '', report_id or '')
    report_path = os.path.join(REPORTS_DIR, f"{report_id}.json")
    if not report_id or not os.path.isfile(report_path):
        flash('Звіт не знайдено або застарів', 'warning')
        return redirect(url_for('students.import_from_excel'))

    row_num = request.form.get('row', type=int)
    fields_to_apply = request.form.getlist('fields')

    if not row_num or not fields_to_apply:
        flash('Не обрано жодного поля для заміни', 'warning')
        return redirect(url_for('students.import_report', report_id=report_id))

    with open(report_path, encoding='utf-8') as f:
        report = json.load(f)

    target = None
    for w in report.get('warnings', []):
        if w.get('row') == row_num and w.get('action') == 'replace_fields':
            target = w
            break

    if not target:
        flash('Запис не знайдено', 'warning')
        return redirect(url_for('students.import_report', report_id=report_id))

    sid = target.get('student_id')
    if not sid:
        flash('Некоректний запис (немає student_id)', 'error')
        return redirect(url_for('students.import_report', report_id=report_id))

    conflicts_by_field = {c['field']: c for c in (target.get('conflicts') or [])}
    to_apply = {
        f: conflicts_by_field[f]['new']
        for f in fields_to_apply
        if f in conflicts_by_field
    }

    if not to_apply:
        flash('Жодного поля для заміни не знайдено', 'warning')
        return redirect(url_for('students.import_report', report_id=report_id))

    conn = get_db()
    try:
        student = conn.execute(
            "SELECT last_name_UA, first_name_UA FROM students WHERE id = ?",
            (sid,),
        ).fetchone()
        if not student:
            flash(f'Студента ID {sid} не знайдено (можливо, видалений)', 'error')
            return redirect(url_for('students.import_report', report_id=report_id))

        # ── Розділяємо edebo_code (per-enrollment) від інших полів ──
        # Витягуємо edebo_code з to_apply, щоб він не пішов у
        # UPDATE students (students.edebo_code — лише КЕШ primary
        # enrollment'а, тому записуємо через update_enrollment_edebo_code).
        edc_conflict = next(
            (c for c in (target.get('conflicts') or [])
             if c['field'] == 'edebo_code'),
            None,
        )
        edebo_value = to_apply.pop('edebo_code', None)

        # Якщо міняємо укр. ФІО — перерахувати англійські варіанти
        new_last = to_apply.get('last_name_UA') or student['last_name_UA'] or ''
        new_first = to_apply.get('first_name_UA') or student['first_name_UA'] or ''
        if 'last_name_UA' in to_apply or 'first_name_UA' in to_apply:
            last_eng, first_eng = generate_english_name(new_last, new_first)
            to_apply['last_name_ENG'] = last_eng
            to_apply['first_name_ENG'] = first_eng

        # 1) UPDATE students для останніх полів (без edebo)
        if to_apply:
            set_clause = ', '.join(f"{k} = ?" for k in to_apply)
            conn.execute(
                f"UPDATE students SET {set_clause} WHERE id = ?",
                list(to_apply.values()) + [sid],
            )

        # 2) edebo_code — через enrollment (per-enrollment)
        if edebo_value and edc_conflict:
            edc_gid = edc_conflict.get('group_id')
            if edc_gid:
                from routes.enrollments_service import (
                    update_enrollment_edebo_code,
                )
                update_enrollment_edebo_code(
                    sid, edc_gid, edebo_value, conn=conn,
                )

        conn.commit()

        # ── Список фактично застосованих полів (edebo потрапив у
        #    to_apply.pop(...) вище — його треба врахувати окремо) ──
        applied_field_set = set(to_apply.keys())
        if edebo_value and edc_conflict:
            applied_field_set.add('edebo_code')

        # Позначаємо, що саме застосували
        applied = target.setdefault('applied_fields', [])
        applied.extend(sorted(applied_field_set))

        # Прибираємо застосовані конфлікти (по applied_field_set,
        # а НЕ по to_apply — інакше edebo_code залишиться в remaining
        # назавжди, і warning ніколи не стане resolved)
        remaining = [
            c for c in (target.get('conflicts') or [])
            if c['field'] not in applied_field_set
        ]
        target['conflicts'] = remaining

        if remaining:
            target['reason'] = (
                f"Студент ID {sid} — залишилось {len(remaining)} розбіжностей "
                f"(вирішено {len(applied_field_set)})"
            )
        else:
            target['resolved'] = True
            target['resolved_at'] = datetime.now().strftime('%d.%m.%Y %H:%M')
            target['resolved_by'] = current_username()

        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False)

        log_action(
            current_username(),
            f'замінив поля студента ID {sid} за звітом імпорту',
            details=f'поля: {", ".join(sorted(applied_field_set))}',
        )
        flash(
            f'Оновлено поля: {", ".join(sorted(applied_field_set))}',
            'success',
        )

    except Exception as e:
        conn.rollback()
        logger.error(
            f'import_report_apply_fields failed '
            f'(report={report_id}, student={sid}): {e}',
            exc_info=True,
        )
        flash(f'Помилка оновлення: {e}', 'error')
    finally:
        conn.close()

    return redirect(url_for('students.import_report', report_id=report_id))
    
@students_bp.route(
    '/import_from_excel/report/<report_id>/apply-all-fields',
    methods=['POST'],
)
@permission_required('import_from_excel')
def import_report_apply_all_fields(report_id):
    """Замінити ВСІ обрані галочками поля у ВСІХ конфліктах одним POST.

    Приймає payload — JSON виду {row_num: [field1, field2, ...]}.
    Обробляє тільки warnings з action='replace_fields'.
    Для кожного студента — один UPDATE з усіма його полями.

    Ключове: на відміну від одиночного apply-fields, обробляє
    декілька студентів за один запит — зручно, коли у звіті 10+
    рядків з розбіжностями.
    """
    report_id = re.sub(r'[^0-9a-f]', '', report_id or '')
    report_path = os.path.join(REPORTS_DIR, f"{report_id}.json")
    if not report_id or not os.path.isfile(report_path):
        flash('Звіт не знайдено або застарів', 'warning')
        return redirect(url_for('students.import_from_excel'))

    payload_raw = (request.form.get('payload') or '').strip()
    if not payload_raw:
        flash('Не обрано жодного поля для заміни', 'warning')
        return redirect(url_for('students.import_report', report_id=report_id))

    try:
        payload = json.loads(payload_raw)
    except (ValueError, TypeError):
        flash('Некоректний формат даних', 'error')
        return redirect(url_for('students.import_report', report_id=report_id))

    if not isinstance(payload, dict) or not payload:
        flash('Некоректний формат даних', 'error')
        return redirect(url_for('students.import_report', report_id=report_id))

    with open(report_path, encoding='utf-8') as f:
        report = json.load(f)

    warnings_list = report.get('warnings') or []
    warning_by_row = {
        w.get('row'): w
        for w in warnings_list
        if w.get('action') == 'replace_fields' and not w.get('resolved')
    }

    conn = get_db()
    applied_rows = []
    failed_rows = []
    total_fields = 0

    try:
        for row_key, fields in payload.items():
            try:
                row_num = int(row_key)
            except (ValueError, TypeError):
                continue

            warning = warning_by_row.get(row_num)
            if not warning:
                continue

            if not isinstance(fields, list) or not fields:
                continue

            sid = warning.get('student_id')
            if not sid:
                failed_rows.append((row_num, 'немає student_id'))
                continue

            # Фільтруємо тільки ті поля, що реально є в conflicts
            # (захист від підміни форми через DevTools)
            conflicts_by_field = {
                c['field']: c for c in (warning.get('conflicts') or [])
            }
            to_apply = {
                f: conflicts_by_field[f]['new']
                for f in fields
                if f in conflicts_by_field
            }
            if not to_apply:
                continue

            student = conn.execute(
                "SELECT last_name_UA, first_name_UA FROM students WHERE id = ?",
                (sid,),
            ).fetchone()
            if not student:
                failed_rows.append((row_num, f'студент ID {sid} не знайдено'))
                continue

            # ── Розділяємо edebo_code (per-enrollment) від інших полів ──
            # Витягуємо edebo_code з to_apply, щоб він не пішов у
            # UPDATE students (students.edebo_code — лише КЕШ primary
            # enrollment'а, тому записуємо через update_enrollment_edebo_code).
            edc_conflict = next(
                (c for c in (warning.get('conflicts') or [])
                 if c['field'] == 'edebo_code'),
                None,
            )
            edebo_value = to_apply.pop('edebo_code', None)

            # ФІО: якщо міняємо укр. прізвище/ім'я — перегенерувати ENG
            new_last = (
                to_apply.get('last_name_UA')
                or student['last_name_UA'] or ''
            )
            new_first = (
                to_apply.get('first_name_UA')
                or student['first_name_UA'] or ''
            )
            if 'last_name_UA' in to_apply or 'first_name_UA' in to_apply:
                last_eng, first_eng = generate_english_name(new_last, new_first)
                to_apply['last_name_ENG'] = last_eng
                to_apply['first_name_ENG'] = first_eng

            # 1) UPDATE students для останніх полів (без edebo)
            if to_apply:
                set_clause = ', '.join(f"{k} = ?" for k in to_apply)
                conn.execute(
                    f"UPDATE students SET {set_clause} WHERE id = ?",
                    list(to_apply.values()) + [sid],
                )

            # 2) edebo_code — через enrollment (per-enrollment)
            if edebo_value and edc_conflict:
                edc_gid = edc_conflict.get('group_id')
                if edc_gid:
                    from routes.enrollments_service import (
                        update_enrollment_edebo_code,
                    )
                    update_enrollment_edebo_code(
                        sid, edc_gid, edebo_value, conn=conn,
                    )

            # ── Список фактично застосованих полів ──
            applied_field_set = set(to_apply.keys())
            if edebo_value and edc_conflict:
                applied_field_set.add('edebo_code')

            # Оновлюємо warning у звіті
            applied = warning.setdefault('applied_fields', [])
            applied.extend(sorted(applied_field_set))

            remaining = [
                c for c in (warning.get('conflicts') or [])
                if c['field'] not in applied_field_set
            ]
            warning['conflicts'] = remaining
            if remaining:
                warning['reason'] = (
                    f"Студент ID {sid} — залишилось {len(remaining)} "
                    f"розбіжностей (вирішено {len(applied_field_set)})"
                )
            else:
                warning['resolved'] = True
                warning['resolved_at'] = datetime.now().strftime('%d.%m.%Y %H:%M')
                warning['resolved_by'] = current_username()

            applied_rows.append((row_num, sid, sorted(applied_field_set)))
            total_fields += len(applied_field_set)

        conn.commit()

        # Перезаписуємо JSON-звіт
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False)

        if applied_rows:
            log_action(
                current_username(),
                f'масово замінив поля за звітом імпорту: '
                f'{len(applied_rows)} студент(ів), {total_fields} полів',
                details='; '.join(
                    f'рядок {r} → ID {sid} ({", ".join(flds)})'
                    for r, sid, flds in applied_rows
                ),
            )
            flash(
                f'Оновлено {len(applied_rows)} студент(ів), '
                f'замінено {total_fields} полів',
                'success',
            )
        else:
            flash('Жодного поля не оновлено', 'warning')

        if failed_rows:
            flash(
                'Не вдалося обробити: ' + '; '.join(
                    f'рядок {r} ({msg})' for r, msg in failed_rows
                ),
                'warning',
            )

    except Exception as e:
        conn.rollback()
        logger.error(
            f'import_report_apply_all_fields failed '
            f'(report={report_id}): {e}',
            exc_info=True,
        )
        flash(f'Помилка: {e}', 'error')
    finally:
        conn.close()

    return redirect(url_for('students.import_report', report_id=report_id))