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
                    edebo_code = (row[3]
                                  if len(row) > 3 and row[3] else '')

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
                    phone = None
                    phone_backup = None
                    if len(row) > 6 and row[6] not in (None, ''):
                        parts = [p.strip()
                                 for p in str(row[6]).split(',') if p.strip()]
                        if parts:
                            phone = parts[0]
                        if len(parts) > 1:
                            phone_backup = parts[1]

                    # ── H: Email ──
                    email = (str(row[7]).strip()
                             if len(row) > 7 and row[7] not in (None, '')
                             else None)

                    # ── I: ІПН ──
                    tax_id = (str(row[8]).strip()
                              if len(row) > 8 and row[8] not in (None, '')
                              else None)

                    # ── J: Адреса ──
                    address_of_residence = (
                        str(row[9]).strip()
                        if len(row) > 9 and row[9] not in (None, '')
                        else None)

                    # ── K: Стать ──
                    gender = (_normalize_gender(row[10])
                              if len(row) > 10 else None)

                    if not full_name:
                        continue

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

                    name_parts = str(full_name).strip().split()
                    if len(name_parts) != 3:
                        skipped_rows.append({
                            'row': i,
                            'fio': str(full_name).strip(),
                            'birth_date': str(birth_date_raw or '').strip(),
                            'reason': (f"Невірний формат ПІБ: «{full_name}» "
                                       f"(очікується 3 слова через пробіл)"),
                        })
                        continue
                    last_name, first_name, middle_name = name_parts

                    if gender is None:
                        gender = detect_gender_from_middle_name(middle_name)
                        if gender is None:
                            gender_missing += 1
                        else:
                            gender_auto += 1

                    if isinstance(birth_date_raw, datetime):
                        birth_date = birth_date_raw.strftime("%d.%m.%Y")
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
                        # tax_id → точний збіг
                        if tax_id:
                            existing = conn.execute(
                                SELECT_EXISTING + " WHERE tax_id = ?",
                                (tax_id,),
                            ).fetchone()
                        # fallback: ПІБ + дата
                        if not existing:
                            existing = conn.execute(
                                SELECT_EXISTING + """
                                WHERE last_name_UA=? AND first_name_UA=?
                                  AND middle_name_UA=? AND birth_date=?
                                """,
                                (last_name, first_name, middle_name,
                                 birth_date),
                            ).fetchone()

                    # ══════════════════════════════════════════════════════
                    # ── СЦЕНАРІЙ 1: новий студент ──
                    # ══════════════════════════════════════════════════════
                    if not existing:
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
                            edebo_code=edebo_code, 
                            conn=conn,
                        )
                        inserted += 1
                        continue

                    # ══════════════════════════════════════════════════════
                    # ── СЦЕНАРІЙ 2: людина вже в базі ──
                    # ══════════════════════════════════════════════════════
                    sid = existing['id']
                    was_archived = bool(existing['archived'])

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

                    if was_reactivation:
                        enrollments_reactivated += 1
                    else:
                        enrollments_added += 1

                    # ── Оновлення порожніх полів картки ──
                    if update_existing_data:
                        updates = {}
                        # ВАЖЛИВО: оновлюємо ТІЛЬКИ порожні поля — не
                        # затираємо те, що адмін вже вписав вручну.
                        if phone and not (existing['phone'] or '').strip():
                            updates['phone'] = phone
                        if phone_backup and not (existing['phone_backup'] or '').strip():
                            updates['phone_backup'] = phone_backup
                        if email and not (existing['email'] or '').strip():
                            updates['email'] = email
                        if tax_id and not (existing['tax_id'] or '').strip():
                            updates['tax_id'] = tax_id
                        if edebo_code and not (existing['edebo_code'] or '').strip():
                            updates['edebo_code'] = edebo_code
                        if address_of_residence and not (existing['address_of_residence'] or '').strip():
                            updates['address_of_residence'] = address_of_residence
                        if (program_credits_override
                                and not existing['program_credits_override']):
                            updates['program_credits_override'] = program_credits_override
                        if gender and not (existing['gender'] or '').strip():
                            updates['gender'] = gender

                        if updates:
                            set_clause = ', '.join(f"{k} = ?"
                                                   for k in updates)
                            conn.execute(
                                f"UPDATE students SET {set_clause} "
                                f"WHERE id = ?",
                                list(updates.values()) + [sid],
                            )
                            data_updated += 1

                    # ── Попередження про реактивацію з архіву ──
                    if was_archived:
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
    return render_template('import_excel_report.html', report=report)