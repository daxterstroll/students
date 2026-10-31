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

    Колонки (A→K, 0-базовий індекс):
        0  A  group_id                  (обов.)
        1  B  ПІБ повністю (3 слова)    (обов.)
        2  C  birth_date                (обов.)
        3  D  edebo_code
        4  E  license (short_name_ua)
        5  F  program_credits_override
        6  G  phone (основний[, резервний])
        7  H  email
        8  I  tax_id                    (РНОКПП)
        9  J  address_of_residence
        10 K  gender                    (M/F/Ч/Ж; порожнє → авто за по батькові)

    last_name_ENG / first_name_ENG НЕ читаються з файлу - генеруються
    автоматично через generate_english_name() (транслітерація КМУ №55-2010).
    id / archived / photo теж не імпортуються: id автоінкрементний,
    archived = FALSE, photo = NULL за замовчуванням.

    Військові дані з Excel більше НЕ імпортуються - вони заповнюються
    вручну на картці студента або через PDF «Резерв+» (сторінка
    «Військові дані» або пакетний імпорт PDF).
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
        inserted = 0
        skipped_rows = []  # список dict: {'row': i, 'fio': ..., 'birth_date': ..., 'reason': ...}
        gender_auto = 0
        gender_missing = 0

        role = session.get('role')
        user_group_ids = session.get('group_ids', [])

        if role == 'admin':
            allowed_group_ids = {
                row['id'] for row in conn.execute("SELECT id FROM groups WHERE archived = FALSE").fetchall()
            }
        else:
            if not user_group_ids:
                allowed_group_ids = set()
            else:
                placeholders = ','.join('?' * len(user_group_ids))
                allowed_group_ids = {
                    row['id'] for row in conn.execute(
                        f"SELECT id FROM groups WHERE id IN ({placeholders}) AND archived = FALSE",
                        user_group_ids
                    ).fetchall()
                }

        # Каталог ліцензій - для зіставлення тексту з колонки E з
        # institution_licenses (за short_name_ua, без регістру).
        license_lookup = {}
        for row in conn.execute("SELECT id, short_name_ua FROM institution_licenses").fetchall():
            if row['short_name_ua']:
                license_lookup[row['short_name_ua'].strip().lower()] = row['id']

        try:
            wb = openpyxl.load_workbook(filepath)
            sheet = wb.active

            for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                try:
                    if not row or len(row) < 4:
                        skipped_rows.append({
                            'row': i,
                            'fio': '—',
                            'birth_date': '—',
                            'reason': 'Порожній рядок або менше 4 колонок'
                        })
                        continue

                    group_id = row[0]
                    full_name = row[1]
                    birth_date_raw = row[2]
                    edebo_code = row[3] if len(row) > 3 and row[3] else ''

                    # Ліцензія вступу (колонка E, необов'язкова)
                    license_id = None
                    if len(row) > 4 and row[4] not in (None, ''):
                        license_key = str(row[4]).strip().lower()
                        license_id = license_lookup.get(license_key)
                        if license_id is None:
                            flash(f"⚠️ Рядок {i}: ліцензію '{row[4]}' не знайдено в каталозі - поле пропущено, студент імпортований без неї")

                    # Кредити скороченої програми (колонка F, необов'язкова)
                    program_credits_override = None
                    if len(row) > 5 and row[5] not in (None, ''):
                        try:
                            program_credits_override = int(row[5])
                        except (ValueError, TypeError):
                            flash(f"⚠️ Рядок {i}: некоректне значення кредитів скороченої програми '{row[5]}' - поле пропущено, студент імпортований без нього")

                    # Телефон (колонка G, необов'язкова) - основний і
                    # резервний номер через кому.
                    phone = None
                    phone_backup = None
                    if len(row) > 6 and row[6] not in (None, ''):
                        phone_parts = [p.strip() for p in str(row[6]).split(',') if p.strip()]
                        if phone_parts:
                            phone = phone_parts[0]
                        if len(phone_parts) > 1:
                            phone_backup = phone_parts[1]

                    # Email (колонка H, необов'язкова)
                    email = str(row[7]).strip() if len(row) > 7 and row[7] not in (None, '') else None

                    # Ідентифікаційний код / РНОКПП (колонка I, необов'язкова)
                    tax_id = str(row[8]).strip() if len(row) > 8 and row[8] not in (None, '') else None

                    # Адреса проживання (колонка J, необов'язкова)
                    address_of_residence = str(row[9]).strip() if len(row) > 9 and row[9] not in (None, '') else None

                    # Стать (колонка K, необов'язкова) — нормалізуємо або
                    # визначаємо за по батькові нижче
                    gender = _normalize_gender(row[10]) if len(row) > 10 else None

                    if not full_name:
                        continue

                    try:
                        group_id = int(group_id)
                    except (ValueError, TypeError):
                        skipped_rows.append({
                            'row': i,
                            'fio': full_name.strip() if full_name else '—',
                            'birth_date': str(birth_date_raw or '').strip(),
                            'reason': f"Некоректний ID групи: '{row[0]}'"
                        })
                        continue

                    if group_id not in allowed_group_ids:
                        skipped_rows.append({
                            'row': i,
                            'fio': full_name.strip() if full_name else '—',
                            'birth_date': str(birth_date_raw or '').strip(),
                            'reason': f"Група ID {group_id} не існує або недоступна"
                        })
                        continue

                    name_parts = full_name.strip().split()
                    if len(name_parts) != 3:
                        skipped_rows.append({
                            'row': i,
                            'fio': full_name.strip(),
                            'birth_date': str(birth_date_raw or '').strip(),
                            'reason': f"Невірний формат ПІБ: «{full_name}» (очікується рівно 3 слова через пробіл)"
                        })
                        continue
                    last_name, first_name, middle_name = name_parts

                    # Якщо стать не вказана в Excel — визначаємо за по батькові
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

                    existing = conn.execute("""
                        SELECT id FROM students
                        WHERE last_name_UA=? AND first_name_UA=? AND middle_name_UA=? AND birth_date=?
                    """, (last_name, first_name, middle_name, birth_date)).fetchone()
                    if existing:
                        skipped_rows.append({
                            'row': i,
                            'fio': f"{last_name} {first_name} {middle_name}",
                            'birth_date': birth_date,
                            'reason': f"Студент уже є в базі (ID {existing['id']}) — дублікат"
                        })
                        continue

                    # Автоматична генерація англійських ПІБ (КМУ №55-2010) -
                    # з Excel ці поля НЕ читаються.
                    last_name_eng, first_name_eng = generate_english_name(last_name, first_name)

                    conn.execute("""
                        INSERT INTO students (
                            last_name_UA, first_name_UA, middle_name_UA,
                            last_name_ENG, first_name_ENG, birth_date, group_id, edebo_code,
                            license_id, program_credits_override, phone, phone_backup, email,
                            tax_id, address_of_residence, gender
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        last_name, first_name, middle_name,
                        last_name_eng, first_name_eng, birth_date, group_id, edebo_code,
                        license_id, program_credits_override, phone, phone_backup, email,
                        tax_id, address_of_residence, gender,
                    ))

                    inserted += 1

                except Exception as e:
                    logger.debug(f"Пропущено рядок {i} при імпорті студентів з Excel: {e}")
                    skipped_rows.append({
                        'row': i,
                        'fio': full_name.strip() if full_name else '—',
                        'birth_date': str(birth_date_raw or '').strip(),
                        'reason': f"Помилка обробки: {e}"
                    })
                    continue

            wb.close()
            conn.commit()

        except Exception as e:
            conn.rollback()
            logger.error(f"Помилка при імпорті студентів з Excel (файл: {filename}): {e}", exc_info=True)
            flash(f"⚠️ Помилка при читанні файлу: {e}")
        finally:
            conn.close()

        log_action(
            current_username(),
            f"імпорт студентів з Excel: додано {inserted}, пропущено {len(skipped_rows)}",
            details=f"файл: {filename}, стать авто: {gender_auto}, без статі: {gender_missing}"
        )

        # Підсумкове повідомлення з деталями по статі
        summary = f"✅ Імпорт завершено. Додано: {inserted}, пропущено: {len(skipped_rows)}"
        if gender_auto:
            summary += f", стать визначено автоматично: {gender_auto}"
        if gender_missing:
            summary += f", без статі: {gender_missing} (заповніть вручну)"
        flash(summary)

        if gender_missing:
            flash(f"⚠️ {gender_missing} студентів без по батькові — стать не визначена. "
                  f"Знайдіть їх через фільтр у списку студентів (або в картці кожного) та заповніть вручну.",
                  "warning")

        # Якщо є пропущені — показуємо детальний звіт, а не одразу список
        if skipped_rows:
            import uuid as _uuid
            os.makedirs(REPORTS_DIR, exist_ok=True)
            report_id = _uuid.uuid4().hex
            report_path = os.path.join(REPORTS_DIR, f"{report_id}.json")
            with open(report_path, 'w', encoding='utf-8') as f:
                json.dump({
                    'filename': filename,
                    'inserted': inserted,
                    'skipped_rows': skipped_rows,
                    'gender_auto': gender_auto,
                    'gender_missing': gender_missing,
                    'when': datetime.now().strftime('%d.%m.%Y %H:%M'),
                }, f, ensure_ascii=False)
            return redirect(url_for('students.import_report', report_id=report_id))

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