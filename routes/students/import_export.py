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
import sqlite3
from routes.utils import generate_english_name
from werkzeug.utils import secure_filename
from datetime import datetime
import openpyxl
import os


@students_bp.route('/import_from_excel', methods=['GET', 'POST'])
@permission_required('import_from_excel')
def import_from_excel():
    """Імпорт студентів з Excel-файлу."""
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
        skipped = 0

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
                        skipped += 1
                        continue

                    group_id = row[0]
                    full_name = row[1]
                    birth_date_raw = row[2]
                    edebo_code = row[3] if len(row) > 3 and row[3] else ''

                    # Ліцензія вступу (колонка E, необов'язкова) - за
                    # короткою назвою з каталогу ("Київська", "Львівська"
                    # тощо, без регістру). Не знайдено - лишаємо
                    # порожнім і попереджаємо, рядок все одно імпортується.
                    license_id = None
                    if len(row) > 4 and row[4] not in (None, ''):
                        license_key = str(row[4]).strip().lower()
                        license_id = license_lookup.get(license_key)
                        if license_id is None:
                            flash(f"⚠️ Рядок {i}: ліцензію '{row[4]}' не знайдено в каталозі - поле пропущено, студент імпортований без неї")

                    # Кредити скороченої програми (колонка F, необов'язкова) -
                    # вступ з визнанням частини кредитів попереднього
                    # диплома. Порожнє/некоректне значення = звичайний
                    # студент за програмою групи, рядок все одно імпортується.
                    program_credits_override = None
                    if len(row) > 5 and row[5] not in (None, ''):
                        try:
                            program_credits_override = int(row[5])
                        except (ValueError, TypeError):
                            flash(f"⚠️ Рядок {i}: некоректне значення кредитів скороченої програми '{row[5]}' - поле пропущено, студент імпортований без нього")

                    raw_military = list(row[6:15]) if len(row) > 6 else []
                    military_data = raw_military + [None] * max(0, 9 - len(raw_military))

                    # Телефон (колонка P, необов'язкова) - основний і
                    # резервний номер через кому, напр.
                    # "+380991234567, +380991234568". Другий номер
                    # необов'язковий; якщо коми немає - записується
                    # лише основний.
                    phone = None
                    phone_backup = None
                    if len(row) > 15 and row[15] not in (None, ''):
                        phone_parts = [p.strip() for p in str(row[15]).split(',') if p.strip()]
                        if phone_parts:
                            phone = phone_parts[0]
                        if len(phone_parts) > 1:
                            phone_backup = phone_parts[1]

                    # Email (колонка Q, необов'язкова).
                    email = str(row[16]).strip() if len(row) > 16 and row[16] not in (None, '') else None

                    if not full_name:
                        continue

                    try:
                        group_id = int(group_id)
                    except (ValueError, TypeError):
                        flash(f"❗ Рядок {i}: некоректний ID групи '{row[0]}'")
                        skipped += 1
                        continue

                    if group_id not in allowed_group_ids:
                        flash(f"❗ Рядок {i}: група {group_id} не існує або недоступна")
                        skipped += 1
                        continue

                    name_parts = full_name.strip().split()
                    if len(name_parts) != 3:
                        flash(f"❗ Рядок {i}: невірний формат ПІБ '{full_name}'")
                        skipped += 1
                        continue
                    last_name, first_name, middle_name = name_parts

                    if isinstance(birth_date_raw, datetime):
                        birth_date = birth_date_raw.strftime("%d.%m.%Y")
                    else:
                        birth_date = str(birth_date_raw).strip()

                    existing = conn.execute("""
                        SELECT id FROM students
                        WHERE last_name_UA=? AND first_name_UA=? AND middle_name_UA=? AND birth_date=?
                    """, (last_name, first_name, middle_name, birth_date)).fetchone()
                    if existing:
                        skipped += 1
                        continue

                    last_name_eng, first_name_eng = generate_english_name(last_name, first_name)

                    conn.execute("""
                        INSERT INTO students (
                            last_name_UA, first_name_UA, middle_name_UA,
                            last_name_ENG, first_name_ENG, birth_date, group_id, edebo_code,
                            license_id, program_credits_override, phone, phone_backup, email
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (last_name, first_name, middle_name, last_name_eng, first_name_eng,
                          birth_date, group_id, edebo_code, license_id, program_credits_override,
                          phone, phone_backup, email))
                    student_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

                    if any(military_data):
                        issued_VOD_raw = military_data[2]
                        if isinstance(issued_VOD_raw, datetime):
                            issued_VOD = issued_VOD_raw.strftime("%d.%m.%Y")
                        elif isinstance(issued_VOD_raw, str):
                            issued_VOD = issued_VOD_raw.strip().replace('-', '.')
                            try:
                                datetime.strptime(issued_VOD, "%d.%m.%Y")
                            except ValueError:
                                issued_VOD = ''
                        else:
                            issued_VOD = ''

                        conn.execute("""
                            INSERT INTO military (
                                student_id, registration_number_of_the_DRPVR,
                                military_registration_document, issued_VOD,
                                military_accounting_specialty_number, military_rank,
                                change_credentials, reason_for_changing_credentials,
                                being_on_military_registration, address_of_residence
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (student_id, military_data[0], military_data[1], issued_VOD,
                              military_data[3], military_data[4], military_data[5],
                              military_data[6], military_data[7], military_data[8]))

                    inserted += 1

                except Exception as e:
                    logger.debug(f"Пропущено рядок {i} при імпорті студентів з Excel: {e}")
                    flash(f"⚠️ Помилка в рядку {i}: {e}")
                    skipped += 1
                    continue

            conn.commit()

        except Exception as e:
            conn.rollback()
            logger.error(f"Помилка при імпорті студентів з Excel (файл: {filename}): {e}", exc_info=True)
            flash(f"⚠️ Помилка при читанні файлу: {e}")
        finally:
            conn.close()

        log_action(
            current_username(),
            f"імпорт студентів з Excel: додано {inserted}, пропущено {skipped}",
            details=f"файл: {filename}"
        )
        flash(f"✅ Імпорт завершено. Додано: {inserted}, пропущено: {skipped}")
        return redirect(url_for('students.student_list'))

    return render_template('import_excel.html')
