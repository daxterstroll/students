# -*- coding: utf-8 -*-
"""
Публічна сторінка "оновити мої дані" за одноразовим токен-посиланням
(/update-info/<token>) - НЕ вимагає авторизації. Той самий принцип
ізоляції, що й routes/public_apply.py: жодного імпорту з routes/admin.py
чи routes/students.py, лише читання/запис у власну таблицю
update_requests (+ спільна attachments для сканів).

На відміну від /apply (повністю анонімна форма, створює лише
"чернетку" в карантині pending_students), тут посилання прив'язане до
КОНКРЕТНОГО, вже наявного студента - тому головна межа безпеки не
honeypot, а токен (випадковий, непередбачуваний) + суворий allowlist
полів на сервері: навіть якщо хтось руками допише в запит зайві поля,
буде прийнято лише те, що адмін дозволив явно (update_requests.allowed_fields).

Токен дійсний 1 добу АБО до першого успішного заповнення - що настане
раніше (перевіряється при кожному зверненні, окремого фонового
завдання для "протухання" не потрібно).

Заповнені дані НЕ застосовуються одразу - лягають на модерацію
(status='submitted'), адмін підтверджує вручну на окремій сторінці
(routes/admin.py: update_requests_review) так само вибірково, поле за
полем, як і з дублікатами заявок на реєстрацію.
"""
import os
import json
from datetime import datetime

from flask import Blueprint, render_template, request
from routes.db import get_db
from routes.utils import save_multiple_attachments
from routes.photo import process_and_save_pending_photo

public_update_bp = Blueprint('public_update', __name__)

ALLOWED_SCAN_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.webp', '.pdf'}
MAX_SCAN_SIZE_BYTES = 15 * 1024 * 1024
MAX_SCAN_FILES = 5

# Прості поля студента, які можна дозволити редагувати - назва форми:підпис.
SIMPLE_FIELDS = {
    'last_name_UA': "Прізвище (українською)",
    'first_name_UA': "Ім'я (українською)",
    'middle_name_UA': "По батькові (українською)",
    'phone': "Телефон",
    'phone_backup': "Резервний телефон",
    'email': "Email",
    'tax_id': "Ідентифікаційний код",
    'edebo_code': "Код ЄДЕБО",
}
# Цілі розділи (з власними підполями і сканами) - той самий набір, що
# й у публічній анкеті /apply.
SECTION_LABELS = {
    'passport': "Паспортні дані",
    'education_document': "Документ про освіту",
    'military': "Військові дані",
}


def _validate_scan_files(files, max_files, label):
    if len(files) > max_files:
        return f"Забагато файлів{label} (максимум {max_files})."
    for sf in files:
        ext = os.path.splitext(sf.filename)[1].lower()
        if ext not in ALLOWED_SCAN_EXTENSIONS:
            return f"Файл «{sf.filename}»{label}: непідтримуваний формат. Дозволені: JPG, PNG, WEBP, PDF."
        sf.seek(0, os.SEEK_END)
        size = sf.tell()
        sf.seek(0)
        if size > MAX_SCAN_SIZE_BYTES:
            size_mb = round(size / 1024 / 1024, 1)
            return f"Файл «{sf.filename}»{label} завеликий ({size_mb} МБ, максимум {MAX_SCAN_SIZE_BYTES // 1024 // 1024} МБ)."
    return None


def _iso_to_ua_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return value


def _load_request(token):
    conn = get_db()
    row = conn.execute("SELECT * FROM update_requests WHERE token = ?", (token,)).fetchone()
    return conn, row


def _is_expired(row):
    try:
        expires = datetime.strptime(row['expires_at'], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        expires = datetime.strptime(row['expires_at'], "%Y-%m-%d %H:%M:%S.%f")
    return datetime.now() > expires


@public_update_bp.route('/update-info/<token>', methods=['GET', 'POST'])
def update_info(token):
    conn, row = _load_request(token)
    if not row:
        conn.close()
        return render_template('public_update_invalid.html', reason="not_found"), 404

    if row['status'] != 'pending':
        conn.close()
        return render_template('public_update_invalid.html', reason="used")

    if _is_expired(row):
        conn.close()
        return render_template('public_update_invalid.html', reason="expired")

    student = conn.execute(
        "SELECT last_name_UA, first_name_UA, middle_name_UA, phone, phone_backup, email, tax_id, edebo_code, photo "
        "FROM students WHERE id = ?", (row['student_id'],)
    ).fetchone()
    if not student:
        conn.close()
        return render_template('public_update_invalid.html', reason="not_found"), 404

    allowed_fields = json.loads(row['allowed_fields'])

    if request.method == 'GET':
        conn.close()
        return render_template(
            'public_update_form.html',
            student=student, allowed_fields=allowed_fields,
            simple_fields=SIMPLE_FIELDS, section_labels=SECTION_LABELS,
        )

    # ---- POST ----
    if (request.form.get('website') or '').strip():
        # Honeypot - мовчки "приймаємо", нічого не зберігаючи.
        conn.close()
        return render_template('public_update_thanks.html')

    def f(name):
        return (request.form.get(name) or '').strip() or None

    submitted = {}

    # Прості поля - СУВОРО лише ті, що є в allowed_fields (сервер, не
    # клієнт, вирішує, що приймати - навіть якщо хтось допише зайве
    # поле в POST-запит руками, воно буде проігнороване).
    for key in SIMPLE_FIELDS:
        if key in allowed_fields:
            submitted[key] = f(key)

    photo_path = None
    if 'photo' in allowed_fields:
        photo_file = request.files.get('photo')
        if photo_file and photo_file.filename:
            try:
                photo_path = process_and_save_pending_photo(photo_file.read())
            except ValueError as e:
                conn.close()
                return render_template(
                    'public_update_form.html', student=student, allowed_fields=allowed_fields,
                    simple_fields=SIMPLE_FIELDS, section_labels=SECTION_LABELS,
                    error=f"Фото: {e}", form=request.form,
                )

    if 'passport' in allowed_fields:
        scans = [sf for sf in request.files.getlist('passport_scans') if sf and sf.filename]
        err = _validate_scan_files(scans, MAX_SCAN_FILES, " паспорта")
        if err:
            conn.close()
            return render_template(
                'public_update_form.html', student=student, allowed_fields=allowed_fields,
                simple_fields=SIMPLE_FIELDS, section_labels=SECTION_LABELS,
                error=err, form=request.form,
            )
        if any([f('passport_number'), f('passport_document_type')]):
            submitted['passport'] = {
                'document_type': f('passport_document_type') if f('passport_document_type') in ('Паспорт (книжка)', 'ID-картка') else None,
                'series': f('passport_series'),
                'number': f('passport_number'),
                'issued_by': f('passport_issued_by'),
                'issue_date': _iso_to_ua_date(f('passport_issue_date')),
                'valid_until': _iso_to_ua_date(f('passport_valid_until')),
                'unique_number': f('passport_unique_number'),
            }

    if 'education_document' in allowed_fields:
        scans = [sf for sf in request.files.getlist('document_scans') if sf and sf.filename]
        err = _validate_scan_files(scans, MAX_SCAN_FILES, " документа про освіту")
        if err:
            conn.close()
            return render_template(
                'public_update_form.html', student=student, allowed_fields=allowed_fields,
                simple_fields=SIMPLE_FIELDS, section_labels=SECTION_LABELS,
                error=err, form=request.form,
            )
        if any([f('document_type'), f('document_number')]):
            submitted['education_document'] = {
                'document_type': f('document_type'),
                'series': f('document_series'),
                'number': f('document_number'),
                'institution': f('document_institution'),
                'country': f('document_country'),
                'completion_date': _iso_to_ua_date(f('document_date')),
            }

    if 'military' in allowed_fields:
        scans = [sf for sf in request.files.getlist('military_scans') if sf and sf.filename]
        err = _validate_scan_files(scans, MAX_SCAN_FILES, " військового обліку")
        if err:
            conn.close()
            return render_template(
                'public_update_form.html', student=student, allowed_fields=allowed_fields,
                simple_fields=SIMPLE_FIELDS, section_labels=SECTION_LABELS,
                error=err, form=request.form,
            )
        submitted['military'] = {
            'registration_number_of_the_DRPVR': f('military_registration_number_drpvr'),
            'military_registration_document': f('military_registration_document'),
            'issued_VOD': _iso_to_ua_date(f('military_issued_vod')),
            'military_accounting_specialty_number': f('military_specialty_number'),
            'military_rank': f('military_rank'),
            'address_of_residence': f('military_address'),
        }

    conn.execute(
        "UPDATE update_requests SET status='submitted', submitted_data=?, photo_path=?, submitted_at=datetime('now','localtime') WHERE id=?",
        (json.dumps(submitted, ensure_ascii=False), photo_path, row['id'])
    )
    conn.commit()

    # Скани - ті самі файли, вже провалідовані вище, повторно читаємо
    # списки (стрім у request.files не вичерпується validate-функцією,
    # яка лише .seek()-ає, тому це безпечно).
    if 'passport' in allowed_fields:
        scans = [sf for sf in request.files.getlist('passport_scans') if sf and sf.filename]
        if scans:
            save_multiple_attachments(conn, 'update_request_passport', row['id'], scans, 'update_requests')
    if 'education_document' in allowed_fields:
        scans = [sf for sf in request.files.getlist('document_scans') if sf and sf.filename]
        if scans:
            save_multiple_attachments(conn, 'update_request_education', row['id'], scans, 'update_requests')
    if 'military' in allowed_fields:
        scans = [sf for sf in request.files.getlist('military_scans') if sf and sf.filename]
        if scans:
            save_multiple_attachments(conn, 'update_request_military', row['id'], scans, 'update_requests')
    conn.commit()
    conn.close()

    return render_template('public_update_thanks.html')
