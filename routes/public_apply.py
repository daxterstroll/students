# -*- coding: utf-8 -*-
"""
Публічна анкета самореєстрації абітурієнта/студента - НЕ вимагає
авторизації, доступна без сесії. Свідомо ІЗОЛЬОВАНА від решти
застосунку.
"""
import os
import re
import tempfile
from datetime import datetime, date

from flask import Blueprint, render_template, request, jsonify
from routes.db import get_db
from routes.utils import generate_english_name, save_multiple_attachments, logger, title_case_ukrainian
from routes.photo import process_and_save_pending_photo, load_and_validate_image
from routes.config import PROJECT_ROOT
from routes.passport_types import (
    PASSPORT_TYPES, PASSPORT_TYPES_ID_STYLE, passport_template_context,
)

public_apply_bp = Blueprint('public_apply', __name__)

ALLOWED_SCAN_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.webp', '.pdf'}
MAX_SCAN_SIZE_BYTES = 15 * 1024 * 1024
MAX_SCAN_FILES = 5
MAX_PDF_SIZE_BYTES = 15 * 1024 * 1024

MIN_APPLICANT_AGE_YEARS = 14
MAX_APPLICANT_AGE_YEARS = 100
PHONE_RE = re.compile(r'^(?:\+?380|0)\d{9}$')

PENDING_PDFS_DIR = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'pending_pdfs')
PENDING_STUDENTS_DIR = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'pending_students')

@public_apply_bp.context_processor
def _inject_passport_types():
    """Додає passport_types/... у ВСІ render_template цього блюпринта."""
    return passport_template_context()

def _validate_birth_date(ua_date_str):
    try:
        d = datetime.strptime(ua_date_str, "%d.%m.%Y").date()
    except (ValueError, TypeError):
        return "Некоректна дата народження."
    today = date.today()
    if d > today:
        return "Дата народження не може бути в майбутньому."
    age = (today - d).days / 365.25
    if age < MIN_APPLICANT_AGE_YEARS:
        return f"За вказаною датою народження вік менше {MIN_APPLICANT_AGE_YEARS} років - перевірте дату."
    if age > MAX_APPLICANT_AGE_YEARS:
        return "Перевірте, будь ласка, дату народження - вказаний вік виглядає нереалістично великим."
    return None


def _validate_phone(value, label):
    if not value:
        return None
    digits = re.sub(r'[\s\-()]', '', value)
    if not PHONE_RE.match(digits):
        return f"{label}: перевірте формат номера (напр. +380991234567 або 0991234567)."
    return None


def _birth_date_bounds():
    today = date.today()

    def _safe_date(year, month, day):
        try:
            return date(year, month, day)
        except ValueError:
            return date(year, month, day - 1)

    return {
        'min_birth_date': _safe_date(today.year - MAX_APPLICANT_AGE_YEARS, today.month, today.day).isoformat(),
        'max_birth_date': _safe_date(today.year - MIN_APPLICANT_AGE_YEARS, today.month, today.day).isoformat(),
    }


def _iso_to_ua_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return value


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


def _parse_pdf_ajax():
    """AJAX-парсинг PDF. Викликається з apply() при action=parse_pdf."""
    if (request.form.get('website') or '').strip():
        return jsonify({'ok': False, 'error': 'spam'}), 400

    file = request.files.get('pdf_file')
    if not file or not file.filename:
        return jsonify({'ok': False, 'error': 'Оберіть PDF-файл'}), 400

    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'ok': False, 'error': 'Підтримуються лише PDF-файли'}), 400

    pdf_bytes = file.read()
    if len(pdf_bytes) > MAX_PDF_SIZE_BYTES:
        return jsonify({
            'ok': False,
            'error': f'Файл завеликий (максимум {MAX_PDF_SIZE_BYTES // 1024 // 1024} МБ)',
        }), 400

    if not pdf_bytes.startswith(b'%PDF'):
        return jsonify({'ok': False, 'error': 'Файл не є PDF'}), 400

    try:
        from routes.parse_military_pdf import parse_reserve_plus_pdf, extract_photo_from_pdf
    except ImportError as e:
        return jsonify({'ok': False, 'error': f'Парсер PDF недоступний: {e}'}), 500

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name
        parsed = parse_reserve_plus_pdf(tmp_path)
        photo_bytes = extract_photo_from_pdf(tmp_path)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'ok': False, 'error': f'Не вдалося розпізнати PDF: {e}'}), 400
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    try:
        full_name = title_case_ukrainian((parsed.get('full_name') or '').strip())
        name_parts = full_name.split()
        last_name = name_parts[0] if len(name_parts) > 0 else ''
        first_name = name_parts[1] if len(name_parts) > 1 else ''
        middle_name = ' '.join(name_parts[2:]) if len(name_parts) > 2 else ''

        birth_date_iso = ''
        bd = (parsed.get('birth_date') or '').strip()
        if bd:
            try:
                birth_date_iso = datetime.strptime(bd, '%d.%m.%Y').strftime('%Y-%m-%d')
            except ValueError:
                birth_date_iso = ''

        phone_raw = (parsed.get('phone') or '').strip()
        phone = phone_raw
        if phone_raw:
            digits = re.sub(r'[\s\-()]', '', phone_raw)
            if len(digits) == 10 and digits.startswith('0'):
                phone = '+38' + digits
            elif len(digits) == 12 and digits.startswith('380'):
                phone = '+' + digits
            elif len(digits) == 13 and digits.startswith('+380'):
                phone = digits

        fields = {
            'last_name_UA': last_name,
            'first_name_UA': first_name,
            'middle_name_UA': middle_name,
            'birth_date': birth_date_iso,
            'tax_id': parsed.get('tax_id') or '',
            'phone': phone,
            'email': parsed.get('email') or '',
            'military_registration_number_drpvr': parsed.get('registry_number') or '',
            'military_specialty_number': parsed.get('vos_code') or '',
            'military_rank': parsed.get('military_rank') or '',
            'military_address': parsed.get('address') or '',
        }

        photo_url = None
        photo_path = None
        photo_data_url = None
        if photo_bytes:
            try:
                photo_path = process_and_save_pending_photo(photo_bytes)
                photo_url = f'/static/{photo_path}'
                import base64
                abs_path = os.path.join(PROJECT_ROOT, 'static', photo_path)
                with open(abs_path, 'rb') as fp:
                    raw = fp.read()
                photo_data_url = 'data:image/jpeg;base64,' + base64.b64encode(raw).decode('ascii')
            except Exception:
                import traceback
                traceback.print_exc()

        # Зберігаємо оригінальний PDF АБСОЛЮТНО (PROJECT_ROOT/static/...)
        import uuid as _uuid
        os.makedirs(PENDING_PDFS_DIR, exist_ok=True)
        pdf_safe_name = f"{_uuid.uuid4().hex}.pdf"
        pdf_abs_path = os.path.join(PENDING_PDFS_DIR, pdf_safe_name)
        with open(pdf_abs_path, 'wb') as f:
            f.write(pdf_bytes)
        pdf_rel_path = f"uploads/pending_pdfs/{pdf_safe_name}"

        logger.info(
            f"[parse_pdf_apply] saved: abs={pdf_abs_path}, rel={pdf_rel_path}, "
            f"exists={os.path.exists(pdf_abs_path)}"
        )

        return jsonify({
            'ok': True,
            'fields': fields,
            'photo_url': photo_url,
            'photo_path': photo_path,
            'photo_data_url': photo_data_url,
            'pdf_path': pdf_rel_path,
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'ok': False, 'error': f'Серверна помилка обробки PDF: {e}'}), 500


@public_apply_bp.before_request
def _limit_request_size():
    max_total = 80 * 1024 * 1024
    if request.content_length and request.content_length > max_total:
        return render_template(
            'public_apply.html',
            error="Загальний розмір прикріплених файлів завеликий. Спробуйте завантажити менше файлів або стисніть їх.",
            **_birth_date_bounds(),
        ), 413


@public_apply_bp.route('/apply', methods=['GET', 'POST'])
def apply():
    if request.method == 'GET':
        return render_template('public_apply.html', **_birth_date_bounds())

    if request.form.get('action') == 'parse_pdf':
        try:
            return _parse_pdf_ajax()
        except Exception as e:
            import traceback
            traceback.print_exc()
            return jsonify({'ok': False, 'error': f'Критична серверна помилка: {e}'}), 500

    # ── Звичайна обробка форми ──
    if (request.form.get('website') or '').strip():
        return render_template('public_apply_thanks.html')

    def f(name):
        return (request.form.get(name) or '').strip() or None

    last_name_ua = f('last_name_UA')
    first_name_ua = f('first_name_UA')
    birth_date = _iso_to_ua_date(f('birth_date'))
    phone = f('phone')

    if not last_name_ua or not first_name_ua or not birth_date or not phone:
        return render_template(
            'public_apply.html',
            error="Заповніть, будь ласка, прізвище, ім'я, дату народження і телефон - без них заявку не можна обробити.",
            form=request.form,
            **_birth_date_bounds(),
        )

    validation_error = _validate_birth_date(birth_date)
    if not validation_error:
        validation_error = _validate_phone(phone, "Телефон")
    if not validation_error:
        validation_error = _validate_phone(f('phone_backup'), "Резервний телефон")
    if validation_error:
        return render_template('public_apply.html', error=validation_error, form=request.form, **_birth_date_bounds())

    photo_file = request.files.get('photo')
    photo_bytes = None
    prefilled_photo_path = (request.form.get('prefilled_photo_path') or '').strip()

    if photo_file and photo_file.filename:
        photo_bytes = photo_file.read()
        try:
            load_and_validate_image(photo_bytes)
        except ValueError as e:
            return render_template('public_apply.html', error=f"Фото: {e}", form=request.form, **_birth_date_bounds())
        prefilled_photo_path = ''
    elif prefilled_photo_path:
        if not prefilled_photo_path.startswith('uploads/pending_photos/'):
            prefilled_photo_path = ''
        else:
            full = os.path.join(PROJECT_ROOT, 'static', prefilled_photo_path)
            if not os.path.exists(full):
                prefilled_photo_path = ''

    scan_files = [sf for sf in request.files.getlist('document_scans') if sf and sf.filename]
    scan_error = _validate_scan_files(scan_files, MAX_SCAN_FILES, " документа про освіту")
    if scan_error:
        return render_template('public_apply.html', error=scan_error, form=request.form, **_birth_date_bounds())

    passport_scan_files = [sf for sf in request.files.getlist('passport_scans') if sf and sf.filename]
    passport_scan_error = _validate_scan_files(passport_scan_files, MAX_SCAN_FILES, " паспорта")
    if passport_scan_error:
        return render_template('public_apply.html', error=passport_scan_error, form=request.form, **_birth_date_bounds())

    military_scan_files = [sf for sf in request.files.getlist('military_scans') if sf and sf.filename]
    military_scan_error = _validate_scan_files(military_scan_files, MAX_SCAN_FILES, " військового обліку")
    if military_scan_error:
        return render_template('public_apply.html', error=military_scan_error, form=request.form, **_birth_date_bounds())

    is_foreign_document = 0
    reduced_program_claim = 0
    military_being_registered = 0

    last_name_eng, first_name_eng = generate_english_name(last_name_ua, first_name_ua)

    document_date = _iso_to_ua_date(f('document_date'))
    military_issued_vod = _iso_to_ua_date(f('military_issued_vod'))

    passport_document_type = f('passport_document_type')
    if passport_document_type not in PASSPORT_TYPES:
        passport_document_type = None
    passport_issue_date = _iso_to_ua_date(f('passport_issue_date'))
    passport_valid_until = _iso_to_ua_date(f('passport_valid_until')) if passport_document_type in PASSPORT_TYPES_ID_STYLE else None

    conn = get_db()
    cur = conn.execute("""
        INSERT INTO pending_students (
            submitted_ip,
            last_name_UA, first_name_UA, middle_name_UA, last_name_ENG, first_name_ENG, birth_date, tax_id,
            phone, phone_backup, email,
            document_type, document_series, document_number, document_institution, document_country, document_date,
            is_foreign_document, foreign_reference_number, foreign_reference_institution,
            foreign_reference_country, foreign_reference_issue_date,
            recognition_certificate_number, recognition_issuer, recognition_date,
            reduced_program_claim, reduced_program_specialty, reduced_program_institution,
            passport_document_type, passport_series, passport_number, passport_issued_by,
            passport_issue_date, passport_valid_until, passport_unique_number,
            military_registration_number_drpvr, military_registration_document, military_issued_vod,
            military_specialty_number, military_rank, military_being_registered, military_address,
            military_change_credentials, military_change_reason
        ) VALUES (?, ?,?,?,?,?,?,?, ?,?,?, ?,?,?,?,?,?, ?,?,?,?,?,?,?,?, ?,?,?, ?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?)
    """, (
        request.remote_addr,
        last_name_ua, first_name_ua, f('middle_name_UA'), last_name_eng, first_name_eng, birth_date, f('tax_id'),
        phone, f('phone_backup'), f('email'),
        f('document_type'), f('document_series'), f('document_number'), f('document_institution'),
        f('document_country'), document_date,
        is_foreign_document, f('reference_number'), f('reference_institution'),
        f('reference_country'), _iso_to_ua_date(f('reference_issue_date')),
        f('recognition_certificate_number'), f('recognition_issuer'), _iso_to_ua_date(f('recognition_date')),
        reduced_program_claim, None, None,
        passport_document_type, f('passport_series'), f('passport_number'), f('passport_issued_by'),
        passport_issue_date, passport_valid_until, f('passport_unique_number'),
        f('military_registration_number_drpvr'), f('military_registration_document'), military_issued_vod,
        f('military_specialty_number'), f('military_rank'), military_being_registered, f('military_address'),
        f('military_change_credentials'), f('military_change_reason'),
    ))
    pending_id = cur.lastrowid

    # ═══════════════════════════════════════════════════════════
    # PDF «Резерв+» — обробляємо файл із поля pdf_file (єдиний шлях,
    # без дублювання через prefilled_pdf_path). Поле в HTML
    # прив'язане до форми через атрибут form="applyForm", тому при
    # сабміті файл приходить сюди.
    # ═══════════════════════════════════════════════════════════
    pdf_file = request.files.get('pdf_file')
    logger.info(
        f"[apply] pdf_file: filename={getattr(pdf_file, 'filename', None)!r}"
    )
    if pdf_file and pdf_file.filename and pdf_file.filename.lower().endswith('.pdf'):
        pdf_bytes = pdf_file.read()
        logger.info(f"[apply] pdf_file size: {len(pdf_bytes)} bytes, "
                    f"starts_with_pdf={pdf_bytes.startswith(b'%PDF')}")
        if pdf_bytes.startswith(b'%PDF'):
            import uuid as _uuid
            os.makedirs(PENDING_STUDENTS_DIR, exist_ok=True)
            safe_name = f"{_uuid.uuid4().hex}.pdf"
            dest_pdf = os.path.join(PENDING_STUDENTS_DIR, safe_name)
            with open(dest_pdf, 'wb') as fp:
                fp.write(pdf_bytes)
            rel_pdf = f"uploads/pending_students/{safe_name}"
            conn.execute(
                "INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_by) "
                "VALUES (?, ?, ?, ?, ?)",
                ('pending_student_military', pending_id, rel_pdf,
                 pdf_file.filename or 'Резерв+.pdf', None),
            )
            logger.info(
                f"[apply] PDF saved from form: {rel_pdf} ({len(pdf_bytes)} bytes), "
                f"abs={dest_pdf}"
            )

    if photo_bytes:
        try:
            photo_path = process_and_save_pending_photo(photo_bytes)
            conn.execute("UPDATE pending_students SET photo_path=? WHERE id=?", (photo_path, pending_id))
        except ValueError:
            pass
    elif prefilled_photo_path:
        conn.execute("UPDATE pending_students SET photo_path=? WHERE id=?", (prefilled_photo_path, pending_id))

    if scan_files:
        save_multiple_attachments(conn, 'pending_student', pending_id, scan_files, 'pending_students')
    if passport_scan_files:
        save_multiple_attachments(conn, 'pending_student_passport', pending_id, passport_scan_files, 'pending_students')
    if military_scan_files:
        save_multiple_attachments(conn, 'pending_student_military', pending_id, military_scan_files, 'pending_students')

    conn.commit()
    conn.close()

    return render_template('public_apply_thanks.html')