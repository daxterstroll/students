# -*- coding: utf-8 -*-
"""
Публічна сторінка "оновити мої дані" за одноразовим токен-посиланням.
"""
import os
import re
import json
import tempfile
from datetime import datetime

from flask import Blueprint, render_template, request, jsonify
from routes.db import get_db
from routes.utils import save_multiple_attachments, logger
from routes.photo import process_and_save_pending_photo
from routes.config import PROJECT_ROOT

public_update_bp = Blueprint('public_update', __name__)

ALLOWED_SCAN_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.webp', '.pdf'}
MAX_SCAN_SIZE_BYTES = 15 * 1024 * 1024
MAX_SCAN_FILES = 5
MAX_PDF_SIZE_BYTES = 15 * 1024 * 1024

PENDING_PDFS_DIR = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'pending_pdfs')
UPDATE_REQUESTS_DIR = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'update_requests')

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
            return f"Файл «{sf.filename}»{label}: непідтримуваний формат."
        sf.seek(0, os.SEEK_END)
        size = sf.tell()
        sf.seek(0)
        if size > MAX_SCAN_SIZE_BYTES:
            size_mb = round(size / 1024 / 1024, 1)
            return f"Файл «{sf.filename}»{label} завеликий ({size_mb} МБ)."
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


def _parse_pdf_ajax_response(pdf_bytes):
    """Парсить PDF «Резерв+» і повертає JSON. Нічого в БД не записує."""
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
        from routes.utils import title_case_ukrainian
    except ImportError:
        def title_case_ukrainian(s): return s or ''

    full_name = title_case_ukrainian((parsed.get('full_name') or '').strip())
    name_parts = full_name.split()
    last_name = name_parts[0] if len(name_parts) > 0 else ''
    first_name = name_parts[1] if len(name_parts) > 1 else ''
    middle_name = ' '.join(name_parts[2:]) if len(name_parts) > 2 else ''

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
        'tax_id': parsed.get('tax_id') or '',
        'phone': phone,
        'email': parsed.get('email') or '',
        'military_status_badge': parsed.get('status_badge') or '',
        'military_valid_until': parsed.get('valid_until') or '',
        'military_registration_category': parsed.get('registration_category') or '',
        'military_tck_name': parsed.get('tck_name') or '',
        'military_rank': parsed.get('military_rank') or '',
        'military_registry_number': parsed.get('registry_number') or '',
        'military_vos_code': parsed.get('vos_code') or '',
        'military_vos_description': parsed.get('vos_description') or '',
        'military_deferral_type': parsed.get('deferral_type') or '',
        'military_deferral_until': parsed.get('deferral_until') or '',
        'military_vlk_decision': parsed.get('vlk_decision') or '',
        'military_vlk_date': parsed.get('vlk_date') or '',
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

    import uuid as _uuid
    os.makedirs(PENDING_PDFS_DIR, exist_ok=True)
    pdf_safe_name = f"{_uuid.uuid4().hex}.pdf"
    pdf_abs_path = os.path.join(PENDING_PDFS_DIR, pdf_safe_name)
    with open(pdf_abs_path, 'wb') as f:
        f.write(pdf_bytes)
    pdf_rel_path = f"uploads/pending_pdfs/{pdf_safe_name}"

    logger.info(
        f"[parse_pdf_update] saved: abs={pdf_abs_path}, rel={pdf_rel_path}, "
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

    if request.method == 'POST' and request.form.get('action') == 'parse_pdf':
        file = request.files.get('pdf_file')
        if not file or not file.filename:
            conn.close()
            return jsonify({'ok': False, 'error': 'Оберіть PDF-файл'}), 400
        if not file.filename.lower().endswith('.pdf'):
            conn.close()
            return jsonify({'ok': False, 'error': 'Підтримуються лише PDF-файли'}), 400
        pdf_bytes = file.read()
        if len(pdf_bytes) > MAX_PDF_SIZE_BYTES:
            conn.close()
            return jsonify({'ok': False, 'error': 'Файл завеликий'}), 400
        conn.close()
        return _parse_pdf_ajax_response(pdf_bytes)

    if request.method == 'GET':
        conn.close()
        return render_template(
            'public_update_form.html',
            student=student, allowed_fields=allowed_fields,
            simple_fields=SIMPLE_FIELDS, section_labels=SECTION_LABELS,
        )

    # ---- POST ----
    if (request.form.get('website') or '').strip():
        conn.close()
        return render_template('public_update_thanks.html')

    def f(name):
        return (request.form.get(name) or '').strip() or None

    submitted = {}

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
            passport_doc_type = f('passport_document_type') if f('passport_document_type') in ('Паспорт (книжка)', 'ID-картка') else None
            submitted['passport'] = {
                'document_type': passport_doc_type,
                'series': f('passport_series'),
                'number': f('passport_number'),
                'issued_by': f('passport_issued_by'),
                'issue_date': _iso_to_ua_date(f('passport_issue_date')),
                'valid_until': _iso_to_ua_date(f('passport_valid_until')) if passport_doc_type == 'ID-картка' else None,
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

    if 'education_document_reference' in allowed_fields:
        if any([f('reference_number'), f('reference_institution')]):
            submitted.setdefault('education_document', {})
            submitted['education_document']['reference'] = {
                'number': f('reference_number'),
                'institution': f('reference_institution'),
                'country': f('reference_country'),
                'issue_date': _iso_to_ua_date(f('reference_issue_date')),
            }

    if 'education_document_recognition' in allowed_fields:
        if any([f('recognition_certificate_number'), f('recognition_issuer')]):
            submitted.setdefault('education_document', {})
            submitted['education_document']['recognition'] = {
                'certificate_number': f('recognition_certificate_number'),
                'issuer': f('recognition_issuer'),
                'date': _iso_to_ua_date(f('recognition_date')),
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
            'status_badge': f('military_status_badge'),
            'valid_until': f('military_valid_until'),
            'registration_category': f('military_registration_category'),
            'tck_name': f('military_tck_name'),
            'military_rank': f('military_rank'),
            'registry_number': f('military_registry_number'),
            'vos_code': f('military_vos_code'),
            'vos_description': f('military_vos_description'),
            'deferral_type': f('military_deferral_type'),
            'deferral_until': f('military_deferral_until'),
            'vlk_decision': f('military_vlk_decision'),
            'vlk_date': f('military_vlk_date'),
            'address_of_residence': f('military_address'),
            'registration_number_of_the_DRPVR': f('military_registration_number_drpvr'),
            'military_registration_document': f('military_registration_document'),
            'issued_VOD': _iso_to_ua_date(f('military_issued_vod')),
            'military_accounting_specialty_number': f('military_specialty_number'),
        }

    conn.execute(
        "UPDATE update_requests SET status='submitted', submitted_data=?, photo_path=?, submitted_at=datetime('now','localtime') WHERE id=?",
        (json.dumps(submitted, ensure_ascii=False), photo_path, row['id'])
    )
    conn.commit()

    # ═══════════════════════════════════════════════════════════
    # Скани
    # ═══════════════════════════════════════════════════════════
    if 'passport' in allowed_fields:
        scans = [sf for sf in request.files.getlist('passport_scans') if sf and sf.filename]
        if scans:
            save_multiple_attachments(conn, 'update_request_passport', row['id'], scans, 'update_requests')
    if 'education_document' in allowed_fields:
        scans = [sf for sf in request.files.getlist('document_scans') if sf and sf.filename]
        if scans:
            save_multiple_attachments(conn, 'update_request_education', row['id'], scans, 'update_requests')

    # ── PDF «Резерв+» — обробляємо НЕЗАЛЕЖНО від 'military' у
    # allowed_fields. Логіка: якщо студент якимось чином завантажив
    # PDF через «Швидке заповнення» (навіть якщо адмін не дозволив
    # military) — файл все одно має бути в сканах заявки, щоб адмін
    # міг його побачити й вирішити, чи застосовувати.
    if 'military' in allowed_fields:
        scans = [sf for sf in request.files.getlist('military_scans') if sf and sf.filename]
        if scans:
            save_multiple_attachments(conn, 'update_request_military', row['id'], scans, 'update_requests')

    # Спроба 1: файл напряму з поля pdf_file
    pdf_file = request.files.get('pdf_file')
    logger.info(
        f"[update_info] pdf_file: filename={getattr(pdf_file, 'filename', None)!r}"
    )
    pdf_saved = False

    if pdf_file and pdf_file.filename and pdf_file.filename.lower().endswith('.pdf'):
        pdf_bytes = pdf_file.read()
        logger.info(f"[update_info] pdf_file size: {len(pdf_bytes)} bytes")
        if pdf_bytes.startswith(b'%PDF'):
            import uuid as _uuid
            os.makedirs(UPDATE_REQUESTS_DIR, exist_ok=True)
            safe_name = f"{_uuid.uuid4().hex}.pdf"
            dest_pdf = os.path.join(UPDATE_REQUESTS_DIR, safe_name)
            with open(dest_pdf, 'wb') as fp:
                fp.write(pdf_bytes)
            rel_pdf = f"uploads/update_requests/{safe_name}"
            conn.execute(
                "INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_by) "
                "VALUES (?, ?, ?, ?, ?)",
                ('update_request_military', row['id'], rel_pdf,
                 pdf_file.filename or 'Резерв+.pdf', None),
            )
            logger.info(
                f"[update_info] PDF saved from form: {rel_pdf} ({len(pdf_bytes)} bytes), "
                f"abs={dest_pdf}"
            )
            pdf_saved = True
        else:
            logger.warning(
                f"[update_info] pdf_file does NOT start with %PDF: "
                f"{pdf_bytes[:20]!r}"
            )

    # Спроба 2 (fallback): якщо pdf_file порожній, але є
    # prefilled_pdf_path — копіюємо файл з pending_pdfs/ (його туди
    # зберіг AJAX-парсинг).
    if not pdf_saved:
        prefilled_pdf = (request.form.get('prefilled_pdf_path') or '').strip()
        logger.info(
            f"[update_info] fallback: prefilled_pdf_path={prefilled_pdf!r}"
        )
        if prefilled_pdf and prefilled_pdf.startswith('uploads/pending_pdfs/'):
            src_pdf = os.path.join(PROJECT_ROOT, 'static', prefilled_pdf)
            logger.info(
                f"[update_info] fallback: src_pdf={src_pdf}, "
                f"exists={os.path.exists(src_pdf)}"
            )
            if os.path.exists(src_pdf):
                import shutil
                import uuid as _uuid
                os.makedirs(UPDATE_REQUESTS_DIR, exist_ok=True)
                safe_name = f"{_uuid.uuid4().hex}.pdf"
                dest_pdf = os.path.join(UPDATE_REQUESTS_DIR, safe_name)
                shutil.copy2(src_pdf, dest_pdf)
                rel_pdf = f"uploads/update_requests/{safe_name}"
                conn.execute(
                    "INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_by) "
                    "VALUES (?, ?, ?, ?, ?)",
                    ('update_request_military', row['id'], rel_pdf, 'Резерв+.pdf', None),
                )
                logger.info(f"[update_info] PDF saved from pending_pdfs: {rel_pdf}")
                try:
                    os.remove(src_pdf)
                except OSError:
                    pass
            else:
                logger.warning(f"[update_info] fallback PDF not found: {src_pdf}")

    conn.commit()
    conn.close()
    return render_template('public_update_thanks.html')