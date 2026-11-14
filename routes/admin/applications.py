"""
routes/admin/applications.py
=============================
Заявки на реєстрацію (з публічної анкети /apply) і заявки на
оновлення даних наявного студента (з токен-посилань /update-info) -
модерація, вибіркове застосування полів, ручна переобрізка фото.
"""
from flask import render_template, request, redirect, url_for, flash, session, send_file
from routes.db import get_db
from routes.utils import log_action, permission_required, logger
from routes.helpers import current_username, sort_ukrainian
from routes.admin import admin_bp
import sqlite3
import os
from routes.utils import get_attachments, save_multiple_attachments, generate_english_name
from routes.admin.documents import translate_to_en
from datetime import datetime
import json
from routes.passport_types import (
    PASSPORT_TYPES_ID_STYLE, PASSPORT_TYPE_DEFAULT,
    normalize_passport_type, normalize_passport_number,
    passport_template_context,
)


# ═══════════════════════════════════════════════════════════════════════
# Допоміжні функції (перенос військових даних з PDF у картку студента)
# ═══════════════════════════════════════════════════════════════════════

def _promote_pending_military_to_student(conn, pending_id, student_id, military_id, username):
    """
    При підтвердженні заявки з /apply:
      1. Переприв'язує всі скани «pending_student_military» до студента
         (entity_type='military', entity_id=military_id).
      2. Якщо серед сканів є PDF — перепарсить його і переносить УСІ
         нові військові поля в military (status_badge, tck_name,
         registry_number, vos_code, vos_description, deferral_type,
         deferral_until, vlk_decision, vlk_date, disability_*, etc.).
      3. Якщо в PDF є фото і в студента ще немає фото — зберігає його.

    Використовує COALESCE — НЕ затирає наявні (введені вручну) значення
    порожніми з PDF. Якщо PDF розпізнав поле, а в БД воно порожнє —
    воно заповнюється. Якщо в БД вже щось є — лишається як є.

    Повертає (success: bool, message: str).
    """
    logger.info(
        f"[promote_apply] START pending_id={pending_id}, "
        f"student_id={student_id}, military_id={military_id}"
    )

    military_attachments = conn.execute(
        "SELECT id, file_path, original_name FROM attachments "
        "WHERE entity_type='pending_student_military' AND entity_id=?",
        (pending_id,)
    ).fetchall()

    logger.info(f"[promote_apply] attachments: {len(military_attachments)}")
    for att in military_attachments:
        logger.info(
            f"[promote_apply]   id={att['id']}, "
            f"file_path={att['file_path']!r}, "
            f"original_name={att['original_name']!r}"
        )

    if not military_attachments:
        return False, "Немає сканів військового обліку в заявці"

    pdf_attachment = None
    for att in military_attachments:
        conn.execute(
            "UPDATE attachments SET entity_type='military', entity_id=? WHERE id=?",
            (military_id, att['id'])
        )
        fp = (att['file_path'] or '').lower()
        on = (att['original_name'] or '').lower()
        if pdf_attachment is None and (fp.endswith('.pdf') or on.endswith('.pdf')):
            pdf_attachment = att

    if not pdf_attachment:
        return False, "PDF «Резерв+» не знайдено серед сканів"

    try:
        from routes.parse_military_pdf import parse_reserve_plus_pdf, extract_photo_from_pdf
    except ImportError as e:
        logger.error(f"[promote_apply] parse_military_pdf: {e}", exc_info=True)
        return False, f"Модуль парсера PDF недоступний: {e}"

    pdf_abs = os.path.join('static', pdf_attachment['file_path'])
    if not os.path.exists(pdf_abs):
        return False, f"PDF не знайдено на диску: {pdf_abs}"

    try:
        parsed = parse_reserve_plus_pdf(pdf_abs)
    except Exception as e:
        logger.error(f"[promote_apply] parse_reserve_plus_pdf: {e}", exc_info=True)
        return False, f"Не вдалося розпізнати PDF: {e}"

    try:
        photo_bytes = extract_photo_from_pdf(pdf_abs)
    except Exception as e:
        logger.warning(f"[promote_apply] extract_photo_from_pdf: {e}")
        photo_bytes = None

    new_military = {
        'status_badge': parsed.get('status_badge') or '',
        'valid_until': parsed.get('valid_until') or '',
        'registration_category': parsed.get('registration_category') or '',
        'removal_basis': parsed.get('removal_basis') or '',
        'tck_name': parsed.get('tck_name') or '',
        'military_rank': parsed.get('military_rank') or '',
        'registry_number': parsed.get('registry_number') or '',
        'vos_code': parsed.get('vos_code') or '',
        'vos_description': parsed.get('vos_description') or '',
        'deferral_type': parsed.get('deferral_type') or '',
        'deferral_until': parsed.get('deferral_until') or '',
        'police_reason': parsed.get('police_reason') or '',
        'police_date': parsed.get('police_date') or '',
        'vlk_decision': parsed.get('vlk_decision') or '',
        'vlk_date': parsed.get('vlk_date') or '',
        'disability_group': parsed.get('disability_group') or '',
        'disability_valid_until': parsed.get('disability_valid_until') or '',
        'disability_reason': parsed.get('disability_reason') or '',
        'formed_at': parsed.get('formed_at') or '',
        'data_update_date': parsed.get('data_update_date') or '',
        'military_registration_document': parsed.get('military_registration_document') or '',
    }

    string_cols = [c for c in new_military.keys() if c != 'data_verified_on_time']
    set_clause = ', '.join(
        f"{c} = COALESCE(NULLIF(?, ''), {c})" for c in string_cols
    )
    verified_value = 1 if parsed.get('data_verified_on_time') else 0
    try:
        conn.execute(
            f"""
            UPDATE military SET
                {set_clause},
                data_verified_on_time = CASE
                    WHEN COALESCE(data_verified_on_time, 0) = 0 THEN ?
                    ELSE data_verified_on_time
                END
            WHERE id = ?
            """,
            [new_military[c] for c in string_cols]
            + [verified_value, military_id]
        )
        logger.info(
            f"[promote_apply] military UPDATE OK: "
            f"status_badge={new_military['status_badge']!r}, "
            f"tck_name={new_military['tck_name']!r}, "
            f"vos_code={new_military['vos_code']!r}"
        )
    except Exception as e:
        logger.error(f"[promote_apply] military UPDATE: {e}", exc_info=True)
        return False, f"Не вдалося оновити військові дані: {e}"

    student_updates = {}
    if parsed.get('address'):
        student_updates['address_of_residence'] = parsed['address']
    if parsed.get('phone'):
        student_updates['phone'] = parsed['phone']
    if parsed.get('email'):
        student_updates['email'] = parsed['email']
    if student_updates:
        try:
            cols = list(student_updates.keys())
            set_clause = ', '.join(f"{c} = COALESCE(NULLIF({c}, ''), ?)" for c in cols)
            conn.execute(
                f"UPDATE students SET {set_clause} WHERE id = ?",
                [student_updates[c] for c in cols] + [student_id],
            )
        except Exception as e:
            logger.warning(f"[promote_apply] students address/contacts update: {e}")

    if photo_bytes:
        try:
            current = conn.execute(
                "SELECT photo FROM students WHERE id = ?", (student_id,)
            ).fetchone()
            if not current or not current['photo']:
                from routes.photo import (
                    load_and_validate_image, auto_center_crop_box,
                    crop_and_resize, save_final_image,
                )
                img = load_and_validate_image(photo_bytes)
                crop = auto_center_crop_box(img.width, img.height)
                final_img = crop_and_resize(img, crop)
                new_rel = save_final_image(final_img, student_id)
                conn.execute(
                    "UPDATE students SET photo = ? WHERE id = ?",
                    (new_rel, student_id)
                )
                logger.info(f"[promote_apply] photo saved: {new_rel}")
        except Exception as e:
            logger.warning(f"[promote_apply] photo save: {e}")

    return True, "Військові дані з PDF перенесено"


def _promote_update_request_military_to_student(conn, request_id, student_id, username):
    """
    При застосуванні /update-info: якщо серед сканів
    'update_request_military' є PDF «Резерв+» — перепарсить його і
    перенесе ВСІ військові поля в military студента.

    Використовує COALESCE — не затирає наявні значення.

    Повертає (success: bool, message: str).
    """
    logger.info(
        f"[promote_update] START request_id={request_id}, student_id={student_id}"
    )

    attachments = conn.execute(
        "SELECT id, file_path, original_name FROM attachments "
        "WHERE entity_type='update_request_military' AND entity_id=?",
        (request_id,)
    ).fetchall()

    logger.info(f"[promote_update] attachments: {len(attachments)}")
    for att in attachments:
        logger.info(
            f"[promote_update]   id={att['id']}, "
            f"file_path={att['file_path']!r}, "
            f"original_name={att['original_name']!r}"
        )

    if not attachments:
        return False, "Немає сканів військового обліку"

    pdf_attachment = None
    for att in attachments:
        fp = (att['file_path'] or '').lower()
        on = (att['original_name'] or '').lower()
        if fp.endswith('.pdf') or on.endswith('.pdf'):
            pdf_attachment = att
            break

    if not pdf_attachment:
        return False, "PDF «Резерв+» не знайдено серед сканів"

    pdf_abs = os.path.join('static', pdf_attachment['file_path'])
    if not os.path.exists(pdf_abs):
        return False, f"PDF не знайдено на диску: {pdf_abs}"

    try:
        from routes.parse_military_pdf import parse_reserve_plus_pdf, extract_photo_from_pdf
        parsed = parse_reserve_plus_pdf(pdf_abs)
        photo_bytes = extract_photo_from_pdf(pdf_abs)
    except Exception as e:
        logger.error(f"[promote_update] parse: {e}", exc_info=True)
        return False, f"Не вдалося розпізнати PDF: {e}"

    military_row = conn.execute(
        "SELECT id FROM military WHERE student_id=?", (student_id,)
    ).fetchone()

    new_fields = {
        'status_badge': parsed.get('status_badge') or '',
        'valid_until': parsed.get('valid_until') or '',
        'registration_category': parsed.get('registration_category') or '',
        'removal_basis': parsed.get('removal_basis') or '',
        'tck_name': parsed.get('tck_name') or '',
        'military_rank': parsed.get('military_rank') or '',
        'registry_number': parsed.get('registry_number') or '',
        'vos_code': parsed.get('vos_code') or '',
        'vos_description': parsed.get('vos_description') or '',
        'deferral_type': parsed.get('deferral_type') or '',
        'deferral_until': parsed.get('deferral_until') or '',
        'police_reason': parsed.get('police_reason') or '',
        'police_date': parsed.get('police_date') or '',
        'vlk_decision': parsed.get('vlk_decision') or '',
        'vlk_date': parsed.get('vlk_date') or '',
        'disability_group': parsed.get('disability_group') or '',
        'disability_valid_until': parsed.get('disability_valid_until') or '',
        'disability_reason': parsed.get('disability_reason') or '',
        'formed_at': parsed.get('formed_at') or '',
        'data_update_date': parsed.get('data_update_date') or '',
    }

    verified_value = 1 if parsed.get('data_verified_on_time') else 0
    cols = list(new_fields.keys())

    try:
        if military_row:
            set_clause = ', '.join(
                f"{c} = COALESCE(NULLIF(?, ''), {c})" for c in cols
            )
            conn.execute(
                f"""
                UPDATE military SET
                    {set_clause},
                    data_verified_on_time = CASE
                        WHEN COALESCE(data_verified_on_time, 0) = 0 THEN ?
                        ELSE data_verified_on_time
                    END
                WHERE id=?
                """,
                [new_fields[c] for c in cols]
                + [verified_value, military_row['id']]
            )
            military_id = military_row['id']
        else:
            all_cols = cols + ['data_verified_on_time']
            placeholders = ', '.join('?' for _ in all_cols)
            conn.execute(
                f"INSERT INTO military (student_id, {', '.join(all_cols)}) "
                f"VALUES (?, {placeholders})",
                [student_id] + [new_fields[c] for c in cols] + [verified_value]
            )
            military_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    except Exception as e:
        logger.error(f"[promote_update] military UPDATE/INSERT: {e}", exc_info=True)
        return False, f"Не вдалося оновити військові дані: {e}"

    student_updates = {}
    if parsed.get('address'):
        student_updates['address_of_residence'] = parsed['address']
    if parsed.get('phone'):
        student_updates['phone'] = parsed['phone']
    if parsed.get('email'):
        student_updates['email'] = parsed['email']
    if student_updates:
        try:
            cols = list(student_updates.keys())
            set_clause = ', '.join(f"{c} = COALESCE(NULLIF({c}, ''), ?)" for c in cols)
            conn.execute(
                f"UPDATE students SET {set_clause} WHERE id = ?",
                [student_updates[c] for c in cols] + [student_id],
            )
        except Exception as e:
            logger.warning(f"[promote_update] students address/contacts update: {e}")

    if photo_bytes:
        try:
            cur = conn.execute(
                "SELECT photo FROM students WHERE id=?", (student_id,)
            ).fetchone()
            if not cur or not cur['photo']:
                from routes.photo import (
                    load_and_validate_image, auto_center_crop_box,
                    crop_and_resize, save_final_image,
                )
                img = load_and_validate_image(photo_bytes)
                crop = auto_center_crop_box(img.width, img.height)
                final = crop_and_resize(img, crop)
                new_rel = save_final_image(final, student_id)
                conn.execute("UPDATE students SET photo=? WHERE id=?", (new_rel, student_id))
        except Exception as e:
            logger.warning(f"[promote_update] photo: {e}")
    
    # ── Переприв'язуємо скани (в т.ч. цей PDF) до military студента, ──
    # щоб вони не «зникали» разом із заявкою після застосування.
    try:
        pending_atts = conn.execute(
            "SELECT id FROM attachments "
            "WHERE entity_type='update_request_military' AND entity_id=?",
            (request_id,)
        ).fetchall()
        for att in pending_atts:
            conn.execute(
                "UPDATE attachments SET entity_type='military', entity_id=? WHERE id=?",
                (military_id, att['id'])
            )
    except Exception as e:
        logger.warning(f"[promote_update] reattach scans: {e}")

    return True, "Військові дані з PDF перенесено"


# ═══════════════════════════════════════════════════════════════════════
# Допоміжні функції для затвердження заявок
# ═══════════════════════════════════════════════════════════════════════

def _pending_has_military_data(row) -> bool:
    """
    Чи є в заявці хоч одне заповнене військове поле (нові або legacy).
    `military_data_verified_on_time` — лише чекбокс, не рахуємо як дані.
    """
    fields = (
        'military_registration_number_drpvr',
        'military_specialty_number',
        'military_issued_vod',
        'military_change_credentials',
        'military_change_reason',
        'military_registry_number',
        'military_status_badge',
        'military_valid_until',
        'military_registration_category',
        'military_removal_basis',
        'military_tck_name',
        'military_vos_code',
        'military_vos_description',
        'military_deferral_type',
        'military_deferral_until',
        'military_police_reason',
        'military_police_date',
        'military_vlk_decision',
        'military_vlk_date',
        'military_disability_group',
        'military_disability_valid_until',
        'military_disability_reason',
        'military_formed_at',
        'military_data_update_date',
        'military_registration_document',
        'military_rank',
    )
    for f in fields:
        try:
            if row[f]:
                return True
        except (KeyError, IndexError):
            continue
    return False


def _military_payload_from_pending(row) -> dict:
    """Будує dict {column_in_military_table: value} з полів заявки."""
    return {
        'registry_number': row['military_registry_number']
                           or row['military_registration_number_drpvr'],
        'formed_at': row['military_formed_at'] or row['military_issued_vod'],
        'vos_code': row['military_vos_code'] or row['military_specialty_number'],
        'military_registration_document': row['military_registration_document'],
        'military_rank': row['military_rank'],
        'change_credentials': row['military_change_credentials'],
        'reason_for_changing_credentials': row['military_change_reason'],
        'status_badge': row['military_status_badge'],
        'valid_until': row['military_valid_until'],
        'registration_category': row['military_registration_category'],
        'removal_basis': row['military_removal_basis'],
        'tck_name': row['military_tck_name'],
        'vos_description': row['military_vos_description'],
        'deferral_type': row['military_deferral_type'],
        'deferral_until': row['military_deferral_until'],
        'police_reason': row['military_police_reason'],
        'police_date': row['military_police_date'],
        'vlk_decision': row['military_vlk_decision'],
        'vlk_date': row['military_vlk_date'],
        'disability_group': row['military_disability_group'],
        'disability_valid_until': row['military_disability_valid_until'],
        'disability_reason': row['military_disability_reason'],
        'data_verified_on_time': row['military_data_verified_on_time'] or 0,
        'data_update_date': row['military_data_update_date'],
    }


def _find_existing_passport(conn, student_id, group_id, document_type, series, number):
    """Шукає наявний паспорт того самого студента у тій самій групі."""
    return conn.execute("""
        SELECT id FROM passport_documents
        WHERE student_id = ?
          AND COALESCE(group_id, -1) = COALESCE(?, -1)
          AND document_type = ?
          AND NORM_SERIES(series) = NORM_SERIES(?)
          AND NORM_PASSPORT(number, document_type)
              = NORM_PASSPORT(?, document_type)
        ORDER BY id ASC
        LIMIT 1
    """, (student_id, group_id, document_type, series, number)).fetchone()


def _find_existing_education_doc(conn, student_id, group_id, document_type, document_number):
    """Шукає наявний документ про освіту того самого студента у тій самій групі."""
    return conn.execute("""
        SELECT id FROM education_documents
        WHERE student_id = ?
          AND COALESCE(group_id, -1) = COALESCE(?, -1)
          AND document_type = ?
          AND NORM_DOC_NUM(document_number) = NORM_DOC_NUM(?)
        LIMIT 1
    """, (student_id, group_id, document_type or '', document_number or '')).fetchone()

def _load_manual_edits(row) -> list:
    """Розбирає JSON-поле pending_students.manual_edits."""
    if 'manual_edits' not in row.keys() or not row['manual_edits']:
        return []
    try:
        data = json.loads(row['manual_edits'])
    except (ValueError, TypeError):
        return []
    return data if isinstance(data, list) else []


# Підписи полів заявки — для manual_edits (щоб у «Підтверджені» було видно
# людську назву, а не key з БД).
_PENDING_FIELD_LABELS = {
    'last_name_UA': 'Прізвище (укр.)',
    'first_name_UA': "Ім'я (укр.)",
    'middle_name_UA': 'По батькові',
    'last_name_ENG': 'Прізвище (лат.)',
    'first_name_ENG': "Ім'я (лат.)",
    'birth_date': 'Дата народження',
    'tax_id': 'ІПН',
    'phone': 'Телефон',
    'phone_backup': 'Резервний телефон',
    'email': 'Email',
    'document_type': 'Документ: тип',
    'document_series': 'Документ: серія',
    'document_number': 'Документ: номер',
    'document_institution': 'Документ: заклад',
    'document_country': 'Документ: країна',
    'document_date': 'Документ: дата',
    'foreign_reference_number': 'Довідка: номер',
    'foreign_reference_institution': 'Довідка: заклад',
    'foreign_reference_country': 'Довідка: країна',
    'foreign_reference_issue_date': 'Довідка: дата',
    'recognition_certificate_number': 'Свідоцтво: номер',
    'recognition_issuer': 'Свідоцтво: орган',
    'recognition_date': 'Свідоцтво: дата',
    'passport_document_type': 'Паспорт: тип',
    'passport_series': 'Паспорт: серія',
    'passport_number': 'Паспорт: номер',
    'passport_issued_by': 'Паспорт: ким видано',
    'passport_issue_date': 'Паспорт: дата видачі',
    'passport_valid_until': 'Паспорт: дійсний до',
    'passport_unique_number': 'Паспорт: ЄДДР',
    'military_registration_number_drpvr': 'Військ.: реєстр. номер',
    'military_registration_document': 'Військ.: документ',
    'military_issued_vod': 'Військ.: дата документа',
    'military_specialty_number': 'Військ.: номер ВОС',
    'military_rank': 'Військ.: звання',
    'military_address': 'Військ.: адреса',
    'military_change_credentials': 'Військ.: зміна даних',
    'military_change_reason': 'Військ.: причина зміни',
    'military_status_badge': 'Військ.: статус',
    'military_valid_until': 'Військ.: дійсний до',
    'military_registration_category': 'Військ.: категорія обліку',
    'military_removal_basis': 'Військ.: підстава зняття',
    'military_tck_name': 'Військ.: ТЦК',
    'military_registry_number': 'Військ.: номер в реєстрі',
    'military_vos_code': 'Військ.: код ВОС',
    'military_vos_description': 'Військ.: опис ВОС',
    'military_deferral_type': 'Військ.: тип відстрочки',
    'military_deferral_until': 'Військ.: відстрочка до',
    'military_police_reason': 'Військ.: причина поліції',
    'military_police_date': 'Військ.: дата поліції',
    'military_vlk_decision': 'Військ.: постанова ВЛК',
    'military_vlk_date': 'Військ.: дата ВЛК',
    'military_disability_group': 'Військ.: група інвалідності',
    'military_disability_valid_until': 'Військ.: інвалідність до',
    'military_disability_reason': 'Військ.: причина інвалідності',
    'military_formed_at': 'Військ.: документ сформовано',
    'military_data_update_date': 'Військ.: дата уточнення',
    'military_data_verified_on_time': 'Військ.: дані уточнено вчасно',
}

def _load_applied_changes(row) -> list:
    """
    Розбирає JSON-поле pending_students.applied_changes.
    Повертає список [{'label','old','new'}, ...] або [].
    """
    if 'applied_changes' not in row.keys() or not row['applied_changes']:
        return []
    try:
        data = json.loads(row['applied_changes'])
    except (ValueError, TypeError):
        return []
    if not isinstance(data, list):
        return []
    return data


# ═══════════════════════════════════════════════════════════════════════
# Заявки на реєстрацію (/apply)
# ═══════════════════════════════════════════════════════════════════════

@admin_bp.route('/admin/pending_students')
@permission_required('manage_pending_students')
def pending_students():
    """Список заявок з публічної анкети самореєстрації."""
    conn = get_db()
    conn.row_factory = sqlite3.Row

    status_filter = request.args.get('status', 'new')
    if status_filter not in ('new', 'approved', 'rejected'):
        status_filter = 'new'

    rows = conn.execute("""
        SELECT * FROM pending_students WHERE status = ?
        ORDER BY created_at DESC
    """, (status_filter,)).fetchall()
    rows = [dict(r) for r in rows]
    for r in rows:
        r['scans_count'] = conn.execute(
            "SELECT COUNT(*) FROM attachments WHERE entity_type='pending_student' AND entity_id=?",
            (r['id'],)
        ).fetchone()[0]
        r['has_military_scan'] = conn.execute(
            "SELECT 1 FROM attachments "
            "WHERE entity_type='pending_student_military' AND entity_id=? "
            "AND (LOWER(original_name) LIKE '%.pdf' OR LOWER(file_path) LIKE '%.pdf') "
            "LIMIT 1",
            (r['id'],)
        ).fetchone() is not None

    duplicates_by_id = {}
    if status_filter == 'new':
        for row in rows:
            existing_student = conn.execute("""
                SELECT id, last_name_UA, first_name_UA FROM students
                WHERE LOWER_UA(last_name_UA) = LOWER_UA(?) AND LOWER_UA(first_name_UA) = LOWER_UA(?)
                  AND birth_date = ?
            """, (row['last_name_UA'], row['first_name_UA'], row['birth_date'])).fetchone()

            other_pending_count = conn.execute("""
                SELECT COUNT(*) FROM pending_students
                WHERE status = 'new' AND id != ?
                  AND LOWER_UA(last_name_UA) = LOWER_UA(?) AND LOWER_UA(first_name_UA) = LOWER_UA(?)
                  AND birth_date = ?
            """, (row['id'], row['last_name_UA'], row['first_name_UA'], row['birth_date'])).fetchone()[0]

            if existing_student or other_pending_count:
                duplicates_by_id[row['id']] = {
                    'existing_student': existing_student,
                    'other_pending_count': other_pending_count,
                }

    counts = {
        s: conn.execute("SELECT COUNT(*) FROM pending_students WHERE status=?", (s,)).fetchone()[0]
        for s in ('new', 'approved', 'rejected')
    }

    conn.close()
    return render_template(
        'admin_pending_students.html',
        rows=rows, status_filter=status_filter, counts=counts,
        duplicates_by_id=duplicates_by_id,
    )


@admin_bp.route('/admin/pending_students/<int:pending_id>', methods=['GET', 'POST'])
@permission_required('manage_pending_students')
def pending_student_review(pending_id):
    """Перегляд однієї заявки: виправлення, підтвердження або відхилення."""
    conn = get_db()
    conn.row_factory = sqlite3.Row

    row = conn.execute("SELECT * FROM pending_students WHERE id=?", (pending_id,)).fetchone()
    if not row:
        conn.close()
        flash("Заявку не знайдено", "error")
        return redirect(url_for('admin.pending_students'))

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'save':
            fields = [
                'last_name_UA', 'first_name_UA', 'middle_name_UA', 'last_name_ENG', 'first_name_ENG', 'birth_date', 'tax_id',
                'phone', 'phone_backup', 'email',
                'document_type', 'document_series', 'document_number', 'document_institution',
                'document_country', 'document_date',
                'foreign_reference_number', 'foreign_reference_institution', 'foreign_reference_country',
                'foreign_reference_issue_date', 'recognition_certificate_number', 'recognition_issuer', 'recognition_date',
                'passport_document_type', 'passport_series', 'passport_number', 'passport_issued_by',
                'passport_issue_date', 'passport_valid_until', 'passport_unique_number',
                'military_registration_number_drpvr', 'military_registration_document', 'military_issued_vod',
                'military_specialty_number', 'military_rank', 'military_address',
                'military_change_credentials', 'military_change_reason',
                'military_status_badge', 'military_valid_until', 'military_registration_category',
                'military_removal_basis', 'military_tck_name', 'military_registry_number',
                'military_vos_code', 'military_vos_description',
                'military_deferral_type', 'military_deferral_until',
                'military_police_reason', 'military_police_date',
                'military_vlk_decision', 'military_vlk_date',
                'military_disability_group', 'military_disability_valid_until', 'military_disability_reason',
                'military_formed_at', 'military_data_update_date',
            ]
            from routes.utils import normalize_phone_pair

            values = {f: (request.form.get(f) or '').strip() or None for f in fields}
            # ── ЗБІР ПРАВОК АДМІНА (old → new), з накопиченням ──
            # row — це поточний рядок pending_students із БД (стан ДО збереження).
            # Якщо адмін редагує ту саму заявку вдруге — зберігаємо ПЕРВІСНЕ old,
            # а new оновлюємо до актуального, щоб diff відображав:
            #   «як було подано студентом» → «як стало після всіх правок адміна».
            values = {f: (request.form.get(f) or '').strip() or None for f in fields}
            values['military_data_verified_on_time'] = 1 if request.form.get('military_data_verified_on_time') else 0
            values['phone'], values['phone_backup'] = normalize_phone_pair(
                values.get('phone'), values.get('phone_backup')
            )

            if not values.get('passport_number'):
                values['passport_document_type'] = None
                values['passport_series'] = None
                values['passport_issued_by'] = None
                values['passport_issue_date'] = None
                values['passport_valid_until'] = None
                values['passport_unique_number'] = None
            elif values.get('passport_document_type') not in PASSPORT_TYPES_ID_STYLE:
                values['passport_valid_until'] = None
            existing_manual = _load_manual_edits(row)
            accumulated = {
                item['field']: dict(item)
                for item in existing_manual
                if isinstance(item, dict) and 'field' in item
            }

            diff_fields = list(fields) + ['military_data_verified_on_time']
            # military_data_verified_on_time не входить у fields, але це теж редаговане
            # поле — додаємо окремо, щоб зміна галочки потрапляла в manual_edits.
            diff_fields = list(fields) + ['military_data_verified_on_time']

            for f in diff_fields:
                if f not in row.keys():
                    continue
                old_val = row[f]
                new_val = values.get(f)

                if f == 'military_data_verified_on_time':
                    old_s = 'так' if int(old_val or 0) else 'ні'
                    new_s = 'так' if int(new_val or 0) else 'ні'
                else:
                    old_s = str(old_val or '').strip()
                    new_s = str(new_val or '').strip()

                if old_s == new_s:
                    continue

                if f in accumulated:
                    accumulated[f]['new'] = new_s
                else:
                    accumulated[f] = {
                        'field': f,
                        'label': _PENDING_FIELD_LABELS.get(f, f),
                        'old': old_s,
                        'new': new_s,
                    }
                old_val = row[f]
                new_val = values.get(f)

                if f == 'military_data_verified_on_time':
                    old_s = 'так' if int(old_val or 0) else 'ні'
                    new_s = 'так' if int(new_val or 0) else 'ні'
                else:
                    old_s = str(old_val or '').strip()
                    new_s = str(new_val or '').strip()

                if old_s == new_s:
                    continue

                if f in accumulated:
                    accumulated[f]['new'] = new_s
                else:
                    accumulated[f] = {
                        'field': f,
                        'label': _PENDING_FIELD_LABELS.get(f, f),
                        'old': old_s,
                        'new': new_s,
                    }

            # Прибираємо записи, де old == new (адмін змінив туди й назад)
            cleaned_manual = [
                item for item in accumulated.values()
                if item['old'] != item['new']
            ]
            manual_edits_json = (
                json.dumps(cleaned_manual, ensure_ascii=False)
                if cleaned_manual else None
            )

            if not values['last_name_UA'] or not values['first_name_UA'] or not values['birth_date']:
                flash("Прізвище, ім'я і дата народження обов'язкові", "error")
            else:
                update_fields = [f for f in fields] + ['military_data_verified_on_time']
                update_values = [values[f] for f in fields] + [values['military_data_verified_on_time']]

                set_clause = ", ".join(f"{k}=?" for k in update_fields)
                conn.execute(
                    f"UPDATE pending_students SET {set_clause}, manual_edits=? WHERE id=?",
                    update_values + [manual_edits_json, pending_id]
                )
                conn.commit()
                flash("Зміни збережено", "success")
            conn.close()
            return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
            
            
            
            
            

        elif action == 'reject':
            note = (request.form.get('review_note') or '').strip() or None
            conn.execute(
                "UPDATE pending_students SET status='rejected', reviewed_by=?, "
                "reviewed_at=datetime('now','localtime'), review_note=? WHERE id=?",
                (current_username(), note, pending_id)
            )
            conn.commit()
            log_action(
                current_username(),
                f"відхилив заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} (заявка ID {pending_id})",
                details=note or ''
            )
            conn.close()
            flash("Заявку відхилено", "success")
            return redirect(url_for('admin.pending_students'))

        elif action == 'delete':
            if row['photo_path']:
                photo_abs = os.path.join('static', row['photo_path'])
                if os.path.exists(photo_abs):
                    os.remove(photo_abs)
            for entity_type in ('pending_student', 'pending_student_passport', 'pending_student_military'):
                for att in get_attachments(conn, entity_type, pending_id):
                    att_abs = os.path.join('static', att['file_path'])
                    if os.path.exists(att_abs):
                        os.remove(att_abs)
                conn.execute(
                    "DELETE FROM attachments WHERE entity_type=? AND entity_id=?",
                    (entity_type, pending_id)
                )
            conn.execute("DELETE FROM pending_students WHERE id=?", (pending_id,))
            conn.commit()
            log_action(
                current_username(),
                f"видалив заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} (заявка ID {pending_id})"
            )
            conn.close()
            flash("Заявку видалено", "success")
            return redirect(url_for('admin.pending_students', status=row['status']))

        elif action == 'approve':
            group_id = request.form.get('group_id')
            if not group_id:
                flash("Оберіть групу, щоб підтвердити заявку", "error")
                conn.close()
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
            group_id = int(group_id)
            license_id = request.form.get('license_id') or None
            program_credits_override_raw = (request.form.get('program_credits_override') or '').strip()
            program_credits_override = int(program_credits_override_raw) if program_credits_override_raw else None

            force_new_student = request.form.get('force_new_student') == '1'

            # ── Змінні, які заповнюються під час обробки й потім ідуть у diff ──
            applied_changes = []
            existing_student = None
            education_doc_id = None
            passport_id = None
            military_id = None

            try:
                if not force_new_student:
                    if row['tax_id']:
                        existing_student = conn.execute(
                            "SELECT id FROM students WHERE tax_id = ?", (row['tax_id'],)
                        ).fetchone()

                    if not existing_student:
                        existing_student = conn.execute("""
                            SELECT id FROM students
                            WHERE LOWER_UA(last_name_UA) = LOWER_UA(?)
                              AND LOWER_UA(first_name_UA) = LOWER_UA(?)
                              AND birth_date = ?
                        """, (row['last_name_UA'], row['first_name_UA'], row['birth_date'])).fetchone()

                if not existing_student and row['phone']:
                    from routes.utils import normalize_phone_ua
                    norm_phone = normalize_phone_ua(row['phone'])
                    if norm_phone and norm_phone.startswith('+380'):
                        existing_student = conn.execute(
                            "SELECT id FROM students WHERE phone = ? OR phone_backup = ? LIMIT 1",
                            (norm_phone, norm_phone)
                        ).fetchone()

                from routes.enrollments_service import add_enrollment
                from routes.utils import normalize_phone_pair
                from routes.gender import detect_gender_from_middle_name

                if existing_student:
                    student_id = existing_student['id']
                    conn.execute("UPDATE students SET archived = FALSE WHERE id = ?", (student_id,))

                    has_active = conn.execute(
                        "SELECT 1 FROM student_enrollments "
                        "WHERE student_id = ? AND status = 'active' LIMIT 1",
                        (student_id,),
                    ).fetchone() is not None

                    add_enrollment(
                        student_id, group_id,
                        is_primary=not has_active,
                        license_id=license_id,
                        edebo_code=None,
                        program_credits_override=program_credits_override,
                        conn=conn,
                    )
                else:
                    gender = detect_gender_from_middle_name(row['middle_name_UA'])
                    cur = conn.execute("""
                        INSERT INTO students (
                            last_name_UA, first_name_UA, middle_name_UA, last_name_ENG, first_name_ENG,
                            birth_date, tax_id, group_id, license_id,
                            phone, phone_backup, email, gender
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        row['last_name_UA'], row['first_name_UA'], row['middle_name_UA'],
                        row['last_name_ENG'], row['first_name_ENG'], row['birth_date'], row['tax_id'],
                        group_id, license_id,
                        *normalize_phone_pair(row['phone'], row['phone_backup']),
                        row['email'],
                        gender,
                    ))
                    student_id = cur.lastrowid

                    add_enrollment(
                        student_id, group_id,
                        is_primary=True,
                        license_id=license_id,
                        edebo_code=None,
                        program_credits_override=program_credits_override,
                        conn=conn,
                    )

                # ── Назва групи для diff ──
                _grp = conn.execute(
                    "SELECT name FROM groups WHERE id = ?", (group_id,)
                ).fetchone()
                group_label = _grp['name'] if _grp else f"#{group_id}"

                if existing_student:
                    applied_changes.append({
                        'label': 'Навчання',
                        'old': '',
                        'new': f'додано в групу «{group_label}»',
                    })
                else:
                    applied_changes.append({
                        'label': 'Студент',
                        'old': '',
                        'new': f'створено нового студента (ID {student_id})',
                    })

                if row['photo_path']:
                    try:
                        from routes.photo import photo_path_for_student, _ensure_dir
                        _ensure_dir()
                        src_path = os.path.join('static', row['photo_path'])
                        if os.path.exists(src_path):
                            dest_path = photo_path_for_student(student_id)
                            with open(src_path, 'rb') as src, open(dest_path, 'wb') as dst:
                                dst.write(src.read())
                            conn.execute(
                                "UPDATE students SET photo=? WHERE id=?",
                                (f"uploads/photos/student_{student_id}.jpg", student_id)
                            )
                            applied_changes.append({
                                'label': 'Фото',
                                'old': '',
                                'new': 'перенесено із заявки',
                            })
                    except Exception as e:
                        logger.error(
                            f"Не вдалося перенести фото заявки {pending_id} "
                            f"студенту {student_id}: {e}"
                        )

                if row['document_type'] or row['document_number']:
                    doc_number = ' '.join(
                        x for x in [row['document_series'], row['document_number']] if x
                    ) or ''
                    _doc_type = row['document_type'] or ''

                    existing_doc = _find_existing_education_doc(
                        conn, student_id, group_id, _doc_type, doc_number
                    )

                    if existing_doc:
                        education_doc_id = existing_doc['id']
                        conn.execute("""
                            UPDATE education_documents SET
                                document_type_en    = COALESCE(NULLIF(?, ''), document_type_en),
                                institution_name    = COALESCE(NULLIF(?, ''), institution_name),
                                institution_name_en = COALESCE(NULLIF(?, ''), institution_name_en),
                                country             = COALESCE(NULLIF(?, ''), country),
                                country_en          = COALESCE(NULLIF(?, ''), country_en),
                                completion_date     = COALESCE(NULLIF(?, ''), completion_date)
                            WHERE id = ?
                        """, (
                            translate_to_en(_doc_type) or None,
                            row['document_institution'] or None,
                            translate_to_en(row['document_institution'] or '') or None,
                            row['document_country'] or None,
                            translate_to_en(row['document_country'] or '') or None,
                            row['document_date'] or None,
                            education_doc_id,
                        ))
                    else:
                        cur_doc = conn.execute("""
                            INSERT INTO education_documents (
                                student_id, group_id, document_type, document_type_en, document_number,
                                institution_name, institution_name_en, country, country_en, completion_date
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            student_id, group_id,
                            _doc_type,
                            translate_to_en(_doc_type),
                            doc_number,
                            row['document_institution'] or '',
                            translate_to_en(row['document_institution'] or ''),
                            row['document_country'] or '',
                            translate_to_en(row['document_country'] or ''),
                            row['document_date'] or '',
                        ))
                        education_doc_id = cur_doc.lastrowid

                    applied_changes.append({
                        'label': 'Документ про освіту',
                        'old': '',
                        'new': (
                            f'тип: {_doc_type or "—"}, № {doc_number or "—"}'
                        ),
                    })

                    conn.execute(
                        "UPDATE attachments SET entity_type='education_document', entity_id=? "
                        "WHERE entity_type='pending_student' AND entity_id=?",
                        (education_doc_id, pending_id)
                    )

                    if (row['foreign_reference_number'] or row['foreign_reference_institution']
                            or row['recognition_certificate_number'] or row['recognition_issuer']):
                        existing_foreign = conn.execute(
                            "SELECT id FROM foreign_education_docs WHERE education_doc_id=?",
                            (education_doc_id,)
                        ).fetchone()

                        ref_inst_en = translate_to_en(row['foreign_reference_institution']) if row['foreign_reference_institution'] else None
                        ref_country_en = translate_to_en(row['foreign_reference_country']) if row['foreign_reference_country'] else None
                        rec_issuer_en = translate_to_en(row['recognition_issuer']) if row['recognition_issuer'] else None

                        if existing_foreign:
                            conn.execute("""
                                UPDATE foreign_education_docs SET
                                    reference_number               = COALESCE(NULLIF(reference_number, ''), ?),
                                    reference_institution          = COALESCE(NULLIF(reference_institution, ''), ?),
                                    reference_institution_en       = COALESCE(NULLIF(reference_institution_en, ''), ?),
                                    reference_country              = COALESCE(NULLIF(reference_country, ''), ?),
                                    reference_country_en           = COALESCE(NULLIF(reference_country_en, ''), ?),
                                    reference_issue_date           = COALESCE(NULLIF(reference_issue_date, ''), ?),
                                    recognition_certificate_number = COALESCE(NULLIF(recognition_certificate_number, ''), ?),
                                    recognition_issuer             = COALESCE(NULLIF(recognition_issuer, ''), ?),
                                    recognition_issuer_en          = COALESCE(NULLIF(recognition_issuer_en, ''), ?),
                                    recognition_date               = COALESCE(NULLIF(recognition_date, ''), ?)
                                WHERE education_doc_id=?
                            """, (
                                row['foreign_reference_number'],
                                row['foreign_reference_institution'],
                                ref_inst_en,
                                row['foreign_reference_country'],
                                ref_country_en,
                                row['foreign_reference_issue_date'],
                                row['recognition_certificate_number'],
                                row['recognition_issuer'],
                                rec_issuer_en,
                                row['recognition_date'],
                                education_doc_id,
                            ))
                        else:
                            conn.execute("""
                                INSERT INTO foreign_education_docs (
                                    education_doc_id, reference_number, reference_institution,
                                    reference_institution_en, reference_country, reference_country_en,
                                    reference_issue_date, recognition_certificate_number,
                                    recognition_issuer, recognition_issuer_en, recognition_date
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                education_doc_id,
                                row['foreign_reference_number'],
                                row['foreign_reference_institution'],
                                ref_inst_en,
                                row['foreign_reference_country'],
                                ref_country_en,
                                row['foreign_reference_issue_date'],
                                row['recognition_certificate_number'],
                                row['recognition_issuer'],
                                rec_issuer_en,
                                row['recognition_date'],
                            ))
                else:
                    conn.execute(
                        "UPDATE attachments SET entity_type='student', entity_id=? "
                        "WHERE entity_type='pending_student' AND entity_id=?",
                        (student_id, pending_id)
                    )

                if row['passport_number']:
                    _passport_doc_type = (
                        normalize_passport_type(row['passport_document_type'])
                        or PASSPORT_TYPE_DEFAULT
                    )
                    _passport_number = normalize_passport_number(
                        row['passport_number'], _passport_doc_type
                    )

                    existing_passport = _find_existing_passport(
                        conn, student_id, group_id,
                        _passport_doc_type, row['passport_series'], _passport_number,
                    )

                    if existing_passport:
                        passport_id = existing_passport['id']
                        conn.execute("""
                            UPDATE passport_documents SET
                                issued_by     = COALESCE(NULLIF(issued_by, ''), ?),
                                issue_date    = COALESCE(NULLIF(issue_date, ''), ?),
                                valid_until   = COALESCE(NULLIF(valid_until, ''), ?),
                                unique_number = COALESCE(NULLIF(unique_number, ''), ?)
                            WHERE id = ?
                        """, (
                            row['passport_issued_by'],
                            row['passport_issue_date'],
                            row['passport_valid_until'],
                            row['passport_unique_number'],
                            passport_id,
                        ))
                    else:
                        cur_passport = conn.execute("""
                            INSERT INTO passport_documents (
                                student_id, group_id, document_type, series, number,
                                issued_by, issue_date, valid_until, unique_number
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            student_id, group_id, _passport_doc_type,
                            row['passport_series'], _passport_number,
                            row['passport_issued_by'], row['passport_issue_date'],
                            row['passport_valid_until'], row['passport_unique_number'],
                        ))
                        passport_id = cur_passport.lastrowid

                    applied_changes.append({
                        'label': 'Паспорт',
                        'old': '',
                        'new': f'тип: {_passport_doc_type}, № {_passport_number}',
                    })

                    conn.execute(
                        "UPDATE attachments SET entity_type='passport_document', entity_id=? "
                        "WHERE entity_type='pending_student_passport' AND entity_id=?",
                        (passport_id, pending_id)
                    )
                else:
                    conn.execute(
                        "UPDATE attachments SET entity_type='student', entity_id=? "
                        "WHERE entity_type='pending_student_passport' AND entity_id=?",
                        (student_id, pending_id)
                    )

                has_military_scans = conn.execute(
                    "SELECT 1 FROM attachments "
                    "WHERE entity_type='pending_student_military' AND entity_id=? LIMIT 1",
                    (pending_id,)
                ).fetchone() is not None

                if has_military_scans or _pending_has_military_data(row):
                    military_data = _military_payload_from_pending(row)
                    existing_military = conn.execute(
                        "SELECT id FROM military WHERE student_id = ?", (student_id,)
                    ).fetchone()

                    if existing_military:
                        string_keys = [
                            k for k in military_data if k != 'data_verified_on_time'
                        ]
                        set_clause = ', '.join(
                            f"{k} = COALESCE(NULLIF(?, ''), {k})" for k in string_keys
                        )
                        conn.execute(f"""
                            UPDATE military SET
                                {set_clause},
                                data_verified_on_time = CASE
                                    WHEN COALESCE(data_verified_on_time, 0) = 0 THEN ?
                                    ELSE data_verified_on_time
                                END
                            WHERE id = ?
                        """, [military_data[k] for k in string_keys]
                              + [military_data['data_verified_on_time'], existing_military['id']])
                        military_id = existing_military['id']
                    else:
                        cols = list(military_data.keys())
                        placeholders = ', '.join('?' for _ in cols)
                        cur_military = conn.execute(
                            f"INSERT INTO military (student_id, {', '.join(cols)}) "
                            f"VALUES (?, {placeholders})",
                            [student_id] + list(military_data.values())
                        )
                        military_id = cur_military.lastrowid

                    if row['military_address']:
                        conn.execute(
                            "UPDATE students SET address_of_residence = "
                            "COALESCE(NULLIF(address_of_residence, ''), ?) WHERE id = ?",
                            (row['military_address'], student_id),
                        )

                    # ── Diff по кожному воєнному полю (тільки для заповнених у заявці) ──
                    _mil_labels_for_diff = {
                        'status_badge': 'Статус',
                        'registration_category': 'Категорія обліку',
                        'valid_until': 'Дійсний до',
                        'removal_basis': 'Підстава зняття',
                        'tck_name': 'ТЦК',
                        'military_rank': 'Військове звання',
                        'registry_number': 'Номер у реєстрі',
                        'vos_code': 'Код ВОС',
                        'vos_description': 'Опис ВОС',
                        'deferral_type': 'Тип відстрочки',
                        'deferral_until': 'Відстрочка до',
                        'vlk_decision': 'Постанова ВЛК',
                        'vlk_date': 'Дата ВЛК',
                        'disability_group': 'Група інвалідності',
                        'disability_valid_until': 'Інвалідність до',
                        'disability_reason': 'Причина інвалідності',
                        'police_reason': 'Причина поліції',
                        'police_date': 'Дата поліції',
                        'formed_at': 'Документ сформовано',
                        'data_update_date': 'Дата уточнення',
                        'military_registration_document': 'Документ військ. обліку',
                        'change_credentials': 'Зміна даних',
                        'reason_for_changing_credentials': 'Причина зміни',
                    }

                    for _col, _label in _mil_labels_for_diff.items():
                        _new = (military_data.get(_col) or '').strip()
                        if not _new:
                            continue
                        _old = ''
                        if existing_military:
                            _old_row = conn.execute(
                                f"SELECT {_col} AS v FROM military WHERE id = ?",
                                (existing_military['id'],),
                            ).fetchone()
                            _old = (_old_row['v'] if _old_row else '') or ''
                        if str(_old).strip() != _new:
                            applied_changes.append({
                                'label': f'Військ.: {_label}',
                                'old': str(_old),
                                'new': _new,
                            })

                    conn.execute(
                        "UPDATE attachments SET entity_type='military', entity_id=? "
                        "WHERE entity_type='pending_student_military' AND entity_id=?",
                        (military_id, pending_id)
                    )
                else:
                    conn.execute(
                        "UPDATE attachments SET entity_type='student', entity_id=? "
                        "WHERE entity_type='pending_student_military' AND entity_id=?",
                        (student_id, pending_id)
                    )

                conn.execute(
                    "UPDATE pending_students SET status='approved', reviewed_by=?, "
                    "reviewed_at=datetime('now','localtime'), resulting_student_id=?, "
                    "applied_changes=? WHERE id=?",
                    (
                        current_username(),
                        student_id,
                        json.dumps(applied_changes, ensure_ascii=False) if applied_changes else None,
                        pending_id,
                    ),
                )

                if military_id:
                    promote_ok, promote_msg = _promote_pending_military_to_student(
                        conn, pending_id, student_id, military_id, current_username()
                    )
                    if not promote_ok:
                        flash(f"Військові дані не перенесено з PDF: {promote_msg}", "warning")

                conn.commit()

            except sqlite3.IntegrityError as e:
                conn.rollback()
                logger.error(
                    f"[approve] pending_id={pending_id} integrity: {e}",
                    exc_info=True,
                )
                flash(
                    f"Помилка: дані з заявки конфліктують із наявними в базі ({e}). "
                    f"Можливо, паспорт/документ уже додано вручну — перевірте картку студента.",
                    "error",
                )
                conn.close()
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
            except Exception as e:
                conn.rollback()
                logger.error(
                    f"[approve] pending_id={pending_id}: {e}", exc_info=True,
                )
                flash(f"Помилка підтвердження заявки: {e}", "error")
                conn.close()
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

            log_action(
                current_username(),
                f"підтвердив заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} -> студент ID {student_id}",
                group_ids=[group_id],
            )
            conn.close()
            flash(f"Студента {row['last_name_UA']} {row['first_name_UA']} додано", "success")
            return redirect(url_for('students.student_details', student_id=student_id))

        elif action == 'update_existing':
            existing_id = request.form.get('existing_student_id')
            if not existing_id:
                flash("Не вказано, якого студента оновлювати", "error")
                conn.close()
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
            existing_id = int(existing_id)

            # Поточні значення студента (для diff old → new)
            _old_row = conn.execute(
                "SELECT * FROM students WHERE id = ?", (existing_id,)
            ).fetchone()
            existing_full_for_diff = dict(_old_row) if _old_row else {}

            applied_changes = []

            try:
                updated_parts = []

                PERSONAL_FIELD_MAP = [
                    ('last_name_UA',   'last_name_UA',   'Прізвище (укр.)'),
                    ('first_name_UA',  'first_name_UA',  "Ім'я (укр.)"),
                    ('middle_name_UA', 'middle_name_UA', 'По батькові'),
                    ('last_name_ENG',  'last_name_ENG',  'Прізвище (лат.)'),
                    ('first_name_ENG', 'first_name_ENG', "Ім'я (лат.)"),
                    ('birth_date',     'birth_date',     'Дата народження'),
                    ('tax_id',         'tax_id',         'ІПН'),
                ]
                personal_updates = {}
                personal_labels = []
                for form_key, db_col, label in PERSONAL_FIELD_MAP:
                    if request.form.get(f'update_{form_key}'):
                        personal_updates[db_col] = row[db_col]
                        personal_labels.append(label)

                if personal_updates:
                    set_clause = ', '.join(f"{col}=?" for col in personal_updates)
                    conn.execute(
                        f"UPDATE students SET {set_clause} WHERE id=?",
                        list(personal_updates.values()) + [existing_id],
                    )
                    updated_parts.append('особисті: ' + ', '.join(personal_labels))

                    for form_key, db_col, label in PERSONAL_FIELD_MAP:
                        if db_col not in personal_updates:
                            continue
                        old_val = existing_full_for_diff.get(db_col) or ''
                        new_val = personal_updates[db_col] or ''
                        if str(old_val) != str(new_val):
                            applied_changes.append({
                                'label': label,
                                'old': str(old_val),
                                'new': str(new_val),
                            })

                from routes.utils import normalize_phone_pair

                update_phone  = bool(request.form.get('update_phone'))
                update_backup = bool(request.form.get('update_phone_backup'))
                update_email  = bool(request.form.get('update_email'))

                contacts_updates = {}
                if update_phone or update_backup:
                    current = conn.execute(
                        "SELECT phone, phone_backup FROM students WHERE id=?",
                        (existing_id,),
                    ).fetchone()
                    cur_phone  = current['phone']        if current else None
                    cur_backup = current['phone_backup'] if current else None

                    new_phone_val  = row['phone']        if update_phone  else cur_phone
                    new_backup_val = row['phone_backup'] if update_backup else cur_backup
                    norm_phone, norm_backup = normalize_phone_pair(
                        new_phone_val, new_backup_val
                    )

                    if update_phone:
                        contacts_updates['phone'] = norm_phone
                    if update_backup:
                        contacts_updates['phone_backup'] = norm_backup

                if update_email:
                    contacts_updates['email'] = row['email']

                if contacts_updates:
                    set_clause = ', '.join(f"{col}=?" for col in contacts_updates)
                    conn.execute(
                        f"UPDATE students SET {set_clause} WHERE id=?",
                        list(contacts_updates.values()) + [existing_id],
                    )
                    updated_parts.append('контакти: ' + ', '.join(contacts_updates.keys()))

                    contacts_labels = {
                        'phone': 'Телефон',
                        'phone_backup': 'Резервний телефон',
                        'email': 'Email',
                    }
                    for col, new_val in contacts_updates.items():
                        old_val = existing_full_for_diff.get(col) or ''
                        if str(old_val) != str(new_val or ''):
                            applied_changes.append({
                                'label': contacts_labels.get(col, col),
                                'old': str(old_val),
                                'new': str(new_val or ''),
                            })

                if request.form.get('update_photo') and row['photo_path']:
                    try:
                        from routes.photo import photo_path_for_student, _ensure_dir
                        _ensure_dir()
                        src_path = os.path.join('static', row['photo_path'])
                        if os.path.exists(src_path):
                            dest_path = photo_path_for_student(existing_id)
                            with open(src_path, 'rb') as src, open(dest_path, 'wb') as dst:
                                dst.write(src.read())
                            conn.execute(
                                "UPDATE students SET photo=? WHERE id=?",
                                (f"uploads/photos/student_{existing_id}.jpg", existing_id)
                            )
                            updated_parts.append('фото')
                            had_photo = bool(existing_full_for_diff.get('photo'))
                            applied_changes.append({
                                'label': 'Фото',
                                'old': 'старе фото' if had_photo else '— порожньо —',
                                'new': 'фото із заявки',
                            })
                    except Exception as e:
                        logger.error(
                            f"Не вдалося перенести фото заявки {pending_id} "
                            f"студенту {existing_id}: {e}"
                        )

                # ── Документ про освіту (dedup + підтримка кількох груп) ──
                doc_ids_for_scans = []
                want_apply_doc = bool(
                    request.form.get('add_document') or request.form.get('update_document')
                )
                if want_apply_doc and (row['document_type'] or row['document_number']):
                    doc_number = ' '.join(
                        x for x in [row['document_series'], row['document_number']] if x
                    ) or ''
                    _doc_type = row['document_type'] or ''

                    target_group_ids = []
                    for gs in request.form.getlist('target_group_ids_document'):
                        try:
                            target_group_ids.append(int(gs))
                        except (ValueError, TypeError):
                            pass
                    if not target_group_ids:
                        _grp_row = conn.execute(
                            "SELECT group_id FROM students WHERE id = ?", (existing_id,)
                        ).fetchone()
                        _gid = _grp_row['group_id'] if _grp_row else None
                        target_group_ids = [_gid] if _gid else [None]

                    for _doc_gid in target_group_ids:
                        existing_doc = _find_existing_education_doc(
                            conn, existing_id, _doc_gid, _doc_type, doc_number
                        )

                        # Fallback: тип в заявці може бути іншою
                        # формулюванням того самого документа (напр.
                        # "Атестат про повну середню освіту" ↔ "Атестат
                        # про повну загальну середню освіту"). Якщо за
                        # парою (тип + номер) не знайшли, але за САМИМ
                        # номером у цій групі є РІВНО ОДИН документ —
                        # оновлюємо його, а не створюємо дубль.
                        if not existing_doc and doc_number:
                            _by_number = conn.execute("""
                                SELECT id FROM education_documents
                                WHERE student_id = ?
                                  AND COALESCE(group_id, -1) = COALESCE(?, -1)
                                  AND NORM_DOC_NUM(document_number) = NORM_DOC_NUM(?)
                            """, (existing_id, _doc_gid, doc_number)).fetchall()
                            if len(_by_number) == 1:
                                existing_doc = _by_number[0]
                                logger.info(
                                    f"[update_existing] document type mismatch, "
                                    f"fallback by number: doc id={existing_doc['id']}, "
                                    f"request type={_doc_type!r}, "
                                    f"db type={dict(_by_number[0]).get('document_type')!r}"
                                )

                        if existing_doc:
                            doc_id = existing_doc['id']
                            conn.execute("""
                                UPDATE education_documents SET
                                    document_type_en    = COALESCE(NULLIF(?, ''), document_type_en),
                                    institution_name    = COALESCE(NULLIF(?, ''), institution_name),
                                    institution_name_en = COALESCE(NULLIF(?, ''), institution_name_en),
                                    country             = COALESCE(NULLIF(?, ''), country),
                                    country_en          = COALESCE(NULLIF(?, ''), country_en),
                                    completion_date     = COALESCE(NULLIF(?, ''), completion_date)
                                WHERE id = ?
                            """, (
                                translate_to_en(_doc_type) or None,
                                row['document_institution'] or None,
                                translate_to_en(row['document_institution'] or '') or None,
                                row['document_country'] or None,
                                translate_to_en(row['document_country'] or '') or None,
                                row['document_date'] or None,
                                doc_id,
                            ))
                            updated_parts.append(
                                f'документ про освіту (оновлено, група ID {_doc_gid})'
                            )
                            applied_changes.append({
                                'label': 'Документ про освіту',
                                'old': '',
                                'new': f'оновлено (група ID {_doc_gid}, тип: {_doc_type or "—"}, № {doc_number or "—"})',
                            })
                        else:
                            cur_doc = conn.execute("""
                                INSERT INTO education_documents (
                                    student_id, group_id, document_type, document_type_en,
                                    document_number, institution_name, institution_name_en,
                                    country, country_en, completion_date
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                existing_id, _doc_gid,
                                _doc_type, translate_to_en(_doc_type),
                                doc_number,
                                row['document_institution'] or '',
                                translate_to_en(row['document_institution'] or ''),
                                row['document_country'] or '',
                                translate_to_en(row['document_country'] or ''),
                                row['document_date'] or '',
                            ))
                            doc_id = cur_doc.lastrowid
                            updated_parts.append(
                                f'документ про освіту (додано, група ID {_doc_gid})'
                            )
                            applied_changes.append({
                                'label': 'Документ про освіту',
                                'old': '',
                                'new': f'додано (група ID {_doc_gid}, тип: {_doc_type or "—"}, № {doc_number or "—"})',
                            })
                        doc_ids_for_scans.append(doc_id)

                    if doc_ids_for_scans:
                        conn.execute(
                            "UPDATE attachments SET entity_type='education_document', entity_id=? "
                            "WHERE entity_type='pending_student' AND entity_id=?",
                            (doc_ids_for_scans[0], pending_id)
                        )
                        if len(doc_ids_for_scans) > 1:
                            import shutil as _shutil, uuid as _uuid_scan
                            src_atts = conn.execute(
                                "SELECT file_path, original_name, uploaded_by FROM attachments "
                                "WHERE entity_type='education_document' AND entity_id=?",
                                (doc_ids_for_scans[0],)
                            ).fetchall()
                            for extra_doc_id in doc_ids_for_scans[1:]:
                                for att in src_atts:
                                    src_rel = att['file_path']
                                    src_abs = os.path.join('static', src_rel)
                                    if not os.path.exists(src_abs):
                                        continue
                                    ext = os.path.splitext(src_rel)[1]
                                    dir_path = os.path.dirname(src_rel)
                                    new_rel = f"{dir_path}/{_uuid_scan.uuid4().hex}{ext}"
                                    try:
                                        _shutil.copy2(src_abs, os.path.join('static', new_rel))
                                        conn.execute(
                                            "INSERT INTO attachments "
                                            "(entity_type, entity_id, file_path, original_name, uploaded_by) "
                                            "VALUES ('education_document', ?, ?, ?, ?)",
                                            (extra_doc_id, new_rel, att['original_name'], att['uploaded_by'])
                                        )
                                    except Exception as _e:
                                        logger.warning(
                                            f"Не вдалося скопіювати скан для документа {extra_doc_id}: {_e}"
                                        )

                conn.execute(
                    "UPDATE attachments SET entity_type='student', entity_id=? "
                    "WHERE entity_type='pending_student' AND entity_id=?",
                    (existing_id, pending_id)
                )

                # ── Паспорт (dedup + підтримка кількох груп) ──
                passport_ids_for_scans = []
                if (request.form.get('add_passport') or request.form.get('update_passport')) and row['passport_number']:
                    _passport_doc_type = (
                        normalize_passport_type(row['passport_document_type'])
                        or PASSPORT_TYPE_DEFAULT
                    )
                    _passport_number = normalize_passport_number(
                        row['passport_number'], _passport_doc_type
                    )

                    target_group_ids = []
                    for gs in request.form.getlist('target_group_ids_passport'):
                        try:
                            target_group_ids.append(int(gs))
                        except (ValueError, TypeError):
                            pass
                    if not target_group_ids:
                        _grp_row = conn.execute(
                            "SELECT group_id FROM students WHERE id = ?", (existing_id,)
                        ).fetchone()
                        _gid = _grp_row['group_id'] if _grp_row else None
                        target_group_ids = [_gid] if _gid else [None]

                    for _doc_gid in target_group_ids:
                        existing_p = _find_existing_passport(
                            conn, existing_id, _doc_gid,
                            _passport_doc_type, row['passport_series'], _passport_number,
                        )

                        # Fallback по номеру (тип в заявці — синоним типу в БД)
                        if not existing_p and _passport_number:
                            _by_number = conn.execute("""
                                SELECT id FROM passport_documents
                                WHERE student_id = ?
                                  AND COALESCE(group_id, -1) = COALESCE(?, -1)
                                  AND NORM_PASSPORT(number, document_type)
                                      = NORM_PASSPORT(?, document_type)
                            """, (existing_id, _doc_gid, _passport_number)).fetchall()
                            if len(_by_number) == 1:
                                existing_p = _by_number[0]

                        if existing_p:
                            passport_id = existing_p['id']
                            conn.execute("""
                                UPDATE passport_documents SET
                                    issued_by     = COALESCE(NULLIF(issued_by, ''), ?),
                                    issue_date    = COALESCE(NULLIF(issue_date, ''), ?),
                                    valid_until   = COALESCE(NULLIF(valid_until, ''), ?),
                                    unique_number = COALESCE(NULLIF(unique_number, ''), ?)
                                WHERE id = ?
                            """, (
                                row['passport_issued_by'], row['passport_issue_date'],
                                row['passport_valid_until'], row['passport_unique_number'],
                                passport_id,
                            ))
                            updated_parts.append(
                                f'паспортні дані (заповнено порожні поля, група ID {_doc_gid})'
                            )
                            applied_changes.append({
                                'label': 'Паспорт',
                                'old': '',
                                'new': f'заповнено порожні поля (група ID {_doc_gid})',
                            })
                        else:
                            cur_p = conn.execute("""
                                INSERT INTO passport_documents (
                                    student_id, group_id, document_type, series, number,
                                    issued_by, issue_date, valid_until, unique_number
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                existing_id, _doc_gid, _passport_doc_type,
                                row['passport_series'], _passport_number,
                                row['passport_issued_by'], row['passport_issue_date'],
                                row['passport_valid_until'], row['passport_unique_number'],
                            ))
                            passport_id = cur_p.lastrowid
                            updated_parts.append(
                                f'паспортні дані (додано, група ID {_doc_gid})'
                            )
                            applied_changes.append({
                                'label': 'Паспорт',
                                'old': '',
                                'new': f'додано (група ID {_doc_gid}, тип: {_passport_doc_type}, № {_passport_number})',
                            })
                        passport_ids_for_scans.append(passport_id)

                    if passport_ids_for_scans:
                        conn.execute(
                            "UPDATE attachments SET entity_type='passport_document', entity_id=? "
                            "WHERE entity_type='pending_student_passport' AND entity_id=?",
                            (passport_ids_for_scans[0], pending_id)
                        )
                        if len(passport_ids_for_scans) > 1:
                            import shutil as _shutil, uuid as _uuid_scan
                            src_atts = conn.execute(
                                "SELECT file_path, original_name, uploaded_by FROM attachments "
                                "WHERE entity_type='passport_document' AND entity_id=?",
                                (passport_ids_for_scans[0],)
                            ).fetchall()
                            for extra_id in passport_ids_for_scans[1:]:
                                for att in src_atts:
                                    src_rel = att['file_path']
                                    src_abs = os.path.join('static', src_rel)
                                    if not os.path.exists(src_abs):
                                        continue
                                    ext = os.path.splitext(src_rel)[1]
                                    dir_path = os.path.dirname(src_rel)
                                    new_rel = f"{dir_path}/{_uuid_scan.uuid4().hex}{ext}"
                                    try:
                                        _shutil.copy2(src_abs, os.path.join('static', new_rel))
                                        conn.execute(
                                            "INSERT INTO attachments "
                                            "(entity_type, entity_id, file_path, original_name, uploaded_by) "
                                            "VALUES ('passport_document', ?, ?, ?, ?)",
                                            (extra_id, new_rel, att['original_name'], att['uploaded_by'])
                                        )
                                    except Exception as _e:
                                        logger.warning(f"Не вдалося скопіювати скан паспорта: {_e}")

                conn.execute(
                    "UPDATE attachments SET entity_type='student', entity_id=? "
                    "WHERE entity_type='pending_student_passport' AND entity_id=?",
                    (existing_id, pending_id)
                )

                # ── Військові дані ──
                if request.form.get('add_military') and _pending_has_military_data(row):
                    military_data = _military_payload_from_pending(row)
                    existing_military = conn.execute(
                        "SELECT id FROM military WHERE student_id=?", (existing_id,)
                    ).fetchone()

                    if existing_military:
                        string_keys = [
                            k for k in military_data if k != 'data_verified_on_time'
                        ]
                        set_clause = ', '.join(
                            f"{k} = COALESCE(NULLIF(?, ''), {k})" for k in string_keys
                        )
                        conn.execute(f"""
                            UPDATE military SET
                                {set_clause},
                                data_verified_on_time = CASE
                                    WHEN COALESCE(data_verified_on_time, 0) = 0 THEN ?
                                    ELSE data_verified_on_time
                                END
                            WHERE id = ?
                        """, [military_data[k] for k in string_keys]
                              + [military_data['data_verified_on_time'], existing_military['id']])
                        military_id_for_scans = existing_military['id']
                    else:
                        cols = list(military_data.keys())
                        placeholders = ', '.join('?' for _ in cols)
                        cur_military = conn.execute(
                            f"INSERT INTO military (student_id, {', '.join(cols)}) "
                            f"VALUES (?, {placeholders})",
                            [existing_id] + list(military_data.values())
                        )
                        military_id_for_scans = cur_military.lastrowid

                    if row['military_address']:
                        conn.execute(
                            "UPDATE students SET address_of_residence = "
                            "COALESCE(NULLIF(address_of_residence, ''), ?) WHERE id = ?",
                            (row['military_address'], existing_id),
                        )
                    updated_parts.append('військовий облік')
                    applied_changes.append({
                        'label': 'Військові дані',
                        'old': '',
                        'new': 'додано/оновлено з заявки',
                    })
                    conn.execute(
                        "UPDATE attachments SET entity_type='military', entity_id=? "
                        "WHERE entity_type='pending_student_military' AND entity_id=?",
                        (military_id_for_scans, pending_id)
                    )
                # ── Скани військового обліку, що залишились неприкріпленими ──
                has_pending_mil_scans = conn.execute(
                    "SELECT 1 FROM attachments "
                    "WHERE entity_type='pending_student_military' AND entity_id=? LIMIT 1",
                    (pending_id,),
                ).fetchone() is not None

                if has_pending_mil_scans:
                    mil_row = conn.execute(
                        "SELECT id FROM military WHERE student_id = ?", (existing_id,)
                    ).fetchone()
                    if mil_row:
                        target_mil_id = mil_row['id']
                    else:
                        target_mil_id = conn.execute(
                            "INSERT INTO military (student_id) VALUES (?)", (existing_id,)
                        ).lastrowid
                    conn.execute(
                        "UPDATE attachments SET entity_type='military', entity_id=? "
                        "WHERE entity_type='pending_student_military' AND entity_id=?",
                        (target_mil_id, pending_id),
                    )
                    updated_parts.append('скани військового обліку')
                    applied_changes.append({
                        'label': 'Скани військового обліку',
                        'old': '',
                        'new': 'перенесено в картку студента',
                    })

                if not updated_parts:
                    conn.rollback()
                    conn.close()
                    flash("Не обрано жодного пункту для оновлення", "error")
                    return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

                conn.execute(
                    "UPDATE pending_students SET status='approved', reviewed_by=?, "
                    "reviewed_at=datetime('now','localtime'), resulting_student_id=?, "
                    "applied_changes=? WHERE id=?",
                    (
                        current_username(),
                        existing_id,
                        json.dumps(applied_changes, ensure_ascii=False) if applied_changes else None,
                        pending_id,
                    ),
                )
                conn.commit()

            except sqlite3.IntegrityError as e:
                conn.rollback()
                logger.error(
                    f"[update_existing] pending_id={pending_id}, existing_id={existing_id} "
                    f"integrity: {e}", exc_info=True,
                )
                flash(
                    f"Помилка: дані з заявки конфліктують із наявними в базі ({e}). "
                    f"Можливо, паспорт/документ уже додано вручну.",
                    "error",
                )
                conn.close()
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
            except Exception as e:
                conn.rollback()
                logger.error(
                    f"[update_existing] pending_id={pending_id}, existing_id={existing_id}: {e}",
                    exc_info=True,
                )
                flash(f"Помилка оновлення студента: {e}", "error")
                conn.close()
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

            log_action(
                current_username(),
                f"оновив наявного студента (ID {existing_id}) даними із заявки на реєстрацію {pending_id}",
                details=', '.join(updated_parts)
            )
            conn.close()
            flash(f"Оновлено: {', '.join(updated_parts)}", "success")
            return redirect(url_for('students.student_details', student_id=existing_id))

        conn.close()
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    groups = conn.execute("""
        SELECT id, name, start_year, study_form, program_credits,
               name || ' (' || start_year || ', ' || study_form || ', ' || program_credits || ' кредитів)' AS display_name
        FROM groups WHERE archived = FALSE ORDER BY name COLLATE UKRAINIAN
    """).fetchall()
    licenses = conn.execute(
        "SELECT id, short_name_ua, name_ua FROM institution_licenses WHERE is_active=1 ORDER BY id"
    ).fetchall()

    existing_student = conn.execute("""
        SELECT id, last_name_UA, first_name_UA FROM students
        WHERE LOWER_UA(last_name_UA) = LOWER_UA(?) AND LOWER_UA(first_name_UA) = LOWER_UA(?) AND birth_date = ?
    """, (row['last_name_UA'], row['first_name_UA'], row['birth_date'])).fetchone()

    scans = get_attachments(conn, 'pending_student', pending_id)
    passport_scans = get_attachments(conn, 'pending_student_passport', pending_id)
    military_scans = get_attachments(conn, 'pending_student_military', pending_id)

    # ── Ініціалізація змінних для шаблону (щоб Jinja не падала при existing_student=None) ──
    existing_full = None
    existing_enrollments = []
    existing_has_passport = False
    existing_passport = None
    existing_edu_docs_count = 0
    existing_edu_docs = []
    existing_matching_edu_doc = None
    existing_military = None
    active_enrolls = []
    education_docs_by_group = {}
    legacy_edu_docs = []
    passports_by_group = {}
    legacy_passport = None

    if existing_student:
        sid = existing_student['id']

        ex_row = conn.execute("""
            SELECT s.*,
                   g.name AS primary_group_name,
                   il.short_name_ua AS primary_license_short
            FROM students s
            LEFT JOIN groups g ON s.group_id = g.id
            LEFT JOIN institution_licenses il ON il.id = s.license_id
            WHERE s.id = ?
        """, (sid,)).fetchone()
        existing_full = dict(ex_row) if ex_row else None

        try:
            from routes.enrollments_service import list_student_enrollments
            existing_enrollments = list_student_enrollments(sid, conn=conn)
        except Exception as _e:
            logger.warning(
                f"pending_student_review: не вдалося завантажити "
                f"enrollments для студента {sid}: {_e}"
            )

        existing_passport = conn.execute(
            "SELECT * FROM passport_documents WHERE student_id = ? ORDER BY id DESC LIMIT 1",
            (sid,),
        ).fetchone()
        existing_has_passport = existing_passport is not None

        existing_edu_docs = [
            dict(r) for r in conn.execute(
                "SELECT * FROM education_documents WHERE student_id = ? ORDER BY id DESC",
                (sid,),
            ).fetchall()
        ]
        existing_edu_docs_count = len(existing_edu_docs)

        def _norm_doc_num(s):
            return ''.join(ch for ch in (s or '').upper() if ch.isalnum())

        app_doc_num = _norm_doc_num(
            ' '.join(x for x in [row['document_series'], row['document_number']] if x)
            or (row['document_number'] or '')
        )
        existing_matching_edu_doc = None
        if app_doc_num:
            for ed in existing_edu_docs:
                if _norm_doc_num(ed.get('document_number')) == app_doc_num:
                    existing_matching_edu_doc = ed
                    break

        existing_military = conn.execute(
            "SELECT * FROM military WHERE student_id = ? LIMIT 1",
            (sid,),
        ).fetchone()

        # ═══════════════════════════════════════════════════════════════
        #  Дані для блоків «Що вже є в базі» — усередині if existing_student:
        # ═══════════════════════════════════════════════════════════════

        # 1) Активні enrollments — ПЕРШИМ ділом (до циклів нижче)
        active_enrolls = [e for e in existing_enrollments if e.get('status') == 'active']

        # Legacy-fallback: студент без enrollment'ів, але з students.group_id
        if not active_enrolls and existing_full and existing_full.get('group_id'):
            _grp = conn.execute(
                "SELECT id, name FROM groups WHERE id = ?",
                (existing_full['group_id'],),
            ).fetchone()
            if _grp:
                active_enrolls = [{
                    'group_id': _grp['id'],
                    'group_name': _grp['name'],
                    'is_primary': True,
                    'status': 'active',
                }]

        # 2) Документи про освіту по кожній активній групі
        education_docs_by_group = {}
        for e in active_enrolls:
            docs = conn.execute("""
                SELECT id, document_type, document_type_en, document_number,
                       institution_name, country, completion_date
                FROM education_documents
                WHERE student_id = ? AND group_id = ?
                ORDER BY id DESC
            """, (sid, e['group_id'])).fetchall()
            if docs:
                lst = []
                for d in docs:
                    d = dict(d)
                    d['_matches_request'] = bool(
                        app_doc_num
                        and _norm_doc_num(d.get('document_number')) == app_doc_num
                    )
                    lst.append(d)
                education_docs_by_group[e['group_id']] = lst

        # 3) Legacy-документи про освіту (без прив'язки до групи)
        for d in conn.execute("""
            SELECT id, document_type, document_number, institution_name,
                   country, completion_date
            FROM education_documents
            WHERE student_id = ? AND group_id IS NULL
            ORDER BY id DESC
        """, (sid,)).fetchall():
            d = dict(d)
            d['_matches_request'] = bool(
                app_doc_num
                and _norm_doc_num(d.get('document_number')) == app_doc_num
            )
            legacy_edu_docs.append(d)

        # 4) Паспорти по кожній активній групі
        for e in active_enrolls:
            p = conn.execute("""
                SELECT id, document_type, series, number, issued_by, issue_date,
                       valid_until, unique_number
                FROM passport_documents
                WHERE student_id = ? AND COALESCE(group_id, -1) = ?
                ORDER BY id DESC LIMIT 1
            """, (sid, e['group_id'])).fetchone()
            if p:
                passports_by_group[e['group_id']] = dict(p)

        # 5) Legacy-паспорт (без прив'язки до групи)
        p_legacy = conn.execute("""
            SELECT id, document_type, series, number, issued_by, issue_date,
                   valid_until, unique_number
            FROM passport_documents
            WHERE student_id = ? AND group_id IS NULL
            ORDER BY id DESC LIMIT 1
        """, (sid,)).fetchone()
        if p_legacy:
            legacy_passport = dict(p_legacy)
            passports_by_group[None] = legacy_passport

    # ── Застосовані зміни (для підтверджених/відхилених заявок) ──
    applied_changes = _load_applied_changes(row)
    manual_edits = _load_manual_edits(row)

    conn.close()
    return render_template(
        'admin_pending_student_review.html',
        row=row, groups=groups, licenses=licenses,
        existing_student=existing_student,
        manual_edits=manual_edits,
        existing_full=existing_full,
        existing_enrollments=existing_enrollments,
        existing_has_passport=existing_has_passport,
        existing_passport=existing_passport,
        existing_edu_docs_count=existing_edu_docs_count,
        existing_edu_docs=existing_edu_docs,
        existing_matching_edu_doc=existing_matching_edu_doc,
        existing_military=existing_military,
        active_enrolls=active_enrolls,
        education_docs_by_group=education_docs_by_group,
        legacy_edu_docs=legacy_edu_docs,
        passports_by_group=passports_by_group,
        legacy_passport=legacy_passport,
        scans=scans, passport_scans=passport_scans, military_scans=military_scans,
        applied_changes=applied_changes,
        **passport_template_context(),
    )


@admin_bp.route('/admin/pending_students/<int:pending_id>/photo', methods=['POST'])
@permission_required('manage_pending_students')
def pending_student_photo(pending_id):
    """Ручна (пере)обрізка фото абітурієнта."""
    conn = get_db()
    row = conn.execute("SELECT photo_path FROM pending_students WHERE id=?", (pending_id,)).fetchone()
    if not row:
        conn.close()
        flash("Заявку не знайдено", "error")
        return redirect(url_for('admin.pending_students'))

    file = request.files.get('photo_file')
    if file and file.filename:
        file_bytes = file.read()
    else:
        if not row['photo_path']:
            conn.close()
            flash("Немає наявного фото для переобрізки - завантажте файл", "error")
            return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
        src_path = os.path.join('static', row['photo_path'])
        if not os.path.exists(src_path):
            conn.close()
            flash("Файл наявного фото не знайдено на диску", "error")
            return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
        with open(src_path, 'rb') as f:
            file_bytes = f.read()

    try:
        crop_box = (
            float(request.form['crop_x']),
            float(request.form['crop_y']),
            float(request.form['crop_w']),
            float(request.form['crop_h']),
        )
    except (KeyError, ValueError):
        conn.close()
        flash("Некоректні дані обрізки фото - спробуйте ще раз", "error")
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    from routes.photo import process_and_save_pending_photo_with_crop
    try:
        new_path = process_and_save_pending_photo_with_crop(
            file_bytes, crop_box, old_rel_path=row['photo_path']
        )
    except ValueError as e:
        conn.close()
        flash(str(e), "error")
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    conn.execute("UPDATE pending_students SET photo_path=? WHERE id=?", (new_path, pending_id))
    conn.commit()
    conn.close()
    flash("Фото оновлено", "success")
    return redirect(url_for('admin.pending_student_review', pending_id=pending_id))


# ═══════════════════════════════════════════════════════════════════════
# Заявки на оновлення даних (/update-info)
# ═══════════════════════════════════════════════════════════════════════

@admin_bp.route('/admin/update_requests')
@permission_required('manage_update_requests')
def update_requests():
    """Список заявок на оновлення даних наявного студента."""
    conn = get_db()
    conn.row_factory = sqlite3.Row

    status_filter = request.args.get('status', 'submitted')
    if status_filter not in ('pending', 'submitted', 'approved', 'rejected'):
        status_filter = 'submitted'

    rows = conn.execute("""
        SELECT ur.*, s.last_name_UA, s.first_name_UA
        FROM update_requests ur
        JOIN students s ON s.id = ur.student_id
        WHERE ur.status = ?
        ORDER BY ur.created_at DESC
    """, (status_filter,)).fetchall()

    counts = {
        s: conn.execute("SELECT COUNT(*) FROM update_requests WHERE status=?", (s,)).fetchone()[0]
        for s in ('pending', 'submitted', 'approved', 'rejected')
    }

    conn.close()
    return render_template(
        'admin_update_requests.html',
        rows=rows, status_filter=status_filter, counts=counts
    )


@admin_bp.route('/admin/update_requests/<int:request_id>/photo', methods=['POST'])
@permission_required('manage_update_requests')
def update_request_photo(request_id):
    """Ручна (пере)обрізка фото, надісланого через /update-info."""
    conn = get_db()
    row = conn.execute("SELECT photo_path FROM update_requests WHERE id=?", (request_id,)).fetchone()
    if not row:
        conn.close()
        flash("Заявку не знайдено", "error")
        return redirect(url_for('admin.update_requests'))

    file = request.files.get('photo_file')
    if file and file.filename:
        file_bytes = file.read()
    else:
        if not row['photo_path']:
            conn.close()
            flash("Немає наявного фото для переобрізки - завантажте файл", "error")
            return redirect(url_for('admin.update_request_review', request_id=request_id))
        src_path = os.path.join('static', row['photo_path'])
        if not os.path.exists(src_path):
            conn.close()
            flash("Файл наявного фото не знайдено на диску", "error")
            return redirect(url_for('admin.update_request_review', request_id=request_id))
        with open(src_path, 'rb') as f:
            file_bytes = f.read()

    try:
        crop_box = (
            float(request.form['crop_x']),
            float(request.form['crop_y']),
            float(request.form['crop_w']),
            float(request.form['crop_h']),
        )
    except (KeyError, ValueError):
        conn.close()
        flash("Некоректні дані обрізки фото - спробуйте ще раз", "error")
        return redirect(url_for('admin.update_request_review', request_id=request_id))

    from routes.photo import process_and_save_pending_photo_with_crop
    try:
        new_path = process_and_save_pending_photo_with_crop(
            file_bytes, crop_box, old_rel_path=row['photo_path']
        )
    except ValueError as e:
        conn.close()
        flash(str(e), "error")
        return redirect(url_for('admin.update_request_review', request_id=request_id))

    conn.execute("UPDATE update_requests SET photo_path=? WHERE id=?", (new_path, request_id))
    conn.commit()
    conn.close()
    flash("Фото оновлено", "success")
    return redirect(url_for('admin.update_request_review', request_id=request_id))


@admin_bp.route('/admin/update_requests/<int:request_id>', methods=['GET', 'POST'])
@permission_required('manage_update_requests')
def update_request_review(request_id):
    """Перегляд однієї заявки на оновлення."""
    conn = get_db()
    conn.row_factory = sqlite3.Row
    row = conn.execute("""
        SELECT ur.*, s.last_name_UA, s.first_name_UA, s.middle_name_UA
        FROM update_requests ur JOIN students s ON s.id = ur.student_id
        WHERE ur.id = ?
    """, (request_id,)).fetchone()
    if not row:
        conn.close()
        flash("Заявку не знайдено", "error")
        return redirect(url_for('admin.update_requests'))

    # ── ВИПРАВЛЕНО: student_id визначаємо ДО POST-гілки (для GET-запиту теж потрібен) ──
    student_id = row['student_id']

    allowed_fields = json.loads(row['allowed_fields'])
    submitted = json.loads(row['submitted_data']) if row['submitted_data'] else {}

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'delete':
            for entity_type in ('update_request_passport', 'update_request_education', 'update_request_military'):
                for att in get_attachments(conn, entity_type, request_id):
                    still_used = conn.execute(
                        "SELECT 1 FROM attachments WHERE file_path=? AND id!=? LIMIT 1",
                        (att['file_path'], att['id'])
                    ).fetchone()
                    att_path = os.path.join('static', att['file_path'])
                    if not still_used and os.path.exists(att_path):
                        os.remove(att_path)
                conn.execute(
                    "DELETE FROM attachments WHERE entity_type=? AND entity_id=?",
                    (entity_type, request_id)
                )
            if row['photo_path']:
                photo_abs = os.path.join('static', row['photo_path'])
                if os.path.exists(photo_abs):
                    os.remove(photo_abs)
            conn.execute("DELETE FROM update_requests WHERE id=?", (request_id,))
            conn.commit()
            log_action(current_username(), f"видалив заявку на оновлення даних ID {request_id}")
            conn.close()
            flash("Заявку видалено", "success")
            return redirect(url_for('admin.update_requests'))

        elif action == 'reject':
            note = (request.form.get('review_note') or '').strip() or None
            conn.execute(
                "UPDATE update_requests SET status='rejected', reviewed_by=?, "
                "reviewed_at=datetime('now','localtime'), review_note=? WHERE id=?",
                (current_username(), note, request_id)
            )
            conn.commit()
            log_action(
                current_username(),
                f"відхилив заявку на оновлення даних ID {request_id}",
                details=note or ''
            )
            conn.close()
            flash("Заявку відхилено", "success")
            return redirect(url_for('admin.update_requests'))

        elif action == 'apply':
            applied_parts = []
            applied_field_keys = []

            try:
                simple_updates = {}
                for key in ('last_name_UA', 'first_name_UA', 'middle_name_UA', 'phone',
                            'phone_backup', 'email', 'tax_id', 'edebo_code'):
                    if key in allowed_fields and request.form.get(f'apply_{key}'):
                        simple_updates[key] = submitted.get(key)

                if 'last_name_UA' in simple_updates or 'first_name_UA' in simple_updates:
                    new_last_ua = simple_updates.get('last_name_UA', row['last_name_UA'])
                    new_first_ua = simple_updates.get('first_name_UA', row['first_name_UA'])
                    last_eng, first_eng = generate_english_name(new_last_ua, new_first_ua)
                    simple_updates['last_name_ENG'] = last_eng
                    simple_updates['first_name_ENG'] = first_eng

                if simple_updates:
                    set_clause = ", ".join(f"{k}=?" for k in simple_updates)
                    conn.execute(
                        f"UPDATE students SET {set_clause} WHERE id=?",
                        list(simple_updates.values()) + [student_id]
                    )
                    if 'edebo_code' in simple_updates:
                        _primary = conn.execute(
                            "SELECT group_id FROM students WHERE id = ?", (student_id,)
                        ).fetchone()
                        if _primary and _primary['group_id']:
                            conn.execute(
                                "UPDATE student_enrollments SET edebo_code = ? "
                                "WHERE student_id = ? AND group_id = ? AND is_primary = 1",
                                (simple_updates['edebo_code'], student_id, _primary['group_id']),
                            )
                    applied_parts.append(", ".join(simple_updates.keys()))
                    applied_field_keys.extend(simple_updates.keys())

                if 'photo' in allowed_fields and row['photo_path'] and request.form.get('apply_photo'):
                    from routes.photo import photo_path_for_student, _ensure_dir
                    _ensure_dir()
                    src_path = os.path.join('static', row['photo_path'])
                    if os.path.exists(src_path):
                        dest_path = photo_path_for_student(student_id)
                        with open(src_path, 'rb') as src, open(dest_path, 'wb') as dst:
                            dst.write(src.read())
                        conn.execute(
                            "UPDATE students SET photo=? WHERE id=?",
                            (f"uploads/photos/student_{student_id}.jpg", student_id)
                        )
                        applied_parts.append('фото')
                        applied_field_keys.append('photo')

                # ── Паспорт (dedup) ──
                if 'passport' in allowed_fields and submitted.get('passport') and request.form.get('apply_passport'):
                    p = submitted['passport']

                    _target_gid = request.form.get('passport_target_group_id', type=int)
                    if not _target_gid:
                        _grp_row = conn.execute(
                            "SELECT group_id FROM students WHERE id = ?", (student_id,)
                        ).fetchone()
                        _target_gid = _grp_row['group_id'] if _grp_row else None

                    if _target_gid:
                        _valid = conn.execute("""
                            SELECT 1 FROM student_enrollments
                            WHERE student_id = ? AND group_id = ? AND status = 'active'
                            UNION
                            SELECT 1 FROM students WHERE id = ? AND group_id = ?
                        """, (student_id, _target_gid, student_id, _target_gid)).fetchone()
                        if not _valid:
                            _target_gid = None

                    _doc_gid = _target_gid

                    _p_doc_type = (
                        normalize_passport_type(p.get('document_type'))
                        or PASSPORT_TYPE_DEFAULT
                    )
                    _p_number = normalize_passport_number(
                        p.get('number'), _p_doc_type
                    )

                    existing_p = _find_existing_passport(
                        conn, student_id, _doc_gid,
                        _p_doc_type, p.get('series'), _p_number,
                    )

                    if existing_p:
                        conn.execute("""
                            UPDATE passport_documents SET
                                issued_by     = COALESCE(NULLIF(issued_by, ''), ?),
                                issue_date    = COALESCE(NULLIF(issue_date, ''), ?),
                                valid_until   = COALESCE(NULLIF(valid_until, ''), ?),
                                unique_number = COALESCE(NULLIF(unique_number, ''), ?)
                            WHERE id = ?
                        """, (
                            p.get('issued_by'), p.get('issue_date'),
                            p.get('valid_until'), p.get('unique_number'),
                            existing_p['id'],
                        ))
                        new_passport_id = existing_p['id']
                    else:
                        cur_p = conn.execute("""
                            INSERT INTO passport_documents (
                                student_id, group_id, document_type, series, number,
                                issued_by, issue_date, valid_until, unique_number
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            student_id, _doc_gid,
                            _p_doc_type,
                            p.get('series'), _p_number,
                            p.get('issued_by'), p.get('issue_date'),
                            p.get('valid_until'), p.get('unique_number'),
                        ))
                        new_passport_id = cur_p.lastrowid

                    conn.execute("""
                        INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_at, uploaded_by)
                        SELECT 'passport_document', ?, file_path, original_name, uploaded_at, uploaded_by
                        FROM attachments WHERE entity_type='update_request_passport' AND entity_id=?
                    """, (new_passport_id, request_id))
                    applied_parts.append('паспортні дані')
                    applied_field_keys.append('passport')

                d = submitted.get('education_document') or {}
                education_doc_id = None

                if 'education_document' in allowed_fields and d.get('document_type') and request.form.get('apply_education_document'):
                    document_number = ((d.get('series') + ' ') if d.get('series') else '') + (d.get('number') or '')
                    _grp_row = conn.execute(
                        "SELECT group_id FROM students WHERE id = ?", (student_id,)
                    ).fetchone()
                    _doc_gid = _grp_row['group_id'] if _grp_row else None

                    existing_d = _find_existing_education_doc(
                        conn, student_id, _doc_gid,
                        d.get('document_type') or '', document_number,
                    )
                    if not existing_d and document_number:
                        _by_number = conn.execute("""
                            SELECT id FROM education_documents
                            WHERE student_id = ?
                              AND COALESCE(group_id, -1) = COALESCE(?, -1)
                              AND NORM_DOC_NUM(document_number) = NORM_DOC_NUM(?)
                        """, (student_id, _doc_gid, document_number)).fetchall()
                        if len(_by_number) == 1:
                            existing_d = _by_number[0]

                    if existing_d:
                        education_doc_id = existing_d['id']
                        conn.execute("""
                            UPDATE education_documents SET
                                document_type_en    = COALESCE(NULLIF(?, ''), document_type_en),
                                institution_name    = COALESCE(NULLIF(?, ''), institution_name),
                                institution_name_en = COALESCE(NULLIF(?, ''), institution_name_en),
                                country             = COALESCE(NULLIF(?, ''), country),
                                country_en          = COALESCE(NULLIF(?, ''), country_en),
                                completion_date     = COALESCE(NULLIF(?, ''), completion_date)
                            WHERE id = ?
                        """, (
                            translate_to_en(d.get('document_type') or '') or None,
                            d.get('institution') or None,
                            translate_to_en(d.get('institution') or '') or None,
                            d.get('country') or None,
                            translate_to_en(d.get('country') or '') or None,
                            d.get('completion_date') or None,
                            education_doc_id,
                        ))
                    else:
                        cur_d = conn.execute("""
                            INSERT INTO education_documents (
                                student_id, group_id, document_type, document_type_en, document_number,
                                institution_name, institution_name_en, country, country_en, completion_date
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            student_id, _doc_gid,
                            d.get('document_type') or '',
                            translate_to_en(d.get('document_type') or ''),
                            document_number,
                            d.get('institution') or '',
                            translate_to_en(d.get('institution') or ''),
                            d.get('country') or '',
                            translate_to_en(d.get('country') or ''),
                            d.get('completion_date') or '',
                        ))
                        education_doc_id = cur_d.lastrowid

                    conn.execute("""
                        INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_at, uploaded_by)
                        SELECT 'education_document', ?, file_path, original_name, uploaded_at, uploaded_by
                        FROM attachments WHERE entity_type='update_request_education' AND entity_id=?
                    """, (education_doc_id, request_id))
                    applied_parts.append('документ про освіту')
                    applied_field_keys.append('education_document')

                reference = d.get('reference') or {}
                recognition = d.get('recognition') or {}
                apply_reference = (
                    'education_document_reference' in allowed_fields and reference
                    and request.form.get('apply_education_reference')
                )
                apply_recognition = (
                    'education_document_recognition' in allowed_fields and recognition
                    and request.form.get('apply_education_recognition')
                )

                if apply_reference or apply_recognition:
                    if education_doc_id is None:
                        existing_doc = conn.execute(
                            "SELECT id FROM education_documents WHERE student_id=? ORDER BY id DESC LIMIT 1",
                            (student_id,)
                        ).fetchone()
                        education_doc_id = existing_doc[0] if existing_doc else None

                    if education_doc_id is None:
                        flash("Немає жодного документа про освіту, до якого можна додати довідку/свідоцтво.", "error")
                    else:
                        existing_foreign = conn.execute(
                            "SELECT id FROM foreign_education_docs WHERE education_doc_id=?",
                            (education_doc_id,)
                        ).fetchone()

                        ref_number = reference.get('number') if apply_reference else None
                        ref_institution = reference.get('institution') if apply_reference else None
                        ref_institution_en = translate_to_en(ref_institution) if ref_institution else None
                        ref_country = reference.get('country') if apply_reference else None
                        ref_country_en = translate_to_en(ref_country) if ref_country else None
                        ref_issue_date = reference.get('issue_date') if apply_reference else None
                        rec_number = recognition.get('certificate_number') if apply_recognition else None
                        rec_issuer = recognition.get('issuer') if apply_recognition else None
                        rec_issuer_en = translate_to_en(rec_issuer) if rec_issuer else None
                        rec_date = recognition.get('date') if apply_recognition else None

                        if existing_foreign:
                            conn.execute("""
                                UPDATE foreign_education_docs SET
                                    reference_number               = COALESCE(?, reference_number),
                                    reference_institution          = COALESCE(?, reference_institution),
                                    reference_institution_en       = COALESCE(?, reference_institution_en),
                                    reference_country              = COALESCE(?, reference_country),
                                    reference_country_en           = COALESCE(?, reference_country_en),
                                    reference_issue_date           = COALESCE(?, reference_issue_date),
                                    recognition_certificate_number = COALESCE(?, recognition_certificate_number),
                                    recognition_issuer             = COALESCE(?, recognition_issuer),
                                    recognition_issuer_en          = COALESCE(?, recognition_issuer_en),
                                    recognition_date               = COALESCE(?, recognition_date)
                                WHERE education_doc_id=?
                            """, (
                                ref_number, ref_institution, ref_institution_en, ref_country, ref_country_en,
                                ref_issue_date, rec_number, rec_issuer, rec_issuer_en, rec_date,
                                education_doc_id
                            ))
                        else:
                            conn.execute("""
                                INSERT INTO foreign_education_docs (
                                    education_doc_id, reference_number, reference_institution, reference_institution_en,
                                    reference_country, reference_country_en, reference_issue_date,
                                    recognition_certificate_number, recognition_issuer, recognition_issuer_en, recognition_date
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                education_doc_id, ref_number, ref_institution, ref_institution_en,
                                ref_country, ref_country_en, ref_issue_date,
                                rec_number, rec_issuer, rec_issuer_en, rec_date
                            ))

                        if apply_reference:
                            applied_parts.append('академічна довідка')
                            applied_field_keys.append('education_document_reference')
                        if apply_recognition:
                            applied_parts.append('свідоцтво про визнання')
                            applied_field_keys.append('education_document_recognition')

                if 'military' in allowed_fields and request.form.get('apply_military'):
                    promote_ok, promote_msg = _promote_update_request_military_to_student(
                        conn, request_id, student_id, current_username()
                    )
                    if not promote_ok:
                        logger.info(f"[apply_update] fallback на форму: {promote_msg}")
                        m = submitted.get('military') or {}
                        existing_military = conn.execute(
                            "SELECT id FROM military WHERE student_id=?", (student_id,)
                        ).fetchone()

                        form_fields = {
                            'status_badge': m.get('status_badge'),
                            'valid_until': m.get('valid_until'),
                            'registration_category': m.get('registration_category'),
                            'removal_basis': m.get('removal_basis'),
                            'tck_name': m.get('tck_name'),
                            'military_rank': m.get('military_rank'),
                            'registry_number': m.get('registry_number'),
                            'vos_code': m.get('vos_code'),
                            'vos_description': m.get('vos_description'),
                            'deferral_type': m.get('deferral_type'),
                            'deferral_until': m.get('deferral_until'),
                            'police_reason': m.get('police_reason'),
                            'police_date': m.get('police_date'),
                            'vlk_decision': m.get('vlk_decision'),
                            'vlk_date': m.get('vlk_date'),
                            'disability_group': m.get('disability_group'),
                            'disability_valid_until': m.get('disability_valid_until'),
                            'disability_reason': m.get('disability_reason'),
                            'formed_at': m.get('formed_at'),
                            'military_registration_document': m.get('military_registration_document'),
                        }
                        updates = {k: v for k, v in form_fields.items() if v}

                        if m.get('address_of_residence'):
                            conn.execute(
                                "UPDATE students SET address_of_residence = COALESCE(NULLIF(address_of_residence, ''), ?) WHERE id = ?",
                                (m.get('address_of_residence'), student_id),
                            )

                        if existing_military:
                            military_id = existing_military['id']
                            if updates:
                                set_clause = ', '.join(
                                    f"{k} = COALESCE(NULLIF(?, ''), {k})" for k in updates
                                )
                                conn.execute(
                                    f"UPDATE military SET {set_clause} WHERE id=?",
                                    list(updates.values()) + [military_id]
                                )
                        else:
                            cols = ['student_id'] + list(updates.keys())
                            placeholders = ', '.join('?' for _ in cols)
                            cur_m = conn.execute(
                                f"INSERT INTO military ({', '.join(cols)}) VALUES ({placeholders})",
                                [student_id] + list(updates.values())
                            )
                            military_id = cur_m.lastrowid

                        conn.execute("""
                            INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_at, uploaded_by)
                            SELECT 'military', ?, file_path, original_name, uploaded_at, uploaded_by
                            FROM attachments WHERE entity_type='update_request_military' AND entity_id=?
                        """, (military_id, request_id))
                        applied_parts.append('військові дані (з форми)')
                    else:
                        applied_parts.append('військові дані (з PDF)')
                    applied_field_keys.append('military')

                if not applied_parts:
                    conn.rollback()
                    conn.close()
                    flash("Не обрано жодного пункту для застосування", "error")
                    return redirect(url_for('admin.update_request_review', request_id=request_id))

                conn.execute(
                    "UPDATE update_requests SET status='approved', reviewed_by=?, "
                    "reviewed_at=datetime('now','localtime'), applied_fields=? WHERE id=?",
                    (current_username(), json.dumps(applied_field_keys), request_id)
                )
                conn.commit()

            except sqlite3.IntegrityError as e:
                conn.rollback()
                logger.error(
                    f"[update_request apply] request_id={request_id}, student_id={student_id} "
                    f"integrity: {e}", exc_info=True,
                )
                flash(
                    f"Помилка: дані з заявки конфліктують із наявними в базі ({e}). "
                    f"Можливо, паспорт/документ уже додано вручну.",
                    "error",
                )
                conn.close()
                return redirect(url_for('admin.update_request_review', request_id=request_id))
            except Exception as e:
                conn.rollback()
                logger.error(
                    f"[update_request apply] request_id={request_id}: {e}",
                    exc_info=True,
                )
                flash(f"Помилка застосування заявки: {e}", "error")
                conn.close()
                return redirect(url_for('admin.update_request_review', request_id=request_id))

            log_action(
                current_username(),
                f"застосував заявку на оновлення даних: {row['last_name_UA']} {row['first_name_UA']} (ID {student_id})",
                details=", ".join(applied_parts)
            )
            conn.close()
            flash(f"Застосовано: {', '.join(applied_parts)}", "success")
            return redirect(url_for('students.student_details', student_id=student_id))

        conn.close()
        return redirect(url_for('admin.update_request_review', request_id=request_id))

    passport_scans = get_attachments(conn, 'update_request_passport', request_id)
    education_scans = get_attachments(conn, 'update_request_education', request_id)
    military_scans = get_attachments(conn, 'update_request_military', request_id)
    applied_fields = json.loads(row['applied_fields']) if row['applied_fields'] else []

    # Активні enrollments студента для вибору цільової групи
    enrollments = conn.execute("""
        SELECT e.group_id, e.is_primary, g.name AS group_name
        FROM student_enrollments e
        JOIN groups g ON g.id = e.group_id
        WHERE e.student_id = ? AND e.status = 'active'
        ORDER BY e.is_primary DESC, e.id
    """, (student_id,)).fetchall()
    enrollments = [dict(e) for e in enrollments]

    # Legacy-fallback: студент без enrollment'ів
    # ── ВИПРАВЛЕНО: не перезаписуємо змінну row (це рядок заявки update_requests!) ──
    if not enrollments:
        _legacy_row = conn.execute("""
            SELECT s.group_id, g.name AS group_name
            FROM students s LEFT JOIN groups g ON g.id = s.group_id
            WHERE s.id = ? AND s.group_id IS NOT NULL
        """, (student_id,)).fetchone()
        if _legacy_row:
            enrollments = [{
                'group_id': _legacy_row['group_id'],
                'group_name': _legacy_row['group_name'],
                'is_primary': True,
            }]

    # Найновіший паспорт для КОЖНОЇ групи окремо
    passports_by_group = {}
    for e in enrollments:
        _gid = e['group_id'] or -1
        p = conn.execute("""
            SELECT document_type, series, number, issued_by, issue_date, valid_until
            FROM passport_documents
            WHERE student_id = ? AND COALESCE(group_id, -1) = ?
            ORDER BY id DESC LIMIT 1
        """, (student_id, _gid)).fetchone()
        if p:
            passports_by_group[e['group_id']] = dict(p)

    conn.close()
    return render_template(
        'admin_update_request_review.html',
        row=row, allowed_fields=allowed_fields, submitted=submitted, applied_fields=applied_fields,
        passport_scans=passport_scans, education_scans=education_scans, military_scans=military_scans,
        enrollments=enrollments,
        passports_by_group=passports_by_group,
        **passport_template_context(),
    )