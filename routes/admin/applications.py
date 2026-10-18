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


@admin_bp.route('/admin/pending_students')
@permission_required('manage_pending_students')
def pending_students():
    """
    Список заявок з публічної анкети самореєстрації (routes/public_apply.py).
    За замовчуванням - лише необроблені ("new"), опрацьовані ховаються
    в окрему вкладку історії. Для кожної заявки одразу видно, чи є
    ймовірний дублікат (та сама ПІБ+дата народження) серед уже
    наявних студентів або серед інших необроблених заявок.
    """
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
            "SELECT COUNT(*) FROM attachments WHERE entity_type='pending_student' AND entity_id=?", (r['id'],)
        ).fetchone()[0]

    # Дублікати рахуємо лише для вкладки "нові" - для вже опрацьованих
    # заявок це неактуально.
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
    """
    Перегляд однієї заявки: можна виправити будь-яке поле (студенти з
    телефону одруковуються), а тоді або підтвердити (обравши групу і,
    за потреби, ліцензію та кредити скороченої програми - саме тут
    заявка стає реальним студентом), або відхилити з приміткою.
    """
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
            ]
            values = {f: (request.form.get(f) or '').strip() or None for f in fields}
            # "Дійсний до" стосується лише ID-картки - JS ховає поле для
            # книжки, але форму теоретично можна надіслати і з вимкненим JS.
            if values.get('passport_document_type') != 'ID-картка':
                values['passport_valid_until'] = None

            if not values['last_name_UA'] or not values['first_name_UA'] or not values['birth_date']:
                flash("Прізвище, ім'я і дата народження обов'язкові", "error")
            else:
                set_clause = ", ".join(f"{k}=?" for k in fields)
                conn.execute(
                    f"UPDATE pending_students SET {set_clause} WHERE id=?",
                    list(values[k] for k in fields) + [pending_id]
                )
                conn.commit()
                flash("Зміни збережено", "success")
            conn.close()
            return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

        elif action == 'reject':
            note = (request.form.get('review_note') or '').strip() or None
            conn.execute(
                "UPDATE pending_students SET status='rejected', reviewed_by=?, reviewed_at=datetime('now','localtime'), review_note=? WHERE id=?",
                (current_username(), note, pending_id)
            )
            conn.commit()
            log_action(current_username(), f"відхилив заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} (заявка ID {pending_id})", details=note or '')
            conn.close()
            flash("Заявку відхилено", "success")
            return redirect(url_for('admin.pending_students'))

        elif action == 'delete':
            # Видаляє саму заявку і її файли (фото, скани) - реального
            # студента, якщо заявку вже підтвердили, це НЕ чіпає: він
            # уже самостійний запис, не залежний від заявки.
            if row['photo_path']:
                photo_abs = os.path.join('static', row['photo_path'])
                if os.path.exists(photo_abs):
                    os.remove(photo_abs)
            for att in get_attachments(conn, 'pending_student', pending_id):
                att_abs = os.path.join('static', att['file_path'])
                if os.path.exists(att_abs):
                    os.remove(att_abs)
            for att in get_attachments(conn, 'pending_student_passport', pending_id):
                att_abs = os.path.join('static', att['file_path'])
                if os.path.exists(att_abs):
                    os.remove(att_abs)
            for att in get_attachments(conn, 'pending_student_military', pending_id):
                att_abs = os.path.join('static', att['file_path'])
                if os.path.exists(att_abs):
                    os.remove(att_abs)
            conn.execute("DELETE FROM attachments WHERE entity_type='pending_student' AND entity_id=?", (pending_id,))
            conn.execute("DELETE FROM attachments WHERE entity_type='pending_student_passport' AND entity_id=?", (pending_id,))
            conn.execute("DELETE FROM attachments WHERE entity_type='pending_student_military' AND entity_id=?", (pending_id,))
            conn.execute("DELETE FROM pending_students WHERE id=?", (pending_id,))
            conn.commit()
            log_action(current_username(), f"видалив заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} (заявка ID {pending_id})")
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

            cur = conn.execute("""
                INSERT INTO students (
                    last_name_UA, first_name_UA, middle_name_UA, last_name_ENG, first_name_ENG, birth_date, tax_id,
                    group_id, license_id, phone, phone_backup, email, program_credits_override
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row['last_name_UA'], row['first_name_UA'], row['middle_name_UA'], row['last_name_ENG'], row['first_name_ENG'],
                row['birth_date'], row['tax_id'], group_id, license_id, row['phone'], row['phone_backup'], row['email'],
                program_credits_override,
            ))
            student_id = cur.lastrowid

            # Фото - копіюємо байти з "карантинної" папки в стандартне
            # місце зберігання фото студентів (той самий шлях, що й
            # для звичайного завантаження фото на картці студента).
            if row['photo_path']:
                try:
                    from routes.photo import photo_path_for_student, _ensure_dir, PHOTOS_DIR
                    _ensure_dir()
                    src_path = os.path.join('static', row['photo_path'])
                    if os.path.exists(src_path):
                        dest_path = photo_path_for_student(student_id)
                        with open(src_path, 'rb') as src, open(dest_path, 'wb') as dst:
                            dst.write(src.read())
                        conn.execute("UPDATE students SET photo=? WHERE id=?",
                                     (f"uploads/photos/student_{student_id}.jpg", student_id))
                except Exception as e:
                    logger.error(f"Не вдалося перенести фото заявки {pending_id} студенту {student_id}: {e}")

            if row['document_type'] or row['document_number']:
                doc_number = ' '.join(x for x in [row['document_series'], row['document_number']] if x) or ''
                cur_doc = conn.execute("""
                    INSERT INTO education_documents (
                        student_id, document_type, document_type_en, document_number,
                        institution_name, institution_name_en, country, country_en, completion_date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    student_id, row['document_type'] or '', translate_to_en(row['document_type'] or ''), doc_number,
                    row['document_institution'] or '', translate_to_en(row['document_institution'] or ''),
                    row['document_country'] or '', translate_to_en(row['document_country'] or ''),
                    row['document_date'] or '',
                ))
                # Скани, завантажені разом із заявкою, - це скани саме
                # цього документа: переприв'язуємо (сам файл на диску
                # лишається на місці, змінюється лише запис у attachments).
                conn.execute(
                    "UPDATE attachments SET entity_type='education_document', entity_id=? WHERE entity_type='pending_student' AND entity_id=?",
                    (cur_doc.lastrowid, pending_id)
                )

                # Академічна довідка / Свідоцтво про визнання - абітурієнт
                # міг позначити одне, обидва чи жодне на формі /apply;
                # записуємо лише те, що реально заповнене.
                if row['foreign_reference_number'] or row['foreign_reference_institution'] or row['recognition_certificate_number'] or row['recognition_issuer']:
                    conn.execute("""
                        INSERT INTO foreign_education_docs (
                            education_doc_id, reference_number, reference_institution, reference_institution_en,
                            reference_country, reference_country_en, reference_issue_date,
                            recognition_certificate_number, recognition_issuer, recognition_issuer_en, recognition_date
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        cur_doc.lastrowid,
                        row['foreign_reference_number'], row['foreign_reference_institution'],
                        translate_to_en(row['foreign_reference_institution']) if row['foreign_reference_institution'] else None,
                        row['foreign_reference_country'],
                        translate_to_en(row['foreign_reference_country']) if row['foreign_reference_country'] else None,
                        row['foreign_reference_issue_date'],
                        row['recognition_certificate_number'], row['recognition_issuer'],
                        translate_to_en(row['recognition_issuer']) if row['recognition_issuer'] else None,
                        row['recognition_date'],
                    ))
            else:
                # Немає окремого документа про освіту, куди прив'язати
                # скани, - лишаємо їх загальними вкладеннями студента.
                conn.execute(
                    "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student' AND entity_id=?",
                    (student_id, pending_id)
                )

            if row['passport_number']:
                cur_passport = conn.execute("""
                    INSERT INTO passport_documents (
                        student_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    student_id, row['passport_document_type'] or 'Паспорт (книжка)', row['passport_series'],
                    row['passport_number'], row['passport_issued_by'], row['passport_issue_date'],
                    row['passport_valid_until'], row['passport_unique_number'],
                ))
                conn.execute(
                    "UPDATE attachments SET entity_type='passport_document', entity_id=? WHERE entity_type='pending_student_passport' AND entity_id=?",
                    (cur_passport.lastrowid, pending_id)
                )
            else:
                conn.execute(
                    "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student_passport' AND entity_id=?",
                    (student_id, pending_id)
                )

            if any([row['military_registration_number_drpvr'], row['military_registration_document'], row['military_rank']]):
                cur_military = conn.execute("""
                    INSERT INTO military (
                        student_id, registration_number_of_the_DRPVR, military_registration_document, issued_VOD,
                        military_accounting_specialty_number, military_rank, address_of_residence,
                        change_credentials, reason_for_changing_credentials
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    student_id, row['military_registration_number_drpvr'], row['military_registration_document'],
                    row['military_issued_vod'], row['military_specialty_number'], row['military_rank'],
                    row['military_address'], row['military_change_credentials'], row['military_change_reason'],
                ))
                conn.execute(
                    "UPDATE attachments SET entity_type='military', entity_id=? WHERE entity_type='pending_student_military' AND entity_id=?",
                    (cur_military.lastrowid, pending_id)
                )
            else:
                conn.execute(
                    "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student_military' AND entity_id=?",
                    (student_id, pending_id)
                )

            conn.execute(
                "UPDATE pending_students SET status='approved', reviewed_by=?, reviewed_at=datetime('now','localtime'), resulting_student_id=? WHERE id=?",
                (current_username(), student_id, pending_id)
            )
            conn.commit()
            log_action(
                current_username(),
                f"підтвердив заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} -> студент ID {student_id}",
                group_ids=[group_id],
            )
            conn.close()
            flash(f"Студента {row['last_name_UA']} {row['first_name_UA']} додано", "success")
            return redirect(url_for('students.student_details', student_id=student_id))

        elif action == 'update_existing':
            # Не створює нового студента - обраними пунктами оновлює
            # ВЖЕ НАЯВНОГО (той самий, на якого вказує "можливий
            # дублікат"). Кожен пункт - окрема галочка, щоб адмін сам
            # вирішував, що саме брати з заявки, а що лишити як є.
            existing_id = request.form.get('existing_student_id')
            if not existing_id:
                flash("Не вказано, якого студента оновлювати", "error")
                conn.close()
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))
            existing_id = int(existing_id)

            updated_parts = []

            if request.form.get('update_personal'):
                conn.execute("""
                    UPDATE students SET last_name_UA=?, first_name_UA=?, middle_name_UA=?,
                                         last_name_ENG=?, first_name_ENG=?, birth_date=?, tax_id=?
                    WHERE id=?
                """, (
                    row['last_name_UA'], row['first_name_UA'], row['middle_name_UA'],
                    row['last_name_ENG'], row['first_name_ENG'], row['birth_date'], row['tax_id'], existing_id
                ))
                updated_parts.append('особисті дані')

            if request.form.get('update_contacts'):
                conn.execute(
                    "UPDATE students SET phone=?, phone_backup=?, email=? WHERE id=?",
                    (row['phone'], row['phone_backup'], row['email'], existing_id)
                )
                updated_parts.append('контакти')

            if request.form.get('update_photo') and row['photo_path']:
                try:
                    from routes.photo import photo_path_for_student, _ensure_dir
                    _ensure_dir()
                    src_path = os.path.join('static', row['photo_path'])
                    if os.path.exists(src_path):
                        dest_path = photo_path_for_student(existing_id)
                        with open(src_path, 'rb') as src, open(dest_path, 'wb') as dst:
                            dst.write(src.read())
                        conn.execute("UPDATE students SET photo=? WHERE id=?",
                                     (f"uploads/photos/student_{existing_id}.jpg", existing_id))
                        updated_parts.append('фото')
                except Exception as e:
                    logger.error(f"Не вдалося перенести фото заявки {pending_id} студенту {existing_id}: {e}")

            doc_id_for_scans = None
            if request.form.get('add_document') and (row['document_type'] or row['document_number']):
                doc_number = ' '.join(x for x in [row['document_series'], row['document_number']] if x) or ''
                cur_doc = conn.execute("""
                    INSERT INTO education_documents (
                        student_id, document_type, document_type_en, document_number,
                        institution_name, institution_name_en, country, country_en, completion_date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    existing_id, row['document_type'] or '', translate_to_en(row['document_type'] or ''), doc_number,
                    row['document_institution'] or '', translate_to_en(row['document_institution'] or ''),
                    row['document_country'] or '', translate_to_en(row['document_country'] or ''),
                    row['document_date'] or '',
                ))
                doc_id_for_scans = cur_doc.lastrowid
                updated_parts.append('документ про освіту (додано як новий)')

            if request.form.get('add_reference_recognition') and (
                row['foreign_reference_number'] or row['foreign_reference_institution']
                or row['recognition_certificate_number'] or row['recognition_issuer']
            ):
                target_doc_id = doc_id_for_scans
                if target_doc_id is None:
                    existing_doc = conn.execute(
                        "SELECT id FROM education_documents WHERE student_id=? ORDER BY id DESC LIMIT 1", (existing_id,)
                    ).fetchone()
                    target_doc_id = existing_doc[0] if existing_doc else None

                if target_doc_id is None:
                    flash("Немає жодного документа про освіту студента, до якого можна додати довідку/свідоцтво.", "error")
                else:
                    existing_foreign = conn.execute(
                        "SELECT id FROM foreign_education_docs WHERE education_doc_id=?", (target_doc_id,)
                    ).fetchone()
                    ref_institution_en = translate_to_en(row['foreign_reference_institution']) if row['foreign_reference_institution'] else None
                    ref_country_en = translate_to_en(row['foreign_reference_country']) if row['foreign_reference_country'] else None
                    rec_issuer_en = translate_to_en(row['recognition_issuer']) if row['recognition_issuer'] else None

                    if existing_foreign:
                        conn.execute("""
                            UPDATE foreign_education_docs SET
                                reference_number=COALESCE(?, reference_number),
                                reference_institution=COALESCE(?, reference_institution),
                                reference_institution_en=COALESCE(?, reference_institution_en),
                                reference_country=COALESCE(?, reference_country),
                                reference_country_en=COALESCE(?, reference_country_en),
                                reference_issue_date=COALESCE(?, reference_issue_date),
                                recognition_certificate_number=COALESCE(?, recognition_certificate_number),
                                recognition_issuer=COALESCE(?, recognition_issuer),
                                recognition_issuer_en=COALESCE(?, recognition_issuer_en),
                                recognition_date=COALESCE(?, recognition_date)
                            WHERE education_doc_id=?
                        """, (
                            row['foreign_reference_number'], row['foreign_reference_institution'], ref_institution_en,
                            row['foreign_reference_country'], ref_country_en, row['foreign_reference_issue_date'],
                            row['recognition_certificate_number'], row['recognition_issuer'], rec_issuer_en,
                            row['recognition_date'], target_doc_id,
                        ))
                    else:
                        conn.execute("""
                            INSERT INTO foreign_education_docs (
                                education_doc_id, reference_number, reference_institution, reference_institution_en,
                                reference_country, reference_country_en, reference_issue_date,
                                recognition_certificate_number, recognition_issuer, recognition_issuer_en, recognition_date
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            target_doc_id, row['foreign_reference_number'], row['foreign_reference_institution'], ref_institution_en,
                            row['foreign_reference_country'], ref_country_en, row['foreign_reference_issue_date'],
                            row['recognition_certificate_number'], row['recognition_issuer'], rec_issuer_en, row['recognition_date'],
                        ))
                    updated_parts.append('академічна довідка/свідоцтво про визнання')

            if doc_id_for_scans:
                conn.execute(
                    "UPDATE attachments SET entity_type='education_document', entity_id=? WHERE entity_type='pending_student' AND entity_id=?",
                    (doc_id_for_scans, pending_id)
                )
            # Будь-які скани, що лишились непереприв'язаними (документ
            # не додавали чи галочку не ставили) - чіпляємо як загальні
            # файли студента, щоб не загубились.
            conn.execute(
                "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student' AND entity_id=?",
                (existing_id, pending_id)
            )

            passport_id_for_scans = None
            if request.form.get('add_passport') and row['passport_number']:
                cur_passport = conn.execute("""
                    INSERT INTO passport_documents (
                        student_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    existing_id, row['passport_document_type'] or 'Паспорт (книжка)', row['passport_series'],
                    row['passport_number'], row['passport_issued_by'], row['passport_issue_date'],
                    row['passport_valid_until'], row['passport_unique_number'],
                ))
                passport_id_for_scans = cur_passport.lastrowid
                updated_parts.append('паспортні дані (додано як новий запис)')

            if passport_id_for_scans:
                conn.execute(
                    "UPDATE attachments SET entity_type='passport_document', entity_id=? WHERE entity_type='pending_student_passport' AND entity_id=?",
                    (passport_id_for_scans, pending_id)
                )
            conn.execute(
                "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student_passport' AND entity_id=?",
                (existing_id, pending_id)
            )

            if request.form.get('add_military') and any([
                row['military_registration_number_drpvr'], row['military_registration_document'], row['military_rank']
            ]):
                existing_military = conn.execute("SELECT id FROM military WHERE student_id=?", (existing_id,)).fetchone()
                if existing_military:
                    conn.execute("""
                        UPDATE military SET registration_number_of_the_DRPVR=?, military_registration_document=?,
                                             issued_VOD=?, military_accounting_specialty_number=?, military_rank=?,
                                             address_of_residence=?, change_credentials=?, reason_for_changing_credentials=?
                        WHERE id=?
                    """, (
                        row['military_registration_number_drpvr'], row['military_registration_document'],
                        row['military_issued_vod'], row['military_specialty_number'], row['military_rank'],
                        row['military_address'], row['military_change_credentials'], row['military_change_reason'],
                        existing_military['id']
                    ))
                    military_id_for_scans = existing_military['id']
                else:
                    cur_military = conn.execute("""
                        INSERT INTO military (
                            student_id, registration_number_of_the_DRPVR, military_registration_document, issued_VOD,
                            military_accounting_specialty_number, military_rank, address_of_residence,
                            change_credentials, reason_for_changing_credentials
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        existing_id, row['military_registration_number_drpvr'], row['military_registration_document'],
                        row['military_issued_vod'], row['military_specialty_number'], row['military_rank'],
                        row['military_address'], row['military_change_credentials'], row['military_change_reason'],
                    ))
                    military_id_for_scans = cur_military.lastrowid
                updated_parts.append('військовий облік')
                conn.execute(
                    "UPDATE attachments SET entity_type='military', entity_id=? WHERE entity_type='pending_student_military' AND entity_id=?",
                    (military_id_for_scans, pending_id)
                )
            conn.execute(
                "UPDATE attachments SET entity_type='student', entity_id=? WHERE entity_type='pending_student_military' AND entity_id=?",
                (existing_id, pending_id)
            )

            if not updated_parts:
                conn.rollback()
                conn.close()
                flash("Не обрано жодного пункту для оновлення", "error")
                return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

            conn.execute(
                "UPDATE pending_students SET status='approved', reviewed_by=?, reviewed_at=datetime('now','localtime'), resulting_student_id=? WHERE id=?",
                (current_username(), existing_id, pending_id)
            )
            conn.commit()
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
    licenses = conn.execute("SELECT id, short_name_ua, name_ua FROM institution_licenses WHERE is_active=1 ORDER BY id").fetchall()

    existing_student = conn.execute("""
        SELECT id, last_name_UA, first_name_UA FROM students
        WHERE LOWER_UA(last_name_UA) = LOWER_UA(?) AND LOWER_UA(first_name_UA) = LOWER_UA(?) AND birth_date = ?
    """, (row['last_name_UA'], row['first_name_UA'], row['birth_date'])).fetchone()

    scans = get_attachments(conn, 'pending_student', pending_id)
    passport_scans = get_attachments(conn, 'pending_student_passport', pending_id)
    military_scans = get_attachments(conn, 'pending_student_military', pending_id)

    conn.close()
    return render_template(
        'admin_pending_student_review.html',
        row=row, groups=groups, licenses=licenses, existing_student=existing_student,
        scans=scans, passport_scans=passport_scans, military_scans=military_scans,
    )


@admin_bp.route('/admin/pending_students/<int:pending_id>/photo', methods=['POST'])
@permission_required('manage_pending_students')
def pending_student_photo(pending_id):
    """
    Ручна (пере)обрізка фото абітурієнта на сторінці перегляду заявки -
    коли автоматична обрізка по центру (яку робить сама публічна форма)
    вийшла невдало. Той самий Cropper.js-підхід, що й на картці
    студента (students.upload_photo), лише зберігає результат у
    "карантинну" папку заявки, а не в фото реального студента.
    """
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
        # Новий файл не завантажували - переобрізаємо той самий, що вже
        # є (адмін просто підправляє рамку на наявному фото).
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
        new_path = process_and_save_pending_photo_with_crop(file_bytes, crop_box, old_rel_path=row['photo_path'])
    except ValueError as e:
        conn.close()
        flash(str(e), "error")
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    conn.execute("UPDATE pending_students SET photo_path=? WHERE id=?", (new_path, pending_id))
    conn.commit()
    conn.close()
    flash("Фото оновлено", "success")
    return redirect(url_for('admin.pending_student_review', pending_id=pending_id))


@admin_bp.route('/admin/update_requests')
@permission_required('manage_students')
def update_requests():
    """
    Список заявок на оновлення даних наявного студента (з токен-
    посилань, які адмін сам створює на картці студента - routes/
    students.generate_update_link). За замовчуванням - лише ті, що
    студент уже заповнив і чекають на модерацію (status='submitted').
    """
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
    return render_template('admin_update_requests.html', rows=rows, status_filter=status_filter, counts=counts)


@admin_bp.route('/admin/update_requests/<int:request_id>/photo', methods=['POST'])
@permission_required('manage_students')
def update_request_photo(request_id):
    """
    Ручна (пере)обрізка фото, надісланого студентом через /update-info -
    той самий Cropper.js-підхід, що й для заявок на реєстрацію
    (admin.pending_student_photo), лише зберігає результат у поле
    update_requests.photo_path замість pending_students.photo_path.
    """
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
        # Новий файл не завантажували - переобрізаємо той самий, що вже
        # є (адмін просто підправляє рамку на наявному фото).
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
        new_path = process_and_save_pending_photo_with_crop(file_bytes, crop_box, old_rel_path=row['photo_path'])
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
@permission_required('manage_students')
def update_request_review(request_id):
    """Перегляд однієї заявки на оновлення - вибіркове застосування
    полів до наявного студента (та сама логіка "яке поле застосувати",
    що й для дублікатів заявок на реєстрацію)."""
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

    allowed_fields = json.loads(row['allowed_fields'])
    submitted = json.loads(row['submitted_data']) if row['submitted_data'] else {}

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'delete':
            for entity_type in ('update_request_passport', 'update_request_education', 'update_request_military'):
                for att in get_attachments(conn, entity_type, request_id):
                    # Після застосування заявки той самий файл копіюється і
                    # на реальний запис студента (passport_document/
                    # education_document/military) - видаляти файл з диска
                    # можна, лише якщо жоден ІНШИЙ рядок attachments на
                    # нього більше не посилається, інакше документ студента
                    # лишиться з "мертвим" посиланням.
                    still_used = conn.execute(
                        "SELECT 1 FROM attachments WHERE file_path=? AND id!=? LIMIT 1",
                        (att['file_path'], att['id'])
                    ).fetchone()
                    att_path = os.path.join('static', att['file_path'])
                    if not still_used and os.path.exists(att_path):
                        os.remove(att_path)
                conn.execute("DELETE FROM attachments WHERE entity_type=? AND entity_id=?", (entity_type, request_id))
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
                "UPDATE update_requests SET status='rejected', reviewed_by=?, reviewed_at=datetime('now','localtime'), review_note=? WHERE id=?",
                (current_username(), note, request_id)
            )
            conn.commit()
            log_action(current_username(), f"відхилив заявку на оновлення даних ID {request_id}", details=note or '')
            conn.close()
            flash("Заявку відхилено", "success")
            return redirect(url_for('admin.update_requests'))

        elif action == 'apply':
            student_id = row['student_id']
            applied_parts = []
            # Які саме секції/поля реально застосовано (на відміну від
            # applied_parts, який лише для flash-повідомлення й логу,
            # applied_field_keys зберігається в update_requests, щоб
            # сторінку заявки можна було відкрити й пізніше і побачити,
            # що саме було прийнято - а не лише "заявка застосована".
            applied_field_keys = []

            simple_updates = {}
            for key in ('last_name_UA', 'first_name_UA', 'middle_name_UA', 'phone', 'phone_backup', 'email', 'tax_id', 'edebo_code'):
                if key in allowed_fields and request.form.get(f'apply_{key}'):
                    simple_updates[key] = submitted.get(key)
            if 'last_name_UA' in simple_updates or 'first_name_UA' in simple_updates:
                # Англійська транслітерація генерується автоматично при
                # створенні студента (add_student) - але раніше тут, при
                # застосуванні зміни прізвища/імені із заявки, лишалась
                # НЕ оновленою, доки адмін не підправить її вручну окремо.
                # Беремо нове значення там, де воно змінюється, і поточне
                # студентське там, де ні (напр. якщо міняється лише
                # прізвище - ім'я для транслітерації береться з row).
                new_last_ua = simple_updates.get('last_name_UA', row['last_name_UA'])
                new_first_ua = simple_updates.get('first_name_UA', row['first_name_UA'])
                last_eng, first_eng = generate_english_name(new_last_ua, new_first_ua)
                simple_updates['last_name_ENG'] = last_eng
                simple_updates['first_name_ENG'] = first_eng
            if simple_updates:
                set_clause = ", ".join(f"{k}=?" for k in simple_updates)
                conn.execute(f"UPDATE students SET {set_clause} WHERE id=?", list(simple_updates.values()) + [student_id])
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
                    conn.execute("UPDATE students SET photo=? WHERE id=?", (f"uploads/photos/student_{student_id}.jpg", student_id))
                    applied_parts.append('фото')
                    applied_field_keys.append('photo')

            if 'passport' in allowed_fields and submitted.get('passport') and request.form.get('apply_passport'):
                p = submitted['passport']
                cur_p = conn.execute("""
                    INSERT INTO passport_documents (student_id, document_type, series, number, issued_by, issue_date, valid_until, unique_number)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (student_id, p.get('document_type') or 'Паспорт (книжка)', p.get('series'), p.get('number'),
                      p.get('issued_by'), p.get('issue_date'), p.get('valid_until'), p.get('unique_number')))
                # Копіюємо (а не переносимо) скани - інакше сторінка заявки,
                # яку адмін може відкрити повторно пізніше, лишилась би без
                # жодного файлу, бо він "переїхав" би на сам документ.
                conn.execute("""
                    INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_at, uploaded_by)
                    SELECT 'passport_document', ?, file_path, original_name, uploaded_at, uploaded_by
                    FROM attachments WHERE entity_type='update_request_passport' AND entity_id=?
                """, (cur_p.lastrowid, request_id))
                applied_parts.append('паспортні дані (новий запис)')
                applied_field_keys.append('passport')

            d = submitted.get('education_document') or {}
            education_doc_id = None  # id новоствореного (чи наявного) документа - для довідки/свідоцтва нижче

            if 'education_document' in allowed_fields and d.get('document_type') and request.form.get('apply_education_document'):
                # education_documents не має окремої колонки "серія" (на
                # відміну від passport_documents) - раніше серія, яку
                # студент вводив у заявці, просто ніде не опинялась:
                # не показувалась при перегляді й не застосовувалась.
                # Об'єднуємо її з номером, як і скрізь у системі
                # (так само, як серія+номер паспорта зводяться в один рядок).
                document_number = ((d.get('series') + ' ') if d.get('series') else '') + (d.get('number') or '')
                cur_d = conn.execute("""
                    INSERT INTO education_documents (student_id, document_type, document_type_en, document_number,
                        institution_name, institution_name_en, country, country_en, completion_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (student_id, d.get('document_type') or '', translate_to_en(d.get('document_type') or ''), document_number,
                      d.get('institution') or '', translate_to_en(d.get('institution') or ''),
                      d.get('country') or '', translate_to_en(d.get('country') or ''),
                      d.get('completion_date') or ''))
                education_doc_id = cur_d.lastrowid
                conn.execute("""
                    INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_at, uploaded_by)
                    SELECT 'education_document', ?, file_path, original_name, uploaded_at, uploaded_by
                    FROM attachments WHERE entity_type='update_request_education' AND entity_id=?
                """, (education_doc_id, request_id))
                applied_parts.append('документ про освіту (новий запис)')
                applied_field_keys.append('education_document')

            # Академічна довідка / Свідоцтво про визнання - незалежні від
            # базового документа: студент міг подати лише їх (без нового
            # документа), щоб долучити до вже наявного. Якщо вище не
            # створено нового документа - прикріплюємо до найновішого
            # наявного запису студента; якщо такого взагалі немає,
            # прикріпити нема до чого, тож пропускаємо з позначкою в логах.
            reference = d.get('reference') or {}
            recognition = d.get('recognition') or {}
            apply_reference = ('education_document_reference' in allowed_fields and reference
                                and request.form.get('apply_education_reference'))
            apply_recognition = ('education_document_recognition' in allowed_fields and recognition
                                  and request.form.get('apply_education_recognition'))

            if apply_reference or apply_recognition:
                if education_doc_id is None:
                    existing_doc = conn.execute(
                        "SELECT id FROM education_documents WHERE student_id=? ORDER BY id DESC LIMIT 1", (student_id,)
                    ).fetchone()
                    education_doc_id = existing_doc[0] if existing_doc else None

                if education_doc_id is None:
                    flash("Немає жодного документа про освіту, до якого можна додати довідку/свідоцтво - спершу додайте основний документ.", "error")
                else:
                    existing_foreign = conn.execute(
                        "SELECT id FROM foreign_education_docs WHERE education_doc_id=?", (education_doc_id,)
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
                        # COALESCE - щоб застосування лише довідки (без
                        # свідоцтва) не затирало вже наявне свідоцтво на
                        # цьому ж документі, і навпаки.
                        conn.execute("""
                            UPDATE foreign_education_docs SET
                                reference_number=COALESCE(?, reference_number),
                                reference_institution=COALESCE(?, reference_institution),
                                reference_institution_en=COALESCE(?, reference_institution_en),
                                reference_country=COALESCE(?, reference_country),
                                reference_country_en=COALESCE(?, reference_country_en),
                                reference_issue_date=COALESCE(?, reference_issue_date),
                                recognition_certificate_number=COALESCE(?, recognition_certificate_number),
                                recognition_issuer=COALESCE(?, recognition_issuer),
                                recognition_issuer_en=COALESCE(?, recognition_issuer_en),
                                recognition_date=COALESCE(?, recognition_date)
                            WHERE education_doc_id=?
                        """, (ref_number, ref_institution, ref_institution_en, ref_country, ref_country_en, ref_issue_date,
                              rec_number, rec_issuer, rec_issuer_en, rec_date, education_doc_id))
                    else:
                        conn.execute("""
                            INSERT INTO foreign_education_docs (
                                education_doc_id, reference_number, reference_institution, reference_institution_en,
                                reference_country, reference_country_en, reference_issue_date,
                                recognition_certificate_number, recognition_issuer, recognition_issuer_en, recognition_date
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (education_doc_id, ref_number, ref_institution, ref_institution_en, ref_country, ref_country_en,
                              ref_issue_date, rec_number, rec_issuer, rec_issuer_en, rec_date))

                    if apply_reference:
                        applied_parts.append('академічна довідка')
                        applied_field_keys.append('education_document_reference')
                    if apply_recognition:
                        applied_parts.append('свідоцтво про визнання')
                        applied_field_keys.append('education_document_recognition')

            if 'military' in allowed_fields and submitted.get('military') and request.form.get('apply_military'):
                m = submitted['military']
                existing_military = conn.execute("SELECT id FROM military WHERE student_id=?", (student_id,)).fetchone()
                if existing_military:
                    conn.execute("""
                        UPDATE military SET registration_number_of_the_DRPVR=?, military_registration_document=?,
                            issued_VOD=?, military_accounting_specialty_number=?, military_rank=?, address_of_residence=?
                        WHERE id=?
                    """, (m.get('registration_number_of_the_DRPVR'), m.get('military_registration_document'),
                          m.get('issued_VOD'), m.get('military_accounting_specialty_number'), m.get('military_rank'),
                          m.get('address_of_residence'), existing_military['id']))
                    military_id = existing_military['id']
                else:
                    cur_m = conn.execute("""
                        INSERT INTO military (student_id, registration_number_of_the_DRPVR, military_registration_document,
                            issued_VOD, military_accounting_specialty_number, military_rank, address_of_residence)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (student_id, m.get('registration_number_of_the_DRPVR'), m.get('military_registration_document'),
                          m.get('issued_VOD'), m.get('military_accounting_specialty_number'), m.get('military_rank'),
                          m.get('address_of_residence')))
                    military_id = cur_m.lastrowid
                conn.execute("""
                    INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_at, uploaded_by)
                    SELECT 'military', ?, file_path, original_name, uploaded_at, uploaded_by
                    FROM attachments WHERE entity_type='update_request_military' AND entity_id=?
                """, (military_id, request_id))
                applied_parts.append('військові дані')
                applied_field_keys.append('military')

            if not applied_parts:
                conn.rollback()
                conn.close()
                flash("Не обрано жодного пункту для застосування", "error")
                return redirect(url_for('admin.update_request_review', request_id=request_id))

            conn.execute(
                "UPDATE update_requests SET status='approved', reviewed_by=?, reviewed_at=datetime('now','localtime'), applied_fields=? WHERE id=?",
                (current_username(), json.dumps(applied_field_keys), request_id)
            )
            conn.commit()
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

    conn.close()
    return render_template(
        'admin_update_request_review.html',
        row=row, allowed_fields=allowed_fields, submitted=submitted, applied_fields=applied_fields,
        passport_scans=passport_scans, education_scans=education_scans, military_scans=military_scans,
    )