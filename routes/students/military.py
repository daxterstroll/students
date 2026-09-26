"""
routes/students/military.py
============================
Військовий облік студента: перегляд/редагування (один запис на
студента) + скановані копії (додавання, перегляд, видалення окремого
файлу), видалення всього запису.
"""
from flask import render_template, request, redirect, url_for, session, flash
from routes.db import get_db
from routes.utils import log_action, login_required, permission_required, logger
from routes.helpers import current_username
from routes.students import students_bp
import sqlite3
from routes.utils import get_attachments, save_multiple_attachments
from datetime import datetime
import os


@students_bp.route('/students/<int:student_id>/military/add', methods=['GET', 'POST'])
@login_required('')
def add_military(student_id):
    """Форма додавання даних військового обліку студенту, який ще їх не має."""
    if request.method == 'POST':
        issued_VOD_raw = request.form.get('issued_VOD', '').strip()
        if issued_VOD_raw:
            issued_VOD_clean = issued_VOD_raw.replace("-", ".")
            try:
                datetime.strptime(issued_VOD_clean, "%d.%m.%Y")
                issued_VOD = issued_VOD_clean
            except ValueError:
                flash("Невірний формат дати. Введіть у форматі ДД.ММ.РРРР")
                return render_template('add_military.html', student_id=student_id)
        else:
            issued_VOD = None

        data = (
            student_id,
            request.form['registration_number_of_the_DRPVR'],
            request.form['military_registration_document'],
            issued_VOD,
            request.form['military_accounting_specialty_number'],
            request.form['military_rank'],
            request.form['change_credentials'],
            request.form['reason_for_changing_credentials'],
            request.form['being_on_military_registration'],
            request.form['address_of_residence'],
        )

        conn = get_db()
        conn.execute("""
            INSERT INTO military (
                student_id, registration_number_of_the_DRPVR, military_registration_document,
                issued_VOD, military_accounting_specialty_number, military_rank,
                change_credentials, reason_for_changing_credentials,
                being_on_military_registration, address_of_residence
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, data)
        conn.commit()

        student_row = conn.execute(
            "SELECT last_name_UA, first_name_UA, group_id FROM students WHERE id=?", (student_id,)
        ).fetchone()
        conn.close()

        log_action(
            current_username(),
            f"додав військові дані: {student_row['last_name_UA']} {student_row['first_name_UA']} (ID {student_id})",
            group_ids=[student_row['group_id']]
        )
        return redirect(url_for('students.student_list'))

    return render_template('add_military.html', student_id=student_id)


@students_bp.route('/students/<int:student_id>/military', methods=['GET', 'POST'])
@login_required('')
def military_data(student_id):
    """Форма перегляду/редагування наявних даних військового обліку студента."""
    MAX_FILES = 5
    conn = get_db()
    military = conn.execute("SELECT * FROM military WHERE student_id = ?", (student_id,)).fetchone()

    if request.method == 'POST':
        issued_VOD_raw = request.form['issued_VOD'].strip()
        issued_VOD_clean = issued_VOD_raw.replace("-", ".")
        try:
            datetime.strptime(issued_VOD_clean, "%d.%m.%Y")
            issued_VOD = issued_VOD_clean
        except ValueError:
            flash("Невірний формат дати. Введіть у форматі ДД.ММ.РРРР")
            military_attachments = get_attachments(conn, 'military', military['id']) if military else []
            return render_template('edit_military.html', student_id=student_id, military=military,
                                    military_attachments=military_attachments, max_files=MAX_FILES)

        new_scans = [f for f in request.files.getlist('military_scans') if f and f.filename]
        existing_count = get_attachments(conn, 'military', military['id']) if military else []
        if len(existing_count) + len(new_scans) > MAX_FILES:
            flash(f'Забагато файлів - максимум {MAX_FILES} на запис (вже є {len(existing_count)}).', 'error')
            return render_template('edit_military.html', student_id=student_id, military=military,
                                    military_attachments=existing_count, max_files=MAX_FILES)

        data = (
            request.form['registration_number_of_the_DRPVR'],
            request.form['military_registration_document'],
            issued_VOD,
            request.form['military_accounting_specialty_number'],
            request.form['military_rank'],
            request.form['change_credentials'],
            request.form['reason_for_changing_credentials'],
            request.form['being_on_military_registration'],
            request.form['address_of_residence'],
            student_id
        )
        if military:
            conn.execute("""
                UPDATE military SET
                    registration_number_of_the_DRPVR=?, military_registration_document=?,
                    issued_VOD=?, military_accounting_specialty_number=?, military_rank=?,
                    change_credentials=?, reason_for_changing_credentials=?,
                    being_on_military_registration=?, address_of_residence=?
                WHERE student_id=?
            """, data)
            military_id = military['id']
        else:
            cur = conn.execute("""
                INSERT INTO military (
                    registration_number_of_the_DRPVR, military_registration_document,
                    issued_VOD, military_accounting_specialty_number, military_rank,
                    change_credentials, reason_for_changing_credentials,
                    being_on_military_registration, address_of_residence, student_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, data)
            military_id = cur.lastrowid
        conn.commit()

        if new_scans:
            save_multiple_attachments(conn, 'military', military_id, new_scans, 'military_records', current_username())
            conn.commit()

        student_row = conn.execute(
            "SELECT last_name_UA, first_name_UA, group_id FROM students WHERE id=?", (student_id,)
        ).fetchone()
        conn.close()

        action_name = "редагував" if military else "додав"
        log_action(
            current_username(),
            f"{action_name} військові дані: {student_row['last_name_UA']} {student_row['first_name_UA']} (ID {student_id})",
            group_ids=[student_row['group_id']]
        )
        return redirect(url_for('students.student_list'))

    military_attachments = get_attachments(conn, 'military', military['id']) if military else []
    conn.close()
    return render_template('edit_military.html', student_id=student_id, military=military,
                            military_attachments=military_attachments, max_files=MAX_FILES)


@students_bp.route('/students/<int:student_id>/military/delete_attachment', methods=['POST'])
@login_required('')
def delete_military_attachment(student_id):
    """Видаляє один скан із військових даних студента (файл з диска + запис)."""
    conn = get_db()
    attachment_id = request.form.get('attachment_id')
    att = conn.execute("SELECT file_path FROM attachments WHERE id=?", (attachment_id,)).fetchone()
    if att:
        att_path = os.path.join('static', att['file_path'])
        if os.path.exists(att_path):
            os.remove(att_path)
        conn.execute("DELETE FROM attachments WHERE id=?", (attachment_id,))
        conn.commit()
        flash('Скан видалено', 'success')
    else:
        flash('Файл не знайдено', 'error')
    conn.close()
    return redirect(url_for('students.military_data', student_id=student_id))


@students_bp.route('/students/<int:student_id>/military/delete')
@permission_required('manage_students')
def delete_military(student_id):
    """Видаляє запис військового обліку студента (разом зі сканами)."""
    conn = get_db()
    student_row = conn.execute(
        "SELECT last_name_UA, first_name_UA, group_id FROM students WHERE id=?", (student_id,)
    ).fetchone()
    military_row = conn.execute("SELECT id FROM military WHERE student_id = ?", (student_id,)).fetchone()
    if military_row:
        for att in get_attachments(conn, 'military', military_row['id']):
            att_path = os.path.join('static', att['file_path'])
            if os.path.exists(att_path):
                os.remove(att_path)
        conn.execute("DELETE FROM attachments WHERE entity_type='military' AND entity_id=?", (military_row['id'],))
    conn.execute("DELETE FROM military WHERE student_id = ?", (student_id,))
    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"ВИДАЛИВ військові дані: {student_row['last_name_UA']} {student_row['first_name_UA']} (ID {student_id})",
        group_ids=[student_row['group_id']] if student_row else []
    )
    return redirect(url_for('students.student_list'))
