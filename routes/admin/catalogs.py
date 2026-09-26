"""
routes/admin/catalogs.py
========================
Довідники: спеціальності, ступені, освітні програми, назви
кваліфікацій, акредитації, ліцензії (+ переведення між ліцензіями,
звіт по ліцензіях), дипломи, масове призначення періодів навчання.
"""
from flask import render_template, request, redirect, url_for, flash, session, send_file
from routes.db import get_db
from routes.utils import log_action, permission_required, logger
from routes.helpers import current_username, sort_ukrainian
from routes.admin import admin_bp
import sqlite3
import os
import uuid
from datetime import datetime


@admin_bp.route('/admin/study_periods/bulk_assign', methods=['GET', 'POST'])
@permission_required('study_periods')
def manage_study_periods_bulk_assign():
    """Масове призначення періоду навчання (філія, дати, група) одразу декільком студентам — обраній групі цілком та/або довільному списку студентів."""
    db = get_db()
    cursor = db.cursor()

    # -------------------- POST --------------------
    if request.method == 'POST':
        mode = request.form.get('mode')
        group_id = request.form.get('group_id', type=int)
        student_ids_raw = request.form.get('student_ids_hidden', '')
        extra_student_ids = [int(x) for x in student_ids_raw.split(',') if x.strip().isdigit()]

        target_student_ids = set(extra_student_ids)

        if group_id and mode in ('whole_group', 'some_from_group'):
            cursor.execute("""
                SELECT id FROM students WHERE group_id = ? AND archived = FALSE
            """, (group_id,))
            target_student_ids.update(r['id'] for r in cursor.fetchall())

        if not target_student_ids:
            flash('Не обрано жодного студента', 'danger')
            return redirect(url_for('admin.manage_study_periods_bulk_assign'))

        filiya = (request.form.get('filiya') or '').strip()
        filiya_en = (request.form.get('filiya_en') or '').strip() or None
        group_name = (request.form.get('group_name') or '').strip() or None
        start_date = (request.form.get('start_date') or '').strip() or None
        end_date = (request.form.get('end_date') or '').strip() or None
        period_order = request.form.get('period_order', type=int) or 0
        note = (request.form.get('note') or '').strip() or None

        if not filiya:
            flash('Потрібно вказати філію', 'danger')
            return redirect(url_for('admin.manage_study_periods_bulk_assign',
                                  group_id=group_id or '',
                                  student_ids=','.join(map(str, extra_student_ids)),
                                  mode=mode))

        added = 0
        try:
            for student_id in target_student_ids:
                cursor.execute("""
                    INSERT INTO student_study_periods
                        (student_id, filiya, filiya_en, group_name, start_date, end_date, period_order, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (student_id, filiya, filiya_en, group_name, start_date, end_date, period_order, note))
                added += 1

            db.commit()
            log_action(
                current_username(),
                f"масово присвоїв період навчання ({filiya}) для {added} студент(ів)"
            )
            flash(f'Період навчання успішно додано {added} студент(ам)', 'success')
        except sqlite3.Error as e:
            db.rollback()
            logger.error(f"Помилка масового призначення періоду навчання (group_id={group_id}): {e}", exc_info=True)
            flash(f'Помилка збереження: {e}', 'danger')

        return redirect(url_for('admin.manage_study_periods_bulk_assign',
                              group_id=group_id or '',
                              student_ids=','.join(map(str, extra_student_ids)),
                              mode=mode))

    # -------------------- GET --------------------
    cursor.execute("SELECT id, name FROM groups WHERE archived = FALSE ORDER BY name")
    groups = cursor.fetchall()

    # Всі студенти з інформацією про групу (для JS-фільтрації)
    cursor.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA, 
               s.group_id, g.name AS group_name
        FROM students s
        LEFT JOIN groups g ON s.group_id = g.id
        WHERE s.archived = FALSE
        ORDER BY s.last_name_UA, s.first_name_UA
    """)
    all_students = cursor.fetchall()

    # Параметри
    mode = request.args.get('mode', 'any_students')
    group_id = request.args.get('group_id', type=int)
    student_ids_param = request.args.get('student_ids', '')
    extra_student_ids = [int(x) for x in student_ids_param.split(',') if x.strip().isdigit()]

    target_student_ids = set(extra_student_ids)

    if group_id and mode in ('whole_group', 'some_from_group'):
        cursor.execute("""
            SELECT id FROM students WHERE group_id = ? AND archived = FALSE
        """, (group_id,))
        target_student_ids.update(r['id'] for r in cursor.fetchall())

    # Прев'ю
    target_students = []
    if target_student_ids:
        placeholders = ','.join('?' for _ in target_student_ids)
        cursor.execute(f"""
            SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA, g.name AS group_name
            FROM students s
            LEFT JOIN groups g ON s.group_id = g.id
            WHERE s.id IN ({placeholders})
            ORDER BY s.last_name_UA, s.first_name_UA
        """, tuple(target_student_ids))
        target_students_rows = cursor.fetchall()

        cursor.execute(f"""
            SELECT student_id, filiya, filiya_en, group_name, start_date, end_date, period_order, note
            FROM student_study_periods
            WHERE student_id IN ({placeholders})
            ORDER BY student_id, period_order ASC, start_date ASC
        """, tuple(target_student_ids))
        periods_rows = cursor.fetchall()

        periods_by_student = {}
        for p in periods_rows:
            periods_by_student.setdefault(p['student_id'], []).append(p)

        for s in target_students_rows:
            s_dict = dict(s)
            s_dict['existing_periods'] = periods_by_student.get(s['id'], [])
            target_students.append(s_dict)

    return render_template(
        'manage_study_periods_bulk_assign.html',
        groups=groups,
        all_students=all_students,
        selected_group_id=group_id,
        selected_student_ids=extra_student_ids,
        target_students=target_students,
        selected_mode=mode
    )


@admin_bp.route('/admin/manage_diplomas', methods=['GET', 'POST'])
@permission_required('manage_diplomas')
def manage_diplomas():
    """Перегляд і масове збереження номерів диплома/додатка для всіх студентів обраної групи."""
    conn = get_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    if request.method == "POST":
        group_id = request.form.get("group_id")
        cursor.execute("SELECT s.id FROM students s WHERE s.group_id = ?", (group_id,))
        students = cursor.fetchall()

        for student in students:
            student_id = student['id']
            diploma_number = request.form.get(f'diploma_number_{student_id}', '').strip()
            appendix_number = request.form.get(f'appendix_number_{student_id}', '').strip()
            if diploma_number:
                diploma_number = diploma_number.zfill(6)
            cursor.execute("SELECT id FROM diplomas WHERE student_id=?", (student_id,))
            exists = cursor.fetchone()
            if exists:
                cursor.execute("""
                    UPDATE diplomas SET diploma_number=?, appendix_number=? WHERE student_id=?
                """, (diploma_number, appendix_number, student_id))
            else:
                cursor.execute("""
                    INSERT INTO diplomas(student_id, diploma_number, appendix_number) VALUES (?, ?, ?)
                """, (student_id, diploma_number, appendix_number))

        conn.commit()

        group_row = cursor.execute("SELECT name FROM groups WHERE id=?", (group_id,)).fetchone()
        group_name = group_row['name'] if group_row else f"ID {group_id}"
        log_action(
            current_username(),
            f"зберіг дипломи для групи: {group_name} (ID {group_id})",
            details=f"студентів оброблено: {len(students)}"
        )

        flash("Дані збережено")
        return redirect(url_for('admin.manage_diplomas', group_id=group_id))

    cursor.execute("SELECT id, name, start_year FROM groups WHERE archived = FALSE ORDER BY name")
    groups = cursor.fetchall()
    selected_group = request.args.get("group_id")
    if selected_group is not None:
        selected_group = int(selected_group)

    students = []
    if selected_group:
        cursor.execute("""
            SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
                   d.diploma_number, d.appendix_number
            FROM students s
            LEFT JOIN diplomas d ON s.id = d.student_id
            WHERE s.group_id = ?
        """, (selected_group,))
        students = cursor.fetchall()
        students = sort_ukrainian(
            students,
            key_func=lambda s: f"{s['last_name_UA']} {s['first_name_UA']} {s['middle_name_UA']}"
        )

    return render_template("manage_diplomas.html", groups=groups, students=students, selected_group=selected_group)


@admin_bp.route('/admin/manage_accreditations', methods=['GET', 'POST'])
@permission_required('manage_accreditations')
def manage_accreditations():
    """CRUD-сторінка довідників акредитацій (ступінь + спеціальність + текст українською/англійською) для підстановки в документи."""
    conn = get_db()
    cursor = conn.cursor()

    if request.method == 'POST' and 'add' in request.form:
        degree = request.form.get('degree')
        specialty = request.form.get('specialty')
        text_ua = request.form.get('text_ua')
        text_en = request.form.get('text_en')
        cursor.execute("""
            INSERT INTO accreditations (degree, specialty, text_ua, text_en) VALUES (?, ?, ?, ?)
        """, (degree, specialty, text_ua, text_en))
        conn.commit()
        log_action(
            current_username(),
            f"додав акредитацію: {degree} / {specialty}"
        )
        return redirect(url_for('admin.manage_accreditations'))

    if request.method == 'POST' and 'edit' in request.form:
        acc_id = request.form.get('id')
        degree = request.form.get('degree')
        specialty = request.form.get('specialty')
        text_ua = request.form.get('text_ua')
        text_en = request.form.get('text_en')
        cursor.execute("""
            UPDATE accreditations SET degree=?, specialty=?, text_ua=?, text_en=? WHERE id=?
        """, (degree, specialty, text_ua, text_en, acc_id))
        conn.commit()
        log_action(
            current_username(),
            f"редагував акредитацію ID {acc_id}: {degree} / {specialty}"
        )
        return redirect(url_for('admin.manage_accreditations'))

    if request.method == 'POST' and 'delete' in request.form:
        acc_id = request.form.get('id')
        row = cursor.execute("SELECT degree, specialty FROM accreditations WHERE id=?", (acc_id,)).fetchone()
        cursor.execute("DELETE FROM accreditations WHERE id=?", (acc_id,))
        conn.commit()
        log_action(
            current_username(),
            f"ВИДАЛИВ акредитацію ID {acc_id}: {row[0] if row else ''} / {row[1] if row else ''}"
        )
        return redirect(url_for('admin.manage_accreditations'))

    cursor.execute("SELECT id, degree, specialty, text_ua, text_en FROM accreditations ORDER BY degree, specialty")
    accreditations = cursor.fetchall()
    cursor.execute("SELECT DISTINCT specialty FROM groups WHERE archived = FALSE ORDER BY specialty")
    groups = cursor.fetchall()

    return render_template('manage_accreditations.html', accreditations=accreditations, groups=groups)


@admin_bp.route('/admin/manage_specialties', methods=['GET', 'POST'])
@permission_required('manage_specialties')
def manage_specialties():
    """
    Каталог галузей знань і спеціальностей (Постанова КМУ №1021 від
    30.08.2024, чинна з 01.11.2024). Дозволяє:
      - вмикати/вимикати актуальність спеціальності для цього закладу
        (is_active) - неактивні не показуються у випадаючому списку на
        сторінці "Групи", але не видаляються (щоб не зламати наявні
        групи, які вже на них посилаються);
      - додавати власні спеціальності, яких немає в офіційному переліку
        (is_custom=1) - позначаються окремо, щоб було видно, що це не
        з держреєстру;
      - редагувати назву вже наявного запису (напр. якщо офіційний
        текст пізніше уточнили).
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'toggle_active':
            code = request.form.get('code')
            conn.execute("UPDATE specialties SET is_active = 1 - is_active WHERE code = ?", (code,))
            conn.commit()
            row = conn.execute("SELECT name_ua, is_active FROM specialties WHERE code = ?", (code,)).fetchone()
            log_action(
                current_username(),
                f"{'увімкнув' if row['is_active'] else 'вимкнув'} спеціальність: {code} {row['name_ua']}"
            )

        elif action == 'add':
            code = (request.form.get('code') or '').strip()
            name_ua = (request.form.get('name_ua') or '').strip()
            short_name = (request.form.get('short_name') or '').strip() or None
            name_en = (request.form.get('name_en') or '').strip() or None
            field_code = request.form.get('knowledge_field_code')

            if not code or not name_ua or not field_code:
                flash("Заповніть код, назву і галузь знань.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO specialties (code, name_ua, short_name, name_en, knowledge_field_code, is_active, is_custom) VALUES (?, ?, ?, ?, ?, 1, 1)",
                        (code, name_ua, short_name, name_en, field_code)
                    )
                    conn.commit()
                    flash(f"Спеціальність «{code} {name_ua}» додано.", "success")
                    log_action(current_username(), f"додав власну спеціальність: {code} {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Спеціальність з кодом «{code}» вже існує.", "error")

        elif action == 'edit':
            code = request.form.get('code')
            name_ua = (request.form.get('name_ua') or '').strip()
            short_name = (request.form.get('short_name') or '').strip() or None
            name_en = (request.form.get('name_en') or '').strip() or None
            field_code = request.form.get('knowledge_field_code')
            if not name_ua or not field_code:
                flash("Заповніть назву і галузь знань.", "error")
            else:
                conn.execute(
                    "UPDATE specialties SET name_ua = ?, short_name = ?, name_en = ?, knowledge_field_code = ? WHERE code = ?",
                    (name_ua, short_name, name_en, field_code, code)
                )
                conn.commit()
                flash("Спеціальність оновлено.", "success")
                log_action(current_username(), f"редагував спеціальність: {code} {name_ua}")

        elif action == 'edit_field':
            field_code = request.form.get('field_code')
            field_name_ua = (request.form.get('field_name_ua') or '').strip()
            field_name_en = (request.form.get('field_name_en') or '').strip() or None
            if not field_name_ua:
                flash("Заповніть назву галузі знань.", "error")
            else:
                conn.execute(
                    "UPDATE knowledge_fields SET name_ua = ?, name_en = ? WHERE code = ?",
                    (field_name_ua, field_name_en, field_code)
                )
                conn.commit()
                flash("Галузь знань оновлено.", "success")
                log_action(current_username(), f"редагував галузь знань: {field_code} {field_name_ua}")

        elif action == 'delete':
            code = request.form.get('code')
            in_use = conn.execute("SELECT COUNT(*) AS c FROM groups WHERE specialty_code = ?", (code,)).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - є {in_use} груп(и), що посилаються на цю спеціальність. "
                      f"Вимкніть актуальність замість видалення.", "error")
            else:
                row = conn.execute("SELECT name_ua FROM specialties WHERE code = ?", (code,)).fetchone()
                conn.execute("DELETE FROM specialties WHERE code = ?", (code,))
                conn.commit()
                flash("Спеціальність видалено.", "success")
                log_action(current_username(), f"видалив спеціальність: {code} {row['name_ua'] if row else ''}")

        elif action == 'add_field':
            field_code = (request.form.get('field_code') or '').strip()
            field_name_ua = (request.form.get('field_name_ua') or '').strip()
            field_name_en = (request.form.get('field_name_en') or '').strip() or None
            if not field_code or not field_name_ua:
                flash("Заповніть код і назву галузі знань.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO knowledge_fields (code, name_ua, name_en) VALUES (?, ?, ?)",
                        (field_code, field_name_ua, field_name_en)
                    )
                    conn.commit()
                    flash(f"Галузь знань «{field_code} {field_name_ua}» додано.", "success")
                    log_action(current_username(), f"додав галузь знань: {field_code} {field_name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Галузь знань з кодом «{field_code}» вже існує.", "error")

        elif action == 'delete_field':
            field_code = request.form.get('field_code')
            in_use = conn.execute(
                "SELECT COUNT(*) AS c FROM specialties WHERE knowledge_field_code = ?", (field_code,)
            ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - на цю галузь знань посилається {in_use} спеціальність(і). "
                      f"Спершу перенесіть або видаліть їх.", "error")
            else:
                row = conn.execute("SELECT name_ua FROM knowledge_fields WHERE code = ?", (field_code,)).fetchone()
                conn.execute("DELETE FROM knowledge_fields WHERE code = ?", (field_code,))
                conn.commit()
                flash("Галузь знань видалено.", "success")
                log_action(current_username(), f"видалив галузь знань: {field_code} {row['name_ua'] if row else ''}")

    fields = conn.execute("""
        SELECT k.code, k.name_ua, k.name_en,
               (SELECT COUNT(*) FROM specialties s WHERE s.knowledge_field_code = k.code) AS specialties_count
        FROM knowledge_fields k ORDER BY k.code
    """).fetchall()
    specialties = conn.execute("""
        SELECT s.code, s.name_ua, s.short_name, s.name_en, s.is_active, s.is_custom, k.code AS field_code, k.name_ua AS field_name,
               (SELECT COUNT(*) FROM groups g WHERE g.specialty_code = s.code) AS groups_count
        FROM specialties s JOIN knowledge_fields k ON k.code = s.knowledge_field_code
        ORDER BY k.code, substr(s.code,1,1), CAST(substr(s.code,2) AS INTEGER)
    """).fetchall()
    conn.close()

    return render_template("manage_specialties.html", fields=fields, specialties=specialties)


@admin_bp.route('/admin/manage_degree_levels', methods=['GET', 'POST'])
@permission_required('manage_degree_levels')
def manage_degree_levels():
    """
    Каталог ступенів (Бакалавр/Магістр тощо) - раніше жорстко прописані
    в коді двома варіантами, тепер редагований список: вмикання/
    вимикання актуальності, додавання нових ступенів, редагування
    назв. Той самий принцип, що й у "Спеціальностях".
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'toggle_active':
            level_id = request.form.get('id')
            conn.execute("UPDATE degree_levels SET is_active = 1 - is_active WHERE id = ?", (level_id,))
            conn.commit()
            row = conn.execute("SELECT name_ua, is_active FROM degree_levels WHERE id = ?", (level_id,)).fetchone()
            if row:
                log_action(
                    current_username(),
                    f"{'увімкнув' if row['is_active'] else 'вимкнув'} ступінь: {row['name_ua']}"
                )

        elif action == 'add':
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву ступеня.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO degree_levels (name_ua, name_en, is_active) VALUES (?, ?, 1)",
                        (name_ua, name_en)
                    )
                    conn.commit()
                    flash(f"Ступінь «{name_ua}» додано.", "success")
                    log_action(current_username(), f"додав ступінь: {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Ступінь «{name_ua}» вже існує.", "error")

        elif action == 'edit':
            level_id = request.form.get('id')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву ступеня.", "error")
            else:
                try:
                    conn.execute(
                        "UPDATE degree_levels SET name_ua = ?, name_en = ? WHERE id = ?",
                        (name_ua, name_en, level_id)
                    )
                    conn.commit()
                    flash("Ступінь оновлено.", "success")
                    log_action(current_username(), f"редагував ступінь: {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Ступінь «{name_ua}» вже існує.", "error")

        elif action == 'delete':
            level_id = request.form.get('id')
            row = conn.execute("SELECT name_ua FROM degree_levels WHERE id = ?", (level_id,)).fetchone()
            in_use = conn.execute(
                "SELECT COUNT(*) AS c FROM groups WHERE degree_level = ?", (row['name_ua'] if row else '',)
            ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - є {in_use} груп(и) з цим ступенем. Вимкніть актуальність замість видалення.", "error")
            else:
                conn.execute("DELETE FROM degree_levels WHERE id = ?", (level_id,))
                conn.commit()
                flash("Ступінь видалено.", "success")
                log_action(current_username(), f"видалив ступінь: {row['name_ua'] if row else ''}")

    degree_levels = conn.execute("""
        SELECT d.id, d.name_ua, d.name_en, d.is_active,
               (SELECT COUNT(*) FROM groups g WHERE g.degree_level = d.name_ua) AS groups_count
        FROM degree_levels d ORDER BY d.id
    """).fetchall()
    conn.close()

    return render_template("manage_degree_levels.html", degree_levels=degree_levels)


@admin_bp.route('/admin/manage_licenses', methods=['GET', 'POST'])
@permission_required('manage_licenses')
def manage_licenses():
    """
    Каталог ліцензій, під якими студенти вступають до закладу
    (Львівська, Київська, Польська тощо) - властивість СТУДЕНТА, а не
    групи, оскільки в одній групі можуть навчатись студенти з різних
    ліцензій. Той самий принцип CRUD, що й у "Ступенях"/"Спеціальностях".
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'toggle_active':
            license_id = request.form.get('id')
            conn.execute("UPDATE institution_licenses SET is_active = 1 - is_active WHERE id = ?", (license_id,))
            conn.commit()
            row = conn.execute("SELECT short_name_ua, is_active FROM institution_licenses WHERE id = ?", (license_id,)).fetchone()
            if row:
                log_action(current_username(), f"{'увімкнув' if row['is_active'] else 'вимкнув'} ліцензію: {row['short_name_ua']}")

        elif action == 'add':
            name_ua = (request.form.get('name_ua') or '').strip()
            short_name_ua = (request.form.get('short_name_ua') or '').strip() or None
            name_en = (request.form.get('name_en') or '').strip() or None
            short_name_en = (request.form.get('short_name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть повну назву ліцензії.", "error")
            else:
                conn.execute(
                    "INSERT INTO institution_licenses (name_ua, short_name_ua, name_en, short_name_en, is_active) VALUES (?, ?, ?, ?, 1)",
                    (name_ua, short_name_ua, name_en, short_name_en)
                )
                conn.commit()
                flash(f"Ліцензію «{short_name_ua or name_ua}» додано.", "success")
                log_action(current_username(), f"додав ліцензію: {short_name_ua or name_ua}")

        elif action == 'edit':
            license_id = request.form.get('id')
            name_ua = (request.form.get('name_ua') or '').strip()
            short_name_ua = (request.form.get('short_name_ua') or '').strip() or None
            name_en = (request.form.get('name_en') or '').strip() or None
            short_name_en = (request.form.get('short_name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть повну назву ліцензії.", "error")
            else:
                conn.execute(
                    "UPDATE institution_licenses SET name_ua = ?, short_name_ua = ?, name_en = ?, short_name_en = ? WHERE id = ?",
                    (name_ua, short_name_ua, name_en, short_name_en, license_id)
                )
                conn.commit()
                flash("Ліцензію оновлено.", "success")
                log_action(current_username(), f"редагував ліцензію (ID {license_id})")

        elif action == 'delete':
            license_id = request.form.get('id')
            row = conn.execute("SELECT short_name_ua, name_ua FROM institution_licenses WHERE id = ?", (license_id,)).fetchone()
            in_use = conn.execute(
                "SELECT COUNT(*) AS c FROM students WHERE license_id = ?", (license_id,)
            ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - {in_use} студент(ів) мають цю ліцензію. Вимкніть актуальність замість видалення.", "error")
            else:
                conn.execute("DELETE FROM institution_licenses WHERE id = ?", (license_id,))
                conn.commit()
                flash("Ліцензію видалено.", "success")
                log_action(current_username(), f"видалив ліцензію: {row['short_name_ua'] or row['name_ua'] if row else ''}")

    licenses = conn.execute("""
        SELECT l.id, l.name_ua, l.short_name_ua, l.name_en, l.short_name_en, l.is_active,
               (SELECT COUNT(*) FROM students s WHERE s.license_id = l.id) AS students_count
        FROM institution_licenses l ORDER BY l.id
    """).fetchall()
    conn.close()

    return render_template("manage_licenses.html", licenses=licenses)


@admin_bp.route('/admin/license_transfer', methods=['GET', 'POST'])
@permission_required('manage_license_transfer')
def license_transfer():
    """
    Крок 1 -> 2 переведення студентів між ліцензіями. Джерело студентів
    для наказу - з активних груп (позначили групу цілком) і/або окремі
    студенти з різних груп (позначили напряму) - обидва джерела
    зливаються в один спільний список на кроці перегляду.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        target_license_id = request.form.get('target_license_id')
        group_ids = request.form.getlist('group_ids')
        student_ids = request.form.getlist('student_ids')

        if not target_license_id:
            flash("Оберіть ліцензію призначення.", "error")
            conn.close()
            return redirect(url_for('admin.license_transfer'))
        if not group_ids and not student_ids:
            flash("Оберіть хоча б одну групу або одного студента.", "error")
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        target_license = conn.execute("SELECT id, short_name_ua, name_ua FROM institution_licenses WHERE id = ?", (target_license_id,)).fetchone()
        if not target_license:
            flash("Ліцензію призначення не знайдено.", "error")
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        # Збираємо студентів з обох джерел в один набір (без дублів,
        # якщо той самий студент прийшов і через групу, і напряму)
        student_id_set = set(int(x) for x in student_ids)
        for group_id in group_ids:
            rows = conn.execute(
                "SELECT id FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0", (group_id,)
            ).fetchall()
            student_id_set.update(r['id'] for r in rows)

        if not student_id_set:
            flash("У обраних групах немає активних студентів.", "error")
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        placeholders = ','.join('?' for _ in student_id_set)
        students = conn.execute(f"""
            SELECT s.id, TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS full_name,
                   g.name AS group_name, l.short_name_ua AS current_license_name
            FROM students s
            LEFT JOIN groups g ON g.id = s.group_id
            LEFT JOIN institution_licenses l ON l.id = s.license_id
            WHERE s.id IN ({placeholders})
            ORDER BY s.last_name_UA COLLATE UKRAINIAN
        """, list(student_id_set)).fetchall()

        conn.close()
        return render_template("license_transfer_review.html", students=students, target_license=target_license)

    licenses = conn.execute("SELECT id, name_ua, short_name_ua FROM institution_licenses WHERE is_active = 1 ORDER BY id").fetchall()
    groups = conn.execute("""
        SELECT id, name,
               (SELECT COUNT(*) FROM students s WHERE s.group_id = groups.id AND COALESCE(s.archived, 0) = 0) AS student_count
        FROM groups WHERE archived = FALSE ORDER BY name
    """).fetchall()
    all_students = conn.execute("""
        SELECT s.id, TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS full_name,
               g.name AS group_name, l.short_name_ua AS current_license_name
        FROM students s
        LEFT JOIN groups g ON g.id = s.group_id
        LEFT JOIN institution_licenses l ON l.id = s.license_id
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA COLLATE UKRAINIAN
    """).fetchall()
    conn.close()

    return render_template("license_transfer_select.html", licenses=licenses, groups=groups, all_students=all_students)


@admin_bp.route('/admin/license_transfer/confirm', methods=['POST'])
@permission_required('manage_license_transfer')
def license_transfer_confirm():
    """Остаточне підтвердження наказу про переведення між ліцензіями."""
    order_number = (request.form.get('order_number') or '').strip()
    order_date = request.form.get('order_date')
    target_license_id = request.form.get('target_license_id')
    student_ids = [int(x) for x in request.form.getlist('student_ids')]

    if not order_number or not order_date or not target_license_id or not student_ids:
        flash("Заповніть номер, дату наказу і переконайтесь, що обрано студентів.", "error")
        return redirect(url_for('admin.license_transfer'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    scan_rel_path = None
    scan_file = request.files.get('scan_file')
    if scan_file and scan_file.filename:
        os.makedirs(os.path.join('static', 'uploads', 'license_transfer_orders'), exist_ok=True)
        ext = os.path.splitext(scan_file.filename)[1]
        safe_name = f"{uuid.uuid4().hex}{ext}"
        scan_file.save(os.path.join('static', 'uploads', 'license_transfer_orders', safe_name))
        scan_rel_path = f"uploads/license_transfer_orders/{safe_name}"

    cur = conn.execute(
        "INSERT INTO license_transfer_orders (order_number, order_date, scan_file, created_by) VALUES (?, ?, ?, ?)",
        (order_number, order_date, scan_rel_path, current_username())
    )
    order_id = cur.lastrowid

    for student_id in student_ids:
        prev = conn.execute("SELECT license_id FROM students WHERE id = ?", (student_id,)).fetchone()
        previous_license_id = prev['license_id'] if prev else None
        conn.execute(
            "INSERT INTO license_transfer_order_students (order_id, student_id, previous_license_id, new_license_id) VALUES (?, ?, ?, ?)",
            (order_id, student_id, previous_license_id, target_license_id)
        )
        conn.execute("UPDATE students SET license_id = ? WHERE id = ?", (target_license_id, student_id))

    conn.commit()
    conn.close()

    log_action(
        current_username(),
        f"наказ про переведення між ліцензіями №{order_number} від {order_date}",
        details=f"студентів: {len(student_ids)}, нова ліцензія ID {target_license_id}"
    )
    flash(f"Переведення між ліцензіями оформлено. Студентів: {len(student_ids)}.", "success")
    return redirect(url_for('admin.manage_licenses'))


@admin_bp.route('/admin/license_report')
@permission_required('license_report')
def license_report():
    """
    Звіт по ліцензіях: скільки студентів кожної групи належать до
    якої ліцензії (Ф=Філія/Львів, К=Київ, П=Польща тощо - будь-які
    активні ліцензії з каталогу). Два режими: "Список" (як у
    друкованих списках вступників - курс/спеціальність/студенти з
    позначкою ліцензії) і "Підсумки" (зведена таблиця з кількостями,
    як власноруч ведена Excel-таблиця). Перемикач архівних груп - той
    самий принцип, що й в Аналітиці.
    """
    include_archived = request.args.get('include_archived') == '1'
    view = request.args.get('view', 'list')

    conn = get_db()
    conn.row_factory = sqlite3.Row

    group_cond = "" if include_archived else "WHERE COALESCE(g.archived,0)=0"
    groups = conn.execute(f"""
        SELECT g.id, g.name, g.course, g.specialty, g.degree_level, g.study_form, g.archived, g.program_credits
        FROM groups g {group_cond}
        ORDER BY g.course, g.specialty COLLATE UKRAINIAN, g.name COLLATE UKRAINIAN
    """).fetchall()

    licenses = conn.execute(
        "SELECT id, short_name_ua FROM institution_licenses WHERE is_active = 1 ORDER BY id"
    ).fetchall()
    license_names = [l['short_name_ua'] for l in licenses]

    # Порядок ступенів - за їхнім id в каталозі degree_levels, той
    # самий принцип, що й на дошці "Курси" й в "Архіві груп".
    degree_order = {
        row['name_ua']: row['id']
        for row in conn.execute("SELECT id, name_ua FROM degree_levels").fetchall()
    }

    groups_with_students = []
    summary_rows = []
    grand_totals = {name: 0 for name in license_names}
    grand_totals['—'] = 0  # без ліцензії
    grand_total_all = 0

    # Підсумки по (ступінь + форма навчання) - як блоки "Бакалавр
    # денна"/"Бакалавр заочна" тощо у власноруч веденій таблиці
    breakdown = {}
    # Підсумки просто по курсах, незалежно від спеціальності/ступеня
    course_breakdown = {}

    for g in groups:
        students = conn.execute("""
            SELECT s.id, TRIM(s.last_name_UA || ' ' || s.first_name_UA) AS full_name,
                   l.short_name_ua AS license_short_name, s.archived, s.program_credits_override
            FROM students s LEFT JOIN institution_licenses l ON l.id = s.license_id
            WHERE s.group_id = ? {archived_cond}
            ORDER BY s.last_name_UA COLLATE UKRAINIAN
        """.format(archived_cond="" if include_archived else "AND COALESCE(s.archived,0)=0"), (g['id'],)).fetchall()

        counts = {name: 0 for name in license_names}
        counts['—'] = 0
        for s in students:
            key = s['license_short_name'] or '—'
            counts[key] = counts.get(key, 0) + 1
            grand_totals[key] = grand_totals.get(key, 0) + 1
            grand_total_all += 1

        groups_with_students.append({'group': g, 'students': students, 'counts': {k: v for k, v in counts.items() if v}})

        group_total = sum(counts.values())
        summary_rows.append({'group': g, 'counts': counts, 'total': group_total})

        breakdown_key = (g['degree_level'] or '—', g['study_form'] or '—')
        if breakdown_key not in breakdown:
            breakdown[breakdown_key] = {name: 0 for name in license_names}
            breakdown[breakdown_key]['—'] = 0
        for name, cnt in counts.items():
            breakdown[breakdown_key][name] = breakdown[breakdown_key].get(name, 0) + cnt

        course_key = g['course']
        if course_key not in course_breakdown:
            course_breakdown[course_key] = {name: 0 for name in license_names}
            course_breakdown[course_key]['—'] = 0
        for name, cnt in counts.items():
            course_breakdown[course_key][name] = course_breakdown[course_key].get(name, 0) + cnt

    conn.close()

    # Групуємо для компактного вигляду в режимі "Список": ступінь -> список груп
    groups_by_degree_raw = {}
    for item in groups_with_students:
        groups_by_degree_raw.setdefault(item['group']['degree_level'] or 'Без ступеня', []).append(item)
    groups_by_degree = {
        degree_level: groups_by_degree_raw[degree_level]
        for degree_level in sorted(groups_by_degree_raw.keys(), key=lambda d: (degree_order.get(d, 999), d))
    }

    breakdown_rows = [
        {'degree_level': k[0], 'study_form': k[1], 'counts': v, 'total': sum(v.values())}
        for k, v in sorted(breakdown.items())
    ]
    course_breakdown_rows = [
        {'course': k, 'counts': v, 'total': sum(v.values())}
        for k, v in sorted(course_breakdown.items())
    ]

    return render_template(
        "license_report.html",
        view=view,
        include_archived=include_archived,
        license_names=license_names,
        groups_with_students=groups_with_students,
        groups_by_degree=groups_by_degree,
        summary_rows=summary_rows,
        course_breakdown_rows=course_breakdown_rows,
        breakdown_rows=breakdown_rows,
        grand_totals=grand_totals,
        grand_total_all=grand_total_all,
    )


@admin_bp.route('/admin/license_report/export_word')
@permission_required('license_report')
def license_report_export_word():
    """
    Вивантажує звіт по ліцензіях у .docx у форматі: "ВСТУП {рік}" ->
    "{курс} курс {форма навчання}" -> реальна назва групи -> нумерована
    таблиця "№ / Прізвище І.П. / Ліцензія" (ліцензія - колонка в
    таблиці, а не заголовок блоку). Денна й заочна форми - окремі
    файли (study_form обов'язковий параметр запиту).
    """
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from collections import OrderedDict

    include_archived = request.args.get('include_archived') == '1'
    study_form = request.args.get('study_form')  # 'Денна' або 'Заочна' - обов'язково, окремий файл на форму
    if study_form not in ('Денна', 'Заочна'):
        flash("Оберіть форму навчання (Денна чи Заочна) для вивантаження.", "error")
        return redirect(url_for('admin.license_report'))


    conn = get_db()
    conn.row_factory = sqlite3.Row
    group_cond = "AND COALESCE(g.archived,0)=0" if not include_archived else ""
    student_cond = "AND COALESCE(s.archived,0)=0" if not include_archived else ""

    rows = conn.execute(f"""
        SELECT g.start_year, g.course, g.name AS group_name, g.degree_level,
               g.specialty_code, sp.name_ua AS specialty_name_ua,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA || ' ' || COALESCE(s.middle_name_UA, '')) AS full_name,
               l.short_name_ua AS license_short_name
        FROM groups g
        JOIN students s ON s.group_id = g.id
        LEFT JOIN specialties sp ON sp.code = g.specialty_code
        LEFT JOIN institution_licenses l ON l.id = s.license_id
        WHERE g.study_form = ? {group_cond} {student_cond}
        ORDER BY g.start_year DESC, g.course, g.name COLLATE UKRAINIAN, s.last_name_UA COLLATE UKRAINIAN
    """, (study_form,)).fetchall()
    conn.close()

    # Групуємо: ступінь (спершу Бакалавр, потім Магістр, решта - після)
    # -> рік вступу -> курс -> реальна назва групи -> студенти (ім'я +
    # своя ліцензія як окрема колонка, а не заголовок блоку).
    DEGREE_ORDER = {'Бакалавр': 0, 'Магістр': 1}
    DEGREE_LABELS = {'Бакалавр': 'БАКАЛАВРИ', 'Магістр': 'МАГІСТРИ'}

    grouped = OrderedDict()
    for r in rows:
        degree_key = r['degree_level'] or 'Інше'
        section_key = (r['start_year'], r['course'])
        group_bucket = grouped.setdefault(degree_key, OrderedDict()).setdefault(section_key, OrderedDict()).setdefault(
            r['group_name'], {'specialty_code': r['specialty_code'], 'specialty_name': r['specialty_name_ua'], 'students': []}
        )
        group_bucket['students'].append((r['full_name'], r['license_short_name'] or 'без ліцензії'))

    # Сортуємо ключі ступенів: Бакалавр, потім Магістр, потім усе інше
    # за абеткою - самі роки/курси всередині зберігають порядок з SQL.
    ordered_degrees = sorted(grouped.keys(), key=lambda d: DEGREE_ORDER.get(d, 99))

    doc = Document()
    for section in doc.sections:
        section.left_margin = Cm(2)
        section.right_margin = Cm(1)

    study_form_text = {'Денна': 'денна форма навчання', 'Заочна': 'заочна форма навчання'}.get(study_form, study_form)

    for degree_key in ordered_degrees:
        degree_label = DEGREE_LABELS.get(degree_key, degree_key.upper())
        years = grouped[degree_key]

        for (start_year, course), groups_dict in years.items():
            p = doc.add_paragraph()
            run = p.add_run(f"ВСТУП {start_year}")
            run.bold = True
            run.font.size = Pt(16)

            p_degree = doc.add_paragraph()
            run_degree = p_degree.add_run(degree_label)
            run_degree.bold = True
            run_degree.font.size = Pt(16)

            p2 = doc.add_paragraph()
            run2 = p2.add_run(f"{course} курс {study_form_text}")
            run2.bold = True
            run2.font.size = Pt(16)

            for group_name, group_info in groups_dict.items():
                students = group_info['students']
                specialty_code = group_info['specialty_code'] or ''
                specialty_name = group_info['specialty_name'] or ''
                header_text = group_name
                if specialty_code or specialty_name:
                    header_text += f" ({specialty_code}, {specialty_name})"

                p3 = doc.add_paragraph()
                run3 = p3.add_run(header_text)
                run3.bold = True
                run3.font.size = Pt(14)

                table = doc.add_table(rows=1, cols=3)
                table.style = 'Table Grid'
                hdr = table.rows[0].cells
                hdr[0].text = 'П/н'
                hdr[1].text = 'Прізвище І.П.'
                hdr[2].text = 'Ліцензія'
                for cell in hdr:
                    for para in cell.paragraphs:
                        for r in para.runs:
                            r.bold = True

                table.columns[0].width = Cm(1.5)
                table.columns[1].width = Cm(9)
                table.columns[2].width = Cm(4.5)

                for i, (name, license_name) in enumerate(students, start=1):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(i)
                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[1].text = name
                    row_cells[2].text = license_name

                doc.add_paragraph()

            doc.add_paragraph()

    output_path = os.path.join('static', 'uploads', f"license_report_{uuid.uuid4().hex}.docx")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)

    log_action(current_username(), f"вивантажив звіт по ліцензіях у Word ({study_form})")
    download_name = f"Списки_студентів_{'денна' if study_form == 'Денна' else 'заочна'}.docx"
    return send_file(output_path, as_attachment=True, download_name=download_name)


@admin_bp.route('/admin/manage_educational_programs', methods=['GET', 'POST'])
@permission_required('manage_educational_programs')
def manage_educational_programs():
    """
    Каталог освітніх програм: для кожної трійки "спеціальність + рік
    вступу + ступінь" - своя назва програми (укр./англ.). Ступінь
    важливий окремо від року: та сама спеціальність і рік вступу
    можуть мати зовсім різну програму на рівні Молодшого спеціаліста/
    Бакалавра/Магістра. Використовується на сторінці "Групи" для
    автопідстановки замість вільного вводу.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            specialty_code = request.form.get('specialty_code')
            start_year = request.form.get('start_year')
            degree_level = request.form.get('degree_level')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None

            if not specialty_code or not start_year or not degree_level or not name_ua:
                flash("Оберіть спеціальність, ступінь, вкажіть рік і назву програми.", "error")
            elif not start_year.isdigit():
                flash("Рік має бути числом.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO educational_programs (specialty_code, start_year, degree_level, name_ua, name_en) VALUES (?, ?, ?, ?, ?)",
                        (specialty_code, int(start_year), degree_level, name_ua, name_en)
                    )
                    conn.commit()
                    flash(f"Освітню програму для {specialty_code} ({degree_level}, {start_year}) додано.", "success")
                    log_action(current_username(), f"додав освітню програму: {specialty_code} {degree_level} {start_year} - {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Для спеціальності {specialty_code}, ступеня «{degree_level}» і року {start_year} програма вже існує - відредагуйте наявну.", "error")

        elif action == 'edit':
            program_id = request.form.get('id')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву програми.", "error")
            else:
                conn.execute(
                    "UPDATE educational_programs SET name_ua = ?, name_en = ? WHERE id = ?",
                    (name_ua, name_en, program_id)
                )
                conn.commit()
                flash("Освітню програму оновлено.", "success")
                log_action(current_username(), f"редагував освітню програму ID {program_id}: {name_ua}")

        elif action == 'delete':
            program_id = request.form.get('id')
            row = conn.execute(
                "SELECT specialty_code, start_year, degree_level, name_ua FROM educational_programs WHERE id = ?", (program_id,)
            ).fetchone()
            in_use = 0
            if row:
                in_use = conn.execute(
                    "SELECT COUNT(*) AS c FROM groups WHERE specialty_code = ? AND start_year = ? AND degree_level = ?",
                    (row['specialty_code'], row['start_year'], row['degree_level'])
                ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - є {in_use} груп(и), що використовують цю програму.", "error")
            else:
                conn.execute("DELETE FROM educational_programs WHERE id = ?", (program_id,))
                conn.commit()
                flash("Освітню програму видалено.", "success")
                log_action(current_username(), f"видалив освітню програму: {row['name_ua'] if row else ''}")

    specialties = conn.execute("""
        SELECT s.code, s.name_ua, k.name_ua AS field_name
        FROM specialties s JOIN knowledge_fields k ON k.code = s.knowledge_field_code
        WHERE s.is_active = 1
        ORDER BY substr(s.code,1,1), CAST(substr(s.code,2) AS INTEGER)
    """).fetchall()

    degree_levels = conn.execute(
        "SELECT name_ua FROM degree_levels WHERE is_active = 1 ORDER BY id"
    ).fetchall()

    programs = conn.execute("""
        SELECT p.id, p.specialty_code, p.start_year, p.degree_level, p.name_ua, p.name_en,
               s.name_ua AS specialty_name,
               (SELECT COUNT(*) FROM groups g
                WHERE g.specialty_code = p.specialty_code AND g.start_year = p.start_year AND g.degree_level = p.degree_level) AS groups_count
        FROM educational_programs p JOIN specialties s ON s.code = p.specialty_code
        ORDER BY p.specialty_code, p.start_year DESC, p.degree_level COLLATE UKRAINIAN
    """).fetchall()

    # Групуємо програми по спеціальності для зручного відображення
    programs_by_specialty = {}
    for p in programs:
        programs_by_specialty.setdefault(p['specialty_code'], []).append(p)

    conn.close()

    return render_template(
        "manage_educational_programs.html",
        specialties=specialties,
        degree_levels=degree_levels,
        programs_by_specialty=programs_by_specialty,
    )


@admin_bp.route('/admin/manage_qualification_names', methods=['GET', 'POST'])
@permission_required('manage_qualification_names')
def manage_qualification_names():
    """
    Каталог назв кваліфікації: для кожної трійки "спеціальність + рік
    вступу + ступінь" - своя назва кваліфікації (укр./англ.). Той самий
    принцип, що й "Освітня програма" - формулювання може відрізнятись і
    рік від року, і від ступеня.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            specialty_code = request.form.get('specialty_code')
            start_year = request.form.get('start_year')
            degree_level = request.form.get('degree_level')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None

            if not specialty_code or not start_year or not degree_level or not name_ua:
                flash("Оберіть спеціальність, ступінь, вкажіть рік і назву кваліфікації.", "error")
            elif not start_year.isdigit():
                flash("Рік має бути числом.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO qualification_names (specialty_code, start_year, degree_level, name_ua, name_en) VALUES (?, ?, ?, ?, ?)",
                        (specialty_code, int(start_year), degree_level, name_ua, name_en)
                    )
                    conn.commit()
                    flash(f"Назву кваліфікації для {specialty_code} ({degree_level}, {start_year}) додано.", "success")
                    log_action(current_username(), f"додав назву кваліфікації: {specialty_code} {degree_level} {start_year} - {name_ua}")
                except sqlite3.IntegrityError:
                    flash(f"Для спеціальності {specialty_code}, ступеня «{degree_level}» і року {start_year} назва кваліфікації вже існує - відредагуйте наявну.", "error")

        elif action == 'edit':
            qual_id = request.form.get('id')
            name_ua = (request.form.get('name_ua') or '').strip()
            name_en = (request.form.get('name_en') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву кваліфікації.", "error")
            else:
                conn.execute(
                    "UPDATE qualification_names SET name_ua = ?, name_en = ? WHERE id = ?",
                    (name_ua, name_en, qual_id)
                )
                conn.commit()
                flash("Назву кваліфікації оновлено.", "success")
                log_action(current_username(), f"редагував назву кваліфікації ID {qual_id}: {name_ua}")

        elif action == 'delete':
            qual_id = request.form.get('id')
            row = conn.execute(
                "SELECT specialty_code, start_year, degree_level, name_ua FROM qualification_names WHERE id = ?", (qual_id,)
            ).fetchone()
            in_use = 0
            if row:
                in_use = conn.execute(
                    "SELECT COUNT(*) AS c FROM groups WHERE specialty_code = ? AND start_year = ? AND degree_level = ?",
                    (row['specialty_code'], row['start_year'], row['degree_level'])
                ).fetchone()['c']
            if in_use > 0:
                flash(f"Неможливо видалити - є {in_use} груп(и), що використовують цю назву кваліфікації.", "error")
            else:
                conn.execute("DELETE FROM qualification_names WHERE id = ?", (qual_id,))
                conn.commit()
                flash("Назву кваліфікації видалено.", "success")
                log_action(current_username(), f"видалив назву кваліфікації: {row['name_ua'] if row else ''}")

    specialties = conn.execute("""
        SELECT s.code, s.name_ua, k.name_ua AS field_name
        FROM specialties s JOIN knowledge_fields k ON k.code = s.knowledge_field_code
        WHERE s.is_active = 1
        ORDER BY substr(s.code,1,1), CAST(substr(s.code,2) AS INTEGER)
    """).fetchall()

    degree_levels = conn.execute(
        "SELECT name_ua FROM degree_levels WHERE is_active = 1 ORDER BY id"
    ).fetchall()

    qualifications = conn.execute("""
        SELECT q.id, q.specialty_code, q.start_year, q.degree_level, q.name_ua, q.name_en,
               s.name_ua AS specialty_name,
               (SELECT COUNT(*) FROM groups g
                WHERE g.specialty_code = q.specialty_code AND g.start_year = q.start_year AND g.degree_level = q.degree_level) AS groups_count
        FROM qualification_names q JOIN specialties s ON s.code = q.specialty_code
        ORDER BY q.specialty_code, q.start_year DESC, q.degree_level COLLATE UKRAINIAN
    """).fetchall()

    qualifications_by_specialty = {}
    for q in qualifications:
        qualifications_by_specialty.setdefault(q['specialty_code'], []).append(q)

    conn.close()

    return render_template(
        "manage_qualification_names.html",
        specialties=specialties,
        degree_levels=degree_levels,
        qualifications_by_specialty=qualifications_by_specialty,
    )
