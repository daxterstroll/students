"""
routes/admin/catalogs.py
========================
Довідники: спеціальності, ступені, освітні програми, назви
кваліфікацій, акредитації, ліцензії (+ переведення між ліцензіями,
звіт по ліцензіях), дипломи, масове призначення періодів навчання.

ПІСЛЯ ПЕРЕХОДУ НА ENROLLMENT-BASED ЛІЦЕНЗІЇ:
  - license_id зберігається в student_enrollments (source of truth);
  - students.license_id — кеш ліцензії primary enrollment'а;
  - усі звіти/фільтри по ліцензіях ітерують по enrollments, а не по
    students.group_id (щоб не втрачати паралельних студентів).
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
    """Масове призначення періоду навчання (філія, дати, група) одразу
    декільком студентам — обраній групі цілком та/або довільному списку
    студентів.

    Група розуміється як «усі студенти, які мають активний enrollment у
    цій групі» (або legacy s.group_id, якщо enrollments немає) — тобто
    паралельні студенти з інших primary-груп теж потрапляють у вибірку
    «усій групі».
    """
    db = get_db()
    cursor = db.cursor()

    # -------------------- POST --------------------
    if request.method == 'POST':
        mode = request.form.get('mode')
        group_id = request.form.get('group_id', type=int)
        student_ids_raw = request.form.get('student_ids_hidden', '')
        extra_student_ids = [int(x) for x in student_ids_raw.split(',') if x.strip().isdigit()]

        target_student_ids = set(extra_student_ids)

        if mode == 'whole_group' and group_id:
            # «Усій групі» = primary + паралельні студенти цієї групи
            cursor.execute("""
                SELECT DISTINCT s.id FROM students s
                WHERE COALESCE(s.archived, 0) = 0
                  AND (
                      s.group_id = ?
                      OR EXISTS (
                          SELECT 1 FROM student_enrollments
                          WHERE student_id = s.id AND group_id = ? AND status = 'active'
                      )
                  )
            """, (group_id, group_id))
            target_student_ids.update(r['id'] for r in cursor.fetchall())
        # mode == 'some_from_group' → беремо ЛИШЕ явно вибраних у пикері
        # mode == 'any_students'     → те саме

        if not target_student_ids:
            flash('Не обрано жодного студента', 'danger')
            return redirect(url_for('admin.manage_study_periods_bulk_assign'))

        filiya = (request.form.get('filiya') or '').strip()
        filiya_en = (request.form.get('filiya_en') or '').strip() or None
        group_name = (request.form.get('group_name') or '').strip() or None
        start_date = (request.form.get('start_date') or '').strip() or None
        end_date = (request.form.get('end_date') or '').strip() or None
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
                next_order = cursor.execute(
                    "SELECT COALESCE(MAX(period_order), 0) + 1 FROM student_study_periods WHERE student_id = ?",
                    (student_id,)
                ).fetchone()[0]
                cursor.execute("""
                    INSERT INTO student_study_periods
                        (student_id, filiya, filiya_en, group_name, start_date, end_date, period_order, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (student_id, filiya, filiya_en, group_name, start_date, end_date, next_order, note))
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

    if mode == 'whole_group' and group_id:
        cursor.execute("""
            SELECT DISTINCT s.id FROM students s
            WHERE COALESCE(s.archived, 0) = 0
              AND (
                  s.group_id = ?
                  OR EXISTS (
                      SELECT 1 FROM student_enrollments
                      WHERE student_id = s.id AND group_id = ? AND status = 'active'
                  )
              )
        """, (group_id, group_id))
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
    """Перегляд і масове збереження номерів диплома/додатка для всіх
    студентів обраної групи.

    Номери зберігаються ОКРЕМО ДЛЯ КОЖНОЇ ГРУПИ студента (ключ —
    пара student_id + group_id). Тобто якщо студент паралельно вчиться
    у двох групах і випускається з обох, у кожної групи може бути свій
    номер диплома. Якщо диплом один (студент паралельно вчився у двох
    групах однієї програми) — адміністратор вписує однакові номери на
    сторінках обох груп.

    «Студенти групи» = primary (s.group_id) + паралельні (active
    enrollment у цій групі). Номери показуються САМЕ для обраної групи
    (d.group_id = selected_group), а не «будь-який номер студента».
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    if request.method == "POST":
        group_id = request.form.get("group_id")
        if not group_id:
            flash("Не вказано групу", "error")
            return redirect(url_for('admin.manage_diplomas'))
        group_id = int(group_id)

        cursor.execute("""
            SELECT DISTINCT s.id FROM students s
            WHERE COALESCE(s.archived, 0) = 0
              AND (
                  s.group_id = ?
                  OR EXISTS (
                      SELECT 1 FROM student_enrollments
                      WHERE student_id = s.id AND group_id = ? AND status = 'active'
                  )
              )
        """, (group_id, group_id))
        students = cursor.fetchall()

        for student in students:
            student_id = student['id']
            diploma_number = request.form.get(f'diploma_number_{student_id}', '').strip()
            appendix_number = request.form.get(f'appendix_number_{student_id}', '').strip()
            if diploma_number:
                diploma_number = diploma_number.zfill(6)

            # UPSERT по (student_id, group_id) — а не по student_id,
            # як було раніше. Це дозволяє зберігати РІЗНІ номери для
            # різних груп одного й того самого студента.
            existing = cursor.execute(
                "SELECT id FROM diplomas WHERE student_id = ? AND group_id = ?",
                (student_id, group_id)
            ).fetchone()

            if existing:
                cursor.execute("""
                    UPDATE diplomas
                    SET diploma_number = ?, appendix_number = ?
                    WHERE id = ?
                """, (diploma_number, appendix_number, existing['id']))
            else:
                cursor.execute("""
                    INSERT INTO diplomas (student_id, group_id, diploma_number, appendix_number)
                    VALUES (?, ?, ?, ?)
                """, (student_id, group_id, diploma_number, appendix_number))

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
        # LEFT JOIN diplomas — именно для этой группы (d.group_id = ?).
        # Не «любой номер студента» — иначе на странице группы Б был бы
        # виден номер группы А, что сбивало бы админа.
        cursor.execute("""
            SELECT DISTINCT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
                   d.diploma_number, d.appendix_number
            FROM students s
            LEFT JOIN diplomas d
                   ON s.id = d.student_id AND d.group_id = ?
            WHERE COALESCE(s.archived, 0) = 0
              AND (
                  s.group_id = ?
                  OR EXISTS (
                      SELECT 1 FROM student_enrollments
                      WHERE student_id = s.id AND group_id = ? AND status = 'active'
                  )
              )
        """, (selected_group, selected_group, selected_group))
        students = cursor.fetchall()
        students = sort_ukrainian(
            students,
            key_func=lambda s: f"{s['last_name_UA']} {s['first_name_UA']} {s['middle_name_UA']}"
        )

    conn.close()
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
            isced_f_ua = (request.form.get('isced_f_ua') or '').strip() or None
            isced_f_en = (request.form.get('isced_f_en') or '').strip() or None
            field_code = request.form.get('knowledge_field_code')

            if not code or not name_ua or not field_code:
                flash("Заповніть код, назву і галузь знань.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO specialties (code, name_ua, short_name, name_en, isced_f_ua, isced_f_en, knowledge_field_code, is_active, is_custom) VALUES (?, ?, ?, ?, ?, ?, ?, 1, 1)",
                        (code, name_ua, short_name, name_en, isced_f_ua, isced_f_en, field_code)
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
            isced_f_ua = (request.form.get('isced_f_ua') or '').strip() or None
            isced_f_en = (request.form.get('isced_f_en') or '').strip() or None
            field_code = request.form.get('knowledge_field_code')
            if not name_ua or not field_code:
                flash("Заповніть назву і галузь знань.", "error")
            else:
                conn.execute(
                    "UPDATE specialties SET name_ua = ?, short_name = ?, name_en = ?, isced_f_ua = ?, isced_f_en = ?, knowledge_field_code = ? WHERE code = ?",
                    (name_ua, short_name, name_en, isced_f_ua, isced_f_en, field_code, code)
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
        SELECT s.code, s.name_ua, s.short_name, s.name_en, s.isced_f_ua, s.isced_f_en, s.is_active, s.is_custom,
               k.code AS field_code, k.name_ua AS field_name,
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
            name_suffix = (request.form.get('name_suffix') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву ступеня.", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO degree_levels (name_ua, name_en, is_active, name_suffix) VALUES (?, ?, 1, ?)",
                        (name_ua, name_en, name_suffix)
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
            name_suffix = (request.form.get('name_suffix') or '').strip() or None
            if not name_ua:
                flash("Заповніть назву ступеня.", "error")
            else:
                try:
                    conn.execute(
                        "UPDATE degree_levels SET name_ua = ?, name_en = ?, name_suffix = ? WHERE id = ?",
                        (name_ua, name_en, name_suffix, level_id)
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
        SELECT d.id, d.name_ua, d.name_en, d.is_active, d.name_suffix,
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
    (Львівська, Київська, Польська тощо).

    Ліцензія зберігається на ДВОХ рівнях:
      - student_enrollments.license_id — SOURCE OF TRUTH (ліцензія
        конкретної групи студента — дозволяє паралельне навчання під
        різними ліцензіями);
      - students.license_id — КЕШ ліцензії primary enrollment'а (для
        legacy-коду, який ще не читає enrollment'и).

    Усі підрахунки/перевірки нижче враховують ОБИДВА джерела через
    UNION, щоб не пропустити реальне використання ліцензії.
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

        elif action == 'set_base':
            license_id = request.form.get('id')
            row = conn.execute(
                "SELECT short_name_ua, name_ua FROM institution_licenses WHERE id = ?",
                (license_id,),
            ).fetchone()
            if row:
                conn.execute("UPDATE institution_licenses SET is_base = 0 WHERE is_base = 1")
                conn.execute("UPDATE institution_licenses SET is_base = 1 WHERE id = ?", (license_id,))
                conn.commit()
                label = row['short_name_ua'] or row['name_ua']
                flash(
                    f"Базову ліцензію змінено на «{label}». "
                    f"Дашборд, списки груп і Додаток 12 тепер за замовчуванням "
                    f"показують студентів цієї ліцензії.",
                    'success',
                )
                log_action(
                    current_username(),
                    f"призначив базову ліцензію: {label} (ID {license_id})",
                )
            else:
                flash("Ліцензію не знайдено", 'error')

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
            row = conn.execute(
                "SELECT short_name_ua, name_ua FROM institution_licenses WHERE id = ?",
                (license_id,),
            ).fetchone()
            # Скільки унікальних АКТИВНИХ студентів використовують ліцензію
            # через students.license_id або student_enrollments.license_id.
            in_use = conn.execute("""
                SELECT COUNT(DISTINCT student_id) AS c FROM (
                    SELECT id AS student_id FROM students
                    WHERE license_id = ? AND COALESCE(archived, 0) = 0
                    UNION
                    SELECT e.student_id
                    FROM student_enrollments e
                    JOIN students s ON s.id = e.student_id
                    WHERE e.license_id = ?
                      AND e.status = 'active'
                      AND COALESCE(s.archived, 0) = 0
                )
            """, (license_id, license_id)).fetchone()['c']

            if in_use > 0:
                flash(
                    f"Неможливо видалити - {in_use} студент(ів) мають цю ліцензію "
                    f"(включно з паралельним навчанням). Вимкніть актуальність "
                    f"замість видалення.",
                    "error",
                )
            else:
                conn.execute("DELETE FROM institution_licenses WHERE id = ?", (license_id,))
                conn.commit()
                flash("Ліцензію видалено.", "success")
                log_action(current_username(), f"видалив ліцензію: {row['short_name_ua'] or row['name_ua'] if row else ''}")

    # ── Список ліцензій + кількість унікальних активних студентів ──
    counts_by_license_id = {}
    for r in conn.execute("""
        SELECT license_id, COUNT(DISTINCT student_id) AS c FROM (
            SELECT id AS student_id, license_id FROM students
            WHERE license_id IS NOT NULL AND COALESCE(archived, 0) = 0
            UNION
            SELECT e.student_id, e.license_id
            FROM student_enrollments e
            JOIN students s ON s.id = e.student_id
            WHERE e.license_id IS NOT NULL
              AND e.status = 'active'
              AND COALESCE(s.archived, 0) = 0
        )
        GROUP BY license_id
    """).fetchall():
        counts_by_license_id[r['license_id']] = r['c']

    licenses = []
    for r in conn.execute("""
        SELECT l.id, l.name_ua, l.short_name_ua, l.name_en, l.short_name_en,
               l.is_active, l.is_base
        FROM institution_licenses l
        ORDER BY l.is_base DESC, l.id
    """).fetchall():
        d = dict(r)
        d['students_count'] = counts_by_license_id.get(d['id'], 0)
        licenses.append(d)

    conn.close()

    return render_template("manage_licenses.html", licenses=licenses)


@admin_bp.route('/admin/license_transfer', methods=['GET', 'POST'])
@permission_required('manage_license_transfer')
def license_transfer():
    """Перевод студентов между лицензиями (с выбором scope)."""
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        target_license_id = request.form.get('target_license_id', type=int)
        group_ids = [int(x) for x in request.form.getlist('group_ids') if x]
        student_ids = [int(x) for x in request.form.getlist('student_ids') if x]
        transfer_scope = request.form.get('transfer_scope') or 'all_enrollments'

        if not target_license_id:
            flash('Оберіть ліцензію призначення.', 'error')
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        target = conn.execute(
            "SELECT id, name_ua, short_name_ua FROM institution_licenses WHERE id=?",
            (target_license_id,)
        ).fetchone()
        if not target:
            flash('Ліцензію не знайдено.', 'error')
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        # Собираем множество студентов
        selected = set(student_ids)
        if group_ids:
            placeholders = ','.join('?' * len(group_ids))
            for r in conn.execute(f"""
                SELECT DISTINCT student_id FROM student_enrollments
                WHERE group_id IN ({placeholders}) AND status = 'active'
            """, group_ids).fetchall():
                selected.add(r['student_id'])

        if not selected:
            flash('Не обрано жодного студента.', 'warning')
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        if transfer_scope == 'enrollments_in_groups' and not group_ids:
            flash('Для «Лише навчання у вибраних групах» оберіть хоча б одну групу.', 'error')
            conn.close()
            return redirect(url_for('admin.license_transfer'))

        # Загружаем студентов для review
        placeholders = ','.join('?' * len(selected))
        students = conn.execute(f"""
            SELECT s.id,
                   TRIM(s.last_name_UA || ' ' || s.first_name_UA || ' ' ||
                        COALESCE(s.middle_name_UA, '')) AS full_name,
                   g.name AS group_name,
                   il.short_name_ua AS current_license_name
            FROM students s
            LEFT JOIN groups g ON g.id = s.group_id
            LEFT JOIN institution_licenses il ON il.id = s.license_id
            WHERE s.id IN ({placeholders})
            ORDER BY s.last_name_UA
        """, list(selected)).fetchall()

        conn.close()
        return render_template(
            'license_transfer_review.html',
            students=students,
            target_license=target,
            group_ids=group_ids,
            student_ids=list(selected),
            transfer_scope=transfer_scope,
        )

    # GET — форма
    licenses = conn.execute(
        "SELECT id, name_ua, short_name_ua FROM institution_licenses "
        "WHERE is_active = 1 ORDER BY id"
    ).fetchall()

    groups = conn.execute("""
        SELECT g.id, g.name,
               (SELECT COUNT(*) FROM student_enrollments e
                WHERE e.group_id = g.id AND e.status = 'active') AS student_count
        FROM groups g WHERE COALESCE(g.archived, 0) = 0
        ORDER BY g.name COLLATE UKRAINIAN
    """).fetchall()

    all_students = conn.execute("""
        SELECT s.id,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA || ' ' ||
                    COALESCE(s.middle_name_UA, '')) AS full_name,
               g.name AS group_name,
               il.short_name_ua AS current_license_name
        FROM students s
        LEFT JOIN groups g ON g.id = s.group_id
        LEFT JOIN institution_licenses il ON il.id = s.license_id
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA, s.first_name_UA
    """).fetchall()

    conn.close()
    return render_template(
        'license_transfer_select.html',
        licenses=licenses, groups=groups, all_students=all_students,
    )


@admin_bp.route('/admin/license_transfer/confirm', methods=['POST'])
@permission_required('manage_license_transfer')
def license_transfer_confirm():
    """Фиксирует перевод и создаёт наказ. Меняет license_id в
    student_enrollments (SOURCE OF TRUTH), затем синхронизирует
    students.license_id с новым primary enrollment."""
    from routes.orders_service import create_order

    target_license_id = request.form.get('target_license_id', type=int)
    student_ids = [int(x) for x in request.form.getlist('student_ids') if x]
    order_number = (request.form.get('order_number') or '').strip()
    order_date = (request.form.get('order_date') or '').strip()
    transfer_scope = request.form.get('transfer_scope') or 'all_enrollments'
    group_ids = [int(x) for x in request.form.getlist('group_ids') if x]

    if not target_license_id or not student_ids or not order_number or not order_date:
        flash('Заповніть усі обов\'язкові поля.', 'error')
        return redirect(url_for('admin.license_transfer'))

    conn = get_db()
    conn.row_factory = sqlite3.Row

    try:
        # 1. Собираем «было» ДО изменений — для history и отчёта
        placeholders = ','.join('?' * len(student_ids))
        before_rows = conn.execute(f"""
            SELECT e.student_id, e.group_id, e.license_id
            FROM student_enrollments e
            WHERE e.student_id IN ({placeholders}) AND e.status = 'active'
        """, student_ids).fetchall()

        before_map = {}  # (student_id, group_id) -> previous_license_id
        for r in before_rows:
            before_map[(r['student_id'], r['group_id'])] = r['license_id']

        # 2. Применяем изменения к enrollments
        if transfer_scope == 'all_enrollments':
            conn.execute(f"""
                UPDATE student_enrollments
                SET license_id = ?
                WHERE student_id IN ({placeholders}) AND status = 'active'
            """, [target_license_id] + student_ids)
            updated_count = len(before_rows)
        else:
            if not group_ids:
                flash('Не вказано групи для «Лише навчання у вибраних групах».', 'error')
                conn.rollback()
                conn.close()
                return redirect(url_for('admin.license_transfer'))
            g_ph = ','.join('?' * len(group_ids))
            conn.execute(f"""
                UPDATE student_enrollments
                SET license_id = ?
                WHERE student_id IN ({placeholders}) AND status = 'active'
                  AND group_id IN ({g_ph})
            """, [target_license_id] + student_ids + group_ids)
            updated_count = len([1 for k in before_map if k[1] in group_ids])

        # 3. Синхронизируем кэш students.license_id с primary enrollment
        conn.execute(f"""
            UPDATE students SET license_id = (
                SELECT license_id FROM student_enrollments
                WHERE student_id = students.id AND is_primary = 1
                  AND status = 'active'
                LIMIT 1
            )
            WHERE id IN ({placeholders})
        """, student_ids)

        # 4. Создаём наказ (документ)
        items = [{
            'entity_type': 'student',
            'entity_id': sid,
            'payload': {
                'new_license_id': target_license_id,
                'transfer_scope': transfer_scope,
            },
        } for sid in student_ids]

        scan_files = request.files.getlist('scan_files')
        order_id = create_order(
            'license_transfer', order_number, order_date, items,
            created_by=current_username(), files=scan_files or None,
            upload_subdir='orders', conn=conn,
        )

        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        logger.error(f"license_transfer_confirm: {e}", exc_info=True)
        flash(f'Помилка переведення: {e}', 'error')
        return redirect(url_for('admin.license_transfer'))

    conn.close()

    log_action(
        current_username(),
        f'переведення між ліцензіями №{order_number} від {order_date}',
        details=f'scope={transfer_scope}, студентів={len(student_ids)}, '
                f'enrollment\'ів={updated_count}, order_id={order_id}',
    )
    flash(
        f'Переведено. Змінено {updated_count} навчань у {len(student_ids)} студентів.',
        'success',
    )
    return redirect(url_for('admin.order_detail', order_id=order_id))


@admin_bp.route('/admin/license_report')
@permission_required('license_report')
def license_report():
    """Звіт по ліцензіях.

    Ітеруємо по student_enrollments (студент ↔ група), а не по
    students.group_id — щоб коректно показати студента, який
    паралельно вчиться у 2+ групах під різними ліцензіями (він
    з'явиться в кожній групі під своєю ліцензією з цього enrollment).

    Для студентів без жодного активного enrollment (legacy) —
    fallback на students.group_id + students.license_id (UNION у
    підзапиті це обробляє природно).
    """
    from collections import defaultdict
    from routes.helpers import sort_ukrainian

    conn = get_db()
    conn.row_factory = sqlite3.Row

    view = request.args.get('view', 'list')
    if view not in ('list', 'summary'):
        view = 'list'
    include_archived = request.args.get('include_archived', '0') == '1'

    # ── 1. Усі активні ліцензії (для колонок звіту) ─────────────
    license_rows = conn.execute("""
        SELECT id, short_name_ua, name_ua
        FROM institution_licenses
        WHERE COALESCE(is_active, 1) = 1
        ORDER BY COALESCE(short_name_ua, name_ua)
    """).fetchall()
    license_names = [
        r['short_name_ua'] or r['name_ua'] for r in license_rows
    ]

    # ── 2. Плоский список "студент × група × ліцензія" ─────────
    # Підзапит effective_enrollments: активні enrollment'и + fallback
    # для legacy-студентів, у яких активних enrollment'ів немає, але
    # є students.group_id. UNION автоматично дедуплікує, тому студент
    # з однаковими (group_id, license_id) у обох джерелах не
    # задвоюється.
    student_archive_filter = "" if include_archived else " AND COALESCE(s.archived, 0) = 0"
    group_archive_filter = "" if include_archived else " AND COALESCE(g.archived, 0) = 0"

    flat_rows = conn.execute(f"""
        SELECT
            s.id                AS student_id,
            s.last_name_UA, s.first_name_UA, s.middle_name_UA,
            s.program_credits_override,
            s.archived          AS student_archived,
            g.id                AS gid,
            g.name              AS group_name,
            g.start_year, g.study_form, g.course,
            g.specialty, g.degree_level,
            g.program_credits,
            g.archived          AS group_archived,
            eff.license_id      AS license_id,
            il.short_name_ua    AS license_short_name,
            il.name_ua          AS license_name
        FROM (
            SELECT student_id, group_id, license_id
            FROM student_enrollments
            WHERE status = 'active' AND group_id IS NOT NULL

            UNION

            SELECT s2.id AS student_id, s2.group_id, s2.license_id
            FROM students s2
            WHERE s2.group_id IS NOT NULL
              AND NOT EXISTS (
                  SELECT 1 FROM student_enrollments e2
                  WHERE e2.student_id = s2.id AND e2.status = 'active'
              )
        ) AS eff
        JOIN students s ON s.id = eff.student_id
        JOIN groups g   ON g.id = eff.group_id
        LEFT JOIN institution_licenses il ON il.id = eff.license_id
        WHERE 1=1
              {student_archive_filter}
              {group_archive_filter}
    """).fetchall()

    # ── 3. Групуємо: група → студенти + лічильники ліцензій ────
    by_group = {}   # gid -> {group: {...}, students: [...], counts: {...}}
    for r in flat_rows:
        gid = r['gid']
        if gid not in by_group:
            by_group[gid] = {
                'group': {
                    'id': gid,
                    'name': r['group_name'] or '—',
                    'start_year': r['start_year'],
                    'study_form': r['study_form'],
                    'course': r['course'],
                    'specialty': r['specialty'] or '—',
                    'degree_level': r['degree_level'] or '—',
                    'archived': bool(r['group_archived']),
                    'program_credits': r['program_credits'],
                },
                'students': [],
                'counts': defaultdict(int),
            }
        full_name = ' '.join(
            x for x in (r['last_name_UA'], r['first_name_UA'],
                        r['middle_name_UA'] or '') if x
        )
        lic = r['license_short_name'] or r['license_name'] or '—'
        by_group[gid]['students'].append({
            'full_name': full_name,
            'program_credits_override': r['program_credits_override'],
            'archived': bool(r['student_archived']),
            'license_short_name': lic,
        })
        by_group[gid]['counts'][lic] += 1

    # Сортуємо студентів у кожній групі
    for gid_data in by_group.values():
        gid_data['students'] = sort_ukrainian(
            gid_data['students'], lambda s: s['full_name']
        )

    groups_with_students = list(by_group.values())

    # ── 4. Групуємо по ступенях (для view=list) ────────────────
    groups_by_degree = defaultdict(list)
    for item in groups_with_students:
        groups_by_degree[item['group']['degree_level']].append(item)
    for deg in groups_by_degree:
        groups_by_degree[deg].sort(
            key=lambda i: (i['group']['name'], i['group']['start_year'] or '')
        )

    # ── 5. Зведені таблиці (для view=summary) ──────────────────
    # 5.1 По групах
    summary_rows = []
    for item in groups_with_students:
        counts = dict(item['counts'])
        total = sum(counts.values())
        summary_rows.append({
            'group': item['group'],
            'counts': counts,
            'total': total,
        })
    summary_rows.sort(
        key=lambda r: (r['group']['degree_level'],
                       r['group']['name'],
                       r['group']['start_year'] or '')
    )

    # 5.2 Загальні підсумки по ліцензіях
    grand_totals = defaultdict(int)
    for item in groups_with_students:
        for lic, cnt in item['counts'].items():
            grand_totals[lic] += cnt
    grand_total_all = sum(grand_totals.values())

    # 5.3 По курсах
    course_agg = defaultdict(lambda: defaultdict(int))
    for item in groups_with_students:
        course = item['group']['course']
        if course is None:
            continue
        for lic, cnt in item['counts'].items():
            course_agg[course][lic] += cnt
    course_breakdown_rows = []
    for course in sorted(course_agg.keys(), key=lambda c: (c is None, c)):
        counts = dict(course_agg[course])
        course_breakdown_rows.append({
            'course': course,
            'counts': counts,
            'total': sum(counts.values()),
        })

    # 5.4 По ступенях + формах навчання
    bd_agg = defaultdict(lambda: defaultdict(int))
    for item in groups_with_students:
        key = (item['group']['degree_level'] or '—',
               item['group']['study_form'] or '—')
        for lic, cnt in item['counts'].items():
            bd_agg[key][lic] += cnt
    breakdown_rows = []
    for (degree, study_form) in sorted(bd_agg.keys()):
        counts = dict(bd_agg[(degree, study_form)])
        breakdown_rows.append({
            'degree_level': degree,
            'study_form': study_form,
            'counts': counts,
            'total': sum(counts.values()),
        })

    conn.close()

    return render_template(
        'license_report.html',
        view=view,
        include_archived=include_archived,
        groups_with_students=groups_with_students,
        groups_by_degree=dict(groups_by_degree),
        license_names=license_names,
        summary_rows=summary_rows,
        grand_totals=dict(grand_totals),
        grand_total_all=grand_total_all,
        course_breakdown_rows=course_breakdown_rows,
        breakdown_rows=breakdown_rows,
    )


@admin_bp.route('/admin/license_report/export_word')
@permission_required('license_report')
def license_report_export_word():
    """
    Вивантажує звіт по ліцензіях у .docx у форматі:
    "ВСТУП {рік}" -> "{курс} курс {форма}" -> реальна назва групи ->
    таблиця "№ / Прізвище І.П. / Ліцензія". Денна й заочна форми —
    окремі файли (study_form обов'язковий параметр).

    Ітеруємо по enrollment'ах (плюс fallback на legacy students.group_id),
    щоб паралельні студенти не загубились і ліцензія в таблиці була
    САМЕ та, під якою вони вчаться в цій групі.
    """
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from collections import OrderedDict

    include_archived = request.args.get('include_archived') == '1'
    study_form = request.args.get('study_form')  # 'Денна' або 'Заочна' — обов'язково
    if study_form not in ('Денна', 'Заочна'):
        flash("Оберіть форму навчання (Денна чи Заочна) для вивантаження.", "error")
        return redirect(url_for('admin.license_report'))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    group_cond = "" if include_archived else "AND COALESCE(g.archived, 0) = 0"
    student_cond = "" if include_archived else "AND COALESCE(s.archived, 0) = 0"

    rows = conn.execute(f"""
        SELECT
            g.start_year, g.course, g.name AS group_name, g.degree_level,
            g.specialty_code,
            sp.name_ua AS specialty_name_ua,
            TRIM(s.last_name_UA || ' ' || s.first_name_UA || ' ' ||
                 COALESCE(s.middle_name_UA, '')) AS full_name,
            COALESCE(il.short_name_ua, il.name_ua) AS license_short_name
        FROM (
            SELECT student_id, group_id, license_id
            FROM student_enrollments
            WHERE status = 'active' AND group_id IS NOT NULL

            UNION

            SELECT s2.id AS student_id, s2.group_id, s2.license_id
            FROM students s2
            WHERE s2.group_id IS NOT NULL
              AND NOT EXISTS (
                  SELECT 1 FROM student_enrollments e2
                  WHERE e2.student_id = s2.id AND e2.status = 'active'
              )
        ) AS eff
        JOIN groups g ON g.id = eff.group_id
        JOIN students s ON s.id = eff.student_id
        LEFT JOIN specialties sp ON sp.code = g.specialty_code
        LEFT JOIN institution_licenses il ON il.id = eff.license_id
        WHERE g.study_form = ? {group_cond} {student_cond}
        ORDER BY g.start_year DESC, g.course,
                 g.name COLLATE UKRAINIAN,
                 s.last_name_UA COLLATE UKRAINIAN
    """, (study_form,)).fetchall()
    conn.close()

    # Групуємо: ступінь (спершу Бакалавр, потім Магістр, решта — після)
    # -> рік вступу -> курс -> назва групи -> студенти (ім'я + ліцензія).
    DEGREE_ORDER = {'Бакалавр': 0, 'Магістр': 1}
    DEGREE_LABELS = {'Бакалавр': 'БАКАЛАВРИ', 'Магістр': 'МАГІСТРИ'}

    grouped = OrderedDict()
    for r in rows:
        degree_key = r['degree_level'] or 'Інше'
        section_key = (r['start_year'], r['course'])
        group_bucket = grouped.setdefault(degree_key, OrderedDict()).setdefault(
            section_key, OrderedDict()
        ).setdefault(
            r['group_name'],
            {
                'specialty_code': r['specialty_code'],
                'specialty_name': r['specialty_name_ua'],
                'students': [],
            },
        )
        group_bucket['students'].append(
            (r['full_name'], r['license_short_name'] or 'без ліцензії')
        )

    ordered_degrees = sorted(grouped.keys(), key=lambda d: DEGREE_ORDER.get(d, 99))

    doc = Document()
    for section in doc.sections:
        section.left_margin = Cm(2)
        section.right_margin = Cm(1)

    study_form_text = {'Денна': 'денна форма навчання',
                       'Заочна': 'заочна форма навчання'}.get(study_form, study_form)

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