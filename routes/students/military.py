"""
routes/students/military.py
============================
Військовий облік студента (структура "Резерв+"):

  - перегляд/редагування вручну;
  - завантаження PDF з "Резерв+": сам PDF зберігається як скан-копія
    (entity_type='military'), а розпізнані дані показуються у preview;
  - попередній перегляд змін (diff) ДВОХ типів;
  - застосування змін до відповідних таблиць з фіксацією історії;
  - фото 3×4 з PDF;
  - пакетний імпорт декількох PDF (студент визначається за ПІБ + датою
    народження, з чекбоксами по кожному полю в preview);
  - скановані копії (до 5 файлів, разом з PDF).

Тимчасові папки тепер усередині static/uploads/ (як і решта файлів):
  - static/uploads/tmp_military_pdf/     (поточний PDF при парсингу)
  - static/uploads/tmp_military_bulk/<token>/ (файли пакетного імпорту)
Прямий доступ до них через /static/ заблокований в app.py.
"""

from datetime import datetime
import os
import uuid
import json
import tempfile
import shutil

from flask import (
    render_template, request, redirect, url_for, session, flash,
    current_app, jsonify,
)
from werkzeug.utils import secure_filename

from routes.db import get_db
from routes.utils import log_action, login_required, permission_required, logger
from routes.helpers import current_username
from routes.students import students_bp
from routes.utils import get_attachments, save_multiple_attachments
from routes.parse_military_pdf import parse_reserve_plus_pdf, extract_photo_from_pdf
from routes.photo import (
    process_and_save_pending_photo,
    process_and_save_pending_photo_smart,
    process_and_save_pending_photo_with_crop,
    save_raw_pending_photo,
    load_and_validate_image,
    auto_center_crop_box,
    crop_and_resize,
    save_final_image,
)
from routes.config import PROJECT_ROOT

MAX_FILES = 5

# Тимчасові папки — тепер усередині static/uploads/, як і решта
# файлів, щоб усе зберігалося в одному місці й не залежало від
# робочої директорії процесу.
TMP_MILITARY_PDF_DIR = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'tmp_military_pdf')
TMP_MILITARY_BULK_DIR = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'tmp_military_bulk')


# Військові поля (колонки в military).
MILITARY_FIELDS = [
    ('status_badge', 'Статус'),
    ('valid_until', 'Дійсний до'),
    ('registration_category', 'Категорія обліку'),
    ('removal_basis', 'Підстава для зняття'),
    ('tck_name', 'ТЦК'),
    ('military_rank', 'Звання'),
    ('registry_number', 'Номер запису (ДРПВР)'),
    ('vos_code', 'Код ВОС'),
    ('vos_description', 'Опис ВОС'),
    ('deferral_type', 'Тип відстрочки'),
    ('deferral_until', 'Відстрочка до'),
    ('police_reason', 'Причина звернення до поліції'),
    ('police_date', 'Дата звернення до поліції'),
    ('vlk_decision', 'Постанова ВЛК'),
    ('vlk_date', 'Дата постанови ВЛК'),
    ('disability_group', 'Група інвалідності'),
    ('disability_valid_until', 'Група інвалідності дійсна до'),
    ('disability_reason', 'Причина інвалідності'),
    ('formed_at', 'Документ сформовано'),
    ('data_verified_on_time', 'Дані уточнено вчасно'),
    ('data_update_date', 'Дата уточнення даних'),
]

# Поля картки студента, які беремо з PDF. У military їх НЕМАЄ.
STUDENT_FIELDS = [
    ('birth_date', 'Дата народження'),
    ('tax_id', 'Ідентифікаційний код'),
    ('address_of_residence', 'Адреса проживання'),
    ('phone', 'Телефон'),
    ('email', 'Email'),
]

# Легасі-поля military, які вже не приходять з PDF, але можуть бути
# заповнені вручну.
LEGACY_FIELDS = [
    ('military_registration_document', 'Військово-обліковий документ'),
    ('issued_VOD', 'Виданий ВОД'),
    ('military_accounting_specialty_number', 'Номер ВОС (старе поле)'),
    ('change_credentials', 'Зміна облікових даних'),
    ('reason_for_changing_credentials', 'Підстава для зміни'),
    ('being_on_military_registration', 'Перебування на обліку'),
]


def _field_label(name, scope='military'):
    src = (MILITARY_FIELDS + LEGACY_FIELDS) if scope == 'military' else STUDENT_FIELDS
    for key, label in src:
        if key == name:
            return label
    return name


def _normalize(value):
    if value is None:
        return ''
    return str(value).strip()


def _normalize_bool(value):
    if isinstance(value, bool):
        return 1 if value else 0
    if value in (0, 1):
        return value
    if isinstance(value, str):
        return 1 if value.strip().lower() in ('1', 'true', 'yes', 'on') else 0
    return 0


def _normalize_date(value):
    """Приводить будь-який формат дати до DD.MM.YYYY для порівняння."""
    if not value:
        return ''
    s = str(value).strip()
    import re as _re
    m = _re.match(r'^(\d{4})-(\d{2})-(\d{2})$', s)
    if m:
        return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
    m = _re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
    if m:
        return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    return s


def _collect_form_data(form):
    result = {}
    for key, _label in MILITARY_FIELDS + LEGACY_FIELDS:
        if key == 'data_verified_on_time':
            result[key] = 1 if form.get(key) in ('on', '1', 'true', 'yes') else 0
        else:
            result[key] = _normalize(form.get(key))
    return result


def _current_military(conn, student_id):
    row = conn.execute("SELECT * FROM military WHERE student_id = ?", (student_id,)).fetchone()
    return dict(row) if row else None


def _current_student(conn, student_id):
    row = conn.execute(
        "SELECT id, last_name_UA, first_name_UA, middle_name_UA, group_id, "
        "birth_date, tax_id, address_of_residence, phone, email, photo "
        "FROM students WHERE id = ?", (student_id,)
    ).fetchone()
    return dict(row) if row else None


def _build_diff(old, new, fields):
    """
    Diff з ігноруванням порожніх `new` (щоб імпорт PDF не стирав
    існуючі дані, яких у PDF немає) і нормалізацією дат.
    """
    old = old or {}
    diff = []
    for key, label in fields:
        if key == 'data_verified_on_time':
            old_v = _normalize_bool(old.get(key))
            new_v = _normalize_bool(new.get(key))
            if old_v != new_v:
                diff.append({'field': key, 'label': label, 'old': old_v, 'new': new_v})
            continue

        if 'date' in key:
            old_v = _normalize_date(old.get(key))
            new_v = _normalize_date(new.get(key))
        else:
            old_v = _normalize(old.get(key))
            new_v = _normalize(new.get(key))

        if not new_v:
            continue

        if old_v != new_v:
            diff.append({'field': key, 'label': label, 'old': old_v, 'new': new_v})

    return diff


def _log_history(conn, student_id, username, source, scope, diff):
    for d in diff:
        conn.execute(
            """INSERT INTO military_history
                   (student_id, changed_by, source, scope, field_name, old_value, new_value)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (student_id, username, source, scope, d['field'], d['old'], d['new']),
        )


def _save_military(conn, student_id, data, source, username):
    existing = conn.execute("SELECT * FROM military WHERE student_id = ?", (student_id,)).fetchone()
    diff = _build_diff(dict(existing) if existing else None, data, MILITARY_FIELDS + LEGACY_FIELDS)

    columns = list(data.keys())
    values = [data[c] for c in columns]

    if existing:
        set_clause = ', '.join(f"{c} = ?" for c in columns)
        conn.execute(f"UPDATE military SET {set_clause} WHERE student_id = ?", values + [student_id])
        military_id = existing['id']
    else:
        placeholders = ', '.join('?' for _ in columns)
        conn.execute(
            f"INSERT INTO military (student_id, {', '.join(columns)}) VALUES (?, {placeholders})",
            [student_id] + values,
        )
        military_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

    _log_history(conn, student_id, username, source, 'military', diff)
    return military_id, diff


def _save_student_fields(conn, student_id, data, source, username):
    """
    Записує поля картки студента. UPDATE будується ЛИШЕ за полями,
    які є в `data` — щоб не затирати відсутні поля.
    """
    existing = conn.execute(
        "SELECT birth_date, tax_id, address_of_residence, phone, email "
        "FROM students WHERE id = ?",
        (student_id,),
    ).fetchone()
    diff = _build_diff(dict(existing) if existing else None, data, STUDENT_FIELDS)
    if not diff:
        return diff

    columns = list(data.keys())
    set_clause = ', '.join(f"{c} = ?" for c in columns)
    values = [data[c] for c in columns]
    conn.execute(f"UPDATE students SET {set_clause} WHERE id = ?", values + [student_id])
    _log_history(conn, student_id, username, source, 'student', diff)
    return diff


def _ensure_military_record(conn, student_id):
    """Повертає id запису military для студента. Створює порожній, якщо немає."""
    row = conn.execute("SELECT id FROM military WHERE student_id = ?", (student_id,)).fetchone()
    if row:
        return row['id']
    cur = conn.execute("INSERT INTO military (student_id) VALUES (?)", (student_id,))
    return cur.lastrowid


def _save_pdf_as_attachment(conn, student_id, pdf_bytes, original_filename, username):
    """
    Зберігає PDF у static/uploads/military_records/ і реєструє в
    attachments з entity_type='military'. Повертає rel_path.
    """
    upload_dir = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'military_records')
    os.makedirs(upload_dir, exist_ok=True)

    ext = os.path.splitext(original_filename)[1] or '.pdf'
    safe_name = f"{uuid.uuid4().hex}{ext}"
    file_path = os.path.join(upload_dir, safe_name)
    with open(file_path, 'wb') as f:
        f.write(pdf_bytes)

    rel_path = f"uploads/military_records/{safe_name}"

    military_id = _ensure_military_record(conn, student_id)
    conn.execute(
        "INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_by) "
        "VALUES (?, ?, ?, ?, ?)",
        ('military', military_id, rel_path, original_filename, username),
    )
    return rel_path


# ============================================================
# Маршрути
# ============================================================

@students_bp.route('/students/<int:student_id>/military/add', methods=['GET', 'POST'])
@login_required('')
def add_military(student_id):
    """Сумісність зі старим кодом — редірект на military_data."""
    return redirect(url_for('students.military_data', student_id=student_id))


@students_bp.route('/students/<int:student_id>/military', methods=['GET', 'POST'])
@login_required('')
def military_data(student_id):
    """GET — форма + історія; POST — збереження вручну."""
    conn = get_db()
    student = _current_student(conn, student_id)
    if not student:
        conn.close()
        flash("Студента не знайдено", "error")
        return redirect(url_for('students.student_list'))

    military = _current_military(conn, student_id)

    if request.method == 'POST':
        existing_attachments = get_attachments(conn, 'military', military['id']) if military else []
        new_scans = [f for f in request.files.getlist('military_scans') if f and f.filename]
        if len(existing_attachments) + len(new_scans) > MAX_FILES:
            conn.close()
            flash(f'Забагато файлів — максимум {MAX_FILES} (вже є {len(existing_attachments)}).', 'error')
            return redirect(url_for('students.military_data', student_id=student_id))

        data = _collect_form_data(request.form)

        try:
            military_id, diff = _save_military(
                conn, student_id, data, source='manual', username=current_username(),
            )
            if new_scans:
                save_multiple_attachments(
                    conn, 'military', military_id, new_scans,
                    'military_records', current_username(),
                )
            conn.commit()
        except Exception as e:
            conn.rollback()
            logger.error(f"Помилка збереження військових даних (student_id={student_id}): {e}", exc_info=True)
            conn.close()
            flash(f"Помилка збереження: {e}", "error")
            return redirect(url_for('students.military_data', student_id=student_id))

        action_name = "редагував" if military else "додав"
        log_action(
            current_username(),
            f"{action_name} військові дані: {student['last_name_UA']} {student['first_name_UA']} (ID {student_id})",
            group_ids=[student['group_id']],
            details=f"змінених полів: {len(diff)}",
        )
        conn.close()
        flash(f"Збережено. Змінених полів: {len(diff)}", "success")
        return redirect(url_for('students.military_data', student_id=student_id))

    military_attachments = get_attachments(conn, 'military', military['id']) if military else []
    history = conn.execute(
        """SELECT * FROM military_history
           WHERE student_id = ?
           ORDER BY changed_at DESC, id DESC
           LIMIT 200""",
        (student_id,),
    ).fetchall()
    conn.close()

    return render_template(
        'edit_military.html',
        student=student,
        military=military,
        military_attachments=military_attachments,
        max_files=MAX_FILES,
        military_fields=MILITARY_FIELDS,
        legacy_fields=LEGACY_FIELDS,
        history=history,
        field_label=_field_label,
    )


@students_bp.route('/students/<int:student_id>/military/upload_pdf', methods=['POST'])
@login_required('')
def military_upload_pdf(student_id):
    """
    Парсить PDF, зберігає його як скан-копію (entity_type='military')
    і показує preview з diff-блоками + блоком фото.
    """
    file = request.files.get('pdf_file')
    if not file or not file.filename:
        flash("Оберіть PDF-файл", "error")
        return redirect(url_for('students.military_data', student_id=student_id))
    if not file.filename.lower().endswith('.pdf'):
        flash("Підтримуються лише PDF-файли", "error")
        return redirect(url_for('students.military_data', student_id=student_id))

    pdf_bytes = file.read()
    original_filename = secure_filename(file.filename) or 'upload.pdf'

    conn = get_db()
    student = _current_student(conn, student_id)
    if not student:
        conn.close()
        flash("Студента не знайдено", "error")
        return redirect(url_for('students.student_list'))

    military = _current_military(conn, student_id)
    existing_attachments = get_attachments(conn, 'military', military['id']) if military else []
    if len(existing_attachments) + 1 > MAX_FILES:
        conn.close()
        flash(f'Забагато файлів — максимум {MAX_FILES} (вже є {len(existing_attachments)}). '
              f'Видаліть непотрібні скани перед завантаженням нового PDF.', 'error')
        return redirect(url_for('students.military_data', student_id=student_id))

    # Тимчасова папка — тепер усередині static/uploads/
    os.makedirs(TMP_MILITARY_PDF_DIR, exist_ok=True)
    tmp_path = os.path.join(
        TMP_MILITARY_PDF_DIR,
        f"{student_id}_{int(datetime.now().timestamp())}_{original_filename}"
    )
    with open(tmp_path, 'wb') as f:
        f.write(pdf_bytes)

    try:
        parsed = parse_reserve_plus_pdf(tmp_path)
    except Exception as e:
        logger.error(f"Помилка парсингу PDF (student_id={student_id}): {e}", exc_info=True)
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        conn.close()
        flash(f"Не вдалося розпізнати PDF: {e}", "error")
        return redirect(url_for('students.military_data', student_id=student_id))

    raw_photo_bytes = extract_photo_from_pdf(tmp_path)
    logger.info(
        f"PDF photo extract (student_id={student_id}): "
        f"{len(raw_photo_bytes) if raw_photo_bytes else 0} bytes"
    )

    try:
        os.remove(tmp_path)
    except OSError:
        pass

    # Зберігаємо PDF як скан-копію (не видаляємо!)
    try:
        saved_pdf_rel = _save_pdf_as_attachment(
            conn, student_id, pdf_bytes, original_filename, current_username()
        )
        conn.commit()
        logger.info(f"PDF saved as attachment (student_id={student_id}): {saved_pdf_rel}")
    except Exception as e:
        conn.rollback()
        logger.error(f"Не вдалося зберегти PDF як скан-копію (student_id={student_id}): {e}", exc_info=True)
        flash(f"Увага: PDF не вдалося зберегти в сканах: {e}", "warning")

    new_military = {}
    for key, _label in MILITARY_FIELDS:
        if key == 'data_verified_on_time':
            new_military[key] = 1 if parsed.get('data_verified_on_time') else 0
        else:
            new_military[key] = _normalize(parsed.get(key))

    new_student = {key: _normalize(parsed.get(key)) for key, _label in STUDENT_FIELDS}

    photo_rel_path = None
    raw_photo_path = None
    photo_preview_error = None
    if raw_photo_bytes:
        try:
            raw_photo_path = save_raw_pending_photo(raw_photo_bytes)
            photo_rel_path = process_and_save_pending_photo_smart(raw_photo_bytes)
            logger.info(
                f"PDF photo saved (student_id={student_id}): "
                f"raw={raw_photo_path}, cropped={photo_rel_path}"
            )
        except ValueError as e:
            photo_preview_error = str(e)
            logger.warning(f"PDF photo processing failed (student_id={student_id}): {e}")

    session[f'military_pdf_pending_raw_photo_{student_id}'] = raw_photo_path

    student = _current_student(conn, student_id)
    military = _current_military(conn, student_id)
    conn.close()

    military_diff = _build_diff(military, new_military, MILITARY_FIELDS)
    student_diff = _build_diff(student, new_student, STUDENT_FIELDS)

    session[f'military_pdf_pending_military_{student_id}'] = new_military
    session[f'military_pdf_pending_student_{student_id}'] = new_student
    session[f'military_pdf_pending_photo_{student_id}'] = photo_rel_path

    return render_template(
        'military_pdf_preview.html',
        student=student,
        military_diff=military_diff,
        student_diff=student_diff,
        new_military=new_military,
        new_student=new_student,
        military_fields=MILITARY_FIELDS,
        field_label=_field_label,
        pending_photo_path=photo_rel_path,
        raw_photo_path=raw_photo_path,
        photo_error=photo_preview_error,
    )


@students_bp.route('/students/<int:student_id>/military/recrop_pdf_photo', methods=['POST'])
@login_required('')
def military_pdf_recrop_photo(student_id):
    """Ручна переобрізка фото з PDF (Cropper.js)."""
    raw_key = f'military_pdf_pending_raw_photo_{student_id}'
    cropped_key = f'military_pdf_pending_photo_{student_id}'

    raw_path = session.get(raw_key)
    if not raw_path:
        return jsonify({'error': 'no_raw_photo'}), 400

    raw_abs = os.path.join(PROJECT_ROOT, 'static', raw_path)
    if not os.path.exists(raw_abs):
        return jsonify({'error': 'raw_photo_missing'}), 404

    try:
        crop_x = float(request.form['crop_x'])
        crop_y = float(request.form['crop_y'])
        crop_w = float(request.form['crop_w'])
        crop_h = float(request.form['crop_h'])
    except (KeyError, ValueError):
        return jsonify({'error': 'invalid_crop'}), 400

    try:
        with open(raw_abs, 'rb') as f:
            raw_bytes = f.read()

        old_cropped = session.get(cropped_key)
        new_cropped = process_and_save_pending_photo_with_crop(
            raw_bytes, (crop_x, crop_y, crop_w, crop_h), old_rel_path=old_cropped
        )
        session[cropped_key] = new_cropped

        return jsonify({
            'ok': True,
            'preview_url': url_for(
                'static', filename=new_cropped, v=uuid.uuid4().hex
            ),
        })
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        logger.error(f"Помилка переобрізки фото (student_id={student_id}): {e}", exc_info=True)
        return jsonify({'error': 'server_error'}), 500


@students_bp.route('/students/<int:student_id>/military/apply_pdf', methods=['POST'])
@login_required('')
def military_apply_pdf(student_id):
    """Застосовує зміни з PDF."""
    mil_key = f'military_pdf_pending_military_{student_id}'
    stud_key = f'military_pdf_pending_student_{student_id}'
    photo_key = f'military_pdf_pending_photo_{student_id}'
    raw_key = f'military_pdf_pending_raw_photo_{student_id}'

    new_military = session.get(mil_key)
    new_student = session.get(stud_key)
    pending_photo = session.get(photo_key)

    if not new_military and not new_student and not pending_photo:
        flash("Немає даних для застосування (сесія застаріла). Завантажте PDF ще раз.", "warning")
        return redirect(url_for('students.military_data', student_id=student_id))

    apply_student_keys = set(request.form.getlist('apply_student_field'))
    apply_photo = bool(request.form.get('apply_photo'))

    conn = get_db()
    student = _current_student(conn, student_id)
    if not student:
        conn.close()
        session.pop(mil_key, None)
        session.pop(stud_key, None)
        session.pop(photo_key, None)
        session.pop(raw_key, None)
        flash("Студента не знайдено", "error")
        return redirect(url_for('students.student_list'))

    photo_saved = False
    try:
        mil_diff = []
        if new_military:
            _, mil_diff = _save_military(
                conn, student_id, new_military, source='pdf_import',
                username=current_username(),
            )

        stud_diff = []
        if new_student and apply_student_keys:
            filtered_student = {
                k: v for k, v in new_student.items() if k in apply_student_keys
            }
            stud_diff = _save_student_fields(
                conn, student_id, filtered_student, source='pdf_import',
                username=current_username(),
            )

        if pending_photo and apply_photo:
            src_abs = os.path.join(PROJECT_ROOT, 'static', pending_photo)
            if os.path.exists(src_abs):
                with open(src_abs, 'rb') as f:
                    raw = f.read()
                try:
                    img = load_and_validate_image(raw)
                    crop = auto_center_crop_box(img.width, img.height)
                    final_img = crop_and_resize(img, crop)
                    new_rel = save_final_image(final_img, student_id)
                    conn.execute(
                        "UPDATE students SET photo = ? WHERE id = ?",
                        (new_rel, student_id),
                    )
                    _log_history(
                        conn, student_id, current_username(), 'pdf_import', 'student',
                        [{'field': 'photo', 'label': 'Фото 3×4',
                          'old': student.get('photo') or '', 'new': new_rel}],
                    )
                    photo_saved = True
                except ValueError as e:
                    logger.warning(f"Не вдалося оновити фото (student_id={student_id}): {e}")

        conn.commit()
    except Exception as e:
        conn.rollback()
        logger.error(f"Помилка застосування PDF-даних (student_id={student_id}): {e}", exc_info=True)
        conn.close()
        flash(f"Помилка збереження: {e}", "error")
        return redirect(url_for('students.military_data', student_id=student_id))

    conn.close()

    if pending_photo:
        pending_abs = os.path.join(PROJECT_ROOT, 'static', pending_photo)
        try:
            if os.path.exists(pending_abs):
                os.remove(pending_abs)
        except OSError:
            pass

    raw_photo = session.get(raw_key)
    if raw_photo:
        raw_abs = os.path.join(PROJECT_ROOT, 'static', raw_photo)
        try:
            if os.path.exists(raw_abs):
                os.remove(raw_abs)
        except OSError:
            pass

    session.pop(mil_key, None)
    session.pop(stud_key, None)
    session.pop(photo_key, None)
    session.pop(raw_key, None)

    log_action(
        current_username(),
        f"оновив військові дані з PDF: {student['last_name_UA']} {student['first_name_UA']} (ID {student_id})",
        group_ids=[student['group_id']],
        details=(
            f"військових: {len(mil_diff)}, "
            f"картки: {len(stud_diff)}, "
            f"фото: {'оновлено' if photo_saved else 'без змін'}"
        ),
    )
    flash(
        f"Дані з PDF застосовано. Військових: {len(mil_diff)}, "
        f"картки: {len(stud_diff)}, фото: {'оновлено' if photo_saved else 'без змін'}",
        "success",
    )
    return redirect(url_for('students.military_data', student_id=student_id))


@students_bp.route('/students/<int:student_id>/military/delete_attachment', methods=['POST'])
@login_required('')
def delete_military_attachment(student_id):
    conn = get_db()
    attachment_id = request.form.get('attachment_id')
    att = conn.execute("SELECT file_path FROM attachments WHERE id=?", (attachment_id,)).fetchone()
    if att:
        att_path = os.path.join(PROJECT_ROOT, 'static', att['file_path'])
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
    conn = get_db()
    student_row = conn.execute(
        "SELECT last_name_UA, first_name_UA, group_id FROM students WHERE id=?", (student_id,)
    ).fetchone()
    military_row = conn.execute("SELECT id FROM military WHERE student_id = ?", (student_id,)).fetchone()
    if military_row:
        for att in get_attachments(conn, 'military', military_row['id']):
            att_path = os.path.join(PROJECT_ROOT, 'static', att['file_path'])
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


# ============================================================
# Пакетний імпорт PDF «Резерв+» (декілька файлів одразу)
# ============================================================

@students_bp.route('/admin/military/bulk_import_pdf', methods=['GET', 'POST'])
@login_required('')
def military_bulk_import_pdf():
    """Пакетний імпорт PDF «Резерв+»."""
    if request.method == 'GET':
        return render_template('military_bulk_import_upload.html')

    files = request.files.getlist('pdf_files')
    files = [f for f in files if f and f.filename and f.filename.lower().endswith('.pdf')]

    if not files:
        flash("Оберіть хоча б один PDF-файл", "error")
        return redirect(url_for('students.military_bulk_import_pdf'))

    conn = get_db()
    token = uuid.uuid4().hex
    # Тимчасова папка — тепер усередині static/uploads/
    tmp_dir = os.path.join(TMP_MILITARY_BULK_DIR, token)
    os.makedirs(tmp_dir, exist_ok=True)

    results = []

    for i, file in enumerate(files):
        result = {
            'idx': i,
            'filename': file.filename,
            'student_id': None,
            'student_name': None,
            'student_group': None,
            'error': None,
            'military_diff': [],
            'student_diff': [],
            'new_military': None,
            'new_student': None,
            'has_photo': False,
            'pdf_saved_path': None,
            'photo_tmp_path': None,
        }
        try:
            pdf_bytes = file.read()
            pdf_name = f"pdf_{i}.pdf"
            pdf_disk_path = os.path.join(tmp_dir, pdf_name)
            with open(pdf_disk_path, 'wb') as f:
                f.write(pdf_bytes)
            # Шлях відносно PROJECT_ROOT (щоб потім легко знайти)
            result['pdf_saved_path'] = f"static/uploads/tmp_military_bulk/{token}/{pdf_name}"

            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(pdf_bytes)
                tmp_path = tmp.name
            try:
                parsed = parse_reserve_plus_pdf(tmp_path)
                photo_bytes = extract_photo_from_pdf(tmp_path)
            finally:
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

            full_name = (parsed.get('full_name') or '').strip()
            birth_date = (parsed.get('birth_date') or '').strip()

            if not full_name:
                result['error'] = "У PDF не знайдено ПІБ"
                results.append(result)
                continue
            if not birth_date:
                result['error'] = "У PDF не знайдено дату народження"
                results.append(result)
                continue

            name_parts = full_name.split()
            if len(name_parts) < 2:
                result['error'] = f"Некоректний ПІБ: {full_name!r}"
                results.append(result)
                continue
            last_name = name_parts[0]
            first_name = name_parts[1]
            middle_name = ' '.join(name_parts[2:]) if len(name_parts) > 2 else ''

            student = conn.execute("""
                SELECT id, last_name_UA, first_name_UA, middle_name_UA, birth_date, group_id
                FROM students
                WHERE LOWER_UA(TRIM(last_name_UA)) = LOWER_UA(?)
                  AND LOWER_UA(TRIM(first_name_UA)) = LOWER_UA(?)
                  AND LOWER_UA(TRIM(COALESCE(middle_name_UA, ''))) = LOWER_UA(?)
                  AND TRIM(COALESCE(birth_date, '')) = ?
                LIMIT 1
            """, (last_name, first_name, middle_name, birth_date)).fetchone()

            if not student:
                student = conn.execute("""
                    SELECT id, last_name_UA, first_name_UA, middle_name_UA, birth_date, group_id
                    FROM students
                    WHERE LOWER_UA(TRIM(last_name_UA)) = LOWER_UA(?)
                      AND LOWER_UA(TRIM(first_name_UA)) = LOWER_UA(?)
                      AND TRIM(COALESCE(birth_date, '')) = ?
                    LIMIT 1
                """, (last_name, first_name, birth_date)).fetchone()

            if not student:
                result['error'] = (
                    f"Студента не знайдено: {last_name} {first_name} {middle_name}, "
                    f"дата народження {birth_date}"
                )
                results.append(result)
                continue

            result['student_id'] = student['id']
            result['student_name'] = (
                f"{student['last_name_UA']} {student['first_name_UA']} "
                f"{student['middle_name_UA'] or ''}".strip()
            )
            result['student_group'] = student['group_id']

            new_military = {}
            for key, _label in MILITARY_FIELDS:
                if key == 'data_verified_on_time':
                    new_military[key] = 1 if parsed.get('data_verified_on_time') else 0
                else:
                    new_military[key] = _normalize(parsed.get(key))
            new_student = {key: _normalize(parsed.get(key)) for key, _label in STUDENT_FIELDS}

            military_row = _current_military(conn, student['id'])
            student_row = _current_student(conn, student['id'])

            result['military_diff'] = _build_diff(military_row, new_military, MILITARY_FIELDS)
            result['student_diff'] = _build_diff(student_row, new_student, STUDENT_FIELDS)
            result['new_military'] = new_military
            result['new_student'] = new_student

            if photo_bytes:
                try:
                    photo_rel = process_and_save_pending_photo_smart(photo_bytes)
                    photo_src = os.path.join(PROJECT_ROOT, 'static', photo_rel)
                    photo_dst_name = f"photo_{i}.jpg"
                    photo_dst = os.path.join(tmp_dir, photo_dst_name)
                    if os.path.exists(photo_src):
                        shutil.copy2(photo_src, photo_dst)
                        os.remove(photo_src)
                        result['has_photo'] = True
                        result['photo_tmp_path'] = f"static/uploads/tmp_military_bulk/{token}/{photo_dst_name}"
                except Exception as e:
                    logger.warning(f"Фото не оброблено для {file.filename}: {e}")

        except Exception as e:
            logger.error(f"Помилка обробки {file.filename}: {e}", exc_info=True)
            result['error'] = f"Помилка обробки: {e}"

        results.append(result)

    conn.close()

    state_path = os.path.join(tmp_dir, 'state.json')
    with open(state_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, default=str)

    session['military_bulk_token'] = token

    return render_template(
        'military_bulk_import_preview.html',
        token=token,
        results=results,
        total=len(results),
        success_count=sum(1 for r in results if not r['error']),
        error_count=sum(1 for r in results if r['error']),
    )


@students_bp.route('/admin/military/bulk_import_pdf/<token>/apply', methods=['POST'])
@login_required('')
def military_bulk_import_apply(token):
    """Застосовує зміни з пакетного імпорту (з чекбоксами по полях)."""
    tmp_dir = os.path.join(TMP_MILITARY_BULK_DIR, token)
    state_path = os.path.join(tmp_dir, 'state.json')

    if not os.path.exists(state_path):
        flash("Сесія імпорту не знайдена або застаріла", "warning")
        return redirect(url_for('students.military_bulk_import_pdf'))

    with open(state_path, encoding='utf-8') as f:
        results = json.load(f)

    apply_indices = set(int(x) for x in request.form.getlist('apply') if x.isdigit())

    conn = get_db()
    applied = 0
    errors = 0

    for r in results:
        idx = r.get('idx', -1)
        if r.get('error') or idx not in apply_indices or not r.get('student_id'):
            continue

        try:
            diff_mil_fields = {d['field'] for d in r.get('military_diff') or []}
            apply_mil_fields = set()
            for f in diff_mil_fields:
                if request.form.get(f"apply_mil_{idx}_{f}") == '1':
                    apply_mil_fields.add(f)

            mil_data = {
                k: v for k, v in (r.get('new_military') or {}).items()
                if k in apply_mil_fields
            }

            diff_stud_fields = {d['field'] for d in r.get('student_diff') or []}
            apply_stud_fields = set()
            for f in diff_stud_fields:
                if request.form.get(f"apply_stud_{idx}_{f}") == '1':
                    apply_stud_fields.add(f)

            stud_data = {
                k: v for k, v in (r.get('new_student') or {}).items()
                if k in apply_stud_fields
            }

            mil_id = None
            mil_diff = []
            if mil_data:
                mil_id, mil_diff = _save_military(
                    conn, r['student_id'], mil_data,
                    source='pdf_import', username=current_username(),
                )
            else:
                row = conn.execute(
                    "SELECT id FROM military WHERE student_id = ?", (r['student_id'],)
                ).fetchone()
                mil_id = row['id'] if row else _ensure_military_record(conn, r['student_id'])

            stud_diff = []
            if stud_data:
                stud_diff = _save_student_fields(
                    conn, r['student_id'], stud_data,
                    source='pdf_import', username=current_username(),
                )

            apply_photo = request.form.get(f"apply_photo_{idx}") == '1'
            photo_saved = False
            if r.get('has_photo') and r.get('photo_tmp_path') and apply_photo:
                photo_abs = os.path.join(PROJECT_ROOT, r['photo_tmp_path'].replace('/', os.sep))
                if os.path.exists(photo_abs):
                    with open(photo_abs, 'rb') as f:
                        raw = f.read()
                    try:
                        img = load_and_validate_image(raw)
                        crop = auto_center_crop_box(img.width, img.height)
                        final_img = crop_and_resize(img, crop)
                        new_rel = save_final_image(final_img, r['student_id'])
                        conn.execute(
                            "UPDATE students SET photo = ? WHERE id = ?",
                            (new_rel, r['student_id']),
                        )
                        _log_history(
                            conn, r['student_id'], current_username(),
                            'pdf_import', 'student',
                            [{'field': 'photo', 'label': 'Фото 3×4',
                              'old': '', 'new': new_rel}],
                        )
                        photo_saved = True
                    except ValueError as e:
                        logger.warning(
                            f"Фото не оновлено для student_id={r['student_id']}: {e}"
                        )

            apply_pdf_scan = request.form.get(f"apply_pdf_scan_{idx}") == '1'
            pdf_attached = False
            if apply_pdf_scan and mil_id and r.get('pdf_saved_path'):
                pdf_abs = os.path.join(PROJECT_ROOT, r['pdf_saved_path'].replace('/', os.sep))
                if os.path.exists(pdf_abs):
                    existing = get_attachments(conn, 'military', mil_id)
                    if len(existing) < MAX_FILES:
                        upload_dir = os.path.join(PROJECT_ROOT, 'static', 'uploads', 'military_records')
                        os.makedirs(upload_dir, exist_ok=True)
                        safe_name = f"{uuid.uuid4().hex}.pdf"
                        dest = os.path.join(upload_dir, safe_name)
                        shutil.copy2(pdf_abs, dest)
                        rel = f"uploads/military_records/{safe_name}"
                        conn.execute(
                            "INSERT INTO attachments (entity_type, entity_id, file_path, original_name, uploaded_by) "
                            "VALUES (?, ?, ?, ?, ?)",
                            ('military', mil_id, rel, r['filename'], current_username())
                        )
                        pdf_attached = True

            conn.commit()
            applied += 1

            log_action(
                current_username(),
                f"пакетний імпорт PDF: {r['student_name']} (ID {r['student_id']})",
                details=(
                    f"військових: {len(mil_diff)}, картки: {len(stud_diff)}, "
                    f"фото: {'так' if photo_saved else 'ні'}, "
                    f"PDF-скан: {'так' if pdf_attached else 'ні'}, "
                    f"файл: {r['filename']}"
                ),
            )
        except Exception as e:
            conn.rollback()
            logger.error(f"Помилка застосування {r['filename']}: {e}", exc_info=True)
            errors += 1

    conn.close()

    try:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    except Exception:
        pass
    session.pop('military_bulk_token', None)

    flash(
        f"Застосовано: {applied}. Помилок: {errors}.",
        'success' if errors == 0 else 'warning',
    )
    return redirect(url_for('students.military_bulk_import_pdf'))