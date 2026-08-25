# -*- coding: utf-8 -*-
"""
Масове завантаження фото студентів: багато файлів одразу, кожен файл
зіставляється зі студентом за іменем у назві файлу ("Прізвище Ім'я.jpg"),
обрізається автоматично по центру до 3х4, з можливістю вручну
відредагувати обрізку конкретного фото на кроці попереднього перегляду
перед остаточним підтвердженням.

Той самий загальний підхід, що й у майстрі імпорту оцінок
(routes/import_grades.py): нечітке зіставлення за текстом + стан
майстра в тимчасових файлах на диску, а не в cookie-сесії.
"""
import os
import re
import json
import time
import uuid
import difflib

from flask import Blueprint, render_template, request, redirect, url_for, flash, session, jsonify
from routes.db import get_db
from routes.utils import login_required, permission_required, log_action
from routes.helpers import current_username
from routes.photo import (
    load_and_validate_image, auto_center_crop_box, crop_and_resize,
    save_final_image, MAX_FILE_SIZE_BYTES, ALLOWED_FORMATS,
)

photo_bulk_bp = Blueprint('photo_bulk', __name__)

SESSIONS_DIR = os.path.join(os.getcwd(), 'photo_bulk_sessions')
SESSION_TTL_SECONDS = 6 * 3600


def _cleanup_old_sessions():
    if not os.path.isdir(SESSIONS_DIR):
        return
    now = time.time()
    for name in os.listdir(SESSIONS_DIR):
        path = os.path.join(SESSIONS_DIR, name)
        try:
            if os.path.isdir(path) and now - os.path.getmtime(path) > SESSION_TTL_SECONDS:
                import shutil
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass


def _session_dir(token):
    return os.path.join(SESSIONS_DIR, token)


def _manifest_path(token):
    return os.path.join(_session_dir(token), 'manifest.json')


def _original_path(token, file_key):
    return os.path.join(_session_dir(token), f"orig_{file_key}.jpg")


def _preview_path(token, file_key):
    return os.path.join(_session_dir(token), f"preview_{file_key}.jpg")


def _save_manifest(token, data):
    with open(_manifest_path(token), 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)


def _load_manifest(token):
    path = _manifest_path(token)
    if not os.path.isfile(path):
        return None
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def _norm_name(s):
    """Нормалізує ПІБ для порівняння: без регістру, без розширення
    файлу, без пунктуації і зайвих пробілів/підкреслень."""
    s = os.path.splitext(str(s or ''))[0]
    s = s.replace('_', ' ').replace('-', ' ')
    s = s.replace('ʼ', "'").replace('’', "'")
    s = re.sub(r"[^\w\s']", ' ', s.lower(), flags=re.UNICODE)
    return ' '.join(s.split())


def _match_filename_to_student(filename, db_students):
    """
    db_students: список dict {'id', 'full_name'}.
    Повертає (student_id | None, score, status) - status: 'auto'
    (впевнений збіг), 'suggest' (найкраще припущення, перевірити) або
    'none' (не знайдено).
    """
    norm = _norm_name(filename)
    if not norm:
        return None, 0.0, 'none'

    best_id, best_score = None, 0.0
    for st in db_students:
        cand = _norm_name(st['full_name'])
        score = difflib.SequenceMatcher(None, norm, cand).ratio()
        # Також перевіряємо збіг у "зворотньому" порядку (Ім'я Прізвище
        # замість Прізвище Ім'я) - файли іноді називають по-різному.
        cand_parts = cand.split()
        if len(cand_parts) >= 2:
            reversed_cand = ' '.join([cand_parts[1], cand_parts[0]] + cand_parts[2:])
            score = max(score, difflib.SequenceMatcher(None, norm, reversed_cand).ratio())
        if score > best_score:
            best_id, best_score = st['id'], score

    if best_score >= 0.85:
        return best_id, best_score, 'auto'
    if best_score >= 0.55:
        return best_id, best_score, 'suggest'
    return None, best_score, 'none'


@photo_bulk_bp.route('/admin/photos/bulk', methods=['GET', 'POST'])
@permission_required('manage_students')
def bulk_upload():
    """Крок 1: вибір групи і завантаження одразу декількох файлів фото."""
    _cleanup_old_sessions()
    conn = get_db()
    role = session.get('role')
    group_ids = session.get('group_ids', [])

    if role == 'admin':
        groups = conn.execute("""
            SELECT id, name, start_year, study_form,
                   name || ' (' || start_year || ', ' || study_form || ')' AS display_name
            FROM groups WHERE archived = FALSE ORDER BY name, start_year
        """).fetchall()
    else:
        placeholders = ','.join('?' for _ in group_ids) if group_ids else "''"
        groups = conn.execute(f"""
            SELECT id, name, start_year, study_form,
                   name || ' (' || start_year || ', ' || study_form || ')' AS display_name
            FROM groups WHERE id IN ({placeholders}) ORDER BY name, start_year
        """, group_ids).fetchall() if group_ids else []

    if request.method == 'POST':
        group_id = request.form.get('group_id', type=int)
        files = request.files.getlist('photo_files')

        if not group_id:
            conn.close()
            flash('Оберіть групу', 'error')
            return redirect(url_for('photo_bulk.bulk_upload'))
        if role != 'admin' and group_id not in group_ids:
            conn.close()
            flash('Ви не маєте доступу до цієї групи', 'error')
            return redirect(url_for('photo_bulk.bulk_upload'))
        if not files or all(f.filename == '' for f in files):
            conn.close()
            flash('Оберіть хоча б один файл фотографії', 'error')
            return redirect(url_for('photo_bulk.bulk_upload'))

        db_students = [dict(r) for r in conn.execute("""
            SELECT id, TRIM(last_name_UA || ' ' || first_name_UA) AS full_name
            FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0
        """, (group_id,)).fetchall()]
        conn.close()

        if not db_students:
            flash('У цій групі немає студентів', 'warning')
            return redirect(url_for('photo_bulk.bulk_upload'))

        token = uuid.uuid4().hex
        os.makedirs(_session_dir(token), exist_ok=True)

        items = []
        skipped = []
        for f in files:
            if f.filename == '':
                continue
            file_bytes = f.read()
            try:
                img = load_and_validate_image(file_bytes)
            except ValueError as e:
                skipped.append({'filename': f.filename, 'error': str(e)})
                continue

            file_key = uuid.uuid4().hex
            img.save(_original_path(token, file_key), 'JPEG', quality=95)

            crop_box = auto_center_crop_box(img.width, img.height)
            preview_img = crop_and_resize(img, crop_box)
            preview_img.save(_preview_path(token, file_key), 'JPEG', quality=85)

            student_id, score, status = _match_filename_to_student(f.filename, db_students)
            items.append({
                'file_key': file_key,
                'filename': f.filename,
                'crop_box': list(crop_box),
                'orig_w': img.width, 'orig_h': img.height,
                'student_id': student_id,
                'score': round(score, 2),
                'status': status,
            })

        _save_manifest(token, {
            'group_id': group_id,
            'photos': items,
            'skipped': skipped,
            'user_id': session.get('user_id'),
        })

        if skipped:
            flash(f"{len(skipped)} файл(ів) пропущено через помилки формату/розміру - деталі на наступному кроці", 'warning')

        if not items:
            flash('Жоден файл не вдалося обробити', 'danger')
            return redirect(url_for('photo_bulk.bulk_upload'))

        return redirect(url_for('photo_bulk.match', token=token))

    conn.close()
    return render_template('photo_bulk_upload.html', groups=groups)


@photo_bulk_bp.route('/admin/photos/bulk/<token>/match', methods=['GET', 'POST'])
@permission_required('manage_students')
def match(token):
    """Крок 2: зіставлення файлів зі студентами (авто + вручну)."""
    manifest = _load_manifest(token)
    if not manifest:
        flash('Сесію завантаження не знайдено або вона застаріла - почніть заново', 'warning')
        return redirect(url_for('photo_bulk.bulk_upload'))

    conn = get_db()
    db_students = [dict(r) for r in conn.execute("""
        SELECT id, TRIM(last_name_UA || ' ' || first_name_UA) AS full_name
        FROM students WHERE group_id = ? AND COALESCE(archived, 0) = 0
        ORDER BY last_name_UA
    """, (manifest['group_id'],)).fetchall()]
    group = conn.execute("SELECT name FROM groups WHERE id = ?", (manifest['group_id'],)).fetchone()
    conn.close()

    if request.method == 'POST':
        valid_ids = {s['id'] for s in db_students}
        for item in manifest['photos']:
            sid = request.form.get(f"student_{item['file_key']}", type=int)
            item['student_id'] = sid if sid in valid_ids else None
        _save_manifest(token, manifest)
        return redirect(url_for('photo_bulk.preview', token=token))

    return render_template(
        'photo_bulk_match.html',
        token=token, manifest=manifest, group=group, db_students=db_students,
    )


@photo_bulk_bp.route('/admin/photos/bulk/<token>/preview')
@permission_required('manage_students')
def preview(token):
    """Крок 3: попередній перегляд обрізаних фото з можливістю
    відредагувати обрізку кожного окремо перед підтвердженням."""
    manifest = _load_manifest(token)
    if not manifest:
        flash('Сесію завантаження не знайдено або вона застаріла - почніть заново', 'warning')
        return redirect(url_for('photo_bulk.bulk_upload'))

    conn = get_db()
    names = {r['id']: r['full_name'] for r in conn.execute("""
        SELECT id, TRIM(last_name_UA || ' ' || first_name_UA) AS full_name
        FROM students WHERE group_id = ?
    """, (manifest['group_id'],)).fetchall()}
    conn.close()

    matched = [i for i in manifest['photos'] if i['student_id']]
    unmatched = [i for i in manifest['photos'] if not i['student_id']]
    for i in matched:
        i['student_name'] = names.get(i['student_id'], '?')

    return render_template(
        'photo_bulk_preview.html',
        token=token, matched=matched, unmatched=unmatched, skipped=manifest.get('skipped', []),
    )


@photo_bulk_bp.route('/admin/photos/bulk/<token>/recrop/<file_key>', methods=['GET', 'POST'])
@permission_required('manage_students')
def recrop(token, file_key):
    """
    GET: віддає оригінальне зображення (для показу в Cropper.js на
        кроці попереднього перегляду).
    POST: приймає нові координати обрізки (crop_x/y/w/h), перераховує і
        зберігає новий preview-файл. AJAX, повертає JSON.
    """
    manifest = _load_manifest(token)
    if not manifest:
        return jsonify({'error': 'session_not_found'}), 404

    item = next((i for i in manifest['photos'] if i['file_key'] == file_key), None)
    if not item:
        return jsonify({'error': 'file_not_found'}), 404

    if request.method == 'GET':
        from flask import send_file
        return send_file(_original_path(token, file_key), mimetype='image/jpeg')

    try:
        crop_box = (
            float(request.form['crop_x']), float(request.form['crop_y']),
            float(request.form['crop_w']), float(request.form['crop_h']),
        )
    except (KeyError, ValueError):
        return jsonify({'error': 'invalid_crop'}), 400

    from PIL import Image
    img = Image.open(_original_path(token, file_key))
    preview_img = crop_and_resize(img, crop_box)
    preview_img.save(_preview_path(token, file_key), 'JPEG', quality=85)

    item['crop_box'] = list(crop_box)
    _save_manifest(token, manifest)

    return jsonify({'ok': True, 'preview_url': url_for('photo_bulk.preview_image', token=token, file_key=file_key, v=uuid.uuid4().hex)})


@photo_bulk_bp.route('/admin/photos/bulk/<token>/preview_image/<file_key>')
@permission_required('manage_students')
def preview_image(token, file_key):
    """Віддає поточний (обрізаний) preview-файл конкретного фото."""
    from flask import send_file
    path = _preview_path(token, file_key)
    if not os.path.isfile(path):
        return '', 404
    return send_file(path, mimetype='image/jpeg')


@photo_bulk_bp.route('/admin/photos/bulk/<token>/confirm', methods=['POST'])
@permission_required('manage_students')
def confirm(token):
    """Крок 4: остаточне збереження всіх зіставлених фото."""
    manifest = _load_manifest(token)
    if not manifest:
        flash('Сесію завантаження не знайдено або вона застаріла', 'warning')
        return redirect(url_for('photo_bulk.bulk_upload'))

    conn = get_db()
    valid_ids = {r['id'] for r in conn.execute(
        "SELECT id FROM students WHERE group_id = ?", (manifest['group_id'],)
    ).fetchall()}

    from PIL import Image
    saved = 0
    for item in manifest['photos']:
        sid = item.get('student_id')
        if not sid or sid not in valid_ids:
            continue
        img = Image.open(_original_path(token, item['file_key']))
        final_img = crop_and_resize(img, item['crop_box'])
        rel_path = save_final_image(final_img, sid)
        conn.execute("UPDATE students SET photo = ? WHERE id = ?", (rel_path, sid))
        saved += 1

    conn.commit()
    group_name_row = conn.execute("SELECT name FROM groups WHERE id = ?", (manifest['group_id'],)).fetchone()
    conn.close()

    log_action(
        current_username(),
        f"масове завантаження фото: {group_name_row['name'] if group_name_row else manifest['group_id']}",
        group_ids=[manifest['group_id']],
        details=f"збережено {saved} фото"
    )

    # Прибираємо тимчасові файли сесії
    import shutil
    shutil.rmtree(_session_dir(token), ignore_errors=True)

    flash(f"Збережено {saved} фото", 'success')
    return redirect(url_for('students.student_list'))
