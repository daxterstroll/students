# -*- coding: utf-8 -*-
"""
Масове завантаження фото студентів без прив'язки до групи.

Студенти шукаються по ВСІХ активних студентах, до яких має доступ
користувач (адмін — усі, працівник — зі своїх груп). Кожен файл
зіставляється зі студентом за іменем у назві файлу
("Прізвище Ім'я.jpg"), фото обрізається автоматично навколо обличчя
до 3×4, з підтримкою сканів із кількома однаковими фото 3×4 на
одному аркуші (обирається одне, найближче до центру). На кроці
попереднього перегляду обрізку можна поправити вручну.
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
    load_and_validate_image, crop_and_resize, save_final_image,
    smart_crop_to_id_photo_with_info,
    MAX_FILE_SIZE_BYTES, ALLOWED_FORMATS,
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
    def _default(o):
        # numpy-скаляри (np.int64, np.float64 тощо) не серіалізуються
        # стандартним json-енкодером, хоча поводяться як числа.
        # Перетворюємо їх у звичайні Python int/float.
        try:
            import numpy as np
            if isinstance(o, np.integer):
                return int(o)
            if isinstance(o, np.floating):
                return float(o)
            if isinstance(o, np.ndarray):
                return o.tolist()
        except ImportError:
            pass
        raise TypeError(f'Object of type {o.__class__.__name__} is not JSON serializable')

    with open(_manifest_path(token), 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, default=_default)


def _load_manifest(token):
    path = _manifest_path(token)
    if not os.path.isfile(path):
        return None
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def _norm_name(s):
    """Нормалізує ПІБ: без розширення, регістру, пунктуації, зайвих
    пробілів та підкреслень."""
    s = os.path.splitext(str(s or ''))[0]
    s = s.replace('_', ' ').replace('-', ' ')
    s = s.replace('ʼ', "'").replace('’', "'")
    s = re.sub(r"[^\w\s']", ' ', s.lower(), flags=re.UNICODE)
    return ' '.join(s.split())


def _build_student_name_index(db_students):
    """
    Передобчислює нормалізовані варіанти ПІБ для кожного студента.
    variants — усі варіанти, з якими порівнюється filename:
      - «Прізвище Ім'я»
      - «Ім'я Прізвище»
      - «Прізвище Ім'я По-батькові»
      - «Ім'я Прізвище По-батькові»
    Порівняння з кожним варіантом через fuzz.token_set_ratio —
    стійко до зайвих слів в обидві сторони (напр. filename
    «Іваненко Іван» → 100% із «Іваненко Іван Петрович» у БД).
    """
    result = []
    for st in db_students:
        last = st.get('last_name_UA') or ''
        first = st.get('first_name_UA') or ''
        middle = st.get('middle_name_UA') or ''

        variants = []
        for parts in (
            (last, first),
            (first, last),
            (last, first, middle),
            (first, last, middle),
        ):
            v = _norm_name(' '.join(p for p in parts if p))
            if v and v not in variants:
                variants.append(v)

        result.append({
            'id': st['id'],
            'full_name': st['full_name'],
            'last': _norm_name(last),
            'first': _norm_name(first),
            'middle': _norm_name(middle),
            'variants': variants,
        })
    return result

def _match_filename_to_student(filename, index):
    """
    Повертає (student_id | None, score, status, candidates).

    status:
      'auto'      — впевнений збіг (1 явний переможець, score >= 0.85)
      'suggest'   — найкраще припущення, перевірте (0.55 <= score < 0.85)
      'ambiguous' — 2+ студенти дають майже однаковий бал (різниця < 5%).
                    student_id = None, candidates = список кандидатів.
                    Вгадувати НЕ МОЖНА: однакові ПІБ у БД — це нормально
                    (тезки, або один студент змінив прізвище — тоді
                    в базі обидва варіанти). Нехай обере людина.
      'none'      — нічого схожого не знайдено (score < 0.55)

    candidates — список [{'id': ..., 'full_name': ...}, ...]
    (порожній, якщо status != 'ambiguous')
    """
    from rapidfuzz import fuzz

    norm = _norm_name(filename)
    if not norm:
        return None, 0.0, 'none', []

    file_tokens = norm.split()
    has_initials = len(file_tokens) >= 2 and any(
        len(t) == 1 for t in file_tokens[1:]
    )

    # Рахуємо найкращий бал для КОЖНОГО студента окремо
    scored = []  # [(score, student_id, full_name)]
    for item in index:
        best_for_item = 0.0
        for variant in item['variants']:
            s = fuzz.token_set_ratio(norm, variant) / 100.0
            if s > best_for_item:
                best_for_item = s

        if has_initials and item['last']:
            surname_score = fuzz.ratio(file_tokens[0], item['last']) / 100.0
            if surname_score >= 0.85:
                file_initials = [t[0].lower() for t in file_tokens[1:] if len(t) == 1]
                expected = []
                if item['first']:
                    expected.append(item['first'][0].lower())
                if item['middle']:
                    expected.append(item['middle'][0].lower())
                if file_initials and file_initials == expected[:len(file_initials)]:
                    if surname_score > best_for_item:
                        best_for_item = surname_score

        if best_for_item > 0:
            scored.append((best_for_item, item['id'], item['full_name']))

    if not scored:
        return None, 0.0, 'none', []

    scored.sort(key=lambda t: t[0], reverse=True)
    best_score, best_id, _best_name = scored[0]

    if best_score < 0.55:
        return None, best_score, 'none', []

    # Неоднозначність: скільки ще студентів мають бал у межах 5% від лідера?
    AMBIGUITY_DELTA = 0.05
    close_calls = [
        (sid, name) for sc, sid, name in scored
        if sc >= best_score - AMBIGUITY_DELTA
    ]

    if len(close_calls) > 1:
        candidates = [{'id': sid, 'full_name': name} for sid, name in close_calls]
        return None, best_score, 'ambiguous', candidates

    status = 'auto' if best_score >= 0.85 else 'suggest'
    return best_id, best_score, status, []
    
def _get_accessible_students(conn, role, group_ids):
    """
    Усі активні студенти, до яких має доступ користувач:
      - адмін -> усі
      - працівник -> студенти зі своїх груп (primary + паралельні)
    Повертає [{'id', 'full_name', 'last_name_UA', 'first_name_UA',
              'middle_name_UA'}, ...], відсортованих за ПІБ.
    """
    if role == 'admin':
        rows = conn.execute("""
            SELECT id, last_name_UA, first_name_UA, middle_name_UA,
                   TRIM(last_name_UA || ' ' || first_name_UA
                        || ' ' || COALESCE(middle_name_UA, '')) AS full_name
            FROM students
            WHERE COALESCE(archived, 0) = 0
            ORDER BY last_name_UA COLLATE UKRAINIAN,
                     first_name_UA COLLATE UKRAINIAN
        """).fetchall()
        return [dict(r) for r in rows]

    if not group_ids:
        return []

    placeholders = ','.join('?' for _ in group_ids)
    rows = conn.execute(f"""
        SELECT DISTINCT s.id, s.last_name_UA, s.first_name_UA,
               s.middle_name_UA,
               TRIM(s.last_name_UA || ' ' || s.first_name_UA
                    || ' ' || COALESCE(s.middle_name_UA, '')) AS full_name
        FROM students s
        LEFT JOIN student_enrollments e
               ON e.student_id = s.id AND e.status = 'active'
        WHERE COALESCE(s.archived, 0) = 0
          AND (s.group_id IN ({placeholders}) OR e.group_id IN ({placeholders}))
        ORDER BY s.last_name_UA COLLATE UKRAINIAN,
                 s.first_name_UA COLLATE UKRAINIAN
    """, list(group_ids) + list(group_ids)).fetchall()
    return [dict(r) for r in rows]

def _validate_student_ids(conn, allowed_ids, requested_ids):
    """Перевіряє, що всі ID зібрані з allowed_ids і ще активні."""
    allowed = set(allowed_ids)
    valid = set()
    for sid in requested_ids:
        if sid in allowed:
            valid.add(sid)
    return valid


# ============================================================
# Крок 1: завантаження файлів
# ============================================================

@photo_bulk_bp.route('/admin/photos/bulk', methods=['GET', 'POST'])
@permission_required('bulk_photos')
def bulk_upload():
    _cleanup_old_sessions()

    if request.method == 'GET':
        return render_template('photo_bulk_upload.html')

    conn = get_db()
    role = session.get('role')
    group_ids = session.get('group_ids', [])

    files = request.files.getlist('photo_files')
    if not files or all(f.filename == '' for f in files):
        conn.close()
        flash('Оберіть хоча б один файл фотографії', 'error')
        return redirect(url_for('photo_bulk.bulk_upload'))

    db_students = _get_accessible_students(conn, role, group_ids)
    conn.close()

    if not db_students:
        flash(
            'Немає доступних студентів — перевірте, що у вас є активні групи, '
            'або зверніться до адміністратора.', 'warning',
        )
        return redirect(url_for('photo_bulk.bulk_upload'))

    name_index = _build_student_name_index(db_students)

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

        # smart_crop_to_id_photo_with_info: обрізає навколо обличчя й
        # повертає, скільки облич знайдено. Якщо 2+ — це, найімовірніше,
        # сканована розкладка з кількох однакових фото 3×4 (обрано одне).
        crop_box, face_count = smart_crop_to_id_photo_with_info(img)
        preview_img = crop_and_resize(img, crop_box)
        preview_img.save(_preview_path(token, file_key), 'JPEG', quality=85)

        student_id, score, status, candidates = _match_filename_to_student(
            f.filename, name_index,
        )
        items.append({
            'file_key': file_key,
            'filename': f.filename,
            'crop_box': [int(v) for v in crop_box],
            'orig_w': img.width, 'orig_h': img.height,
            'student_id': student_id,
            'score': round(score, 2),
            'status': status,
            'candidates': candidates,   # ← НОВЕ: список кандидатів (для ambiguous)
            'face_count': face_count,
        })

    _save_manifest(token, {
        'allowed_student_ids': [s['id'] for s in db_students],
        'photos': items,
        'skipped': skipped,
        'user_id': session.get('user_id'),
    })

    if skipped:
        flash(
            f"{len(skipped)} файл(ів) пропущено через помилки формату/"
            f"розміру — деталі на наступному кроці", 'warning',
        )

    if not items:
        flash('Жоден файл не вдалося обробити', 'danger')
        return redirect(url_for('photo_bulk.bulk_upload'))

    return redirect(url_for('photo_bulk.match', token=token))


# ============================================================
# Крок 2: зіставлення зі студентами
# ============================================================

@photo_bulk_bp.route('/admin/photos/bulk/<token>/match', methods=['GET', 'POST'])
@permission_required('bulk_photos')
def match(token):
    manifest = _load_manifest(token)
    if not manifest:
        flash('Сесію завантаження не знайдено або вона застаріла — почніть заново', 'warning')
        return redirect(url_for('photo_bulk.bulk_upload'))

    allowed_ids = set(manifest.get('allowed_student_ids') or [])
    if not allowed_ids:
        flash('Сесію пошкоджено — почніть заново', 'warning')
        return redirect(url_for('photo_bulk.bulk_upload'))

    conn = get_db()
    placeholders = ','.join('?' for _ in allowed_ids)
    db_students = [dict(r) for r in conn.execute(f"""
        SELECT id,
               TRIM(last_name_UA || ' ' || first_name_UA
                    || ' ' || COALESCE(middle_name_UA, '')) AS full_name
        FROM students
        WHERE id IN ({placeholders}) AND COALESCE(archived, 0) = 0
        ORDER BY last_name_UA COLLATE UKRAINIAN,
                 first_name_UA COLLATE UKRAINIAN
    """, list(allowed_ids)).fetchall()]
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
        token=token, manifest=manifest, db_students=db_students,
    )


# ============================================================
# Крок 3: попередній перегляд + ручна переобрізка
# ============================================================

@photo_bulk_bp.route('/admin/photos/bulk/<token>/preview')
@permission_required('bulk_photos')
def preview(token):
    manifest = _load_manifest(token)
    if not manifest:
        flash('Сесію завантаження не знайдено або вона застаріла — почніть заново', 'warning')
        return redirect(url_for('photo_bulk.bulk_upload'))

    allowed_ids = set(manifest.get('allowed_student_ids') or [])
    conn = get_db()
    if allowed_ids:
        placeholders = ','.join('?' for _ in allowed_ids)
        names = {r['id']: r['full_name'] for r in conn.execute(f"""
            SELECT id, TRIM(last_name_UA || ' ' || first_name_UA
                            || ' ' || COALESCE(middle_name_UA, '')) AS full_name
            FROM students WHERE id IN ({placeholders})
        """, list(allowed_ids)).fetchall()}
    else:
        names = {}
    conn.close()

    matched = [i for i in manifest['photos'] if i['student_id']]
    unmatched = [i for i in manifest['photos'] if not i['student_id']]
    for i in matched:
        i['student_name'] = names.get(i['student_id'], '?')

    return render_template(
        'photo_bulk_preview.html',
        token=token, matched=matched, unmatched=unmatched,
        skipped=manifest.get('skipped', []),
    )


@photo_bulk_bp.route('/admin/photos/bulk/<token>/recrop/<file_key>', methods=['GET', 'POST'])
@permission_required('bulk_photos')
def recrop(token, file_key):
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
    # Ручна переобрізка — користувач сам вирішив, скільки фото лишити,
    # тому face_count для UI-бейджа знімаємо (він уже неактуальний).
    item['face_count'] = None
    _save_manifest(token, manifest)

    return jsonify({
        'ok': True,
        'preview_url': url_for(
            'photo_bulk.preview_image', token=token, file_key=file_key,
            v=uuid.uuid4().hex,
        ),
    })


@photo_bulk_bp.route('/admin/photos/bulk/<token>/preview_image/<file_key>')
@permission_required('bulk_photos')
def preview_image(token, file_key):
    from flask import send_file
    path = _preview_path(token, file_key)
    if not os.path.isfile(path):
        return '', 404
    return send_file(path, mimetype='image/jpeg')


# ============================================================
# Крок 4: збереження
# ============================================================

@photo_bulk_bp.route('/admin/photos/bulk/<token>/confirm', methods=['POST'])
@permission_required('bulk_photos')
def confirm(token):
    manifest = _load_manifest(token)
    if not manifest:
        flash('Сесію завантаження не знайдено або вона застаріла', 'warning')
        return redirect(url_for('photo_bulk.bulk_upload'))

    allowed_ids = set(manifest.get('allowed_student_ids') or [])
    conn = get_db()
    if allowed_ids:
        placeholders = ','.join('?' for _ in allowed_ids)
        valid_ids = {r['id'] for r in conn.execute(f"""
            SELECT id FROM students
            WHERE id IN ({placeholders}) AND COALESCE(archived, 0) = 0
        """, list(allowed_ids)).fetchall()}
    else:
        valid_ids = set()

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
    conn.close()

    log_action(
        current_username(),
        f"масове завантаження фото: збережено {saved} фото",
    )

    import shutil
    shutil.rmtree(_session_dir(token), ignore_errors=True)

    flash(f"Збережено {saved} фото", 'success')
    return redirect(url_for('students.student_list'))