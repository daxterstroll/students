# -*- coding: utf-8 -*-
"""
Обробка фото студента: обрізка (вручну через Cropper.js або автоматично
по центру для масового завантаження) до стандартного розміру фото на
документи (3х4 см).
"""
import io
import os
import uuid

from PIL import Image, ImageOps

TARGET_WIDTH = 600
TARGET_HEIGHT = 800
TARGET_RATIO = TARGET_WIDTH / TARGET_HEIGHT  # 0.75 (3:4)

ALLOWED_FORMATS = {'JPEG', 'PNG', 'WEBP'}
MAX_FILE_SIZE_BYTES = 8 * 1024 * 1024  # 8 МБ

PHOTOS_DIR = os.path.join(os.getcwd(), 'static', 'uploads', 'photos')


def _ensure_dir():
    os.makedirs(PHOTOS_DIR, exist_ok=True)


def photo_path_for_student(student_id):
    return os.path.join(PHOTOS_DIR, f"student_{student_id}.jpg")


def delete_photo(student_id):
    path = photo_path_for_student(student_id)
    if os.path.exists(path):
        os.remove(path)
        return True
    return False


def load_and_validate_image(file_bytes):
    if len(file_bytes) > MAX_FILE_SIZE_BYTES:
        size_mb = round(len(file_bytes) / 1024 / 1024, 1)
        max_mb = MAX_FILE_SIZE_BYTES // 1024 // 1024
        raise ValueError(f"Файл завеликий ({size_mb} МБ) - максимальний розмір {max_mb} МБ.")
    try:
        img = Image.open(io.BytesIO(file_bytes))
        img_format = img.format
        img = ImageOps.exif_transpose(img)
    except Exception as e:
        raise ValueError(f"Не вдалося розпізнати зображення: {e}")
    if img_format not in ALLOWED_FORMATS:
        raise ValueError(
            f"Формат {img_format or 'невідомий'} не підтримується. "
            f"Дозволені формати: JPEG, PNG, WEBP."
        )
    return img.convert("RGB")


def auto_center_crop_box(img_width, img_height):
    current_ratio = img_width / img_height
    if current_ratio > TARGET_RATIO:
        crop_h = img_height
        crop_w = round(crop_h * TARGET_RATIO)
    else:
        crop_w = img_width
        crop_h = round(crop_w / TARGET_RATIO)
    x = (img_width - crop_w) // 2
    y = (img_height - crop_h) // 2
    return (x, y, crop_w, crop_h)


def crop_and_resize(img, crop_box):
    x, y, w, h = crop_box
    x, y, w, h = round(x), round(y), round(w), round(h)
    x = max(0, min(x, img.width - 1))
    y = max(0, min(y, img.height - 1))
    w = max(1, min(w, img.width - x))
    h = max(1, min(h, img.height - y))
    cropped = img.crop((x, y, x + w, y + h))
    return cropped.resize((TARGET_WIDTH, TARGET_HEIGHT), Image.LANCZOS)


def save_final_image(final_img, student_id):
    _ensure_dir()
    dest_path = photo_path_for_student(student_id)
    tmp_path = dest_path + f".{uuid.uuid4().hex}.tmp"
    final_img.save(tmp_path, "JPEG", quality=92)
    os.replace(tmp_path, dest_path)
    return f"uploads/photos/student_{student_id}.jpg"


def process_and_save_photo(file_bytes, crop_box, student_id):
    img = load_and_validate_image(file_bytes)
    final_img = crop_and_resize(img, crop_box)
    return save_final_image(final_img, student_id)


PENDING_PHOTOS_DIR = os.path.join(os.getcwd(), 'static', 'uploads', 'pending_photos')


def _ensure_pending_dir():
    os.makedirs(PENDING_PHOTOS_DIR, exist_ok=True)


def process_and_save_pending_photo(file_bytes):
    img = load_and_validate_image(file_bytes)
    crop_box = auto_center_crop_box(img.width, img.height)
    final_img = crop_and_resize(img, crop_box)
    _ensure_pending_dir()
    filename = f"{uuid.uuid4().hex}.jpg"
    dest_path = os.path.join(PENDING_PHOTOS_DIR, filename)
    tmp_path = dest_path + ".tmp"
    final_img.save(tmp_path, "JPEG", quality=92)
    os.replace(tmp_path, dest_path)
    return f"uploads/pending_photos/{filename}"


def process_and_save_pending_photo_with_crop(file_bytes, crop_box, old_rel_path=None):
    img = load_and_validate_image(file_bytes)
    final_img = crop_and_resize(img, crop_box)
    _ensure_pending_dir()
    filename = f"{uuid.uuid4().hex}.jpg"
    dest_path = os.path.join(PENDING_PHOTOS_DIR, filename)
    tmp_path = dest_path + ".tmp"
    final_img.save(tmp_path, "JPEG", quality=92)
    os.replace(tmp_path, dest_path)

    if old_rel_path:
        old_abs_path = os.path.join('static', old_rel_path)
        if os.path.exists(old_abs_path):
            try:
                os.remove(old_abs_path)
            except OSError:
                pass
    return f"uploads/pending_photos/{filename}"


# --------------------------------------------------------------------
# Розумна обрізка фото на документи (3×4) з пошуком обличчя (OpenCV).
# --------------------------------------------------------------------

_face_cascade = None


def _get_face_cascade():
    global _face_cascade
    if _face_cascade is not None:
        return _face_cascade
    try:
        import cv2
        cascade_path = os.path.join(
            cv2.data.haarcascades,
            'haarcascade_frontalface_default.xml',
        )
        if not os.path.exists(cascade_path):
            return None
        _face_cascade = cv2.CascadeClassifier(cascade_path)
        return _face_cascade
    except Exception:
        return None


def smart_crop_to_id_photo(img):
    """
    Обрізка 3×4 навколо обличчя (як на фото на документи).
    """
    try:
        import cv2
        import numpy as _np
    except ImportError:
        return auto_center_crop_box(img.width, img.height)

    cascade = _get_face_cascade()
    if cascade is None:
        return auto_center_crop_box(img.width, img.height)

    arr = _np.array(img)
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    h, w = gray.shape

    faces = cascade.detectMultiScale(
        gray, scaleFactor=1.1, minNeighbors=5,
        minSize=(80, 80), flags=cv2.CASCADE_SCALE_IMAGE,
    )
    if len(faces) == 0:
        return auto_center_crop_box(img.width, img.height)

    fx, fy, fw, fh = max(faces, key=lambda r: r[2] * r[3])
    face_cx = fx + fw // 2

    crop_h = int(fh * 2.0)
    crop_w = int(crop_h * TARGET_RATIO)

    x = face_cx - crop_w // 2
    y = fy - int(fh * 0.3)

    if y < 0:
        y = 0
    if y + crop_h > img.height:
        if y > 0:
            y = max(0, img.height - crop_h)
        if y + crop_h > img.height:
            crop_h = img.height - y

    if x < 0:
        x = 0
    if x + crop_w > img.width:
        if x > 0:
            x = max(0, img.width - crop_w)
        if x + crop_w > img.width:
            crop_w = img.width - x

    current_ratio = crop_w / crop_h if crop_h else 0
    if abs(current_ratio - TARGET_RATIO) > 0.01:
        if current_ratio > TARGET_RATIO:
            new_w = int(crop_h * TARGET_RATIO)
            x += (crop_w - new_w) // 2
            crop_w = new_w
        else:
            new_h = int(crop_w / TARGET_RATIO)
            y += (crop_h - new_h) // 2
            crop_h = new_h

    return (x, y, crop_w, crop_h)


def process_and_save_pending_photo_smart(file_bytes):
    img = load_and_validate_image(file_bytes)
    crop_box = smart_crop_to_id_photo(img)
    final_img = crop_and_resize(img, crop_box)

    _ensure_pending_dir()
    filename = f"{uuid.uuid4().hex}.jpg"
    dest_path = os.path.join(PENDING_PHOTOS_DIR, filename)
    tmp_path = dest_path + ".tmp"
    final_img.save(tmp_path, "JPEG", quality=92)
    os.replace(tmp_path, dest_path)
    return f"uploads/pending_photos/{filename}"


def save_raw_pending_photo(file_bytes):
    """
    Зберігає «сире» зображення у pending_photos БЕЗ обрізки — для
    подальшого ручного кропу через Cropper.js.
    """
    img = load_and_validate_image(file_bytes)
    _ensure_pending_dir()
    filename = f"{uuid.uuid4().hex}_raw.jpg"
    dest_path = os.path.join(PENDING_PHOTOS_DIR, filename)
    tmp_path = dest_path + ".tmp"
    img.save(tmp_path, "JPEG", quality=95)
    os.replace(tmp_path, dest_path)
    return f"uploads/pending_photos/{filename}"