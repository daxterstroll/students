# -*- coding: utf-8 -*-
"""
Обробка фото студента: обрізка (вручну через Cropper.js або автоматично
по центру для масового завантаження) до стандартного розміру фото на
документи (3х4 см).
"""
import io
import os
import uuid

from PIL import Image

# Фінальний розмір файлу - співвідношення сторін точно 3:4 (як
# стандартне фото на документи), у роздільності, достатній для чіткого
# друку/вставки в Word-документ у реальному розмірі 3х4 см.
TARGET_WIDTH = 600
TARGET_HEIGHT = 800
TARGET_RATIO = TARGET_WIDTH / TARGET_HEIGHT  # 0.75 (3:4)

# Дозволені формати оригінального завантаженого файлу (те, що реально
# розпізнає Pillow за вмістом файлу, а не лише за розширенням).
ALLOWED_FORMATS = {'JPEG', 'PNG', 'WEBP'}

# Максимальний розмір оригінального файлу до обрізки. Це окреме,
# набагато суворіше обмеження, ніж загальний MAX_CONTENT_LENGTH застосунку
# (32 МБ, розрахований на великі Excel-файли) - фото на документи не
# повинно важити більше кількох мегабайт.
MAX_FILE_SIZE_BYTES = 8 * 1024 * 1024  # 8 МБ

PHOTOS_DIR = os.path.join(os.getcwd(), 'static', 'uploads', 'photos')


def _ensure_dir():
    os.makedirs(PHOTOS_DIR, exist_ok=True)


def photo_path_for_student(student_id):
    """Шлях до файлу фото студента (завжди один файл на студента -
    повторне завантаження просто перезаписує попереднє)."""
    return os.path.join(PHOTOS_DIR, f"student_{student_id}.jpg")


def delete_photo(student_id):
    """Видаляє файл фото студента з диска (якщо він є). Повертає True,
    якщо файл дійсно існував і був видалений, False - якщо файлу й так не було."""
    path = photo_path_for_student(student_id)
    if os.path.exists(path):
        os.remove(path)
        return True
    return False


def load_and_validate_image(file_bytes):
    """
    Перевіряє розмір і формат байтів зображення, повертає відкритий
    Pillow Image (у режимі RGB) або кидає ValueError із зрозумілим
    повідомленням, якщо щось не так.
    """
    if len(file_bytes) > MAX_FILE_SIZE_BYTES:
        size_mb = round(len(file_bytes) / 1024 / 1024, 1)
        max_mb = MAX_FILE_SIZE_BYTES // 1024 // 1024
        raise ValueError(f"Файл завеликий ({size_mb} МБ) - максимальний розмір {max_mb} МБ.")

    try:
        img = Image.open(io.BytesIO(file_bytes))
        img_format = img.format
    except Exception as e:
        raise ValueError(f"Не вдалося розпізнати зображення: {e}")

    if img_format not in ALLOWED_FORMATS:
        raise ValueError(
            f"Формат {img_format or 'невідомий'} не підтримується. "
            f"Дозволені формати: JPEG, PNG, WEBP."
        )

    return img.convert("RGB")


def auto_center_crop_box(img_width, img_height):
    """
    Обчислює найбільший можливий прямокутник 3:4 по центру зображення -
    для автоматичної обрізки без участі людини (масове завантаження).
    Повертає (x, y, w, h).
    """
    current_ratio = img_width / img_height
    if current_ratio > TARGET_RATIO:
        # зображення ширше, ніж треба - обрізаємо по боках
        crop_h = img_height
        crop_w = round(crop_h * TARGET_RATIO)
    else:
        # зображення вище, ніж треба - обрізаємо зверху/знизу
        crop_w = img_width
        crop_h = round(crop_w / TARGET_RATIO)
    x = (img_width - crop_w) // 2
    y = (img_height - crop_h) // 2
    return (x, y, crop_w, crop_h)


def crop_and_resize(img, crop_box):
    """Обрізає img за crop_box (x, y, w, h у пікселях img) і дотягує до
    фінального розміру TARGET_WIDTH x TARGET_HEIGHT. Повертає новий Image."""
    x, y, w, h = crop_box
    x, y, w, h = round(x), round(y), round(w), round(h)
    x = max(0, min(x, img.width - 1))
    y = max(0, min(y, img.height - 1))
    w = max(1, min(w, img.width - x))
    h = max(1, min(h, img.height - y))

    cropped = img.crop((x, y, x + w, y + h))
    return cropped.resize((TARGET_WIDTH, TARGET_HEIGHT), Image.LANCZOS)


def save_final_image(final_img, student_id):
    """Зберігає вже обрізане й дотягнуте до потрібного розміру
    зображення як фото студента (атомарно). Повертає відносний шлях."""
    _ensure_dir()
    dest_path = photo_path_for_student(student_id)
    tmp_path = dest_path + f".{uuid.uuid4().hex}.tmp"
    final_img.save(tmp_path, "JPEG", quality=92)
    os.replace(tmp_path, dest_path)
    return f"uploads/photos/student_{student_id}.jpg"


def process_and_save_photo(file_bytes, crop_box, student_id):
    """
    Ручна обрізка одного фото (сторінка студента, Cropper.js).

    file_bytes: байти оригінального завантаженого зображення.
    crop_box: (x, y, width, height) - прямокутник обрізки в пікселях
        ОРИГІНАЛЬНОГО зображення (координати від Cropper.js).
    student_id: для якого студента.

    Повертає відносний шлях до збереженого файлу або кидає ValueError.
    """
    img = load_and_validate_image(file_bytes)
    final_img = crop_and_resize(img, crop_box)
    return save_final_image(final_img, student_id)
