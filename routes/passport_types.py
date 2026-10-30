"""
routes/passport_types.py
========================
Єдине джерело істини для типів паспортного документа.

Раніше в коді було зашито лише 2 значення ('Паспорт (книжка)' /
'ID-картка') у ~10 місцях (Python + 5 шаблонів + CHECK у БД). Будь-яка
зміна списку типів вимагала синхронної правки в усіх цих місцях, і
одне пропущене місце ламало або валідацію, або відображення.

Тепер список типів і їхні "стилі" (ID чи книжковий) зберігаються тут,
решта коду та шаблони імпортують звідси.
"""

# ── 4 типи документів (порядок показу в <select>) ──
PASSPORT_TYPE_BIRTH_CERT = 'Свідоцтво про народження'
PASSPORT_TYPE_BOOK       = 'Паспорт громадянина України'
PASSPORT_TYPE_ID_CHIP    = 'Паспорт громадянина України з безконтактним електронним носієм'
PASSPORT_TYPE_ID_NO_CHIP = 'Паспорт громадянина України, що не містить безконтактного електронного носія'

PASSPORT_TYPES = [
    PASSPORT_TYPE_BIRTH_CERT,
    PASSPORT_TYPE_BOOK,
    PASSPORT_TYPE_ID_CHIP,
    PASSPORT_TYPE_ID_NO_CHIP,
]

# ── "ID-стиль" (без серії; з «Дійсний до» і «Номер запису в ЄДДР») ──
PASSPORT_TYPES_ID_STYLE = {
    PASSPORT_TYPE_ID_CHIP,
    PASSPORT_TYPE_ID_NO_CHIP,
}

# ── "Книжковий стиль" (з серією; без «Дійсний до» і ЄДДР) ──
PASSPORT_TYPES_BOOK_STYLE = {
    PASSPORT_TYPE_BIRTH_CERT,
    PASSPORT_TYPE_BOOK,
}

# ── Тип за замовчуванням для нових записів ──
PASSPORT_TYPE_DEFAULT = PASSPORT_TYPE_ID_CHIP

# ── Застарілі назви → нові (для нормалізації старих даних у заявках) ──
LEGACY_PASSPORT_TYPES = {
    'Паспорт (книжка)': PASSPORT_TYPE_BOOK,
    'ID-картка':        PASSPORT_TYPE_ID_CHIP,
}


def normalize_passport_type(value):
    """
    Перетворює застарілі назви типів ('Паспорт (книжка)' / 'ID-картка')
    на нові. Якщо значення вже нове, порожнє або невідоме — повертає
    як є (без помилки).
    """
    if not value:
        return value
    return LEGACY_PASSPORT_TYPES.get(value, value)


def passport_template_context():
    """
    Готовий словник для `render_template(..., **passport_template_context())`
    або для `@bp.context_processor`. Містить усе, що потрібно шаблонам:
      - passport_types:            список для <select>
      - passport_types_id_style:   перелік "ID-стилю" (для JS)
      - passport_type_default:     тип за замовчуванням
    """
    return {
        'passport_types': PASSPORT_TYPES,
        'passport_types_id_style': list(PASSPORT_TYPES_ID_STYLE),
        'passport_type_default': PASSPORT_TYPE_DEFAULT,
    }