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
 
def normalize_passport_number(number, document_type):
    """
    Додає провідні нулі до номера паспорта, якщо Excel/користувач їх
    втратив (Excel зберігає "039628" як число 39628 — нулі зникають).

    Правила:
      - Паспорт-книжка: очікується 6 цифр (39628 → 039628)
      - ID-картка / ID-картка без чипа: 9 цифр (39628 → 000039628)
      - Свідоцтво про народження: НЕ чіпаємо (там формат не фіксований,
        часто містить літери)
      - Якщо в номері є не лише цифри (напр. "12345А", "АБ-123") —
        НЕ чіпаємо: це не стандартний числовий номер
      - Якщо номер вже довший за очікувану довжину — НЕ чіпаємо
        (це не наш формат, хай буде як є)

    Приклади:
      normalize_passport_number('39628', 'Паспорт громадянина України')
        → '039628'
      normalize_passport_number('39628', 'Паспорт громадянина України з безконтактним електронним носієм')
        → '000039628'
      normalize_passport_number('039628', 'Паспорт громадянина України')
        → '039628'  (не чіпаємо — вже правильно)
      normalize_passport_number('12345А', 'Паспорт громадянина України')
        → '12345А'  (є літера — не чіпаємо)
      normalize_passport_number('12345678', 'Паспорт громадянина України')
        → '12345678'  (довший за 6 — не чіпаємо)
      normalize_passport_number('1234', 'Свідоцтво про народження')
        → '1234'  (свідоцтво — не чіпаємо)
    """
    if not number or not document_type:
        return number

    s = str(number).strip()
    if not s:
        return number

    # Тільки цифри — інакше залишаємо як є
    if not s.isdigit():
        return s

    if document_type == PASSPORT_TYPE_BOOK:
        target_len = 6
    elif document_type in PASSPORT_TYPES_ID_STYLE:
        target_len = 9
    else:
        # Свідоцтво про народження або невідомий тип — не чіпаємо
        return s

    if len(s) < target_len:
        return s.zfill(target_len)
    return s