# -*- coding: utf-8 -*-
"""
Публічна анкета самореєстрації абітурієнта/студента - НЕ вимагає
авторизації, доступна без сесії. Свідомо ІЗОЛЬОВАНА від решти
застосунку:

  - жодного імпорту з routes/admin.py чи routes/students.py - навіть
    випадково неможливо викликати щось звідти;
  - пише виключно в "карантинну" таблицю pending_students, яка не має
    зовнішніх ключів на groups/students (крім resulting_student_id,
    який проставляється вже ПІСЛЯ того, як адмін вручну підтвердив
    заявку на окремій сторінці /admin/pending_students);
  - тут немає жодного GET-маршруту, який читає чи показує вже подані
    заявки - лише порожня форма і "дякуємо" після відправки.

Захист від спаму - мінімальний і навмисно без капчі/токенів:
прихований honeypot-інпут ("сайт"), який людина ніколи не бачить і не
заповнює, а бот - як правило, заповнює. Такі заявки просто ігноруються
(не зберігаються, але користувачу все одно показується "дякуємо", щоб
бот не зрозумів, що його відсіяли).

Файли (фото і скани документів) валідуються ДО будь-якого запису в
БД - якщо файл поганий (завеликий, не той формат), у БД взагалі
нічого не з'являється, форма просто повертається з поясненням помилки.
"""
import os
import re
from datetime import datetime, date

from flask import Blueprint, render_template, request
from routes.db import get_db
from routes.utils import generate_english_name, save_multiple_attachments
from routes.photo import process_and_save_pending_photo

public_apply_bp = Blueprint('public_apply', __name__)

# Скани документів - будь-який тип файлу, окрім фото, тут не потрібен:
# або фотографія сторінки документа, або PDF.
ALLOWED_SCAN_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.webp', '.pdf'}
MAX_SCAN_SIZE_BYTES = 15 * 1024 * 1024  # 15 МБ на файл
MAX_SCAN_FILES = 5

# Гігієна вхідних даних - форма публічна й дивиться в інтернет, тому
# "мінімум, без якого заявка не має сенсу" (перевірка обов'язкових
# полів нижче) сама собою не рятує від абсурдних значень (2 роки,
# 1850 рік, телефон "asdf"). Валідація серверна, не лише HTML5 -
# HTML5-обмеження (type="date"/"tel") легко обійти прямим POST-запитом
# повз браузер, тож єдина надійна перевірка - тут.
MIN_APPLICANT_AGE_YEARS = 14
MAX_APPLICANT_AGE_YEARS = 100
PHONE_RE = re.compile(r'^(?:\+?380|0)\d{9}$')


def _validate_birth_date(ua_date_str):
    """ua_date_str - вже сконвертована дата у форматі ДД.ММ.РРРР
    (див. _iso_to_ua_date). Повертає None, якщо все гаразд, або текст
    помилки."""
    try:
        d = datetime.strptime(ua_date_str, "%d.%m.%Y").date()
    except (ValueError, TypeError):
        return "Некоректна дата народження."
    today = date.today()
    if d > today:
        return "Дата народження не може бути в майбутньому."
    age = (today - d).days / 365.25
    if age < MIN_APPLICANT_AGE_YEARS:
        return f"За вказаною датою народження вік менше {MIN_APPLICANT_AGE_YEARS} років - перевірте дату."
    if age > MAX_APPLICANT_AGE_YEARS:
        return "Перевірте, будь ласка, дату народження - вказаний вік виглядає нереалістично великим."
    return None


def _validate_phone(value, label):
    """value - сирий рядок з форми (ще без нормалізації). Повертає
    None, якщо все гаразд (включно з порожнім - необов'язкові поля
    штибу "резервний телефон" не повинні блокувати відправку), або
    текст помилки."""
    if not value:
        return None
    digits = re.sub(r'[\s\-()]', '', value)
    if not PHONE_RE.match(digits):
        return f"{label}: перевірте формат номера (напр. +380991234567 або 0991234567)."
    return None


def _birth_date_bounds():
    """min/max для <input type="date"> дати народження - клієнтська
    підказка (браузер сам не дасть обрати дату поза межами), сервер
    все одно перевіряє незалежно (_validate_birth_date вище)."""
    today = date.today()

    def _safe_date(year, month, day):
        try:
            return date(year, month, day)
        except ValueError:
            return date(year, month, day - 1)  # 29 лютого у невисокосному році

    return {
        'min_birth_date': _safe_date(today.year - MAX_APPLICANT_AGE_YEARS, today.month, today.day).isoformat(),
        'max_birth_date': _safe_date(today.year - MIN_APPLICANT_AGE_YEARS, today.month, today.day).isoformat(),
    }


def _iso_to_ua_date(value):
    """<input type="date"> завжди повертає "РРРР-ММ-ДД" (стандарт
    HTML5), а решта системи (students.birth_date, completion_date
    тощо) скрізь зберігає дати як "ДД.ММ.РРРР". Без цієї конвертації
    дата народження виглядала б неправильно скрізь, де вона потім
    показується чи використовується (сортування, документи).
    Повертає None, якщо value порожнє або не розпізнається."""
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return value  # на випадок, якщо браузер із якоїсь причини надіслав інший формат - краще зберегти як є, ніж загубити


def _validate_scan_files(files, max_files, label):
    """Спільна валідація для будь-якого набору сканів (документ про
    освіту чи паспорт) - кількість, розширення, розмір кожного файлу.
    Повертає None (усе гаразд) або текст помилки."""
    if len(files) > max_files:
        return f"Забагато файлів{label} (максимум {max_files})."
    for sf in files:
        ext = os.path.splitext(sf.filename)[1].lower()
        if ext not in ALLOWED_SCAN_EXTENSIONS:
            return f"Файл «{sf.filename}»{label}: непідтримуваний формат. Дозволені: JPG, PNG, WEBP, PDF."
        sf.seek(0, os.SEEK_END)
        size = sf.tell()
        sf.seek(0)
        if size > MAX_SCAN_SIZE_BYTES:
            size_mb = round(size / 1024 / 1024, 1)
            return f"Файл «{sf.filename}»{label} завеликий ({size_mb} МБ, максимум {MAX_SCAN_SIZE_BYTES // 1024 // 1024} МБ)."
    return None


@public_apply_bp.before_request
def _limit_request_size():
    """Груба відсічка занадто великих запитів (кілька фото/сканів одразу
    можуть важити багато) - без цього хтось міг би закидати диск сервера
    величезними файлами через публічну, неавторизовану форму."""
    max_total = 80 * 1024 * 1024  # 80 МБ на весь запит (фото + до 5 сканів)
    if request.content_length and request.content_length > max_total:
        return render_template(
            'public_apply.html',
            error="Загальний розмір прикріплених файлів завеликий. Спробуйте завантажити менше файлів або стисніть їх.",
            **_birth_date_bounds(),
        ), 413


@public_apply_bp.route('/apply', methods=['GET', 'POST'])
def apply():
    if request.method == 'GET':
        return render_template('public_apply.html', **_birth_date_bounds())

    # Honeypot: приховане поле, яке людина не бачить і не заповнює.
    # Заповнене - майже напевно бот. Мовчки "приймаємо" (для бота),
    # нічого не зберігаючи.
    if (request.form.get('website') or '').strip():
        return render_template('public_apply_thanks.html')

    def f(name):
        return (request.form.get(name) or '').strip() or None

    last_name_ua = f('last_name_UA')
    first_name_ua = f('first_name_UA')
    birth_date = _iso_to_ua_date(f('birth_date'))
    phone = f('phone')

    # Мінімум, без якого заявка не має сенсу. Решта полів - за
    # бажанням, щоб не відлякувати заповнення з телефону: адмін потім
    # сам зв'яжеться і доуточнить, якщо чогось бракує.
    if not last_name_ua or not first_name_ua or not birth_date or not phone:
        return render_template(
            'public_apply.html',
            error="Заповніть, будь ласка, прізвище, ім'я, дату народження і телефон - без них заявку не можна обробити.",
            form=request.form,
            **_birth_date_bounds(),
        )

    # Гігієна вхідних даних (див. коментар біля констант вище) - дата
    # народження і формат телефонів. Помилка тут - одразу назад у
    # форму з поясненням, ще до жодного запису у БД.
    validation_error = _validate_birth_date(birth_date)
    if not validation_error:
        validation_error = _validate_phone(phone, "Телефон")
    if not validation_error:
        validation_error = _validate_phone(f('phone_backup'), "Резервний телефон")
    if validation_error:
        return render_template('public_apply.html', error=validation_error, form=request.form, **_birth_date_bounds())

    # ---- Валідація файлів ДО будь-якого запису в БД ----
    photo_file = request.files.get('photo')
    photo_bytes = None
    if photo_file and photo_file.filename:
        photo_bytes = photo_file.read()
        try:
            # Лише перевіряємо валідність тут (формат/розмір) - саму
            # обрізку й збереження на диск робимо вже після успішного
            # запису анкети в БД, щоб не лишати "осиротілі" файли,
            # якщо щось інше в запиті виявиться некоректним.
            from routes.photo import load_and_validate_image
            load_and_validate_image(photo_bytes)
        except ValueError as e:
            return render_template('public_apply.html', error=f"Фото: {e}", form=request.form, **_birth_date_bounds())

    scan_files = [sf for sf in request.files.getlist('document_scans') if sf and sf.filename]
    scan_error = _validate_scan_files(scan_files, MAX_SCAN_FILES, " документа про освіту")
    if scan_error:
        return render_template('public_apply.html', error=scan_error, form=request.form, **_birth_date_bounds())

    passport_scan_files = [sf for sf in request.files.getlist('passport_scans') if sf and sf.filename]
    passport_scan_error = _validate_scan_files(passport_scan_files, MAX_SCAN_FILES, " паспорта")
    if passport_scan_error:
        return render_template('public_apply.html', error=passport_scan_error, form=request.form, **_birth_date_bounds())

    # "Видано за кордоном", "скорочена програма" і "перебуває на
    # обліку" студент на публічній формі більше не позначає сам (ці
    # нюанси визначає адмін під час обробки заявки, дивлячись у сам
    # документ) - тому тут завжди 0, а не читання неіснуючих полів форми.
    is_foreign_document = 0
    reduced_program_claim = 0
    military_being_registered = 0

    # Латиницею студент нічого не вводить - транслітерується
    # автоматично з українського написання, той самий принцип, що й
    # при ручному додаванні студента адміном (generate_english_name).
    last_name_eng, first_name_eng = generate_english_name(last_name_ua, first_name_ua)

    document_date = _iso_to_ua_date(f('document_date'))
    military_issued_vod = _iso_to_ua_date(f('military_issued_vod'))

    passport_document_type = f('passport_document_type')
    if passport_document_type not in ('Паспорт (книжка)', 'ID-картка'):
        passport_document_type = None
    passport_issue_date = _iso_to_ua_date(f('passport_issue_date'))
    passport_valid_until = _iso_to_ua_date(f('passport_valid_until'))

    conn = get_db()
    cur = conn.execute("""
        INSERT INTO pending_students (
            submitted_ip,
            last_name_UA, first_name_UA, middle_name_UA, last_name_ENG, first_name_ENG, birth_date,
            phone, phone_backup, email,
            document_type, document_series, document_number, document_institution, document_country, document_date,
            is_foreign_document, foreign_reference_number, foreign_reference_institution,
            foreign_reference_country, foreign_reference_issue_date,
            recognition_certificate_number, recognition_issuer, recognition_date,
            reduced_program_claim, reduced_program_specialty, reduced_program_institution,
            passport_document_type, passport_series, passport_number, passport_issued_by,
            passport_issue_date, passport_valid_until, passport_unique_number,
            military_registration_number_drpvr, military_registration_document, military_issued_vod,
            military_specialty_number, military_rank, military_being_registered, military_address,
            military_change_credentials, military_change_reason
        ) VALUES (?, ?,?,?,?,?,?, ?,?,?, ?,?,?,?,?,?, ?,?,?,?,?,?,?,?, ?,?,?, ?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?)
    """, (
        request.remote_addr,
        last_name_ua, first_name_ua, f('middle_name_UA'), last_name_eng, first_name_eng, birth_date,
        phone, f('phone_backup'), f('email'),
        f('document_type'), f('document_series'), f('document_number'), f('document_institution'),
        f('document_country'), document_date,
        is_foreign_document, None, None,
        None, None,
        None, None, None,
        reduced_program_claim, None, None,
        passport_document_type, f('passport_series'), f('passport_number'), f('passport_issued_by'),
        passport_issue_date, passport_valid_until, f('passport_unique_number'),
        f('military_registration_number_drpvr'), f('military_registration_document'), military_issued_vod,
        f('military_specialty_number'), f('military_rank'), military_being_registered, f('military_address'),
        f('military_change_credentials'), f('military_change_reason'),
    ))
    pending_id = cur.lastrowid

    # Фото - вже провалідоване вище, лишається обрізати по центру й
    # зберегти (окрема "карантинна" папка, не та, де фото реальних
    # студентів).
    if photo_bytes:
        try:
            photo_path = process_and_save_pending_photo(photo_bytes)
            conn.execute("UPDATE pending_students SET photo_path=? WHERE id=?", (photo_path, pending_id))
        except ValueError:
            pass  # малоймовірно (уже провалідовано вище), але не рвати заявку через фото

    # Скани документів - через ту саму спільну таблицю attachments, що
    # й накази/заморозка тощо (routes/utils.save_multiple_attachments),
    # прив'язані поки що до заявки. Документ про освіту і паспорт -
    # РІЗНІ entity_type, щоб при підтвердженні знати, які скани куди
    # переприв'язувати (до education_document чи до passport_document).
    if scan_files:
        save_multiple_attachments(conn, 'pending_student', pending_id, scan_files, 'pending_students')
    if passport_scan_files:
        save_multiple_attachments(conn, 'pending_student_passport', pending_id, passport_scan_files, 'pending_students')

    conn.commit()
    conn.close()

    return render_template('public_apply_thanks.html')
