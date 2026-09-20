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
from datetime import datetime

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
        ), 413


@public_apply_bp.route('/apply', methods=['GET', 'POST'])
def apply():
    if request.method == 'GET':
        return render_template('public_apply.html')

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
        )

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
            return render_template('public_apply.html', error=f"Фото: {e}", form=request.form)

    scan_files = [sf for sf in request.files.getlist('document_scans') if sf and sf.filename]
    if len(scan_files) > MAX_SCAN_FILES:
        return render_template(
            'public_apply.html',
            error=f"Забагато файлів сканів (максимум {MAX_SCAN_FILES}).",
            form=request.form,
        )
    for sf in scan_files:
        ext = os.path.splitext(sf.filename)[1].lower()
        if ext not in ALLOWED_SCAN_EXTENSIONS:
            return render_template(
                'public_apply.html',
                error=f"Файл «{sf.filename}»: непідтримуваний формат. Дозволені: JPG, PNG, WEBP, PDF.",
                form=request.form,
            )
        sf.seek(0, os.SEEK_END)
        size = sf.tell()
        sf.seek(0)
        if size > MAX_SCAN_SIZE_BYTES:
            size_mb = round(size / 1024 / 1024, 1)
            return render_template(
                'public_apply.html',
                error=f"Файл «{sf.filename}» завеликий ({size_mb} МБ, максимум {MAX_SCAN_SIZE_BYTES // 1024 // 1024} МБ).",
                form=request.form,
            )

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
            military_registration_number_drpvr, military_registration_document, military_issued_vod,
            military_specialty_number, military_rank, military_being_registered, military_address,
            military_change_credentials, military_change_reason
        ) VALUES (?, ?,?,?,?,?,?, ?,?,?, ?,?,?,?,?,?, ?,?,?,?,?,?,?,?, ?,?,?, ?,?,?,?,?,?,?,?,?)
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
    # прив'язані поки що до заявки (entity_type='pending_student');
    # адмін при підтвердженні "переприв'язує" їх до реального студента.
    if scan_files:
        save_multiple_attachments(conn, 'pending_student', pending_id, scan_files, 'pending_students')

    conn.commit()
    conn.close()

    return render_template('public_apply_thanks.html')
