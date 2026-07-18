"""
routes/office_editor.py
========================
Інтеграція з ONLYOFFICE Document Server: перегляд і редагування
згенерованих .docx файлів прямо в браузері замість негайного
завантаження "наосліп".

Використовується двома місцями:
  - students.generate       (routes/students.py)  - один документ студента
  - admin.generate_group_docs (routes/admin.py)   - масова генерація (кілька файлів)

Як це працює для ОДНОГО файлу:
  1. Виклик gen_doc(...) як і раніше створює .docx на диску (тепер - у
     папці office_sessions/, а не одразу в generated_docs/).
  2. create_editing_session(...) реєструє цей файл під випадковим
     doc_id і повертає його.
  3. Маршрут перенаправляє користувача на /office/edit/<doc_id>.
  4. Ця сторінка підключає JS ONLYOFFICE (DocsAPI) і показує редактор
     прямо в браузері.
  5. Document Server (окремий сервіс на цьому ж сервері) сам звертається
     до нас за файлом: GET /office/file/<doc_id>. Це "серверний" виклик
     від Document Server до Flask, браузер користувача в ньому не бере
     участі.
  6. Коли користувач зберігає документ (Ctrl+S, або форс-збереження при
     закритті вкладки - див. customization.forcesave в конфігу),
     Document Server надсилає нам POST /office/callback/<doc_id> з
     посиланням на нову версію файлу - ми довантажуємо її і
     перезаписуємо файл на диску.
  7. Кнопка "Завершити і завантажити" на сторінці редактора веде на
     /office/finalize/<doc_id>. Перед видачею файлу ми САМІ просимо
     Document Server примусово зберегти документ (Command Service,
     команда forcesave - той самий ефект, що Ctrl+S), чекаємо, поки
     callback перезапише файл на диску, і лише тоді віддаємо його
     користувачу. Тобто пам'ятати про Ctrl+S не потрібно.

Для МАСОВОЇ генерації (group_export) усі документи групи реєструються
з однаковим batch_id (routes/admin.py), а /office/batch/<batch_id>/zip
збирає їх у ZIP на льоту - перед пакуванням кожен документ теж
проходить форс-збереження, тож у ZIP потрапляють актуальні версії.

ВАЖЛИВО: цей модуль писався без доступу до реального встановленого
Document Server (він ставиться окремо на вашому Windows-сервері) - тож
перед реальним використанням варто пройти чек-лист з ONLYOFFICE_SETUP.md
і протестувати весь ланцюжок "відкрити -> редагувати -> зберегти" наживо.
Найімовірніші місця, які може знадобитися підправити під конкретну
версію Document Server: назва поля з токеном у callback-запиті та
формат обгортки JWT (усе позначено коментарями "ONLYOFFICE-SPECIFIC"
нижче).
"""

import os
import time
import uuid
import threading
from zipfile import ZipFile

import jwt
import requests
from flask import (
    Blueprint, render_template, request, jsonify,
    send_file, abort, session, flash, redirect, url_for
)

from routes.utils import logger
from routes.helpers import current_username
from routes.config import (
    ONLYOFFICE_PUBLIC_URL,
    ONLYOFFICE_CALLBACK_BASE_URL,
    ONLYOFFICE_JWT_SECRET,
    ONLYOFFICE_VERIFY_SSL,
    ONLYOFFICE_COMMAND_URL,
)

if not ONLYOFFICE_VERIFY_SSL:
    # Придушуємо попередження urllib3 про самопідписаний сертифікат -
    # ми свідомо вимкнули перевірку (див. ONLYOFFICE_VERIFY_SSL у
    # config.py), тому попередження при кожному збереженні документа
    # лише засмічувало б консоль/лог.
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

office_bp = Blueprint('office', __name__, url_prefix='/office')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SESSIONS_DIR = os.path.join(BASE_DIR, 'office_sessions')
os.makedirs(SESSIONS_DIR, exist_ok=True)

# ----------------------------------------------------------------------
# Реєстр активних сесій редагування - у пам'яті процесу.
#
# ОБМЕЖЕННЯ: цього достатньо, поки застосунок працює як ОДИН процес
# waitress (типовий запуск `serve(app, ...)` - саме так налаштовано у
# app.py). Якщо колись переведете застосунок на кілька робочих процесів
# (наприклад, gunicorn з --workers > 1), різні процеси матимуть РІЗНІ
# копії цього словника в пам'яті і не будуть бачити сесії одне одного -
# у такому разі цей реєстр потрібно перенести в спільне сховище (файл
# JSON на диску, SQLite-таблицю чи Redis).
# ----------------------------------------------------------------------
_SESSIONS = {}          # doc_id -> {...}
_BATCHES = {}           # batch_id -> [doc_id, ...] (порядок для сторінки перегляду)
_BATCH_NAMES = {}       # batch_id -> назва пакету (напр. назва групи) для заголовка сторінки
_LOCK = threading.Lock()

SESSION_TTL_SECONDS = 6 * 3600  # 6 годин - забуті/покинуті сесії прибираються автоматично

JWT_ALGORITHM = 'HS256'
DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'


# ============================================================
# Внутрішні допоміжні функції
# ============================================================

def _cleanup_expired_sessions():
    """Видаляє прострочені сесії редагування та їхні тимчасові файли з диска."""
    now = time.time()
    with _LOCK:
        expired = [doc_id for doc_id, s in _SESSIONS.items() if now - s['created_at'] > SESSION_TTL_SECONDS]
        for doc_id in expired:
            s = _SESSIONS.pop(doc_id, None)
            if not s:
                continue
            batch_id = s.get('batch_id')
            if batch_id and batch_id in _BATCHES:
                _BATCHES[batch_id] = [d for d in _BATCHES[batch_id] if d != doc_id]
                if not _BATCHES[batch_id]:
                    _BATCHES.pop(batch_id, None)
                    _BATCH_NAMES.pop(batch_id, None)
            try:
                if os.path.exists(s['path']):
                    os.remove(s['path'])
            except OSError as e:
                logger.warning(f"Не вдалося видалити прострочений файл сесії редагування {s['path']}: {e}")


def _sign_jwt(payload: dict):
    """Підписує конфіг редактора спільним секретом (якщо він заданий)."""
    if not ONLYOFFICE_JWT_SECRET:
        return None
    return jwt.encode(payload, ONLYOFFICE_JWT_SECRET, algorithm=JWT_ALGORITHM)


def _verify_document_server_request(body: dict):
    """
    Перевіряє, що вхідний запит (callback) дійсно надійшов від Document
    Server, підписаного нашим спільним секретом, а не від будь-кого, хто
    просто вгадав URL.

    Якщо ONLYOFFICE_JWT_SECRET не задано - перевірку пропускаємо (JWT
    вимкнено з обох боків за замовчуванням на щойно встановленому
    Document Server).

    ONLYOFFICE-SPECIFIC: залежно від версії/налаштувань Document Server,
    підписаний токен приходить або в заголовку Authorization: Bearer ...,
    або прямо в тілі запиту як поле "token". Перевіряємо обидва варіанти.
    Payload токена може бути або самим тілом запиту, або обгорнутий у
    {"payload": {...}} - теж перевіряємо обидва варіанти.
    """
    if not ONLYOFFICE_JWT_SECRET:
        return True, body

    token = None
    auth_header = request.headers.get('Authorization', '')
    if auth_header.startswith('Bearer '):
        token = auth_header[len('Bearer '):]
    elif isinstance(body, dict) and body.get('token'):
        token = body['token']

    if not token:
        logger.error("Callback від Document Server без JWT-токена, хоча ONLYOFFICE_JWT_SECRET задано - відхилено.")
        return False, None

    try:
        decoded = jwt.decode(token, ONLYOFFICE_JWT_SECRET, algorithms=[JWT_ALGORITHM])
    except jwt.InvalidTokenError as e:
        logger.error(f"Недійсний JWT у callback від Document Server: {e}")
        return False, None

    return True, decoded.get('payload', decoded)


def _get_session_or_404(doc_id):
    with _LOCK:
        s = _SESSIONS.get(doc_id)
    if not s:
        abort(404)
    return s


def _check_owner_or_admin(s):
    """Перевіряє, що поточний користувач має право на цю сесію редагування."""
    if 'user_id' not in session:
        abort(401)
    if s['user_id'] != session['user_id'] and not session.get('is_admin'):
        abort(403)


def _forcesave_and_wait(s, doc_id, timeout=15):
    """
    Просить Document Server негайно зберегти поточний стан документа
    (той самий ефект, що Ctrl+S у редакторі) і чекає, поки callback
    (status=6) фактично перезапише файл на диску.

    Повертає True, якщо файл на диску гарантовано актуальний.
    При False просто віддамо останню збережену версію - гірше не стане.

    Коди відповіді Command Service (поле "error"):
        0 - команду прийнято, збереження ініційовано (чекаємо callback)
        1 - документа з таким key немає у редакторі (ніхто не редагує -
            файл на диску і так актуальний)
        4 - незбережених змін немає (вже все збережено)
        6 - невірний JWT (перевірте, що секрет збігається з Document Server)
    """
    event = s['save_event']
    event.clear()

    # ONLYOFFICE-SPECIFIC: для Command Service підписується саме тіло
    # запиту, і токен кладеться в тіло полем "token".
    payload = {"c": "forcesave", "key": s['key']}
    token = _sign_jwt(payload)
    if token:
        payload["token"] = token

    try:
        resp = requests.post(ONLYOFFICE_COMMAND_URL, json=payload,
                             timeout=10, verify=ONLYOFFICE_VERIFY_SSL)
        resp.raise_for_status()
        result = resp.json()
    except Exception as e:
        logger.warning(f"Command Service недоступний, віддаємо останню збережену версію (doc_id={doc_id}): {e}")
        return False

    error = result.get('error')

    if error == 0:
        saved = event.wait(timeout)
        if not saved:
            logger.warning(f"Не дочекалися callback-збереження за {timeout}с (doc_id={doc_id})")
        return saved

    if error in (1, 4):
        # 1 - документ зараз ніхто не редагує; 4 - змін немає.
        # В обох випадках файл на диску вже актуальний.
        return True

    logger.warning(f"Command Service повернув error={error} (doc_id={doc_id})")
    return False


# ============================================================
# Публічне API для інших маршрутів (students.py / admin.py)
# ============================================================

def create_editing_session(file_path: str, display_name: str, user_id, batch_id: str = None):
    """
    Реєструє нову сесію редагування документа, який уже згенеровано на
    диску (file_path).

    :param file_path: абсолютний шлях до .docx-файлу (файл вже повинен існувати)
    :param display_name: ім'я файлу, яке побачить користувач (напр. "Іваненко_Петро.docx")
    :param user_id: session['user_id'] власника сесії - для перевірки прав доступу при відкритті
    :param batch_id: якщо задано - додає сесію до групи (масова генерація), щоб пізніше
                      зібрати всі файли групи в один ZIP через /office/batch/<batch_id>/zip
    :return: doc_id (str) для посилання /office/edit/<doc_id>
    """
    _cleanup_expired_sessions()
    doc_id = uuid.uuid4().hex
    with _LOCK:
        _SESSIONS[doc_id] = {
            'path': file_path,
            'display_name': display_name,
            'user_id': user_id,
            'created_at': time.time(),
            'key': uuid.uuid4().hex,  # ONLYOFFICE "document key" - унікальний для поточної версії файлу
            'batch_id': batch_id,
            'save_event': threading.Event(),  # сигнал "callback щойно зберіг файл на диск"
        }
        if batch_id:
            _BATCHES.setdefault(batch_id, []).append(doc_id)
    return doc_id


def new_batch_id():
    """Генерує новий ідентифікатор пакету для масової генерації документів."""
    return uuid.uuid4().hex


def set_batch_name(batch_id, name):
    """Зберігає людиночитабельну назву пакету (напр. назву групи) для заголовка сторінки списку."""
    if not batch_id:
        return
    with _LOCK:
        _BATCH_NAMES[batch_id] = name


def get_batch_name(batch_id, default=''):
    """Повертає збережену назву пакету або default, якщо її нема."""
    with _LOCK:
        return _BATCH_NAMES.get(batch_id, default)


def get_batch_items(batch_id):
    """Повертає список [(doc_id, display_name), ...] для збірки ZIP масової генерації."""
    with _LOCK:
        doc_ids = list(_BATCHES.get(batch_id, []))
        return [(d, _SESSIONS[d]['display_name']) for d in doc_ids if d in _SESSIONS]


def get_batch_preview_items(batch_id):
    """
    Повертає список словників для сторінки перегляду масової генерації
    (group_docs_preview.html): [{'doc_id', 'name', 'filename'}, ...].

    'name' - людиночитабельне ім'я студента (без розширення), 'filename' -
    ім'я .docx-файлу. Оскільки в реєстрі сесій зберігається лише
    display_name (ім'я файлу), 'name' відновлюємо з нього, прибравши
    розширення .docx і повернувши підкреслення назад у пробіли.
    """
    with _LOCK:
        doc_ids = list(_BATCHES.get(batch_id, []))
        items = []
        for d in doc_ids:
            s = _SESSIONS.get(d)
            if not s:
                continue
            filename = s['display_name']
            name = filename
            if name.lower().endswith('.docx'):
                name = name[:-len('.docx')]
            name = name.replace('_', ' ')
            items.append({'doc_id': d, 'name': name, 'filename': filename})
        return items


# ============================================================
# Маршрути
# ============================================================

@office_bp.route('/edit/<doc_id>')
def edit(doc_id):
    """Сторінка з вбудованим редактором ONLYOFFICE для одного документа."""
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    s = _get_session_or_404(doc_id)
    _check_owner_or_admin(s)

    file_url = f"{ONLYOFFICE_CALLBACK_BASE_URL}/office/file/{doc_id}"
    callback_url = f"{ONLYOFFICE_CALLBACK_BASE_URL}/office/callback/{doc_id}"

    editor_config = {
        "document": {
            "fileType": "docx",
            "key": s['key'],
            "title": s['display_name'],
            "url": file_url,
            "permissions": {"edit": True, "download": True, "print": True, "chat": False, "comment": False},
        },
        "documentType": "word",
        "editorConfig": {
            "callbackUrl": callback_url,
            "lang": "uk",
            "user": {
                "id": str(s['user_id']),
                "name": current_username(),
            },
            "customization": {
                # forcesave: якщо користувач просто закриє вкладку, не
                # натиснувши явно "Зберегти", Document Server все одно
                # надішле нам callback зі status=6 (force-save) перед
                # закриттям редактора - інакше зміни могли б загубитися.
                "forcesave": True,
            },
        },
    }

    token = _sign_jwt(editor_config)
    if token:
        editor_config["token"] = token

    return render_template(
        'office_editor.html',
        doc_id=doc_id,
        display_name=s['display_name'],
        batch_id=s.get('batch_id'),
        onlyoffice_public_url=ONLYOFFICE_PUBLIC_URL.rstrip('/') + '/',
        editor_config=editor_config,
    )


@office_bp.route('/file/<doc_id>')
def serve_file(doc_id):
    """
    Віддає сам .docx файл Document Server'у.

    Цей маршрут викликає НЕ браузер користувача, а сам Document Server
    (сервер-до-сервера запит на локальній машині) - тому тут навмисно
    немає перевірки сесії (session), лише існування doc_id. Захист
    забезпечується непередбачуваністю doc_id (128-бітний uuid4,
    практично неможливо підібрати) та обмеженим часом життя сесії
    (SESSION_TTL_SECONDS).
    """
    s = _get_session_or_404(doc_id)
    if not os.path.exists(s['path']):
        abort(404)
    return send_file(s['path'], as_attachment=False,
                      download_name=s['display_name'],
                      mimetype=DOCX_MIME,
                      conditional=False)


@office_bp.route('/callback/<doc_id>', methods=['POST'])
def callback(doc_id):
    """
    Callback від Document Server про стан документа. Викликається
    автоматично самим Document Server, формат тіла - стандартний
    ONLYOFFICE callback (https://api.onlyoffice.com/docs/docs-api/additional-api/callback-handler/).

    Поле "status" (найважливіше):
        0 - немає документа з таким key
        1 - документ зараз редагується
        2 - документ готовий для збереження (усі співавтори вийшли з редактора)
        3 - помилка при збереженні документа на боці Document Server
        4 - документ закрито без змін
        6 - force-save під час редагування (напр. користувач закрив вкладку
            або наш /office/finalize ініціював збереження через Command Service)
        7 - помилка force-save

    При status 2 або 6 у тілі приходить поле "url" - тимчасове посилання,
    звідки потрібно завантажити нову версію файлу.

    Document Server ЧЕКАЄ у відповідь JSON {"error": 0} - будь-яка інша
    відповідь (або HTTP-помилка) інтерпретується як "не вдалося
    зберегти", і Document Server покаже користувачу помилку.
    """
    s = _get_session_or_404(doc_id)

    body = request.get_json(force=True, silent=True) or {}

    ok, payload = _verify_document_server_request(body)
    if not ok:
        # error:1 навмисно (а не 0) - якщо це спрацювало через
        # неправильно налаштований спільний секрет під час первинного
        # налаштування, адміну одразу видно помилку в Document Server,
        # а не тиху "начебто успішну" відповідь при тому, що насправді
        # нічого не зберігається.
        logger.error(f"Відхилено callback з невірним підписом (doc_id={doc_id})")
        return jsonify({"error": 1})
    body = payload or body

    status = body.get('status')

    if status in (2, 6):
        download_url = body.get('url')
        if not download_url:
            logger.error(f"Callback status={status} без посилання на файл (doc_id={doc_id})")
            return jsonify({"error": 1})
        try:
            # verify=ONLYOFFICE_VERIFY_SSL: Document Server сам зараз
            # використовує самопідписаний сертифікат (той самий, що і
            # для основного сайту), тому стандартна перевірка SSL тут
            # відхиляє з'єднання (requests не довіряє самопідписаним
            # сертифікатам за замовчуванням, так само як спочатку не
            # довіряв і браузер). Коли на сервері зʼявиться сертифікат
            # від довіреного CA, поставте ONLYOFFICE_VERIFY_SSL=true
            # в .env, щоб повернути повну перевірку.
            resp = requests.get(download_url, timeout=30, verify=ONLYOFFICE_VERIFY_SSL)
            resp.raise_for_status()
            tmp_path = s['path'] + '.tmp'
            with open(tmp_path, 'wb') as f:
                f.write(resp.content)
            os.replace(tmp_path, s['path'])
            logger.info(f"Збережено відредагований документ '{s['display_name']}' (doc_id={doc_id})")
            # Сигналізуємо _forcesave_and_wait (якщо він зараз чекає у
            # /office/finalize або /office/batch/.../zip), що файл на
            # диску вже актуальний і його можна віддавати користувачу.
            s['save_event'].set()
        except Exception as e:
            logger.error(f"Не вдалося зберегти відредагований документ (doc_id={doc_id}): {e}", exc_info=True)
            return jsonify({"error": 1})

    return jsonify({"error": 0})


@office_bp.route('/finalize/<doc_id>')
def finalize(doc_id):
    """Кнопка «Завершити і завантажити» - віддає користувачу поточну (можливо, відредаговану) версію файлу."""
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    s = _get_session_or_404(doc_id)
    _check_owner_or_admin(s)

    # Перед видачею просимо Document Server зберегти всі незбережені
    # зміни (аналог Ctrl+S) і чекаємо, поки callback перезапише файл.
    _forcesave_and_wait(s, doc_id)

    if not os.path.exists(s['path']):
        flash('Файл не знайдено (можливо, сесія редагування застаріла)', 'danger')
        return redirect(url_for('students.student_list'))

    return send_file(s['path'], as_attachment=True,
                      download_name=s['display_name'],
                      mimetype=DOCX_MIME)


@office_bp.route('/batch/<batch_id>')
def batch_view(batch_id):
    """
    Сторінка зі списком усіх документів пакету (масова генерація) -
    та сама, що показується одразу після генерації. Дозволяє
    повернутися до списку з редактора окремого документа і завантажити
    підсумковий ZIP, не перегенеровуючи документи заново.
    """
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    items = get_batch_preview_items(batch_id)
    if not items:
        flash('Пакет документів не знайдено (можливо, сесія застаріла)', 'danger')
        return redirect(url_for('admin.group_export'))

    # Назву групи, збережену при генерації, показуємо в заголовку.
    group_name = get_batch_name(batch_id, default='Згенеровані документи')
    return render_template('group_docs_preview.html', items=items,
                           batch_id=batch_id, group_name=group_name)


@office_bp.route('/batch/<batch_id>/zip')
def download_batch_zip(batch_id):
    """
    Збирає всі документи пакету (масова генерація) в один ZIP - у
    поточному стані кожного файлу (враховуючи будь-які зміни, збережені
    через редактор для окремих студентів).
    """
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    items = get_batch_items(batch_id)
    if not items:
        flash('Пакет документів не знайдено (можливо, сесія застаріла)', 'danger')
        return redirect(url_for('admin.group_export'))

    # Форс-збереження кожного документа пакета, який зараз відкрито в
    # редакторі, щоб у ZIP потрапили актуальні версії.
    for doc_id, _name in items:
        s = _SESSIONS.get(doc_id)
        if s:
            _forcesave_and_wait(s, doc_id, timeout=10)

    zip_path = os.path.join(SESSIONS_DIR, f"batch_{batch_id}.zip")
    with ZipFile(zip_path, 'w') as zipf:
        for doc_id, display_name in items:
            s = _SESSIONS.get(doc_id)
            if s and os.path.exists(s['path']):
                zipf.write(s['path'], arcname=display_name)

    return send_file(zip_path, as_attachment=True, download_name=f"{batch_id}.zip")