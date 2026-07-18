"""
config.py
=========
Конфігурація застосунку. Усі значення можна перевизначити змінними
середовища - або системними (Windows: "Змінні середовища" в
властивостях системи), або найпростіше - файлом `.env` поруч з app.py
(підтримується автоматично, якщо встановлено пакет python-dotenv;
якщо пакета немає - просто використовуються значення за замовчуванням
нижче, без помилки).

Приклад .env файлу (покласти поруч з app.py, НЕ комітити в git):

    SECRET_KEY=довгий-випадковий-рядок
    ONLYOFFICE_JWT_SECRET=ще-один-довгий-випадковий-рядок
    ONLYOFFICE_PUBLIC_URL=https://ваш-сервер/onlyoffice/
    ONLYOFFICE_CALLBACK_BASE_URL=http://localhost:5000
"""

import logging
import os
import secrets

logger = logging.getLogger('Students')

try:
    from dotenv import load_dotenv
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    load_dotenv(os.path.join(BASE_DIR, '.env'))
except ImportError:
    # python-dotenv не встановлено - це нормально, просто працюємо лише
    # зі змінними середовища ОС. Встановити: pip install python-dotenv
    pass


# --------------------------------------------------------------------
# SECRET_KEY - підпис сесійних cookie.
# --------------------------------------------------------------------
# Якщо ключ витікає або залишається однаковим "за замовчуванням" у всіх
# інсталяціях (як було раніше - захардкоджений рядок прямо в коді
# репозиторію), будь-хто, хто бачив цей код (включно з git-історією),
# може підробити сесію будь-якого користувача, зокрема адміністратора.
#
# Якщо змінна середовища не задана, генерується випадковий ключ на час
# роботи процесу (це означає, що після кожного перезапуску сервера всі
# активні сесії "злітають" - користувачам доведеться залогінитись
# знову, але це набагато безпечніше, ніж один статичний ключ, видимий у
# публічному коді).
#
# Для стабільної роботи в продакшні задайте власний постійний ключ:
#   Windows (PowerShell), для поточного користувача назавжди:
#       [Environment]::SetEnvironmentVariable("SECRET_KEY", "ваш-ключ", "User")
#   або просто додайте рядок SECRET_KEY=... у файл .env поруч з app.py.
#   Згенерувати випадковий ключ:
#       python -c "import secrets; print(secrets.token_hex(32))"
SECRET_KEY = os.environ.get('SECRET_KEY') or secrets.token_hex(32)

if not os.environ.get('SECRET_KEY'):
    logger.warning(
        "Змінна середовища SECRET_KEY не задана - згенеровано тимчасовий "
        "ключ для цього запуску. Усі сесії користувачів будуть скинуті "
        "при наступному перезапуску сервера. Задайте SECRET_KEY у "
        "середовищі (або файлі .env) для стабільної роботи в продакшні."
    )

# Сайт віддається лише по HTTPS (див. nginx.conf) - url_for(..., _external=True)
# та редіректи повинні за замовчуванням генерувати https://, а не http://.
PREFERRED_URL_SCHEME = os.environ.get('PREFERRED_URL_SCHEME', 'https')


# --------------------------------------------------------------------
# ONLYOFFICE Document Server - перегляд/редагування згенерованих .docx
# просто в браузері (routes/office_editor.py).
# --------------------------------------------------------------------
# ONLYOFFICE_PUBLIC_URL - адреса, яку використовує БРАУЗЕР користувача,
# щоб завантажити редактор (JS-бібліотека Document Server). Повинна
# бути доступна ззовні, тому вказуємо публічний HTTPS-шлях, проксійований
# через nginx (див. location /onlyoffice/ у nginx.conf), а не внутрішній
# порт Document Server напряму.
ONLYOFFICE_PUBLIC_URL = os.environ.get('ONLYOFFICE_PUBLIC_URL', 'https://localhost/onlyoffice/')

# ONLYOFFICE_CALLBACK_BASE_URL - адреса, за якою САМ Document Server
# (окремий процес на цьому ж Windows-сервері) звертається до НАШОГО
# Flask-застосунку, щоб забрати файл на редагування і надіслати назад
# збережену версію. Оскільки обидва процеси працюють на одній машині,
# найпростіше і найшвидше - звертатися напряму на localhost:5000, в обхід
# nginx (не потрібно валідного SSL-ланцюжка для внутрішніх викликів).
ONLYOFFICE_CALLBACK_BASE_URL = os.environ.get('ONLYOFFICE_CALLBACK_BASE_URL', 'http://localhost:5000')

# ONLYOFFICE_JWT_SECRET - спільний секрет для підпису запitів між нашим
# застосунком і Document Server. Має ЗБІГАТИСЯ з секретом, який ви
# вкажете в налаштуваннях самого Document Server (файл local.json,
# розділ services.CoAuthoring.secret, або через змінну середовища
# JWT_SECRET при встановленні). Якщо залишити порожнім - перевірка JWT
# вимикається (підходить лише для першого тестування на localhost;
# для реальної роботи обов'язково задайте секрет і тут, і в Document
# Server - інакше будь-хто, хто вгадає doc_id, зможе підмінити callback
# збереження).
ONLYOFFICE_JWT_SECRET = os.environ.get('ONLYOFFICE_JWT_SECRET', '')

if not ONLYOFFICE_JWT_SECRET:
    logger.warning(
        "ONLYOFFICE_JWT_SECRET не задано - перевірка підпису запитів від "
        "Document Server вимкнена. Це прийнятно лише для первинного "
        "тестування на localhost. Перед реальним використанням задайте "
        "однаковий секрет тут і в налаштуваннях Document Server."
    )

# ONLYOFFICE_VERIFY_SSL - чи перевіряти дійсність SSL-сертифіката, коли
# наш Flask сам звертається до Document Server (наприклад, щоб забрати
# щойно збережену версію документа після редагування). Поки Document
# Server використовує самопідписаний сертифікат (типова ситуація для
# внутрішнього/тестового розгортання), стандартна перевірка сертифіката
# відхиляє з'єднання - так само, як спочатку відхиляв і браузер, доки
# ви вручну не підтвердили виняток. Коли на сервері з'явиться сертифікат
# від довіреного центру сертифікації (CA) - задайте ONLYOFFICE_VERIFY_SSL=true.
ONLYOFFICE_VERIFY_SSL = os.environ.get('ONLYOFFICE_VERIFY_SSL', 'false').strip().lower() in ('1', 'true', 'yes')

if not ONLYOFFICE_VERIFY_SSL:
    logger.warning(
        "ONLYOFFICE_VERIFY_SSL вимкнено - перевірка SSL-сертифіката Document Server "
        "при завантаженні збережених документів пропускається (типово для самопідписаного "
        "сертифіката). Коли встановите довірений сертифікат, задайте ONLYOFFICE_VERIFY_SSL=true."
    )

# ONLYOFFICE_COMMAND_URL - адреса службового "Command Service" Document
# Server, яку наш Flask використовує, щоб самому ІНІЦІЮВАТИ примусове
# збереження документа (той самий ефект, що й натискання користувачем
# Ctrl+S), перш ніж видати фінальний файл на кнопці "Завершити і
# завантажити". Document Server і Flask - на одній машині, тому
# найпростіше і найшвидше звертатися напряму на порт, куди встановлено
# Document Server (за замовчуванням у нашій інструкції - 8081), в обхід
# nginx і самопідписаного сертифіката.
ONLYOFFICE_COMMAND_URL = os.environ.get(
    'ONLYOFFICE_COMMAND_URL', 'http://localhost:8081/coauthoring/CommandService.ashx'
)