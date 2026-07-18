import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="docxcompose")

import argparse
import json
import os

from flask import Flask, render_template, request, abort
from werkzeug.middleware.proxy_fix import ProxyFix

from routes.config import SECRET_KEY, PREFERRED_URL_SCHEME
from routes.auth import auth_bp
from routes.students import students_bp
from routes.admin import admin_bp
from routes.office_editor import office_bp
from routes.utils import logger
from routes.gen_docx import format_grade

# Инициализация приложения Flask
#
# ПРИМІТКА: templates фізично лежать УСЕРЕДИНІ static/ (так історично
# склалося в цьому проєкті, і ми свідомо залишаємо цю структуру як є).
# Проблема в тому, що Flask за замовчуванням роздає ВЕСЬ вміст
# static_folder на URL /static/<шлях> - тобто без додаткового захисту
# будь-хто міг би відкрити https://ваш-сайт/static/templates/login.html
# і побачити сирий, невідрендерений Jinja-код сторінки. Замість
# перенесення файлів (що вимагало б змінювати шлях на сервері) - нижче
# доданий окремий блок (`block_templates_static_access`), який просто
# забороняє прямий HTTP-доступ до підпапки templates всередині /static/,
# лишаючи решту static-файлів (CSS/JS/зображення/шаблони word) доступними
# як і раніше.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'static', 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static')
)


@app.before_request
def block_templates_static_access():
    """
    Забороняє пряме звернення до /static/templates/... ззовні.

    Без цього будь-хто міг би завантажити сирий, невідрендерений
    .html-шаблон (з коментарями, назвами внутрішніх макросів тощо)
    напряму через статичну роздачу Flask, в обхід звичайного
    рендерингу через render_template(). Усі інші файли під /static/
    (CSS, JS, зображення, .docx-шаблони) продовжують роздаватися як
    і раніше - обмеження стосується лише підпапки templates.
    """
    if request.path.startswith('/static/templates/'):
        abort(404)

# ProxyFix: застосунок працює за nginx (SSL termination + reverse proxy,
# див. nginx.conf). Без цього middleware Flask/Werkzeug не довіряють
# заголовкам X-Forwarded-For / X-Forwarded-Proto, які виставляє nginx, і
# request.remote_addr для КОЖНОГО запиту показував би адресу самого
# nginx (127.0.0.1), а не реальну адресу користувача. Це прямо ламало
# аудит-логування спроб входу в auth.py (лог "IP: 127.0.0.1" для всіх
# користувачів однаково - неможливо розслідувати підозрілі спроби входу).
# x_for=1, x_proto=1 - довіряємо рівно одному "хопу" проксі (nginx на
# тому ж сервері); якщо колись додасте ще один проксі/балансувальник
# попереду nginx, це число потрібно буде збільшити.
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

app.secret_key = SECRET_KEY

# Кукі сесії: сайт віддається лише по HTTPS (nginx), тож кука сесії
# ніколи не повинна піти по звичайному HTTP та не повинна бути доступна
# з JavaScript.
app.config.update(
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    PREFERRED_URL_SCHEME=PREFERRED_URL_SCHEME,
    # Обмеження розміру запиту (імпорт Excel/Word-шаблонів). Значення
    # узгоджене з client_max_body_size у nginx.conf - якщо збільшите
    # одне, збільшуйте і друге.
    MAX_CONTENT_LENGTH=32 * 1024 * 1024,  # 32 МБ
)


@app.template_filter('fromjson')
def fromjson(value):
    """Парсить JSON-рядок у список або словник"""
    if value is None or value == '':
        return []
    try:
        return json.loads(value)
    except (json.JSONDecodeError, TypeError):
        return []  # або повертайте {}, якщо очікуєте словник


# Регистрация blueprint'ов
app.register_blueprint(auth_bp)
app.register_blueprint(admin_bp)
app.register_blueprint(students_bp)
app.register_blueprint(office_bp)
app.jinja_env.filters['format_grade'] = format_grade


@app.errorhandler(404)
def handle_404(e):
    """Показує дружню 404-сторінку. Якщо шаблон 404.html відсутній - віддає простий текст, а не падає повторно."""
    try:
        return render_template('404.html'), 404
    except Exception:
        return "404 - Сторінку не знайдено", 404


@app.errorhandler(500)
def handle_500(e):
    """
    Раніше необроблені виключення на рівні застосунку (поза try/except
    у самих маршрутах) нічим не логувались у наш файловий лог
    (app.log) - Flask/waitress просто повертали користувачу 500-у
    сторінку, а деталі помилки залишались лише в консолі процесу
    (яку на Windows-сервері, запущеному як фонова служба, ніхто не
    бачить). Тепер будь-яка непіймана помилка гарантовано потрапляє в
    app.log з повним traceback.
    """
    logger.error(f"Необроблена помилка застосунку: {e}", exc_info=True)
    try:
        return render_template('500.html'), 500
    except Exception:
        return "500 - Внутрішня помилка сервера. Її вже записано в журнал.", 500


# --- Запуск приложения ---
if __name__ == '__main__':
    """
    Запускает Flask-приложение в зависимости от выбранного режима.

    Аргументы командной строки:
        --mode (str): Режим запуска ('debug' для локального сервера, 'production' для waitress).
                      За замовчуванням - 'production' (див. примітку з міркувань безпеки нижче).
        --host (str): Хост (за замовчуванням 'localhost' - саме так налаштований nginx.conf,
                      що проксує на localhost:5000).
        --port (int): Порт (за замовчуванням 5000 - саме так налаштований nginx.conf).

    Приклади:
        python app.py                                   # production, localhost:5000 (звичайний запуск за nginx)
        python app.py --mode debug                       # локальна розробка з debug-режимом Flask
        python app.py --mode production --host 0.0.0.0 --port 8080

    ВАЖЛИВО ПРО БЕЗПЕКУ: раніше --mode за замовчуванням дорівнював
    'debug'. Якщо застосунок коли-небудь запускали (наприклад, вручну
    перезапускали на сервері, або через Планувальник завдань Windows)
    БЕЗ явного прапорця --mode production, він тихо стартував у
    Flask debug-режимі. У debug-режимі Flask вмикає інтерактивну
    веб-консоль Werkzeug: при будь-якій необробленій помилці кожен, хто
    відкриє сторінку помилки, отримує консоль з можливістю виконати
    ДОВІЛЬНИЙ python-код прямо на сервері (це офіційно задокументована
    поведінка Werkzeug, не баг). Оскільки застосунок дивиться назовні
    через nginx на 443 порт, це означає віддалене виконання коду для
    будь-кого, хто зловить помилку на сайті. Тепер за замовчуванням -
    production (waitress, без debug-консолі); debug потрібно вмикати
    явно і свідомо лише на своїй машині розробника.
    """
    parser = argparse.ArgumentParser(description='Запуск Flask-приложения в указанном режиме.')
    parser.add_argument('--mode', choices=['debug', 'production'], default='production',
                        help="Режим запуска: 'production' (waitress, за замовчуванням) або 'debug' (лише для локальної розробки!)")
    parser.add_argument('--host', default='localhost',
                        help='Хост (за замовчуванням: localhost - узгоджено з nginx.conf)')
    parser.add_argument('--port', type=int, default=5000,
                        help='Порт (за замовчуванням: 5000 - узгоджено з nginx.conf)')
    args = parser.parse_args()

    if args.mode == 'debug':
        logger.warning(
            f"Застосунок запущено в DEBUG-режимі (host={args.host}, port={args.port}). "
            "НІКОЛИ не використовуйте цей режим, якщо сервер доступний ззовні через nginx/інтернет - "
            "debug-консоль Werkzeug дозволяє виконання довільного коду."
        )
        app.run(debug=True, host=args.host, port=args.port)
    else:  # args.mode == 'production'
        from waitress import serve
        logger.info(f"Запуск застосунку у production-режимі на {args.host}:{args.port}")
        serve(app, host=args.host, port=args.port)