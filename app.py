import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="docxcompose")

import argparse
import json
import os

from flask import Flask, render_template, request, abort, session
from werkzeug.middleware.proxy_fix import ProxyFix

from routes.config import SECRET_KEY, PREFERRED_URL_SCHEME
from routes.auth import auth_bp
from routes.students import students_bp
from routes.admin import admin_bp
from routes.office_editor import office_bp
from routes.import_grades import import_grades_bp
from routes.analytics import analytics_bp
from routes.photo_bulk import photo_bulk_bp
from routes.public_apply import public_apply_bp
from routes.public_update import public_update_bp
from routes.utils import logger
from routes.gen_docx import format_grade

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Гарантуємо, що всі відносні шляхи в застосунку працюють відносно
# кореня проєкту. На Windows-сервері, де застосунок запускається як
# фонова служба (nssm / Task Scheduler), CWD = C:\Windows\System32, і
# без цього рядка os.path.join('static', 'uploads', ...) у різних
# модулях (save_multiple_attachments, routes/photo.py) зберігає файли
# не туди, звідки Flask їх віддає.
os.chdir(BASE_DIR)

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'static', 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static')
)


@app.before_request
def block_templates_static_access():
    """
    Забороняє прямий доступ через /static/ до:

      1. Папки templates/ - без цього можна було б завантажити сирий
         невідрендерений Jinja-шаблон.

      2. Тимчасових папок військових PDF - там персональні дані
         (скан-копії «Резерв+»), які не повинні відкриватися за
         прямим URL /static/uploads/tmp_military_* /... поки їх не
         підтвердили і не перенесли в постійну папку.
    """
    if request.path.startswith('/static/templates/'):
        abort(404)
    if (request.path.startswith('/static/uploads/tmp_military_pdf/') or
            request.path.startswith('/static/uploads/tmp_military_bulk/')):
        abort(404)


@app.context_processor
def inject_nav_context():
    """
    Єдина точка, яка готує все для навбара:
      - nav_sections  - структура меню (routes/nav.py)
      - nav_badges    - числові бейджі (нові заявки, нові запити)
      - has_perm(p)   - True, якщо адмін або має дозвіл p
      - is_admin, perms - для backward-сумісності зі старими шаблонами

    Повертає порожні значення для незалогінених - навбар у цьому
    випадку взагалі не рендериться (див. layout.html).
    """
    if not session.get('user_id'):
        return {
            'nav_sections': [],
            'nav_badges': {},
            'has_perm': lambda p: False,
            'is_admin': False,
            'perms': [],
        }

    # Бейджі з кількістю нових заявок - та сама логіка, що й була в
    # inject_pending_counts (тепер об'єднана в один processor).
    try:
        from routes.db import get_db
        conn = get_db()
        new_applications = conn.execute(
            "SELECT COUNT(*) FROM pending_students WHERE status='new'"
        ).fetchone()[0]
        new_update_requests = conn.execute(
            "SELECT COUNT(*) FROM update_requests WHERE status='submitted'"
        ).fetchone()[0]
        conn.close()
    except Exception:
        new_applications = 0
        new_update_requests = 0

    perms = set(session.get('permissions', []))
    is_admin = bool(session.get('is_admin', False))

    def has_perm(permission):
        """True, якщо адмін АБО в користувача явно видано цей дозвіл.
        None/'' трактується як 'без перевірки' - True."""
        if not permission:
            return True
        return is_admin or permission in perms

    from routes.nav import NAV_SECTIONS

    return {
        'nav_sections': NAV_SECTIONS,
        'nav_badges': {
            'nav_pending_applications': new_applications,
            'nav_pending_updates': new_update_requests,
        },
        'has_perm': has_perm,
        'is_admin': is_admin,
        'perms': list(perms),
    }

app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

app.secret_key = SECRET_KEY

app.config.update(
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    PREFERRED_URL_SCHEME=PREFERRED_URL_SCHEME,
    MAX_CONTENT_LENGTH=32 * 1024 * 1024,
)


@app.template_filter('fromjson')
def fromjson(value):
    if value is None or value == '':
        return []
    try:
        return json.loads(value)
    except (json.JSONDecodeError, TypeError):
        return []


app.register_blueprint(auth_bp)
app.register_blueprint(admin_bp)
app.register_blueprint(students_bp)
app.register_blueprint(office_bp)
app.register_blueprint(import_grades_bp)
app.register_blueprint(analytics_bp)
app.register_blueprint(photo_bulk_bp)
app.register_blueprint(public_apply_bp)
app.register_blueprint(public_update_bp)
app.jinja_env.filters['format_grade'] = format_grade


@app.errorhandler(404)
def handle_404(e):
    try:
        return render_template('404.html'), 404
    except Exception:
        return "404 - Сторінку не знайдено", 404


@app.errorhandler(500)
def handle_500(e):
    logger.error(f"Необроблена помилка застосунку: {e}", exc_info=True)
    try:
        return render_template('500.html'), 500
    except Exception:
        return "500 - Внутрішня помилка сервера. Її вже записано в журнал.", 500


if __name__ == '__main__':
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
            "НІКОЛИ не використовуйте цей режим, якщо сервер доступний ззовні."
        )
        app.run(debug=True, host=args.host, port=args.port)
    else:
        from waitress import serve
        logger.info(f"Запуск застосунку у production-режимі на {args.host}:{args.port}")
        serve(app, host=args.host, port=args.port)