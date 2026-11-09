import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="docxcompose")

import argparse
import json
import os
import time

from flask import Flask, render_template, request, abort, session, redirect, url_for
from werkzeug.middleware.proxy_fix import ProxyFix

from routes.config import SECRET_KEY, PREFERRED_URL_SCHEME
from routes.auth import auth_bp
from routes.students import students_bp
from routes.admin import admin_bp
from routes.office_editor import office_bp
from routes.import_grades import import_grades_bp
from routes.analytics import analytics_bp
from routes.photo_bulk import photo_bulk_bp
from routes.military_reports import military_reports_bp
from routes.public_apply import public_apply_bp
from routes.public_update import public_update_bp
from routes.exam_sessions import exam_sessions_bp
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


# ═══════════════════════════════════════════════════════════════════
#  before_request №1: блокування прямого доступу через /static/
# ═══════════════════════════════════════════════════════════════════
@app.before_request
def block_templates_static_access():
    """Забороняє прямий доступ через /static/ до:
       - templates/ (сирі Jinja-шаблони);
       - тимчасових папок військових PDF (скан-копії «Резерв+»);
       - сесійних папок, у яких лежать згенеровані .docx,
         Excel-стани імпорту, JSON-звіти та тимчасові фото —
         усе це з персональними даними студентів і не повинно
         відкриватися за прямим URL, навіть якщо файл існує.
    """
    blocked_prefixes = (
        '/static/templates/',
        '/static/uploads/tmp_military_pdf/',
        '/static/uploads/tmp_military_bulk/',
        '/static/uploads/import_reports/',
        '/static/uploads/import_sessions/',
        '/static/uploads/office_sessions/',
        '/static/uploads/photo_bulk_sessions/',
        '/static/uploads/license_reports/',
    )
    for prefix in blocked_prefixes:
        if request.path.startswith(prefix):
            abort(404)


# ═══════════════════════════════════════════════════════════════════
#  before_request №2: TTL-оновлення прав у сесії
# ═══════════════════════════════════════════════════════════════════
@app.before_request
def refresh_session_permissions():
    """
    Перечитує права користувача з БД кожні PERMS_TTL секунд.

    Навіщо: адмін редагує права через /admin/users/<id>/edit, вони
    змінюються в БД, але в сесії користувача залишається стара копія
    (кеш ставиться при логіні). Через це забране право не діє, поки
    користувач не вийде і не ввійде заново, а навбар показує старі
    пункти. TTL робить зміни «живими» без перелогіну.

    Чому TTL, а не на кожен запит: SELECT на кожен запит — це зайве
    навантаження. 5 секунд — компроміс: адмін змінює права, а через
    кілька секунд вони вже діють у користувача.

    ВАЖЛИВО: після зміни прав користувач повинен ПЕРЕЗАВАНТАЖИТИ
    сторінку в браузері — інакше HTML у нього вже відрендерений зі
    старим навбаром, і нові права він побачить лише при переході
    на іншу сторінку (або після F5).
    """
    from routes.db import get_db  # локальний імпорт — щоб не плодити
                                   # циклічну залежність app.py ↔ routes.db

    PERMS_TTL = 5

    if not session.get('user_id'):
        return  # незалогінений — нічого не робимо

    now = time.time()
    last_check = session.get('perm_check_at', 0)
    if now - last_check < PERMS_TTL:
        return  # ще рано — віримо кешу

    try:
        conn = get_db()
        try:
            row = conn.execute(
                "SELECT role, is_admin, permissions FROM users WHERE id = ?",
                (session['user_id'],)
            ).fetchone()
        finally:
            conn.close()
    except Exception as e:
        logger.error(
            f"before_request refresh_session_permissions: не вдалося "
            f"перевірити права (user_id={session.get('user_id')}): {e}",
            exc_info=True,
        )
        return  # не валимо запит через тимчасову помилку БД

    if row is None:
        # Користувача видалили, поки він був залогінений — розлогінюємо.
        logger.info(
            f"before_request: користувач id={session.get('user_id')} "
            f"зник з БД — примусовий logout"
        )
        session.clear()
        return redirect(url_for('auth.login'))

    try:
        perms = json.loads(row['permissions'] or '[]')
        if not isinstance(perms, list):
            perms = []
    except (ValueError, TypeError):
        perms = []

    is_admin = bool(row['is_admin']) or (row['role'] == 'admin')

    # Оновлюємо лише якщо дійсно щось змінилось — щоб не переписувати
    # cookie на кожен запит.
    changed = (
        session.get('is_admin') != is_admin
        or session.get('role') != row['role']
        or session.get('permissions') != perms
    )
    if changed:
        session['is_admin'] = is_admin
        session['role'] = row['role']
        session['permissions'] = perms
        session.modified = True

    session['perm_check_at'] = now
    session.modified = True


# ═══════════════════════════════════════════════════════════════════
#  context_processor для навбара
# ═══════════════════════════════════════════════════════════════════
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
    MAX_CONTENT_LENGTH=90 * 1024 * 1024,
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
app.register_blueprint(military_reports_bp)
app.register_blueprint(public_apply_bp)
app.register_blueprint(public_update_bp)
app.register_blueprint(exam_sessions_bp)
app.jinja_env.filters['format_grade'] = format_grade


@app.template_filter('format_phone')
def format_phone_filter(value):
    if not value:
        return '—'
    from routes.utils import normalize_phone_ua
    canonical = normalize_phone_ua(value) or ''
    if len(canonical) == 13 and canonical.startswith('+380'):
        return f"+380 ({canonical[4:6]}) {canonical[6:9]}-{canonical[9:11]}-{canonical[11:13]}"
    return str(value)


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