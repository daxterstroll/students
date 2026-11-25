import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="docxcompose")

from datetime import datetime, timedelta
import argparse
import json
import os
import threading
import time

from flask import Flask, render_template, request, abort, session, redirect, url_for
from werkzeug.middleware.proxy_fix import ProxyFix

from routes.config import SECRET_KEY, PREFERRED_URL_SCHEME
from routes.security import (
    authorize_and_serve_static,
    generate_csrf_token,
    validate_csrf_request,
)
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
#  before_request №1: блокування /static/templates/ та захист /static/uploads/
# ═══════════════════════════════════════════════════════════════════
@app.before_request
def block_templates_static_access():
    """Блокує прямий доступ до шаблонів і тимчасових каталогів та
    контролює доступ до /static/uploads/ на основі сесії, прав і скоупа."""
    return authorize_and_serve_static()


@app.before_request
def enforce_csrf():
    """Глобальна перевірка CSRF-токена для мутуючих HTTP-методів."""
    return validate_csrf_request()


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

    from routes.user_scope import get_allowed_group_ids

    try:
        conn = get_db()
        try:
            row = conn.execute(
                "SELECT role, is_admin, permissions FROM users WHERE id = ?",
                (session['user_id'],)
            ).fetchone()
            allowed_gids = get_allowed_group_ids(conn, session['user_id']) if row is not None else []
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
        or session.get('group_ids') != allowed_gids
    )
    if changed:
        session['is_admin'] = is_admin
        session['role'] = row['role']
        session['permissions'] = perms
        session['group_ids'] = allowed_gids
        session.modified = True

    session['perm_check_at'] = now
    session.modified = True


_FACULTY_SCHEMA_LOCK = threading.Lock()
_FACULTY_SCHEMA_READY = False


@app.before_request
def ensure_faculties_migrated():
    """Одноразово за процес ідемпотентно створює таблицю факультетів та колонки."""
    global _FACULTY_SCHEMA_READY
    if _FACULTY_SCHEMA_READY:
        return
    with _FACULTY_SCHEMA_LOCK:
        if _FACULTY_SCHEMA_READY:
            return
        from routes.db import get_db
        from routes.utils import ensure_faculty_schema
        conn = None
        try:
            conn = get_db()
            ensure_faculty_schema(conn)
            _FACULTY_SCHEMA_READY = True
        except Exception as e:
            logger.warning(f"Не вдалося виконати міграцію факультетів: {e}")
        finally:
            if conn is not None:
                conn.close()


_CATALOG_SYNC_LOCK = threading.Lock()
_CATALOG_SYNCED_YEAR = None


@app.before_request
def ensure_annual_catalogs():
    """
    Раз на календарний рік (або після перезапуску сервера) гарантує, що
    довідники `educational_programs` та `qualification_names` автоматично
    поповнені записами на поточний рік за даними попередніх років — навіть
    якщо адміністратор одразу відкриває «Групи», не заходячи у довідники.
    """
    global _CATALOG_SYNCED_YEAR
    if not session.get('user_id'):
        return

    current_year = datetime.now().year
    if _CATALOG_SYNCED_YEAR == current_year:
        return

    with _CATALOG_SYNC_LOCK:
        if _CATALOG_SYNCED_YEAR == current_year:
            return
        from routes.db import get_db
        from routes.admin.catalogs import auto_fill_catalog_for_year
        conn = None
        try:
            conn = get_db()
            c_prog = auto_fill_catalog_for_year(
                conn, 'educational_programs', target_year=current_year
            )
            c_qual = auto_fill_catalog_for_year(
                conn, 'qualification_names', target_year=current_year
            )
            if c_prog or c_qual:
                logger.info(
                    f"Автосинхронізація довідників на {current_year} рік: "
                    f"освітні програми={c_prog}, назви кваліфікацій={c_qual}"
                )
            _CATALOG_SYNCED_YEAR = current_year
        except Exception as e:
            logger.warning(f"Не вдалося автосинхронізувати довідники за {current_year} рік: {e}")
        finally:
            if conn is not None:
                conn.close()


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
            'eu_groups_catalog': {'degree_order': {}, 'groups': {}},
        }

    eu_groups_catalog = {'degree_order': {}, 'groups': {}}
    try:
        from routes.db import get_db
        conn = get_db()
        try:
            new_applications = conn.execute(
                "SELECT COUNT(*) FROM pending_students WHERE status='new'"
            ).fetchone()[0]
            new_update_requests = conn.execute(
                "SELECT COUNT(*) FROM update_requests WHERE status='submitted'"
            ).fetchone()[0]
            try:
                degree_rows = conn.execute(
                    "SELECT name_ua, id FROM degree_levels"
                ).fetchall()
                eu_groups_catalog['degree_order'] = {
                    r['name_ua']: r['id'] for r in degree_rows if r['name_ua']
                }
                group_rows = conn.execute(
                    """
                    SELECT id, name, start_year, study_form, degree_level,
                           course, program_credits, specialty
                    FROM groups
                    ORDER BY start_year DESC, name COLLATE UKRAINIAN
                    """
                ).fetchall()
                eu_groups_catalog['groups'] = {
                    str(g['id']): {
                        'name': g['name'] or '',
                        'year': g['start_year'],
                        'form': g['study_form'] or '',
                        'degree': g['degree_level'] or 'Без ступеня',
                        'course': g['course'],
                        'credits': g['program_credits'],
                        'specialty': g['specialty'] or '',
                    }
                    for g in group_rows
                }
            except Exception:
                pass
        finally:
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
        'eu_groups_catalog': eu_groups_catalog,
    }


app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

app.secret_key = SECRET_KEY

app.config.update(
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
    PREFERRED_URL_SCHEME=PREFERRED_URL_SCHEME,
    MAX_CONTENT_LENGTH=90 * 1024 * 1024,
)
app.jinja_env.globals['csrf_token'] = generate_csrf_token


@app.after_request
def add_security_headers(response):
    """Базові HTTP-заголовки безпеки для всіх відповідей застосунку."""
    response.headers.setdefault('X-Content-Type-Options', 'nosniff')
    response.headers.setdefault('X-Frame-Options', 'SAMEORIGIN')
    response.headers.setdefault('Referrer-Policy', 'strict-origin-when-cross-origin')
    response.headers.setdefault('Permissions-Policy', 'camera=(), microphone=(), geolocation=()')
    if request.is_secure:
        response.headers.setdefault(
            'Strict-Transport-Security',
            'max-age=31536000; includeSubDomains',
        )
    return response


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