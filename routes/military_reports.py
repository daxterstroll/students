"""
routes/military_reports.py
==========================
Військовий облік студентів — Дашборд та списки за Додатком 5
(Постанова КМУ №916 від 30.07.2025).

Розділення студентів на 4 групи (згідно з Додатком 5):
  1. Офіцери      — чоловіки з офіцерським званням
  2. Рядові       — чоловіки з рядовим/сержантським/старшинським званням
  3. Жінки        — всі жінки незалежно від звання (окремий список у постанові)
  4. Призовники   — status_badge = 'Призовник' незалежно від іншого

Пріоритет визначення:
  1. Призовник (за status_badge)
  2. Жінка (за gender)
  3. Явне перевизначення (military.category_override)
  4. Автоматично за званням (fuzzy match з довідником)

Виключені з усіх 4 груп (не входять в Додаток 5):
  - Невійськовозобов'язаний
  - Знято з обліку

Категорія не визначена (None) → потрапляє в «Проблемні»
"""

import sqlite3
from datetime import date

from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify
from routes.db import get_db
from routes.utils import log_action, permission_required, logger
from routes.helpers import current_username

military_reports_bp = Blueprint('military_reports', __name__)


# ═══════════════════════════════════════════════════════════════════════
# Довідники звань (для fuzzy-розпізнавання категорії)
# ═══════════════════════════════════════════════════════════════════════

OFFICER_RANKS = {
    'молодший лейтенант', 'лейтенант', 'старший лейтенант',
    'капітан', 'капітан 2 рангу', 'капітан 1 рангу',
    'майор', 'підполковник', 'полковник',
    'генерал-майор', 'генерал-лейтенант', 'генерал',
    # медична служба
    'молодший лейтенант медичної служби',
    'лейтенант медичної служби',
    'старший лейтенант медичної служби',
    'капітан медичної служби',
    'майор медичної служби',
    'підполковник медичної служби',
    'полковник медичної служби',
    # юстиція
    'молодший лейтенант юстиції',
    'лейтенант юстиції',
    'старший лейтенант юстиції',
    'капітан юстиції',
    'майор юстиції',
    'підполковник юстиції',
    'полковник юстиції',
    # ветеринарна служба
    'молодший лейтенант ветеринарної служби',
    'лейтенант ветеринарної служби',
    'старший лейтенант ветеринарної служби',
    'капітан ветеринарної служби',
    'майор ветеринарної служби',
    'підполковник ветеринарної служби',
    'полковник ветеринарної служби',
}

RANK_AND_FILE = {
    'солдат', 'рядовий', 'матрос',
    'старший солдат', 'старший матрос',
    'молодший сержант', 'сержант', 'старший сержант',
    'старшина', 'головний старшина', 'головний корабельний старшина',
    'прапорщик', 'старший прапорщик',
    'рекрут',
}

EXCLUDED_STATUSES = {
    "Невійськовозобов'язаний",
    'Невійськовозобов’язаний',
    'Знято з обліку',
}

# ── Позначки графи 18 «Примітка» (Постанова КМУ №812 від 10.06.2026) ──
# За новою редакцією ці значення є обов'язковими для фіксації факту
# вибуття військовозобов'язаного. Записи з цими позначками:
#   - НЕ входять у Додаток 5 (жодна з 4 груп)
#   - ВХОДЯТЬ у Додаток 12 окремими рядками (3/4/5)
#   - Зберігаються у звітах до кінця поточного календарного року
NOTE_EXCLUDED_BY_AGE = 'Виключено з військового обліку за віком'
NOTE_DISMISSED       = 'Звільнено з роботи, завершено навчання (відраховано)'
NOTE_MOBILIZED       = 'Військовослужбовець'

NOTE_18_LABELS = {
    NOTE_EXCLUDED_BY_AGE: 'Виключено з військового обліку за віком',
    NOTE_DISMISSED:       'Звільнено з роботи, завершено навчання (відраховано)',
    NOTE_MOBILIZED:       'Військовослужбовець',
}

# Позначки, які виключають студента з Додатка 5 (усіх 4 груп)
NOTE_18_EXCLUDES_FROM_DODATOK_5 = {NOTE_EXCLUDED_BY_AGE, NOTE_DISMISSED, NOTE_MOBILIZED}

def _has_military_data(status_badge, military_rank, vos_code, tck_name):
    """
    Чи є у студента хоч якісь заповнені військові дані?
    Використовується для визначення "жінка без військового обліку"
    (такі не показуються ніде — ні в групах, ні в проблемних, ні в звітах).
    """
    for v in (status_badge, military_rank, vos_code, tck_name):
        if v and str(v).strip():
            return True
    return False


def _is_woman_without_military(gender, military_id, status_badge,
                                military_rank, vos_code, tck_name):
    """
    Жінка без військових даних — не показується ніде:
      - немає запису military, АБО
      - запис є, але всі ключові поля порожні (status_badge, rank, vos, tck).
    """
    if gender != 'F':
        return False
    if not military_id:
        return True
    return not _has_military_data(
        status_badge, military_rank, vos_code, tck_name
    )

def _is_dismissed_this_year(dismissed_at, report_year):
    """
    Чи був студент відрахований/випущений у ВКАЗАНОМУ звітному році.

    Примітка до Списків (№812): відомості про звільнених, відрахованих
    або виключених з обліку осіб зберігаються в списках ДО КІНЦЯ
    поточного календарного року. Тобто в звіті за 2026 рік вони ще
    показуються, а в звіті за 2027 — вже ні.

    Повертає True, якщо dismissed_at має рік == report_year.
    """
    if not dismissed_at:
        return False
    try:
        year = int(str(dismissed_at).split('-')[0])
        return year == report_year
    except (ValueError, IndexError):
        return False

# ═══════════════════════════════════════════════════════════════════════
# Допоміжні функції
# ═══════════════════════════════════════════════════════════════════════

def _normalize_rank(rank):
    """
    Нормалізація звання для порівняння:
      - нижній регістр
      - прямий апостроф
      - видалення ' запасу', ' у відставці'
      - видалення '-' (порожнє)
    """
    if not rank:
        return ''
    r = str(rank).strip().lower()
    r = r.replace('’', "'").replace('ʼ', "'").replace('‘', "'")
    r = r.replace(' запасу', '').replace(' у відставці', '')
    r = ' '.join(r.split())  # нормалізуємо пробіли
    if r in ('-', '—', '–', ''):
        return ''
    return r


def _detect_category(rank, status_badge, gender, override=None, note_18=None):
    """
    Визначає категорію для Додатку 5.

    Повертає:
        'officers' | 'rank' | 'women' | 'conscripts' | None

    None = категорія не визначена АБО студент вибув (позначка графи 18).

    Параметр note_18 (Постанова №812):
      - 'Виключено за віком'          → None (не показуємо в Додатку 5)
      - 'Звільнено (відраховано)'     → None (не показуємо в Додатку 5,
                                        але є в Додатку 12 окремим рядком)
      - 'Військовослужбовець'         → None (не показуємо в Додатку 5,
                                        є в Додатку 12 окремим рядком)
      - None                          → звичайна логіка визначення категорії
    """
    # 0. Виключені з обліку (Постанова №812, графа 18)
    if note_18 in NOTE_18_EXCLUDES_FROM_DODATOK_5:
        return None

    # 1. Призовник — найвищий пріоритет
    if status_badge == 'Призовник':
        return 'conscripts'

    # 2. Жінка — друга за пріоритетом (окремий список у постанові)
    if gender == 'F':
        return 'women'

    # 3. Явне перевизначення адміном
    if override == 'officer':
        return 'officers'
    if override == 'rank':
        return 'rank'

    # 4. Автоматично за званням
    r = _normalize_rank(rank)
    if not r:
        return None

    # Точне співпадіння
    if r in OFFICER_RANKS:
        return 'officers'
    if r in RANK_AND_FILE:
        return 'rank'

    # Fuzzy match (для опечаток типу "молодший лейтинант")
    try:
        from rapidfuzz import fuzz, process
        best_officer = process.extractOne(r, OFFICER_RANKS, scorer=fuzz.ratio)
        best_rank = process.extractOne(r, RANK_AND_FILE, scorer=fuzz.ratio)
        score_o = best_officer[1] if best_officer else 0
        score_r = best_rank[1] if best_rank else 0
        if score_o >= 85 and score_o > score_r:
            return 'officers'
        if score_r >= 85:
            return 'rank'
    except ImportError:
        pass

    # Не розпізнано — адмін вирішить через override
    return None


def _get_report_dates():
    """
    Повертає дати для звітів. Автоматично масштабується на будь-який рік.

    annual                   — 01.01 поточного року (щорічний звіт, п. 40)
    current                  — сьогодні (проміжний список для ТЦК)
    next_deadline            — 25.01 наступного року (дедлайн щорічної здачі)
    next_report_date         — 01.01 наступного року (підготовка наступного звіту)
    reconciliation_deadline  — 01.12 поточного року (звіряння з ТЦК, №812)
    """
    today = date.today()
    year = today.year
    return {
        'annual': date(year, 1, 1),
        'current': today,
        'next_deadline': date(year + 1, 1, 25),
        'next_report_date': date(year + 1, 1, 1),
        'reconciliation_deadline': date(year, 12, 1),
    }

def _get_base_license_id(conn):
    """
    Повертає id базової ліцензії (institution_licenses.is_base = 1).
    За Постановою №812 військовий облік ведеться для конкретного закладу
    (Львівська філія ПВНЗ «Європейський університет»). Тому за замовчуванням
    на дашборді, у списках груп і в Додатку 12 показуються лише студенти
    цієї ліцензії (+ студенти без ліцензії, щоб їх не загубити).
    """
    row = conn.execute(
        "SELECT id FROM institution_licenses WHERE is_base = 1 LIMIT 1"
    ).fetchone()
    if not row:
        return None
    return row[0]


def _get_base_license_info(conn):
    """Повертає (id, short_name_ua, name_ua) базової ліцензії або (None, '', '')."""
    row = conn.execute(
        "SELECT id, short_name_ua, name_ua FROM institution_licenses "
        "WHERE is_base = 1 LIMIT 1"
    ).fetchone()
    if not row:
        return None, '', ''
    return row[0], (row[1] or ''), (row[2] or '')



def _format_date_ua(d):
    """Форматує date → 'ДД.ММ.РРРР'."""
    return d.strftime('%d.%m.%Y') if d else ''


def _count_category_rows(conn, category):
    """
    Повертає список студентів певної категорії.
    Використовується на сторінках списків (Крок 5+).
    """
    rows = conn.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.gender, s.tax_id,
               m.id AS military_id, m.status_badge, m.military_rank,
               m.category_override, m.tck_name, m.vos_code, m.registry_number
        FROM students s
        INNER JOIN military m ON m.student_id = s.id
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA COLLATE UKRAINIAN, s.first_name_UA COLLATE UKRAINIAN
    """).fetchall()

    result = []
    for r in rows:
        status = r['status_badge'] or ''
        if status in EXCLUDED_STATUSES:
            continue
        cat = _detect_category(
            r['military_rank'], status, r['gender'], r['category_override']
        )
        if cat == category:
            result.append(dict(r))
    return result


# ═══════════════════════════════════════════════════════════════════════
# Маршрути
# ═══════════════════════════════════════════════════════════════════════

@military_reports_bp.route('/admin/military_reports')
@permission_required('manage_military_reports')
def dashboard():
    """
    Головна сторінка модуля:
      - 4 лічильники груп (офіцери / рядові / жінки / призовники)
      - Бейдж дедлайну щорічної здачі
      - Стан підготовки щорічного звіту (які документи вже згенеровано)
      - Проблемні записи (без ТЦК / без ВОС / без категорії / без статі / без military)
      - Посилання на Додаток 12 і довідник TCK

    Лічильники рахують ЛИШЕ діючих студентів:
      - не в архіві (archived = 0)
      - не заморожених (немає активного запису в frozen_students)
    Заморожені рахуються окремо (для довідки) і не входять у звіт.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    # Фільтр по базовій ліцензії (Постанова №812)
    show_all_licenses = request.args.get('show_all_licenses') == '1'
    base_license_id, base_license_short, base_license_name = _get_base_license_info(conn)

    students = conn.execute("""
        SELECT s.id, s.gender, s.dismissed_at, s.license_id,
               m.id AS military_id, m.status_badge, m.military_rank,
               m.category_override, m.tck_name, m.vos_code, m.note_18,
               fz.id AS frozen_id
        FROM students s
        LEFT JOIN military m ON m.student_id = s.id
        LEFT JOIN frozen_students fz
               ON fz.student_id = s.id AND fz.resolved_at IS NULL
        WHERE COALESCE(s.archived, 0) = 0
    """).fetchall()

    # Дати (потрібні нижче для report_year)
    dates = _get_report_dates()
    report_year = dates['annual'].year

    counts = {
        'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0,
        # Нові лічильники (Постанова №812, графа 18)
        'excluded_by_age': 0,
        'dismissed_this_year': 0,
        'mobilized': 0,
    }
    issues = {
        'no_tck': 0,
        'no_vos': 0,
        'no_category': 0,
        'no_gender': 0,
        'no_military': 0,
    }
    frozen_count = 0
    total_students = 0

    for s in students:
        # ── 0. Фільтр ліцензії (Постанова №812) ──
        # Студенти без ліцензії (NULL) показуються ЗАВЖДИ (з бейджем),
        # бо інакше вони б "зникли" з поля зору і їх ніхто б не заповнив.
        if not show_all_licenses:
            if s['license_id'] is not None and s['license_id'] != base_license_id:
                continue

        # ── 1. Заморожені — рахуємо окремо, далі не йдемо ──
        if s['frozen_id']:
            if s['military_id']:
                status = s['status_badge'] or ''
                if status not in EXCLUDED_STATUSES and not _is_woman_without_military(
                    s['gender'], s['military_id'], s['status_badge'],
                    s['military_rank'], s['vos_code'], s['tck_name']
                ):
                    frozen_count += 1
            continue

        # ── 2. Активний — рахуємо в загальному охваті ──
        total_students += 1

        # ── 3. Жінка без військових даних — не рахуємо ніде ──
        if _is_woman_without_military(
            s['gender'], s['military_id'], s['status_badge'],
            s['military_rank'], s['vos_code'], s['tck_name']
        ):
            continue

        # ── 4. Без запису military (це вже тільки чоловіки, жінок відсіяли) ──
        if not s['military_id']:
            issues['no_military'] += 1
            continue

        # ── 5. Виключені статуси (Невійськовозобов'язаний / Знято з обліку) ──
        status = s['status_badge'] or ''
        if status in EXCLUDED_STATUSES:
            continue

        # ── 6. Позначка графи 18 (Постанова №812) ──
        #     Виключені за віком / відраховані / мобілізовані
        #     не входять у 4 групи Додатка 5, але показуються
        #     на дашборді окремими лічильниками (для Додатка 12).
        if s['note_18'] == NOTE_EXCLUDED_BY_AGE:
            counts['excluded_by_age'] += 1
            continue
        if s['note_18'] == NOTE_DISMISSED:
            # За приміткою до Списків — показуємо в звіті, якщо
            # відраховано в ПОТОЧНОМУ році
            if _is_dismissed_this_year(s['dismissed_at'], report_year):
                counts['dismissed_this_year'] += 1
            continue
        if s['note_18'] == NOTE_MOBILIZED:
            counts['mobilized'] += 1
            continue

        # ── 7. Без статі ──
        if not s['gender']:
            issues['no_gender'] += 1

        # ── 8. Категорія ──
        category = _detect_category(
            s['military_rank'], status, s['gender'],
            s['category_override'], s['note_18']
        )
        if category:
            counts[category] += 1
        else:
            issues['no_category'] += 1

        # ── 9. Обов'язкові поля ──
        if not s['tck_name']:
            issues['no_tck'] += 1
        if not s['vos_code']:
            issues['no_vos'] += 1

    # Стан підготовки щорічного звіту (з журналу)
    report_types = [
        ('dodatok_5_officers', 'Додаток 5 — Офіцери'),
        ('dodatok_5_rank', 'Додаток 5 — Рядові'),
        ('dodatok_5_women', 'Додаток 5 — Жінки'),
        ('dodatok_5_conscripts', 'Додаток 5 — Призовники'),
        ('dodatok_12', 'Додаток 12 (зведення)'),
    ]

    last_generations = {}
    for rtype, _rlabel in report_types:
        row = conn.execute("""
            SELECT generated_at, generated_by, rows_count
            FROM military_reports_log
            WHERE report_type = ? AND report_date = ?
            ORDER BY generated_at DESC LIMIT 1
        """, (rtype, dates['annual'].isoformat())).fetchone()
        last_generations[rtype] = dict(row) if row else None

    conn.close()

    return render_template(
        'military_reports/dashboard.html',
        counts=counts,
        issues=issues,
        frozen_count=frozen_count,
        total_students=total_students,
        dates=dates,
        dates_formatted={
            'annual': _format_date_ua(dates['annual']),
            'current': _format_date_ua(dates['current']),
            'next_deadline': _format_date_ua(dates['next_deadline']),
            'reconciliation_deadline': _format_date_ua(dates['reconciliation_deadline']),
        },
        report_types=report_types,
        last_generations=last_generations,
        # Ліцензія (Постанова №812)
        show_all_licenses=show_all_licenses,
        base_license_id=base_license_id,
        base_license_label=base_license_short or base_license_name or '—',
    )
    

# ═══════════════════════════════════════════════════════════════════════
# Тимчасові заглушки для майбутніх сторінок (Крок 5+)
# ═══════════════════════════════════════════════════════════════════════

CATEGORY_LABELS = {
    'officers':   ('Офіцери', 'Офіцерський склад'),
    'rank':       ('Рядові, сержанти, старшини',
                   'Рядовий, сержантський та старшинський склад'),
    'women':      ('Жінки', 'Військовозобов\'язані з числа жінок'),
    'conscripts': ('Призовники', 'Призовники (до призову)'),
}


@military_reports_bp.route('/admin/military_reports/group/<category>')
@permission_required('manage_military_reports')
def group_list(category):
    """
    Список студентів однієї категорії з фільтрами, пагінацією
    та перемикачем дати складання (A/щорічний, B/поточний, C/довільний).

    За замовчуванням заморожені студенти НЕ показуються (вони тимчасово
    не навчаються і не входять у звіт). Чекбокс «Показати заморожених»
    дозволяє побачити їх довідково (з бейджем у рядку).
    """
    if category not in CATEGORY_LABELS:
        flash("Невідома категорія", "error")
        return redirect(url_for('military_reports.dashboard'))

    label, subtitle = CATEGORY_LABELS[category]

    # ── Параметри запиту ─────────────────────────────────────
    search = (request.args.get('search') or '').strip()
    group_id = request.args.get('group_id', type=int)
    tck_filter = (request.args.get('tck') or '').strip()
    include_frozen = request.args.get('include_frozen') == '1'
    include_dismissed = request.args.get('include_dismissed') == '1'
    show_all_licenses = request.args.get('show_all_licenses') == '1'
    report_mode = request.args.get('report_mode', 'annual')
    if report_mode not in ('annual', 'current', 'custom'):
        report_mode = 'annual'
    report_date_custom = (request.args.get('report_date') or '').strip()
    page = max(1, request.args.get('page', 1, type=int))
    per_page = request.args.get('per_page', 50, type=int)
    if per_page not in (20, 50, 100, 200, 500):
        per_page = 50

    # ── Визначення дати ──────────────────────────────────────
    dates = _get_report_dates()
    if report_mode == 'annual':
        report_date = dates['annual']
    elif report_mode == 'current':
        report_date = dates['current']
    else:
        # custom
        from datetime import datetime as _dt
        try:
            report_date = _dt.strptime(report_date_custom, '%d.%m.%Y').date()
        except (ValueError, TypeError):
            report_date = dates['current']

    conn = get_db()
    conn.row_factory = sqlite3.Row

    # ── Довідники для фільтрів ───────────────────────────────
    groups = conn.execute("""
        SELECT id, name, start_year, study_form
        FROM groups WHERE archived = FALSE
        ORDER BY name COLLATE UKRAINIAN
    """).fetchall()

    tck_offices = conn.execute("""
        SELECT id, name_ua, address FROM tck_offices
        WHERE is_active = 1 ORDER BY name_ua COLLATE UKRAINIAN
    """).fetchall()

    # ── Основні дані ─────────────────────────────────────────
    # Приєднуємо frozen_students, щоб знати, хто заморожений.
    # Архівні/відчислені (archived = 1) — завжди відсіюємо.
    # Базова ліцензія (Постанова №812)
    base_license_id, base_license_short, base_license_name = _get_base_license_info(conn)

    rows_raw = conn.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.gender, s.tax_id, s.dismissed_at,
               s.license_id,
               il.short_name_ua AS license_short_name_ua,
               il.name_ua AS license_name_ua,
               m.id AS military_id, m.status_badge, m.military_rank,
               m.category_override, m.tck_name, m.vos_code, m.vos_description,
               m.registry_number, m.military_registration_document,
               m.note_18, m.mobilized_date,
               g.id AS group_id, g.name AS group_name,
               fz.id AS frozen_id, fz.reason AS frozen_reason
        FROM students s
        INNER JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON s.group_id = g.id
        LEFT JOIN institution_licenses il ON il.id = s.license_id
        LEFT JOIN frozen_students fz
               ON fz.student_id = s.id AND fz.resolved_at IS NULL
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA COLLATE UKRAINIAN,
                 s.first_name_UA COLLATE UKRAINIAN
    """).fetchall()

    # ── Фільтрація в Python ──────────────────────────────────
    filtered = []
    for r in rows_raw:
        # ── Фільтр базової ліцензії (Постанова №812) ──
        # Без ліцензії (NULL) — показуємо завжди (з бейджем).
        if not show_all_licenses:
            if r['license_id'] is not None and r['license_id'] != base_license_id:
                continue

        # Заморожені — тільки якщо увімкнено чекбокс
        if r['frozen_id'] and not include_frozen:
            continue

        # Жінка без військових даних — не показуємо ніде
        if _is_woman_without_military(
            r['gender'], r['military_id'], r['status_badge'],
            r['military_rank'], r['vos_code'], r['tck_name']
        ):
            continue

        status = r['status_badge'] or ''
        if status in EXCLUDED_STATUSES:
            continue
        cat = _detect_category(
            r['military_rank'], status, r['gender'], r['category_override']
        )
        if cat != category:
            continue
        if search:
            haystack = (
                f"{r['last_name_UA']} {r['first_name_UA']} "
                f"{r['middle_name_UA'] or ''}"
            ).lower()
            if not all(w in haystack for w in search.lower().split()):
                continue
        if group_id and r['group_id'] != group_id:
            continue
        if tck_filter and (r['tck_name'] or '') != tck_filter:
            continue
        filtered.append(dict(r))

    total = len(filtered)
    total_pages = (total + per_page - 1) // per_page if total else 1
    if page > total_pages:
        page = total_pages
    start = (page - 1) * per_page
    page_rows = filtered[start:start + per_page]

    # ── Паспорти для рядків поточної сторінки ────────────────
    if page_rows:
        ids = [r['id'] for r in page_rows]
        placeholders = ','.join('?' * len(ids))
        passport_rows = conn.execute(f"""
            SELECT student_id, document_type, series, number,
                   issued_by, issue_date, valid_until, unique_number
            FROM passport_documents
            WHERE student_id IN ({placeholders})
            ORDER BY student_id, id DESC
        """, ids).fetchall()
        passports = {}
        for p in passport_rows:
            if p['student_id'] not in passports:
                passports[p['student_id']] = dict(p)
        for r in page_rows:
            r['passport'] = passports.get(r['id'])

    conn.close()

    # ── Список імен ТЦК для «Інше» в модалці ────────────────
    tck_names = {t['name_ua'] for t in tck_offices}
    for r in page_rows:
        r['tck_in_catalog'] = (r.get('tck_name') in tck_names) if r.get('tck_name') else True

    # ── Дані для JS (модалка) ────────────────────────────────
    rows_json = [
        {
            'id': r['id'],
            'full_name': f"{r['last_name_UA']} {r['first_name_UA']} "
                         f"{r['middle_name_UA'] or ''}".strip(),
            'military_rank': r['military_rank'] or '',
            'category_override': r['category_override'] or '',
            'vos_code': r['vos_code'] or '',
            'tck_name': r['tck_name'] or '',
            'tck_in_catalog': r['tck_in_catalog'],
            'registry_number': r['registry_number'] or '',
            'note_18': r['note_18'] or '',
            'mobilized_date': r['mobilized_date'] or '',
            'license_id': r['license_id'],
            'license_short_name_ua': r['license_short_name_ua'] or '',
        }
        for r in page_rows
    ]

    range_start = start + 1 if total else 0
    range_end = min(start + per_page, total)

    return render_template(
        'military_reports/group_list.html',
        category=category,
        label=label,
        subtitle=subtitle,
        rows=page_rows,
        rows_json=rows_json,
        total=total,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        range_start=range_start,
        range_end=range_end,
        search=search,
        group_id=group_id,
        tck_filter=tck_filter,
        include_frozen=include_frozen,
        include_dismissed=include_dismissed,
        show_all_licenses=show_all_licenses,
        base_license_id=base_license_id,
        base_license_label=base_license_short or base_license_name or '—',
        report_mode=report_mode,
        report_date_custom=report_date_custom,
        report_date_ua=_format_date_ua(report_date),
        dates=dates,
        dates_formatted={
            'annual': _format_date_ua(dates['annual']),
            'current': _format_date_ua(dates['current']),
        },
        groups=groups,
        tck_offices=tck_offices,
    )

@military_reports_bp.route(
    '/admin/military_reports/student/<int:student_id>/quick_edit',
    methods=['POST'],
)
@permission_required('manage_military_reports')
def quick_edit(student_id):
    """
    AJAX: швидке редагування 5 полів військових даних прямо зі списку.
    Приймає JSON: {military_rank, category_override, vos_code, tck_name, registry_number}
    """
    data = request.get_json(silent=True) or {}

    conn = get_db()
    conn.row_factory = sqlite3.Row

    student = conn.execute(
        "SELECT id, last_name_UA, first_name_UA FROM students WHERE id = ?",
        (student_id,),
    ).fetchone()
    if not student:
        conn.close()
        return jsonify({'ok': False, 'error': 'Студента не знайдено'}), 404

    military = conn.execute(
        "SELECT id FROM military WHERE student_id = ?", (student_id,)
    ).fetchone()
    if not military:
        conn.close()
        return jsonify(
            {'ok': False, 'error': 'У студента немає запису військового обліку'}
        ), 400

    updates = {}
    for field in ('military_rank', 'category_override', 'vos_code',
                  'tck_name', 'registry_number', 'note_18', 'mobilized_date'):
        if field in data:
            value = (data[field] or '').strip() or None
            if field == 'category_override' and value not in (
                None, 'officer', 'rank'
            ):
                value = None
            if field == 'note_18' and value not in (None, *NOTE_18_LABELS.keys()):
                value = None
            if field == 'mobilized_date' and value:
                # Валідація формату ДД.ММ.РРРР
                try:
                    from datetime import datetime as _dt
                    _dt.strptime(value, '%d.%m.%Y')
                except (ValueError, TypeError):
                    value = None
            updates[field] = value

    if not updates:
        conn.close()
        return jsonify({'ok': False, 'error': 'Немає що оновлювати'}), 400

    set_clause = ', '.join(f"{k} = ?" for k in updates)
    try:
        conn.execute(
            f"UPDATE military SET {set_clause} WHERE student_id = ?",
            list(updates.values()) + [student_id],
        )
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        logger.error(
            f"quick_edit military (student_id={student_id}): {e}",
            exc_info=True,
        )
        return jsonify({'ok': False, 'error': str(e)}), 500

    conn.close()
    log_action(
        current_username(),
        f"швидке редагування військових даних: "
        f"{student['last_name_UA']} {student['first_name_UA']} (ID {student_id})",
        details=f"поля: {', '.join(updates.keys())}",
    )
    return jsonify({'ok': True})


ISSUE_LABELS = {
    'no_military': 'Немає запису військового обліку',
    'no_gender':   'Не визначено стать',
    'no_category': 'Не визначено категорію (звання)',
    'no_tck':      'Не вказано ТЦК',
    'no_vos':      'Не вказано ВОС',
    'no_license':  'Не вказано ліцензію вступу',
}


@military_reports_bp.route('/admin/military_reports/issues')
@permission_required('manage_military_reports')
def issues_list():
    """
    Сторінка «Проблемні» — студенти з неповними військовими даними,
    які заважають коректній генерації Додатка 5.

    Типи проблем:
      - no_military  — студент взагалі без запису в military
      - no_gender    — стать не визначена (впливає на категорію)
      - no_category  — звання не розпізнане → категорія не визначена
      - no_tck       — порожнє поле tck_name
      - no_vos       — порожнє поле vos_code

    Фільтри: чекбокси по типах проблем (за замовчуванням усі увімкнені).
    Студент показується, якщо має ХОЧА Б ОДНУ з увімкнених проблем.

    Заморожені — за замовчуванням приховані (як на сторінках груп),
    чекбокс «Показати заморожених» вмикає їх.
    """
    # ── Параметри ────────────────────────────────────────────
    enabled = set()
    for key in ISSUE_LABELS:
        if request.args.get(f'issue_{key}', '1') == '1':
            enabled.add(key)
    # Якщо всі вимкнено — показуємо всі (щоб сторінка не була пустою)
    if not enabled:
        enabled = set(ISSUE_LABELS.keys())

    include_frozen = request.args.get('include_frozen') == '1'
    search = (request.args.get('search') or '').strip()
    group_id = request.args.get('group_id', type=int)
    page = max(1, request.args.get('page', 1, type=int))
    per_page = request.args.get('per_page', 50, type=int)
    if per_page not in (20, 50, 100, 200, 500):
        per_page = 50

    conn = get_db()
    conn.row_factory = sqlite3.Row

    groups = conn.execute("""
        SELECT id, name, start_year, study_form
        FROM groups WHERE archived = FALSE
        ORDER BY name COLLATE UKRAINIAN
    """).fetchall()

    # Всі активні студенти (архівні/відчислені — відсіюємо)
    rows_raw = conn.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.gender, s.license_id,
               m.id AS military_id, m.status_badge, m.military_rank,
               m.category_override, m.tck_name, m.vos_code, m.note_18,
               g.id AS group_id, g.name AS group_name,
               fz.id AS frozen_id
        FROM students s
        LEFT JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON s.group_id = g.id
        LEFT JOIN frozen_students fz
               ON fz.student_id = s.id AND fz.resolved_at IS NULL
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA COLLATE UKRAINIAN,
                 s.first_name_UA COLLATE UKRAINIAN
    """).fetchall()

    # ── Аналіз проблем кожного студента ─────────────────────
    filtered = []
    for r in rows_raw:
        # Заморожені
        if r['frozen_id'] and not include_frozen:
            continue

        # Жінка без військових даних — не показуємо ніде
        if _is_woman_without_military(
            r['gender'], r['military_id'], r['status_badge'],
            r['military_rank'], r['vos_code'], r['tck_name']
        ):
            continue

        problems = []

        # 0. Без ліцензії вступу (Постанова №812)
        if not r['license_id'] and 'no_license' in enabled:
            problems.append('no_license')

        # 1. Без запису military
        if not r['military_id']:
            if 'no_military' in enabled:
                problems.append('no_military')
        else:
            status = r['status_badge'] or ''

            # Виключені (Невійськовозобов'язаний / Знято з обліку)
            # не входять у Додаток 5, тому не є "проблемними"
            if status in EXCLUDED_STATUSES:
                continue

            # 2. Без статі
            if not r['gender'] and 'no_gender' in enabled:
                problems.append('no_gender')

            # 3. Категорія не визначена
            cat = _detect_category(
                r['military_rank'], status, r['gender'], r['category_override']
            )
            if not cat and 'no_category' in enabled:
                problems.append('no_category')

            # 4. Без ТЦК
            if not r['tck_name'] and 'no_tck' in enabled:
                problems.append('no_tck')

            # 5. Без ВОС
            if not r['vos_code'] and 'no_vos' in enabled:
                problems.append('no_vos')

        if not problems:
            continue

        # Додаткові фільтри
        if search:
            haystack = (
                f"{r['last_name_UA']} {r['first_name_UA']} "
                f"{r['middle_name_UA'] or ''}"
            ).lower()
            if not all(w in haystack for w in search.lower().split()):
                continue
        if group_id and r['group_id'] != group_id:
            continue

        row = dict(r)
        row['problems'] = problems
        filtered.append(row)

    total = len(filtered)
    total_pages = (total + per_page - 1) // per_page if total else 1
    if page > total_pages:
        page = total_pages
    start = (page - 1) * per_page
    page_rows = filtered[start:start + per_page]

    # ── Підрахунок проблем по категоріях (для бейджів) ──────
    # Рахуємо по ВСІХ активних (не лише сторінці), але з тим самим
    # фільтром заморожених — щоб бейджик показував актуальне число
    issue_counts = {k: 0 for k in ISSUE_LABELS}
    for r in rows_raw:
        if r['frozen_id'] and not include_frozen:
            continue
        if _is_woman_without_military(
            r['gender'], r['military_id'], r['status_badge'],
            r['military_rank'], r['vos_code'], r['tck_name']
        ):
            continue
        if r['note_18'] in (NOTE_EXCLUDED_BY_AGE, NOTE_DISMISSED, NOTE_MOBILIZED):
            continue

        # Проблема: без ліцензії
        if not r['license_id']:
            issue_counts['no_license'] += 1

        if not r['military_id']:
            issue_counts['no_military'] += 1
            continue
        status = r['status_badge'] or ''
        if status in EXCLUDED_STATUSES:
            continue
        if not r['gender']:
            issue_counts['no_gender'] += 1
        cat = _detect_category(
            r['military_rank'], status, r['gender'], r['category_override']
        )
        if not cat:
            issue_counts['no_category'] += 1
        if not r['tck_name']:
            issue_counts['no_tck'] += 1
        if not r['vos_code']:
            issue_counts['no_vos'] += 1

    conn.close()

    range_start = start + 1 if total else 0
    range_end = min(start + per_page, total)

    return render_template(
        'military_reports/issues.html',
        rows=page_rows,
        total=total,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        range_start=range_start,
        range_end=range_end,
        search=search,
        group_id=group_id,
        include_frozen=include_frozen,
        enabled_issues=enabled,
        issue_labels=ISSUE_LABELS,
        issue_counts=issue_counts,
        groups=groups,
    )

@military_reports_bp.route('/admin/military_reports/summary')
@permission_required('manage_military_reports')
def summary():
    """
    Додаток 12 — Відомість оперативного обліку призовників,
    військовозобов'язаних та резервістів.

    Структура (згідно з Постановою №916, додаток 12):
      1.   Військовозобов'язані (у тому числі резервісти), із них:
      1.1  — офіцерського складу
      1.2  — рядового, сержантського та старшинського складу
      1.3  — жінки
      2.   Призовники

    Поля «Мають мобілізаційні розпорядження», «Заброньовані»,
    «Не заброньовані та не мають моб.розпоряджень» поки що показуються
    прочерком — у БД немає відповідних полів (не входить у поточний обсяг).

    Поля «Зараховано з 1 січня {рік}» і «Відраховано з 1 січня {рік}»
    рахуються:
      - Зараховано: студенти, чия група має start_year = рік
      - Відраховано: накази про відрахування з датою в цьому році

    ⚠️ ВАЖЛИВО: цифри відображають ПОТОЧНИЙ стан БД. Історичних зрізів
    на довільну минулу дату система не має (бо не ведеться історія змін
    статусів). Це явно показується попередженням на сторінці.
    """
    # ── Параметри дати ───────────────────────────────────────
    report_mode = request.args.get('report_mode', 'annual')
    if report_mode not in ('annual', 'current', 'custom'):
        report_mode = 'annual'
    report_date_custom = (request.args.get('report_date') or '').strip()

    dates = _get_report_dates()
    from datetime import datetime as _dt
    if report_mode == 'annual':
        report_date = dates['annual']
    elif report_mode == 'current':
        report_date = dates['current']
    else:
        try:
            report_date = _dt.strptime(report_date_custom, '%d.%m.%Y').date()
        except (ValueError, TypeError):
            report_date = dates['current']

    report_year = report_date.year

    conn = get_db()
    conn.row_factory = sqlite3.Row

    # ── Основні дані (як у dashboard) ────────────────────────
    # Відсіюємо archived та заморожених.
    # Базова ліцензія (Постанова №812)
    show_all_licenses = request.args.get('show_all_licenses') == '1'
    base_license_id, base_license_short, base_license_name = _get_base_license_info(conn)

    students = conn.execute("""
        SELECT s.id, s.gender, s.dismissed_at, s.license_id,
               m.id AS military_id, m.status_badge, m.military_rank,
               m.category_override, m.vos_code, m.tck_name, m.note_18,
               g.start_year AS group_start_year
        FROM students s
        LEFT JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON s.group_id = g.id
        WHERE COALESCE(s.archived, 0) = 0
          AND NOT EXISTS (
              SELECT 1 FROM frozen_students fz
              WHERE fz.student_id = s.id AND fz.resolved_at IS NULL
          )
    """).fetchall()

    # Лічильники по категоріях + по роках
    counts = {'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0}
    enrolled_this_year = {'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0}

    # Нові лічильники (Постанова №812, графа 18) — для рядків 3/4/5
    excluded_by_age = 0
    dismissed_this_year = 0
    mobilized = 0

    for s in students:
        # ── Фільтр базової ліцензії (Постанова №812) ──
        if not show_all_licenses:
            if s['license_id'] is not None and s['license_id'] != base_license_id:
                continue

        if not s['military_id']:
            continue
        status = s['status_badge'] or ''
        if status in EXCLUDED_STATUSES:
            continue

        # Жінка без військових даних — не входить у звіт
        if _is_woman_without_military(
            s['gender'], s['military_id'], s['status_badge'],
            s['military_rank'], s['vos_code'], s['tck_name']
        ):
            continue

        # ── Позначки графи 18 (Постанова №812) ──
        note_18 = s['note_18']

        if note_18 == NOTE_EXCLUDED_BY_AGE:
            excluded_by_age += 1
            continue

        if note_18 == NOTE_MOBILIZED:
            mobilized += 1
            continue

        if note_18 == NOTE_DISMISSED:
            # Враховуємо тільки якщо відраховано в поточному році
            # (примітка до Списків: до кінця поточного календарного року)
            if _is_dismissed_this_year(s['dismissed_at'], report_year):
                dismissed_this_year += 1
            continue

        # Звичайні (на обліку) — рядки 1/2
        category = _detect_category(
            s['military_rank'], status, s['gender'],
            s['category_override'], note_18
        )
        if not category:
            continue
        counts[category] += 1
        if s['group_start_year'] == report_year:
            enrolled_this_year[category] += 1

    # ── Відраховані цього року ───────────────────────────────
    # Беремо всі накази за рік і для кожного студента визначаємо
    # категорію за ПОТОЧНИМ записом military (він міг змінитись після
    # відрахування, але для статистики це прийнятне наближення).
    expelled_rows = conn.execute("""
        SELECT eos.student_id,
               m.status_badge, m.military_rank, m.category_override,
               s.gender
        FROM expulsion_order_students eos
        INNER JOIN expulsion_orders eo ON eo.id = eos.order_id
        LEFT JOIN students s ON s.id = eos.student_id
        LEFT JOIN military m ON m.student_id = eos.student_id
        WHERE eo.order_date >= ? AND eo.order_date < ?
    """, (f"{report_year}-01-01", f"{report_year + 1}-01-01")).fetchall()

    expelled_this_year = {'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0}
    for e in expelled_rows:
        status = e['status_badge'] or ''
        if status in EXCLUDED_STATUSES:
            continue
        category = _detect_category(
            e['military_rank'], status, e['gender'], e['category_override']
        )
        if category:
            expelled_this_year[category] += 1

    conn.close()

    # ── Агрегування для Додатка 12 ───────────────────────────
    mobilized_total = counts['officers'] + counts['rank'] + counts['women']
    total_all = mobilized_total + counts['conscripts']

    enrolled_total = (
        enrolled_this_year['officers']
        + enrolled_this_year['rank']
        + enrolled_this_year['women']
        + enrolled_this_year['conscripts']
    )
    expelled_total = (
        expelled_this_year['officers']
        + expelled_this_year['rank']
        + expelled_this_year['women']
        + expelled_this_year['conscripts']
    )

    # Рядки відповідно до форми Додатка 12
    # (моб.розпорядження / заброньовані / не заброньовані — поки None)
    rows = [
        {
            'num': '1.',
            'name': "Військовозобов'язані (у тому числі резервісти), із них:",
            'total': mobilized_total,
            'mob_order': None,
            'reserved': None,
            'not_reserved': None,
            'enrolled': enrolled_this_year['officers'] + enrolled_this_year['rank'] + enrolled_this_year['women'],
            'expelled': expelled_this_year['officers'] + expelled_this_year['rank'] + expelled_this_year['women'],
            'bold': True,
        },
        {
            'num': '1.1',
            'name': '— офіцерського складу',
            'total': counts['officers'],
            'mob_order': None,
            'reserved': None,
            'not_reserved': None,
            'enrolled': enrolled_this_year['officers'],
            'expelled': expelled_this_year['officers'],
            'bold': False,
        },
        {
            'num': '1.2',
            'name': '— рядового, сержантського та старшинського складу',
            'total': counts['rank'],
            'mob_order': None,
            'reserved': None,
            'not_reserved': None,
            'enrolled': enrolled_this_year['rank'],
            'expelled': expelled_this_year['rank'],
            'bold': False,
        },
        {
            'num': '1.3',
            'name': '— жінки',
            'total': counts['women'],
            'mob_order': None,
            'reserved': None,
            'not_reserved': None,
            'enrolled': enrolled_this_year['women'],
            'expelled': expelled_this_year['women'],
            'bold': False,
        },
        {
            'num': '2.',
            'name': 'Призовники',
            'total': counts['conscripts'],
            'mob_order': None,
            'reserved': None,
            'not_reserved': None,
            'enrolled': enrolled_this_year['conscripts'],
            'expelled': expelled_this_year['conscripts'],
            'bold': True,
        },
    ]

    return render_template(
        'military_reports/summary.html',
        rows=rows,
        total_all=total_all,
        mobilized_total=mobilized_total,
        conscripts_total=counts['conscripts'],
        enrolled_total=enrolled_total,
        expelled_total=expelled_total,
        # Нові (Постанова №812)
        excluded_by_age=excluded_by_age,
        mobilized=mobilized,
        dismissed_this_year=dismissed_this_year,
        show_all_licenses=show_all_licenses,
        base_license_id=base_license_id,
        base_license_label=base_license_short or base_license_name or '—',
        report_mode=report_mode,
        report_date_custom=report_date_custom,
        report_date_ua=_format_date_ua(report_date),
        report_year=report_year,
        dates_formatted={
            'annual': _format_date_ua(dates['annual']),
            'current': _format_date_ua(dates['current']),
        },
    )

# ═══════════════════════════════════════════════════════════════════════
# Генерація .xlsx (Додаток 5, Додаток 12)
# ═══════════════════════════════════════════════════════════════════════

def _fetch_dodatok_5_rows(conn, category, report_date, show_all_licenses,
                          base_license_id, base_license_label):
    """
    Збирає рядки для Додатка 5 з повними полями (для генерації .xlsx).
    Логіка фільтрації — та сама, що й у group_list(), але БЕЗ пагінації
    і без include_frozen / include_dismissed (у звіт потрапляють лише
    активні на обліку).

    Повертає список dict.
    """
    report_year = report_date.year

    # Групи з start_year == report_year (для колонки 23 «Зараховано»)
    # і всі дані по студентах + military + passport + license
    students = conn.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.gender, s.tax_id, s.dismissed_at,
               s.license_id,
               il.short_name_ua AS license_short,
               m.status_badge, m.military_rank, m.category_override,
               m.registry_number, m.vos_code, m.vos_description,
               m.military_registration_document,
               m.address_of_residence,
               m.tck_name,
               m.deferral_type, m.deferral_until,
               m.note_18,
               m.change_credentials,
               m.reason_for_changing_credentials,
               g.start_year AS group_start_year,
               fz.id AS frozen_id
        FROM students s
        INNER JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON s.group_id = g.id
        LEFT JOIN institution_licenses il ON il.id = s.license_id
        LEFT JOIN frozen_students fz
               ON fz.student_id = s.id AND fz.resolved_at IS NULL
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA COLLATE UKRAINIAN,
                 s.first_name_UA COLLATE UKRAINIAN
    """).fetchall()

    # Останні паспорти — одним запитом
    passport_map = {}
    if students:
        ids = [s['id'] for s in students]
        placeholders = ','.join('?' * len(ids))
        for p in conn.execute(f"""
            SELECT student_id, document_type, series, number,
                   issued_by, issue_date
            FROM passport_documents
            WHERE student_id IN ({placeholders})
            ORDER BY student_id, id DESC
        """, ids):
            if p['student_id'] not in passport_map:
                passport_map[p['student_id']] = dict(p)

    result = []
    for s in students:
        # Заморожені — не в звіт
        if s['frozen_id']:
            continue

        # Жінка без військових даних — не в звіт
        if _is_woman_without_military(
            s['gender'], s['id'], s['status_badge'],
            s['military_rank'], s['vos_code'], s['tck_name']
        ):
            continue

        # Виключені статуси
        status = s['status_badge'] or ''
        if status in EXCLUDED_STATUSES:
            continue

        # Вибулі за графою 18 — не в Додаток 5
        if s['note_18'] in NOTE_18_EXCLUDES_FROM_DODATOK_5:
            continue

        # Фільтр ліцензії
        if not show_all_licenses:
            if s['license_id'] is not None and s['license_id'] != base_license_id:
                continue

        # Категорія
        cat = _detect_category(
            s['military_rank'], status, s['gender'],
            s['category_override'], s['note_18']
        )
        if cat != category:
            continue

        # ПІБ
        full_name = ' '.join(
            x for x in [
                s['last_name_UA'], s['first_name_UA'], s['middle_name_UA']
            ] if x
        ).strip()

        # Паспорт (склеюємо)
        passport_text = ''
        p = passport_map.get(s['id'])
        if p:
            parts = []
            if p.get('series'):
                parts.append(p['series'])
            if p.get('number'):
                parts.append(p['number'])
            if p.get('issued_by'):
                parts.append(p['issued_by'])
            if p.get('issue_date'):
                parts.append(p['issue_date'])
            passport_text = ', '.join(parts)

        # ВОС + опис
        vos_text = ''
        if s['vos_code']:
            vos_text = s['vos_code']
            if s['vos_description']:
                vos_text += f" — {s['vos_description']}"

        # Відстрочка
        deferral_text = ''
        if s['deferral_type']:
            deferral_text = s['deferral_type']
            if s['deferral_until']:
                deferral_text += f" до {s['deferral_until']}"

        # Зараховано (колонка 23) — просто позначка, якщо група цього року
        enrolled_note = ''
        if s['group_start_year'] == report_year:
            enrolled_note = f"зараховано {s['group_start_year']}"

        # Відраховано (колонка 24) — тільки якщо dismissed_at цього року
        expelled_note = ''
        if s['dismissed_at']:
            try:
                dy = int(str(s['dismissed_at']).split('-')[0])
                if dy == report_year:
                    expelled_note = f"відраховано {s['dismissed_at']}"
            except (ValueError, IndexError):
                pass

        result.append({
            'full_name': full_name,
            'birth_date': s['birth_date'] or '',
            'tax_id': s['tax_id'] or '',
            'status_badge': status,
            'military_rank': s['military_rank'] or '',
            'registry_number': s['registry_number'] or '',
            'vos': vos_text,
            'military_document': s['military_registration_document'] or '',
            'passport': passport_text,
            'address': s['address_of_residence'] or '',
            'tck_name': s['tck_name'] or '',
            'deferral': deferral_text,
            'note_18': s['note_18'] or '',
            'change_credentials': s['change_credentials'] or '',
            'reason_for_change': s['reason_for_changing_credentials'] or '',
            'license_short': s['license_short'] or '',
            'enrolled_note': enrolled_note,
            'expelled_note': expelled_note,
        })

    return result


@military_reports_bp.route('/admin/military_reports/export/dodatok_5/<category>')
@permission_required('manage_military_reports')
def export_dodatok_5(category):
    """
    Генерація .xlsx Додатка 5 для однієї категорії.

    Query-параметри:
        report_mode=annual|current|custom  (як на сторінці)
        report_date=ДД.ММ.РРРР (для custom)
        show_all_licenses=1  (як на сторінці)
    """
    from flask import send_file
    from routes.military_reports_export import generate_dodatok_5_xlsx
    from datetime import datetime as _dt

    if category not in CATEGORY_LABELS:
        flash("Невідома категорія", "error")
        return redirect(url_for('military_reports.dashboard'))

    label, _subtitle = CATEGORY_LABELS[category]

    # ── Параметри дати ──
    report_mode = request.args.get('report_mode', 'annual')
    if report_mode not in ('annual', 'current', 'custom'):
        report_mode = 'annual'
    report_date_custom = (request.args.get('report_date') or '').strip()

    dates = _get_report_dates()
    if report_mode == 'annual':
        report_date = dates['annual']
    elif report_mode == 'current':
        report_date = dates['current']
    else:
        try:
            report_date = _dt.strptime(report_date_custom, '%d.%m.%Y').date()
        except (ValueError, TypeError):
            report_date = dates['current']

    show_all_licenses = request.args.get('show_all_licenses') == '1'

    conn = get_db()
    conn.row_factory = sqlite3.Row

    base_license_id, base_license_short, base_license_name = _get_base_license_info(conn)

    rows = _fetch_dodatok_5_rows(
        conn, category, report_date, show_all_licenses,
        base_license_id, base_license_short or base_license_name,
    )
    conn.close()

    if not rows:
        flash(
            "Немає студентів для генерації — перевірте, чи є активні "
            "студенти цієї категорії для обраної ліцензії.",
            "warning",
        )
        return redirect(url_for(
            'military_reports.group_list',
            category=category,
            report_mode=report_mode,
            report_date=report_date_custom if report_mode == 'custom' else None,
            show_all_licenses='1' if show_all_licenses else None,
        ))

    xlsx_io = generate_dodatok_5_xlsx(
        category, report_date, rows,
        base_license_label=base_license_short or base_license_name,
    )

    # ── Логування в журнал ──
    rtype = f'dodatok_5_{category}'
    filename = (
        f"Додаток_5_{label.replace(' ', '_').replace(',', '')}_"
        f"{_format_date_ua(report_date).replace('.', '-')}.xlsx"
    )
    try:
        log_conn = get_db()
        log_conn.execute(
            "INSERT INTO military_reports_log "
            "(report_type, report_date, generated_by, rows_count, filename) "
            "VALUES (?, ?, ?, ?, ?)",
            (rtype, report_date.isoformat(), current_username(),
             len(rows), filename),
        )
        log_conn.commit()
        log_conn.close()
    except Exception as e:
        logger.error(f"Не вдалося записати в журнал генерацій: {e}", exc_info=True)

    log_action(
        current_username(),
        f"згенерував Додаток 5 ({label}): {len(rows)} осіб, "
        f"станом на {_format_date_ua(report_date)}"
    )

    return send_file(
        xlsx_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )

def _fetch_dodatok_12_data(conn, report_date, show_all_licenses, base_license_id):
    """
    Збирає підрахунки для Додатка 12.
    Логіка фільтрації — та сама, що й у summary().
    """
    report_year = report_date.year

    students = conn.execute("""
        SELECT s.id, s.gender, s.dismissed_at, s.license_id,
               m.id AS military_id, m.status_badge, m.military_rank,
               m.category_override, m.vos_code, m.tck_name, m.note_18,
               g.start_year AS group_start_year
        FROM students s
        LEFT JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON s.group_id = g.id
        WHERE COALESCE(s.archived, 0) = 0
          AND NOT EXISTS (
              SELECT 1 FROM frozen_students fz
              WHERE fz.student_id = s.id AND fz.resolved_at IS NULL
          )
    """).fetchall()

    counts = {'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0}
    enrolled = {'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0}
    expelled = {'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0}
    excluded_by_age = 0
    dismissed_this_year = 0
    mobilized = 0

    for s in students:
        if not show_all_licenses:
            if s['license_id'] is not None and s['license_id'] != base_license_id:
                continue

        if not s['military_id']:
            continue
        status = s['status_badge'] or ''
        if status in EXCLUDED_STATUSES:
            continue

        if _is_woman_without_military(
            s['gender'], s['military_id'], s['status_badge'],
            s['military_rank'], s['vos_code'], s['tck_name']
        ):
            continue

        note_18 = s['note_18']

        if note_18 == NOTE_EXCLUDED_BY_AGE:
            excluded_by_age += 1
            continue

        if note_18 == NOTE_MOBILIZED:
            mobilized += 1
            continue

        if note_18 == NOTE_DISMISSED:
            if _is_dismissed_this_year(s['dismissed_at'], report_year):
                dismissed_this_year += 1
                # Категорію теж враховуємо (для колонок «відраховано»)
                cat = _detect_category(
                    s['military_rank'], status, s['gender'],
                    s['category_override'], note_18=None
                )
                if cat:
                    expelled[cat] += 1
            continue

        category = _detect_category(
            s['military_rank'], status, s['gender'],
            s['category_override'], note_18
        )
        if not category:
            continue
        counts[category] += 1
        if s['group_start_year'] == report_year:
            enrolled[category] += 1

    return {
        'officers': counts['officers'],
        'rank': counts['rank'],
        'women': counts['women'],
        'conscripts': counts['conscripts'],
        'excluded_by_age': excluded_by_age,
        'dismissed_this_year': dismissed_this_year,
        'mobilized': mobilized,
        'enrolled': enrolled,
        'expelled': expelled,
    }


@military_reports_bp.route('/admin/military_reports/export/dodatok_12')
@permission_required('manage_military_reports')
def export_dodatok_12():
    """
    Генерація .xlsx Додатка 12 (зведена відомість).

    Query-параметри:
        report_mode=annual|current|custom
        report_date=ДД.ММ.РРРР (для custom)
        show_all_licenses=1
    """
    from flask import send_file
    from routes.military_reports_export import generate_dodatok_12_xlsx
    from datetime import datetime as _dt

    # ── Параметри дати ──
    report_mode = request.args.get('report_mode', 'annual')
    if report_mode not in ('annual', 'current', 'custom'):
        report_mode = 'annual'
    report_date_custom = (request.args.get('report_date') or '').strip()

    dates = _get_report_dates()
    if report_mode == 'annual':
        report_date = dates['annual']
    elif report_mode == 'current':
        report_date = dates['current']
    else:
        try:
            report_date = _dt.strptime(report_date_custom, '%d.%m.%Y').date()
        except (ValueError, TypeError):
            report_date = dates['current']

    show_all_licenses = request.args.get('show_all_licenses') == '1'

    conn = get_db()
    conn.row_factory = sqlite3.Row

    base_license_id, base_license_short, base_license_name = _get_base_license_info(conn)

    data = _fetch_dodatok_12_data(
        conn, report_date, show_all_licenses, base_license_id,
    )
    conn.close()

    data['base_license_label'] = base_license_short or base_license_name or '—'
    data['show_all_licenses'] = show_all_licenses

    xlsx_io = generate_dodatok_12_xlsx(data, report_date)

    # ── Логування ──
    filename = (
        f"Додаток_12_"
        f"{_format_date_ua(report_date).replace('.', '-')}.xlsx"
    )
    total_students = (
        data['officers'] + data['rank'] + data['women'] + data['conscripts']
        + data['excluded_by_age'] + data['dismissed_this_year'] + data['mobilized']
    )
    try:
        log_conn = get_db()
        log_conn.execute(
            "INSERT INTO military_reports_log "
            "(report_type, report_date, generated_by, rows_count, filename) "
            "VALUES (?, ?, ?, ?, ?)",
            ('dodatok_12', report_date.isoformat(), current_username(),
             total_students, filename),
        )
        log_conn.commit()
        log_conn.close()
    except Exception as e:
        logger.error(f"Не вдалося записати в журнал генерацій: {e}", exc_info=True)

    log_action(
        current_username(),
        f"згенерував Додаток 12: {total_students} осіб, "
        f"станом на {_format_date_ua(report_date)}"
    )

    return send_file(
        xlsx_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )


@military_reports_bp.route('/admin/military_reports/tck', methods=['GET', 'POST'])
@permission_required('manage_users')
def tck_offices():
    """
    Довідник ТЦК — CRUD-сторінка (тільки адмін).

    Дозволені дії:
      - add              — додати новий ТЦК
      - edit             — редагувати назву/адресу; опційно каскадно
                           оновити назву у студентів (military.tck_name)
      - delete           — видалити (тільки якщо не використовується)
      - toggle           — деактивувати / активувати
      - add_unknown      — додати один ТЦК зі списку «невідомих»
                           (тих, що вже вказані у студентів, але відсутні
                           в довіднику). Опційно об'єднати всі варіанти
                           написання в одне значення
      - add_all_unknown  — додати всі невідомі ТЦК однією дією
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row

    if request.method == 'POST':
        action = request.form.get('action')

        # ── Додати ───────────────────────────────────────────
        if action == 'add':
            name_ua = (request.form.get('name_ua') or '').strip()
            address = (request.form.get('address') or '').strip() or None
            if not name_ua:
                flash("Вкажіть назву ТЦК", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO tck_offices (name_ua, address) VALUES (?, ?)",
                        (name_ua, address),
                    )
                    conn.commit()
                    log_action(current_username(),
                               f"додав ТЦК у довідник: {name_ua}")
                    flash(f"ТЦК «{name_ua}» додано", "success")
                except sqlite3.IntegrityError:
                    flash(f"ТЦК з назвою «{name_ua}» вже існує", "error")

        # ── Редагувати ───────────────────────────────────────
        elif action == 'edit':
            tck_id = request.form.get('id', type=int)
            old_name = (request.form.get('old_name_ua') or '').strip()
            new_name = (request.form.get('name_ua') or '').strip()
            new_address = (request.form.get('address') or '').strip() or None
            cascade = request.form.get('cascade') == '1'

            if not tck_id or not new_name:
                flash("Вкажіть назву ТЦК", "error")
            else:
                try:
                    conn.execute(
                        "UPDATE tck_offices SET name_ua = ?, address = ? WHERE id = ?",
                        (new_name, new_address, tck_id),
                    )
                    cascaded = 0
                    if cascade and old_name and old_name != new_name:
                        cur = conn.execute(
                            "UPDATE military SET tck_name = ? WHERE tck_name = ?",
                            (new_name, old_name),
                        )
                        cascaded = cur.rowcount
                    conn.commit()
                    details = (
                        f"перейменовано з «{old_name}»" if old_name != new_name else "оновлено"
                    )
                    if cascaded:
                        details += f", оновлено у {cascaded} студентів"
                    log_action(current_username(),
                               f"редагував ТЦК (ID {tck_id}): {new_name}",
                               details=details)
                    flash(
                        f"ТЦК «{new_name}» оновлено" +
                        (f" (у {cascaded} студентів)" if cascaded else ""),
                        "success",
                    )
                except sqlite3.IntegrityError:
                    flash(f"ТЦК з назвою «{new_name}» вже існує", "error")

        # ── Видалити ─────────────────────────────────────────
        elif action == 'delete':
            tck_id = request.form.get('id', type=int)
            row = conn.execute(
                "SELECT name_ua FROM tck_offices WHERE id = ?", (tck_id,)
            ).fetchone()
            if not row:
                flash("ТЦК не знайдено", "error")
            else:
                in_use = conn.execute(
                    "SELECT COUNT(*) FROM military "
                    "WHERE TRIM(LOWER_UA(tck_name)) = TRIM(LOWER_UA(?))",
                    (row['name_ua'],),
                ).fetchone()[0]
                if in_use > 0:
                    flash(
                        f"Неможливо видалити ТЦК «{row['name_ua']}» — "
                        f"він використовується у {in_use} студентів. "
                        f"Спершу змініть ТЦК у цих студентів або "
                        f"деактивуйте запис (перемикач «Активний»).",
                        "error",
                    )
                else:
                    conn.execute("DELETE FROM tck_offices WHERE id = ?", (tck_id,))
                    conn.commit()
                    log_action(current_username(),
                               f"ВИДАЛИВ ТЦК з довідника: {row['name_ua']}")
                    flash(f"ТЦК «{row['name_ua']}» видалено", "success")

        # ── Деактивувати / активувати ────────────────────────
        elif action == 'toggle':
            tck_id = request.form.get('id', type=int)
            row = conn.execute(
                "SELECT name_ua, is_active FROM tck_offices WHERE id = ?",
                (tck_id,),
            ).fetchone()
            if row:
                new_state = 0 if row['is_active'] else 1
                conn.execute(
                    "UPDATE tck_offices SET is_active = ? WHERE id = ?",
                    (new_state, tck_id),
                )
                conn.commit()
                log_action(current_username(),
                           f"{'увімкнув' if new_state else 'вимкнув'} ТЦК: {row['name_ua']}")
                flash(
                    f"ТЦК «{row['name_ua']}» тепер "
                    f"{'активний' if new_state else 'неактивний'}",
                    "success",
                )

        # ── Додати один «невідомий» ТЦК ──────────────────────
        elif action == 'add_unknown':
            name_ua = (request.form.get('name_ua') or '').strip()
            cascade = request.form.get('cascade') == '1'
            if not name_ua:
                flash("Не вказано назву ТЦК", "error")
            else:
                try:
                    conn.execute(
                        "INSERT INTO tck_offices (name_ua) VALUES (?)",
                        (name_ua,),
                    )
                    merged = 0
                    if cascade:
                        # Об'єднуємо всі варіанти написання в одне значення:
                        # знаходимо всі tck_name у military, чия нормалізована
                        # форма співпадає з нормалізованою формою name_ua,
                        # і замінюємо на name_ua.
                        variants = conn.execute("""
                            SELECT DISTINCT tck_name FROM military
                            WHERE tck_name IS NOT NULL AND TRIM(tck_name) != ''
                              AND TRIM(LOWER_UA(tck_name)) = TRIM(LOWER_UA(?))
                        """, (name_ua,)).fetchall()
                        for v in variants:
                            if (v['tck_name'] or '').strip() != name_ua:
                                cur = conn.execute(
                                    "UPDATE military SET tck_name = ? WHERE tck_name = ?",
                                    (name_ua, v['tck_name']),
                                )
                                merged += cur.rowcount
                    conn.commit()
                    msg = f"ТЦК «{name_ua}» додано в довідник"
                    if merged:
                        msg += f" (об'єднано {merged} студентів з різними написаннями)"
                    log_action(current_username(),
                               f"додав ТЦК зі списку невідомих: {name_ua}",
                               details=f"об'єднано студентів: {merged}")
                    flash(msg, "success")
                except sqlite3.IntegrityError:
                    flash(f"ТЦК з назвою «{name_ua}» вже існує", "error")

        # ── Додати ВСІ невідомі ТЦК ──────────────────────────
        elif action == 'add_all_unknown':
            unknown_rows = conn.execute("""
                SELECT MIN(m.tck_name) AS canonical_name
                FROM military m
                WHERE m.tck_name IS NOT NULL AND TRIM(m.tck_name) != ''
                  AND NOT EXISTS (
                      SELECT 1 FROM tck_offices t
                      WHERE TRIM(LOWER_UA(t.name_ua)) = TRIM(LOWER_UA(m.tck_name))
                  )
                GROUP BY TRIM(LOWER_UA(m.tck_name))
            """).fetchall()

            added, skipped = 0, 0
            for u in unknown_rows:
                name = (u['canonical_name'] or '').strip()
                if not name:
                    continue
                try:
                    conn.execute(
                        "INSERT INTO tck_offices (name_ua) VALUES (?)",
                        (name,),
                    )
                    added += 1
                except sqlite3.IntegrityError:
                    skipped += 1
            conn.commit()
            log_action(current_username(),
                       f"додав усі невідомі ТЦК у довідник: {added}",
                       details=f"пропущено дублів: {skipped}")
            if added:
                flash(
                    f"Додано {added} ТЦК у довідник" +
                    (f" (пропущено {skipped} дублів)" if skipped else ""),
                    "success",
                )
            else:
                flash("Немає нових ТЦК для додавання", "warning")

        conn.close()
        return redirect(url_for('military_reports.tck_offices'))

    # ── GET: список ТЦК + статистика + невідомі ─────────────
    search = (request.args.get('search') or '').strip()
    show = request.args.get('show', 'all')  # all | active | inactive
    if show not in ('all', 'active', 'inactive'):
        show = 'all'

    where_parts = []
    params = []
    if show == 'active':
        where_parts.append("is_active = 1")
    elif show == 'inactive':
        where_parts.append("is_active = 0")
    if search:
        where_parts.append("(name_ua LIKE ? OR address LIKE ?)")
        params.extend([f'%{search}%', f'%{search}%'])

    where_sql = ("WHERE " + " AND ".join(where_parts)) if where_parts else ""

    rows_raw = conn.execute(f"""
        SELECT t.id, t.name_ua, t.address, t.is_active,
               -- Активні на обліку (без note_18, не архівні, не заморожені)
               (SELECT COUNT(*) FROM students s
                INNER JOIN military m ON m.student_id = s.id
                WHERE TRIM(LOWER_UA(m.tck_name)) = TRIM(LOWER_UA(t.name_ua))
                  AND COALESCE(s.archived, 0) = 0
                  AND m.note_18 IS NULL
                  AND NOT EXISTS (SELECT 1 FROM frozen_students fz
                                 WHERE fz.student_id = s.id AND fz.resolved_at IS NULL)
               ) AS students_count_active,
               -- Архівні
               (SELECT COUNT(*) FROM students s
                INNER JOIN military m ON m.student_id = s.id
                WHERE TRIM(LOWER_UA(m.tck_name)) = TRIM(LOWER_UA(t.name_ua))
                  AND COALESCE(s.archived, 0) = 1
               ) AS students_count_archived,
               -- Заморожені
               (SELECT COUNT(*) FROM students s
                INNER JOIN military m ON m.student_id = s.id
                WHERE TRIM(LOWER_UA(m.tck_name)) = TRIM(LOWER_UA(t.name_ua))
                  AND COALESCE(s.archived, 0) = 0
                  AND EXISTS (SELECT 1 FROM frozen_students fz
                             WHERE fz.student_id = s.id AND fz.resolved_at IS NULL)
               ) AS students_count_frozen,
               -- Вибулі за графою 18 (виключено за віком / відраховано / мобілізовано)
               (SELECT COUNT(*) FROM students s
                INNER JOIN military m ON m.student_id = s.id
                WHERE TRIM(LOWER_UA(m.tck_name)) = TRIM(LOWER_UA(t.name_ua))
                  AND COALESCE(s.archived, 0) = 0
                  AND m.note_18 IS NOT NULL
                  AND NOT EXISTS (SELECT 1 FROM frozen_students fz
                                 WHERE fz.student_id = s.id AND fz.resolved_at IS NULL)
               ) AS students_count_note18
        FROM tck_offices t
        {where_sql}
        ORDER BY t.name_ua COLLATE UKRAINIAN
    """, params).fetchall()

    rows = []
    for r in rows_raw:
        item = dict(r)
        item['students_count_all'] = (
            item['students_count_active']
            + item['students_count_archived']
            + item['students_count_frozen']
            + item['students_count_note18']
        )
        rows.append(item)

    # Невідомі ТЦК: є у military, але відсутні в довіднику.
    #
    # SQLite не підтримує GROUP_CONCAT(DISTINCT x, separator) —
    # з DISTINCT дозволено тільки один аргумент. Тому список варіантів
    # написання збираємо в корельованому підзапиті з власним DISTINCT
    # (без separator — беремо дефолтну кому і потім у Python замінюємо
    # на ' | ' для читабельності).
    unknown_tck_raw = conn.execute("""
        SELECT
            g.canonical_name,
            g.students_count,
            g.variants_count,
            (SELECT GROUP_CONCAT(v.variant)
             FROM (
                 SELECT DISTINCT TRIM(m2.tck_name) AS variant
                 FROM military m2
                 WHERE TRIM(LOWER_UA(m2.tck_name)) = g.canonical_key
                   AND m2.tck_name IS NOT NULL
                   AND TRIM(m2.tck_name) != ''
                 ORDER BY TRIM(m2.tck_name)
             ) v
            ) AS variants
        FROM (
            SELECT
                MIN(m.tck_name) AS canonical_name,
                TRIM(LOWER_UA(m.tck_name)) AS canonical_key,
                COUNT(*) AS students_count,
                COUNT(DISTINCT TRIM(m.tck_name)) AS variants_count
            FROM military m
            WHERE m.tck_name IS NOT NULL
              AND TRIM(m.tck_name) != ''
              AND NOT EXISTS (
                  SELECT 1 FROM tck_offices t
                  WHERE TRIM(LOWER_UA(t.name_ua)) = TRIM(LOWER_UA(m.tck_name))
              )
            GROUP BY TRIM(LOWER_UA(m.tck_name))
        ) g
        ORDER BY g.students_count DESC, g.canonical_name COLLATE UKRAINIAN
    """).fetchall()

    # Перетворюємо Row у dict і форматуємо список варіантів
    # (коми → ' | ') для читабельного відображення в template
    unknown_tck = []
    for u in unknown_tck_raw:
        item = dict(u)
        if item.get('variants'):
            item['variants'] = ' | '.join(
                v.strip() for v in item['variants'].split(',') if v.strip()
            )
        else:
            item['variants'] = ''
        unknown_tck.append(item)
    # Загальна статистика
    stats = {
        'total': conn.execute("SELECT COUNT(*) FROM tck_offices").fetchone()[0],
        'active': conn.execute(
            "SELECT COUNT(*) FROM tck_offices WHERE is_active = 1"
        ).fetchone()[0],
        'inactive': conn.execute(
            "SELECT COUNT(*) FROM tck_offices WHERE is_active = 0"
        ).fetchone()[0],
        'in_use': conn.execute(
            "SELECT COUNT(DISTINCT tck_name) FROM military "
            "WHERE tck_name IS NOT NULL AND tck_name != ''"
        ).fetchone()[0],
        'unknown_count': len(unknown_tck),
        'unknown_students': sum(u['students_count'] for u in unknown_tck),
    }

    conn.close()

    return render_template(
        'military_reports/tck.html',
        rows=rows,
        unknown_tck=unknown_tck,
        search=search,
        show=show,
        stats=stats,
    )
    
@military_reports_bp.route('/admin/military_reports/tck/<int:tck_id>/students')
@permission_required('manage_users')
def tck_students(tck_id):
    """
    AJAX: список студентів, які використовують конкретний ТЦК.

    Query-параметри:
        include_archived=1  — показати також архівних / відрахованих
        include_frozen=1    — показати також заморожених
    За замовчуванням обидва вимкнені → показуються лише активні.

    Повертає JSON зі списком студентів + розбивка по категоріях
    (active / archived / frozen / total_all) для показу в модалці.
    """
    include_archived = request.args.get('include_archived') == '1'
    include_frozen = request.args.get('include_frozen') == '1'
    include_dismissed = request.args.get('include_dismissed') == '1'

    conn = get_db()
    conn.row_factory = sqlite3.Row

    tck = conn.execute(
        "SELECT id, name_ua, address FROM tck_offices WHERE id = ?",
        (tck_id,),
    ).fetchone()

    if not tck:
        conn.close()
        return jsonify({'ok': False, 'error': 'ТЦК не знайдено'}), 404

    # ── Базовий SQL з умовними фільтрами ─────────────────────
    base_sql = """
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.gender, s.archived,
               m.id AS military_id, m.status_badge, m.military_rank,
               m.category_override, m.vos_code,
               g.name AS group_name,
               fz.id AS frozen_id
        FROM students s
        INNER JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON s.group_id = g.id
        LEFT JOIN frozen_students fz
               ON fz.student_id = s.id AND fz.resolved_at IS NULL
        WHERE TRIM(LOWER_UA(m.tck_name)) = TRIM(LOWER_UA(?))
    """
    extra_conditions = []
    if not include_archived:
        extra_conditions.append("COALESCE(s.archived, 0) = 0")
    if not include_frozen:
        extra_conditions.append("fz.id IS NULL")
    if extra_conditions:
        base_sql += " AND " + " AND ".join(extra_conditions)
    base_sql += """
        ORDER BY s.last_name_UA COLLATE UKRAINIAN,
                 s.first_name_UA COLLATE UKRAINIAN
    """

    rows = conn.execute(base_sql, (tck['name_ua'],)).fetchall()

    # ── Розбивка по категоріях (для відображення в шапці модалки) ──
    counts_row = conn.execute("""
        SELECT
            COUNT(*) AS total_all,
            SUM(CASE WHEN COALESCE(s.archived, 0) = 0 AND fz.id IS NULL
                     THEN 1 ELSE 0 END) AS total_active,
            SUM(CASE WHEN COALESCE(s.archived, 0) = 1
                     THEN 1 ELSE 0 END) AS total_archived,
            SUM(CASE WHEN COALESCE(s.archived, 0) = 0 AND fz.id IS NOT NULL
                     THEN 1 ELSE 0 END) AS total_frozen
        FROM students s
        INNER JOIN military m ON m.student_id = s.id
        LEFT JOIN frozen_students fz
               ON fz.student_id = s.id AND fz.resolved_at IS NULL
        WHERE TRIM(LOWER_UA(m.tck_name)) = TRIM(LOWER_UA(?))
    """, (tck['name_ua'],)).fetchone()

    CATEGORY_LABELS_SHORT = {
        'officers':   'Офіцер',
        'rank':       'Рядовий/сержант',
        'women':      'Жінка',
        'conscripts': 'Призовник',
    }

    students = []
    for r in rows:
        status = r['status_badge'] or ''
        category = _detect_category(
            r['military_rank'], status, r['gender'], r['category_override']
        )
        students.append({
            'id': r['id'],
            'full_name': (
                f"{r['last_name_UA']} {r['first_name_UA']} "
                f"{r['middle_name_UA'] or ''}"
            ).strip(),
            'birth_date': r['birth_date'] or '',
            'category': CATEGORY_LABELS_SHORT.get(category, ''),
            'category_key': category or '',
            'rank': r['military_rank'] or '',
            'vos': r['vos_code'] or '',
            'status': status,
            'group': r['group_name'] or '',
            'gender': r['gender'] or '',
            'is_frozen': bool(r['frozen_id']),
            'is_archived': bool(r['archived']),
            'is_excluded': status in EXCLUDED_STATUSES,
            'note_18': r['note_18'] or '',
        })

    conn.close()

    return jsonify({
        'ok': True,
        'tck_name': tck['name_ua'],
        'tck_address': tck['address'] or '',
        'students': students,
        'shown': len(students),
        'counts': {
            'total_all':      counts_row['total_all'] or 0,
            'total_active':   counts_row['total_active'] or 0,
            'total_archived': counts_row['total_archived'] or 0,
            'total_frozen':   counts_row['total_frozen'] or 0,
        },
    })