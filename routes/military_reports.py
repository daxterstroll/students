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

import io
import os
import re
import sqlite3
import hashlib
import zipfile
from datetime import date

from flask import (
    Blueprint, render_template, redirect, url_for,
    flash, request, jsonify, session, send_file,
)
from routes.db import get_db
from routes.utils import log_action, permission_required, logger, PROJECT_ROOT
from routes.helpers import current_username
from routes.orders_service import (
    get_enrollment_student_ids,
    orders_tables_ready,
)

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


def _year_from_order_date(order_date):
    """Рік з DD.MM.YYYY / DD/MM/YYYY / YYYY-MM-DD / YYYY."""
    if not order_date:
        return None
    s = str(order_date).strip()
    for sep in ('.', '/'):
        parts = s.split(sep)
        if len(parts) == 3 and len(parts[2]) == 4 and parts[2].isdigit():
            try:
                return int(parts[2])
            except ValueError:
                return None
    if len(s) >= 4 and s[0:4].isdigit():
        if len(s) == 4 or s[4] in '-./':
            try:
                return int(s[0:4])
            except ValueError:
                return None
    return None


def _load_enrollment_years(conn):
    """
    student_id → рік наказу про зарахування (найраніший, якщо кілька).
    Порожній dict, якщо таблиці orders ще немає.
    """
    if not orders_tables_ready(conn=conn):
        return {}
    rows = conn.execute("""
        SELECT oi.entity_id AS student_id, o.order_date
        FROM orders o
        JOIN order_items oi ON oi.order_id = o.id
        WHERE o.status = 'active'
          AND o.order_type = 'enrollment'
          AND oi.entity_type = 'student'
    """).fetchall()
    result = {}
    for r in rows:
        y = _year_from_order_date(r['order_date'])
        if y is None:
            continue
        sid = r['student_id']
        prev = result.get(sid)
        if prev is None or y < prev:
            result[sid] = y
    return result


def _parse_order_date_to_date(order_date):
    """DD.MM.YYYY / DD/MM/YYYY / YYYY-MM-DD → datetime.date або None."""
    from datetime import date as _date
    if not order_date:
        return None
    s = str(order_date).strip()
    for sep in ('.', '/'):
        parts = s.split(sep)
        if len(parts) == 3 and len(parts[2]) == 4:
            try:
                d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                return _date(y, m, d)
            except ValueError:
                return None
    if len(s) >= 10 and s[4] == '-' and s[7] == '-':
        try:
            return _date(int(s[0:4]), int(s[5:7]), int(s[8:10]))
        except ValueError:
            return None
    return None


def _load_enrollment_dates(conn):
    """student_id → date найранішого наказу enrollment."""
    if not orders_tables_ready(conn=conn):
        return {}
    rows = conn.execute("""
        SELECT oi.entity_id AS student_id, o.order_date
        FROM orders o
        JOIN order_items oi ON oi.order_id = o.id
        WHERE o.status = 'active'
          AND o.order_type = 'enrollment'
          AND oi.entity_type = 'student'
    """).fetchall()
    result = {}
    for r in rows:
        d = _parse_order_date_to_date(r['order_date'])
        if d is None:
            continue
        sid = r['student_id']
        prev = result.get(sid)
        if prev is None or d < prev:
            result[sid] = d
    return result


def _load_enrollment_or_renewal_orders(conn, report_date=None):
    """
    student_id → текст наказу про зарахування або поновлення для 17 графи Додатка 5.
    Береться активний наказ типу 'enrollment' або 'license_transfer' (ПОНОВЛЕННЯ)
    з order_items (entity_type='student') або student_enrollments.order_id.
    Якщо є кілька наказів — береться найактуальніший (спочатку <= report_date,
    інакше найновіший активний наказ студента).
    """
    if not orders_tables_ready(conn=conn):
        return {}

    candidates = {}

    def _add(sid, order_id, order_type, order_number, order_date):
        if sid is None or not order_number:
            return
        sid = int(sid)
        p_date = _parse_order_date_to_date(order_date)
        candidates.setdefault(sid, {})[int(order_id)] = (
            p_date,
            int(order_id),
            order_type,
            (order_number or '').strip(),
            (order_date or '').strip(),
        )

    rows = conn.execute("""
        SELECT oi.entity_id AS student_id,
               o.id AS order_id,
               o.order_type,
               o.order_number,
               o.order_date
        FROM orders o
        JOIN order_items oi ON oi.order_id = o.id
        WHERE o.status = 'active'
          AND o.order_type IN ('enrollment', 'license_transfer')
          AND oi.entity_type = 'student'
    """).fetchall()
    for r in rows:
        _add(r['student_id'], r['order_id'], r['order_type'], r['order_number'], r['order_date'])

    enr_rows = conn.execute("""
        SELECT e.student_id,
               o.id AS order_id,
               o.order_type,
               o.order_number,
               o.order_date
        FROM student_enrollments e
        JOIN orders o ON o.id = e.order_id
        WHERE e.status = 'active'
          AND o.status = 'active'
          AND o.order_type IN ('enrollment', 'license_transfer')
    """).fetchall()
    for r in enr_rows:
        _add(r['student_id'], r['order_id'], r['order_type'], r['order_number'], r['order_date'])

    result = {}
    for sid, order_map in candidates.items():
        items = list(order_map.values())
        if report_date is not None:
            eligible = [it for it in items if it[0] is not None and it[0] <= report_date]
            pool = eligible if eligible else items
        else:
            pool = items
        best = max(pool, key=lambda it: (it[0] or date.min, it[1]))
        p_date, _oid, otype, num, raw_date = best
        date_str = _format_date_ua(p_date) if p_date is not None else raw_date
        prefix = 'Наказ про поновлення' if otype == 'license_transfer' else 'Наказ про зарахування'
        if date_str:
            result[sid] = f"{prefix} №{num} від {date_str}"
        else:
            result[sid] = f"{prefix} №{num}"
    return result


STUDY_FORM_OPTIONS = ('Денна', 'Заочна', 'Дистанційна')


def _empty_study_form_counts():
    return {'Денна': 0, 'Заочна': 0, 'Дистанційна': 0, 'other': 0}


def _bump_study_form(bucket, study_form):
    key = study_form if study_form in STUDY_FORM_OPTIONS else 'other'
    bucket[key] += 1


def _was_enrolled_as_of(report_date, group_start_year, enrollment_date):
    """
    Чи студент уже мав бути на навчанні станом на report_date.
    - є наказ enrollment → дата наказу <= report_date;
    - немає → умовна дата 1 вересня group_start_year.
    """
    from datetime import date as _date
    if enrollment_date is not None:
        return enrollment_date <= report_date
    if group_start_year:
        try:
            y = int(group_start_year)
        except (TypeError, ValueError):
            return True
        return _date(y, 9, 1) <= report_date
    return True




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

def _load_enrollment_licenses(conn) -> dict:
    """student_id → set(license_id) усіх активних enrollment'ів.
    Один SQL на весь дашборд замість N+1 у циклі."""
    rows = conn.execute("""
        SELECT student_id, license_id
        FROM student_enrollments
        WHERE status = 'active' AND license_id IS NOT NULL
    """).fetchall()
    result = {}
    for r in rows:
        sid = r['student_id'] if hasattr(r, 'keys') else r[0]
        lid = r['license_id'] if hasattr(r, 'keys') else r[1]
        result.setdefault(sid, set()).add(lid)
    return result


def _load_enrollment_license_shorts(conn) -> dict:
    """student_id → список коротких назв ліцензій (для відображення)."""
    rows = conn.execute("""
        SELECT e.student_id,
               COALESCE(il.short_name_ua, il.name_ua) AS short
        FROM student_enrollments e
        JOIN institution_licenses il ON il.id = e.license_id
        WHERE e.status = 'active' AND e.license_id IS NOT NULL
        ORDER BY il.id
    """).fetchall()
    result = {}
    for r in rows:
        sid = r['student_id'] if hasattr(r, 'keys') else r[0]
        sh = r['short'] if hasattr(r, 'keys') else r[1]
        result.setdefault(sid, []).append(sh)
    return result


def _has_base_license(student_id, enrollment_licenses, base_license_id):
    """True, якщо студент має активний enrollment під базовою ліцензією,
    АБО взагалі не має активних enrollment'ів з ліцензією (тоді показуємо
    з бейджем «без ліцензії» — дані потребують виправлення)."""
    if base_license_id is None:
        return True
    lic_ids = enrollment_licenses.get(student_id, set())
    if not lic_ids:
        return True
    return base_license_id in lic_ids


def _normalize_tck_key(value):
    return (value or '').strip().lower()

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
    enrollment_licenses = _load_enrollment_licenses(conn)
    students = conn.execute("""
        SELECT s.id, s.gender, s.dismissed_at, s.license_id,
               m.id AS military_id, m.status_badge, m.military_rank,
               m.category_override, m.tck_name, m.vos_code, m.note_18,
               g.study_form AS study_form,
               fz.id AS frozen_id
        FROM students s
        LEFT JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON s.group_id = g.id
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
    study_form_totals = _empty_study_form_counts()
    study_form_by_cat = {
        'officers': _empty_study_form_counts(),
        'rank': _empty_study_form_counts(),
        'women': _empty_study_form_counts(),
        'conscripts': _empty_study_form_counts(),
    }
    issues = {
        'no_tck': 0,
        'no_vos': 0,
        'no_category': 0,
        'no_gender': 0,
        'no_military': 0,
        'no_license': 0,
        'no_enrollment_order': 0,
    }
    frozen_count = 0
    total_students = 0

    enrolled_ids = set()
    if orders_tables_ready(conn=conn):
        enrolled_ids = get_enrollment_student_ids(conn=conn)

    for s in students:
        # ── 0. Фільтр ліцензії (Постанова №812) ──
        # Студент належить базовій ліцензії, якщо хоч один його АКТИВНИЙ
        # enrollment — під нею. Студенти без ліцензій показуються з бейджем.
        if not show_all_licenses:
            if not _has_base_license(s['id'], enrollment_licenses, base_license_id):
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
        _bump_study_form(study_form_totals, s['study_form'])

        # ── 3. Жінка без військових даних — не рахуємо ніде ──
        if _is_woman_without_military(
            s['gender'], s['military_id'], s['status_badge'],
            s['military_rank'], s['vos_code'], s['tck_name']
        ):
            continue

        status = s['status_badge'] or ''

        # ── 4. Для наявних записів military перевіряємо виключення до підрахунку проблем ──
        if s['military_id']:
            # Виключені статуси (Невійськовозобов'язаний / Знято з обліку)
            if status in EXCLUDED_STATUSES:
                continue

            # Позначка графи 18 (Постанова №812)
            if s['note_18'] == NOTE_EXCLUDED_BY_AGE:
                counts['excluded_by_age'] += 1
                continue
            if s['note_18'] == NOTE_DISMISSED:
                if _is_dismissed_this_year(s['dismissed_at'], report_year):
                    counts['dismissed_this_year'] += 1
                continue
            if s['note_18'] == NOTE_MOBILIZED:
                counts['mobilized'] += 1
                continue

        # ── 5. Проблеми ліцензії та наказу про зарахування (для всіх не-виключених) ──
        if not enrollment_licenses.get(s['id']) and not s['license_id']:
            issues['no_license'] += 1
        if s['id'] not in enrolled_ids:
            issues['no_enrollment_order'] += 1

        # ── 6. Без запису military (це вже тільки чоловіки, жінок відсіяли) ──
        if not s['military_id']:
            issues['no_military'] += 1
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
            _bump_study_form(study_form_by_cat[category], s['study_form'])
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
        study_form_totals=study_form_totals,
        study_form_by_cat=study_form_by_cat,
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
    study_form = (request.args.get('study_form') or '').strip()
    if study_form not in STUDY_FORM_OPTIONS:
        study_form = ''
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
               g.start_year AS group_start_year,
               g.study_form AS study_form,
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
    enrollment_dates = _load_enrollment_dates(conn)
    filtered = []
    enrollment_licenses = _load_enrollment_licenses(conn)
    enrollment_license_shorts = _load_enrollment_license_shorts(conn)   
    for r in rows_raw:
        # ── Фільтр базової ліцензії (Постанова №812) ──
        # Без ліцензії (NULL) — показуємо завжди (з бейджем).
        if not show_all_licenses:
            if not _has_base_license(r['id'], enrollment_licenses, base_license_id):
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

        # Вибулі за графою 18 (Постанова №812):
        #   - за віком і мобілізовані не входять у 4 групи Додатка 5
        #   - відраховані цього року показуються лише з чекбоксом include_dismissed
        note_18 = r['note_18']
        if note_18 in (NOTE_EXCLUDED_BY_AGE, NOTE_MOBILIZED):
            continue
        if note_18 == NOTE_DISMISSED:
            if not include_dismissed or not _is_dismissed_this_year(r['dismissed_at'], report_date.year):
                continue
            cat = _detect_category(
                r['military_rank'], status, r['gender'], r['category_override'], note_18=None
            )
        else:
            cat = _detect_category(
                r['military_rank'], status, r['gender'], r['category_override'], note_18
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
        if study_form and (r['study_form'] or '') != study_form:
            continue
        if tck_filter and _normalize_tck_key(r['tck_name']) != _normalize_tck_key(tck_filter):
            continue
        if not _was_enrolled_as_of(
            report_date,
            r['group_start_year'],
            enrollment_dates.get(r['id']),
        ):
            continue
        row = dict(r)
        row['enrollment_license_shorts'] = enrollment_license_shorts.get(r['id'], [])
        filtered.append(row)

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
            'license_short_names': r.get('enrollment_license_shorts') or [],
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
        study_form=study_form,
        study_form_options=STUDY_FORM_OPTIONS,
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
        report_year=report_date.year, 
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
        "SELECT * FROM military WHERE student_id = ?", (student_id,)
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

    diff_rows = []
    for k, new_val in updates.items():
        old_val = military[k] if k in military.keys() else None
        old_norm = (str(old_val).strip() if old_val is not None else '')
        new_norm = (str(new_val).strip() if new_val is not None else '')
        if old_norm != new_norm:
            diff_rows.append((k, old_norm, new_norm))

    set_clause = ', '.join(f"{k} = ?" for k in updates)
    try:
        conn.execute(
            f"UPDATE military SET {set_clause} WHERE student_id = ?",
            list(updates.values()) + [student_id],
        )
        if updates.get('note_18') == NOTE_DISMISSED:
            conn.execute(
                "UPDATE students SET dismissed_at = COALESCE(dismissed_at, date('now')) WHERE id = ?",
                (student_id,),
            )
        username = current_username()
        for field_name, old_val, new_val in diff_rows:
            conn.execute(
                """INSERT INTO military_history
                       (student_id, changed_by, source, scope, field_name, old_value, new_value)
                   VALUES (?, ?, 'quick_edit', 'military', ?, ?, ?)""",
                (student_id, username, field_name, old_val, new_val),
            )
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        logger.error(
            f"quick_edit military (student_id={student_id}): {e}",
            exc_info=True,
        )
        return jsonify({'ok': False, 'error': 'Помилка збереження військових даних'}), 500

    conn.close()
    _quick_labels = {
        'category': 'Категорія (звання)',
        'rank': 'Військове звання',
        'vos_code': 'ВОС',
        'tck_name': 'ТЦК та СП',
        'registration_type': 'Тип обліку',
        'note_18': 'Примітка (постанова)',
        'is_liable': 'Військовозобовʼязаний',
        'deferment_until': 'Відстрочка до',
        'conscription_status': 'Статус обліку',
        'data_verified_on_time': 'Дані звірені вчасно',
    }
    quick_changes = [
        {
            'field': fn,
            'label': _quick_labels.get(fn, fn),
            'old': ov or '—',
            'new': nv or '—',
        }
        for fn, ov, nv in diff_rows
    ]
    full_fio = f"{student['last_name_UA']} {student['first_name_UA']}".strip()
    log_action(
        current_username(),
        f"швидке редагування військових даних: {full_fio} (ID {student_id})",
        group_ids=[student['group_id']] if 'group_id' in student.keys() and student['group_id'] else None,
        category='update',
        section='Військовий облік',
        entity_type='student',
        entity_id=student_id,
        entity_name=full_fio,
        changes=quick_changes,
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
    'no_enrollment_order': 'Немає наказу про зарахування',
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
    show_all_licenses = request.args.get('show_all_licenses') == '1'
    search = (request.args.get('search') or '').strip()
    group_id = request.args.get('group_id', type=int)
    study_form = (request.args.get('study_form') or '').strip()
    if study_form not in STUDY_FORM_OPTIONS:
        study_form = ''
    page = max(1, request.args.get('page', 1, type=int))
    per_page = request.args.get('per_page', 50, type=int)
    if per_page not in (20, 50, 100, 200, 500):
        per_page = 50

    conn = get_db()
    conn.row_factory = sqlite3.Row

    base_license_id, base_license_short, base_license_name = _get_base_license_info(conn)
    enrollment_licenses = _load_enrollment_licenses(conn)

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
               g.study_form AS study_form,
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
    enrolled_ids = set()
    if orders_tables_ready(conn=conn):
        enrolled_ids = get_enrollment_student_ids(conn=conn)

    filtered = []
    issue_counts = {k: 0 for k in ISSUE_LABELS}

    for r in rows_raw:
        # Фільтр базової ліцензії (Постанова №812)
        if not show_all_licenses:
            if not _has_base_license(r['id'], enrollment_licenses, base_license_id):
                continue

        # Заморожені
        if r['frozen_id'] and not include_frozen:
            continue

        # Жінка без військових даних — не показуємо ніде
        if _is_woman_without_military(
            r['gender'], r['military_id'], r['status_badge'],
            r['military_rank'], r['vos_code'], r['tck_name']
        ):
            continue

        status = r['status_badge'] or ''

        # Виключені за статусом або графою 18 не входять у Додаток 5,
        # тому не є «проблемними» для Додатка 5
        if r['military_id']:
            if status in EXCLUDED_STATUSES:
                continue
            if r['note_18'] in NOTE_18_EXCLUDES_FROM_DODATOK_5:
                continue

        has_license = bool(enrollment_licenses.get(r['id']) or r['license_id'])
        has_enrollment = r['id'] in enrolled_ids

        # Підрахунок бейджів по всій вибірці (до пошуку/групи)
        if not has_license:
            issue_counts['no_license'] += 1
        if not has_enrollment:
            issue_counts['no_enrollment_order'] += 1

        if not r['military_id']:
            issue_counts['no_military'] += 1
        else:
            if not r['gender']:
                issue_counts['no_gender'] += 1
            cat = _detect_category(
                r['military_rank'], status, r['gender'], r['category_override'], r['note_18']
            )
            if not cat:
                issue_counts['no_category'] += 1
            if not r['tck_name']:
                issue_counts['no_tck'] += 1
            if not r['vos_code']:
                issue_counts['no_vos'] += 1

        problems = []

        # 0. Без ліцензії вступу (Постанова №812)
        if not has_license and 'no_license' in enabled:
            problems.append('no_license')

        if not has_enrollment and 'no_enrollment_order' in enabled:
            problems.append('no_enrollment_order')

        # 1. Без запису military
        if not r['military_id']:
            if 'no_military' in enabled:
                problems.append('no_military')
        else:
            # 2. Без статі
            if not r['gender'] and 'no_gender' in enabled:
                problems.append('no_gender')

            # 3. Категорія не визначена
            cat = _detect_category(
                r['military_rank'], status, r['gender'], r['category_override'], r['note_18']
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
        if study_form and (r['study_form'] or '') != study_form:
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
        study_form=study_form,
        study_form_options=STUDY_FORM_OPTIONS,
        include_frozen=include_frozen,
        show_all_licenses=show_all_licenses,
        base_license_label=base_license_short or base_license_name or '—',
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
    рахуються за календарним роком обраної дати звіту (report_year):
      - Зараховано: є активний наказ order_type='enrollment',
        рік order_date == report_year
        (без fallback на start_year групи)
      - Відраховано: накази order_type='expulsion' з order_date у report_year

    ⚠️ ВАЖЛИВО: цифри відображають ПОТОЧНИЙ стан БД. Історичних зрізів
    на довільну минулу дату система не має (бо не ведеться історія змін
    статусів). Це явно показується попередженням на сторінці.

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

    metrics = _compute_dodatok_12_metrics(
        conn, report_date, show_all_licenses, base_license_id
    )
    conn.close()

    counts = {
        'officers': metrics['officers'],
        'rank': metrics['rank'],
        'women': metrics['women'],
        'conscripts': metrics['conscripts'],
    }
    enrolled_this_year = metrics['enrolled']
    expelled_this_year = metrics['expelled']
    excluded_by_age = metrics['excluded_by_age']
    dismissed_this_year = metrics['dismissed_this_year']
    mobilized = metrics['mobilized']

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
                          base_license_id, base_license_label, tck_filter=None,
                          study_form=None):
    """
    Збирає рядки для Додатка 5 з повними полями (для генерації .xlsx).
    Логіка фільтрації — та сама, що й у group_list(), але БЕЗ пагінації
    і без include_frozen / include_dismissed (у звіт потрапляють лише
    активні на обліку).

    tck_filter: точна назва ТЦК (як у military.tck_name) або None/'' = усі.
    study_form: форма навчання ('Денна' / 'Заочна' / 'Дистанційна') або None/''.

    Повертає список dict.
    """
    report_year = report_date.year
    enrollment_dates = _load_enrollment_dates(conn)
    enrollment_years = _load_enrollment_years(conn)
    enrollment_licenses = _load_enrollment_licenses(conn)
    enrollment_license_shorts = _load_enrollment_license_shorts(conn)
    enrollment_or_renewal_orders = _load_enrollment_or_renewal_orders(conn, report_date=report_date)

    # Всі дані по студентах + military + passport + license
    students = conn.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.gender, s.tax_id, s.dismissed_at,
               s.license_id,
               s.address_of_residence,          
               il.short_name_ua AS license_short,
               m.id AS military_id, m.status_badge, m.military_rank, m.category_override,
               m.registry_number, m.vos_code, m.vos_description,
               m.military_registration_document,
               m.tck_name,
               m.deferral_type, m.deferral_until,
               m.note_18,
               m.change_credentials,
               m.reason_for_changing_credentials,
               g.start_year AS group_start_year,
               g.study_form AS study_form,
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
            s['gender'], s['military_id'], s['status_badge'],
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
            if not _has_base_license(s['id'], enrollment_licenses, base_license_id):
                continue

        # Категорія
        cat = _detect_category(
            s['military_rank'], status, s['gender'],
            s['category_override'], s['note_18']
        )
        if cat != category:
            continue

        # Фільтр ТЦК (якщо задано)
        if tck_filter and _normalize_tck_key(s['tck_name']) != _normalize_tck_key(tck_filter):
            continue
        if study_form and (s['study_form'] or '') != study_form:
            continue

        # Станом на report_date
        if not _was_enrolled_as_of(
            report_date,
            s['group_start_year'],
            enrollment_dates.get(s['id']),
        ):
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

        # Зараховано (колонка 23) — пріоритет наказу про зарахування, fallback на start_year групи
        enrolled_note = ''
        enroll_order_year = enrollment_years.get(s['id'])
        if enroll_order_year is not None:
            if enroll_order_year == report_year:
                enrolled_note = f"зараховано {enroll_order_year}"
        elif s['group_start_year'] == report_year:
            enrolled_note = f"зараховано {s['group_start_year']}"

        # Відраховано (колонка 24) — тільки якщо dismissed_at цього року
        expelled_note = ''
        if _is_dismissed_this_year(s['dismissed_at'], report_year):
            expelled_note = f"відраховано {s['dismissed_at']}"

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
            'enrollment_order': enrollment_or_renewal_orders.get(s['id'], ''),
            'note_18': s['note_18'] or '',
            'change_credentials': s['change_credentials'] or '',
            'reason_for_change': s['reason_for_changing_credentials'] or '',
            'license_short': ', '.join(enrollment_license_shorts.get(s['id'], [])) or (s['license_short'] or ''),
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
        tck=назва ТЦК  (якщо задано — лише студенти цього ТЦК; інакше всі)
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
    tck_filter = (request.args.get('tck') or '').strip() or None
    study_form = (request.args.get('study_form') or '').strip()
    if study_form not in STUDY_FORM_OPTIONS:
        study_form = None

    conn = get_db()
    conn.row_factory = sqlite3.Row

    base_license_id, base_license_short, base_license_name = _get_base_license_info(conn)

    rows = _fetch_dodatok_5_rows(
        conn, category, report_date, show_all_licenses,
        base_license_id, base_license_short or base_license_name,
        tck_filter=tck_filter,
        study_form=study_form,
    )
    conn.close()

    if not rows:
        flash(
            "Немає студентів для генерації — перевірте категорію, ліцензію"
            + (" та обраний ТЦК." if tck_filter else "."),
            "warning",
        )
        return redirect(url_for(
            'military_reports.group_list',
            category=category,
            report_mode=report_mode,
            report_date=report_date_custom if report_mode == 'custom' else None,
            show_all_licenses='1' if show_all_licenses else None,
            tck=tck_filter,
            study_form=study_form,
        ))

    xlsx_io = generate_dodatok_5_xlsx(
        category, report_date, rows,
        base_license_label=base_license_short or base_license_name,
    )

    # ── Логування в журнал ──
    rtype = f'dodatok_5_{category}'
    tck_slug = ''
    if tck_filter:
        # безпечне коротке ім'я для файлу
        tck_slug = '_' + ''.join(
            c if c.isalnum() or c in '-_' else '_'
            for c in tck_filter
        )[:40].strip('_')
    filename = (
        f"Додаток_5_{label.replace(' ', '_').replace(',', '')}_"
        f"{_format_date_ua(report_date).replace('.', '-')}"
        f"{tck_slug}.xlsx"
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

def _compute_dodatok_12_metrics(conn, report_date, show_all_licenses, base_license_id):
    """
    Єдиний розрахунок показників для Додатка 12 (спільний для веб-перегляду
    summary() та Excel-експорту _fetch_dodatok_12_data()).
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
    enrollment_licenses = _load_enrollment_licenses(conn)
    enrollment_years = _load_enrollment_years(conn)

    counts = {'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0}
    enrolled = {'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0}
    expelled = {'officers': 0, 'rank': 0, 'women': 0, 'conscripts': 0}
    excluded_by_age = 0
    dismissed_this_year = 0
    mobilized = 0

    for s in students:
        if not show_all_licenses:
            if not _has_base_license(s['id'], enrollment_licenses, base_license_id):
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
            continue

        category = _detect_category(
            s['military_rank'], status, s['gender'],
            s['category_override'], note_18
        )
        if not category:
            continue
        counts[category] += 1

        # «Зараховано з 1 січня» — за активним наказом про зарахування у report_year
        enroll_year = enrollment_years.get(s['id'])
        if enroll_year == report_year:
            enrolled[category] += 1

    # «Відраховано з 1 січня» — за активними наказами про відрахування в report_year
    if orders_tables_ready(conn=conn):
        expelled_rows = conn.execute("""
            SELECT oi.entity_id AS student_id,
                   o.order_date,
                   m.id AS military_id, m.status_badge, m.military_rank,
                   m.category_override, m.vos_code, m.tck_name,
                   s.gender, s.license_id
            FROM orders o
            JOIN order_items oi ON oi.order_id = o.id
            LEFT JOIN students s ON s.id = oi.entity_id
            LEFT JOIN military m ON m.student_id = oi.entity_id
            WHERE o.status = 'active'
              AND o.order_type = 'expulsion'
              AND oi.entity_type = 'student'
        """).fetchall()
        for e in expelled_rows:
            if _year_from_order_date(e['order_date']) != report_year:
                continue
            if not show_all_licenses and base_license_id is not None:
                lic_ids = enrollment_licenses.get(e['student_id'], set())
                if lic_ids:
                    if base_license_id not in lic_ids:
                        continue
                elif e['license_id'] is not None and e['license_id'] != base_license_id:
                    continue
            if not e['military_id']:
                continue
            status = e['status_badge'] or ''
            if status in EXCLUDED_STATUSES:
                continue
            if _is_woman_without_military(
                e['gender'], e['military_id'], e['status_badge'],
                e['military_rank'], e['vos_code'], e['tck_name']
            ):
                continue
            category = _detect_category(
                e['military_rank'], status, e['gender'], e['category_override']
            )
            if category:
                expelled[category] += 1

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


def _fetch_dodatok_12_data(conn, report_date, show_all_licenses, base_license_id):
    """
    Збирає підрахунки для Додатка 12.
    Логіка фільтрації — та сама, що й у summary().
    """
    return _compute_dodatok_12_metrics(conn, report_date, show_all_licenses, base_license_id)


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
@permission_required('manage_military_reports')
def tck_offices():
    """
    Довідник ТЦК.

    - is_admin або право manage_military_reports («Військовий облік (дашборд, додатки)»):
      повний CRUD і перегляд невідомих ТЦК.
    """
    conn = get_db()
    conn.row_factory = sqlite3.Row
    can_manage_tck = bool(session.get('is_admin')) or (
        'manage_military_reports' in (session.get('permissions') or [])
    )

    if request.method == 'POST':
        if not can_manage_tck:
            conn.close()
            flash('Недостатньо прав для редагування довідника ТЦК.', 'error')
            return redirect(url_for('military_reports.tck_offices'))
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
                    prev_row = conn.execute("SELECT name_ua, address FROM tck_offices WHERE id = ?", (tck_id,)).fetchone()
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
                    tck_changes = []
                    if prev_row:
                        if (prev_row['name_ua'] or '') != new_name:
                            tck_changes.append({'field': 'name_ua', 'label': 'Назва ТЦК', 'old': prev_row['name_ua'] or '—', 'new': new_name})
                        if (prev_row['address'] or '') != (new_address or ''):
                            tck_changes.append({'field': 'address', 'label': 'Адреса ТЦК', 'old': prev_row['address'] or '—', 'new': new_address or '—'})
                    details = (
                        f"перейменовано з «{old_name}»" if old_name != new_name else "оновлено"
                    )
                    if cascaded:
                        details += f", оновлено у {cascaded} студентів"
                    log_action(
                        current_username(),
                        f"редагував ТЦК (ID {tck_id}): {new_name}",
                        category='update',
                        section='Військовий облік',
                        entity_type='tck',
                        entity_id=tck_id,
                        entity_name=new_name,
                        changes=tck_changes,
                        details=details,
                    )
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
                log_action(
                    current_username(),
                    f"{'увімкнув' if new_state else 'вимкнув'} ТЦК: {row['name_ua']}",
                    category='update',
                    section='Військовий облік',
                    entity_type='tck',
                    entity_id=tck_id,
                    entity_name=row['name_ua'],
                    changes=[{
                        'field': 'is_active',
                        'label': 'Статус ТЦК',
                        'old': 'Активний' if row['is_active'] else 'Неактивний',
                        'new': 'Активний' if new_state else 'Неактивний',
                    }],
                )
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

    # Без права повного доступу: лише ТЦК, де є студенти на обліку
    if not can_manage_tck:
        rows = [r for r in rows if r['students_count_all'] > 0]

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

    # Без права повного доступу не показується блок «невідомі» і масові операції
    if not can_manage_tck:
        unknown_tck = []
        stats = {
            **stats,
            'unknown_count': 0,
            'unknown_students': 0,
        }

    return render_template(
        'military_reports/tck.html',
        rows=rows,
        unknown_tck=unknown_tck,
        search=search,
        show=show,
        stats=stats,
        can_manage_tck=can_manage_tck,
    )
    
@military_reports_bp.route('/admin/military_reports/tck/<int:tck_id>/students')
@permission_required('manage_military_reports')
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
               m.note_18,
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
    if not include_dismissed:
        extra_conditions.append("m.note_18 IS NULL")
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
            SUM(CASE WHEN COALESCE(s.archived, 0) = 0 AND fz.id IS NULL AND m.note_18 IS NULL
                     THEN 1 ELSE 0 END) AS total_active,
            SUM(CASE WHEN COALESCE(s.archived, 0) = 1
                     THEN 1 ELSE 0 END) AS total_archived,
            SUM(CASE WHEN COALESCE(s.archived, 0) = 0 AND fz.id IS NOT NULL
                     THEN 1 ELSE 0 END) AS total_frozen,
            SUM(CASE WHEN COALESCE(s.archived, 0) = 0 AND fz.id IS NULL AND m.note_18 IS NOT NULL
                     THEN 1 ELSE 0 END) AS total_note18
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
            'total_note18':   counts_row['total_note18'] or 0,
        },
    })

# ── Додаток 4 ──────────────────────────────────────────────
from routes.dodatok4_export import generate_dodatok_4_xlsx  # noqa: E402


# ── Додаток 4: відбиток даних і позначка «надіслано» ───────────

_DODATOK4_FINGERPRINT_FIELDS = (
    'last_name_UA', 'first_name_UA', 'middle_name_UA', 'birth_date', 'tax_id',
    'registry_number', 'military_registration_document', 'vos_code',
    'military_rank', 'tck_name', 'change_credentials',
    'reason_for_changing_credentials', 'address_of_residence',
)


_DODATOK4_FIELD_LABELS = {
    'last_name_UA': 'Прізвище',
    'first_name_UA': "Ім'я",
    'middle_name_UA': 'По батькові',
    'birth_date': 'Дата народження',
    'tax_id': 'РНКОПП',
    'registry_number': 'Реєстр. номер',
    'military_registration_document': 'ВОД',
    'vos_code': 'ВОС',
    'military_rank': 'Звання',
    'tck_name': 'ТЦК',
    'change_credentials': 'Зміна облікових даних',
    'reason_for_changing_credentials': 'Підстава зміни',
    'address_of_residence': 'Адреса',
}


def _dodatok4_values(row) -> dict:
    """Поточні значення полів Додатка 4."""
    out = {}
    for key in _DODATOK4_FINGERPRINT_FIELDS:
        try:
            val = row[key] if not isinstance(row, dict) else row.get(key)
        except (KeyError, IndexError, TypeError):
            val = None
        if isinstance(val, str):
            out[key] = val.strip()
        else:
            out[key] = str(val or '')
    return out


def _dodatok4_snapshot_json(row) -> str:
    import json
    return json.dumps(_dodatok4_values(row), ensure_ascii=False)


def _dodatok4_diff(row) -> list:
    """
    Список змін {(label, old, new)} порівняно зі збереженим snapshot.
    Порожній список, якщо snapshot немає (перший друк) або збіг.
    """
    import json
    try:
        raw = row['dodatok4_snapshot'] if not isinstance(row, dict) else row.get('dodatok4_snapshot')
    except (KeyError, IndexError, TypeError):
        raw = None
    if not raw:
        return []
    try:
        old = json.loads(raw)
    except Exception:
        return []
    if not isinstance(old, dict):
        return []
    cur = _dodatok4_values(row)
    changes = []
    for key in _DODATOK4_FINGERPRINT_FIELDS:
        a = (old.get(key) or '').strip()
        b = (cur.get(key) or '').strip()
        if a != b:
            changes.append({
                'field': key,
                'label': _DODATOK4_FIELD_LABELS.get(key, key),
                'old': a or '—',
                'new': b or '—',
            })
    return changes



def _ensure_dodatok4_columns(conn):
    """
    Колонки обліку Додатка 4 створюються у init_db.py.
    Тут лише перевіряємо їхню наявність і попереджаємо, якщо БД стара —
    НЕ робимо ALTER TABLE на льоту (це створює ризик блокувань і
    розсинхрону зі схемою).
    """
    existing = {row[1] for row in conn.execute('PRAGMA table_info(military)')}
    required = {'dodatok4_sent_at', 'dodatok4_fingerprint', 'dodatok4_snapshot'}
    missing = required - existing
    if missing:
        logger.error(
            f"military не має колонок {sorted(missing)} — функціонал «Додаток 4» "
            f"непрацездатний. Запустіть `python init_db.py` для оновлення схеми."
        )


def _dodatok4_fingerprint(row) -> str:
    """Хеш полів, що потрапляють у Додаток 4."""
    parts = []
    for key in _DODATOK4_FINGERPRINT_FIELDS:
        try:
            if isinstance(row, dict):
                val = row.get(key)
            else:
                val = row[key]
        except (KeyError, IndexError, TypeError):
            val = None
        if isinstance(val, str):
            parts.append(val.strip())
        else:
            parts.append(str(val or ''))
    raw = '|'.join(parts)
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()


def _dodatok4_needs_resend(row) -> bool:
    """True, якщо ще не друкували або дані змінились після відправки."""
    fp = _dodatok4_fingerprint(row)
    try:
        prev = row['dodatok4_fingerprint'] if not isinstance(row, dict) else row.get('dodatok4_fingerprint')
    except (KeyError, IndexError, TypeError):
        prev = None
    if not prev:
        return True
    return prev != fp


@military_reports_bp.route('/admin/military_reports/dodatok_4')
@permission_required('manage_military_reports')
def dodatok_4_list():
    """
    Список для Додатка 4.
    За замовчуванням — потребують друку (ніколи не слали або дані змінились).
    """
    search = (request.args.get('search') or '').strip()
    tck_filter = (request.args.get('tck') or '').strip()
    study_form = (request.args.get('study_form') or '').strip()
    if study_form not in STUDY_FORM_OPTIONS:
        study_form = ''
    status = (request.args.get('status') or 'need').strip()
    if status not in ('need', 'sent', 'all'):
        status = 'need'

    conn = get_db()
    conn.row_factory = sqlite3.Row
    _ensure_dodatok4_columns(conn)

    tck_offices = conn.execute("""
        SELECT name_ua FROM tck_offices WHERE is_active = 1
        ORDER BY name_ua COLLATE UKRAINIAN
    """).fetchall()

    rows = conn.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.tax_id,
               m.registry_number, m.military_registration_document,
               m.vos_code, m.military_rank, m.tck_name,
               m.change_credentials, m.reason_for_changing_credentials,
               s.address_of_residence,
               m.dodatok4_sent_at, m.dodatok4_fingerprint, m.dodatok4_snapshot,
               g.name AS group_name, g.study_form AS study_form
        FROM students s
        INNER JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON g.id = s.group_id
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA COLLATE UKRAINIAN,
                 s.first_name_UA COLLATE UKRAINIAN
    """).fetchall()

    filtered = []
    for r in rows:
        needs = _dodatok4_needs_resend(r)
        if status == 'need' and not needs:
            continue
        if status == 'sent' and needs:
            continue
        if tck_filter and _normalize_tck_key(r['tck_name']) != _normalize_tck_key(tck_filter):
            continue
        if study_form and (r['study_form'] or '') != study_form:
            continue
        if search:
            hay = (
                f"{r['last_name_UA']} {r['first_name_UA']} "
                f"{r['middle_name_UA'] or ''}"
            ).lower()
            if not all(w in hay for w in search.lower().split()):
                continue
        full = ' '.join(
            x for x in (
                r['last_name_UA'], r['first_name_UA'], r['middle_name_UA'] or ''
            ) if x
        )
        changes = _dodatok4_diff(r) if needs else []
        has_fp = bool(r['dodatok4_fingerprint'] or '')
        has_snap = bool(r['dodatok4_snapshot'] or '') if 'dodatok4_snapshot' in r.keys() else False
        filtered.append({
            **dict(r),
            'full_name': full,
            'needs_resend': needs,
            'changes': changes,
            'is_first_print': needs and not has_fp,
            'missing_snapshot': needs and has_fp and not has_snap and not changes,
        })

    conn.close()
    return render_template(
        'military_reports/dodatok_4.html',
        rows=filtered,
        search=search,
        tck_filter=tck_filter,
        study_form=study_form,
        study_form_options=STUDY_FORM_OPTIONS,
        status=status,
        tck_offices=tck_offices,
    )


@military_reports_bp.route('/admin/military_reports/dodatok_4/export/<int:student_id>')
@permission_required('manage_military_reports')
def dodatok_4_export(student_id):
    """Один файл Додатка 4; після успіху — позначка «дані актуальні»."""
    from flask import send_file
    from routes.dodatok4_export import generate_dodatok_4_xlsx
    from datetime import datetime

    conn = get_db()
    conn.row_factory = sqlite3.Row
    _ensure_dodatok4_columns(conn)

    r = conn.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.tax_id, s.address_of_residence,  
               m.registry_number, m.military_registration_document,
               m.vos_code, m.military_rank, m.tck_name,
               m.change_credentials, m.reason_for_changing_credentials,
               m.dodatok4_sent_at, m.dodatok4_fingerprint
        FROM students s
        INNER JOIN military m ON m.student_id = s.id
        WHERE s.id = ?
    """, (student_id,)).fetchone()

    if not r:
        conn.close()
        flash('Студента або військових даних не знайдено.', 'error')
        return redirect(url_for('military_reports.dodatok_4_list'))

    full = ' '.join(
        x for x in (r['last_name_UA'], r['first_name_UA'], r['middle_name_UA'] or '') if x
    )
    data = {
        'tck_header': r['tck_name'] or '',
        'full_name': full,
        'registry_number': r['registry_number'] or '',
        'military_document': r['military_registration_document'] or '',
        'vos_code': r['vos_code'] or '',
        'military_rank': r['military_rank'] or '',
        'birth_date': r['birth_date'] or '',
        'change_text': r['change_credentials'] or '',
        'basis_text': r['reason_for_changing_credentials'] or '',
        'tck_name': r['tck_name'] or '',
        'address': r['address_of_residence'] or '',
        'tax_id': r['tax_id'] or '',
    }

    try:
        buf = generate_dodatok_4_xlsx(data)
    except FileNotFoundError as e:
        conn.close()
        flash(str(e), 'error')
        return redirect(url_for('military_reports.dodatok_4_list'))

    fp = _dodatok4_fingerprint(r)
    snap = _dodatok4_snapshot_json(r)
    conn.execute(
        """
        UPDATE military
        SET dodatok4_sent_at = datetime('now', 'localtime'),
            dodatok4_fingerprint = ?,
            dodatok4_snapshot = ?
        WHERE student_id = ?
        """,
        (fp, snap, student_id),
    )
    conn.commit()
    conn.close()

    safe = ''.join(c if c.isalnum() or c in '-_' else '_' for c in full)[:40]
    filename = f"Dodatok_4_{safe}_{student_id}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


# ═══════════════════════════════════════════════════════════════════════
# Військово-облікові файли «Резерв+» (помісячні паки та контроль оновлень)
# ═══════════════════════════════════════════════════════════════════════

_MONTH_NAMES_UA = {
    1: 'Січень',
    2: 'Лютий',
    3: 'Березень',
    4: 'Квітень',
    5: 'Травень',
    6: 'Червень',
    7: 'Липень',
    8: 'Серпень',
    9: 'Вересень',
    10: 'Жовтень',
    11: 'Листопад',
    12: 'Грудень',
}


def _format_ym_label_ua(ym: str) -> str:
    """'2026-10' → 'Жовтень 2026'."""
    if not ym or '-' not in ym:
        return ym or '—'
    try:
        y_s, m_s = ym.split('-', 1)
        m_name = _MONTH_NAMES_UA.get(int(m_s), m_s)
        return f"{m_name} {int(y_s)}"
    except (ValueError, TypeError):
        return ym


def _parse_ym_and_sort_key(raw_val: str):
    """
    Розбирає дату/час формату 'DD.MM.YYYY HH:MM', 'DD.MM.YYYY',
    'DD/MM/YYYY...', 'YYYY-MM-DD HH:MM:SS' або 'YYYY-MM-DD'.
    Повертає ('YYYY-MM', 'YYYY-MM-DD HH:MM:SS') або (None, '').
    """
    if not raw_val:
        return None, ''
    s = str(raw_val).strip()
    if not s:
        return None, ''

    # DD.MM.YYYY [HH:MM[:SS]] або DD/MM/YYYY
    m = re.match(r'^(\d{1,2})[./](\d{1,2})[./](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?', s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        hh = int(m.group(4) or 0)
        mm = int(m.group(5) or 0)
        ss = int(m.group(6) or 0)
        if 1 <= mo <= 12 and 1 <= d <= 31:
            return f"{y:04d}-{mo:02d}", f"{y:04d}-{mo:02d}-{d:02d} {hh:02d}:{mm:02d}:{ss:02d}"

    # YYYY-MM-DD [HH:MM[:SS]]
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?', s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        hh = int(m.group(4) or 0)
        mm = int(m.group(5) or 0)
        ss = int(m.group(6) or 0)
        if 1 <= mo <= 12 and 1 <= d <= 31:
            return f"{y:04d}-{mo:02d}", f"{y:04d}-{mo:02d}-{d:02d} {hh:02d}:{mm:02d}:{ss:02d}"

    return None, ''


def _sanitize_zip_filename(name: str) -> str:
    """Очищає ПІБ для безпечної назви файлу в архіві."""
    cleaned = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', '_', (name or '').strip())
    cleaned = re.sub(r'\s+', ' ', cleaned).strip(' .')
    return cleaned


def _resolve_attachment_disk_path(rel_path: str):
    """Повертає абсолютний шлях у static/ або None при некоректному шляху."""
    if not rel_path:
        return None
    static_root = os.path.abspath(os.path.join(PROJECT_ROOT, 'static'))
    clean_rel = str(rel_path).replace('\\', '/').lstrip('/')
    if clean_rel.startswith('static/'):
        clean_rel = clean_rel[len('static/'):]
    candidate = os.path.abspath(os.path.join(static_root, clean_rel))
    if candidate != static_root and not candidate.startswith(static_root + os.sep):
        return None
    return candidate


def _recent_months_list(count: int = 12):
    """Останні count календарних місяців ['YYYY-MM', ...] включно з поточним."""
    today = date.today()
    y, m = today.year, today.month
    out = []
    for _ in range(count):
        out.append(f"{y:04d}-{m:02d}")
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return out


_FORMED_AT_IN_PDF_RE = re.compile(
    r'Сформовано\s*[:\-]?\s*(?:(\d{1,2}[./]\d{1,2}[./]\d{4})|(\d{4}-\d{1,2}-\d{1,2}))'
    r'(?:[\s,]+(\d{1,2}:\d{2}(?::\d{2})?))?',
    re.IGNORECASE,
)


def _extract_formed_at_from_text(text: str) -> str:
    """Шукає «Сформовано: ДД.ММ.РРРР[, ГГ:ХХ]» у тексті PDF «Резерв+»."""
    if not text:
        return ''
    m = _FORMED_AT_IN_PDF_RE.search(text)
    if not m:
        return ''
    date_part = m.group(1) or m.group(2) or ''
    time_part = m.group(3) or ''
    return f"{date_part} {time_part}".strip()


def _extract_formed_at_from_pdf(disk_path: str):
    """
    Читає дату «Сформовано:» безпосередньо з PDF-файлу на диску.
    Повертає (formed_at_str, parsed_ok):
      - parsed_ok=True: PDF успішно прочитано (навіть якщо рядок не знайдено)
      - parsed_ok=False: файл відсутній або не вдалося відкрити/розпізнати
    """
    if not disk_path or not os.path.exists(disk_path):
        return '', False

    text_chunks = []
    reader_available = False

    # 1. PyMuPDF (fitz) — найшвидший для масового сканування
    try:
        import fitz  # PyMuPDF
        reader_available = True
        with fitz.open(disk_path) as doc:
            for page in doc[:2]:
                txt = page.get_text() or ''
                if txt:
                    text_chunks.append(txt)
                    found = _extract_formed_at_from_text(txt)
                    if found:
                        return found, True
    except ImportError:
        pass
    except Exception:
        pass

    # 2. pdfplumber — штатний парсер проєкту (routes/parse_military_pdf.py)
    try:
        import pdfplumber
        reader_available = True
        with pdfplumber.open(disk_path) as pdf:
            for page in pdf.pages[:2]:
                txt = page.extract_text() or ''
                if txt:
                    text_chunks.append(txt)
                    found = _extract_formed_at_from_text(txt)
                    if found:
                        return found, True
    except ImportError:
        pass
    except Exception:
        pass

    # 3. pypdf / PyPDF2 fallback
    for mod_name in ('pypdf', 'PyPDF2'):
        try:
            mod = __import__(mod_name)
            reader_available = True
            reader = mod.PdfReader(disk_path)
            for page in reader.pages[:2]:
                txt = page.extract_text() or ''
                if txt:
                    text_chunks.append(txt)
                    found = _extract_formed_at_from_text(txt)
                    if found:
                        return found, True
            break
        except ImportError:
            continue
        except Exception:
            continue

    if text_chunks:
        return _extract_formed_at_from_text('\n'.join(text_chunks)), True

    return '', reader_available


def _ensure_attachment_meta_table(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS attachment_military_meta (
            attachment_id INTEGER PRIMARY KEY,
            file_path TEXT NOT NULL,
            file_size INTEGER NOT NULL,
            file_mtime REAL NOT NULL,
            formed_at TEXT NOT NULL DEFAULT ''
        )
    """)


def _load_attachment_formed_map(conn, att_rows):
    """
    Повертає словник {attachment_id: formed_at_from_pdf_str} для всіх
    військово-облікових PDF-файлів, використовуючи кеш у SQLite та автоматично
    перечитуючи PDF при зміні розміру/mtime/шляху файлу на диску.
    """
    _ensure_attachment_meta_table(conn)
    cached_rows = conn.execute(
        "SELECT attachment_id, file_path, file_size, file_mtime, formed_at "
        "FROM attachment_military_meta"
    ).fetchall()
    cache = {
        r['attachment_id']: (
            r['file_path'], r['file_size'], r['file_mtime'], r['formed_at']
        )
        for r in cached_rows
    }

    result = {}
    updates = []
    for a in att_rows:
        rel_path = a.get('file_path') or ''
        if not rel_path.lower().endswith('.pdf'):
            continue
        disk_path = _resolve_attachment_disk_path(rel_path)
        if not disk_path or not os.path.exists(disk_path):
            continue
        try:
            st = os.stat(disk_path)
        except OSError:
            continue

        aid = a['id']
        cached = cache.get(aid)
        if (
            cached
            and cached[0] == rel_path
            and cached[1] == st.st_size
            and abs(cached[2] - st.st_mtime) < 1e-6
        ):
            result[aid] = cached[3]
            continue

        formed_str, parsed_ok = _extract_formed_at_from_pdf(disk_path)
        result[aid] = formed_str
        if parsed_ok:
            updates.append((aid, rel_path, st.st_size, st.st_mtime, formed_str))

    if updates:
        conn.executemany(
            "INSERT OR REPLACE INTO attachment_military_meta "
            "(attachment_id, file_path, file_size, file_mtime, formed_at) "
            "VALUES (?, ?, ?, ?, ?)",
            updates,
        )
        conn.commit()

    return result


def _collect_military_files_overview(
    conn,
    *,
    selected_ym: str,
    search: str = '',
    group_id=None,
    study_form: str = '',
    category_filter: str = '',
    show_all_licenses: bool = False,
    include_frozen: bool = False,
    pdf_only: bool = False,
):
    """
    Збирає дані про військово-облікові файли («Резерв+») для сторінки контролю
    та завантаження помісячних ZIP-паків.

    Уся помісячна прив'язка будується СТРОГО за датою «Сформовано:», зчитаною
    безпосередньо з самого PDF-файлу (без підміни датою завантаження uploaded_at).
    """
    base_license_id, base_license_short, base_license_name = _get_base_license_info(conn)
    enrollment_licenses = _load_enrollment_licenses(conn)

    students = conn.execute("""
        SELECT s.id, s.last_name_UA, s.first_name_UA, s.middle_name_UA,
               s.birth_date, s.gender, s.license_id,
               s.phone, s.phone_backup, s.email, s.address_of_residence,
               m.id AS military_id, m.status_badge, m.military_rank,
               m.category_override, m.tck_name, m.vos_code, m.note_18,
               m.formed_at, m.military_registration_document,
               g.id AS group_id, g.name AS group_name,
               g.study_form AS study_form,
               fz.id AS frozen_id, fz.reason AS frozen_reason
        FROM students s
        LEFT JOIN military m ON m.student_id = s.id
        LEFT JOIN groups g ON s.group_id = g.id
        LEFT JOIN frozen_students fz
               ON fz.student_id = s.id AND fz.resolved_at IS NULL
        WHERE COALESCE(s.archived, 0) = 0
        ORDER BY s.last_name_UA COLLATE UKRAINIAN,
                 s.first_name_UA COLLATE UKRAINIAN
    """).fetchall()

    att_rows_raw = conn.execute("""
        SELECT id, entity_id, file_path, original_name, uploaded_at, uploaded_by
        FROM attachments
        WHERE entity_type = 'military'
        ORDER BY id ASC
    """).fetchall()
    att_rows = [dict(a) for a in att_rows_raw]
    pdf_formed_map = _load_attachment_formed_map(conn, att_rows)

    atts_by_mil = {}
    for a in att_rows:
        atts_by_mil.setdefault(a['entity_id'], []).append(a)

    ready_rows = []
    missing_rows = []
    outdated_rows = []
    month_student_counts = {}

    for s in students:
        if not show_all_licenses:
            if not _has_base_license(s['id'], enrollment_licenses, base_license_id):
                continue

        if s['frozen_id'] and not include_frozen:
            continue

        if _is_woman_without_military(
            s['gender'], s['military_id'], s['status_badge'],
            s['military_rank'], s['vos_code'], s['tck_name']
        ):
            continue

        status = s['status_badge'] or ''
        if status in EXCLUDED_STATUSES:
            continue
        if s['note_18'] in NOTE_18_EXCLUDES_FROM_DODATOK_5:
            continue

        cat = _detect_category(
            s['military_rank'], status, s['gender'],
            s['category_override'], s['note_18']
        ) if s['military_id'] else None

        if group_id and s['group_id'] != group_id:
            continue
        if study_form and (s['study_form'] or '') != study_form:
            continue
        if category_filter:
            if category_filter == 'other':
                if cat is not None:
                    continue
            elif cat != category_filter:
                continue

        full_name = ' '.join(
            x for x in (s['last_name_UA'], s['first_name_UA'], s['middle_name_UA'] or '') if x
        ).strip()

        if search:
            haystack = ' '.join([
                full_name,
                s['phone'] or '',
                s['phone_backup'] or '',
                s['email'] or '',
                s['group_name'] or '',
            ]).lower()
            if not all(w in haystack for w in search.lower().split()):
                continue

        card_formed_ym, _ = _parse_ym_and_sort_key(s['formed_at'])

        raw_atts = atts_by_mil.get(s['military_id'], []) if s['military_id'] else []
        if pdf_only:
            filtered_atts = [
                a for a in raw_atts
                if (a.get('file_path') or '').lower().endswith('.pdf')
            ]
        else:
            filtered_atts = list(raw_atts)

        processed_atts = []
        files_by_month = {}
        for a in filtered_atts:
            disk_path = _resolve_attachment_disk_path(a.get('file_path') or '')
            exists_on_disk = bool(disk_path and os.path.exists(disk_path))
            is_pdf = (a.get('file_path') or '').lower().endswith('.pdf')

            # Місяць файлу визначається ВИКЛЮЧНО за «Сформовано:» всередині PDF
            raw_pdf_formed = pdf_formed_map.get(a['id'], '') if (is_pdf and exists_on_disk) else ''
            pdf_ym, pdf_sort = _parse_ym_and_sort_key(raw_pdf_formed)
            _, up_sort = _parse_ym_and_sort_key(a.get('uploaded_at'))

            if pdf_ym:
                eff_ym = pdf_ym
                eff_sort = pdf_sort or up_sort
                date_display = raw_pdf_formed
                date_source = 'formed'
            else:
                eff_ym = None
                eff_sort = up_sort
                date_display = ''
                date_source = 'no_formed'

            att_info = {
                'id': a['id'],
                'file_path': a.get('file_path') or '',
                'disk_path': disk_path,
                'exists_on_disk': exists_on_disk,
                'original_name': a.get('original_name') or os.path.basename(a.get('file_path') or 'file'),
                'uploaded_at': a.get('uploaded_at') or '',
                'effective_ym': eff_ym,
                'effective_sort': eff_sort,
                'date_display': date_display,
                'date_source': date_source,
                'is_pdf': is_pdf,
            }
            processed_atts.append(att_info)
            if eff_ym and exists_on_disk:
                files_by_month.setdefault(eff_ym, []).append(att_info)

        # У кожному місяці сортуємо файли від найновішого за «Сформовано:»
        for ym_key, m_files in files_by_month.items():
            m_files.sort(key=lambda x: (x['effective_sort'], x['id']), reverse=True)
            if m_files:
                month_student_counts[ym_key] = month_student_counts.get(ym_key, 0) + 1

        existing_atts = [a for a in processed_atts if a['exists_on_disk']]
        in_month_files = files_by_month.get(selected_ym, [])

        # Найновіший файл: пріоритет файлам із розпізнаною датою «Сформовано:»
        formed_atts = [a for a in existing_atts if a['effective_ym']]
        latest_att = max(
            formed_atts or existing_atts,
            key=lambda a: (a['effective_sort'], a['id']),
            default=None,
        )

        row_data = {
            'id': s['id'],
            'full_name': full_name,
            'group_id': s['group_id'],
            'group_name': s['group_name'] or '',
            'study_form': s['study_form'] or '',
            'category': cat,
            'category_label': CATEGORY_LABELS[cat][0] if cat in CATEGORY_LABELS else 'Не визначена',
            'has_military': bool(s['military_id']),
            'frozen_id': s['frozen_id'],
            'frozen_reason': s['frozen_reason'] or '',
            'phone': (s['phone'] or '').strip(),
            'phone_backup': (s['phone_backup'] or '').strip(),
            'email': (s['email'] or '').strip(),
            'address': (s['address_of_residence'] or '').strip(),
            'formed_at': (s['formed_at'] or '').strip(),
            'formed_ym': card_formed_ym,
            'in_month_files': in_month_files,
            'all_files_count': len(existing_atts),
            'latest_file': latest_att,
            'latest_ym': latest_att['effective_ym'] if latest_att else None,
            'latest_ym_label': (
                _format_ym_label_ua(latest_att['effective_ym'])
                if latest_att and latest_att['effective_ym']
                else ''
            ),
        }

        if in_month_files:
            row_data['file_status'] = 'ready'
            ready_rows.append(row_data)
        elif not existing_atts:
            row_data['file_status'] = 'missing'
            missing_rows.append(row_data)
        else:
            row_data['file_status'] = 'outdated'
            outdated_rows.append(row_data)

    # Доступні місяці: останні 12 місяців + всі місяці з наявними файлами
    all_yms = sorted(
        set(_recent_months_list(12)) | set(month_student_counts.keys()) | ({selected_ym} if selected_ym else set()),
        reverse=True,
    )
    month_options = [
        {
            'ym': ym,
            'label': _format_ym_label_ua(ym),
            'count': month_student_counts.get(ym, 0),
        }
        for ym in all_yms
    ]

    need_contact_rows = outdated_rows + missing_rows
    need_contact_rows.sort(key=lambda r: r['full_name'].lower())

    return {
        'ready_rows': ready_rows,
        'missing_rows': missing_rows,
        'outdated_rows': outdated_rows,
        'need_contact_rows': need_contact_rows,
        'month_options': month_options,
        'month_student_counts': month_student_counts,
        'base_license_label': base_license_short or base_license_name or '—',
    }


@military_reports_bp.route('/admin/military_reports/files')
@permission_required('manage_military_reports')
def military_files():
    """
    Сторінка контролю військово-облікових файлів («Резерв+»):
      - Завантаження помісячного ZIP-паку з файлами, названими ПІБ студентів
      - Списки студентів без файлу та тих, хто ще не оновив файл у вибраному місяці
      - Контакти для зв'язку зі студентами (телефон, резервний телефон, email)
    """
    today_ym = date.today().strftime('%Y-%m')
    selected_ym = (request.args.get('month') or today_ym).strip()
    if not re.match(r'^\d{4}-\d{2}$', selected_ym):
        selected_ym = today_ym

    search = (request.args.get('search') or '').strip()
    group_id = request.args.get('group_id', type=int)
    study_form = (request.args.get('study_form') or '').strip()
    if study_form not in STUDY_FORM_OPTIONS:
        study_form = ''
    category_filter = (request.args.get('category') or '').strip()
    if category_filter not in ('officers', 'rank', 'women', 'conscripts', 'other'):
        category_filter = ''
    show_all_licenses = request.args.get('show_all_licenses') == '1'
    include_frozen = request.args.get('include_frozen') == '1'
    pdf_only = request.args.get('pdf_only') == '1'
    active_tab = (request.args.get('tab') or 'need').strip()
    if active_tab not in ('need', 'missing', 'outdated', 'ready'):
        active_tab = 'need'

    conn = get_db()
    conn.row_factory = sqlite3.Row

    groups = conn.execute("""
        SELECT id, name, start_year, study_form
        FROM groups WHERE archived = FALSE
        ORDER BY name COLLATE UKRAINIAN
    """).fetchall()

    overview = _collect_military_files_overview(
        conn,
        selected_ym=selected_ym,
        search=search,
        group_id=group_id,
        study_form=study_form,
        category_filter=category_filter,
        show_all_licenses=show_all_licenses,
        include_frozen=include_frozen,
        pdf_only=pdf_only,
    )
    conn.close()

    return render_template(
        'military_reports/files.html',
        selected_ym=selected_ym,
        selected_ym_label=_format_ym_label_ua(selected_ym),
        today_ym=today_ym,
        active_tab=active_tab,
        search=search,
        group_id=group_id,
        study_form=study_form,
        study_form_options=STUDY_FORM_OPTIONS,
        category_filter=category_filter,
        category_labels=CATEGORY_LABELS,
        show_all_licenses=show_all_licenses,
        include_frozen=include_frozen,
        pdf_only=pdf_only,
        groups=groups,
        **overview,
    )


@military_reports_bp.route('/admin/military_reports/files/download_pack')
@permission_required('manage_military_reports')
def military_files_download_pack():
    """
    Формує і віддає ZIP-архів файлів «Резерв+» за вибраний місяць.
    Кожен файл усередині архіву називається ПІБ студента (напр. «Іваненко Іван Іванович.pdf»).
    """
    today_ym = date.today().strftime('%Y-%m')
    selected_ym = (request.args.get('month') or today_ym).strip()
    if not re.match(r'^\d{4}-\d{2}$', selected_ym):
        selected_ym = today_ym

    search = (request.args.get('search') or '').strip()
    group_id = request.args.get('group_id', type=int)
    study_form = (request.args.get('study_form') or '').strip()
    if study_form not in STUDY_FORM_OPTIONS:
        study_form = ''
    category_filter = (request.args.get('category') or '').strip()
    if category_filter not in ('officers', 'rank', 'women', 'conscripts', 'other'):
        category_filter = ''
    show_all_licenses = request.args.get('show_all_licenses') == '1'
    include_frozen = request.args.get('include_frozen') == '1'
    pdf_only = request.args.get('pdf_only') == '1'

    conn = get_db()
    conn.row_factory = sqlite3.Row
    overview = _collect_military_files_overview(
        conn,
        selected_ym=selected_ym,
        search=search,
        group_id=group_id,
        study_form=study_form,
        category_filter=category_filter,
        show_all_licenses=show_all_licenses,
        include_frozen=include_frozen,
        pdf_only=pdf_only,
    )
    conn.close()

    ready_rows = overview['ready_rows']
    if not ready_rows:
        flash(
            f"За {_format_ym_label_ua(selected_ym)} немає військово-облікових файлів для завантаження.",
            "warning",
        )
        return redirect(url_for(
            'military_reports.military_files',
            month=selected_ym,
            search=search or None,
            group_id=group_id,
            study_form=study_form or None,
            category=category_filter or None,
            show_all_licenses='1' if show_all_licenses else None,
            include_frozen='1' if include_frozen else None,
            pdf_only='1' if pdf_only else None,
        ))

    buf = io.BytesIO()
    used_arcnames = set()
    files_written = 0
    # Рахуємо повтори ПІБ серед студентів, щоб тезки не перезаписали один одного
    name_counts = {}
    for r in ready_rows:
        base_name = _sanitize_zip_filename(r['full_name']) or f"Студент_{r['id']}"
        name_counts[base_name] = name_counts.get(base_name, 0) + 1

    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for r in ready_rows:
            base_name = _sanitize_zip_filename(r['full_name']) or f"Студент_{r['id']}"
            if name_counts.get(base_name, 0) > 1:
                base_name = f"{base_name} (ID {r['id']})"

            student_files = [a for a in r['in_month_files'] if a['exists_on_disk']][:1]
            for idx, att in enumerate(student_files, start=1):
                ext = os.path.splitext(att['file_path'])[1].lower() or '.pdf'
                arcname = f"{base_name}{ext}" if idx == 1 else f"{base_name} ({idx}){ext}"
                suffix_n = 2
                while arcname.lower() in used_arcnames:
                    arcname = f"{base_name} ({suffix_n}){ext}"
                    suffix_n += 1
                used_arcnames.add(arcname.lower())
                zf.write(att['disk_path'], arcname=arcname)
                files_written += 1

    if files_written == 0:
        flash(
            f"Файли за {_format_ym_label_ua(selected_ym)} відсутні на диску.",
            "warning",
        )
        return redirect(url_for('military_reports.military_files', month=selected_ym))

    buf.seek(0)
    log_action(
        current_username(),
        f"завантажив пак файлів «Резерв+» за {_format_ym_label_ua(selected_ym)}: "
        f"{len(ready_rows)} студ., {files_written} файл(ів)",
    )
    zip_filename = f"Резерв_плюс_{selected_ym}.zip"
    return send_file(
        buf,
        mimetype='application/zip',
        as_attachment=True,
        download_name=zip_filename,
    )