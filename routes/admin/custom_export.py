"""
routes/admin/custom_export.py
==============================
Гнучка система експорту даних студентів:

  - Довільна комбінація фільтрів (вік, стать, ступінь, форма навчання,
    спеціальність, група, ліцензія, наявність паспорта/військових даних/
    документа про освіту/фото, статус навчання, роки вступу тощо).
  - Довільний вибір колонок для виводу (з ~60 доступних полів).
  - Декілька форматів експорту: XLSX, CSV, DOCX, JSON.
  - Пресети популярних запитів (студенти 25+, 30+ на бакалавраті,
    з паспортом громадянина України тощо).
  - Режим "один рядок на студента" (основна група) або
    "один рядок на навчання" (усі enrollment'и — для паралельного
    навчання).

Право: 'custom_export'. Зареєстровано на admin_bp.
"""

from __future__ import annotations

import csv
import io
import json
from datetime import date, datetime
from typing import Any

from flask import render_template, request, send_file, flash, redirect, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from routes.db import get_db
from routes.utils import logger, permission_required
from routes.helpers import current_username
from routes.admin import admin_bp
from routes.passport_types import PASSPORT_TYPES


# ═══════════════════════════════════════════════════════════════════════
# Утиліти
# ═══════════════════════════════════════════════════════════════════════

def _parse_date_ua(s):
    if not s:
        return None
    for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y.%m.%d'):
        try:
            return datetime.strptime(str(s).strip(), fmt).date()
        except (ValueError, TypeError):
            continue
    return None


def _calc_age(birth_date_str, today=None):
    today = today or date.today()
    bd = _parse_date_ua(birth_date_str)
    if not bd:
        return None
    age = today.year - bd.year
    if (today.month, today.day) < (bd.month, bd.day):
        age -= 1
    return age


def _join_non_empty(parts, sep=' '):
    return sep.join(str(p) for p in parts if p not in (None, ''))


def _yes_no(v):
    return 'так' if v else 'ні'


def _get_int(form, key):
    v = (form.get(key) or '').strip()
    if not v:
        return None
    try:
        return int(v)
    except (ValueError, TypeError):
        return None


def _get_list(form, key):
    return [v for v in form.getlist(key) if v]


# ═══════════════════════════════════════════════════════════════════════
# Опис доступних колонок
# ═══════════════════════════════════════════════════════════════════════

COLUMN_GROUPS = [
    {'key': 'personal', 'label': 'Особисті дані', 'columns': [
        ('student_id', 'ID'),
        ('last_name_UA', 'Прізвище (укр.)'),
        ('first_name_UA', "Ім'я (укр.)"),
        ('middle_name_UA', 'По батькові'),
        ('full_name_UA', 'ПІБ (укр.)'),
        ('last_name_ENG', 'Прізвище (англ.)'),
        ('first_name_ENG', "Ім'я (англ.)"),
        ('birth_date', 'Дата народження'),
        ('age', 'Вік'),
        ('gender', 'Стать'),
        ('tax_id', 'ІПН'),
        ('edebo_code', 'Код ЄДЕБО'),
        ('phone', 'Телефон'),
        ('phone_backup', 'Резервний тел.'),
        ('email', 'Email'),
        ('address_of_residence', 'Адреса'),
        ('has_photo', 'Є фото'),
    ]},
    {'key': 'group', 'label': 'Група та навчання', 'columns': [
        ('group_id', 'ID групи'),
        ('group_name', 'Назва групи'),
        ('course', 'Курс'),
        ('start_year', 'Рік вступу (з групи)'),
        ('degree_level', 'Ступінь'),
        ('study_form', 'Форма навчання'),
        ('specialty', 'Спеціальність'),
        ('specialty_code', 'Код спеціальності'),
        ('program_credits', 'Кредити програми'),
        ('program_credits_override', 'Кредити (скорочена)'),
        ('enrollment_status', 'Статус навчання'),
        ('is_primary', 'Основна група'),
    ]},
    {'key': 'license', 'label': 'Ліцензія', 'columns': [
        ('license_short_name_ua', 'Ліцензія (коротко)'),
        ('license_name_ua', 'Ліцензія (повна)'),
    ]},
    {'key': 'passport', 'label': 'Паспорт', 'columns': [
        ('passport_document_type', 'Тип документа'),
        ('passport_series', 'Серія'),
        ('passport_number', 'Номер'),
        ('passport_issued_by', 'Ким видано'),
        ('passport_issue_date', 'Дата видачі'),
        ('passport_valid_until', 'Дійсний до'),
        ('passport_unique_number', 'Запис у ЄДДР'),
        ('passport_full', 'Паспорт (разом)'),
    ]},
    {'key': 'education', 'label': 'Документ про освіту', 'columns': [
        ('edu_document_type', 'Тип документа'),
        ('edu_document_number', 'Номер'),
        ('edu_institution_name', 'Заклад'),
        ('edu_country', 'Країна'),
        ('edu_completion_date', 'Дата завершення'),
    ]},
    {'key': 'military', 'label': 'Військовий облік', 'columns': [
        ('military_status_badge', 'Статус'),
        ('military_registration_category', 'Категорія обліку'),
        ('military_tck_name', 'ТЦК та СП'),
        ('military_rank', 'Військове звання'),
        ('military_vos_code', 'Код ВОС'),
        ('military_vos_description', 'Опис ВОС'),
        ('military_registry_number', 'Номер «Оберіг»'),
        ('military_deferral_type', 'Тип відстрочки'),
        ('military_deferral_until', 'Відстрочка до'),
        ('military_vlk_decision', 'Постанова ВЛК'),
        ('military_vlk_date', 'Дата ВЛК'),
        ('military_disability_group', 'Група інвалідності'),
        ('military_note_18', 'Примітка (гр. 18)'),
        ('military_formed_at', 'Документ сформовано'),
        ('military_data_update_date', 'Дата уточнення'),
    ]},
    {'key': 'diploma', 'label': 'Диплом', 'columns': [
        ('diploma_number', 'Номер диплома'),
        ('appendix_number', 'Номер додатка'),
    ]},
    {'key': 'periods', 'label': 'Періоди навчання', 'columns': [
        ('study_periods', 'Періоди навчання'),
        ('study_periods_dates', 'Дати навчання'),
    ]},
]

ALL_COLUMN_KEYS = [k for g in COLUMN_GROUPS for k, _ in g['columns']]
COLUMN_LABELS = {k: label for g in COLUMN_GROUPS for k, label in g['columns']}

DEFAULT_COLUMNS = [
    'last_name_UA', 'first_name_UA', 'middle_name_UA',
    'birth_date', 'phone', 'email', 'group_name',
]


# ═══════════════════════════════════════════════════════════════════════
# Пресети популярних запитів
# ═══════════════════════════════════════════════════════════════════════

PRESETS = [
    {
        'key': 'age_25',
        'label': 'Студенти старші 25 років',
        'description': 'Усі активні студенти, яким уже виповнилось 26 і більше',
        'filters': {'min_age': 26},
        'columns': [
            'full_name_UA', 'birth_date', 'age', 'phone', 'email',
            'group_name', 'degree_level', 'study_form',
        ],
    },
    {
        'key': 'age_30_bachelor',
        'label': 'Студенти 30+ на бакалавраті',
        'description': 'Активні студенти віком 31+, які навчаються на бакалавра',
        'filters': {'min_age': 31, 'degree_level': 'Бакалавр'},
        'columns': [
            'full_name_UA', 'birth_date', 'age', 'phone', 'email',
            'group_name', 'degree_level', 'course', 'study_form', 'license_short_name_ua',
        ],
    },
    {
        'key': 'has_ukr_passport',
        'label': 'Студенти з паспортом громадянина України',
        'description': 'Ті, у кого додано паспорт-книжку (не ID-картка)',
        'filters': {'passport_type': ['Паспорт громадянина України']},
        'columns': [
            'full_name_UA', 'group_name',
            'passport_document_type', 'passport_series', 'passport_number',
            'passport_issued_by', 'passport_issue_date',
        ],
    },
    {
        'key': 'has_military',
        'label': 'Студенти з військовим обліком',
        'description': 'Ті, у кого є запис у модулі військового обліку',
        'filters': {'has_military': 'yes'},
        'columns': [
            'full_name_UA', 'group_name', 'military_status_badge',
            'military_tck_name', 'military_rank', 'military_vos_code',
            'military_deferral_until',
        ],
    },
    {
        'key': 'has_deferral',
        'label': 'Студенти з відстрочкою',
        'description': 'Ті, у кого є тип відстрочки у військовому обліку',
        'filters': {'has_military': 'yes'},
        'columns': [
            'full_name_UA', 'group_name', 'military_status_badge',
            'military_deferral_type', 'military_deferral_until',
            'military_tck_name',
        ],
        'post_filter': 'has_deferral',
    },
    {
        'key': 'no_passport',
        'label': 'Студенти без паспортних даних',
        'description': 'Активні студенти, у яких ще не доданий жоден паспорт',
        'filters': {'has_passport': 'no'},
        'columns': [
            'full_name_UA', 'birth_date', 'phone', 'group_name',
        ],
    },
    {
        'key': 'no_enrollment_order',
        'label': 'Студенти без наказу про зарахування',
        'description': 'Активні студенти, у яких немає активного наказу про зарахування',
        'filters': {'has_enrollment_order': 'no'},
        'columns': [
            'full_name_UA', 'birth_date', 'group_name', 'license_short_name_ua',
        ],
    },
]


# ═══════════════════════════════════════════════════════════════════════
# Завантаження та побудова записів
# ═══════════════════════════════════════════════════════════════════════

def _load_reference_data(conn):
    """Завантажує всі дані одним пакетом (без N+1 запитів)."""
    ref = {}

    ref['groups'] = {g['id']: dict(g) for g in conn.execute(
        "SELECT * FROM groups"
    ).fetchall()}

    ref['licenses'] = {l['id']: dict(l) for l in conn.execute(
        "SELECT * FROM institution_licenses"
    ).fetchall()}

    ref['passports'] = {}
    for p in conn.execute(
        "SELECT * FROM passport_documents ORDER BY id DESC"
    ).fetchall():
        ref['passports'].setdefault(p['student_id'], dict(p))

    ref['edu_docs'] = {}
    for d in conn.execute(
        "SELECT * FROM education_documents ORDER BY id DESC"
    ).fetchall():
        ref['edu_docs'].setdefault(d['student_id'], dict(d))

    ref['military'] = {m['student_id']: dict(m) for m in conn.execute(
        "SELECT * FROM military"
    ).fetchall()}

    ref['diplomas'] = {}
    for d in conn.execute(
        "SELECT * FROM diplomas ORDER BY id DESC"
    ).fetchall():
        ref['diplomas'].setdefault((d['student_id'], d['group_id']), dict(d))

    ref['periods'] = {}
    for p in conn.execute(
        "SELECT * FROM student_study_periods "
        "ORDER BY student_id, period_order, start_date"
    ).fetchall():
        ref['periods'].setdefault(p['student_id'], []).append(dict(p))

    ref['frozen'] = {
        r['student_id'] for r in conn.execute(
            "SELECT student_id FROM frozen_students WHERE resolved_at IS NULL"
        ).fetchall()
    }

    ref['enrollment_students'] = {
        r['student_id'] for r in conn.execute(
            "SELECT DISTINCT oi.entity_id AS student_id "
            "FROM orders o JOIN order_items oi ON oi.order_id = o.id "
            "WHERE o.order_type = 'enrollment' AND o.status = 'active'"
        ).fetchall()
    }

    ref['enrollments'] = {}
    for e in conn.execute(
        "SELECT * FROM student_enrollments"
    ).fetchall():
        ref['enrollments'].setdefault(e['student_id'], []).append(dict(e))

    return ref


def _build_record(student, enrollment, ref, today):
    s = dict(student)
    e = enrollment or {}
    gid = e.get('group_id')
    group = ref['groups'].get(gid, {}) if gid else {}

    lic_id = e.get('license_id') if e else s.get('license_id')
    license_row = ref['licenses'].get(lic_id, {}) if lic_id else {}

    p = ref['passports'].get(s['id'], {})
    ed = ref['edu_docs'].get(s['id'], {})
    m = ref['military'].get(s['id'], {})
    dip = ref['diplomas'].get((s['id'], gid), {}) if gid else {}

    all_periods = ref['periods'].get(s['id'], [])
    group_periods = [
        pp for pp in all_periods
        if pp.get('group_id') is None or pp.get('group_id') == gid
    ]

    age = _calc_age(s.get('birth_date'), today)

    passport_parts = []
    if p.get('document_type'):
        passport_parts.append(p['document_type'])
    ser_num = _join_non_empty([p.get('series'), p.get('number')])
    if ser_num:
        passport_parts.append(ser_num)
    if p.get('issued_by'):
        passport_parts.append(p['issued_by'])
    if p.get('issue_date'):
        passport_parts.append(p['issue_date'])
    passport_full = ', '.join(pp for pp in passport_parts if pp)

    return {
        # Особисті
        'student_id': s['id'],
        'last_name_UA': s.get('last_name_UA') or '',
        'first_name_UA': s.get('first_name_UA') or '',
        'middle_name_UA': s.get('middle_name_UA') or '',
        'full_name_UA': _join_non_empty(
            [s.get('last_name_UA'), s.get('first_name_UA'), s.get('middle_name_UA')]
        ),
        'last_name_ENG': s.get('last_name_ENG') or '',
        'first_name_ENG': s.get('first_name_ENG') or '',
        'birth_date': s.get('birth_date') or '',
        'age': age if age is not None else '',
        'gender': {'M': 'Чоловіча', 'F': 'Жіноча'}.get(s.get('gender'), ''),
        'tax_id': s.get('tax_id') or '',
        'edebo_code': e.get('edebo_code') or s.get('edebo_code') or '',
        'phone': s.get('phone') or '',
        'phone_backup': s.get('phone_backup') or '',
        'email': s.get('email') or '',
        'address_of_residence': s.get('address_of_residence') or '',
        'has_photo': _yes_no(s.get('photo')),

        # Група
        'group_id': gid or '',
        'group_name': group.get('name') or '',
        'course': group.get('course') or '',
        'start_year': group.get('start_year') or '',
        'degree_level': group.get('degree_level') or '',
        'study_form': group.get('study_form') or '',
        'specialty': group.get('specialty') or '',
        'specialty_code': group.get('specialty_code') or '',
        'program_credits': group.get('program_credits') or '',
        'program_credits_override': e.get('program_credits_override') or '',
        'enrollment_status': e.get('status') or '',
        'is_primary': _yes_no(e.get('is_primary')),

        # Ліцензія
        'license_id': lic_id or '',
        'license_short_name_ua': license_row.get('short_name_ua') or '',
        'license_name_ua': license_row.get('name_ua') or '',

        # Паспорт
        'passport_document_type': p.get('document_type') or '',
        'passport_series': p.get('series') or '',
        'passport_number': p.get('number') or '',
        'passport_issued_by': p.get('issued_by') or '',
        'passport_issue_date': p.get('issue_date') or '',
        'passport_valid_until': p.get('valid_until') or '',
        'passport_unique_number': p.get('unique_number') or '',
        'passport_full': passport_full,

        # Освіта
        'edu_document_type': ed.get('document_type') or '',
        'edu_document_number': ed.get('document_number') or '',
        'edu_institution_name': ed.get('institution_name') or '',
        'edu_country': ed.get('country') or '',
        'edu_completion_date': ed.get('completion_date') or '',

        # Військові
        'military_status_badge': m.get('status_badge') or '',
        'military_registration_category': m.get('registration_category') or '',
        'military_tck_name': m.get('tck_name') or '',
        'military_rank': m.get('military_rank') or '',
        'military_vos_code': m.get('vos_code') or '',
        'military_vos_description': m.get('vos_description') or '',
        'military_registry_number': m.get('registry_number') or '',
        'military_deferral_type': m.get('deferral_type') or '',
        'military_deferral_until': m.get('deferral_until') or '',
        'military_vlk_decision': m.get('vlk_decision') or '',
        'military_vlk_date': m.get('vlk_date') or '',
        'military_disability_group': m.get('disability_group') or '',
        'military_note_18': m.get('note_18') or '',
        'military_formed_at': m.get('formed_at') or '',
        'military_data_update_date': m.get('data_update_date') or '',

        # Диплом
        'diploma_number': dip.get('diploma_number') or '',
        'appendix_number': dip.get('appendix_number') or '',

        # Періоди
        'study_periods': '; '.join(pp.get('filiya') or '' for pp in group_periods),
        'study_periods_dates': '; '.join(
            f"{pp.get('start_date') or '?'}–{pp.get('end_date') or 'т.ч.'}"
            for pp in group_periods
        ),

        # ── службові поля для фільтрів, не в колонках ──
        '_archived': bool(s.get('archived')),
        '_frozen': s['id'] in ref['frozen'],
        '_has_passport': bool(p.get('id')),
        '_passport_type': p.get('document_type') or '',
        '_has_military': bool(m.get('id')),
        '_has_deferral': bool((m.get('deferral_type') or '').strip()),
        '_has_edu_doc': bool(ed.get('id')),
        '_has_enrollment_order': s['id'] in ref['enrollment_students'],
        '_program_credits_override': e.get('program_credits_override'),
        '_group_program_credits': group.get('program_credits'),
        '_is_reduced': (
            e.get('program_credits_override') is not None
            and group.get('program_credits') is not None
            and int(e['program_credits_override']) < int(group['program_credits'])
        ),
    }


def _build_records(conn, mode, include_no_enrollment):
    """mode: 'student' (основна група) | 'enrollment' (усі групи)."""
    ref = _load_reference_data(conn)
    today = date.today()

    students = conn.execute("SELECT * FROM students").fetchall()
    records = []

    for s in students:
        enr_list = ref['enrollments'].get(s['id'], [])

        if mode == 'student':
            primary = next((e for e in enr_list if e.get('is_primary')), None)
            if primary is None and enr_list:
                primary = enr_list[0]
            if primary is not None:
                records.append(_build_record(s, primary, ref, today))
            elif include_no_enrollment:
                records.append(_build_record(s, None, ref, today))
        else:  # enrollment
            if enr_list:
                for e in enr_list:
                    records.append(_build_record(s, e, ref, today))
            elif include_no_enrollment:
                records.append(_build_record(s, None, ref, today))

    return records


# ═══════════════════════════════════════════════════════════════════════
# Фільтрація
# ═══════════════════════════════════════════════════════════════════════

def _extract_filters(form):
    return {
        'min_age': _get_int(form, 'min_age'),
        'max_age': _get_int(form, 'max_age'),
        'gender': (form.get('gender') or '').strip(),
        'group_id': _get_int(form, 'group_id'),
        'degree_level': (form.get('degree_level') or '').strip(),
        'study_form': (form.get('study_form') or '').strip(),
        'course': _get_int(form, 'course'),
        'specialty_code': (form.get('specialty_code') or '').strip(),
        'license_id': _get_int(form, 'license_id'),
        'start_year_from': _get_int(form, 'start_year_from'),
        'start_year_to': _get_int(form, 'start_year_to'),
        'enrollment_status': _get_list(form, 'enrollment_status') or ['active'],
        'is_reduced': (form.get('is_reduced') or '').strip(),
        'include_archived': form.get('include_archived') == '1',
        'include_frozen': form.get('include_frozen') == '1',
        'has_passport': (form.get('has_passport') or '').strip(),
        'passport_type': _get_list(form, 'passport_type'),
        'has_military': (form.get('has_military') or '').strip(),
        'has_education_doc': (form.get('has_education_doc') or '').strip(),
        'has_photo': (form.get('has_photo') or '').strip(),
        'has_enrollment_order': (form.get('has_enrollment_order') or '').strip(),
    }


def _matches(record, filters):
    # Вік
    if filters['min_age'] is not None or filters['max_age'] is not None:
        age = record.get('age')
        if age == '':
            return False
        if filters['min_age'] is not None and age < filters['min_age']:
            return False
        if filters['max_age'] is not None and age > filters['max_age']:
            return False

    if filters['gender'] and record.get('gender') != filters['gender']:
        return False

    if filters['group_id'] and record.get('group_id') != filters['group_id']:
        return False

    if filters['degree_level'] and record.get('degree_level') != filters['degree_level']:
        return False

    if filters['study_form'] and record.get('study_form') != filters['study_form']:
        return False

    if filters['course'] is not None and record.get('course') != filters['course']:
        return False

    if filters['specialty_code'] and record.get('specialty_code') != filters['specialty_code']:
        return False

    if filters['license_id']:
        # Перевіряємо за ID ліцензії конкретного enrollment'а (або
        # students.license_id для legacy-записів без enrollment'а).
        # record['license_id'] заповнюється в _build_record() з
        # e.license_id (пріоритет) або s.license_id (fallback).
        if record.get('license_id') != filters['license_id']:
            return False

    if filters['start_year_from'] is not None:
        try:
            if int(record.get('start_year') or 0) < filters['start_year_from']:
                return False
        except (ValueError, TypeError):
            return False
    if filters['start_year_to'] is not None:
        try:
            if int(record.get('start_year') or 0) > filters['start_year_to']:
                return False
        except (ValueError, TypeError):
            return False

    # Статус навчання
    if filters['enrollment_status']:
        if record.get('enrollment_status') and record['enrollment_status'] not in filters['enrollment_status']:
            return False
        if not record.get('enrollment_status') and 'none' not in filters['enrollment_status']:
            return False

    # Архів / заморозка
    if not filters['include_archived'] and record.get('_archived'):
        return False
    if not filters['include_frozen'] and record.get('_frozen'):
        return False

    if filters['is_reduced'] == 'yes' and not record.get('_is_reduced'):
        return False
    if filters['is_reduced'] == 'no' and record.get('_is_reduced'):
        return False

    if filters['has_passport'] == 'yes' and not record.get('_has_passport'):
        return False
    if filters['has_passport'] == 'no' and record.get('_has_passport'):
        return False

    if filters['passport_type']:
        if record.get('_passport_type') not in filters['passport_type']:
            return False

    if filters['has_military'] == 'yes' and not record.get('_has_military'):
        return False
    if filters['has_military'] == 'no' and record.get('_has_military'):
        return False

    if filters['has_education_doc'] == 'yes' and not record.get('_has_edu_doc'):
        return False
    if filters['has_education_doc'] == 'no' and record.get('_has_edu_doc'):
        return False

    if filters['has_photo'] == 'yes' and record.get('has_photo') != 'так':
        return False
    if filters['has_photo'] == 'no' and record.get('has_photo') == 'так':
        return False

    if filters['has_enrollment_order'] == 'yes' and not record.get('_has_enrollment_order'):
        return False
    if filters['has_enrollment_order'] == 'no' and record.get('_has_enrollment_order'):
        return False

    return True


def _apply_filters(records, filters, post_filter=None):
    result = [r for r in records if _matches(r, filters)]
    if post_filter == 'has_deferral':
        result = [r for r in result if r.get('_has_deferral')]
    return result


# ═══════════════════════════════════════════════════════════════════════
# Сортування
# ═══════════════════════════════════════════════════════════════════════

def _sort_records(records, sort_by, sort_order):
    if not sort_by:
        return records

    reverse = (sort_order or 'asc').lower() == 'desc'
    try:
        return sorted(records, key=lambda r: (r.get(sort_by) == '', r.get(sort_by) or ''),
                      reverse=reverse)
    except Exception:
        return records


# ═══════════════════════════════════════════════════════════════════════
# Експортери
# ═══════════════════════════════════════════════════════════════════════

def _export_xlsx(records, columns):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Експорт'

    header_font = Font(bold=True, color='FFFFFF', size=10)
    header_fill = PatternFill('solid', fgColor='1F2A3A')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for idx, col_key in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=idx, value=COLUMN_LABELS.get(col_key, col_key))
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    for r_idx, rec in enumerate(records, start=2):
        for c_idx, col_key in enumerate(columns, start=1):
            val = rec.get(col_key, '')
            if val == '':
                val = None
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = Alignment(vertical='top', wrap_text=False)
            cell.border = border

    # Автоширина колонок (приблизна, з обмеженням)
    for c_idx, col_key in enumerate(columns, start=1):
        max_len = len(COLUMN_LABELS.get(col_key, col_key))
        for rec in records:
            v = rec.get(col_key, '')
            if v != '' and v is not None:
                max_len = max(max_len, len(str(v)))
        width = min(max_len + 2, 50)
        ws.column_dimensions[get_column_letter(c_idx)].width = width

    ws.freeze_panes = 'A2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _export_csv(records, columns):
    buf = io.StringIO()
    writer = csv.writer(buf, delimiter=';', quoting=csv.QUOTE_MINIMAL)
    writer.writerow([COLUMN_LABELS.get(k, k) for k in columns])
    for rec in records:
        writer.writerow([rec.get(k, '') for k in columns])
    # Excel любить BOM для коректного відображення кирилиці
    data = '\ufeff' + buf.getvalue()
    return io.BytesIO(data.encode('utf-8'))


def _export_docx(records, columns):
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.section import WD_ORIENT

    doc = Document()
    # Альбомна орієнтація для широких таблиць
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(9)

    doc.add_heading('Експорт даних студентів', level=1)

    table = doc.add_table(rows=1, cols=len(columns))
    table.style = 'Table Grid'

    for idx, col_key in enumerate(columns):
        cell = table.rows[0].cells[idx]
        run = cell.paragraphs[0].add_run(COLUMN_LABELS.get(col_key, col_key))
        run.bold = True
        run.font.size = Pt(9)

    for rec in records:
        row = table.add_row()
        for idx, col_key in enumerate(columns):
            row.cells[idx].text = str(rec.get(col_key, '') or '')
            for p in row.cells[idx].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def _export_json(records, columns):
    data = []
    for rec in records:
        item = {}
        for k in columns:
            v = rec.get(k, '')
            if v is None or v == '':
                item[k] = None
            else:
                item[k] = v
        data.append(item)
    buf = io.BytesIO(json.dumps(data, ensure_ascii=False, indent=2, default=str).encode('utf-8'))
    return buf


EXPORTERS = {
    'xlsx': (_export_xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'xlsx'),
    'csv': (_export_csv, 'text/csv; charset=utf-8', 'csv'),
    'docx': (_export_docx, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'docx'),
    'json': (_export_json, 'application/json; charset=utf-8', 'json'),
}


# ═══════════════════════════════════════════════════════════════════════
# Маршрути
# ═══════════════════════════════════════════════════════════════════════

def _prepare_filter_choices(conn):
    """Довідники для випадаючих списків у формі."""
    groups = conn.execute(
        "SELECT id, name, start_year, study_form, specialty, "
        "       specialty_code, degree_level, course "
        "FROM groups WHERE COALESCE(archived, 0) = 0 "
        "ORDER BY name COLLATE UKRAINIAN"
    ).fetchall()

    licenses = conn.execute(
        "SELECT id, short_name_ua, name_ua FROM institution_licenses "
        "WHERE COALESCE(is_active, 1) = 1 "
        "ORDER BY COALESCE(short_name_ua, name_ua)"
    ).fetchall()

    degree_levels = [r['name_ua'] for r in conn.execute(
        "SELECT name_ua FROM degree_levels WHERE COALESCE(is_active, 1) = 1 ORDER BY id"
    ).fetchall()]

    study_forms = sorted({
        r[0] for r in conn.execute(
            "SELECT DISTINCT study_form FROM groups WHERE study_form IS NOT NULL"
        ).fetchall()
    })

    courses = sorted({
        r[0] for r in conn.execute(
            "SELECT DISTINCT course FROM groups WHERE course IS NOT NULL"
        ).fetchall()
    })

    specialties = conn.execute(
        "SELECT DISTINCT specialty_code, specialty FROM groups "
        "WHERE specialty_code IS NOT NULL "
        "ORDER BY specialty_code"
    ).fetchall()

    return {
        'groups': groups,
        'licenses': licenses,
        'degree_levels': degree_levels,
        'study_forms': study_forms,
        'courses': courses,
        'specialties': specialties,
    }


@admin_bp.route('/admin/custom_export', methods=['GET', 'POST'])
@permission_required('custom_export')
def custom_export():
    """Форма конструктора звітів."""
    conn = get_db()
    conn.row_factory = __import__('sqlite3').Row
    choices = _prepare_filter_choices(conn)
    conn.close()

    selected_columns = _get_list(request.form, 'columns') or DEFAULT_COLUMNS
    selected_columns = [c for c in selected_columns if c in ALL_COLUMN_KEYS]

    return render_template(
        'admin_custom_export.html',
        column_groups=COLUMN_GROUPS,
        presets=PRESETS,
        choices=choices,
        passport_types=PASSPORT_TYPES,
        selected_columns=selected_columns,
        mode=request.form.get('mode', 'student'),
        sort_by=request.form.get('sort_by', 'last_name_UA'),
        sort_order=request.form.get('sort_order', 'asc'),
        form=request.form,
    )


@admin_bp.route('/admin/custom_export/preview', methods=['POST'])
@permission_required('custom_export')
def custom_export_preview():
    """AJAX: порахувати кількість записів за поточними фільтрами."""
    import sqlite3
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        mode = request.form.get('mode', 'student')
        include_no = request.form.get('include_no_enrollment') == '1'
        filters = _extract_filters(request.form)
        post_filter = (request.form.get('post_filter') or '').strip() or None

        records = _build_records(conn, mode, include_no)
        filtered = _apply_filters(records, filters, post_filter=post_filter)

        return {'count': len(filtered), 'total': len(records)}
    except Exception as e:
        logger.error(f"custom_export preview error: {e}", exc_info=True)
        return {'error': str(e)}, 500
    finally:
        conn.close()


@admin_bp.route('/admin/custom_export/download', methods=['POST'])
@permission_required('custom_export')
def custom_export_download():
    """Згенерувати і віддати файл."""
    import sqlite3
    conn = get_db()
    conn.row_factory = sqlite3.Row
    try:
        mode = request.form.get('mode', 'student')
        include_no = request.form.get('include_no_enrollment') == '1'
        filters = _extract_filters(request.form)
        post_filter = (request.form.get('post_filter') or '').strip() or None

        columns = [c for c in _get_list(request.form, 'columns') if c in ALL_COLUMN_KEYS]
        if not columns:
            flash('Оберіть хоча б одну колонку для експорту.', 'error')
            return redirect(url_for('admin.custom_export'))

        file_format = (request.form.get('format') or 'xlsx').lower()
        if file_format not in EXPORTERS:
            file_format = 'xlsx'

        sort_by = (request.form.get('sort_by') or 'last_name_UA').strip()
        sort_order = (request.form.get('sort_order') or 'asc').strip()

        records = _build_records(conn, mode, include_no)
        filtered = _apply_filters(records, filters, post_filter=post_filter)
        filtered = _sort_records(filtered, sort_by, sort_order)

        if not filtered:
            flash('За обраними фільтрами не знайдено жодного запису.', 'warning')
            return redirect(url_for('admin.custom_export'))

        exporter, mimetype, ext = EXPORTERS[file_format]
        buf = exporter(filtered, columns)

        filename = f"custom_export_{datetime.now().strftime('%Y-%m-%d_%H%M')}.{ext}"

        logger.info(
            f"Кастомний експорт: {current_username()}, "
            f"{len(filtered)} записів, колонок: {len(columns)}, "
            f"формат: {file_format}, режим: {mode}"
        )

        return send_file(buf, mimetype=mimetype, as_attachment=True, download_name=filename)
    except Exception as e:
        logger.error(f"custom_export download error: {e}", exc_info=True)
        flash(f'Помилка експорту: {e}', 'error')
        return redirect(url_for('admin.custom_export'))
    finally:
        conn.close()