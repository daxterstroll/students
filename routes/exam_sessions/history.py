"""
routes/exam_sessions/history.py
===============================
Журнал змін сесій та відомостей: запис аудит-логів, переклад усіх технічних
кодів/значень на зрозумілу українську мову та сторінки перегляду історії.
"""
from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from typing import Any

from flask import (
    flash,
    has_request_context,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

from routes.db import get_db
from routes.exam_sessions import exam_sessions_bp
from routes.exam_sessions.schema import ensure_exam_session_tables
from routes.exam_sessions.service import allowed_groups_for_user, user_can_access_group
from routes.helpers import current_username, sort_ukrainian
from routes.utils import logger, permission_required


# ═══════════════════════════════════════════════════════════════════
# СЛОВНИКИ ПЕРЕКЛАДУ (УКР)
# ═══════════════════════════════════════════════════════════════════

ENTITY_TYPE_UA: dict[str, str] = {
    'session': 'Сесія',
    'sheet': 'Відомість',
    'sheet_student': 'Оцінка студента',
    'student': 'Студент у відомості',
    'teacher': 'Викладач',
}

# (Назва дії, візуальний тон для маркера)
ACTION_META_UA: dict[str, tuple[str, str]] = {
    'create': ('Створення', 'ok'),
    'update': ('Оновлення', 'info'),
    'delete': ('Видалення', 'miss'),
    'refresh': ('Оновлення списку студентів', 'info'),
    'grade_set': ('Виставлення оцінки', 'ok'),
    'grade_update': ('Зміна оцінки', 'ok'),
    'approve': ('Затвердження відомості', 'ok'),
    'lock': ('Блокування відомості', 'muted'),
    'unlock': ('Розблокування відомості', 'warn'),
    'submit': ('Подання відомості', 'info'),
    'close_by_teacher': ('Закриття відомості викладачем', 'info'),
    'reopen': ('Повторне відкриття відомості', 'warn'),
    'create_retake': ('Створення відомості перездачі', 'warn'),
    'send': ('Надсилання викладачу', 'info'),
    'transfer': ('Перенос оцінок у план', 'ok'),
    'unlink': ('Відвʼязування від плану', 'warn'),
    'copy': ('Копіювання', 'info'),
    'copy_sheet': ('Копіювання відомості', 'info'),
    'copy_session': ('Копіювання сесії', 'info'),
    'create_session': ('Створення сесії', 'ok'),
    'create_sheet': ('Створення відомості', 'ok'),
    'update_meta': ('Оновлення даних', 'info'),
    'save_meta': ('Оновлення даних', 'info'),
    'save_grades': ('Збереження оцінок', 'ok'),
    'bulk_approve': ('Масове затвердження', 'ok'),
    'refresh_students': ('Оновлення списку студентів', 'info'),
    'unlink_from_plan': ('Відвʼязування від плану', 'warn'),
    'delete_sheet': ('Видалення відомості', 'miss'),
    'force_delete_sheet': ('Примусове видалення відомості', 'miss'),
    'delete_session': ('Видалення сесії', 'miss'),
    'token_reset': ('Оновлення токена', 'warn'),
    'reissue_token': ('Оновлення токена', 'warn'),
}

ACTION_UA: dict[str, str] = {k: v[0] for k, v in ACTION_META_UA.items()}

FIELD_NAME_UA: dict[str, str] = {
    'name': 'Назва',
    'status': 'Статус',
    'notes': 'Примітки',
    'academic_year': 'Навчальний рік',
    'semester': 'Семестр',
    'copy_from': 'Скопійовано з',
    'discipline_name': 'Дисципліна',
    'work_topic': 'Тема роботи',
    'entity_kind': 'Тип освітнього компонента',
    'control_form': 'Форма контролю',
    'hours_credits': 'Години / кредити',
    'teacher_name': 'Викладач',
    'teacher_ref_id': 'Викладач (картотека)',
    'exam_date': 'Дата проведення',
    'statement_number': 'Номер відомості',
    'license_id': 'Ліцензія',
    'points_mid': 'Міжсесійний контроль',
    'points_current': 'Поточні бали',
    'points_exam': 'Іспит / Залік',
    'points': 'Підсумковий бал',
    'national_grade': 'Національна оцінка',
    'ects': 'Оцінка ЄКТС',
    'linked_entity_id': 'Обʼєкт у навчальному плані',
    'linked_entity_type': 'Тип обʼєкта в плані',
    'students_count': 'Кількість студентів',
    'group_id': 'Група',
    'is_active': 'Активність',
    'token': 'Посилання доступу',
    'full_name': 'ПІБ',
    'short_name': 'Скорочене імʼя',
    'department': 'Кафедра',
    'approved_by': 'Затвердив',
    'approved_at': 'Дата затвердження',
    'linked_by': 'Переніс у план',
    'linked_at': 'Дата переносу в план',
    'created_by': 'Створив',
    'teacher_id': 'Викладач',
    'sheet_kind': 'Тип відомості',
    'parent_sheet_id': 'Основна відомість',
    'teacher_closed': 'Закрито викладачем',
    'teacher_closed_at': 'Дата закриття викладачем',
    'reopened_at': 'Дата повторного відкриття',
    'reopened_by': 'Повторно відкрив',
}

VALUE_UA: dict[str, str] = {
    # Статуси сесій / відомостей
    'draft': 'Чернетка',
    'open': 'Відкрита',
    'closed': 'Закрита',
    'sent': 'Надіслано викладачу',
    'in_progress': 'В процесі заповнення',
    'submitted': 'Оцінки здано',
    'approved': 'Затверджено',
    'locked': 'Заблоковано',
    # Типи відомостей
    'main': 'Основна відомість',
    'retake': 'Перездача (ліквідація заборгованості)',
    # Типи компонентів
    'subject': 'Предмет',
    'practice': 'Практика',
    'coursework': 'Курсова робота',
    'attestation': 'Атестація',
    # Форми контролю (на випадок англ. кодів)
    'exam': 'Екзамен',
    'credit': 'Залік',
    'diff_credit': 'Диференційований залік',
    'passed': 'Зараховано',
    'not_passed': 'Не зараховано',
    'failed': 'Незадовільно',
    'absent': 'Не зʼявився',
    # Актор / збереження
    'teacher (cabinet)': 'Викладач (кабінет)',
    'teacher (public form)': 'Викладач (публічна форма)',
    'teacher': 'Викладач',
    'admin': 'Адміністратор',
    'system': 'Система',
    # Булеві значення
    'true': 'Так',
    'false': 'Ні',
    '1': 'Так',
    '0': 'Ні',
}

ENTITY_KIND_UA: dict[str, str] = {
    'subject': 'Предмет',
    'practice': 'Практика',
    'coursework': 'Курсова робота',
    'attestation': 'Атестація',
}

DETAIL_KEY_UA: dict[str, str] = {
    'discipline_name': 'Дисципліна',
    'entity_kind': 'Тип компонента',
    'control_form': 'Форма контролю',
    'license_id': 'Ліцензія',
    'teacher_ref_id': 'Викладач',
    'students_count': 'Студентів у списку',
    'created_for_all_licenses': 'Створено для всіх ліцензій',
    'student_name': 'Студент',
    'saved_by': 'Джерело зміни',
    'source_session_id': 'Сесія-джерело',
    'source_sheet_id': 'Відомість-джерело',
    'sheets_copied': 'Скопійовано відомостей',
    'sheets_count': 'Відомостей у сесії',
    'linked_sheets_count': 'Привʼязаних до плану',
    'grades_removed_from_plan': 'Видалено оцінок з плану',
    'removed_grades': 'Видалено оцінок з плану',
    'was_linked': 'Була перенесена в план',
    'linked_kind': 'Тип у плані',
    'linked_target_id': 'ID обʼєкта в плані',
    'old_target_kind': 'Попередній тип у плані',
    'old_target_id': 'ID обʼєкта в плані',
    'target_kind': 'Тип у плані',
    'target_type': 'Тип у плані',
    'target_id': 'ID обʼєкта в плані',
    'target_name': 'Елемент плану',
    'subject_name': 'Предмет у плані',
    'transferred_count': 'Перенесено оцінок',
    'grades_transferred': 'Перенесено оцінок',
    'updated_count': 'Оновлено записів',
    'added_count': 'Додано записів',
    'removed_count': 'Видалено записів',
    'skipped_count': 'Пропущено',
    'name': 'Назва сесії',
    'group_id': 'Група',
    'group_name': 'Група',
    'academic_year': 'Навчальний рік',
    'semester': 'Семестр',
    'session_name': 'Сесія',
    'reason': 'Причина',
    'comment': 'Коментар',
    'notes': 'Примітки',
    'hours_credits': 'Години / кредити',
    'exam_date': 'Дата проведення',
    'statement_number': 'Номер відомості',
    'work_topic': 'Тема роботи',
    'parent_sheet_id': 'Основна відомість',
    'retake_sheet_id': 'Відомість перездачі',
    'retake_students_count': 'Студентів на перездачу',
    'teacher_closed': 'Закрито викладачем',
    'reopened_by': 'Повторно відкрив',
    'graded_count': 'Заповнено оцінок',
    'prev_status': 'Попередній статус',
    'new_status': 'Новий статус',
    'sheet_kind': 'Тип відомості',
}


# ═══════════════════════════════════════════════════════════════════
# РОБОТА ЗІ СХЕМОЮ ТА ЗАПИС В ІСТОРІЮ
# ═══════════════════════════════════════════════════════════════════

def _ensure_history_table(conn: sqlite3.Connection) -> set[str]:
    """Гарантує існування таблиці exam_session_history та повертає її колонки."""
    try:
        ensure_exam_session_tables(conn)
    except Exception:
        pass
    try:
        from routes.exam_sessions.helpers_extra import ensure_exam_extra_columns
        ensure_exam_extra_columns(conn)
    except Exception:
        pass

    exists = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='exam_session_history'"
    ).fetchone()
    if not exists:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS exam_session_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entity_type TEXT NOT NULL,
                entity_id INTEGER,
                session_id INTEGER,
                sheet_id INTEGER,
                action TEXT NOT NULL,
                field_name TEXT,
                old_value TEXT,
                new_value TEXT,
                details TEXT,
                username TEXT,
                created_at TEXT DEFAULT (datetime('now', 'localtime'))
            )
        """)
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_esh_session ON exam_session_history(session_id)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_esh_sheet ON exam_session_history(sheet_id)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_esh_created ON exam_session_history(created_at)"
        )

    return {
        r[1] for r in conn.execute("PRAGMA table_info(exam_session_history)").fetchall()
    }


def _resolve_actor_for_insert(
    explicit_actor: str | None,
    details_obj: Any,
) -> str:
    if explicit_actor and str(explicit_actor).strip():
        return str(explicit_actor).strip()

    if has_request_context():
        uname = session.get('username')
        if session.get('user_id') and uname:
            return str(uname).strip()
        if isinstance(details_obj, dict) and details_obj.get('saved_by'):
            return str(details_obj['saved_by']).strip()
        try:
            u = current_username()
            if u and str(u).strip() and str(u).strip().lower() not in ('none', 'guest', 'гість'):
                return str(u).strip()
        except Exception:
            pass

    if isinstance(details_obj, dict) and details_obj.get('saved_by'):
        return str(details_obj['saved_by']).strip()
    return 'system'


def log_history(
    conn: sqlite3.Connection,
    entity_type: str,
    entity_id: int | None = None,
    session_id: int | None = None,
    sheet_id: int | None = None,
    action: str = 'update',
    field_name: str | None = None,
    old_value: Any = None,
    new_value: Any = None,
    details: Any = None,
    username: str | None = None,
    actor: str | None = None,
    changed_by: str | None = None,
    **kwargs: Any,
) -> None:
    """Сумісний запис події у журнал exam_session_history."""
    try:
        cols = _ensure_history_table(conn)

        details_str: str | None
        if details is None:
            details_str = None
        elif isinstance(details, (dict, list)):
            details_str = json.dumps(details, ensure_ascii=False)
        else:
            details_str = str(details)

        actor_val = _resolve_actor_for_insert(
            username or actor or changed_by or kwargs.get('actor_name'),
            details,
        )

        old_str = None if old_value is None else str(old_value)
        new_str = None if new_value is None else str(new_value)

        row_data: dict[str, Any] = {
            'entity_type': entity_type,
            'entity_id': entity_id,
            'session_id': session_id,
            'sheet_id': sheet_id,
            'action': action,
            'field_name': field_name,
            'old_value': old_str,
            'new_value': new_str,
            'details': details_str,
        }

        for actor_col in ('username', 'changed_by', 'actor', 'actor_name', 'user_name'):
            if actor_col in cols:
                row_data[actor_col] = actor_val

        if 'actor_type' in cols:
            row_data['actor_type'] = kwargs.get('actor_type') or (
                'teacher' if 'teacher' in actor_val.lower() else (
                    'system' if actor_val.lower() == 'system' else 'user'
                )
            )

        now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for ts_col in ('created_at', 'changed_at', 'timestamp'):
            if ts_col in cols:
                row_data[ts_col] = now_str

        insert_cols = [c for c in row_data.keys() if c in cols]
        placeholders = ', '.join('?' * len(insert_cols))
        col_list = ', '.join(insert_cols)
        vals = [row_data[c] for c in insert_cols]

        conn.execute(
            f"INSERT INTO exam_session_history ({col_list}) VALUES ({placeholders})",
            vals,
        )
    except Exception as exc:
        logger.warning(f"[exam_sessions.history] log_history failed: {exc}")


# ═══════════════════════════════════════════════════════════════════
# ФОРМАТУВАННЯ ТА ПЕРЕКЛАД У ЛЮДИНОЧИТАНИЙ ВИГЛЯД
# ═══════════════════════════════════════════════════════════════════

def _parse_details(raw_details: str | None) -> dict[str, Any]:
    if not raw_details:
        return {}
    if isinstance(raw_details, dict):
        return raw_details
    s = str(raw_details).strip()
    if not s:
        return {}
    try:
        parsed = json.loads(s)
        if isinstance(parsed, dict):
            return parsed
        return {'raw': parsed}
    except Exception:
        return {'_text': s}


def _clean_empty(val: Any) -> str:
    if val is None:
        return ''
    s = str(val).strip()
    if s in ('', 'None', 'null', 'NULL'):
        return ''
    return s


def _translate_value(
    val: Any,
    field_name: str | None,
    lookups: dict[str, dict[Any, str]],
) -> str:
    s = _clean_empty(val)
    if not s:
        return ''

    if field_name == 'teacher_ref_id':
        try:
            tname = lookups['teachers'].get(int(s))
            if tname:
                return tname
        except (TypeError, ValueError):
            pass

    if field_name == 'license_id':
        try:
            lname = lookups['licenses'].get(int(s))
            if lname:
                return lname
        except (TypeError, ValueError):
            pass

    if field_name == 'group_id':
        try:
            gname = lookups['groups'].get(int(s))
            if gname:
                return gname
        except (TypeError, ValueError):
            pass

    if field_name == 'copy_from':
        return f"#{s}"

    if field_name in ('points_mid', 'points_current', 'points_exam', 'points'):
        try:
            fval = float(s)
            if fval.is_integer():
                return str(int(fval))
            return f"{fval:g}"
        except (TypeError, ValueError):
            pass

    if field_name == 'is_active':
        if s in ('1', 'True', 'true'):
            return 'Активний'
        if s in ('0', 'False', 'false'):
            return 'Неактивний'

    low = s.lower()
    if low in VALUE_UA:
        return VALUE_UA[low]

    return s


def _translate_actor(
    raw_actor: str | None,
    details: dict[str, Any],
    sheet_teacher_name: str | None,
) -> tuple[str, str]:
    """Повертає (відображуване_імʼя, тип: admin|teacher|system)."""
    saved_by = _clean_empty(details.get('saved_by'))
    raw = _clean_empty(raw_actor)

    candidate = raw or saved_by
    low = candidate.lower()

    if low in ('teacher (cabinet)', 'teacher_cabinet') or saved_by.lower() == 'teacher (cabinet)':
        label = 'Викладач (кабінет)'
        if sheet_teacher_name:
            label = f"{sheet_teacher_name} · кабінет"
        return label, 'teacher'

    if low in ('teacher (public form)', 'teacher_public') or saved_by.lower() == 'teacher (public form)':
        label = 'Викладач (публічне посилання)'
        if sheet_teacher_name:
            label = f"{sheet_teacher_name} · посилання"
        return label, 'teacher'

    if low == 'teacher':
        return (sheet_teacher_name or 'Викладач'), 'teacher'

    if not candidate or low in ('system', 'гість', 'guest', 'anonymous'):
        return 'Система', 'system'

    if low == 'admin':
        return 'Адміністратор (admin)', 'admin'

    return candidate, 'admin'


def _format_timestamp_ua(raw_ts: str | None) -> tuple[str, str]:
    """Перетворює YYYY-MM-DD HH:MM:SS у ('DD.MM.YYYY', 'HH:MM:SS')."""
    if not raw_ts:
        return '—', ''
    s = str(raw_ts).strip().replace('T', ' ')
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try:
            dt = datetime.strptime(s[:len(fmt) + 2].strip(), fmt) if fmt != '%Y-%m-%d %H:%M:%S' else datetime.strptime(s[:19], fmt)
            date_part = dt.strftime('%d.%m.%Y')
            time_part = dt.strftime('%H:%M:%S') if '%H' in fmt else ''
            return date_part, time_part
        except ValueError:
            continue
    parts = s.split(' ', 1)
    return parts[0], (parts[1][:8] if len(parts) > 1 else '')


def _format_details_items(
    details: dict[str, Any],
    action: str,
    entity_type: str,
    lookups: dict[str, dict[Any, str]],
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    """
    Перетворює JSON details у:
      - список зрозумілих пар {'label': ..., 'value': ...}
      - список видалених відомостей (якщо це snapshot видалення сесії)
    """
    if '_text' in details:
        return [{'label': 'Примітка', 'value': str(details['_text'])}], []

    items: list[dict[str, str]] = []
    deleted_sheets: list[dict[str, str]] = []

    if isinstance(details.get('sheets'), list):
        for sh in details['sheets']:
            if isinstance(sh, dict):
                deleted_sheets.append({
                    'id': str(sh.get('id') or '—'),
                    'discipline': str(sh.get('discipline') or sh.get('discipline_name') or 'Без назви'),
                })

    # Ключі, які вже виносяться в заголовок/контекст або мають технічний характер
    skip_keys = {
        'sheets',
        'student_name',
        'saved_by',
        'session_id',
    }
    # Якщо це проста зміна оцінки студента — student_name вже у колонці "Подія"
    if action in ('create', 'delete') and entity_type == 'session':
        # group_name і academic_year і semester корисні у деталях при видаленні/створенні
        pass

    # Комбінуємо old_target_kind + old_target_id або linked_kind + linked_target_id
    kind_key = None
    id_key = None
    for kk, ik in (
        ('old_target_kind', 'old_target_id'),
        ('linked_kind', 'linked_target_id'),
        ('target_kind', 'target_id'),
        ('target_type', 'target_id'),
    ):
        if details.get(kk) or details.get(ik):
            kind_key, id_key = kk, ik
            break

    if kind_key:
        raw_k = _clean_empty(details.get(kind_key))
        raw_i = _clean_empty(details.get(id_key))
        k_ua = ENTITY_KIND_UA.get(raw_k.lower(), raw_k)
        val_txt = f"{k_ua} #{raw_i}".strip() if raw_i else k_ua
        lbl = 'Попередній обʼєкт у плані' if 'old_' in kind_key else 'Обʼєкт у плані'
        if val_txt:
            items.append({'label': lbl, 'value': val_txt})
        skip_keys.update({kind_key, id_key})

    ordered_keys = [
        'discipline_name',
        'name',
        'group_name',
        'group_id',
        'academic_year',
        'semester',
        'entity_kind',
        'control_form',
        'license_id',
        'teacher_ref_id',
        'students_count',
        'sheets_count',
        'sheets_copied',
        'linked_sheets_count',
        'transferred_count',
        'grades_transferred',
        'removed_grades',
        'grades_removed_from_plan',
        'source_session_id',
        'source_sheet_id',
        'created_for_all_licenses',
        'was_linked',
        'target_name',
        'subject_name',
        'reason',
        'comment',
    ]
    seen = set(skip_keys)

    def _emit_key(k: str) -> None:
        if k in seen:
            return
        seen.add(k)
        v = details.get(k)
        if v is None or v == '':
            return
        if k == 'created_for_all_licenses':
            if v:
                items.append({'label': 'Режим', 'value': 'Створено окремо для кожної ліцензії'})
            return
        if k == 'was_linked':
            items.append({'label': 'Була перенесена в план', 'value': 'Так' if v else 'Ні'})
            return
        if k == 'group_id':
            if details.get('group_name'):
                return
            try:
                gname = lookups['groups'].get(int(v))
                items.append({'label': 'Група', 'value': gname or f"#{v}"})
            except (TypeError, ValueError):
                items.append({'label': 'Група', 'value': str(v)})
            return
        if k == 'license_id':
            try:
                lname = lookups['licenses'].get(int(v))
                items.append({'label': 'Ліцензія', 'value': lname or f"#{v}"})
            except (TypeError, ValueError):
                items.append({'label': 'Ліцензія', 'value': str(v)})
            return
        if k == 'teacher_ref_id':
            try:
                tname = lookups['teachers'].get(int(v))
                items.append({'label': 'Викладач', 'value': tname or f"#{v}"})
            except (TypeError, ValueError):
                items.append({'label': 'Викладач', 'value': str(v)})
            return
        if k == 'entity_kind':
            items.append({'label': 'Тип компонента', 'value': ENTITY_KIND_UA.get(str(v).lower(), str(v))})
            return
        if k in ('source_session_id', 'source_sheet_id'):
            lbl = DETAIL_KEY_UA[k]
            items.append({'label': lbl, 'value': f"#{v}"})
            return
        if isinstance(v, bool):
            val_str = 'Так' if v else 'Ні'
        elif isinstance(v, (dict, list)):
            val_str = json.dumps(v, ensure_ascii=False)
        else:
            val_str = VALUE_UA.get(str(v).lower(), str(v))
        label = DETAIL_KEY_UA.get(k, k.replace('_', ' ').capitalize())
        items.append({'label': label, 'value': val_str})

    for k in ordered_keys:
        if k in details:
            _emit_key(k)
    for k in details.keys():
        _emit_key(k)

    return items, deleted_sheets


def _load_lookups(conn: sqlite3.Connection) -> dict[str, dict[Any, str]]:
    teachers = {}
    licenses = {}
    groups = {}
    try:
        for r in conn.execute("SELECT id, short_name, full_name FROM teachers").fetchall():
            teachers[r[0]] = (r[1] or r[2] or f"#{r[0]}").strip()
    except Exception:
        pass
    try:
        for r in conn.execute("SELECT id, short_name_ua, name_ua FROM institution_licenses").fetchall():
            licenses[r[0]] = (r[1] or r[2] or f"Ліцензія #{r[0]}").strip()
    except Exception:
        pass
    try:
        for r in conn.execute("SELECT id, name FROM groups").fetchall():
            groups[r[0]] = (r[1] or f"Група #{r[0]}").strip()
    except Exception:
        pass
    return {'teachers': teachers, 'licenses': licenses, 'groups': groups}


def format_history_row(
    raw: dict[str, Any],
    lookups: dict[str, dict[Any, str]],
    deleted_session_meta: dict[int, dict[str, Any]],
) -> dict[str, Any]:
    """Перетворює рядок з БД у повністю україномовну та структуровану модель перегляду."""
    details = _parse_details(raw.get('details'))
    entity_type = (raw.get('entity_type') or '').strip()
    action = (raw.get('action') or '').strip()
    field_name = _clean_empty(raw.get('field_name'))

    action_label, action_tone = ACTION_META_UA.get(
        action,
        (action.replace('_', ' ').capitalize() or 'Зміна', 'muted'),
    )
    entity_label = ENTITY_TYPE_UA.get(
        entity_type,
        entity_type.replace('_', ' ').capitalize() or 'Запис',
    )

    # Контекст сесії та групи (навіть якщо видалено)
    sess_id = raw.get('session_id')
    sheet_id = raw.get('sheet_id')
    sess_snap = deleted_session_meta.get(sess_id) if sess_id else None

    session_exists = bool(raw.get('live_session_id'))
    sheet_exists = bool(raw.get('live_sheet_id'))

    group_id = raw.get('live_group_id') or details.get('group_id') or (sess_snap or {}).get('group_id')
    group_name = (
        _clean_empty(raw.get('live_group_name'))
        or _clean_empty(details.get('group_name'))
        or _clean_empty((sess_snap or {}).get('group_name'))
    )
    if not group_name and group_id:
        try:
            group_name = lookups['groups'].get(int(group_id), '')
        except (TypeError, ValueError):
            pass

    session_name = (
        _clean_empty(raw.get('live_session_name'))
        or _clean_empty(details.get('session_name'))
        or (_clean_empty(details.get('name')) if entity_type == 'session' else '')
        or _clean_empty((sess_snap or {}).get('name'))
    )
    if not session_name and sess_id:
        session_name = f"Сесія #{sess_id}"

    discipline_name = (
        _clean_empty(raw.get('live_discipline_name'))
        or _clean_empty(details.get('discipline_name'))
    )
    license_name = _clean_empty(raw.get('live_license_name'))
    if not license_name and details.get('license_id'):
        try:
            license_name = lookups['licenses'].get(int(details['license_id']), '')
        except (TypeError, ValueError):
            pass

    student_name = _clean_empty(details.get('student_name'))
    if student_name.lower().startswith('student#'):
        student_name = 'Студент #' + student_name[8:]
    if not student_name and entity_type == 'sheet_student' and raw.get('entity_id'):
        student_name = f"Студент #{raw['entity_id']}"

    teacher_name = _clean_empty(raw.get('live_teacher_name'))
    if not teacher_name and details.get('teacher_ref_id'):
        try:
            teacher_name = lookups['teachers'].get(int(details['teacher_ref_id']), '')
        except (TypeError, ValueError):
            pass

    actor_label, actor_kind = _translate_actor(
        raw.get('raw_actor'),
        details,
        teacher_name,
    )

    date_ua, time_ua = _format_timestamp_ua(raw.get('raw_ts'))

    field_ua = FIELD_NAME_UA.get(field_name, field_name) if field_name else ''
    old_ua = _translate_value(raw.get('old_value'), field_name, lookups)
    new_ua = _translate_value(raw.get('new_value'), field_name, lookups)

    # Зрозумілий заголовок зміни, якщо field_name не передано (напр. approve/lock без field_name)
    if not field_ua and (old_ua or new_ua):
        if action in ('approve', 'lock', 'unlock', 'submit', 'send', 'close_by_teacher', 'reopen'):
            field_ua = 'Статус'
        elif action == 'create' and new_ua:
            field_ua = 'Значення'

    detail_items, deleted_sheets = _format_details_items(details, action, entity_type, lookups)

    # Пошуковий індекс (все українською + ID + оригінал)
    search_parts = [
        date_ua, time_ua, actor_label, action_label, entity_label,
        session_name, group_name, discipline_name, license_name, student_name,
        field_ua, old_ua, new_ua,
    ]
    for it in detail_items:
        search_parts.append(it['label'])
        search_parts.append(it['value'])
    for ds in deleted_sheets:
        search_parts.append(ds['discipline'])
    search_blob = ' '.join(p for p in search_parts if p).lower()

    return {
        'id': raw['id'],
        'date_ua': date_ua,
        'time_ua': time_ua,
        'raw_ts': raw.get('raw_ts') or '',
        'actor_label': actor_label,
        'actor_kind': actor_kind,
        'action': action,
        'action_label': action_label,
        'action_tone': action_tone,
        'entity_type': entity_type,
        'entity_label': entity_label,
        'entity_id': raw.get('entity_id'),
        'session_id': sess_id,
        'sheet_id': sheet_id,
        'session_exists': session_exists,
        'sheet_exists': sheet_exists,
        'session_name': session_name,
        'group_id': group_id,
        'group_name': group_name,
        'discipline_name': discipline_name,
        'license_name': license_name,
        'student_name': student_name,
        'field_ua': field_ua,
        'old_ua': old_ua,
        'new_ua': new_ua,
        'detail_items': detail_items,
        'deleted_sheets': deleted_sheets,
        'search_blob': search_blob,
    }


# ═══════════════════════════════════════════════════════════════════
# ЗАПИТИ ТА ФІЛЬТРАЦІЯ ІСТОРІЇ
# ═══════════════════════════════════════════════════════════════════

def _query_and_render_history(
    *,
    fixed_session_id: int | None = None,
    fixed_sheet_id: int | None = None,
):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    cols = _ensure_history_table(conn)

    # Перевірка контексту для сторінки конкретної сесії або відомості
    context_session = None
    context_sheet = None

    if fixed_sheet_id is not None:
        sh_row = conn.execute("""
            SELECT ss.*, es.group_id, es.name AS session_name, es.academic_year, es.semester,
                   g.name AS group_name, l.short_name_ua AS license_name
            FROM session_sheets ss
            JOIN exam_sessions es ON es.id = ss.session_id
            JOIN groups g ON g.id = es.group_id
            LEFT JOIN institution_licenses l ON l.id = ss.license_id
            WHERE ss.id = ?
        """, (fixed_sheet_id,)).fetchone()
        if sh_row:
            if not user_can_access_group(conn, session, sh_row['group_id']):
                conn.close()
                flash('Немає доступу до цієї відомості', 'error')
                return redirect(url_for('exam_sessions.history_index'))
            context_sheet = dict(sh_row)
            fixed_session_id = context_sheet['session_id']

    if fixed_session_id is not None:
        es_row = conn.execute("""
            SELECT es.*, g.name AS group_name, g.course, g.study_form
            FROM exam_sessions es
            JOIN groups g ON g.id = es.group_id
            WHERE es.id = ?
        """, (fixed_session_id,)).fetchone()
        if es_row:
            if not user_can_access_group(conn, session, es_row['group_id']):
                conn.close()
                flash('Немає доступу до цієї сесії', 'error')
                return redirect(url_for('exam_sessions.history_index'))
            context_session = dict(es_row)

    allowed_groups = allowed_groups_for_user(conn, session)
    allowed_group_ids = {g['id'] for g in allowed_groups}
    allowed_group_names = {
        (g.get('name') or '').strip() for g in allowed_groups if g.get('name')
    }
    is_admin = bool(session.get('is_admin'))

    # Визначаємо колонки актора і часу в наявній таблиці
    actor_candidates = [c for c in ('username', 'changed_by', 'actor', 'actor_name', 'user_name') if c in cols]
    if len(actor_candidates) == 1:
        actor_expr = "NULLIF(TRIM(h." + actor_candidates[0] + "), '')"
    elif len(actor_candidates) > 1:
        actor_parts = ["NULLIF(TRIM(h." + c + "), '')" for c in actor_candidates]
        actor_expr = "COALESCE(" + ", ".join(actor_parts) + ")"
    else:
        actor_expr = "NULL"
    ts_col = next((c for c in ('created_at', 'changed_at', 'timestamp') if c in cols), None)
    ts_expr = f"h.{ts_col}" if ts_col else "NULL"

    # Збираємо метадані видалених сесій, щоб відновити назву/групу/доступ для старих рядків
    deleted_session_meta: dict[int, dict[str, Any]] = {}
    for dr in conn.execute(
        "SELECT session_id, details FROM exam_session_history "
        "WHERE entity_type = 'session' AND action IN ('delete', 'create') AND details IS NOT NULL"
    ).fetchall():
        sid = dr['session_id']
        if not sid:
            continue
        parsed = _parse_details(dr['details'])
        if parsed:
            prev = deleted_session_meta.get(sid, {})
            prev.update({k: v for k, v in parsed.items() if v is not None})
            deleted_session_meta[sid] = prev

    # Параметри фільтрації з query string
    f_session_id = fixed_session_id if fixed_session_id is not None else request.args.get('session_id', type=int)
    f_sheet_id = fixed_sheet_id if fixed_sheet_id is not None else request.args.get('sheet_id', type=int)
    f_group_id = request.args.get('group_id', type=int)
    f_entity_type = (request.args.get('entity_type') or '').strip()
    f_action = (request.args.get('action') or '').strip()
    f_actor = (request.args.get('actor') or '').strip()
    f_date_from = (request.args.get('date_from') or '').strip()
    f_date_to = (request.args.get('date_to') or '').strip()
    f_q = (request.args.get('q') or '').strip()
    page = max(1, request.args.get('page', 1, type=int) or 1)
    per_page = request.args.get('per_page', 50, type=int) or 50
    if per_page not in (25, 50, 100):
        per_page = 50

    where_clauses: list[str] = []
    params: list[Any] = []

    if f_session_id is not None:
        where_clauses.append("h.session_id = ?")
        params.append(f_session_id)
    if f_sheet_id is not None:
        where_clauses.append("h.sheet_id = ?")
        params.append(f_sheet_id)
    if f_entity_type:
        where_clauses.append("h.entity_type = ?")
        params.append(f_entity_type)
    if f_action:
        where_clauses.append("h.action = ?")
        params.append(f_action)
    if ts_col and f_date_from:
        where_clauses.append(f"substr(h.{ts_col}, 1, 10) >= ?")
        params.append(f_date_from)
    if ts_col and f_date_to:
        where_clauses.append(f"substr(h.{ts_col}, 1, 10) <= ?")
        params.append(f_date_to)

    where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""

    raw_rows = conn.execute(f"""
        SELECT h.*,
               {actor_expr} AS raw_actor,
               {ts_expr} AS raw_ts,
               es.id AS live_session_id,
               es.name AS live_session_name,
               es.group_id AS live_group_id,
               g.name AS live_group_name,
               ss.id AS live_sheet_id,
               ss.discipline_name AS live_discipline_name,
               l.short_name_ua AS live_license_name,
               COALESCE(t.short_name, t.full_name, ss.teacher_name) AS live_teacher_name
        FROM exam_session_history h
        LEFT JOIN exam_sessions es ON es.id = h.session_id
        LEFT JOIN groups g ON g.id = es.group_id
        LEFT JOIN session_sheets ss ON ss.id = h.sheet_id
        LEFT JOIN institution_licenses l ON l.id = ss.license_id
        LEFT JOIN teachers t ON t.id = ss.teacher_ref_id
        {where_sql}
        ORDER BY h.id DESC
    """, params).fetchall()

    lookups = _load_lookups(conn)

    # Список сесій у випадаючий фільтр (у межах доступу)
    accessible_sessions = []
    if allowed_group_ids:
        ph = ','.join('?' * len(allowed_group_ids))
        accessible_sessions = [
            dict(r) for r in conn.execute(f"""
                SELECT es.id, es.name, es.academic_year, es.semester, g.name AS group_name
                FROM exam_sessions es
                JOIN groups g ON g.id = es.group_id
                WHERE es.group_id IN ({ph})
                ORDER BY es.created_at DESC
            """, list(allowed_group_ids)).fetchall()
        ]

    # Також зберемо унікальні значення дій/акторів у межах скоупу для фільтрів
    scoped_rows: list[dict[str, Any]] = []
    actors_set: set[str] = set()
    actions_set: set[str] = set()
    entity_types_set: set[str] = set()

    for r in raw_rows:
        rd = dict(r)
        formatted = format_history_row(rd, lookups, deleted_session_meta)

        # Перевірка доступу користувача (навіть для видалених сесій/відомостей)
        if not is_admin:
            gid = formatted.get('group_id')
            gname = (formatted.get('group_name') or '').strip()
            has_access = False
            if gid is not None:
                try:
                    has_access = int(gid) in allowed_group_ids
                except (TypeError, ValueError):
                    has_access = False
            elif gname:
                has_access = gname in allowed_group_names
            if not has_access:
                continue

        if f_group_id is not None:
            try:
                if int(formatted.get('group_id') or -1) != f_group_id:
                    continue
            except (TypeError, ValueError):
                continue

        actors_set.add(formatted['actor_label'])
        if formatted['action']:
            actions_set.add(formatted['action'])
        if formatted['entity_type']:
            entity_types_set.add(formatted['entity_type'])

        if f_actor and formatted['actor_label'] != f_actor:
            continue

        if f_q:
            tokens = [t for t in f_q.lower().split() if t]
            blob = formatted['search_blob']
            if not all(t in blob for t in tokens):
                continue

        scoped_rows.append(formatted)

    # Якщо відкрито сторінку видаленої сесії/відомості, перевіримо що хоч щось доступне або покажемо контекст
    if fixed_session_id is not None and context_session is None and not scoped_rows:
        meta = deleted_session_meta.get(fixed_session_id)
        if meta and not is_admin:
            mgid = meta.get('group_id')
            mgname = (meta.get('group_name') or '').strip()
            if (mgid not in allowed_group_ids) and (mgname not in allowed_group_names):
                conn.close()
                flash('Немає доступу до історії цієї сесії', 'error')
                return redirect(url_for('exam_sessions.history_index'))

    # Додаткові варіанти для селектів (навіть коли вибрано один фільтр)
    if f_entity_type or f_action or f_actor or f_q or f_date_from or f_date_to:
        # Додаємо стандартні варіанти, щоб користувач міг перемкнутись без скидання
        for k in ('session', 'sheet', 'sheet_student'):
            entity_types_set.add(k)
        for k in ('create', 'update', 'grade_set', 'approve', 'lock', 'transfer', 'unlink', 'refresh', 'delete'):
            actions_set.add(k)

    today_ua = datetime.now().strftime('%d.%m.%Y')
    summary = {
        'total': len(scoped_rows),
        'today': sum(1 for x in scoped_rows if x['date_ua'] == today_ua),
        'grades_and_approvals': sum(
            1 for x in scoped_rows if x['action'] in ('grade_set', 'grade_update', 'approve', 'transfer', 'lock')
        ),
        'deletes': sum(1 for x in scoped_rows if x['action'] == 'delete'),
    }

    total_count = len(scoped_rows)
    total_pages = max(1, (total_count + per_page - 1) // per_page)
    if page > total_pages:
        page = total_pages
    start_idx = (page - 1) * per_page
    page_rows = scoped_rows[start_idx:start_idx + per_page]

    action_options = [
        (code, ACTION_UA.get(code, code.replace('_', ' ').capitalize()))
        for code in sorted(actions_set, key=lambda c: ACTION_UA.get(c, c))
    ]
    entity_options = [
        (code, ENTITY_TYPE_UA.get(code, code.replace('_', ' ').capitalize()))
        for code in sorted(entity_types_set, key=lambda c: ENTITY_TYPE_UA.get(c, c))
    ]
    actor_options = sort_ukrainian(list(actors_set), lambda x: x)

    deleted_meta_for_fixed = None
    if fixed_session_id is not None and context_session is None:
        m = deleted_session_meta.get(fixed_session_id) or {}
        deleted_meta_for_fixed = {
            'id': fixed_session_id,
            'name': m.get('name') or f"Сесія #{fixed_session_id}",
            'group_name': m.get('group_name') or lookups['groups'].get(m.get('group_id'), ''),
            'academic_year': m.get('academic_year') or '',
            'semester': m.get('semester') or '',
        }

    conn.close()

    return render_template(
        'exam_sessions/history.html',
        rows=page_rows,
        summary=summary,
        total_count=total_count,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        groups=allowed_groups,
        sessions=accessible_sessions,
        action_options=action_options,
        entity_options=entity_options,
        actor_options=actor_options,
        context_session=context_session,
        context_sheet=context_sheet,
        deleted_session_meta=deleted_meta_for_fixed,
        fixed_session_id=fixed_session_id,
        fixed_sheet_id=fixed_sheet_id,
        filters={
            'session_id': f_session_id,
            'sheet_id': f_sheet_id,
            'group_id': f_group_id,
            'entity_type': f_entity_type,
            'action': f_action,
            'actor': f_actor,
            'date_from': f_date_from,
            'date_to': f_date_to,
            'q': f_q,
            'per_page': per_page,
        },
    )


# ═══════════════════════════════════════════════════════════════════
# МАРШРУТИ
# ═══════════════════════════════════════════════════════════════════

@exam_sessions_bp.route('/admin/exam-sessions/history')
@permission_required('view_exam_session_history')
def history_index():
    """Загальний журнал усіх змін по сесіях та відомостях."""
    return _query_and_render_history()


@exam_sessions_bp.route('/admin/exam-sessions/<int:session_id>/history')
@permission_required('view_exam_session_history')
def session_history(session_id: int):
    """Історія змін по конкретній сесії."""
    return _query_and_render_history(fixed_session_id=session_id)


@exam_sessions_bp.route('/admin/exam-sessions/sheets/<int:sheet_id>/history')
@permission_required('view_exam_session_history')
def sheet_history(sheet_id: int):
    """Історія змін по конкретній відомості."""
    return _query_and_render_history(fixed_sheet_id=sheet_id)
