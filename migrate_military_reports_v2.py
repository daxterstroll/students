"""
migrate_military_reports_v2.py
===============================
Міграція БД під Постанову КМУ №812 від 10.06.2026 (зміни до Порядку
№1487 та Додатка 5).

Що робить:
  1. Додає military.note_18 — позначка з графи 18 «Примітка»:
       - 'Виключено з військового обліку за віком'
       - 'Звільнено з роботи, завершено навчання (відраховано)'
       - 'Військовослужбовець'
  2. Додає military.mobilized_date — дата мобілізації (ДД.ММ.РРРР),
     заповнюється лише коли note_18 = 'Військовослужбовець'.
  3. Додає students.dismissed_at — дата вибуття (РРРР-ММ-ДД).
     Потрібна, щоб записи відрахованих/випускників зберігалися в
     списках лише до кінця поточного календарного року (примітка
     до Списків у №812), а з наступного — не показувались.
  4. Автозаповнює note_18 для вже існуючих АРХІВНИХ студентів:
       - archived=1 AND group_id IS NULL → відраховані
         (у expulsion_confirm() вони отримують саме такий стан)
         → note_18 = 'Звільнено...' + dismissed_at = сьогодні
       - archived=1 AND group_id IS NOT NULL → випускники
         (graduation_confirm() архівує групу, не очищаючи group_id)
         → note_18 НЕ ставимо (випуск ≠ відрахування),
           але dismissed_at = сьогодні (щоб з наступного року не
           показувати в звітах)

Ідемпотентний. Перед запуском робить бекап.

Запуск:
    python migrate_military_reports_v2.py
"""

import os
import shutil
import sqlite3
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'students.db')

# Константи — ті самі значення, що й у routes/military_reports.py
NOTE_EXCLUDED_BY_AGE = 'Виключено з військового обліку за віком'
NOTE_DISMISSED       = 'Звільнено з роботи, завершено навчання (відраховано)'
NOTE_MOBILIZED       = 'Військовослужбовець'

ALLOWED_NOTES = (NOTE_EXCLUDED_BY_AGE, NOTE_DISMISSED, NOTE_MOBILIZED)


def _column_exists(cur, table, column):
    try:
        return column in {row[1] for row in cur.execute(f"PRAGMA table_info({table})")}
    except sqlite3.Error:
        return False


def backup_database():
    if not os.path.exists(DB_PATH):
        print(f"❌ Файл {DB_PATH} не знайдено.")
        sys.exit(1)
    stamp = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    backup_path = f"{DB_PATH}.bak-v2-{stamp}"
    shutil.copy2(DB_PATH, backup_path)
    print(f"💾 Бекап: {os.path.basename(backup_path)}")
    return backup_path


def migrate():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    cur = conn.cursor()

    print()
    print("=" * 70)
    print("  Міграція v2: Постанова КМУ №812 від 10.06.2026")
    print("=" * 70)
    print()

    # ── 1. military.note_18 ──────────────────────────────────────
    if not _column_exists(cur, 'military', 'note_18'):
        cur.execute("""
            ALTER TABLE military ADD COLUMN note_18 TEXT
                CHECK(note_18 IN (
                    'Виключено з військового обліку за віком',
                    'Звільнено з роботи, завершено навчання (відраховано)',
                    'Військовослужбовець'
                ))
        """)
        print("➕ military.note_18 — додано")
    else:
        print("✓ military.note_18 — вже існує")

    # ── 2. military.mobilized_date ───────────────────────────────
    if not _column_exists(cur, 'military', 'mobilized_date'):
        cur.execute("ALTER TABLE military ADD COLUMN mobilized_date TEXT")
        print("➕ military.mobilized_date — додано")
    else:
        print("✓ military.mobilized_date — вже існує")

    # ── 3. students.dismissed_at ─────────────────────────────────
    if not _column_exists(cur, 'students', 'dismissed_at'):
        cur.execute("ALTER TABLE students ADD COLUMN dismissed_at TEXT")
        print("➕ students.dismissed_at — додано")
    else:
        print("✓ students.dismissed_at — вже існує")

    # ── 4. Індекс для швидкого фільтру в звітах ─────────────────
    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_military_note_18
            ON military (note_18)
    """)
    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_students_dismissed_at
            ON students (dismissed_at)
    """)
    print("✓ індекси для note_18 / dismissed_at — створено")

    # ── 5. Автозаповнення для ВЖЕ архівних студентів ────────────
    today_iso = datetime.now().strftime('%Y-%m-%d')

    # 5.1 Відраховані: archived=1 AND group_id IS NULL
    #     → note_18 = 'Звільнено...' + dismissed_at = сьогодні
    expelled_no_note = cur.execute("""
        SELECT s.id, m.id AS military_id
        FROM students s
        INNER JOIN military m ON m.student_id = s.id
        WHERE COALESCE(s.archived, 0) = 1
          AND s.group_id IS NULL
          AND m.note_18 IS NULL
    """).fetchall()

    updated_note = 0
    for row in expelled_no_note:
        cur.execute(
            "UPDATE military SET note_18 = ? WHERE id = ?",
            (NOTE_DISMISSED, row[1]),
        )
        updated_note += 1
    if updated_note:
        print(f"   ↳ проставлено note_18='Звільнено...' для {updated_note} відрахованих")

    # 5.2 dismissed_at для ВСІХ архівних (і відрахованих, і випускників)
    cur.execute(
        "UPDATE students SET dismissed_at = ? "
        "WHERE COALESCE(archived, 0) = 1 AND dismissed_at IS NULL",
        (today_iso,),
    )
    updated_dismissed = cur.rowcount
    if updated_dismissed:
        print(f"   ↳ проставлено dismissed_at='{today_iso}' для {updated_dismissed} архівних")

    # 5.3 Перевірка: чи залишились архівні без dismissed_at
    remaining = cur.execute(
        "SELECT COUNT(*) FROM students "
        "WHERE COALESCE(archived, 0) = 1 AND dismissed_at IS NULL"
    ).fetchone()[0]
    if remaining:
        print(f"   ⚠️ {remaining} архівних студентів залишились без dismissed_at")

    conn.commit()

    # ── Статистика ───────────────────────────────────────────────
    print()
    print("─" * 70)
    print("  Статистика після міграції:")
    print("─" * 70)

    stats = [
        ('Усього military записів',
         "SELECT COUNT(*) FROM military"),
        ('  з них note_18=NULL (на обліку)',
         "SELECT COUNT(*) FROM military WHERE note_18 IS NULL"),
        ('  з них "Виключено за віком"',
         f"SELECT COUNT(*) FROM military WHERE note_18 = '{NOTE_EXCLUDED_BY_AGE}'"),
        ('  з них "Звільнено (відраховано)"',
         f"SELECT COUNT(*) FROM military WHERE note_18 = '{NOTE_DISMISSED}'"),
        ('  з них "Військовослужбовець"',
         f"SELECT COUNT(*) FROM military WHERE note_18 = '{NOTE_MOBILIZED}'"),
        ('Архівних студентів усього',
         "SELECT COUNT(*) FROM students WHERE COALESCE(archived, 0) = 1"),
        ('  з них dismissed_at цього року',
         f"SELECT COUNT(*) FROM students WHERE COALESCE(archived, 0) = 1 "
         f"AND dismissed_at LIKE '{today_iso[:4]}-%'"),
    ]
    for label, sql in stats:
        val = cur.execute(sql).fetchone()[0]
        print(f"  {label:<45} {val:>5}")

    conn.close()

    print()
    print("=" * 70)
    print("  ✅ Міграція v2 завершена успішно")
    print("=" * 70)
    print()
    print("Наступний крок — Крок 2: оновлення backend (routes/military_reports.py).")
    print()


if __name__ == '__main__':
    backup_database()
    try:
        migrate()
    except Exception as e:
        print()
        print(f"❌ ПОМИЛКА: {e}")
        print()
        print("Нічого не зламано — БД у попередньому стані (бекап зроблено).")
        print(f"Відкат: cp students.db.bak-v2-* students.db")
        import traceback
        traceback.print_exc()
        sys.exit(1)