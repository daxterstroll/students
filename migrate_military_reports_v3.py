"""
migrate_military_reports_v3.py
===============================
Міграція: додає institution_licenses.is_base — позначку "базової"
ліцензії (заклад, для якого ведеться військовий облік у модулі
military_reports).

За Постановою №812 військовий облік ведеться для конкретного закладу
(Львівська філія ПВНЗ «Європейський університет»). Тому на сторінках
Дашборд / Списки груп / Додаток 12 за замовчуванням показуються лише
студенти базової ліцензії (+ без ліцензії, з бейджем «потребує
заповнення»). Решта ліцензій (Київська, Польська тощо) доступні
через чекбокс «Показати всі ліцензії».

Базова ліцензія — рівно одна. При спробі встановити нову — попередня
автоматично стає не базовою.

Ідемпотентний. Бекап перед запуском.

Запуск:
    python migrate_military_reports_v3.py
"""

import os
import shutil
import sqlite3
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'students.db')

# Назва базової ліцензії для початкового автозаповнення.
# Якщо у вас інша — змініть тут або вручну на сторінці «Ліцензії».
DEFAULT_BASE_SHORT_NAME = 'Львівська'


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
    backup_path = f"{DB_PATH}.bak-v3-{stamp}"
    shutil.copy2(DB_PATH, backup_path)
    print(f"💾 Бекап: {os.path.basename(backup_path)}")
    return backup_path


def migrate():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    cur = conn.cursor()

    print()
    print("=" * 70)
    print("  Міграція v3: institution_licenses.is_base (базова ліцензія)")
    print("=" * 70)
    print()

    # ── 1. Додати колонку is_base ───────────────────────────────
    if not _column_exists(cur, 'institution_licenses', 'is_base'):
        cur.execute(
            "ALTER TABLE institution_licenses "
            "ADD COLUMN is_base INTEGER NOT NULL DEFAULT 0"
        )
        print("➕ institution_licenses.is_base — додано")
    else:
        print("✓ institution_licenses.is_base — вже існує")

    # ── 2. Автопризначення базової ліцензії ─────────────────────
    # Якщо жодна не є базовою — призначаємо за short_name_ua (або першу)
    base_exists = cur.execute(
        "SELECT COUNT(*) FROM institution_licenses WHERE is_base = 1"
    ).fetchone()[0]

    if base_exists == 0:
        row = cur.execute(
            "SELECT id, name_ua, short_name_ua FROM institution_licenses "
            "WHERE short_name_ua = ? LIMIT 1",
            (DEFAULT_BASE_SHORT_NAME,),
        ).fetchone()

        if not row:
            # Fallback — перша активна ліцензія за id
            row = cur.execute(
                "SELECT id, name_ua, short_name_ua FROM institution_licenses "
                "ORDER BY id LIMIT 1"
            ).fetchone()

        if row:
            cur.execute(
                "UPDATE institution_licenses SET is_base = 1 WHERE id = ?",
                (row[0],),
            )
            print(f"   ↳ базовою призначено: {row[2]} (ID {row[0]})")
        else:
            print("   ⚠️ institution_licenses порожня — призначте базову вручну на сторінці «Ліцензії»")
    else:
        row = cur.execute(
            "SELECT id, short_name_ua FROM institution_licenses WHERE is_base = 1 LIMIT 1"
        ).fetchone()
        print(f"✓ базова ліцензія вже призначена: {row[1]} (ID {row[0]})")

    # ── 3. Перевірка "рівно одна базова" ────────────────────────
    extra = cur.execute(
        "SELECT COUNT(*) FROM institution_licenses WHERE is_base = 1"
    ).fetchone()[0]
    if extra > 1:
        # Залишаємо тільки першу
        keep_id = cur.execute(
            "SELECT id FROM institution_licenses WHERE is_base = 1 ORDER BY id LIMIT 1"
        ).fetchone()[0]
        cur.execute(
            "UPDATE institution_licenses SET is_base = 0 WHERE id != ?",
            (keep_id,),
        )
        print(f"   ⚠️ було кілька базових — залишено тільки ID {keep_id}")

    conn.commit()

    # ── Статистика ───────────────────────────────────────────────
    print()
    print("─" * 70)
    print("  Стан ліцензій після міграції:")
    print("─" * 70)
    print(f"  {'ID':>4}  {'Базова':<8}  {'Коротка назва':<25}  {'Студентів':>10}")
    print(f"  {'─'*4}  {'─'*8}  {'─'*25}  {'─'*10}")
    for r in cur.execute("""
        SELECT il.id, il.is_base, il.short_name_ua,
               (SELECT COUNT(*) FROM students s
                WHERE s.license_id = il.id AND COALESCE(s.archived, 0) = 0) AS cnt
        FROM institution_licenses il
        ORDER BY il.is_base DESC, il.id
    """):
        marker = '★' if r[1] else ''
        print(f"  {r[0]:>4}  {marker:<8}  {r[2] or '—':<25}  {r[3]:>10}")

    # NULL (без ліцензії)
    null_cnt = cur.execute(
        "SELECT COUNT(*) FROM students "
        "WHERE license_id IS NULL AND COALESCE(archived, 0) = 0"
    ).fetchone()[0]
    print(f"  {'—':>4}  {'':<8}  {'(без ліцензії)':<25}  {null_cnt:>10}")

    conn.close()

    print()
    print("=" * 70)
    print("  ✅ Міграція v3 завершена успішно")
    print("=" * 70)
    print()
    print("Наступний крок — Пакет 1, частина 2: UI на сторінці «Ліцензії».")
    print()


if __name__ == '__main__':
    backup_database()
    try:
        migrate()
    except Exception as e:
        print()
        print(f"❌ ПОМИЛКА: {e}")
        print()
        print(f"Відкат: cp students.db.bak-v3-* students.db")
        import traceback
        traceback.print_exc()
        sys.exit(1)