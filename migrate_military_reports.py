"""
migrate_military_reports.py
============================
Одноразова міграція БД для модуля військового обліку студентів
(Додаток 5, Постанова КМУ №916 від 30.07.2025).

Що робить:
  1. Додає students.gender (M/F) — стать студента.
     Автоматично заповнює для існуючих студентів на основі по батькові.
  2. Додає military.category_override ('officer' / 'rank') —
     ручне перевизначення категорії, якщо звання нестандартне і не
     розпізнається автоматично.
  3. Створює tck_offices — довідник ТЦК (67 установ, з наданого CSV).
  4. Створює military_reports_log — журнал генерацій документів
     (для нагадувань про щорічну здачу до 25 січня).

Ідемпотентний — можна запускати повторно безпечно.
Перед запуском робить бекап: students.db → students.db.bak-РРРР-ММ-ДД-ГГ-ХХ-СС

Запуск:
    python migrate_military_reports.py
"""

import os
import shutil
import sqlite3
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'students.db')


# ═══════════════════════════════════════════════════════════════════════
# Допоміжні
# ═══════════════════════════════════════════════════════════════════════

def _column_exists(cur, table, column):
    """Чи існує колонка в таблиці (для ідемпотентності ALTER TABLE)."""
    try:
        return column in {row[1] for row in cur.execute(f"PRAGMA table_info({table})")}
    except sqlite3.Error:
        return False


def _table_exists(cur, table):
    row = cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)
    ).fetchone()
    return row is not None


def detect_gender_from_middle_name(middle_name):
    """
    Визначає стать за по батькові (українська традиція).

    Жіночі закінчення: -івна, -ївна, -овна, -евна, -євна, -ична.
    Все інше — чоловіче. Порожньо — None (треба заповнити вручну).
    """
    if not middle_name:
        return None
    m = middle_name.strip().lower().replace('’', "'").replace('ʼ', "'")
    if m.endswith(('івна', 'ївна', 'овна', 'евна', 'євна', 'ична')):
        return 'F'
    return 'M'


# ═══════════════════════════════════════════════════════════════════════
# 67 ТЦК (з наданого CSV; відновлено коректну українську кодування)
# ═══════════════════════════════════════════════════════════════════════

TCK_OFFICES = [
    # (name_ua, address)
    ('Львівський РТЦК та СП',
     'вул. Дорошенка, 9а, м. Жовква, Львівська обл., Львівський р-н, 80300'),
    ('Личаківсько-Залізничний ОРТЦК та СП',
     'вулиця Батуринська, 2, Львів, Львівська область, 79007'),
    ('2 відділ Самбірського РТЦК та СП (Турка)', None),
    ('Сихівський РТЦК та СП м. Львова',
     'вулиця Стрийська, 105, Львів, Львівська область, 79031'),
    ('4 відділ Львівського РТЦК та СП (Перемишляни)',
     'вулиця Галицька, 26а, м. Перемишляни, Львівський р-н, Львівська обл., 81200'),
    ('1 відділ Львівського РТЦК та СП (Городок)',
     'Львівська обл., Львівський р-он., м. Городок, вул. Львівська, буд. 3, 81500'),
    ('Самбірський РТЦК та СП',
     'вул. Івана Мазепи, 17, м. Самбір, Львівська обл., 81400'),
    ('4 відділ Миколаївського РТЦК та СП (Миколаїв)',
     'м. Пустомити, вул. Грушевського, 36, 81100'),
    ('Перемишлянський РВК', None),
    ('Львівський РТЦК та СП (Жовква)',
     'м. Жовква, вул. Дорошенка, 9 а, 80300'),
    ('Яворівський РТЦК та СП',
     'вул. Львівська, 2, м. Яворів, Львівська обл., 81000'),
    ('Галицько-Франківський ОРТЦК та СП',
     'вулиця Клепарівська, 22, Львів, Львівська область, 79000'),
    ('2 відділ Стрийського РТЦК та СП (Миколаїв Львівська обл.)',
     'площа Ринок, 16, м. Миколаїв, Львівська обл., 81600'),
    ('Стрийський РТЦК та СП',
     'вулиця Остапа Нижанківського, 6, Стрий, Львівська область, 82404'),
    ('Кропивницький міський ТЦК та СП',
     'м. Кропивницький, вул. В. Чміленка, буд. 47, каб. 2'),
    ('Вараський РТЦК та СП',
     'м. Вараш, м-н Перемоги 21-А, Рівненська область'),
    ('4 відділ Івано-Франківського РТЦК та СП (Рогатин)',
     'м. Рогатин, вул. Івана Франка, 4, 77001'),
    ('6 відділ Кропивницького РТЦК та СП (Олександрівка)', None),
    ('2 відділ РТЦК та СП Львівської обл.',
     "м. Кам'янка-Бузька, вул. Шевченка, 57, 80400"),
    ('1 відділ Стрийського РТЦК та СП (Жидачів)',
     'вул. Шашкевича, 17, м. Жидачів, Стрийський район, Львівська область, 81700'),
    ('Калуський РТЦК та СП',
     'м. Калуш, вул. С. Бандери, 25'),
    ('1 відділ Яворівського РТЦК та СП (Мостиська)',
     'м. Мостиська, вул. Міцкевича, 13, 81300'),
    ('Шевченківський РТЦК та СП Львівської обл.',
     'вулиця Батуринська, 2, Львів, Львівська область, 79007'),
    ('2 відділ Шептицького РТЦК та СП (Радехів)',
     'м. Радехів, вул. Стоянівська, 23'),
    ('1 відділ Шептицького РТЦК та СП (Сокаль)',
     'м. Сокаль, вул. Семенюка, 2'),
    ('1 відділ Самбірського РТЦК та СП (Старий Самбір)',
     'м. Старий Самбір, вул. Коцюбинського, 6, 82000'),
    ('Інгульський РТЦК та СП',
     'м. Миколаїв, проспект Миру, 24-а, 54004'),
    ('3 відділ Чернівецького РТЦК та СП (Кіцмань)', None),
    ('2 відділ Броварського РТЦК та СП (Баришівка)', None),
    ('Івано-Франківський МТЦК та СП', None),
    ("2 відділ Львівського РТЦК та СП (Кам'янка-Бузька)", None),
    ('Черкаський ОМТЦК та СП', None),
    ('2 відділ Золочівського РТЦК та СП (Буськ)', None),
    ('Дрогобицький РТЦК та СП', None),
    ('2 відділ Чортківського РТЦК та СП Тернопільської обл.', None),
    ('1 відділ Кременецького РТЦК та СП (Шумськ)', None),
    ('Хустський РТЦК та СП', None),
    ('Броварський РТЦК та СП', None),
    ('Шептицький РТЦК та СП', None),
    ('Ужгородський РТЦК та СП', None),
    ('2 відділ Шепетівського РТЦК та СП (Полонне)', None),
    ('1 відділ Чортківського РТЦК та СП (Гусятин)', None),
    ('1 відділ Коростенського РТЦК та СП (Овруч)', None),
    ('Червоноградський РТЦК та СП', None),
    ('Березівський РТЦК та СП', None),
    ('2 відділ Чернівецького РТЦК та СП (Глибока)', None),
    ('Новокодацький районний у місті Дніпро ТЦК та СП', None),
    ('Вишгородський РТЦК та СП', None),
    ('2 відділ Львівського РТЦК та СП (Городок) Львівської обл.', None),
    ('2 відділ Рівненського РТЦК та СП Рівненської обл.', None),
    ('1 відділ Червоноградського РТЦК та СП', None),
    ('Золочівський РТЦК та СП Львівської обл.',
     'вул. Січових Стрільців, 24, м. Золочів, Львівська обл., 80700'),
    ('Тячівський РТЦК та СП Закарпатської обл.', None),
    ('3 відділ Стрийського РТЦК та СП (Сколе)', None),
    ('2 відділ Первомайського РТЦК та СП (Врадіївка)', None),
    ('Самарський районний у місті Дніпро ТЦК та СП', None),
    ('Чернівецький МТЦК та СП', None),
    ('Рівненський ОМТЦК та СП', None),
    ('3 відділ Львівського РТЦК та СП (Пустомити)', None),
    ('Коростенський РТЦК та СП', None),
    ('Звягельський РТЦК та СП', None),
    ('Одеський РТЦК та СП', None),
    ('Заводський РТЦК та СП м. Миколаїв', None),
    ('Корабельний РТЦК та СП', None),
    ('Центральний РТЦК та СП м. Миколаїв', None),
    ('Луцький ОМТЦК та СП',
     'вул. Теремнівська, 85а, м. Луцьк, Волинська область, 43008'),
]


# ═══════════════════════════════════════════════════════════════════════
# Основна логіка
# ═══════════════════════════════════════════════════════════════════════

def backup_database():
    if not os.path.exists(DB_PATH):
        print(f"❌ Файл {DB_PATH} не знайдено.")
        print("   Запустіть спочатку `python routes/init_db.py` (якщо БД ще немає).")
        sys.exit(1)

    stamp = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    backup_path = f"{DB_PATH}.bak-{stamp}"
    shutil.copy2(DB_PATH, backup_path)
    print(f"💾 Бекап: {os.path.basename(backup_path)}")
    return backup_path


def migrate():
    if not os.path.exists(DB_PATH):
        print(f"❌ Файл {DB_PATH} не знайдено.")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    cur = conn.cursor()

    print()
    print("=" * 66)
    print("  Міграція: модуль військового обліку (Додаток 5, Постанова №916)")
    print("=" * 66)
    print()

    # ── 1. students.gender ────────────────────────────────────────
    if not _column_exists(cur, 'students', 'gender'):
        cur.execute(
            "ALTER TABLE students ADD COLUMN gender TEXT CHECK(gender IN ('M', 'F'))"
        )
        print("➕ students.gender (M/F) — додано")
    else:
        print("✓ students.gender — вже існує")

    # Заповнюємо для тих, у кого NULL
    rows = cur.execute(
        "SELECT id, middle_name_UA FROM students WHERE gender IS NULL"
    ).fetchall()
    if rows:
        updated = 0
        for student_id, middle_name in rows:
            g = detect_gender_from_middle_name(middle_name)
            if g:
                cur.execute(
                    "UPDATE students SET gender = ? WHERE id = ?", (g, student_id)
                )
                updated += 1
        print(f"   ↳ автовизначено стать для {updated} з {len(rows)} студентів")
        if updated < len(rows):
            print(f"   ⚠️ {len(rows) - updated} студентів без по батькові —"
                  f" стать залишилась NULL (заповніть вручну в картці)")

    # ── 2. military.category_override ─────────────────────────────
    if not _column_exists(cur, 'military', 'category_override'):
        cur.execute(
            "ALTER TABLE military ADD COLUMN category_override TEXT "
            "CHECK(category_override IN ('officer', 'rank'))"
        )
        print("➕ military.category_override ('officer'/'rank') — додано")
    else:
        print("✓ military.category_override — вже існує")

    # ── 3. tck_offices ────────────────────────────────────────────
    cur.execute("""
        CREATE TABLE IF NOT EXISTS tck_offices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name_ua TEXT NOT NULL UNIQUE,
            address TEXT,
            is_active INTEGER NOT NULL DEFAULT 1
        )
    """)
    if _table_exists(cur, 'tck_offices'):
        existing = cur.execute("SELECT COUNT(*) FROM tck_offices").fetchone()[0]
        if existing == 0:
            for name, addr in TCK_OFFICES:
                cur.execute(
                    "INSERT OR IGNORE INTO tck_offices (name_ua, address) VALUES (?, ?)",
                    (name, addr),
                )
            inserted = cur.execute("SELECT COUNT(*) FROM tck_offices").fetchone()[0]
            print(f"➕ tck_offices — створено, заповнено {inserted} установ")
        else:
            print(f"✓ tck_offices — вже існує ({existing} установ)")

    # ── 4. military_reports_log ───────────────────────────────────
    cur.execute("""
        CREATE TABLE IF NOT EXISTS military_reports_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_type TEXT NOT NULL,
            report_date TEXT NOT NULL,
            generated_at TEXT DEFAULT (datetime('now','localtime')),
            generated_by TEXT,
            rows_count INTEGER,
            filename TEXT
        )
    """)
    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_military_reports_log_type_date
            ON military_reports_log (report_type, report_date)
    """)
    print("✓ military_reports_log — створено")

    conn.commit()
    conn.close()

    print()
    print("=" * 66)
    print("  ✅ Міграція завершена успішно")
    print("=" * 66)
    print()
    print("Що далі:")
    print("  1. Перевірка БД:   sqlite3 students.db \".schema tck_offices\"")
    print("  2. Перевірка:      sqlite3 students.db \"SELECT COUNT(*) FROM tck_offices\"")
    print("                     (має бути 66-67)")
    print()
    print("Наступний крок — Крок 2: додати поле «Стать» у форму студента.")
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
        print("Якщо хочете відкотитись вручну:")
        print(f"   cp students.db.bak-* students.db")
        import traceback
        traceback.print_exc()
        sys.exit(1)