# -*- coding: utf-8 -*-
"""
Міграція: переносить "Найменування та статус закладу" з рівня ГРУПИ на
рівень СТУДЕНТА (як ліцензію вступу - Львівська/Київська/Польська
тощо). Безпечна для запуску на вже наявній базі.

Що робить:
1. Створює каталог institution_licenses + таблиці наказу переведення.
2. Засіює каталог двома ліцензіями, які раніше були жорстко прописані
   в групах (Київська/Львівська) - з тим самим текстом, тому наявні
   документи не зміняться.
3. Додає students.license_id.
4. ПЕРЕНОСИТЬ ДАНІ: для кожного студента визначає ліцензію за текстом
   institution_name_and_status його ПОТОЧНОЇ групи і проставляє
   відповідний license_id - щоб після міграції нічого не "спорожніло".
   Студенти без групи чи з незрозумілим текстом лишаються без ліцензії
   (license_id=NULL) - доведеться виставити вручну.
5. Поля groups.institution_name_and_status/_en НЕ видаляються з бази
   (про всяк випадок, для старих даних) - просто більше не
   використовуються ні формою, ні генерацією документів.

Запуск:  python migrate_licenses.py
"""
import sqlite3
import os

db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'students.db')
conn = sqlite3.connect(db_path)
conn.row_factory = sqlite3.Row
cur = conn.cursor()

cur.executescript("""
CREATE TABLE IF NOT EXISTS "institution_licenses" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name_ua TEXT NOT NULL,
    short_name_ua TEXT,
    name_en TEXT,
    short_name_en TEXT,
    is_active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS "license_transfer_orders" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number TEXT NOT NULL,
    order_date TEXT NOT NULL,
    scan_file TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS "license_transfer_order_students" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER NOT NULL,
    student_id INTEGER NOT NULL,
    previous_license_id INTEGER,
    new_license_id INTEGER NOT NULL,
    FOREIGN KEY (order_id) REFERENCES license_transfer_orders(id),
    FOREIGN KEY (student_id) REFERENCES students(id),
    FOREIGN KEY (previous_license_id) REFERENCES institution_licenses(id),
    FOREIGN KEY (new_license_id) REFERENCES institution_licenses(id)
);
""")

cur.execute(
    "INSERT OR IGNORE INTO institution_licenses (id, name_ua, short_name_ua, name_en, short_name_en, is_active) "
    "VALUES (1, ?, 'Київська', ?, 'Kyiv', 1)",
    ('Приватний вищий навчальний заклад «Європейський університет». Приватна форма власності. Міністерство освіти і науки України. Ліцензія серія ВО № 00228-022801 від 15/05/2017.',
     "Private Higher Educational Institution 'European University'. Private. Ministry of Education and Science of Ukraine. License series BO № 00228-022801 dated 15/05/2017.")
)
cur.execute(
    "INSERT OR IGNORE INTO institution_licenses (id, name_ua, short_name_ua, name_en, short_name_en, is_active) "
    "VALUES (2, ?, 'Львівська', ?, 'Lviv', 1)",
    ('Львівська філія Приватного вищого навчального закладу «Європейський університет». Приватна форма власності. Міністерство освіти і науки України. Ліцензія серія ВО № 00228-022801 від 15/05/2017.',
     'Lviv Branch of Private Higher Education Establishment «European University». Private. Ministry of  Education and  Science of Ukraine. License series ВO № 00228-022801 from 15/05/2017.')
)

try:
    cur.execute("ALTER TABLE students ADD COLUMN license_id INTEGER REFERENCES institution_licenses(id)")
    print("Додано колонку students.license_id")
except sqlite3.OperationalError as e:
    if "duplicate column" in str(e):
        print("Колонка students.license_id вже існує - пропускаю")
    else:
        raise

# Перенесення даних: за поточною групою студента визначаємо ліцензію
licenses_by_text = {
    row['name_ua']: row['id']
    for row in cur.execute("SELECT id, name_ua FROM institution_licenses")
}

updated = 0
skipped = 0
rows = cur.execute("""
    SELECT s.id AS student_id, g.institution_name_and_status AS inst_text
    FROM students s LEFT JOIN groups g ON s.group_id = g.id
    WHERE s.license_id IS NULL
""").fetchall()
for row in rows:
    license_id = licenses_by_text.get(row['inst_text'])
    if license_id:
        cur.execute("UPDATE students SET license_id = ? WHERE id = ?", (license_id, row['student_id']))
        updated += 1
    else:
        skipped += 1

conn.commit()
print(f"Перенесено ліцензію для {updated} студент(ів).")
if skipped:
    print(f"УВАГА: {skipped} студент(ів) лишились без ліцензії (без групи або "
          f"незнайомий текст закладу) - виставте вручну через картку студента.")

n_licenses = cur.execute("SELECT COUNT(*) FROM institution_licenses").fetchone()[0]
print(f"\nГотово: {n_licenses} ліцензій у каталозі.")
conn.close()
