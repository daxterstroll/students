# -*- coding: utf-8 -*-
"""
Міграція таблиць для процесів "Переведення на курс" / "Заморожені
студенти" / "Відрахування". Безпечна для запуску на вже наявній базі -
не чіпає students/groups/grades тощо, лише додає 5 нових таблиць.

Запуск:  python migrate_course_transfer.py
"""
import sqlite3
import os

db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'students.db')
conn = sqlite3.connect(db_path)
cur = conn.cursor()

cur.executescript("""
CREATE TABLE IF NOT EXISTS "course_transfer_orders" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number TEXT NOT NULL,
    order_date TEXT NOT NULL,
    scan_file TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS "course_transfer_order_groups" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER NOT NULL,
    group_id INTEGER NOT NULL,
    course_from INTEGER NOT NULL,
    course_to INTEGER NOT NULL,
    FOREIGN KEY (order_id) REFERENCES course_transfer_orders(id),
    FOREIGN KEY (group_id) REFERENCES groups(id)
);

CREATE TABLE IF NOT EXISTS "frozen_students" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL,
    previous_group_id INTEGER,
    order_id INTEGER,
    reason TEXT NOT NULL,
    frozen_at TEXT DEFAULT (datetime('now')),
    frozen_by TEXT,
    resolved_at TEXT,
    resolution TEXT,
    resolved_by TEXT,
    FOREIGN KEY (student_id) REFERENCES students(id),
    FOREIGN KEY (previous_group_id) REFERENCES groups(id),
    FOREIGN KEY (order_id) REFERENCES course_transfer_orders(id)
);

CREATE TABLE IF NOT EXISTS "expulsion_orders" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number TEXT NOT NULL,
    order_date TEXT NOT NULL,
    scan_file TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS "expulsion_order_students" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER NOT NULL,
    student_id INTEGER NOT NULL,
    previous_group_id INTEGER,
    reason TEXT NOT NULL,
    FOREIGN KEY (order_id) REFERENCES expulsion_orders(id),
    FOREIGN KEY (student_id) REFERENCES students(id),
    FOREIGN KEY (previous_group_id) REFERENCES groups(id)
);
""")

conn.commit()
tables = ['course_transfer_orders', 'course_transfer_order_groups', 'frozen_students',
          'expulsion_orders', 'expulsion_order_students']
for t in tables:
    n = cur.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
    print(f"{t}: {n} записів")
print("\nГотово.")
conn.close()