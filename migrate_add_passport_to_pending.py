# -*- coding: utf-8 -*-
"""
Міграція: паспортні дані (книжка або ID-картка) у самій публічній
анкеті /apply, так само як уже є документ про освіту - зберігаються
поки що в pending_students (окремими полями з префіксом passport_),
адмін при підтвердженні заявки переносить їх у окрему таблицю
passport_documents (якщо вона ще не створена - див.
migrate_add_passport_documents.py, запустіть його ПЕРШИМ).

Запуск: python migrate_add_passport_to_pending.py
"""
import sqlite3
import os

db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'students.db')
conn = sqlite3.connect(db_path)
cur = conn.cursor()

NEW_COLUMNS = [
    ("passport_document_type", "TEXT"),
    ("passport_series", "TEXT"),
    ("passport_number", "TEXT"),
    ("passport_issued_by", "TEXT"),
    ("passport_issue_date", "TEXT"),
    ("passport_valid_until", "TEXT"),
    ("passport_unique_number", "TEXT"),
]

cur.execute("PRAGMA table_info(pending_students)")
existing = {row[1] for row in cur.fetchall()}

for col, col_type in NEW_COLUMNS:
    if col not in existing:
        cur.execute(f"ALTER TABLE pending_students ADD COLUMN {col} {col_type}")
        print(f"Додано pending_students.{col}")
    else:
        print(f"pending_students.{col} вже існує - пропущено")

conn.commit()
conn.close()
print("Готово.")
