# -*- coding: utf-8 -*-
"""
Міграція: універсальна таблиця вкладень (кілька файлів там, де раніше
можна було прикріпити лише один) - наказ про переведення на курс,
наказ про відрахування, підстава заморозки студента.

Створює таблицю "attachments" (entity_type + entity_id - до чого
прикріплено; наявні одиничні поля scan_file/document_file НЕ
видаляються і НЕ чіпаються - лишаються як є, для сумісності зі старими
записами, зробленими до цієї міграції).

Безпечна для повторного запуску. Запуск: python migrate_add_attachments.py
"""
import sqlite3
import os

db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'students.db')
conn = sqlite3.connect(db_path)
cur = conn.cursor()

cur.execute("""
    CREATE TABLE IF NOT EXISTS "attachments" (
        "id" INTEGER PRIMARY KEY AUTOINCREMENT,
        "entity_type" TEXT NOT NULL,   -- 'course_transfer_order' | 'expulsion_order' | 'frozen_student'
        "entity_id" INTEGER NOT NULL,
        "file_path" TEXT NOT NULL,
        "original_name" TEXT,
        "uploaded_at" TEXT DEFAULT (datetime('now','localtime')),
        "uploaded_by" TEXT
    )
""")
cur.execute("""
    CREATE INDEX IF NOT EXISTS "idx_attachments_entity" ON "attachments" ("entity_type", "entity_id")
""")

conn.commit()
conn.close()
print("Готово: таблиця attachments створена (або вже існувала).")
