"""
routes/init_db.py
==================
Одноразовий скрипт для створення бази даних students.db "з нуля" разом
з усіма таблицями та користувачем admin/admin123 за замовчуванням.

Запуск: `python init_db.py` з кореня проєкту (або routes/, шлях до БД
обчислюється відносно розташування цього файлу).

Використовує CREATE TABLE IF NOT EXISTS, тож повторний запуск на вже
існуючій базі не видаляє наявні дані, а лише додає відсутні таблиці.
УВАГА: якщо в майбутньому додається нова колонка до вже існуючої
таблиці (як це один раз уже сталося з таблицею "groups" - див.
коментар нижче), для існуючих баз даних CREATE TABLE IF NOT EXISTS
її не додасть - потрібна окрема ALTER TABLE-міграція.
"""

import sqlite3
import logging
import os
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, 'students.db')

conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

cur.executescript("""
CREATE TABLE IF NOT EXISTS "activity_grades" (
	"id"	INTEGER,
	"student_id"	INTEGER NOT NULL,
	"entity_id"	INTEGER NOT NULL,
	"entity_type"	TEXT NOT NULL CHECK("entity_type" IN ('practice', 'coursework', 'attestation')),
	"grade"	INTEGER,
	"name"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("student_id") REFERENCES "students"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "attestations" (
	"id"	INTEGER,
	"code"	TEXT NOT NULL,
	"name"	TEXT NOT NULL,
	"credits"	INTEGER NOT NULL,
	"type"	TEXT NOT NULL CHECK("type" IN ('Залік', 'Екзамен')),
	"position"	INTEGER NOT NULL,
	"group_id"	INTEGER NOT NULL,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
CREATE TABLE IF NOT EXISTS "courseworks" (
	"id"	INTEGER,
	"code"	TEXT NOT NULL,
	"name"	TEXT NOT NULL,
	"credits"	INTEGER NOT NULL,
	"type"	TEXT NOT NULL CHECK("type" IN ('Залік', 'Екзамен')),
	"position"	INTEGER NOT NULL,
	"group_id"	INTEGER NOT NULL,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
CREATE TABLE IF NOT EXISTS "education_documents" (
	"id"	INTEGER,
	"student_id"	INTEGER NOT NULL,
	"document_type"	TEXT NOT NULL,
	"document_type_en"	TEXT NOT NULL,
	"document_number"	TEXT NOT NULL,
	"institution_name"	TEXT NOT NULL,
	"institution_name_en"	TEXT NOT NULL,
	"country"	TEXT NOT NULL,
	"country_en"	TEXT NOT NULL,
	"completion_date"	TEXT NOT NULL,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("student_id") REFERENCES "students"("id")
);
CREATE TABLE IF NOT EXISTS "foreign_education_docs" (
	"id"	INTEGER,
	"education_doc_id"	INTEGER NOT NULL,
	"reference_number"	TEXT,
	"reference_institution"	TEXT,
	"reference_institution_en"	TEXT,
	"reference_country"	TEXT,
	"reference_country_en"	TEXT,
	"reference_issue_date"	TEXT,
	"recognition_certificate_number"	TEXT,
	"recognition_issuer"	TEXT,
	"recognition_issuer_en"	TEXT,
	"recognition_date"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("education_doc_id") REFERENCES "education_documents"("id")
);
CREATE TABLE IF NOT EXISTS "grades" (
	"id"	INTEGER,
	"student_id"	INTEGER,
	"subject_id"	INTEGER,
	"grade"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("subject_id") REFERENCES "subjects"("id"),
	FOREIGN KEY("student_id") REFERENCES "students"("id")
);
CREATE TABLE IF NOT EXISTS "groups" (
	"id"	INTEGER,
	"name"	TEXT NOT NULL,
	"start_year"	INTEGER NOT NULL,
	"study_form"	TEXT NOT NULL CHECK("study_form" IN ('Денна', 'Заочна')),
	"program_credits"	INTEGER NOT NULL,
	"degree_level"	TEXT NOT NULL,
	"degree_level_en"	TEXT NOT NULL,
	"knowledge_area"	TEXT NOT NULL,
	"knowledge_area_en"	TEXT NOT NULL,
	"specialty"	TEXT NOT NULL,
	"specialty_en"	TEXT NOT NULL,
	"educational_program"	TEXT NOT NULL,
	"educational_program_en"	TEXT NOT NULL,
	"qualification_name"	TEXT NOT NULL,
	"qualification_name_en"	TEXT NOT NULL,
	"course"	INTEGER NOT NULL DEFAULT 1,
	"institution_name_and_status"	TEXT,
	"institution_name_and_status_en"	TEXT,
	"entry_requirements"	TEXT,
	"entry_requirements_en"	TEXT,
	"learning_outcomes"	TEXT,
	"learning_outcomes_en"	TEXT,
	"program_includes"	TEXT,
	"program_includes_en"	TEXT,
	"archived"	BOOLEAN DEFAULT FALSE,
	PRIMARY KEY("id" AUTOINCREMENT),
	UNIQUE("name","start_year")
);
CREATE TABLE IF NOT EXISTS "military" (
	"id"	INTEGER,
	"student_id"	INTEGER,
	"registration_number_of_the_DRPVR"	TEXT,
	"military_registration_document"	TEXT,
	"issued_VOD"	TEXT,
	"military_accounting_specialty_number"	TEXT,
	"military_rank"	TEXT,
	"change_credentials"	TEXT,
	"reason_for_changing_credentials"	TEXT,
	"being_on_military_registration"	TEXT,
	"address_of_residence"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("student_id") REFERENCES "students"("id")
);
CREATE TABLE IF NOT EXISTS "practices" (
	"id"	INTEGER,
	"code"	TEXT NOT NULL,
	"name"	TEXT NOT NULL,
	"credits"	INTEGER NOT NULL,
	"type"	TEXT NOT NULL CHECK("type" IN ('Залік', 'Екзамен')),
	"position"	INTEGER NOT NULL,
	"group_id"	INTEGER NOT NULL,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
CREATE TABLE IF NOT EXISTS "students" (
	"id"	INTEGER,
	"last_name_UA"	TEXT,
	"first_name_UA"	TEXT,
	"middle_name_UA"	TEXT,
	"last_name_ENG"	TEXT,
	"first_name_ENG"	TEXT,
	"birth_date"	TEXT,
	"group_id"	INTEGER,
	"edebo_code"	VARCHAR(50),
	"archived"	BOOLEAN DEFAULT FALSE,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
CREATE TABLE IF NOT EXISTS "subjects" (
	"id"	INTEGER,
	"code"	TEXT,
	"name"	TEXT NOT NULL,
	"credits"	INTEGER,
	"group_id"	INTEGER,
	"position"	INTEGER DEFAULT 0,
	"type"	TEXT DEFAULT 'Залік' CHECK("type" IN ('Залік', 'Екзамен')),
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
CREATE TABLE IF NOT EXISTS "user_groups" (
	"user_id"	INTEGER,
	"group_id"	INTEGER,
	PRIMARY KEY("user_id","group_id"),
	FOREIGN KEY("group_id") REFERENCES "groups"("id") ON DELETE CASCADE,
	FOREIGN KEY("user_id") REFERENCES "users"("id") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "users" (
	"id"	INTEGER,
	"username"	TEXT UNIQUE,
	"password_hash"	TEXT,
	"role"	TEXT NOT NULL CHECK("role" IN ('admin', 'user')),
	"is_admin"	INTEGER DEFAULT 0,
	"permissions"	TEXT DEFAULT '[]',
	PRIMARY KEY("id" AUTOINCREMENT)
);
CREATE TABLE accreditations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    degree TEXT NOT NULL,          
    specialty TEXT NOT NULL,       
    text_ua TEXT,
    text_en TEXT
);
CREATE TABLE diplomas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL,
    diploma_number TEXT,
    appendix_number TEXT,
    FOREIGN KEY (student_id) REFERENCES students(id)
);
CREATE TABLE student_study_periods (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    filiya TEXT NOT NULL,             
    filiya_en TEXT NOT NULL,             
    group_name TEXT,                  
    start_date TEXT,                   
    end_date TEXT,                    
    period_order INTEGER NOT NULL DEFAULT 0,  
    note TEXT                         
);

CREATE INDEX idx_study_periods_student ON student_study_periods(student_id);

CREATE TABLE IF NOT EXISTS "document_templates" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL UNIQUE,      -- ім'я файлу в папці template_word/, напр. "template_tck.docx"
    display_name TEXT,                  -- людська назва для випадаючого списку (необов'язково)
    description TEXT,                   -- опис призначення шаблону (необов'язково)
    admin_only INTEGER NOT NULL DEFAULT 0,  -- 1 = бачать і можуть обрати лише адміністратори
    hidden INTEGER NOT NULL DEFAULT 0,      -- 1 = шаблон не показується в списках вибору при генерації (для всіх, включно з адмінами)
    uploaded_by TEXT,
    uploaded_at TEXT DEFAULT (datetime('now'))
);
""")

# Додаємо користувачів.
# ВИПРАВЛЕНО: тут раніше було
#     for u, p, r, g in [('admin', 'admin123', 'admin', '1', '[]')]:
# - кортеж мав 5 елементів (логін, пароль, роль, is_admin, permissions),
# а розпаковувався лише в 4 змінні (u, p, r, g). Це кидало
# `ValueError: too many values to unpack` і скрипт init_db.py падав
# щоразу при спробі його запустити, ще ДО створення користувача-адміна.
# Крім того, INSERT нижче мав 5 плейсхолдерів (?, ?, ?, ?, ?), але
# отримував лише 4 значення (u, hash, r, g) - це також викликало б
# помилку sqlite3 навіть після виправлення розпакування кортежу.
DEFAULT_USERS = [
    # (username, password, role, is_admin, permissions_json)
    ('admin', 'admin123', 'admin', 1, '[]'),
]

for username, password, role, is_admin, permissions in DEFAULT_USERS:
    try:
        cur.execute(
            "INSERT OR IGNORE INTO users (username, password_hash, role, is_admin, permissions) VALUES (?, ?, ?, ?, ?)",
            (username, generate_password_hash(password), role, is_admin, permissions)
        )
    except sqlite3.Error as e:
        logging.error(f"Не вдалося створити користувача за замовчуванням '{username}': {e}")

conn.commit()
conn.close()
print("✅ DB та користувачі створені.")
print("⚠️  УВАГА: користувач 'admin' створений зі стандартним паролем 'admin123'. "
      "Обов'язково змініть його одразу після першого входу через "
      "/admin/users/<id>/change-password.")