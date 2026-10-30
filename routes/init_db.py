"""
routes/init_db.py
==================
Одноразовий скрипт для створення БД students.db "з нуля" з усіма
таблицями, базовими довідниками та користувачем admin/admin123.

Також ІДЕМПОТЕНТНО доводить до актуального стану вже існуючу базу:
усі колонки й таблиці, які раніше додавались окремими міграційними
скриптами, тепер створюються/додаються тут автоматично.

Файл ЗАМІНЮЄ собою:
  - routes/init_db.py (початкове створення БД)
  - migrate_military_reports.py     (v1: gender, category_override, tck_offices, log)
  - migrate_military_reports_v2.py  (поля «Резерв+», military_history)
  - migrate_military_reports_v3.py  (Постанова №812: note_18, dismissed_at, is_base)

Запуск: `python init_db.py` з кореня проєкту (шлях до БД обчислюється
відносно розташування файлу - не залежить від CWD).

Повторний запуск БЕЗПЕЧНИЙ: наявні дані не видаляються, лише
додаються відсутні колонки й таблиці. Якщо БД уже існує та має дані -
спершу створюється резервна копія у тій самій папці.
"""

import os
import shutil
import sqlite3
from datetime import datetime

from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, 'students.db')

# ═══════════════════════════════════════════════════════════════════════
# Резервна копія (лише якщо БД уже існує і має реальний вміст)
# ═══════════════════════════════════════════════════════════════════════
if os.path.exists(DB_PATH) and os.path.getsize(DB_PATH) > 4096:
    stamp = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    backup_path = f"{DB_PATH}.bak-{stamp}"
    shutil.copy2(DB_PATH, backup_path)
    print(f"💾 Бекап: {os.path.basename(backup_path)}")

conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()


# ═══════════════════════════════════════════════════════════════════════
# 1. СХЕМА — CREATE TABLE IF NOT EXISTS
# ═══════════════════════════════════════════════════════════════════════
# Усі нові колонки одразу вписані в CREATE TABLE - щоб на чистій
# інсталяції одразу отримувати актуальну схему. Для вже існуючих
# баз колонки додаються нижче через _add_column_if_missing().
# ═══════════════════════════════════════════════════════════════════════
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
	"full_program_only"	INTEGER NOT NULL DEFAULT 0,
	"reduced_credits"	INTEGER,
	"reduced_type"	TEXT,
	"reduced_only"	INTEGER NOT NULL DEFAULT 0,
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
	"full_program_only"	INTEGER NOT NULL DEFAULT 0,
	"reduced_credits"	INTEGER,
	"reduced_type"	TEXT,
	"reduced_only"	INTEGER NOT NULL DEFAULT 0,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
CREATE TABLE IF NOT EXISTS "passport_documents" (
	"id"	INTEGER,
	"student_id"	INTEGER NOT NULL,
	"document_type"	TEXT NOT NULL CHECK("document_type" IN (
		'Свідоцтво про народження',
		'Паспорт громадянина України',
		'Паспорт громадянина України з безконтактним електронним носієм',
		'Паспорт громадянина України, що не містить безконтактного електронного носія'
	)),
	"series"	TEXT,
	"number"	TEXT NOT NULL,
	"issued_by"	TEXT,
	"issue_date"	TEXT,
	"valid_until"	TEXT,
	"unique_number"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("student_id") REFERENCES "students"("id")
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
	"isced_f_ua"	TEXT,
	"isced_f_en"	TEXT,
	"specialty_code"	TEXT,
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
	"entry_requirements_reduced"	TEXT,
	"entry_requirements_reduced_en"	TEXT,
	"program_includes_reduced"	TEXT,
	"program_includes_reduced_en"	TEXT,
	"learning_outcomes_reduced"	TEXT,
	"learning_outcomes_reduced_en"	TEXT,
	"archived"	BOOLEAN DEFAULT FALSE,
	PRIMARY KEY("id" AUTOINCREMENT),
	UNIQUE("name","start_year","study_form","program_credits")
);
-- ═══════════════════════════════════════════════════════════════════
-- military — ПОВНА структура «Резерв+» (оригінальні + v1/v2/v3)
-- ═══════════════════════════════════════════════════════════════════
CREATE TABLE IF NOT EXISTS "military" (
	"id"	INTEGER,
	"student_id"	INTEGER,
	-- Оригінальні легасі-поля (ручне заповнення)
	"registration_number_of_the_DRPVR"	TEXT,
	"military_registration_document"	TEXT,
	"issued_VOD"	TEXT,
	"military_accounting_specialty_number"	TEXT,
	"military_rank"	TEXT,
	"change_credentials"	TEXT,
	"reason_for_changing_credentials"	TEXT,
	"being_on_military_registration"	TEXT,
	"address_of_residence"	TEXT,
	-- v1: явне перевизначення категорії
	"category_override"	TEXT CHECK("category_override" IN ('officer', 'rank')),
	-- v2: поля з PDF «Резерв+»
	"status_badge"	TEXT,
	"valid_until"	TEXT,
	"registration_category"	TEXT,
	"removal_basis"	TEXT,
	"tck_name"	TEXT,
	"registry_number"	TEXT,
	"vos_code"	TEXT,
	"vos_description"	TEXT,
	"deferral_type"	TEXT,
	"deferral_until"	TEXT,
	"police_reason"	TEXT,
	"police_date"	TEXT,
	"vlk_decision"	TEXT,
	"vlk_date"	TEXT,
	"disability_group"	TEXT,
	"disability_valid_until"	TEXT,
	"disability_reason"	TEXT,
	"formed_at"	TEXT,
	"data_verified_on_time"	INTEGER DEFAULT 0,
	"data_update_date"	TEXT,
	"phone"	TEXT,
	"email"	TEXT,
	-- v3: Постанова КМУ №812 від 10.06.2026 (графа 18)
	"note_18"	TEXT,
	"mobilized_date"	TEXT,
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
	"full_program_only"	INTEGER NOT NULL DEFAULT 0,
	"reduced_credits"	INTEGER,
	"reduced_type"	TEXT,
	"reduced_only"	INTEGER NOT NULL DEFAULT 0,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
-- ═══════════════════════════════════════════════════════════════════
-- students — з gender (v1), address_of_residence (v2), dismissed_at (v3)
-- ═══════════════════════════════════════════════════════════════════
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
	"tax_id"	TEXT,
	"photo"	TEXT,
	"license_id"	INTEGER,
	"archived"	BOOLEAN DEFAULT FALSE,
	"program_credits_override"	INTEGER,
	"phone"	TEXT,
	"phone_backup"	TEXT,
	"email"	TEXT,
	"address_of_residence"	TEXT,
	"gender"	TEXT CHECK("gender" IN ('M', 'F')),
	"dismissed_at"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT),
	FOREIGN KEY("group_id") REFERENCES "groups"("id"),
	FOREIGN KEY("license_id") REFERENCES "institution_licenses"("id")
);
CREATE TABLE IF NOT EXISTS "subjects" (
	"id"	INTEGER,
	"code"	TEXT,
	"name"	TEXT NOT NULL,
	"credits"	INTEGER,
	"group_id"	INTEGER,
	"position"	INTEGER DEFAULT 0,
	"type"	TEXT DEFAULT 'Залік' CHECK("type" IN ('Залік', 'Екзамен')),
	"full_program_only"	INTEGER NOT NULL DEFAULT 0,
	"reduced_credits"	INTEGER,
	"reduced_type"	TEXT,
	"reduced_only"	INTEGER NOT NULL DEFAULT 0,
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
CREATE TABLE IF NOT EXISTS accreditations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    degree TEXT NOT NULL,
    specialty TEXT NOT NULL,
    text_ua TEXT,
    text_en TEXT
);
CREATE TABLE IF NOT EXISTS diplomas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL,
    diploma_number TEXT,
    appendix_number TEXT,
    FOREIGN KEY (student_id) REFERENCES students(id)
);
CREATE TABLE IF NOT EXISTS student_study_periods (
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
CREATE INDEX IF NOT EXISTS idx_study_periods_student ON student_study_periods(student_id);

CREATE TABLE IF NOT EXISTS "document_templates" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL UNIQUE,
    display_name TEXT,
    description TEXT,
    admin_only INTEGER NOT NULL DEFAULT 0,
    hidden INTEGER NOT NULL DEFAULT 0,
    uploaded_by TEXT,
    uploaded_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS "degree_levels" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name_ua TEXT NOT NULL UNIQUE,
    name_en TEXT,
    is_active INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS "knowledge_fields" (
    code TEXT PRIMARY KEY,
    name_ua TEXT NOT NULL,
    name_en TEXT
);
CREATE TABLE IF NOT EXISTS "specialties" (
    code TEXT PRIMARY KEY,
    name_ua TEXT NOT NULL,
    short_name TEXT,
    name_en TEXT,
    isced_f_ua TEXT,
    isced_f_en TEXT,
    knowledge_field_code TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1,
    is_custom INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (knowledge_field_code) REFERENCES knowledge_fields(code)
);
CREATE TABLE IF NOT EXISTS "educational_programs" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    specialty_code TEXT NOT NULL,
    start_year INTEGER NOT NULL,
    degree_level TEXT NOT NULL,
    name_ua TEXT NOT NULL,
    name_en TEXT,
    UNIQUE (specialty_code, start_year, degree_level),
    FOREIGN KEY (specialty_code) REFERENCES specialties(code)
);
CREATE TABLE IF NOT EXISTS "qualification_names" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    specialty_code TEXT NOT NULL,
    start_year INTEGER NOT NULL,
    degree_level TEXT NOT NULL,
    name_ua TEXT NOT NULL,
    name_en TEXT,
    UNIQUE (specialty_code, start_year, degree_level),
    FOREIGN KEY (specialty_code) REFERENCES specialties(code)
);
CREATE TABLE IF NOT EXISTS "update_requests" (
    "id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "token" TEXT NOT NULL UNIQUE,
    "student_id" INTEGER NOT NULL,
    "allowed_fields" TEXT NOT NULL,
    "created_at" TEXT DEFAULT (datetime('now','localtime')),
    "created_by" TEXT,
    "expires_at" TEXT NOT NULL,
    "status" TEXT NOT NULL DEFAULT 'pending' CHECK("status" IN ('pending','submitted','approved','rejected')),
    "submitted_data" TEXT,
    "photo_path" TEXT,
    "submitted_at" TEXT,
    "reviewed_by" TEXT,
    "reviewed_at" TEXT,
    "review_note" TEXT,
    "applied_fields" TEXT,
    FOREIGN KEY("student_id") REFERENCES "students"("id") ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_update_requests_token ON update_requests(token);
CREATE INDEX IF NOT EXISTS idx_update_requests_student ON update_requests(student_id);

CREATE TABLE IF NOT EXISTS "pending_students" (
    "id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "created_at" TEXT DEFAULT (datetime('now','localtime')),
    "submitted_ip" TEXT,
    "photo_path" TEXT,
    "last_name_UA" TEXT NOT NULL,
    "first_name_UA" TEXT NOT NULL,
    "middle_name_UA" TEXT,
    "last_name_ENG" TEXT,
    "first_name_ENG" TEXT,
    "birth_date" TEXT,
    "phone" TEXT,
    "tax_id" TEXT,
    "phone_backup" TEXT,
    "email" TEXT,
    "document_type" TEXT,
    "document_series" TEXT,
    "document_number" TEXT,
    "document_institution" TEXT,
    "document_country" TEXT,
    "document_date" TEXT,
    "is_foreign_document" INTEGER NOT NULL DEFAULT 0,
    "foreign_reference_number" TEXT,
    "foreign_reference_institution" TEXT,
    "foreign_reference_country" TEXT,
    "foreign_reference_issue_date" TEXT,
    "recognition_certificate_number" TEXT,
    "recognition_issuer" TEXT,
    "recognition_date" TEXT,
    "reduced_program_claim" INTEGER NOT NULL DEFAULT 0,
    "reduced_program_specialty" TEXT,
    "reduced_program_institution" TEXT,
    "passport_document_type" TEXT,
    "passport_series" TEXT,
    "passport_number" TEXT,
    "passport_issued_by" TEXT,
    "passport_issue_date" TEXT,
    "passport_valid_until" TEXT,
    "passport_unique_number" TEXT,
    "military_registration_number_drpvr" TEXT,
    "military_registration_document" TEXT,
    "military_issued_vod" TEXT,
    "military_specialty_number" TEXT,
    "military_rank" TEXT,
    "military_being_registered" INTEGER NOT NULL DEFAULT 0,
    "military_address" TEXT,
    "military_change_credentials" TEXT,
    "military_change_reason" TEXT,
    "status" TEXT NOT NULL DEFAULT 'new' CHECK("status" IN ('new', 'approved', 'rejected')),
    "reviewed_by" TEXT,
    "reviewed_at" TEXT,
    "review_note" TEXT,
    "resulting_student_id" INTEGER,
    FOREIGN KEY("resulting_student_id") REFERENCES "students"("id") ON DELETE SET NULL
);
CREATE INDEX IF NOT EXISTS idx_pending_students_status ON pending_students(status);

CREATE TABLE IF NOT EXISTS "attachments" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_type TEXT NOT NULL,
    entity_id INTEGER NOT NULL,
    file_path TEXT NOT NULL,
    original_name TEXT,
    uploaded_at TEXT DEFAULT (datetime('now','localtime')),
    uploaded_by TEXT
);
CREATE INDEX IF NOT EXISTS "idx_attachments_entity" ON "attachments" ("entity_type", "entity_id");

CREATE TABLE IF NOT EXISTS "course_transfer_orders" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number TEXT NOT NULL,
    order_date TEXT NOT NULL,
    scan_file TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
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
    document_file TEXT,
    frozen_at TEXT DEFAULT (datetime('now', 'localtime')),
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
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
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
-- ═══════════════════════════════════════════════════════════════════
-- institution_licenses — з is_base (v3, Постанова №812)
-- ═══════════════════════════════════════════════════════════════════
CREATE TABLE IF NOT EXISTS "institution_licenses" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name_ua TEXT NOT NULL,
    short_name_ua TEXT,
    name_en TEXT,
    short_name_en TEXT,
    is_active INTEGER NOT NULL DEFAULT 1,
    is_base INTEGER NOT NULL DEFAULT 0
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
-- ═══════════════════════════════════════════════════════════════════
-- НОВІ таблиці (з міграцій military_reports)
-- ═══════════════════════════════════════════════════════════════════
-- v1: довідник ТЦК
CREATE TABLE IF NOT EXISTS "tck_offices" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name_ua TEXT NOT NULL UNIQUE,
    address TEXT,
    is_active INTEGER NOT NULL DEFAULT 1
);
-- v1: журнал генерацій звітів
CREATE TABLE IF NOT EXISTS "military_reports_log" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_type TEXT NOT NULL,
    report_date TEXT NOT NULL,
    generated_at TEXT DEFAULT (datetime('now','localtime')),
    generated_by TEXT,
    rows_count INTEGER,
    filename TEXT
);
CREATE INDEX IF NOT EXISTS idx_military_reports_log_type_date
    ON military_reports_log (report_type, report_date);
-- v2: історія змін військових даних (і полів картки студента)
CREATE TABLE IF NOT EXISTS "military_history" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL,
    changed_at TEXT DEFAULT (datetime('now','localtime')),
    changed_by TEXT,
    source TEXT,           -- 'manual' | 'pdf_import'
    scope TEXT,            -- 'military' | 'student'
    field_name TEXT NOT NULL,
    old_value TEXT,
    new_value TEXT,
    FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_military_history_student
    ON military_history (student_id, id);
""")

# ═══════════════════════════════════════════════════════════════════════
# 2. ДАНІ — ліцензії, ступені, ТЦК, спеціальності, користувачі
# ═══════════════════════════════════════════════════════════════════════

# ── Ліцензії. v3: Львівська філія — базова (is_base=1) ──
cur.executescript("""
INSERT OR IGNORE INTO institution_licenses
    (id, name_ua, short_name_ua, name_en, short_name_en, is_active, is_base)
VALUES (
    1,
    'Приватний вищий навчальний заклад «Європейський університет». Приватна форма власності. Міністерство освіти і науки України. Ліцензія серія ВО № 00228-022801 від 15/05/2017.',
    'Київська',
    'Private Higher Educational Institution ''European University''. Private. Ministry of Education and Science of Ukraine. License series BO № 00228-022801 dated 15/05/2017.',
    'Kyiv',
    1, 0
);
INSERT OR IGNORE INTO institution_licenses
    (id, name_ua, short_name_ua, name_en, short_name_en, is_active, is_base)
VALUES (
    2,
    'Львівська філія Приватного вищого навчального закладу «Європейський університет». Приватна форма власності. Міністерство освіти і науки України. Ліцензія серія ВО № 00228-022801 від 15/05/2017.',
    'Львівська',
    'Lviv Branch of Private Higher Education Establishment «European University». Private. Ministry of  Education and  Science of Ukraine. License series ВO № 00228-022801 from 15/05/2017.',
    'Lviv',
    1, 1
);

-- Якщо існуюча БД: жодна ліцензія ще не позначена як базова - ставимо Львівську
UPDATE institution_licenses SET is_base = 1
WHERE short_name_ua = 'Львівська'
  AND (SELECT COUNT(*) FROM institution_licenses WHERE is_base = 1) = 0;
""")

# ── Ступені ──
cur.executescript("""
INSERT OR IGNORE INTO degree_levels (name_ua, name_en, is_active) VALUES ('Молодший бакалавр', 'Junior Bachelor', 0);
INSERT OR IGNORE INTO degree_levels (name_ua, name_en, is_active) VALUES ('Бакалавр', 'Bachelor', 1);
INSERT OR IGNORE INTO degree_levels (name_ua, name_en, is_active) VALUES ('Магістр', 'Master Degree', 1);
INSERT OR IGNORE INTO degree_levels (name_ua, name_en, is_active) VALUES ('Доктор філософії', 'Doctor of Philosophy', 0);
""")

# ── 67 ТЦК (зі migrate_military_reports.py v1) ──
TCK_OFFICES = [
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
for name, addr in TCK_OFFICES:
    cur.execute(
        "INSERT OR IGNORE INTO tck_offices (name_ua, address) VALUES (?, ?)",
        (name, addr)
    )
print(f"✅ tck_offices: {cur.execute('SELECT COUNT(*) FROM tck_offices').fetchone()[0]} установ")

# ── Галузі знань + спеціальності (Постанова КМУ №1021 від 30.08.2024) ──
# ⚠️ ЦЕЙ БЛОК НЕЗМІННИЙ - скопіюйте як є зі свого поточного
#    init_db.py (розділ з INSERT INTO knowledge_fields та
#    INSERT INTO specialties). Він не стосується міграцій
#    military_reports і не потребує редагування.
#
#    Нижче - лише маркер, щоб ви не забули його вставити:
#
#    cur.executescript("""
#        INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('A', 'Освіта', 'Education');
#        ...
#        INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A1', '...', '...', 'A', 1);
#        ...
#    """)
# ───────────────────────────────────────────────────────────────────────

# ── Користувач admin за замовчуванням ──
DEFAULT_USERS = [
    ('admin', 'admin123', 'admin', 1, '[]'),
]
for username, password, role, is_admin, permissions in DEFAULT_USERS:
    try:
        cur.execute(
            "INSERT OR IGNORE INTO users (username, password_hash, role, is_admin, permissions) "
            "VALUES (?, ?, ?, ?, ?)",
            (username, generate_password_hash(password), role, is_admin, permissions)
        )
    except sqlite3.Error as e:
        print(f"⚠️ Не вдалося створити '{username}': {e}")


# ═══════════════════════════════════════════════════════════════════════
# 3. МІГРАЦІЇ ДЛЯ ВЖЕ ІСНУЮЧИХ БАЗ
# ═══════════════════════════════════════════════════════════════════════
# CREATE TABLE IF NOT EXISTS не додає колонки в уже існуючу таблицю,
# тому перевіряємо через PRAGMA table_info і додаємо лише відсутні.
# ═══════════════════════════════════════════════════════════════════════
def _add_column_if_missing(table, column, coltype):
    existing = {row[1] for row in cur.execute(f'PRAGMA table_info("{table}")')}
    if column not in existing:
        cur.execute(f'ALTER TABLE "{table}" ADD COLUMN "{column}" {coltype}')
        print(f"➕ {table}.{column}")


print("\n── Міграція існуючої БД (якщо є) ──")

# students
_add_column_if_missing('students', 'gender', "TEXT CHECK(gender IN ('M', 'F'))")
_add_column_if_missing('students', 'address_of_residence', 'TEXT')
_add_column_if_missing('students', 'dismissed_at', 'TEXT')

# military
_add_column_if_missing('military', 'category_override',
                       "TEXT CHECK(category_override IN ('officer', 'rank'))")
_add_column_if_missing('military', 'status_badge', 'TEXT')
_add_column_if_missing('military', 'valid_until', 'TEXT')
_add_column_if_missing('military', 'registration_category', 'TEXT')
_add_column_if_missing('military', 'removal_basis', 'TEXT')
_add_column_if_missing('military', 'tck_name', 'TEXT')
_add_column_if_missing('military', 'registry_number', 'TEXT')
_add_column_if_missing('military', 'vos_code', 'TEXT')
_add_column_if_missing('military', 'vos_description', 'TEXT')
_add_column_if_missing('military', 'deferral_type', 'TEXT')
_add_column_if_missing('military', 'deferral_until', 'TEXT')
_add_column_if_missing('military', 'police_reason', 'TEXT')
_add_column_if_missing('military', 'police_date', 'TEXT')
_add_column_if_missing('military', 'vlk_decision', 'TEXT')
_add_column_if_missing('military', 'vlk_date', 'TEXT')
_add_column_if_missing('military', 'disability_group', 'TEXT')
_add_column_if_missing('military', 'disability_valid_until', 'TEXT')
_add_column_if_missing('military', 'disability_reason', 'TEXT')
_add_column_if_missing('military', 'formed_at', 'TEXT')
_add_column_if_missing('military', 'data_verified_on_time', 'INTEGER DEFAULT 0')
_add_column_if_missing('military', 'data_update_date', 'TEXT')
_add_column_if_missing('military', 'phone', 'TEXT')
_add_column_if_missing('military', 'email', 'TEXT')
_add_column_if_missing('military', 'note_18', 'TEXT')
_add_column_if_missing('military', 'mobilized_date', 'TEXT')

# institution_licenses
_add_column_if_missing('institution_licenses', 'is_base', 'INTEGER NOT NULL DEFAULT 0')
cur.execute("""
    UPDATE institution_licenses SET is_base = 1
    WHERE short_name_ua = 'Львівська'
      AND (SELECT COUNT(*) FROM institution_licenses WHERE is_base = 1) = 0
""")

# Відомі колонки з попередніх версій init_db.py (залишаємо для сумісності)
_add_column_if_missing('update_requests', 'applied_fields', 'TEXT')
_add_column_if_missing('specialties', 'isced_f_ua', 'TEXT')
_add_column_if_missing('specialties', 'isced_f_en', 'TEXT')
_add_column_if_missing('groups', 'isced_f_ua', 'TEXT')
_add_column_if_missing('groups', 'isced_f_en', 'TEXT')

# ── Автопідтягування ISCED-F для вже існуючих груп ──
cur.execute("""
    UPDATE groups SET
        isced_f_ua = (SELECT s.isced_f_ua FROM specialties s WHERE s.code = groups.specialty_code),
        isced_f_en = (SELECT s.isced_f_en FROM specialties s WHERE s.code = groups.specialty_code)
    WHERE specialty_code IS NOT NULL
      AND isced_f_ua IS NULL AND isced_f_en IS NULL
      AND EXISTS (SELECT 1 FROM specialties s WHERE s.code = groups.specialty_code
                  AND (s.isced_f_ua IS NOT NULL OR s.isced_f_en IS NOT NULL))
""")
if cur.rowcount:
    print(f"➕ Підтягнуто ISCED-F для {cur.rowcount} існуючих груп")


# ═══════════════════════════════════════════════════════════════════════
# 4. ДАНІ-МІГРАЦІЯ ДЛЯ ІСНУЮЧИХ СТУДЕНТІВ
# ═══════════════════════════════════════════════════════════════════════
def detect_gender_from_middle_name(middle_name):
    """Стать за по батькові (з migrate_military_reports.py)."""
    if not middle_name:
        return None
    m = middle_name.strip().lower().replace('’', "'").replace('ʼ', "'")
    if m.endswith(('івна', 'ївна', 'овна', 'евна', 'євна', 'ична')):
        return 'F'
    return 'M'


# Автозаповнення gender для тих, у кого NULL
rows = cur.execute(
    "SELECT id, middle_name_UA FROM students WHERE gender IS NULL"
).fetchall()
if rows:
    updated = 0
    for student_id, middle_name in rows:
        g = detect_gender_from_middle_name(middle_name)
        if g:
            cur.execute("UPDATE students SET gender = ? WHERE id = ?", (g, student_id))
            updated += 1
    if updated:
        print(f"➕ Автовизначено стать для {updated} з {len(rows)} студентів")
    missing = len(rows) - updated
    if missing:
        print(f"   ⚠️ {missing} студентів без по батькові — стать лишилась NULL")


# ═══════════════════════════════════════════════════════════════════════
# 5. ФІНАЛ
# ═══════════════════════════════════════════════════════════════════════
conn.commit()
conn.close()
print()
print("✅ БД готова до роботи.")
print("⚠️  Якщо користувач 'admin' щойно створений — стандартний пароль "
      "'admin123'. Змініть його одразу після першого входу через "
      "/admin/users/<id>/change-password.")