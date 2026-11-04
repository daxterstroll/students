"""
routes/init_db.py
==================
Одноразовий скрипт для створення БД students.db "з нуля" та
ідемпотентного доведення вже існуючої бази до актуальної схеми.

Замінює собою:
  - routes/init_db.py (початкове створення)
  - migrate_military_reports.py     (v1: gender, category_override, tck_offices, log)
  - migrate_military_reports_v2.py  (поля «Резерв+», military_history)
  - migrate_military_reports_v3.py  (Постанова №812: note_18, dismissed_at, is_base)
  - єдина модель наказів: orders / order_items
    (expulsion_* і course_transfer_* більше не створюються; лише orders / order_items)

Запуск: `python init_db.py` з кореня проєкту
(шлях до БД обчислюється відносно розташування файлу - не залежить від CWD).

Порядок виконання:
  0. Бекап (якщо БД уже існує і має дані).
  1. CREATE TABLE IF NOT EXISTS - схема "з нуля" (усі колонки відразу).
  2. ALTER TABLE ... ADD COLUMN - для вже існуючих баз (додає лише те,
     чого бракує; на чистій БД нічого не робить - бо колонки вже є з
     кроку 1).
  3. SEED-дані (INSERT OR IGNORE): ліцензії, ступені, галузі, спеціальності,
     ТЦК, admin-користувач. ВИКОНУЄТЬСЯ ПІСЛЯ кроку 2, щоб на існуючій
     базі нові колонки вже існували до INSERT.
  4. Одноразові UPDATE для заповнення нових полів у вже існуючих записах.
  5. commit + звіт.

Повторний запуск БЕЗПЕЧНИЙ: наявні дані не видаляються, лише додаються
відсутні колонки й таблиці. Якщо БД уже існує та містить дані — спершу
створюється резервна копія у тій самій папці.
"""

import os
import shutil
import sqlite3
from datetime import datetime

from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, 'students.db')


# ═══════════════════════════════════════════════════════════════════════
# 0. Резервна копія (лише якщо БД уже існує і має реальний вміст)
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
# На чистій БД одразу створює актуальну схему з усіма колонками.
# На вже існуючій БД нічого не змінює (колонки додаються у кроці 2).
# ═══════════════════════════════════════════════════════════════════════
cur.executescript("""
-- ─── Довідники (спочатку — на них посилаються інші таблиці) ───
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
CREATE TABLE IF NOT EXISTS "degree_levels" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name_ua TEXT NOT NULL UNIQUE,
    name_en TEXT,
    is_active INTEGER NOT NULL DEFAULT 1,
    name_suffix TEXT
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
CREATE TABLE IF NOT EXISTS "institution_licenses" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name_ua TEXT NOT NULL,
    short_name_ua TEXT,
    name_en TEXT,
    short_name_en TEXT,
    is_active INTEGER NOT NULL DEFAULT 1,
    is_base INTEGER NOT NULL DEFAULT 0
);

-- ─── Групи й студенти ───
CREATE TABLE IF NOT EXISTS "groups" (
    "id"    INTEGER,
    "name"  TEXT NOT NULL,
    "start_year"    INTEGER NOT NULL,
    "study_form"    TEXT NOT NULL CHECK("study_form" IN ('Денна', 'Заочна', 'Дистанційна')),
    "program_credits"   INTEGER NOT NULL,
    "degree_level"  TEXT NOT NULL,
    "degree_level_en"   TEXT NOT NULL,
    "knowledge_area"    TEXT NOT NULL,
    "knowledge_area_en" TEXT NOT NULL,
    "specialty" TEXT NOT NULL,
    "specialty_en"  TEXT NOT NULL,
    "isced_f_ua"    TEXT,
    "isced_f_en"    TEXT,
    "specialty_code"    TEXT,
    "educational_program"   TEXT NOT NULL,
    "educational_program_en"    TEXT NOT NULL,
    "qualification_name"    TEXT NOT NULL,
    "qualification_name_en" TEXT NOT NULL,
    "course"    INTEGER NOT NULL DEFAULT 1,
    "institution_name_and_status"   TEXT,
    "institution_name_and_status_en"    TEXT,
    "entry_requirements"    TEXT,
    "entry_requirements_en" TEXT,
    "learning_outcomes" TEXT,
    "learning_outcomes_en"  TEXT,
    "program_includes"  TEXT,
    "program_includes_en"   TEXT,
    "entry_requirements_reduced"    TEXT,
    "entry_requirements_reduced_en" TEXT,
    "program_includes_reduced"  TEXT,
    "program_includes_reduced_en"   TEXT,
    "learning_outcomes_reduced" TEXT,
    "learning_outcomes_reduced_en"  TEXT,
    "archived"  BOOLEAN DEFAULT FALSE,
    PRIMARY KEY("id" AUTOINCREMENT),
    UNIQUE("name","start_year","study_form","program_credits")
);
CREATE TABLE IF NOT EXISTS "students" (
    "id"    INTEGER,
    "last_name_UA"  TEXT,
    "first_name_UA" TEXT,
    "middle_name_UA"    TEXT,
    "last_name_ENG" TEXT,
    "first_name_ENG"    TEXT,
    "birth_date"    TEXT,
    "group_id"  INTEGER,
    "edebo_code"    VARCHAR(50),
    "tax_id"    TEXT,
    "photo" TEXT,
    "license_id"    INTEGER,
    "archived"  BOOLEAN DEFAULT FALSE,
    "program_credits_override"  INTEGER,
    "phone" TEXT,
    "phone_backup"  TEXT,
    "email" TEXT,
    "address_of_residence"  TEXT,
    "gender"    TEXT CHECK("gender" IN ('M', 'F')),
    "dismissed_at"  TEXT,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("group_id") REFERENCES "groups"("id"),
    FOREIGN KEY("license_id") REFERENCES "institution_licenses"("id")
);

CREATE TABLE IF NOT EXISTS student_enrollments (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id  INTEGER NOT NULL,
    group_id    INTEGER NOT NULL,
    is_primary  INTEGER NOT NULL DEFAULT 0,
    program_credits_override INTEGER,
    status      TEXT NOT NULL DEFAULT 'active',
    license_id  INTEGER,
    edebo_code  TEXT,                    -- ЄДЕБО для ЦЬОГО навчання
    order_id    INTEGER,
    notes       TEXT,
    is_correction INTEGER NOT NULL DEFAULT 0,
    correction_note TEXT,
    correction_by TEXT,
    correction_at TEXT,
    created_at  TEXT DEFAULT (datetime('now', 'localtime')),
    UNIQUE (student_id, group_id),
    FOREIGN KEY (student_id) REFERENCES students(id),
    FOREIGN KEY (group_id) REFERENCES groups(id),
    FOREIGN KEY (license_id) REFERENCES institution_licenses(id)
);

-- ─── Військовий облік ───
CREATE TABLE IF NOT EXISTS "military" (
    "id"    INTEGER,
    "student_id"    INTEGER,
    "registration_number_of_the_DRPVR"  TEXT,
    "military_registration_document"    TEXT,
    "issued_VOD"    TEXT,
    "military_accounting_specialty_number"  TEXT,
    "military_rank" TEXT,
    "change_credentials"    TEXT,
    "reason_for_changing_credentials"   TEXT,
    "being_on_military_registration"    TEXT,
    "address_of_residence"  TEXT,
    "category_override" TEXT CHECK("category_override" IN ('officer', 'rank')),
    "status_badge"  TEXT,
    "valid_until"   TEXT,
    "registration_category" TEXT,
    "removal_basis" TEXT,
    "tck_name"  TEXT,
    "registry_number"   TEXT,
    "vos_code"  TEXT,
    "vos_description"   TEXT,
    "deferral_type" TEXT,
    "deferral_until"    TEXT,
    "police_reason" TEXT,
    "police_date"   TEXT,
    "vlk_decision"  TEXT,
    "vlk_date"  TEXT,
    "disability_group"  TEXT,
    "disability_valid_until"    TEXT,
    "disability_reason" TEXT,
    "formed_at" TEXT,
    "data_verified_on_time" INTEGER DEFAULT 0,
    "data_update_date"  TEXT,
    "phone" TEXT,
    "email" TEXT,
    "note_18"       TEXT,
    "mobilized_date"    TEXT,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("student_id") REFERENCES "students"("id")
);

-- ─── Освітні компоненти ───
CREATE TABLE IF NOT EXISTS "subjects" (
    "id"    INTEGER,
    "code"  TEXT,
    "name"  TEXT NOT NULL,
    "credits"   INTEGER,
    "group_id"  INTEGER,
    "position"  INTEGER DEFAULT 0,
    "type"  TEXT DEFAULT 'Залік' CHECK("type" IN ('Залік', 'Екзамен')),
    "full_program_only" INTEGER NOT NULL DEFAULT 0,
    "reduced_credits"   INTEGER,
    "reduced_type"  TEXT,
    "reduced_only"  INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
CREATE TABLE IF NOT EXISTS "practices" (
    "id"    INTEGER,
    "code"  TEXT NOT NULL,
    "name"  TEXT NOT NULL,
    "credits"   INTEGER NOT NULL,
    "type"  TEXT NOT NULL CHECK("type" IN ('Залік', 'Екзамен')),
    "position"  INTEGER NOT NULL,
    "group_id"  INTEGER NOT NULL,
    "full_program_only" INTEGER NOT NULL DEFAULT 0,
    "reduced_credits"   INTEGER,
    "reduced_type"  TEXT,
    "reduced_only"  INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
CREATE TABLE IF NOT EXISTS "courseworks" (
    "id"    INTEGER,
    "code"  TEXT NOT NULL,
    "name"  TEXT NOT NULL,
    "credits"   INTEGER NOT NULL,
    "type"  TEXT NOT NULL CHECK("type" IN ('Залік', 'Екзамен')),
    "position"  INTEGER NOT NULL,
    "group_id"  INTEGER NOT NULL,
    "full_program_only" INTEGER NOT NULL DEFAULT 0,
    "reduced_credits"   INTEGER,
    "reduced_type"  TEXT,
    "reduced_only"  INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("group_id") REFERENCES "groups"("id")
);
CREATE TABLE IF NOT EXISTS "attestations" (
    "id"    INTEGER,
    "code"  TEXT NOT NULL,
    "name"  TEXT NOT NULL,
    "credits"   INTEGER NOT NULL,
    "type"  TEXT NOT NULL CHECK("type" IN ('Залік', 'Екзамен')),
    "position"  INTEGER NOT NULL,
    "group_id"  INTEGER NOT NULL,
    "full_program_only" INTEGER NOT NULL DEFAULT 0,
    "reduced_credits"   INTEGER,
    "reduced_type"  TEXT,
    "reduced_only"  INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("group_id") REFERENCES "groups"("id")
);

-- ─── Оцінки ───
CREATE TABLE IF NOT EXISTS "grades" (
    "id"    INTEGER,
    "student_id"    INTEGER,
    "subject_id"    INTEGER,
    "grade" TEXT,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("subject_id") REFERENCES "subjects"("id"),
    FOREIGN KEY("student_id") REFERENCES "students"("id")
);
CREATE TABLE IF NOT EXISTS "activity_grades" (
    "id"    INTEGER,
    "student_id"    INTEGER NOT NULL,
    "entity_id" INTEGER NOT NULL,
    "entity_type"   TEXT NOT NULL CHECK("entity_type" IN ('practice', 'coursework', 'attestation')),
    "grade" INTEGER,
    "name"  TEXT,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("student_id") REFERENCES "students"("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- ─── Документи ───
CREATE TABLE IF NOT EXISTS "passport_documents" (
    "id"    INTEGER,
    "student_id"    INTEGER NOT NULL,
    "document_type" TEXT NOT NULL CHECK("document_type" IN (
        'Свідоцтво про народження',
        'Паспорт громадянина України',
        'Паспорт громадянина України з безконтактним електронним носієм',
        'Паспорт громадянина України, що не містить безконтактного електронного носія'
    )),
    "series"    TEXT,
    "number"    TEXT NOT NULL,
    "issued_by" TEXT,
    "issue_date"    TEXT,
    "valid_until"   TEXT,
    "unique_number" TEXT,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("student_id") REFERENCES "students"("id")
);
CREATE TABLE IF NOT EXISTS "education_documents" (
    "id"    INTEGER,
    "student_id"    INTEGER NOT NULL,
    "document_type" TEXT NOT NULL,
    "document_type_en"  TEXT NOT NULL,
    "document_number"   TEXT NOT NULL,
    "institution_name"  TEXT NOT NULL,
    "institution_name_en"   TEXT NOT NULL,
    "country"   TEXT NOT NULL,
    "country_en"    TEXT NOT NULL,
    "completion_date"   TEXT NOT NULL,
    "group_id"      INTEGER REFERENCES groups(id),
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("student_id") REFERENCES "students"("id")
);
CREATE TABLE IF NOT EXISTS "foreign_education_docs" (
    "id"    INTEGER,
    "education_doc_id"  INTEGER NOT NULL,
    "reference_number"  TEXT,
    "reference_institution" TEXT,
    "reference_institution_en"  TEXT,
    "reference_country" TEXT,
    "reference_country_en"  TEXT,
    "reference_issue_date"  TEXT,
    "recognition_certificate_number"    TEXT,
    "recognition_issuer"    TEXT,
    "recognition_issuer_en" TEXT,
    "recognition_date"  TEXT,
    PRIMARY KEY("id" AUTOINCREMENT),
    FOREIGN KEY("education_doc_id") REFERENCES "education_documents"("id")
);
CREATE TABLE IF NOT EXISTS "diplomas" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL,
    group_id INTEGER,
    diploma_number TEXT,
    appendix_number TEXT,
    FOREIGN KEY (student_id) REFERENCES students(id),
    FOREIGN KEY (group_id) REFERENCES groups(id)
);
CREATE TABLE IF NOT EXISTS "accreditations" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    degree TEXT NOT NULL,
    specialty TEXT NOT NULL,
    text_ua TEXT,
    text_en TEXT
);
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

-- ─── Періоди навчання ───
CREATE TABLE IF NOT EXISTS "student_study_periods" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    group_id INTEGER REFERENCES groups(id),
    filiya TEXT NOT NULL,
    filiya_en TEXT NOT NULL,
    group_name TEXT,
    start_date TEXT,
    end_date TEXT,
    period_order INTEGER NOT NULL DEFAULT 0,
    note TEXT
);
CREATE INDEX IF NOT EXISTS idx_study_periods_student ON student_study_periods(student_id);

-- ─── Шаблони документів ───
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

-- ─── Користувачі ───
CREATE TABLE IF NOT EXISTS "users" (
    "id"    INTEGER,
    "username"  TEXT UNIQUE,
    "password_hash" TEXT,
    "role"  TEXT NOT NULL CHECK("role" IN ('admin', 'user')),
    "is_admin"  INTEGER DEFAULT 0,
    "permissions"   TEXT DEFAULT '[]',
    PRIMARY KEY("id" AUTOINCREMENT)
);
CREATE TABLE IF NOT EXISTS "user_groups" (
    "user_id"   INTEGER,
    "group_id"  INTEGER,
    PRIMARY KEY("user_id","group_id"),
    FOREIGN KEY("group_id") REFERENCES "groups"("id") ON DELETE CASCADE,
    FOREIGN KEY("user_id") REFERENCES "users"("id") ON DELETE CASCADE
);

-- ─── Заявки ───
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

-- ─── Накази й рух студентів ───
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
    FOREIGN KEY (previous_group_id) REFERENCES groups(id)
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

-- ─── Військовий облік: довідники й журнали ───
CREATE TABLE IF NOT EXISTS "tck_offices" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name_ua TEXT NOT NULL UNIQUE,
    address TEXT,
    is_active INTEGER NOT NULL DEFAULT 1
);
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

CREATE TABLE IF NOT EXISTS "student_name_history" (
    "id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "student_id" INTEGER NOT NULL,
    "order_id" INTEGER,
    "changed_at" TEXT DEFAULT (datetime('now', 'localtime')),
    "changed_by" TEXT,
    "old_last_name_UA" TEXT, "old_first_name_UA" TEXT, "old_middle_name_UA" TEXT,
    "old_last_name_ENG" TEXT, "old_first_name_ENG" TEXT,
    FOREIGN KEY("student_id") REFERENCES "students"("id") ON DELETE CASCADE,
    FOREIGN KEY("order_id") REFERENCES "orders"("id")
);
CREATE INDEX IF NOT EXISTS "idx_student_name_history_student"
    ON "student_name_history" ("student_id", "id");    
    
CREATE TABLE IF NOT EXISTS "military_history" (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL,
    changed_at TEXT DEFAULT (datetime('now','localtime')),
    changed_by TEXT,
    source TEXT,
    scope TEXT,
    field_name TEXT NOT NULL,
    old_value TEXT,
    new_value TEXT,
    FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_military_history_student
    ON military_history (student_id, id);
    
CREATE TABLE IF NOT EXISTS "orders" (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    order_type   TEXT NOT NULL,
        -- enrollment | course_transfer | license_transfer
        -- | freeze | unfreeze | expulsion | graduation
    order_number TEXT NOT NULL,
    order_date   TEXT NOT NULL,          -- DD.MM.YYYY (як у старих таблицях)
    title        TEXT,
    notes        TEXT,
    created_by   TEXT,
    created_at   TEXT DEFAULT (datetime('now', 'localtime')),
    status       TEXT NOT NULL DEFAULT 'active',
        -- active | cancelled
    -- Зв'язок зі старим записом (для міграції і перехідного періоду)
    legacy_table TEXT,                  -- напр. 'course_transfer_orders'
    legacy_id    INTEGER
);

CREATE INDEX IF NOT EXISTS "idx_orders_type_date"
    ON "orders" ("order_type", "order_date");
CREATE INDEX IF NOT EXISTS "idx_orders_number"
    ON "orders" ("order_number");
CREATE INDEX IF NOT EXISTS "idx_orders_legacy"
    ON "orders" ("legacy_table", "legacy_id");

CREATE TABLE IF NOT EXISTS "order_items" (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id    INTEGER NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
    entity_type TEXT NOT NULL,          -- 'student' | 'group'
    entity_id   INTEGER NOT NULL,
    -- Контекст дії (JSON-рядок):
    -- course_transfer: {"course_from":2,"course_to":3}
    -- expulsion:       {"reason":"...","previous_group_id":12}
    -- license_transfer:{"previous_license_id":1,"new_license_id":2}
    -- enrollment:      {"group_id":5}
    -- freeze:          {"reason":"...","previous_group_id":5}
    payload     TEXT,
    UNIQUE (order_id, entity_type, entity_id)
);

CREATE INDEX IF NOT EXISTS "idx_order_items_entity"
    ON "order_items" ("entity_type", "entity_id");
CREATE INDEX IF NOT EXISTS "idx_order_items_order"
    ON "order_items" ("order_id");    

""")




# ═══════════════════════════════════════════════════════════════════════
# 2. МІГРАЦІЇ ДЛЯ ВЖЕ ІСНУЮЧИХ БАЗ
# ═══════════════════════════════════════════════════════════════════════
# ЦЕЙ БЛОК ІДЕ ПЕРЕД SEED-ДАНИМИ: якщо БД уже існує, тут додаються
# колонки, яких у ній ще немає (CREATE TABLE IF NOT EXISTS їх не додає),
# і лише після цього INSERT-и в секції 3 матимуть усі потрібні колонки.
# На чистій БД нічого не робить - колонки вже є з кроку 1.
# ═══════════════════════════════════════════════════════════════════════
def _add_column_if_missing(table, column, coltype):
    """Додає колонку, якщо її ще немає (ідемпотентно)."""
    existing = {row[1] for row in cur.execute(f'PRAGMA table_info("{table}")')}
    if column not in existing:
        cur.execute(f'ALTER TABLE "{table}" ADD COLUMN "{column}" {coltype}')
        print(f"➕ {table}.{column}")


# Legacy-таблиці відрахування більше не використовуються (модель orders).
for _t in (
    'expulsion_order_students', 'expulsion_orders',
    'course_transfer_order_groups', 'course_transfer_orders',
):
    row = cur.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (_t,)
    ).fetchone()
    if row:
        cur.execute(f'DROP TABLE IF EXISTS "{_t}"')
        print(f"🗑️  прибрано legacy-таблицю {_t}")


print("── Міграція схеми ──")

# students
_add_column_if_missing('students', 'gender', "TEXT CHECK(gender IN ('M', 'F'))")
_add_column_if_missing('students', 'address_of_residence', 'TEXT')
_add_column_if_missing('students', 'dismissed_at', 'TEXT')

# student_study_periods — привязка періодів до конкретного enrollment'а
_add_column_if_missing('student_study_periods', 'group_id', 'INTEGER REFERENCES groups(id)')
cur.execute("CREATE INDEX IF NOT EXISTS idx_study_periods_group ON student_study_periods(group_id)")

# student_enrollments — ЄДЕБО на рівні навчання
_add_column_if_missing('student_enrollments', 'edebo_code', 'TEXT')

# student_enrollments — кредити скороченої програми на рівні навчання
_add_column_if_missing('student_enrollments', 'program_credits_override', 'INTEGER')

# Одноразова міграція: students.program_credits_override → primary enrollment
cur.execute("""
    UPDATE student_enrollments
    SET program_credits_override = (
        SELECT s.program_credits_override FROM students s
        WHERE s.id = student_enrollments.student_id
          AND s.program_credits_override IS NOT NULL
    )
    WHERE program_credits_override IS NULL
      AND is_primary = 1
      AND EXISTS (
          SELECT 1 FROM students s
          WHERE s.id = student_enrollments.student_id
            AND s.program_credits_override IS NOT NULL
      )
""")
if cur.rowcount:
    print(f"➕ student_enrollments.program_credits_override: перенесено з students для {cur.rowcount} primary-записів")

# Позначка «додано/змінено без наказу» (адмінська корекція помилки).
_add_column_if_missing('student_enrollments', 'is_correction', 'INTEGER NOT NULL DEFAULT 0')
_add_column_if_missing('student_enrollments', 'correction_note', 'TEXT')
_add_column_if_missing('student_enrollments', 'correction_by', 'TEXT')
_add_column_if_missing('student_enrollments', 'correction_at', 'TEXT')

# Одноразовий backfill: старий edebo_code з students → у primary enrollment.
# Робимо ЛИШЕ якщо в enrollment ще порожньо, і тільки для primary
# (щоб у паралельного студента не «затекло» одне значення у всі групи).
cur.execute("""
    UPDATE student_enrollments
    SET edebo_code = (
        SELECT s.edebo_code FROM students s
        WHERE s.id = student_enrollments.student_id
          AND s.edebo_code IS NOT NULL AND s.edebo_code != ''
    )
    WHERE edebo_code IS NULL
      AND is_primary = 1
      AND EXISTS (
          SELECT 1 FROM students s
          WHERE s.id = student_enrollments.student_id
            AND s.edebo_code IS NOT NULL AND s.edebo_code != ''
      )
""")
if cur.rowcount:
    print(f"➕ student_enrollments.edebo_code: перенесено з students для {cur.rowcount} primary-записів")

# Одноразове заповнення: старі періоди (без group_id) прив'язуємо до
# ОСНОВНОЇ групи студента на момент міграції — далі студент зможе
# змінити це вручну в картці студента.
cur.execute("""
    UPDATE student_study_periods
    SET group_id = (SELECT s.group_id FROM students s WHERE s.id = student_study_periods.student_id)
    WHERE group_id IS NULL
      AND EXISTS (
          SELECT 1 FROM students s
          WHERE s.id = student_study_periods.student_id
            AND s.group_id IS NOT NULL
      )
""")
if cur.rowcount:
    print(f"➕ study_periods: прив'язано до групи {cur.rowcount} записів")

# diplomas — ключ тепер (student_id, group_id), а не просто student_id
_add_column_if_missing('diplomas', 'group_id', 'INTEGER')

# education_documents — документ про освіту для конкретної групи
# (напр. для бакалаврату — шкільний атестат, для магістратури — диплом бакалавра)
_add_column_if_missing('education_documents', 'group_id', 'INTEGER REFERENCES groups(id)')
cur.execute("CREATE INDEX IF NOT EXISTS idx_education_documents_group ON education_documents(group_id)")

# Backfill: старі документи → до ОСНОВНОЇ групи студента
cur.execute("""
    UPDATE education_documents
    SET group_id = (SELECT s.group_id FROM students s WHERE s.id = education_documents.student_id)
    WHERE group_id IS NULL
      AND EXISTS (SELECT 1 FROM students s WHERE s.id = education_documents.student_id AND s.group_id IS NOT NULL)
""")
if cur.rowcount:
    print(f"➕ education_documents: прив'язано до групи {cur.rowcount} записів")

# military — усі поля з v1/v2/v3
for _col, _type in [
    ('category_override', "TEXT CHECK(category_override IN ('officer', 'rank'))"),
    ('status_badge', 'TEXT'),
    ('valid_until', 'TEXT'),
    ('registration_category', 'TEXT'),
    ('removal_basis', 'TEXT'),
    ('tck_name', 'TEXT'),
    ('registry_number', 'TEXT'),
    ('vos_code', 'TEXT'),
    ('vos_description', 'TEXT'),
    ('deferral_type', 'TEXT'),
    ('deferral_until', 'TEXT'),
    ('police_reason', 'TEXT'),
    ('police_date', 'TEXT'),
    ('vlk_decision', 'TEXT'),
    ('vlk_date', 'TEXT'),
    ('disability_group', 'TEXT'),
    ('disability_valid_until', 'TEXT'),
    ('disability_reason', 'TEXT'),
    ('formed_at', 'TEXT'),
    ('data_verified_on_time', 'INTEGER DEFAULT 0'),
    ('data_update_date', 'TEXT'),
    ('phone', 'TEXT'),
    ('email', 'TEXT'),
    ('note_18', 'TEXT'),
    ('mobilized_date', 'TEXT'),
]:
    _add_column_if_missing('military', _col, _type)

# institution_licenses
_add_column_if_missing('institution_licenses', 'is_base', 'INTEGER NOT NULL DEFAULT 0')

# degree_levels
_add_column_if_missing('degree_levels', 'name_suffix', 'TEXT')

# Історичні колонки (для сумісності з попередніми версіями)
_add_column_if_missing('update_requests', 'applied_fields', 'TEXT')
_add_column_if_missing('specialties', 'isced_f_ua', 'TEXT')
_add_column_if_missing('specialties', 'isced_f_en', 'TEXT')
_add_column_if_missing('groups', 'isced_f_ua', 'TEXT')
_add_column_if_missing('groups', 'isced_f_en', 'TEXT')


# ── groups.study_form: додати 'Дистанційна' до CHECK ──
_groups_sql_row = cur.execute(
    "SELECT sql FROM sqlite_master WHERE type='table' AND name='groups'"
).fetchone()
if _groups_sql_row and "'Дистанційна'" not in (_groups_sql_row[0] or ''):
    print("🔁 groups: розширюю CHECK для study_form (перезбір таблиці)")

    # На випадок, якщо попередній запуск упав на середині — прибрати недобудовану таблицю.
    cur.execute('DROP TABLE IF EXISTS "groups_new"')

    # 1. Гарантуємо, що всі NOT NULL-колонки існують і що в них немає NULL.
    _not_null_fixups = [
        ('degree_level', "TEXT NOT NULL DEFAULT ''"),
        ('degree_level_en', "TEXT NOT NULL DEFAULT ''"),
        ('knowledge_area', "TEXT NOT NULL DEFAULT ''"),
        ('knowledge_area_en', "TEXT NOT NULL DEFAULT ''"),
        ('specialty', "TEXT NOT NULL DEFAULT ''"),
        ('specialty_en', "TEXT NOT NULL DEFAULT ''"),
        ('educational_program', "TEXT NOT NULL DEFAULT ''"),
        ('educational_program_en', "TEXT NOT NULL DEFAULT ''"),
        ('qualification_name', "TEXT NOT NULL DEFAULT ''"),
        ('qualification_name_en', "TEXT NOT NULL DEFAULT ''"),
        ('name', "TEXT NOT NULL DEFAULT ''"),
        ('study_form', "TEXT NOT NULL DEFAULT ''"),
        ('course', "INTEGER NOT NULL DEFAULT 1"),
        ('start_year', "INTEGER NOT NULL DEFAULT 0"),
        ('program_credits', "INTEGER NOT NULL DEFAULT 0"),
    ]
    for _col, _coltype in _not_null_fixups:
        _add_column_if_missing('groups', _col, _coltype)
    for _col, _ in _not_null_fixups:
        cur.execute(f'UPDATE "groups" SET "{_col}" = \'\' WHERE "{_col}" IS NULL')
    cur.execute('UPDATE "groups" SET "course" = 1 WHERE "course" IS NULL')
    cur.execute('UPDATE "groups" SET "start_year" = 0 WHERE "start_year" IS NULL')
    cur.execute('UPDATE "groups" SET "program_credits" = 0 WHERE "program_credits" IS NULL')

    # 2. Інші (nullable) колонки, яких може не бути в старій таблиці.
    _other_cols = [
        ('isced_f_ua', 'TEXT'),
        ('isced_f_en', 'TEXT'),
        ('specialty_code', 'TEXT'),
        ('institution_name_and_status', 'TEXT'),
        ('institution_name_and_status_en', 'TEXT'),
        ('entry_requirements', 'TEXT'),
        ('entry_requirements_en', 'TEXT'),
        ('learning_outcomes', 'TEXT'),
        ('learning_outcomes_en', 'TEXT'),
        ('program_includes', 'TEXT'),
        ('program_includes_en', 'TEXT'),
        ('entry_requirements_reduced', 'TEXT'),
        ('entry_requirements_reduced_en', 'TEXT'),
        ('program_includes_reduced', 'TEXT'),
        ('program_includes_reduced_en', 'TEXT'),
        ('learning_outcomes_reduced', 'TEXT'),
        ('learning_outcomes_reduced_en', 'TEXT'),
        ('archived', 'BOOLEAN DEFAULT FALSE'),
    ]
    for _col, _coltype in _other_cols:
        _add_column_if_missing('groups', _col, _coltype)

    # 3. Явний порядок колонок для INSERT ... SELECT (не SELECT *!).
    #    Якщо використати зірочку, при різному порядку колонок у старій
    #    таблиці (через ALTER ADD COLUMN) дані тихо потраплять не в ті
    #    колонки нової таблиці.
    _all_cols_ordered = [
        'id', 'name', 'start_year', 'study_form', 'program_credits',
        'degree_level', 'degree_level_en', 'knowledge_area', 'knowledge_area_en',
        'specialty', 'specialty_en', 'isced_f_ua', 'isced_f_en', 'specialty_code',
        'educational_program', 'educational_program_en',
        'qualification_name', 'qualification_name_en', 'course',
        'institution_name_and_status', 'institution_name_and_status_en',
        'entry_requirements', 'entry_requirements_en',
        'learning_outcomes', 'learning_outcomes_en',
        'program_includes', 'program_includes_en',
        'entry_requirements_reduced', 'entry_requirements_reduced_en',
        'program_includes_reduced', 'program_includes_reduced_en',
        'learning_outcomes_reduced', 'learning_outcomes_reduced_en',
        'archived',
    ]
    _existing = {row[1] for row in cur.execute('PRAGMA table_info("groups")')}
    _cols_csv = ', '.join(f'"{c}"' for c in _all_cols_ordered if c in _existing)

    # 4. Пересборка. PRAGMA foreign_keys = OFF — обов'язково ПОЗА транзакцією.
    conn.commit()
    cur.execute("PRAGMA foreign_keys = OFF")
    try:
        cur.executescript(f"""
            CREATE TABLE "groups_new" (
                "id" INTEGER,
                "name" TEXT NOT NULL,
                "start_year" INTEGER NOT NULL,
                "study_form" TEXT NOT NULL CHECK("study_form" IN ('Денна', 'Заочна', 'Дистанційна')),
                "program_credits" INTEGER NOT NULL,
                "degree_level" TEXT NOT NULL,
                "degree_level_en" TEXT NOT NULL,
                "knowledge_area" TEXT NOT NULL,
                "knowledge_area_en" TEXT NOT NULL,
                "specialty" TEXT NOT NULL,
                "specialty_en" TEXT NOT NULL,
                "isced_f_ua" TEXT,
                "isced_f_en" TEXT,
                "specialty_code" TEXT,
                "educational_program" TEXT NOT NULL,
                "educational_program_en" TEXT NOT NULL,
                "qualification_name" TEXT NOT NULL,
                "qualification_name_en" TEXT NOT NULL,
                "course" INTEGER NOT NULL DEFAULT 1,
                "institution_name_and_status" TEXT,
                "institution_name_and_status_en" TEXT,
                "entry_requirements" TEXT,
                "entry_requirements_en" TEXT,
                "learning_outcomes" TEXT,
                "learning_outcomes_en" TEXT,
                "program_includes" TEXT,
                "program_includes_en" TEXT,
                "entry_requirements_reduced" TEXT,
                "entry_requirements_reduced_en" TEXT,
                "program_includes_reduced" TEXT,
                "program_includes_reduced_en" TEXT,
                "learning_outcomes_reduced" TEXT,
                "learning_outcomes_reduced_en" TEXT,
                "archived" BOOLEAN DEFAULT FALSE,
                PRIMARY KEY("id" AUTOINCREMENT),
                UNIQUE("name","start_year","study_form","program_credits")
            );
            INSERT INTO "groups_new" ({_cols_csv}) SELECT {_cols_csv} FROM "groups";
            DROP TABLE "groups";
            ALTER TABLE "groups_new" RENAME TO "groups";
        """)
        conn.commit()
        _n = cur.execute('SELECT COUNT(*) FROM groups').fetchone()[0]
        print(f"   ✅ перезбудовано, записів збережено: {_n}")
    finally:
        try:
            conn.rollback()
        except Exception:
            pass
        cur.execute("PRAGMA foreign_keys = ON")

# ── student_enrollments: видалити дублікати пар (student_id, group_id) ──
# Якщо БД створювалась старою версією CREATE TABLE — UNIQUE-констрейнту
# не було, і в таблиці могли накопичитись дублікати. Один такий
# дублікат ламає картку студента (groups_data містить ту саму групу
# двічі → шаблон рендерить 2 таби + повторює оцінки/документи/періоди).
#
# Залишаємо одну запис на пару, з пріоритетом:
#   1) status='active'
#   2) is_primary=1
#   3) найбільший id (найсвіжіша)
_dup_count = cur.execute("""
    SELECT COUNT(*) FROM (
        SELECT student_id, group_id
        FROM student_enrollments
        GROUP BY student_id, group_id
        HAVING COUNT(*) > 1
    )
""").fetchone()[0]

if _dup_count:
    print(f"⚠️ student_enrollments: знайдено {_dup_count} дублікатів пар (student_id, group_id) — чистка")
    cur.execute("""
        DELETE FROM student_enrollments
        WHERE id NOT IN (
            SELECT id FROM (
                SELECT id,
                       ROW_NUMBER() OVER (
                           PARTITION BY student_id, group_id
                           ORDER BY
                               CASE status
                                   WHEN 'active' THEN 0
                                   WHEN 'frozen' THEN 1
                                   WHEN 'completed' THEN 2
                                   WHEN 'graduated' THEN 3
                                   WHEN 'expelled' THEN 4
                                   ELSE 5
                               END,
                               is_primary DESC,
                               id DESC
                       ) AS rn
                FROM student_enrollments
            ) WHERE rn = 1
        )
    """)
    print(f"   видалено зайвих записів: {_dup_count}")

# Після чистки — UNIQUE INDEX, який унеможливить нові дублікати.
# Синтаксис partial — не заважає старим SQLite, які ігнорують
# ON CONFLICT без явного рівня; тут достатньо звичайного UNIQUE.
try:
    cur.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_enrollments_unique_student_group
        ON student_enrollments (student_id, group_id)
    """)
except sqlite3.IntegrityError as e:
    print(f"⚠️ Не вдалося створити UNIQUE INDEX на student_enrollments: {e}")



# ═══════════════════════════════════════════════════════════════════════
# 3. СИД-ДАНІ: ліцензії, ступені, галузі, спеціальності, ТЦК, admin
# ═══════════════════════════════════════════════════════════════════════

# ─── Ліцензії. v3: Львівська філія — базова ───
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
""")

# ─── Ступені (name_suffix — для назви групи) ───
cur.executescript("""
INSERT OR IGNORE INTO degree_levels (name_ua, name_en, is_active, name_suffix) VALUES ('Молодший бакалавр', 'Junior Bachelor', 0, 'мб');
INSERT OR IGNORE INTO degree_levels (name_ua, name_en, is_active, name_suffix) VALUES ('Бакалавр', 'Bachelor', 1, NULL);
INSERT OR IGNORE INTO degree_levels (name_ua, name_en, is_active, name_suffix) VALUES ('Магістр', 'Master Degree', 1, 'м');
INSERT OR IGNORE INTO degree_levels (name_ua, name_en, is_active, name_suffix) VALUES ('Доктор філософії', 'Doctor of Philosophy', 0, 'ф');
""")

# ─── Галузі знань (Постанова КМУ №1021 від 30.08.2024) ───
cur.executescript("""
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('A', 'Освіта', 'Education');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('B', 'Культура, мистецтво та гуманітарні науки', 'Culture, Arts and Humanities');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('C', 'Соціальні науки, журналістика та інформація', 'Social Sciences, Journalism and Information');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('D', 'Бізнес, адміністрування та право', 'Business, Administration and Law');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('E', 'Природничі науки, математика та статистика', 'Natural Sciences, Mathematics and Statistics');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('F', 'Інформаційні технології', 'Information Technology');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('G', 'Інженерія, виробництво та будівництво', 'Engineering, Manufacturing and Construction');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('H', 'Сільське, лісове, рибне господарство та ветеринарна медицина', 'Agriculture, Forestry, Fisheries and Veterinary Medicine');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('I', 'Охорона здоров’я та соціальне забезпечення', 'Health and Welfare');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('J', 'Транспорт та послуги', 'Transport and Services');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('K', 'Безпека та оборона', 'Security and Defence');
""")

# ─── Спеціальності (Постанова КМУ №1021 від 30.08.2024) ───
cur.executescript("""
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A1', 'Освітні науки', '0111 Education science', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A2', 'Дошкільна освіта', '0112 Training for pre-school teachers', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A3', 'Початкова освіта', '0113 Teacher training without subject specialisation', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A4', 'Середня освіта (за предметними спеціальностями)', '0114 Teacher training with subject specialisation', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A5', 'Професійна освіта (за спеціалізаціями)', '0114 Teacher training with subject specialisation', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A6', 'Спеціальна освіта (за спеціалізаціями)', '0113 Teacher training without subject specialisation', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A7', 'Фізична культура і спорт', '1014 Sports', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B1', 'Аудіовізуальне мистецтво та медіавиробництво', '0211 Audio-visual techniques and media production', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B2', 'Дизайн', '0212 Fashion, interior and industrial design', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B3', 'Декоративне мистецтво та ремесла', '0214 Handicrafts', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B4', 'Образотворче мистецтво та реставрація', '0213 Fine arts', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B5', 'Музичне мистецтво', '0215 Music and performing arts', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B6', 'Перформативні мистецтва', '0215 Music and performing arts', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B7', 'Релігієзнавство', '0221 Religion and theology', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B8', 'Богослов’я', '0221 Religion and theology', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B9', 'Історія та археологія', '0222 History and archaeology', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B10', 'Філософія', '0223 Philosophy and ethics', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B11', 'Філологія (за спеціалізаціями)', '0231 Language acquisition', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B12', 'Культурологія та музеєзнавство', '0314 Sociology and cultural studies', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B13', 'Бібліотечна, інформаційна та архівна справа', '0322 Library, information and archival studies', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B14', 'Організація соціокультурної діяльності', '0413 Management and administration', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C1', 'Економіка', '0311 Economics', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C2', 'Політологія', '0312 Political sciences and civics', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C3', 'Міжнародні відносини', '0312 Political sciences and civics', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C4', 'Психологія', '0313 Psychology', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C5', 'Соціологія', '0314 Sociology and cultural studies', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C6', 'Географія та регіональні студії', '0314 Sociology and cultural studies', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C7', 'Журналістика', '0321 Journalism and reporting', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D1', 'Облік і оподаткування', '0411 Accounting and taxation', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D2', 'Фінанси, банківська справа, страхування та фондовий ринок', '0412 Finance, banking and insurance', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D3', 'Менеджмент', '0413 Management and administration', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D4', 'Публічне управління та адміністрування', '0413 Management and administration', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D5', 'Маркетинг', '0414 Marketing and advertising', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D6', 'Секретарська та офісна справа', '0415 Secretarial and office work', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D7', 'Торгівля', '0416 Wholesale and retail sales', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D8', 'Право', '0421 Law', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D9', 'Міжнародне право', '0421 Law', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E1', 'Біологія та біохімія', '0511 Biology', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E2', 'Екологія', '0521 Environmental sciences', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E3', 'Хімія', '0531 Chemistry', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E4', 'Науки про Землю', '0532 Earth sciences', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E5', 'Фізика та астрономія', '0533 Physics', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E6', 'Прикладна фізика та наноматеріали', '0533 Physics', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E7', 'Математика', '0541 Mathematics', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E8', 'Статистика', '0542 Statistics', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F1', 'Прикладна математика', '0541 Mathematics', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F2', 'Інженерія програмного забезпечення', '0613 Software and applications development and analysis', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F3', 'Комп’ютерні науки', '0613 Software and applications development and analysis', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F4', 'Системний аналіз та наука про дані', '0688 Inter-disciplinary programmes and qualifications involving Information and Communication Technologies', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F5', 'Кібербезпека та захист інформації', '0612 Database and network design and administration', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F6', 'Інформаційні системи і технології', '0612 Database and network design and administration', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F7', 'Комп’ютерна інженерія', '0612 Database and network design and administration', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G1', 'Хімічні технології та інженерія', '0711 Chemical engineering and processes', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G2', 'Технології захисту навколишнього середовища', '0712 Environmental protection technology', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G3', 'Електрична інженерія', '0713 Electricity and energy', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G4', 'Енерговиробництво (за спеціалізацією)', '0713 Electricity and energy', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G5', 'Електроніка, електронні комунікації, приладобудування та радіотехніка', '0714 Electronics and automation', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G6', 'Інформаційно-вимірювальні технології', '0714 Electronics and automation', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G7', 'Автоматизація, комп’ютерно-інтегровані технології та робототехніка', '0714 Electronics and automation', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G8', 'Матеріалознавство', '0788 Inter-disciplinary programmes and qualifications involving engineering, manufacturing and construction', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G9', 'Прикладна механіка', '0715 Mechanics and metal trades', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G10', 'Металургія', '0715 Mechanics and metal trades', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G11', 'Машинобудування (за спеціалізаціями)', '0715 Mechanics and metal trades', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G12', 'Авіаційна та ракетно-космічна техніка', '0716 Motor vehicles, ships and aircraft', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G13', 'Харчові технології', '0721 Food processing', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G14', 'Деревообробні та меблеві технології', '0722 Materials (glass, paper, plastic and wood)', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G15', 'Технології легкої промисловості', '0723 Textiles (clothes, footwear and leather)', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G16', 'Гірництво та нафтогазові технології', '0724 Mining and extraction', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G17', 'Архітектура та містобудування', '0731 Architecture and town planning', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G18', 'Геодезія та землеустрій', '0532 Earth sciences', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G19', 'Будівництво та цивільна інженерія', '0732 Building and civil engineering', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G20', 'Видавництво та поліграфія', '0211 Audio-visual techniques and media production', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G21', 'Біотехнології та біоінженерія', '0588 Inter-disciplinary programmes and qualifications involving natural sciences, mathematics and statistics', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G22', 'Біомедична інженерія', '0588 Inter-disciplinary programmes and qualifications involving natural sciences, mathematics and statistics', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H1', 'Агрономія', '0811 Crop and livestock production', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H2', 'Тваринництво', '0811 Crop and livestock production', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H3', 'Садово-паркове господарство', '0812 Horticulture', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H4', 'Лісове господарство', '0821 Forestry', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H5', 'Водні біоресурси та аквакультура', '0831 Fisheries', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H6', 'Ветеринарна медицина', '0841 Veterinary', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H7', 'Агроінженерія', '0788 Inter-disciplinary programmes and qualifications involving engineering, manufacturing and construction', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I1', 'Стоматологія', '0911 Dental studies', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I2', 'Медицина', '0912 Medicine', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I3', 'Педіатрія', '0912 Medicine', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I4', 'Медична психологія', '0912 Medicine', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I5', 'Медсестринство (за спеціалізаціями)', '0913 Nursing and midwifery', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I6', 'Технології медичної діагностики та лікування (за спеціалізаціями)', '0914 Medical diagnostic and treatment technology', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I7', 'Терапія та реабілітація (за спеціалізаціями)', '0915 Therapy and rehabilitation', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I8', 'Фармація (за спеціалізаціями)', '0711 Chemical engineering and processes', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I9', 'Громадське здоров’я', '0988 Inter-disciplinary programmes and qualifications involving health and welfare', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I10', 'Соціальна робота та консультування', '0921 Care of the elderly and of disabled adults', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I11', 'Дитячі та молодіжні служби', '0922 Child care and youth services', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J1', 'Послуги краси', '1012 Hair and beauty services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J2', 'Готельно-ресторанна справа та кейтеринг', '1013 Hotel, restaurants and catering', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J3', 'Туризм та рекреація', '1015 Travel, tourism and leisure', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J4', 'Охорона праці', '1022 Occupational health and safety', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J5', 'Морський та внутрішній водний транспорт', '1041 Transport services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J6', 'Авіаційний транспорт', '1041 Transport services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J7', 'Залізничний транспорт', '1041 Transport services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J8', 'Автомобільний транспорт', '1041 Transport services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K1', 'Державна безпека', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K2', 'Безпека державного кордону', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K3', 'Національна безпека (за окремими сферами забезпечення і видами діяльності)***', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K4', 'Управління інформаційною безпекою', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K5', 'Військове управління (за видами збройних сил)', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K6', 'Забезпечення військ (сил)', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K7', 'Озброєння та військова техніка', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K8', 'Пожежна безпека', '1032 Protection of persons and property', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K9', 'Правоохоронна діяльність', '1032 Protection of persons and property', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K10', 'Цивільна безпека', '1032 Protection of persons and property', 'K', 1);
""")

# ─── 67 ТЦК та СП (v1) ───
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
for _name, _addr in TCK_OFFICES:
    cur.execute(
        "INSERT OR IGNORE INTO tck_offices (name_ua, address) VALUES (?, ?)",
        (_name, _addr)
    )
print(f"✅ tck_offices: {cur.execute('SELECT COUNT(*) FROM tck_offices').fetchone()[0]} установ")

# ─── admin-користувач ───
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
# 4. ОДНОРАЗОВЕ ЗАПОВНЕННЯ ДАНИХ ДЛЯ ІСНУЮЧИХ ЗАПИСІВ
# ═══════════════════════════════════════════════════════════════════════

# ── 4.1. is_base для ліцензій ──
cur.execute("""
    UPDATE institution_licenses SET is_base = 1
    WHERE short_name_ua = 'Львівська'
      AND (SELECT COUNT(*) FROM institution_licenses WHERE is_base = 1) = 0
""")

# ── 4.2. name_suffix для ступенів ──
cur.executescript("""
    UPDATE degree_levels SET name_suffix = 'мб' WHERE name_ua = 'Молодший бакалавр' AND name_suffix IS NULL;
    UPDATE degree_levels SET name_suffix = 'м'  WHERE name_ua = 'Магістр'           AND name_suffix IS NULL;
    UPDATE degree_levels SET name_suffix = 'ф'  WHERE name_ua = 'Доктор філософії'  AND name_suffix IS NULL;
""")

# ── 4.3. ISCED-F у вже існуючих групах ──
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
    print(f"➕ Підтягнуто ISCED-F з довідника для {cur.rowcount} існуючих груп")

# ── 4.4. Автоматичне заповнення gender за по батькові ──
def detect_gender_from_middle_name(middle_name):
    if not middle_name:
        return None
    m = middle_name.strip().lower().replace('’', "'").replace('ʼ', "'")
    if m.endswith(('івна', 'ївна', 'овна', 'евна', 'євна', 'ична')):
        return 'F'
    return 'M'
    
# ── 4.5. diplomas: перенос ключа з (student_id) на (student_id, group_id) ──
# Для існуючих записів, де group_id порожнє, заповнюємо його основною
# групою студента (це відповідає старій поведінці: «один номер на
# студента» → «номер для primary-групи»).
cur.execute("""
    UPDATE diplomas
    SET group_id = (
        SELECT s.group_id FROM students s WHERE s.id = diplomas.student_id
    )
    WHERE group_id IS NULL
      AND EXISTS (
          SELECT 1 FROM students s
          WHERE s.id = diplomas.student_id AND s.group_id IS NOT NULL
      )
""")
if cur.rowcount:
    print(f"➕ diplomas: підтягнуто group_id для {cur.rowcount} записів")

# Тепер, коли всі існуючі записи мають group_id, можна створити
# UNIQUE INDEX. Він не заважатиме legacy-записам без group_id
# (напр. студент без групи) — partial index WHERE group_id IS NOT NULL.
cur.execute("""
    CREATE UNIQUE INDEX IF NOT EXISTS idx_diplomas_student_group_unique
    ON diplomas(student_id, group_id)
    WHERE group_id IS NOT NULL
""")

# Перевірка: скільки записів залишилось без group_id (зазвичай це
# студенти без групи — архівні, заморожені, без enrollment'ів).
_orphan = cur.execute(
    "SELECT COUNT(*) FROM diplomas WHERE group_id IS NULL"
).fetchone()[0]
if _orphan:
    print(f"⚠️ diplomas: {_orphan} записів без group_id "
          f"(студенти без групи) — вони не будуть показані в manage_diplomas "
          f"і не використовуватимуться при генерації документів")


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
    if updated < len(rows):
        print(f"   ⚠️ {len(rows) - updated} студентів без по батькові — стать "
              f"лишилась NULL (заповніть вручну в картці)")


# ═══════════════════════════════════════════════════════════════════════
# 5. ФІНАЛ
# ═══════════════════════════════════════════════════════════════════════
conn.commit()
conn.close()

print()
print("=" * 66)
print("  ✅ БД готова до роботи.")
print("=" * 66)
print()
print("Перевірки:")
print('  sqlite3 students.db ".schema degree_levels"                (має бути name_suffix)')
print('  sqlite3 students.db "SELECT name_ua, name_suffix FROM degree_levels"')
print('  sqlite3 students.db "SELECT COUNT(*) FROM specialties"     (має бути ~90)')
print('  sqlite3 students.db "SELECT COUNT(*) FROM tck_offices"     (має бути 66-67)')
print()
print("⚠️  Якщо користувач 'admin' щойно створений — стандартний пароль")
print("    'admin123'. Змініть його одразу після першого входу через")
print("    /admin/users/<id>/change-password.")
print()