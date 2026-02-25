import sqlite3
from werkzeug.security import generate_password_hash

conn = sqlite3.connect('students.db')
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
	FOREIGN KEY("student_id") REFERENCES "students"("id"),
	FOREIGN KEY("subject_id") REFERENCES "subjects"("id")
);
CREATE TABLE IF NOT EXISTS "groups" (
	"id"	INTEGER,
	"name"	TEXT NOT NULL,
	"start_year"	INTEGER NOT NULL,
	"study_form"	TEXT NOT NULL CHECK("study_form" IN ('Денна', 'Заочна')),
	"program_credits"	INTEGER NOT NULL CHECK("program_credits" IN (180, 240)),
	"qualification_name"	TEXT,
	"degree_level"	TEXT,
	"specialty"	TEXT,
	"educational_program"	TEXT,
	"knowledge_area"	TEXT,
	"qualification_name_en"	TEXT,
	"degree_level_en"	TEXT,
	"specialty_en"	TEXT,
	"educational_program_en"	TEXT,
	"knowledge_area_en"	TEXT,
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
	PRIMARY KEY("id" AUTOINCREMENT)
);
""")

for u,p,r in [('admin','admin123','admin')]:
    cur.execute("INSERT OR IGNORE INTO users (username,password_hash,role) VALUES (?, ?, ?)",
                (u, generate_password_hash(p), r))
conn.commit()
conn.close()
print("✅ DB и пользователи созданы.")
