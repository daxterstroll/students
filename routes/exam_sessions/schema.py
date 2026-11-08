"""
Таблиці сесійних відомостей + user_knowledge_areas.
Викликати ensure_exam_session_tables(conn) при старті маршрутів / у init_db.
"""
from __future__ import annotations


def ensure_exam_session_tables(conn) -> None:
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_knowledge_areas (
            user_id INTEGER NOT NULL,
            knowledge_field_code TEXT NOT NULL,
            PRIMARY KEY (user_id, knowledge_field_code),
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS exam_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id INTEGER NOT NULL,
            academic_year TEXT NOT NULL,
            semester INTEGER NOT NULL CHECK (semester IN (1, 2, 3)),
            name TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'draft'
                CHECK (status IN ('draft', 'open', 'closed')),
            notes TEXT,
            created_by TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT,
            FOREIGN KEY (group_id) REFERENCES groups(id)
        )
    """)
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_exam_sessions_group ON exam_sessions(group_id)"
    )

    conn.execute("""
        CREATE TABLE IF NOT EXISTS session_sheets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            -- Вільна назва: предмета в базі ще може не бути
            discipline_name TEXT NOT NULL,
            work_topic TEXT,
            -- Мітка для майбутнього переносу
            entity_kind TEXT NOT NULL DEFAULT 'subject'
                CHECK (entity_kind IN ('subject', 'practice', 'coursework', 'attestation')),
            control_form TEXT NOT NULL DEFAULT 'Залік'
                CHECK (control_form IN ('Залік', 'Екзамен')),
            hours_credits TEXT,
            teacher_name TEXT,
            exam_date TEXT,
            statement_number TEXT,
            -- NULL = студенти групи без license_id; інакше лише ця ліцензія
            license_id INTEGER,
            status TEXT NOT NULL DEFAULT 'draft'
                CHECK (status IN ('draft', 'sent', 'submitted', 'approved', 'locked')),
            token TEXT UNIQUE,
            token_expires_at TEXT,
            -- Після переносу
            linked_entity_type TEXT
                CHECK (linked_entity_type IS NULL OR linked_entity_type IN
                       ('subject', 'practice', 'coursework', 'attestation')),
            linked_entity_id INTEGER,
            linked_at TEXT,
            linked_by TEXT,
            created_by TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            submitted_at TEXT,
            approved_at TEXT,
            approved_by TEXT,
            FOREIGN KEY (session_id) REFERENCES exam_sessions(id) ON DELETE CASCADE,
            FOREIGN KEY (license_id) REFERENCES institution_licenses(id)
        )
    """)
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_session_sheets_session ON session_sheets(session_id)"
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_session_sheets_token ON session_sheets(token)"
    )

    conn.execute("""
        CREATE TABLE IF NOT EXISTS session_sheet_students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sheet_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            -- Snapshot на момент формування відомості
            full_name TEXT NOT NULL,
            edebo_code TEXT,
            work_topic TEXT,
            position INTEGER NOT NULL DEFAULT 0,
            points_mid INTEGER,
            points_current INTEGER,
            points_exam INTEGER,
            points INTEGER,
            national_grade TEXT,
            ects TEXT,
            teacher_sign TEXT,
            UNIQUE (sheet_id, student_id),
            FOREIGN KEY (sheet_id) REFERENCES session_sheets(id) ON DELETE CASCADE,
            FOREIGN KEY (student_id) REFERENCES students(id)
        )
    """)
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_sheet_students_sheet ON session_sheet_students(sheet_id)"
    )

    try:
        conn.execute("ALTER TABLE session_sheets ADD COLUMN work_topic TEXT")
    except Exception:
        pass
    try:
        conn.execute("ALTER TABLE session_sheet_students ADD COLUMN work_topic TEXT")
    except Exception:
        pass
    for col in ('points_mid', 'points_current', 'points_exam'):
        try:
            conn.execute(f"ALTER TABLE session_sheet_students ADD COLUMN {col} INTEGER")
        except Exception:
            pass
    try:
        conn.commit()
    except Exception:
        pass
