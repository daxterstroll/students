"""
Одноразова міграція: картотека викладачів замість session_sheet_batches.

Що робить:
  1. Створює таблицю teachers.
  2. Додає session_sheets.teacher_ref_id.
  3. Перезбирає session_sheets з новим статусом 'in_progress'.
  4. Збирає унікальні teacher_name → створює записи в teachers з токеном.
  5. Прив'язує session_sheets.teacher_ref_id.
  6. Видаляє старі таблиці пакетів (session_sheet_batches, session_sheet_batch_items).

Запуск: python migrate_teachers.py
"""
import os
import secrets
import shutil
import sqlite3
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'students.db')

# Бекап
if os.path.exists(DB_PATH):
    stamp = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    backup = f"{DB_PATH}.bak-teachers-{stamp}"
    shutil.copy2(DB_PATH, backup)
    print(f"💾 Бекап: {os.path.basename(backup)}")

conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row
cur = conn.cursor()

# ─── 1. teachers ──────────────────────────────────────────────
cur.executescript("""
    CREATE TABLE IF NOT EXISTS teachers (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        full_name       TEXT NOT NULL,
        short_name      TEXT,
        email           TEXT,
        phone           TEXT,
        department      TEXT,
        user_id         INTEGER,
        token           TEXT UNIQUE NOT NULL,
        token_created_at TEXT DEFAULT (datetime('now','localtime')),
        is_active       INTEGER NOT NULL DEFAULT 1,
        notes           TEXT,
        created_by      TEXT,
        created_at      TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
    );
    CREATE INDEX IF NOT EXISTS idx_teachers_token ON teachers(token);
    CREATE INDEX IF NOT EXISTS idx_teachers_active ON teachers(is_active);
""")
print("✅ teachers створено")

# ─── 2. session_sheets.teacher_ref_id ─────────────────────────
existing_cols = {r[1] for r in cur.execute('PRAGMA table_info(session_sheets)').fetchall()}
if 'teacher_ref_id' not in existing_cols:
    cur.execute('ALTER TABLE session_sheets ADD COLUMN teacher_ref_id INTEGER REFERENCES teachers(id)')
    print("➕ session_sheets.teacher_ref_id")

# ─── 3. Перезбірка session_sheets зі статусом in_progress ─────
_sql_row = cur.execute(
    "SELECT sql FROM sqlite_master WHERE type='table' AND name='session_sheets'"
).fetchone()
needs_rebuild = _sql_row and 'in_progress' not in (_sql_row[0] or '')

if needs_rebuild:
    print("🔁 session_sheets: додаю статус in_progress — перезбираю таблицю")
    cur.execute('DROP TABLE IF EXISTS session_sheets_new')
    conn.commit()
    cur.execute("PRAGMA foreign_keys = OFF")
    try:
        # Тимчасово відв'язуємо дочірні таблиці
        has_items = cur.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='session_sheet_batch_items'"
        ).fetchone()

        cur.executescript("""
            CREATE TABLE "session_sheets_new" (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id INTEGER NOT NULL,
                discipline_name TEXT NOT NULL,
                work_topic TEXT,
                entity_kind TEXT NOT NULL DEFAULT 'subject'
                    CHECK (entity_kind IN ('subject','practice','coursework','attestation')),
                control_form TEXT NOT NULL DEFAULT 'Залік'
                    CHECK (control_form IN ('Залік','Екзамен')),
                hours_credits TEXT,
                teacher_name TEXT,
                teacher_ref_id INTEGER REFERENCES teachers(id),
                exam_date TEXT,
                statement_number TEXT,
                license_id INTEGER,
                status TEXT NOT NULL DEFAULT 'draft'
                    CHECK (status IN ('draft','sent','in_progress','submitted','approved','locked')),
                token TEXT UNIQUE,
                token_expires_at TEXT,
                linked_entity_type TEXT
                    CHECK (linked_entity_type IS NULL OR linked_entity_type IN
                           ('subject','practice','coursework','attestation')),
                linked_entity_id INTEGER,
                linked_at TEXT,
                linked_by TEXT,
                created_by TEXT,
                created_at TEXT DEFAULT (datetime('now','localtime')),
                submitted_at TEXT,
                approved_at TEXT,
                approved_by TEXT,
                FOREIGN KEY (session_id) REFERENCES exam_sessions(id) ON DELETE CASCADE,
                FOREIGN KEY (license_id) REFERENCES institution_licenses(id)
            );
            INSERT INTO session_sheets_new
                (id, session_id, discipline_name, work_topic, entity_kind,
                 control_form, hours_credits, teacher_name, teacher_ref_id,
                 exam_date, statement_number, license_id, status,
                 token, token_expires_at,
                 linked_entity_type, linked_entity_id, linked_at, linked_by,
                 created_by, created_at, submitted_at, approved_at, approved_by)
            SELECT
                id, session_id, discipline_name, work_topic, entity_kind,
                control_form, hours_credits, teacher_name,
                NULL,
                exam_date, statement_number, license_id, status,
                token, token_expires_at,
                linked_entity_type, linked_entity_id, linked_at, linked_by,
                created_by, created_at, submitted_at, approved_at, approved_by
            FROM session_sheets;
            DROP TABLE session_sheets;
            ALTER TABLE session_sheets_new RENAME TO session_sheets;
            CREATE INDEX IF NOT EXISTS idx_session_sheets_session ON session_sheets(session_id);
            CREATE INDEX IF NOT EXISTS idx_session_sheets_token ON session_sheets(token);
            CREATE INDEX IF NOT EXISTS idx_session_sheets_teacher ON session_sheets(teacher_ref_id);
        """)
        conn.commit()
        print("   ✅ перезбудовано")
    finally:
        cur.execute("PRAGMA foreign_keys = ON")
else:
    # Просто індекси
    cur.execute("CREATE INDEX IF NOT EXISTS idx_session_sheets_teacher ON session_sheets(teacher_ref_id)")

# ─── 4. Заповнюємо картотеку з наявних teacher_name ─────────
names = [
    r[0] for r in cur.execute("""
        SELECT DISTINCT teacher_name
        FROM session_sheets
        WHERE teacher_name IS NOT NULL AND teacher_name != ''
          AND teacher_ref_id IS NULL
    """).fetchall()
]

created = 0
for name in names:
    token = secrets.token_urlsafe(24)
    cur.execute("""
        INSERT INTO teachers (full_name, short_name, token, created_by)
        VALUES (?, ?, ?, 'migration')
    """, (name, name, token))
    tid = cur.lastrowid
    cur.execute("""
        UPDATE session_sheets SET teacher_ref_id = ?
        WHERE teacher_name = ? AND teacher_ref_id IS NULL
    """, (tid, name))
    created += 1

# ─── 5. Видаляємо старі пакети ────────────────────────────────
for t in ('session_sheet_batch_items', 'session_sheet_batches'):
    if cur.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (t,)
    ).fetchone():
        cur.execute(f'DROP TABLE "{t}"')
        print(f'🗑️  видалено {t}')

conn.commit()
conn.close()

print()
print("=" * 60)
print(f"✅ Створено записів у teachers: {created}")
print(f"✅ Прив'язано відомостей: {len(names)}")
print("=" * 60)