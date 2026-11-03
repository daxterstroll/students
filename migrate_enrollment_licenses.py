# migrate_enrollment_licenses.py — положить в корень и запустить один раз
import sqlite3, os

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'students.db')
conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row

# 1. Добавить колонку
try:
    conn.execute("ALTER TABLE student_enrollments ADD COLUMN license_id INTEGER "
                 "REFERENCES institution_licenses(id)")
    print("✓ колонка student_enrollments.license_id добавлена")
except sqlite3.OperationalError as e:
    if 'duplicate column' in str(e).lower():
        print("→ колонка уже существует, пропускаем")
    else:
        raise

# 2. Backfill: копируем students.license_id в каждый enrollment
#    (для тех, у кого ещё NULL)
conn.execute("""
    UPDATE student_enrollments
    SET license_id = (SELECT license_id FROM students WHERE id = student_id)
    WHERE license_id IS NULL
""")
print(f"✓ backfill: {conn.total_changes} enrollments обновлено")

# 3. Синхронизировать students.license_id с primary-enrollment
#    (чтобы старая логика, читающая students.license_id, была консистентна)
conn.execute("""
    UPDATE students
    SET license_id = (
        SELECT e.license_id FROM student_enrollments e
        WHERE e.student_id = students.id AND e.is_primary = 1 AND e.status = 'active'
        LIMIT 1
    )
    WHERE id IN (
        SELECT student_id FROM student_enrollments
        WHERE is_primary = 1 AND status = 'active' AND license_id IS NOT NULL
    )
""")
print(f"✓ students.license_id синхронизирован с primary-enrollment")

conn.commit()
conn.close()
print("Готово.")