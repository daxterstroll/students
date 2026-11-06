"""
Одноразова міграція: приводить усі збережені номери телефонів
у students і pending_students до канонічного формату +380XXXXXXXXX.
"""
import os
import sqlite3
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from routes.utils import normalize_phone_ua

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'students.db')

conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row

total_changed = 0

for table in ('students', 'pending_students'):
    try:
        rows = conn.execute(
            f"SELECT id, phone, phone_backup FROM {table}"
        ).fetchall()
    except sqlite3.OperationalError:
        print(f"⚠️  Таблиця {table} відсутня — пропускаю")
        continue

    changed = 0
    for r in rows:
        old_phone = r['phone']
        old_backup = r['phone_backup']
        new_phone = normalize_phone_ua(old_phone)
        new_backup = normalize_phone_ua(old_backup)
        if new_phone != old_phone or new_backup != old_backup:
            conn.execute(
                f"UPDATE {table} SET phone = ?, phone_backup = ? WHERE id = ?",
                (new_phone, new_backup, r['id'])
            )
            changed += 1
    print(f"  {table}: нормалізовано {changed} з {len(rows)}")
    total_changed += changed

conn.commit()
conn.close()
print(f"✅ Готово. Всього змінено: {total_changed}")