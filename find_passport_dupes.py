# find_passport_dupes.py
"""
Одноразовий скрипт: знайти дублікати паспортів у БД.

Запуск з кореня проєкту:
    python find_passport_dupes.py
"""
import sqlite3
import os

from routes.passport_types import (
    normalize_passport_series,
    normalize_passport_number,
)

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'students.db')

conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row

# Ті самі функції, що їх реєструє routes/db.py
conn.create_function(
    "NORM_SERIES", 1,
    lambda s: normalize_passport_series(s) or '',
)
conn.create_function(
    "NORM_PASSPORT", 2,
    lambda n, t: normalize_passport_number(n, t) or '',
)

dupes = conn.execute("""
    SELECT
        student_id,
        COALESCE(group_id, -1) AS grp,
        document_type,
        NORM_SERIES(series) AS norm_series,
        NORM_PASSPORT(number, document_type) AS norm_number,
        COUNT(*) AS cnt,
        GROUP_CONCAT(id) AS ids,
        GROUP_CONCAT(number) AS numbers
    FROM passport_documents
    GROUP BY student_id, COALESCE(group_id, -1), document_type,
             NORM_SERIES(series), NORM_PASSPORT(number, document_type)
    HAVING COUNT(*) > 1
    ORDER BY student_id, grp
""").fetchall()

if not dupes:
    print("Дублікатів не знайдено.")
else:
    print(f"Знайдено {len(dupes)} груп дублікатів:\n")
    for d in dupes:
        print(f"  student_id = {d['student_id']}   group_id = {d['grp']}")
        print(f"  тип: {d['document_type']}")
        print(f"  нормалізовано: серія={d['norm_series']!r}  номер={d['norm_number']!r}")
        print(f"  id:      {d['ids']}")
        print(f"  номери:  {d['numbers']}")
        print()

conn.close()