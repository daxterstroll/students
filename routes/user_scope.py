"""
Розширений доступ користувача до контингенту.

Правила:
  - Явно обрані групи (user_groups) завжди доступні.
  - Критерії «спеціальність / ступінь / галузь» між собою комбінуються через
    ПЕРЕТИН (AND): «Бакалавр» + «галузь F» → лише бакалаврські групи галузі F.
  - Якщо обрано лише один тип критерію — фільтр тільки за ним.
  - Якщо жоден критерій і жодна група не обрані — порожній список (немає доступу).
  - Адмін обходить цей механізм на рівні permission_required / is_admin.
"""
from __future__ import annotations


def ensure_user_scope_tables(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_specialties (
            user_id INTEGER NOT NULL,
            specialty_code TEXT NOT NULL,
            PRIMARY KEY (user_id, specialty_code),
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_degree_levels (
            user_id INTEGER NOT NULL,
            degree_level TEXT NOT NULL,
            PRIMARY KEY (user_id, degree_level),
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_knowledge_areas (
            user_id INTEGER NOT NULL,
            knowledge_field_code TEXT NOT NULL,
            PRIMARY KEY (user_id, knowledge_field_code),
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        )
    """)


def _row_val(r, key, idx=0):
    if hasattr(r, 'keys'):
        return r[key]
    return r[idx]


def get_allowed_group_ids(conn, user_id) -> list:
    ensure_user_scope_tables(conn)

    # 1) Явні групи
    explicit = {
        int(_row_val(r, 'group_id'))
        for r in conn.execute(
            "SELECT group_id FROM user_groups WHERE user_id = ?", (user_id,)
        ).fetchall()
    }

    specs = [
        _row_val(r, 'specialty_code')
        for r in conn.execute(
            "SELECT specialty_code FROM user_specialties WHERE user_id = ?",
            (user_id,),
        ).fetchall()
    ]
    degrees = [
        _row_val(r, 'degree_level')
        for r in conn.execute(
            "SELECT degree_level FROM user_degree_levels WHERE user_id = ?",
            (user_id,),
        ).fetchall()
    ]
    k_codes = [
        _row_val(r, 'knowledge_field_code')
        for r in conn.execute(
            "SELECT knowledge_field_code FROM user_knowledge_areas WHERE user_id = ?",
            (user_id,),
        ).fetchall()
    ]

    # 2) Критерії каталогу — AND між типами, що задані
    # candidates is None, поки жоден фільтр не застосований
    candidates = None

    def apply_filter(ids_set: set):
        nonlocal candidates
        if candidates is None:
            candidates = set(ids_set)
        else:
            candidates &= ids_set

    if specs:
        ph = ','.join('?' * len(specs))
        rows = conn.execute(
            f"SELECT id FROM groups WHERE specialty_code IN ({ph}) AND COALESCE(archived,0)=0",
            specs,
        ).fetchall()
        apply_filter({int(_row_val(r, 'id')) for r in rows})

    if degrees:
        ph = ','.join('?' * len(degrees))
        rows = conn.execute(
            f"SELECT id FROM groups WHERE degree_level IN ({ph}) AND COALESCE(archived,0)=0",
            degrees,
        ).fetchall()
        apply_filter({int(_row_val(r, 'id')) for r in rows})

    if k_codes:
        ph = ','.join('?' * len(k_codes))
        # групи через specialty → knowledge_field_code
        rows1 = conn.execute(
            f"""
            SELECT g.id FROM groups g
            JOIN specialties s ON s.code = g.specialty_code
            WHERE s.knowledge_field_code IN ({ph}) AND COALESCE(g.archived,0)=0
            """,
            k_codes,
        ).fetchall()
        # або groups.knowledge_area збігається з назвою/кодом галузі
        rows2 = conn.execute(
            f"""
            SELECT g.id FROM groups g
            JOIN knowledge_fields kf ON (
                g.knowledge_area = kf.name_ua OR g.knowledge_area = kf.code
            )
            WHERE kf.code IN ({ph}) AND COALESCE(g.archived,0)=0
            """,
            k_codes,
        ).fetchall()
        k_ids = {int(_row_val(r, 'id')) for r in rows1} | {
            int(_row_val(r, 'id')) for r in rows2
        }
        apply_filter(k_ids)

    result = set(explicit)
    if candidates is not None:
        result |= candidates

    return sorted(result)
