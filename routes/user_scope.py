"""
Розширений доступ користувача до контингенту:
групи + спеціальності + ступені → множина group_id.
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


def get_allowed_group_ids(conn, user_id) -> list:
    """
    Повертає список id груп, доступних користувачу:
    - прямі user_groups
    - усі групи з specialty_code ∈ user_specialties
    - усі групи з degree_level ∈ user_degree_levels
    """
    ensure_user_scope_tables(conn)
    ids = set()

    for r in conn.execute(
        "SELECT group_id FROM user_groups WHERE user_id = ?", (user_id,)
    ).fetchall():
        ids.add(r[0] if not hasattr(r, 'keys') else r['group_id'])

    specs = [
        r[0] if not hasattr(r, 'keys') else r['specialty_code']
        for r in conn.execute(
            "SELECT specialty_code FROM user_specialties WHERE user_id = ?",
            (user_id,),
        ).fetchall()
    ]
    if specs:
        ph = ','.join('?' * len(specs))
        for r in conn.execute(
            f"SELECT id FROM groups WHERE specialty_code IN ({ph}) AND COALESCE(archived,0)=0",
            specs,
        ).fetchall():
            ids.add(r[0] if not hasattr(r, 'keys') else r['id'])

    degrees = [
        r[0] if not hasattr(r, 'keys') else r['degree_level']
        for r in conn.execute(
            "SELECT degree_level FROM user_degree_levels WHERE user_id = ?",
            (user_id,),
        ).fetchall()
    ]
    if degrees:
        ph = ','.join('?' * len(degrees))
        for r in conn.execute(
            f"SELECT id FROM groups WHERE degree_level IN ({ph}) AND COALESCE(archived,0)=0",
            degrees,
        ).fetchall():
            ids.add(r[0] if not hasattr(r, 'keys') else r['id'])

    return sorted(ids)
