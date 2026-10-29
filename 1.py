from routes.db import get_db
conn = get_db()
row = conn.execute(
    "SELECT status_badge FROM military WHERE student_id = ?", (123,)
).fetchone()
val = row['status_badge'] if row else None
print("status_badge =", repr(val))
if val:
    print("коди символів:", [hex(ord(c)) for c in val if not c.isalpha() and c != ' '])
conn.close()