import sqlite3

db = "e:/ea-cli/tmp_alembic_autogen.db"
con = sqlite3.connect(db)
cur = con.cursor()
print("Tables:")
for row in cur.execute("SELECT name FROM sqlite_master WHERE type='table'"):
    print(" -", row[0])
    cols = cur.execute(f"PRAGMA table_info({row[0]})").fetchall()
    for col in cols:
        print(f"    {col[1]}: {col[2]}")
con.close()
