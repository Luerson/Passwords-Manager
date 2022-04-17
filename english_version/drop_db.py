import sqlite3

conn = sqlite3.connect("users.db")

conn.execute("DROP TABLE users")
conn.execute("DROP TABLE passwords")

conn.close()