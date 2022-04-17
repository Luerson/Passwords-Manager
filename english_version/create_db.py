import sqlite3

conn = sqlite3.connect("users.db")

conn.execute("""CREATE TABLE users (id INTEGER NOT NULL, username TEXT NOT NULL
             ,password BINARY NOT NULL, email BINARY NOT NULL, key BINARY NOT NULL, PRIMARY KEY(id))""")
conn.commit()

conn.execute("""CREATE TABLE passwords (user_id INTEGER NOT NULL, title TEXT NOT NULL
                , password BINARY NOT NULL, FOREIGN KEY (user_id) REFERENCES users(id))""")
conn.commit()

conn.close()
