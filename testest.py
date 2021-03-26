import sqlite3

conn = sqlite3.connect("testtest.db", isolation_level=None)
c = conn.cursor()