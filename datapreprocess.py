import sqlite3

conn = sqlite3.connect("stock_list.db", isolation_level=None)
c = conn.cursor()
c.execute("select code,name from stock_list")
listCode = c.fetchall()
codes = []
for i in range(2400,len(listCode)):
    codes.append([listCode[i][0],listCode[i][1]])
conn.close()

conn = sqlite3.connect("stock_pyun.db", isolation_level=None)
c = conn.cursor()
c.execute("CREATE TABLE IF NOT EXISTS stock_pyun (code, name)")
for codep,namep in codes:
    print(codep+",,,,"+namep)
    c.execute("INSERT OR IGNORE INTO stock_pyun VALUES( ?, ?)",
              ((codep, namep)))

