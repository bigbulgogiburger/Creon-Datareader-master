import sqlite3

conn = sqlite3.connect("stock_list.db", isolation_level=None)
c = conn.cursor()
c.execute("select code,name from stock_list")
listCode = c.fetchall()
codes = []
for i in range(2400,len(listCode)):
    codes.append([listCode[i][0],listCode[i][1]])
conn.close()
#
# conn = sqlite3.connect("stock_pyun.db", isolation_level=None)
# c = conn.cursor()
# codes_for_pyun =[]
# c.execute("CREATE TABLE IF NOT EXISTS stock_pyun (code primary key, name)")
# for codep,namep in codes:
#     print(namep)
#     codes_for_pyun.append((codep,namep))
#
# c.executemany("INSERT OR IGNORE INTO stock_pyun VALUES( ?, ?)",
#               codes_for_pyun)


#백업 만들어 놓음.
conn = sqlite3.connect("stock_pyun_backup.db", isolation_level=None)
c = conn.cursor()
c.execute("CREATE TABLE IF NOT EXISTS stock_pyun_backup (code primary key, name)")
codes_for_pyun =[]

for codep,namep in codes:
    print(namep)
    codes_for_pyun.append((codep, namep))
c.executemany("INSERT OR IGNORE INTO stock_pyun_backup VALUES( ?, ?)",
          codes_for_pyun)

