import sqlite3

import win32com.client


class import_list:


    conn = sqlite3.connect("stock_list.db", isolation_level=None)
    c = conn.cursor()
    def run(self):
        self.c.execute("select code from stock_list")
        listCode = self.c.fetchall()
        codes = []
        for i in range(0,len(listCode)):
            print(listCode[i][0])
            codes.append(listCode[i][0])

        print(codes)

        return codes


class stock_list:
    # 연결 여부 체크
    print('dkkkkkkkkkkkkkkk')
    def run(self):
        print("sssssssssssssssssss")

        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            exit()

        # 종목코드 리스트 구하기
        objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        codeList = objCpCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = objCpCodeMgr.GetStockListByMarket(2)  # 코스닥

        conn = sqlite3.connect("stock_list.db", isolation_level=None)
        c = conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS stock_list (code, name)")
        for i, code in enumerate(codeList):


            secondCode = objCpCodeMgr.GetStockSectionKind(code)
            name = objCpCodeMgr.CodeToName(code)
            stdPrice = objCpCodeMgr.GetStockStdPrice(code)
            c.execute("INSERT OR IGNORE INTO stock_list VALUES( ?, ?)",
                      ((code,name)))
            print(name)


        print("코스닥 종목코드", len(codeList2))
        for i, code in enumerate(codeList2):
            secondCode = objCpCodeMgr.GetStockSectionKind(code)
            name = objCpCodeMgr.CodeToName(code)
            stdPrice = objCpCodeMgr.GetStockStdPrice(code)
            c.execute("INSERT OR IGNORE INTO stock_list VALUES( ?, ?)",
                      ((code, name)))
            print(i)


        print("거래소 + 코스닥 종목코드 ", len(codeList) + len(codeList2))
        c.close()


