import win32com.client
import sqlite3


def ReqeustData(obj):
    # 데이터 요청
    obj.BlockRequest()

    # 통신 결과 확인
    rqStatus = obj.GetDibStatus()
    rqRet = obj.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        return False

    conn = sqlite3.connect("stock_kind.db", isolation_level=None)
    c = conn.cursor()
    # DB 연결
    code = ['A005390']

    for codenum in code:

        c.execute("CREATE TABLE IF NOT EXISTS " + code[0] +
                  "(date integer, open integer, high integer, low integer,close integer, diff integer, vol integer)")
        # sql문 실행 - 테이블 생성
        # 일자별 정보 데이터 처리
        count = obj.GetHeaderValue(1)  # 데이터 개수
        for i in range(count):
            date = obj.GetDataValue(0, i)  # 일자
            open = obj.GetDataValue(1, i)  # 시가
            high = obj.GetDataValue(2, i)  # 고가
            low = obj.GetDataValue(3, i)  # 저가
            close = obj.GetDataValue(4, i)  # 종가
            diff = obj.GetDataValue(5, i)  # 종가
            vol = obj.GetDataValue(6, i)  # 종가
            print(date, open, high, low, close, diff, vol)
            c.execute("INSERT OR IGNORE INTO "+ code[0] +" VALUES( ?, ?, ?, ?, ?, ?, ?)", ((date,open,high,low,close,diff,vol)))
            return True


# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 일자별 object 구하기
objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
objStockWeek.SetInputValue(0, 'A005930')  # 종목 코드 - 삼성전자

# 최초 데이터 요청
ret = ReqeustData(objStockWeek)
if ret == False:
    exit()

# 연속 데이터 요청
# 예제는 5번만 연속 통신 하도록 함.
NextCount = 1
while objStockWeek.Continue:  # 연속 조회처리
    NextCount += 1;
    if (NextCount > 1000000):
        break
    ret = ReqeustData(objStockWeek)
    if ret == False:
        exit()