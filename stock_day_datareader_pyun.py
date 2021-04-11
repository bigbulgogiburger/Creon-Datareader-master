import win32com.client
import sqlite3
import datetime
import pandas as pd
import numpy as np

def ReqeustData(obj):
    # 데이터 요청
    obj.BlockRequest()

    # 통신 결과 확인
    rqStatus = obj.GetDibStatus()
    rqRet = obj.GetDibMsg1()

    if rqStatus != 0:
        return False
    ## 현재 받은 리스트 로드하기
    connn = sqlite3.connect("stock_db(day)_2401~_dh.db", isolation_level=None)
    c2 = connn.cursor()
    # DB 연결
    print(obj.GetHeaderValue(0))
    c2.execute("CREATE TABLE IF NOT EXISTS " + obj.GetHeaderValue(0) +
              " (day date primary key, cur_pr integer,high_pr integer, low_pr integer, clo_pr integer, pr_diff integer, acc_vol integer"
              ", for_stor integer, for_stor_diff integer, for perc integer, com_buy_vol integer, oot_cur_pr integer"
              ", oot_high_pr, oot_low_pr, oot_clo_pr, oot_pr_diff, oot_vol)")
    # sql문 실행 - 테이블 생성
    # 일자별 정보 데이터 처리
    count = obj.GetHeaderValue(1)  # 데이터 개수
    print(count)
    for i in range(count):
        day = obj.GetDataValue(0, i)  # 일자
        cur_pr = obj.GetDataValue(1, i)  # 시가
        high_pr = obj.GetDataValue(2, i)  # 고가
        low_pr = obj.GetDataValue(3, i)  # 저가
        clo_pr = obj.GetDataValue(4, i)  # 종가
        pr_diff = obj.GetDataValue(5, i)  # 전일대비
        acc_vol = obj.GetDataValue(6, i)  # 누적거래량
        for_stor = obj.GetDataValue(7, i)  # 외인보유
        for_stor_diff = obj.GetDataValue(8, i)  # 외인보유전일대비
        for_perc = obj.GetDataValue(9, i)  # 외인비중
        com_buy_vol = obj.GetDataValue(12, i)  # 기관순매수수량
        oot_cur_pr = obj.GetDataValue(13, i)  # 시간외단일가시가
        oot_high_pr = obj.GetDataValue(14, i)  # 시간외단일가고가
        oot_low_pr = obj.GetDataValue(15, i)  # 시간외단일가저가
        oot_clo_pr = obj.GetDataValue(16, i)  # 시간외단일가종가
        oot_pr_diff = obj.GetDataValue(18, i)  # 시간외단일가전일대비
        oot_vol = obj.GetDataValue(19, i)  # 시간외단일가거래량

        c2.execute("INSERT OR IGNORE INTO "+ obj.GetHeaderValue(0) +" VALUES( ?, ?, ?, ?, ?, ?, ?,?,?,?,?,?,?,?,?,?,?)",
                  ((day,cur_pr,high_pr,low_pr,clo_pr,pr_diff,acc_vol,for_stor,for_stor_diff,for_perc,com_buy_vol,oot_cur_pr,oot_high_pr,oot_low_pr,oot_clo_pr,oot_pr_diff,oot_vol)))

    return True

def store_list(obj):
    conn = sqlite3.connect("stock(day).db", isolation_level=None)
    c = conn.cursor()



# 연결 여부 체크

class stock_day_collector:

    def run(self,codelist):
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            exit()


        # 일자별 object 구하기
        objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
        ## 현재 받은 리스트 로드하기
        conn = sqlite3.connect("stock_pyun.db", isolation_level=None)
        c = conn.cursor()
        time1 = datetime.datetime.now

        for codenum in codelist:

            objStockWeek.SetInputValue(0, codenum)  # 종목 코드 - 삼성전자

            # 최초 데이터 요청
            ret = ReqeustData(objStockWeek)
            if ret == False:
                exit()

            # 연속 데이터 요청
            # 예제는 10000000번만 연속 통신 하도록 함.
            # 해당 while문을 지우면 중복하는 데이터를 받지 않는다.(최신 36일치만 받음)
            NextCount = 1
            # while objStockWeek.Continue:  # 연속 조회처리
            #     NextCount += 1;
            #     if (NextCount > 100000000):
            #         break
            #     ret = ReqeustData(objStockWeek)
            #     if ret == False:
            #         exit()
            # c.execute("DELETE FROM stock_pyun WHERE code = ?", (codenum,))
            # print(codenum,"추가 완료 : " , datetime.datetime.now)
