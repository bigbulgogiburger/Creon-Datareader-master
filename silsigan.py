# 주식 실시간 데이터를 조회하여 DB에 저장하는 코드

import sys
from PyQt5.QtWidgets import *
import win32com.client
import sqlite3
import time
import pandas as pd
import numpy as np

# 복수 종목 실시간 조회 샘플 (조회는 없고 실시간만 있음)
class CpEvent:
    def set_params(self, client):
        self.client = client

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 코드
        name = self.client.GetHeaderValue(1)  # 종목명
        diff = self.client.GetHeaderValue(2)  # 전일대비
        cur_price = self.client.GetHeaderValue(4)  #시가
        high_price = self.client.GetHeaderValue(5)  # 고가
        low_price = self.client.GetHeaderValue(6)  # 저가
        sell_call = self.client.GetHeaderValue(7)  # 매도호가
        buy_call = self.client.GetHeaderValue(8)  # 매수호가
        acc_vol = self.client.GetHeaderValue(9)  # 누적거래량
        pred_price = self.client.GetHeaderValue(13)  # 현재가 또는 예상체결가
        deal_state = self.client.GetHeaderValue(14)  # 체결상태(체결가 방식)
        acc_sell_deal_vol = self.client.GetHeaderValue(15)  # 누적매도체결수량(체결가방식)
        acc_buy_deal_vol = self.client.GetHeaderValue(16)  # 누적매수체결수량(체결가방식)
        moment_deal_vol = self.client.GetHeaderValue(17)  # 순간체결수량
        timess1 = time.strftime('%Y%m%d')
        date_time_sec= timess1 + str(self.client.GetHeaderValue(18))  # 시간(초)
        exFlag = self.client.GetHeaderValue(19)  # 예상체결가구분플래그
        market_diff_flag = self.client.GetHeaderValue(20)  # 장구분플래그

        conn = sqlite3.connect("stock(cur).db", isolation_level=None)
        c = conn.cursor()

        c.execute("CREATE TABLE IF NOT EXISTS " + code +
                  " (diff real, cur_price integer, high_price integer, low_price integer"
                  ", sell_call integer, buy_call integer, acc_vol integer, pred_price integer, deal_state text, acc_sell_deal_vol integer"
                  ", acc_buy_deal_vol integer , moment_deal_vol integer ,date_time_sec text, exFlag text, market_diff_flag text )")
        # sql문 실행 - 테이블 생성
        c.execute(
            "INSERT OR IGNORE INTO " + code + " VALUES( ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            ((diff,cur_price,high_price,low_price,sell_call,buy_call,acc_vol,pred_price,deal_state,acc_sell_deal_vol,acc_buy_deal_vol,moment_deal_vol,date_time_sec,exFlag,market_diff_flag)))


        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", code,name,diff,cur_price,high_price,low_price,sell_call,buy_call,acc_vol,pred_price,deal_state,acc_sell_deal_vol
                  ,acc_buy_deal_vol,moment_deal_vol,exFlag,market_diff_flag)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", code,name,diff,cur_price,high_price,low_price,sell_call,buy_call,acc_vol,pred_price,deal_state,acc_sell_deal_vol
                  ,acc_buy_deal_vol,moment_deal_vol,exFlag,market_diff_flag)


class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()


class CpMarketEye:
    def Request(self, code, rqField):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        # 관심종목 객체 구하기
        objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        # 요청 필드 세팅 - 종목코드, 종목명, 시간, 대비부호, 대비, 현재가, 거래량
        # rqField = [0,17, 1,2,3,4,10]
        objRq.SetInputValue(0, rqField)  # 요청 필드
        objRq.SetInputValue(1, code)  # 종목코드 or 종목코드 리스트
        print(code)
        objRq.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objRq.GetDibStatus()
        rqRet = objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        conn = sqlite3.connect("stock_cur.db", isolation_level=None)
        c = conn.cursor()
        # DB 연결


        # sql문 실행 - 테이블 생성
        # 일자별 정보 데이터 처리

        cnt = objRq.GetHeaderValue(2)
        print(cnt)
        for i in range(cnt):
            rpCode = objRq.GetDataValue(0, i)  # 코드
            rpName = objRq.GetDataValue(1, i)  # 종목명
            rpTime = objRq.GetDataValue(2, i)  # 시간
            rpDiffFlag = objRq.GetDataValue(3, i)  # 대비부호
            rpDiff = objRq.GetDataValue(4, i)  # 대비
            rpCur = objRq.GetDataValue(5, i)  # 현재가
            rpVol = objRq.GetDataValue(6, i)  # 거래량
            print("자아아",rpCode, rpName, rpTime, rpDiffFlag, rpDiff, rpCur, rpVol)

        return True


class silsigan:

    def __init__(self):
        super().__init__()
        self.isSB = False
        self.objCur = []


    def StopSubscribe(self):
        if self.isSB:
            cnt = len(self.objCur)
            for i in range(cnt):
                self.objCur[i].Unsubscribe()
            print(cnt, "종목 실시간 해지되었음")
        self.isSB = False

        self.objCur = []

    def btnStart_clicked(self):
        self.StopSubscribe();

        # 요청 종목 배열

        data = pd.read_csv('E:/big12/python-project/note/categories/제약기업선정.csv', encoding='utf-8')
        codes = data['code'].tolist()

        cnt = len(codes)
        for i in range(cnt):
            self.objCur.append(CpStockCur())
            self.objCur[i].Subscribe(codes[i])

        print("빼기빼기================-")
        print(cnt, "종목 실시간 현재가 요청 시작")
        self.isSB = True