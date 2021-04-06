import sys
from PyQt5.QtWidgets import *
import win32com.client
import sqlite3
import time
import pandas as pd

class CpEvent:
    instance = None

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 코드
        name = self.client.GetHeaderValue(1)  # 종목명
        diff = self.client.GetHeaderValue(2)  # 전일대비
        cur_price = self.client.GetHeaderValue(4)  # 시가
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
        date_time_sec = timess1 + str(self.client.GetHeaderValue(18))  # 시간(초)
        exFlag = self.client.GetHeaderValue(19)  # 예상체결가구분플래그
        market_diff_flag = self.client.GetHeaderValue(20)  # 장구분플래그

        # if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
        #     print("실시간(예상체결)", timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        # elif (exFlag == ord('2')):  # 장중(체결)
        #     print("실시간(장중 체결)", timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)


class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        CpEvent.instance = self.objStockCur
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()


class CpStockMst:
    def Request(self, code):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        # 현재가 객체 구하기
        objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        objStockMst.SetInputValue(0, code)  # 종목 코드 - 삼성전자
        objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objStockMst.GetDibStatus()
        rqRet = objStockMst.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        # 현재가 정보 조회
        code = objStockMst.GetHeaderValue(0)  # 종목코드
        name = objStockMst.GetHeaderValue(1)  # 종목명
        time = objStockMst.GetHeaderValue(4)  # 시간
        cprice = objStockMst.GetHeaderValue(11)  # 종가
        diff = objStockMst.GetHeaderValue(12)  # 대비
        open = objStockMst.GetHeaderValue(13)  # 시가
        high = objStockMst.GetHeaderValue(14)  # 고가
        low = objStockMst.GetHeaderValue(15)  # 저가
        offer = objStockMst.GetHeaderValue(16)  # 매도호가
        bid = objStockMst.GetHeaderValue(17)  # 매수호가
        vol = objStockMst.GetHeaderValue(18)  # 거래량
        vol_value = objStockMst.GetHeaderValue(19)  # 거래대금

        print("코드 이름 시간 현재가 대비 시가 고가 저가 매도호가 매수호가 거래량 거래대금")
        print(code, name, time, cprice, diff, open, high, low, offer, bid, vol, vol_value)
        return True


class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 150)
        self.isRq = False
        self.objStockMst = CpStockMst()
        self.objStockCur = CpStockCur()

        btn1 = QPushButton("요청 시작", self)
        btn1.move(20, 20)
        btn1.clicked.connect(self.btn1_clicked)

        btn2 = QPushButton("요청 종료", self)
        btn2.move(20, 70)
        btn2.clicked.connect(self.btn2_clicked)

        btn3 = QPushButton("종료", self)
        btn3.move(20, 120)
        btn3.clicked.connect(self.btn3_clicked)

    def StopSubscribe(self):
        if self.isRq:
            self.objStockCur.Unsubscribe()
        self.isRq = False

    def btn1_clicked(self):
        data = pd.read_csv('E:/big12/python-project/note/categories/제약기업선정.csv', encoding='utf-8')
        codes = data['code'].tolist()
        print(codes)
        for code in codes:
            if (self.objStockMst.Request(code) == False):
                exit()

        # 하이닉스 실시간 현재가 요청
        for code in codes:
            self.objStockCur.Subscribe(code)

            print("빼기빼기================-")
            print("실시간 현재가 요청 시작")
            self.isRq = True

    def btn2_clicked(self):
        self.StopSubscribe()

    def btn3_clicked(self):
        self.StopSubscribe()
        exit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
app.exec_()
