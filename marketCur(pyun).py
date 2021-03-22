import sys
from PyQt5.QtWidgets import *
import win32com.client

##현재가 조회
class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()

# 복수 종목 실시간 조회 샘플 (조회는 없고 실시간만 있음)
class CpEvent:
    def set_params(self, client):
        self.client = client

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        diff = self.client.GetHeaderValue(2)  # 대비
        time = self.client.GetHeaderValue(3)  # 시간
        cur_price = self.client.GetHeaderValue(4)  # 시가
        high_price = self.client.GetHeaderValue(5)  # 고가
        low_price = self.client.GetHeaderValue(6)  # 저가
        sell_call = self.client.GetHeaderValue(7)  # 매도호가
        buy_call = self.client.GetHeaderValue(8)  # 매수호가.
        acc_vol = self.client.GetHeaderValue(9)  # 누적거래량
        cprice = self.client.GetHeaderValue(13)  # 현재가
        deal_state = self.client.GetHeaderValue(14)  # 체결
        acc_sell_deal_vol = self.client.GetHeaderValue(15)  # 현재가
        acc_buy_deal_vol = self.client.GetHeaderValue(16)  # 현재가
        moment_deal_vol = self.client.GetHeaderValue(17)  # 순간체결수량
        time_sec = self.client.GetHeaderValue(18)  # 초
        exp_price_com_flag= self.client.GetHeaderValue(19)  # 예상체결가 구분 플래그
        market_diff_flag= self.client.GetHeaderValue(20)  # 장구분 플래그



        print(code,name,diff,time,cur_price,high_price,low_price,sell_call,buy_call,acc_vol,cprice,deal_state,acc_sell_deal_vol,acc_buy_deal_vol,moment_deal_vol,time_sec, exp_price_com_flag,market_diff_flag)



isSB = False
objCur = []


def StopSubscribe(self):
    if self.isSB:
        cnt = len(self.objCur)
        for i in range(cnt):
            self.objCur[i].Unsubscribe()
        print(cnt, "종목 실시간 해지되었음")
    self.isSB = False

    self.objCur = []


    self.StopSubscribe();

# 요청 종목 배열
codes = ["A005930","A003540", "A000660", "A005930", "A035420", "A069500", "Q530031"]

cnt = len(codes)
for i in range(cnt):
   objCur.append(CpStockCur())
   objCur[i].Subscribe(codes[i])

print("빼기빼기================-")
print(cnt, "종목 실시간 현재가 요청 시작")
isSB = True
