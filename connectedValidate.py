import win32com.client

cpcybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = cpcybos.IsConnect
if(bConnect ==0):
    print("연결 안됨")
    exit()