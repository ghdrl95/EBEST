import win32com.client
import pythoncom
import time
import threading
import Login
#계좌의 주식 잔고확인
class Xing_T0424(threading.Thread):
    queryState = 0
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        pythoncom.CoInitialize()
        inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery",XAQueryEvents)

        # 1.stock data
        inXAQuery.LoadFromResFile("C:/eBEST/xingAPI/Res/t0424.res")  # res 등록 (주식 현재가)
        inXAQuery.SetFieldData('t0424InBlock','accno',0, Login.Login.account)
        inXAQuery.SetFieldData('t0424InBlock', 'passwd', 0, '0000')
        inXAQuery.SetFieldData('t0424InBlock', 'prcgb', 0, '1')
        inXAQuery.SetFieldData('t0424InBlock', 'chegb', 0, '2')
        inXAQuery.SetFieldData('t0424InBlock', 'dangb', 0, '0')
        inXAQuery.SetFieldData('t0424InBlock', 'charge', 0, '0')
        result = inXAQuery.Request(0)
        if result >= 0:
            while Xing_T0424.queryState == 0:
                pythoncom.PumpWaitingMessages()
            n = inXAQuery.GetBlockCount("t1833OutBlock1")
            for i in range(n-1):
                str = inXAQuery.GetFieldData("t1833OutBlock1","shcode",i) +", " +inXAQuery.GetFieldData("t1833OutBlock1","hname",i) +", " +inXAQuery.GetFieldData("t1833OutBlock1","sign",i) +", " +inXAQuery.GetFieldData("t1833OutBlock1","signcnt",i)+", " +inXAQuery.GetFieldData("t1833OutBlock1","close",i)+", " +inXAQuery.GetFieldData("t1833OutBlock1","change",i)+", " +inXAQuery.GetFieldData("t1833OutBlock1","diff",i)+", " +inXAQuery.GetFieldData("t1833OutBlock1","volume",i)
                print (str)

        else:
            print("T1833 Error %s " % result)
        # GetBlockCount("블록이름")

class XAQueryEvents:

    def OnReceiveData(self, szTrCode):
        print("ReceiveData %s" % szTrCode)
        Xing_T0424.queryState = 1
    def OnReceiveMessage(self, systemError, messageCode, message):
        print("ReceiveMessage %s %s" % (messageCode,message))