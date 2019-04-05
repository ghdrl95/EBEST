import win32com.client
import pythoncom
import time
import threading

#조건식으로 종목검색
class Xing_T1833(threading.Thread):
    queryState = 0
    def __init__(self):
        threading.Thread.__init__(self)
        self.list = []
    def getList(self):
        self.start()
        while not self.queryState == 2:
            pythoncom.PumpWaitingMessages()

    def run(self):
        pythoncom.CoInitialize()
        inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery",XAQueryEvents)

        # 1.stock data
        inXAQuery.LoadFromResFile("C:/eBEST/xingAPI/Res/t1833.res")  # res 등록 (주식 현재가)
        result = inXAQuery.RequestService("t1833","C:/eBEST/xingAPI/Res/buy.ACF")
        if result >= 0:
            while Xing_T1833.queryState == 0:
                pythoncom.PumpWaitingMessages()
            n = inXAQuery.GetBlockCount("t1833OutBlock1")
            self.list=[]
            for i in range(n-1):
                diction = {'code' : inXAQuery.GetFieldData("t1833OutBlock1","shcode",i), 'name': inXAQuery.GetFieldData("t1833OutBlock1","hname",i), 'value': inXAQuery.GetFieldData("t1833OutBlock1","close",i)}
                #str = inXAQuery.GetFieldData("t1833OutBlock1","shcode",i) +", " +inXAQuery.GetFieldData("t1833OutBlock1","hname",i) +", " +inXAQuery.GetFieldData("t1833OutBlock1","sign",i) +", " +inXAQuery.GetFieldData("t1833OutBlock1","signcnt",i)+", " +inXAQuery.GetFieldData("t1833OutBlock1","close",i)+", " +inXAQuery.GetFieldData("t1833OutBlock1","change",i)+", " +inXAQuery.GetFieldData("t1833OutBlock1","diff",i)+", " +inXAQuery.GetFieldData("t1833OutBlock1","volume",i)
                print (diction)
                self.list.append(diction)

        else:
            print("T1833 Error %s " % result)
        # GetBlockCount("블록이름")
        self.queryState = 2
class XAQueryEvents:

    def OnReceiveData(self, szTrCode):
        print("ReceiveData %s" % szTrCode)
        Xing_T1833.queryState = 1
    def OnReceiveMessage(self, systemError, messageCode, message):
        print("ReceiveMessage %s %s" % (messageCode,message))