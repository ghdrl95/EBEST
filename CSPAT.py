import win32com.client
import pythoncom
import time
import threading
import Login

#주식 주문넣기
class Xing_CSPAT00600(threading.Thread):

    def __init__(self,code,num,price,type,ordercode='00'):
        threading.Thread.__init__(self)
        self.code =code
        self.num = num
        self.price = price
        self.type = type
        self.ordercode = ordercode
    def run(self):
        pythoncom.CoInitialize()
        inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
        inXAQuery.LoadFromResFile("C:/eBEST/xingAPI/Res/CSPAT00600.res")  # res 등록 (주식 현재가)
        inXAQuery.SetFieldData('CSPAT00600InBlock1', 'AccntNo', 0, Login.Login.account)
        inXAQuery.SetFieldData('CSPAT00600InBlock1', 'InptPwd', 0, '0000')
        inXAQuery.SetFieldData('CSPAT00600InBlock1', 'IsuNo', 0, self.code)
        inXAQuery.SetFieldData('CSPAT00600InBlock1', 'OrdQty', 0, self.num)
        inXAQuery.SetFieldData('CSPAT00600InBlock1', 'OrdPrc', 0, self.price)
        inXAQuery.SetFieldData('CSPAT00600InBlock1', 'BnsTpCode', 0, self.type)
        inXAQuery.SetFieldData('CSPAT00600InBlock1', 'OrdprcPtnCode', 0, self.ordercode )
        inXAQuery.SetFieldData('CSPAT00600InBlock1', 'MgntrnCode', 0, '000')
        inXAQuery.SetFieldData('CSPAT00600InBlock1', 'OrdCndiTpCode', 0, '0')
        result = inXAQuery.Request(0)
        if result >= 0:
            print("Xing_CSPAT00600 %s" % '주문완료')
        else:
            print("Xing_CSPAT00600 %s " % result)
class XAQueryEvents:

    def OnReceiveData(self, szTrCode):
        print("ReceiveData %s" % szTrCode)
    def OnReceiveMessage(self, systemError, messageCode, message):
        print("ReceiveMessage %s %s" % (messageCode,message))