import win32com.client
import pythoncom
import sys
import Login

class XASessionEvents:
    logInState = 0

    def OnLogin(self, code, msg):
        print("onLogin method is called")
        print(str(code))
        print(str(msg))
        # 0000이 입력될 때만 로그인 성공
        if str(code) == '0000':
            XASessionEvents.logInState = 1

    def OnLogout(self):
        print("OnLogout method is called")

    def OnDisconnect(self):
        print("OnDisconnect method is called")


# 데이터 조회시 사용하는 클래스
class XAQueryEvents:
    queryState = 0

    def OnReceiveData(self, szTrCode):
        print("ReceiveData %s" % szTrCode)
        XAQueryEvents.queryState = 1

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("ReceiveMessage")

if __name__ == "__main__":

    login = Login.Login()

    login.serverConnect()

    # -- Get data --
    inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)

    # 1.stock data
    inXAQuery.LoadFromResFile("C:/eBEST/xingAPI/Res/t1102.res")                          #res 등록 (주식 현재가)
    inXAQuery.SetFieldData("t1102InBlock", "shcode", 0, "000150")        #종목코드 입력
    inXAQuery.Request(0)

    while XAQueryEvents.queryState == 0:
        pythoncom.PumpWaitingMessages()
    # Get FieldData
    name = inXAQuery.GetFieldData("t1102OutBlock", "hname", 0)
    price = inXAQuery.GetFieldData("t1102OutBlock", "price", 0)

    print("name : %s, price : %s" % (name, price))
    XAQueryEvents.queryState = 0

