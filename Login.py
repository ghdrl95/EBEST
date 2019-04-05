import win32com.client
import pythoncom
import time
import T1833
import T1305
import T8436
import datetime

'''
class XASessionEventHandler:
    login_state = 0

    def OnLogin(self,code,msg):
        if code=="0000":
            print("로그인 성공")
            XASessionEventHandler.login_state=1
        else:
            print("로그인 실패")


instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)
id = "shh821"
passwd = "algo9252"
cert_passwd = "cert"

instXASession.ConnectServer("demo.ebestsec.co.kr", 20001)
instXASession.Login(id, passwd, cert_passwd, 0, 0)

while XASessionEventHandler.login_state == 0 :
    pythoncom.PumpWaitingMessages()
'''
import sys


class XASessionEvents:
    logInState = 0

    def OnLogin(self, code, msg):
        print("onLogin method is called")
        print(str(code))
        print(str(msg))

        # 0000이 입력될 때만 로그인 성공
        if str(code) == '0000':
            XASessionEvents.logInState = 1
        else:
            XASessionEvents.logInState = -1

    def OnLogout(self):
        print("OnLogout method is called")

    def OnDisconnect(self):
        print("OnDisconnect method is called")


class Login:
    server_addr = 'demo.ebestsec.co.kr'
    server_port = 200001
    server_type = 0
    user_id = 'shg955'
    user_pass = 'shh13579'
    user_cert = '공인인증 비밀번호'
    account = ''

    def serverConnect(self):
        inXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvents)
        bConnect = None
        while not bConnect:
            bConnect = inXASession.ConnectServer(Login.server_addr, Login.server_port)

            if not bConnect:
                nErrCode = inXASession.GetLastError()
                strErrMsg = inXASession.GetErrorMessage(nErrCode)
                print(strErrMsg)
                time.sleep(600)

        XASessionEvents.logInState = -1
        while XASessionEvents.logInState == -1:
            inXASession.Login(Login.user_id, Login.user_pass, Login.user_cert, Login.server_type, 0)
            XASessionEvents.logInState = 0
            while XASessionEvents.logInState == 0:
                pythoncom.PumpWaitingMessages()
            if XASessionEvents.logInState == -1:
                print("로그인 실패 재접속 중")
                time.sleep(600)
        # 계좌정보 불러오기
        nCount = inXASession.GetAccountListCount()
        for i in range(nCount):
            print("Account : %d - %s" % (i, inXASession.GetAccountList(i)))
        account = inXASession.GetAccountList(0)


# 개장시간 기다리기
open = datetime.time(9, 0)
close = datetime.time(15, 30)


def waitOpenTime():
    today = datetime.datetime.now()
    nowtime = datetime.time(today.hour, today.minute)
    while not (nowtime > open and nowtime < close):
        nowtime = datetime.datetime.now()
        time.sleep(600)


# 현재 장 시간대 인지 확인
def checkOpen():
    today = datetime.datetime.now()
    nowtime = datetime.time(today.hour, today.minute)
    return (nowtime > open and nowtime < close)


import pandas as pd

if __name__ == "__main__":
    obj = Login()
    obj.serverConnect()
    # 코스피 종목 가져오기
    t8436 = T8436.Xing_T8436()
    t8436.getdatas()
    # codelist.csv에 있는 종목들 데이터 수집하기
    csv_data = pd.read_csv("./datas/codelist.csv")
    #print(csv_data.loc[:,'code'])
    cnt = 0
    for data in csv_data.loc[:,'code']:
        print (data)
        t1305 = T1305.Xing_T1305(data)
        t1305.getdatas()
        cnt += 1
        if cnt / 150 > 1:
            print('요청초과 대기중')
            time.sleep(600)
            cnt=0
    '''
    while True:
        #개장 시간 대기
        #waitOpenTime()


        
        while True:
            #if not checkOpen():
            #    break
            
            t1833 = T1833.Xing_T1833()
            t1833.getList()
            if not t1833.list:
                print("조건충족 주식 없음")
            else:
                print("조건충족 주식 있음")
                for obj in t1833.list:
                    t1305 = T1305.Xing_T1305(obj.code)
                    datas = t1305.getdatas()

            time.sleep(120)
        '''
