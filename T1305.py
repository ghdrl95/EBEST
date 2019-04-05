import win32com.client
import pythoncom
import time
import threading
import pandas as pd
import numpy as np

# 기간별 주가 (일,주,월 시가,종가, 등락율, 거래량 데이터 추출)
class Xing_T1305(threading.Thread):
    queryState = 0

    def __init__(self, code):
        threading.Thread.__init__(self)

        self.code = code
        '''
        self.data_open = []
        self.data_close = []
        self.data_volume = []
        self.data_fpvolume = []
        self.data_covolume = []
        self.data_ppvolume = []
        self.data_diff = []
        '''
        self.datas = [];

    def getdatas(self):
        self.start()
        while self.queryState != 2:
            pythoncom.PumpWaitingMessages()
        return self.datas

    def run(self):
        pythoncom.CoInitialize()
        inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)

        # 1.stock data
        inXAQuery.LoadFromResFile("C:/eBEST/xingAPI/Res/t1305.res")  # res 등록 (주식 현재가)
        inXAQuery.SetFieldData('t1305InBlock', 'shcode', 0, self.code)
        for dwmcode in range(1, 2):

            if dwmcode == 1:
                cnt = '2500'
            elif dwmcode == 2:
                cnt = '1100'
            else:
                cnt = '300'
            # 1 : 일 2 : 주 3 : 월
            inXAQuery.SetFieldData('t1305InBlock', 'dwmcode', 0, str(dwmcode))

            inXAQuery.SetFieldData('t1305InBlock', 'cnt', 0, cnt)
            Xing_T1305.queryState = 0

            result = inXAQuery.Request(0)
            while result < 0:
                print("T1305 Error %s " % result)
                time.sleep(10)
                result = inXAQuery.Request(0)

            while Xing_T1305.queryState == 0:
                pythoncom.PumpWaitingMessages()
            n = inXAQuery.GetBlockCount("t1305OutBlock1")
            datas=[]
            #이평선 만들기
            closes=[]
            n_range = list(range(n))
            n_range.reverse()
            for i in n_range:
                data = []
                #날짜, 시가기준등락율, 고가기준등락율, 저가기준등락율, 종가기준등락율, 거래량, 거래증가율, 체결강도, 소진율, 회전율, 외인순매수, 기관순매수,개인순매수,시가총액,종가,시가,고가,저가
                #날짜
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "date", i))
                #시가기준등락율
                data.append(float(inXAQuery.GetFieldData("t1305OutBlock1", "o_diff", i)))
                #고가기준등락율
                data.append(float(inXAQuery.GetFieldData("t1305OutBlock1", "h_diff", i)))
                #저가기준등락율
                data.append(float(inXAQuery.GetFieldData("t1305OutBlock1", "l_diff", i)))
                #종가기준등락율
                data.append(float(inXAQuery.GetFieldData("t1305OutBlock1", "diff", i)))
                #거래량
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "volume", i))
                #거래증가율
                data.append(float(inXAQuery.GetFieldData("t1305OutBlock1", "diff_vol", i)))
                #체결강도
                data.append(float(inXAQuery.GetFieldData("t1305OutBlock1", "chdegree", i)))
                #소진율
                data.append(float(inXAQuery.GetFieldData("t1305OutBlock1", "sojinrate", i)))
                #회전율
                data.append(float(inXAQuery.GetFieldData("t1305OutBlock1", "changerate", i)))
                fp = int(inXAQuery.GetFieldData("t1305OutBlock1", "fpvolume", i))
                co = int(inXAQuery.GetFieldData("t1305OutBlock1", "covolume", i))
                pp = int(inXAQuery.GetFieldData("t1305OutBlock1", "ppvolume", i))
                #외인순매수
                data.append( 1 if fp>=0 else -1)
                #기관순매수
                data.append(1 if co>=0 else -1)
                #개인순매수
                data.append(1 if pp>=0 else -1)
                #시가총액
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "marketcap", i))
                #종가
                close = int(inXAQuery.GetFieldData("t1305OutBlock1", "close", i))
                closes.append(close)
                #시가 - 게임 끝내려고할때 시가로 팔기
                open = int(inXAQuery.GetFieldData("t1305OutBlock1", "open", i))
                #저가
                low = int(inXAQuery.GetFieldData("t1305OutBlock1", "low", i))
                #고가
                high = int(inXAQuery.GetFieldData("t1305OutBlock1", "high", i))
                #호가 확인 one-hot
                if close < 1000:
                    type = 0
                elif close < 5000:
                    type = 1
                elif close < 10000:
                    type= 2
                elif close < 50000:
                    type=3
                else:
                    type=4
                arr = np.zeros(5,int)
                arr[type] = 1
                data.extend(arr)
                data.append(close)
                data.append(open)
                data.append(low)
                data.append(high)
                #이평선 계산
                if len(closes) >= 120:
                    data.append(sum(closes[-120:]) / 120)
                    data.append(sum(closes[-60:]) / 60)
                    data.append(sum(closes[-20:]) / 20)
                    data.append(sum(closes[-10:]) / 10)
                    data.append(sum(closes[-5:]) / 5)
                elif len(closes) >= 60:
                    data.append(0)
                    data.append(sum(closes[-60:]) / 60)
                    data.append(sum(closes[-20:]) / 20)
                    data.append(sum(closes[-10:]) / 10)
                    data.append(sum(closes[-5:]) / 5)
                elif len(closes) >= 20:
                    data.extend([0,0])
                    data.append(sum(closes[-20:]) / 20)
                    data.append(sum(closes[-10:]) / 10)
                    data.append(sum(closes[-5:]) / 5)
                elif len(closes) >= 10:
                    data.extend([0, 0,0])
                    data.append(sum(closes[-10:]) / 10)
                    data.append(sum(closes[-5:]) / 5)
                elif len(closes)>= 5:
                    data.extend([0, 0,0,0])
                    data.append(sum(closes[-5:]) / 5)
                else:
                    data.extend([0, 0,0,0,0])
                datas.append(data)
                '''
                self.data_open.append(inXAQuery.GetFieldData("t1305OutBlock1","open",i))
                self.data_close.append(inXAQuery.GetFieldData("t1305OutBlock1", "close", i))
                self.data_volume.append(inXAQuery.GetFieldData("t1305OutBlock1", "volume", i))
                self.data_diff.append(inXAQuery.GetFieldData("t1305OutBlock1", "diff", i))
                self.data_fpvolume.append(inXAQuery.GetFieldData("t1305OutBlock1", "fpvolume", i))
                self.data_covolume.append(inXAQuery.GetFieldData("t1305OutBlock1", "covolume", i))
                self.data_ppvolume.append(inXAQuery.GetFieldData("t1305OutBlock1", "ppvolume", i))
                    '''
            # 데이터 저장
            datas.reverse()
            dataframe = pd.DataFrame(datas)
            dataframe.to_csv("./datas/%s_%d.csv" % (self.code, dwmcode), header=False, index=False)

            time.sleep(1)
        '''
        inXAQuery.SetFieldData('t1305InBlock','shcode',0, self.code)
        # 1 : 일 2 : 주 3 : 월
        inXAQuery.SetFieldData('t1305InBlock', 'dwmcode', 0, '2')
        inXAQuery.SetFieldData('t1305InBlock', 'cnt', 0, '1100')
        Xing_T1305.queryState = 0
        result = inXAQuery.Request(0)
        if result >= 0:
            while Xing_T1305.queryState == 0:
                pythoncom.PumpWaitingMessages()
            n = inXAQuery.GetBlockCount("t1305OutBlock1")
            for i in range(n-1):
                data = []
                data.append(inXAQuery.GetFieldData("t1305OutBlock1","open",i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "close", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "volume", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "diff", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "fpvolume", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "covolume", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "ppvolume", i))
                self.datas.append(data)
        else:
            print("T1833 Error %s " % result)
        time.sleep(1)

        inXAQuery.SetFieldData('t1305InBlock', 'shcode', 0, self.code)
        # 1 : 일 2 : 주 3 : 월
        inXAQuery.SetFieldData('t1305InBlock', 'dwmcode', 0, '3')
        inXAQuery.SetFieldData('t1305InBlock', 'cnt', 0, '180')
        Xing_T1305.queryState = 0
        result = inXAQuery.Request(0)
        if result >= 0:
            while Xing_T1305.queryState == 0:
                pythoncom.PumpWaitingMessages()
            n = inXAQuery.GetBlockCount("t1305OutBlock1")
            for i in range(n - 1):
                data = []
                data.append(inXAQuery.GetFieldData("t1305OutBlock1","open",i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "close", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "volume", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "diff", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "fpvolume", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "covolume", i))
                data.append(inXAQuery.GetFieldData("t1305OutBlock1", "ppvolume", i))
                self.datas.append(data)
        else:
            print("T1833 Error %s " % result)
        '''
        # GetBlockCount("블록이름")
        '''
        print("시가 %s" % self.data_open)
        print("종가 %s" % self.data_volume)
        print("거래량 %s" % self.data_open)
        print("등락률 %s" % self.data_diff)
        print("외인거래량 %s" % self.data_fpvolume)
        print("기관거래량 %s" % self.data_covolume)
        print("개인거래량 %s" % self.data_ppvolume)
        '''
        self.queryState = 2


class XAQueryEvents:
    def OnReceiveData(self, szTrCode):
        print("ReceiveData %s" % szTrCode)
        Xing_T1305.queryState = 1

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("ReceiveMessage %s %s" % (messageCode, message))
