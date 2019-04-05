'''
주식종목 조회 클래스 파일
2019.01.12 Hong Gi Shin.
'''
import win32com.client
import pythoncom
import time
import threading
import pandas as pd


class Xing_T8436(threading.Thread):
    queryState = 0

    def __init__(self):
        threading.Thread.__init__(self)

    def getdatas(self):
        self.start()
        while self.queryState != 2:
            pythoncom.PumpWaitingMessages()
        return True

    def run(self):
        pythoncom.CoInitialize()
        inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)

        # 1.stock data
        inXAQuery.LoadFromResFile("C:/eBEST/xingAPI/Res/t8436.res")  # res 등록 (주식 현재가)
        inXAQuery.SetFieldData('t8436InBlock', 'gubun', 0, '1')
        Xing_T8436.queryState = 0

        result = inXAQuery.Request(0)
        while result < 0:
            print("T8436 Error %s " % result)
            time.sleep(10)
            result = inXAQuery.Request(0)

        while Xing_T8436.queryState == 0:
            pythoncom.PumpWaitingMessages()
        n = inXAQuery.GetBlockCount("t8436OutBlock")
        datas = []
        for i in range(n):
            data = []
            # 최고,최저로 5000~10000에 들어오는 종목 코드만 추출

            high = int(inXAQuery.GetFieldData("t8436OutBlock", "uplmtprice", i))
            low = int(inXAQuery.GetFieldData("t8436OutBlock", "dnlmtprice", i))

            if low >= 10000 and high <= 50000:
                data.append(high)
                data.append(low)
                data.append(inXAQuery.GetFieldData("t8436OutBlock", "shcode", i))
                data.append(inXAQuery.GetFieldData("t8436OutBlock", "jnilclose", i))
                datas.append(data)

        # 데이터 저장
        dataframe = pd.DataFrame(datas)
        dataframe.to_csv("./datas/codelist.csv", header=['high', 'low', 'code', 'price'], index=False)

        self.queryState = 2


class XAQueryEvents:
    def OnReceiveData(self, szTrCode):
        print("ReceiveData %s" % szTrCode)
        Xing_T8436.queryState = 1

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("ReceiveMessage %s %s" % (messageCode, message))
