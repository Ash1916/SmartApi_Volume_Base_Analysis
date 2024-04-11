from SmartApi import SmartConnect
import pyotp
from logzero import logger
from datetime import date, timedelta , datetime
from SmartApi.smartApiWebsocket import SmartWebSocket
import websocket
import pandas as pd
import requests
import numpy as np
from matplotlib import pyplot as plt
import openpyxl as xl

A = (datetime.now() - timedelta(days=100))
B = str(A).split(' ')[0]
E = str(B) + ' ' + str('09:15')

C = datetime.now()
D = str(C).split(' ')[0] + ' ' + str('03:30')


Exchange = 'NSE'
Interval = 'ONE_DAY'
Fromdate = E
Todate = D


api_key = 'N6YxYEPv'
username = 'Aakash Verma'
pwd = '9818'
Client_ID = 'A51427597'
smartApi = SmartConnect(api_key)



try:
    token = 'R6JIVSWA2PEKCAUYRUDCMTWF64'
    totp = pyotp.TOTP(token).now()
except Exception as e:
    logger.error("Invalid Token: The provided token is not valid.")
    raise e


correlation_id = "abcde"
data = smartApi.generateSession(Client_ID, pwd, totp)
Token = []


if not data['status']:
    logger.error(data)

else:
    def get_historical_data(exchange, symboltoken, interval, fromdate, todate):
        historicParam = {
            "exchange": exchange,
            "symboltoken": symboltoken,
            "interval": interval,
            "fromdate": fromdate,
            "todate": todate
        }
        try:
            historical_data = smartApi.getCandleData(historicParam)
            return historical_data
            
        except Exception as e:
            logger.exception(f"Historic Api failed: {e}")
            
            
            
    def TokenInfo(Name, exch_seg):
        #data = requests.get('https://margincalculator.angelbroking.com/OpenAPI_File/files/OpenAPIScripMaster.json').json()
        #List = pd.DataFrame.from_dict(data)
        List = pd.read_csv("C:/Users/akaeh/Downloads/token_data.csv")
        List['expiry'] = pd.to_datetime(List['expiry'], format='mixed')
        List = List.astype({'strike': float})
        L = (f"{Name}-EQ")
        M = List[List['symbol'] == L]
        token = M.iloc[0]
        you = token["token"]
        return you
        '''
        M = List['Name'+'-EQ'] = Name
        v = List['exch_seg'] = exch_seg
        Token = List.iloc[0]
        You = Token['token']
        return You
        '''
    
         
    def Ap(list):
        Data = []
        n = 1
        H = len(list) - 1
        for i in range(1,H):
            if list[i-1] - list[i] < 0:
                if list[i] - list[i-1] > list[i-1] * 2:
                    Ap = list[i-1] - list[i]
                    Data.append(Ap)
            
            elif list[i-1] - list[i] > list[i] * 2:
                Ap = list[i-1] - list[i]
                Data.append(Ap)
                
            else:
                pass

                
        return Data
    
    
    
    
    def Main(Stock, Exchange, Interval, Fromdate, Todate):
        try:
            I = TokenInfo(Stock, Exchange)
            
        except:
            IndexError
            h = "BL Value"
            return h
        X = get_historical_data(Exchange, I,  Interval, Fromdate, Todate)
        H = pd.DataFrame(X)
        Y = H['data']
        list = []
        for volume in Y:
            U = volume[5]
            list.append(U)
        
        k = Ap(list)
            
        return len(k)
        


    
    Ly = xl.load_workbook("C:/Users/akaeh/Downloads/Doji, Technical Analysis Scanner (7).xlsx")
    Data_1 = Ly.active
    n = 165
    
    while True:
        # Potencial Stock Name from colum B
        H = Data_1[f"C{n}"].value
        Number = Main(H, Exchange, Interval, Fromdate, Todate )
        if Number == "BL Value":
            break
        else:
            try:
                Data_1[f"M{n}"].value = Number
            except:
                ValueError
                
            print(H,Number)
            
            n += 1
            Ly.save("C:/Users/akaeh/Downloads/Doji, Technical Analysis Scanner (7).xlsx")
            
