import re
import os
import pandas as pd
from datetime import datetime
from datetime import timedelta
import csv
import openpyxl
def tradeDay(date,ticker):
    os.chdir("D:\\high-frequency data\\pharmarcy\\"+ticker)
    arr = os.listdir()
    stockfolder = []
    x=0
    for y in range(len(arr)):
        if date+'.csv' ==arr[y]:
            x=1
    if x==1:
        return True
    else:
        return False
def timespot(date_time_obj):
    hour = date_time_obj.strftime("%H")
    min = date_time_obj.strftime("%M")
    time = int(hour+min)
    if time<700:
        return "pre_trade"
    if time<930:
        if time>=700:
            return "pre_market"
    if time>=930:
        if time<1600:
            return "market"
    if time>=1600:
        if time<=2000:
            return "post_market"
    if time>2000:
        return "post_trade"

def notTrade(date_time_obj,ticker):
    return 1

def one_min(data,date_time_obj):
    one_min_later = date_time_obj + timedelta(minutes=1)
    hour = one_min_later.strftime("%H")
    min = one_min_later.strftime("%M")
    want = data.loc[data['time'] == hour + min]
    if want.empty:
        print("N/A")
    else:
        return want.reset_index(drop=True).loc[0, 'open']

def one_hour(data,date_time_obj):
    one_min_later = date_time_obj + timedelta(minutes=60)
    hour = one_min_later.strftime("%H")
    min = one_min_later.strftime("%M")
    want = data.loc[data['time'] == hour + min]
    if want.empty:
        return "N/A"
    else:
        return want.reset_index(drop=True).loc[0, 'open']

def next_day_opening(date,ticker):
    os.chdir("D:\\high-frequency data\\pharmarcy\\"+ticker)
    arr = os.listdir()
    lol = []
    x = re.search("c", arr[0])
    y = x.start()
    for x in range(len(arr)):
        zz = arr[x]
        yy = zz[:y - 1]
        lol.append(yy)
    spot = lol.index(date)
    if spot < len(arr)-1:
        date = arr[spot + 1]
        date = date[:y - 1]
        data = pd.read_csv(date + '.csv',
                           names=["date", "time", "open", "high", "low", "close", "volume", "split factor", "earnings",
                                  "dividends"])
        data = data.astype(str)
        want = data.loc[data['time'] == '930']
        if want.empty:
            return "N/A"
        else:
            return (want.reset_index(drop=True).loc[0, 'open'])
    else:
        return "N/A"

def next_day_close(date,ticker):
    os.chdir("D:\\high-frequency data\\pharmarcy\\" + ticker)
    arr = os.listdir()
    lol = []
    x = re.search("c", arr[0])
    y = x.start()
    for x in range(len(arr)):
        zz = arr[x]
        yy = zz[:y - 1]
        lol.append(yy)
    spot = lol.index(date)
    if spot < len(arr)-1:
        date = arr[spot + 1]
        date = date[:y - 1]
        data = pd.read_csv(date + '.csv',
                           names=["date", "time", "open", "high", "low", "close", "volume", "split factor", "earnings",
                                  "dividends"])
        data = data.astype(str)
        want = data.loc[data['time'] == '1600']
        if want.empty:
            return "N/A"
        else:
            return (want.reset_index(drop=True).loc[0, 'open'])
    else:
        return "N/A"

def ten_days_close(date,ticker):
    os.chdir("D:\\high-frequency data\\pharmarcy\\" + ticker)
    arr = os.listdir()
    lol = []
    x = re.search("c", arr[0])
    y = x.start()
    for x in range(len(arr)):
        zz = arr[x]
        yy = zz[:y - 1]
        lol.append(yy)
    spot = lol.index(date)
    if spot+10 < len(arr):
        date = arr[spot + 10]
        date = date[:y - 1]
        data = pd.read_csv(date + '.csv',
                           names=["date", "time", "open", "high", "low", "close", "volume", "split factor", "earnings",
                                  "dividends"])
        data = data.astype(str)
        want = data.loc[data['time'] == '1600']
        if want.empty:
            return "N/A"
        else:
            return (want.reset_index(drop=True).loc[0, 'open'])
    else:
        return "N/A"

def one_month_close(date,ticker):
    os.chdir("D:\\high-frequency data\\pharmarcy\\" + ticker)
    arr = os.listdir()
    lol = []
    x = re.search("c", arr[0])
    y = x.start()
    for x in range(len(arr)):
        zz = arr[x]
        yy = zz[:y - 1]
        lol.append(yy)
    spot = lol.index(date)
    if spot+22 < len(arr):
        date = arr[spot + 22]
        date = date[:y - 1]
        data = pd.read_csv(date + '.csv',
                           names=["date", "time", "open", "high", "low", "close", "volume", "split factor", "earnings",
                                  "dividends"])
        data = data.astype(str)
        want = data.loc[data['time'] == '1600']
        if want.empty:
            return "N/A"
        else:
            return (want.reset_index(drop=True).loc[0, 'open'])
    else:
        return "N/A"

def market(date_time_obj,ticker):
    year = "20"+date_time_obj.strftime("%y")
    month = date_time_obj.strftime("%m")
    day = date_time_obj.strftime("%d")
    hour = date_time_obj.strftime("%H")
    min = date_time_obj.strftime("%M")
    second = date_time_obj.strftime("%S")
    date = year+month+day
    os.chdir('D:\\high-frequency data\\pharmarcy\\' + ticker)
    data = pd.read_csv(date + ".csv",
                       names=["date", "time", "open", "high", "low", "close", "volume", "split factor", "earnings",
                              "dividends"])
    data = data.astype(str)
    want = data.loc[data['time'] == hour + min]
    df = pd.DataFrame(columns=['time', 'ticker', 'status','pre','open', 'volume', 'one min', 'one hour','close price','post',
                               'next day opening','next day close','ten days close','one month close'])
    df.at[0, 'time'] =date_time_obj.strftime("%m/%d/%Y, %H:%M:%S")
    df.at[0, 'ticker'] = ticker
    df.at[0, 'pre'] = 'N/A'
    if want.empty:
        df.at[0, 'open'] = "N/A"
        df.at[0, 'volume'] = "N/A"
    else:
        df.at[0, 'open'] = want.reset_index(drop=True).loc[0, 'open']
        df.at[0, 'volume'] = want.reset_index(drop=True).loc[0, 'volume']
    df.at[0,'one min'] = one_min(data,date_time_obj)
    df.at[0,'one hour'] = one_hour(data,date_time_obj)
    close_price = data.loc[data['time'] == '1600']
    if close_price.empty:
        return "N/A"
    else:
        df.at[0, 'close price'] = close_price.reset_index(drop=True).loc[0, 'close']
    df.at[0,'post']= data.loc[len(data)-1,'close']
    df.at[0,'next day opening'] = next_day_opening(date,ticker)
    df.at[0,'next day close'] = next_day_close(date,ticker)
    df.at[0,'ten days close'] = ten_days_close(date,ticker)
    df.at[0,'one month close'] = one_month_close(date,ticker)
    return df



def pre_trade(date_time_obj,ticker):
    year = "20" + date_time_obj.strftime("%y")
    month = date_time_obj.strftime("%m")
    day = date_time_obj.strftime("%d")
    hour = date_time_obj.strftime("%H")
    min = date_time_obj.strftime("%M")
    second = date_time_obj.strftime("%S")
    date = year + month + day
    os.chdir('D:\\high-frequency data\\pharmarcy\\' + ticker)
    data = pd.read_csv(date + ".csv",
                       names=["date", "time", "open", "high", "low", "close", "volume", "split factor", "earnings",
                              "dividends"])
    data = data.astype(str)
    df = pd.DataFrame(columns=['time', 'ticker','status', 'pre','open', 'volume', 'one min', 'one hour','close price','post',
                               'next day opening','next day close','ten days close','one month close'])
    df.at[0, 'time'] =date_time_obj.strftime("%m/%d/%Y, %H:%M:%S")
    df.at[0, 'ticker'] = ticker
    pre = data.loc[data['time'] == '700']
    if pre.empty:
        df.at[0, 'pre']= "N/A"
    else:
        df.at[0, 'pre'] = pre.reset_index(drop=True).loc[0, 'open']
    want = data.loc[data['time'] == '930']
    if want.empty:
        df.at[0, 'open'] = "N/A"
        df.at[0, 'volume'] = "N/A"
    else:
        df.at[0, 'open'] = want.reset_index(drop=True).loc[0, 'open']
        df.at[0, 'volume'] = want.reset_index(drop=True).loc[0, 'volume']
    df.at[0, 'one min'] = 'N/A'
    df.at[0, 'one hour'] = 'N/A'
    close_price = data.loc[data['time'] == '1600']
    if close_price.empty:
        return "N/A"
    else:
        df.at[0, 'close price'] = close_price.reset_index(drop=True).loc[0, 'close']
    df.at[0, 'post'] = data.loc[len(data) - 1, 'close']
    df.at[0, 'next day opening'] = next_day_opening(date, ticker)
    df.at[0, 'next day close'] = next_day_close(date, ticker)
    df.at[0, 'ten days close'] = ten_days_close(date, ticker)
    df.at[0, 'one month close'] = one_month_close(date, ticker)
    df.at[0, 'status'] = 'pre_trade'
    return df

def pre_market(date_time_obj,ticker):
    year = "20" + date_time_obj.strftime("%y")
    month = date_time_obj.strftime("%m")
    day = date_time_obj.strftime("%d")
    hour = date_time_obj.strftime("%H")
    min = date_time_obj.strftime("%M")
    second = date_time_obj.strftime("%S")
    date = year + month + day
    os.chdir('D:\\high-frequency data\\pharmarcy\\' + ticker)
    data = pd.read_csv(date + ".csv",
                       names=["date", "time", "open", "high", "low", "close", "volume", "split factor", "earnings",
                              "dividends"])
    data = data.astype(str)
    want = data.loc[data['time'] == hour + min]
    df = pd.DataFrame(columns=['time', 'ticker', 'status','pre', 'open', 'volume', 'one min', 'one hour', 'close price', 'post',
                               'next day opening', 'next day close', 'ten days close', 'one month close'])
    df.at[0, 'time'] = date_time_obj.strftime("%m/%d/%Y, %H:%M:%S")
    df.at[0, 'ticker'] = ticker
    pre = data.loc[data['time'] == '700']
    if pre.empty:
        df.at[0, 'pre'] = "N/A"
    else:
        df.at[0, 'pre'] = pre.reset_index(drop=True).loc[0, 'open']
    if want.empty:
        df.at[0, 'open'] = "N/A"
        df.at[0, 'volume'] = "N/A"
    else:
        df.at[0, 'open'] = want.reset_index(drop=True).loc[0, 'open']
        df.at[0, 'volume'] = want.reset_index(drop=True).loc[0, 'volume']
    df.at[0, 'one min'] = one_min(data, date_time_obj)
    df.at[0, 'one hour'] = one_hour(data, date_time_obj)
    close_price = data.loc[data['time'] == '1600']
    if close_price.empty:
        return "N/A"
    else:
        df.at[0, 'close price'] = close_price.reset_index(drop=True).loc[0, 'close']
    df.at[0, 'post'] = data.loc[len(data) - 1, 'close']
    df.at[0, 'next day opening'] = next_day_opening(date, ticker)
    df.at[0, 'next day close'] = next_day_close(date, ticker)
    df.at[0, 'ten days close'] = ten_days_close(date, ticker)
    df.at[0, 'one month close'] = one_month_close(date, ticker)
    df.at[0, 'status'] = 'pre_market'
    return df
def post_market(date_time_obj,ticker):
    year = "20" + date_time_obj.strftime("%y")
    month = date_time_obj.strftime("%m")
    day = date_time_obj.strftime("%d")
    hour = date_time_obj.strftime("%H")
    min = date_time_obj.strftime("%M")
    second = date_time_obj.strftime("%S")
    date = year + month + day
    os.chdir('D:\\high-frequency data\\pharmarcy\\' + ticker)
    data = pd.read_csv(date + ".csv",
                       names=["date", "time", "open", "high", "low", "close", "volume", "split factor", "earnings",
                              "dividends"])
    data = data.astype(str)
    want = data.loc[data['time'] == hour + min]
    df = pd.DataFrame(columns=['time', 'ticker','status', 'pre', 'open', 'volume', 'one min', 'one hour', 'close price', 'post',
                               'next day opening', 'next day close', 'ten days close', 'one month close'])
    df.at[0, 'time'] = date_time_obj.strftime("%m/%d/%Y, %H:%M:%S")
    df.at[0, 'ticker'] = ticker
    df.at[0, 'pre'] = 'N/A'
    if want.empty:
        df.at[0, 'open'] = "N/A"
        df.at[0, 'volume'] = "N/A"
    else:
        df.at[0, 'open'] = want.reset_index(drop=True).loc[0, 'open']
        df.at[0, 'volume'] = want.reset_index(drop=True).loc[0, 'volume']

    df.at[0, 'one min'] = one_min(data, date_time_obj)
    df.at[0, 'one hour'] = one_hour(data, date_time_obj)
    df.at[0, 'close price'] = 'N/A'
    if len(data)<1:
        df.at[0, 'post'] = 'N/A'
    else:
        df.at[0, 'post'] = data.loc[len(data) - 1, 'close']
    df.at[0, 'next day opening'] = next_day_opening(date, ticker)
    df.at[0, 'next day close'] = next_day_close(date, ticker)
    df.at[0, 'ten days close'] = ten_days_close(date, ticker)
    df.at[0, 'one month close'] = one_month_close(date, ticker)
    df.at[0, 'status'] = 'post_market'
    return df
def post_trade(date_time_obj,ticker):
    year = "20" + date_time_obj.strftime("%y")
    month = date_time_obj.strftime("%m")
    day = date_time_obj.strftime("%d")
    hour = date_time_obj.strftime("%H")
    min = date_time_obj.strftime("%M")
    second = date_time_obj.strftime("%S")
    date = year + month + day
    os.chdir('D:\\high-frequency data\\pharmarcy\\' + ticker)
    data = pd.read_csv(date + ".csv",
                       names=["date", "time", "open", "high", "low", "close", "volume", "split factor", "earnings",
                              "dividends"])
    data = data.astype(str)
    want = data.loc[data['time'] == hour + min]
    df = pd.DataFrame(columns=['time', 'ticker','status', 'pre', 'open', 'volume', 'one min', 'one hour', 'close price', 'post',
                               'next day opening', 'next day close', 'ten days close', 'one month close'])
    df.at[0, 'time'] = date_time_obj.strftime("%m/%d/%Y, %H:%M:%S")
    df.at[0, 'ticker'] = ticker
    df.at[0, 'pre'] = 'N/A'
    df.at[0, 'open'] = 'N/A'
    df.at[0, 'volume'] = 'N/A'
    df.at[0, 'one min'] = 'N/A'
    df.at[0, 'one hour'] = 'N/A'
    df.at[0, 'close price'] = 'N/A'
    if len(data)<1:
        df.at[0, 'post'] = 'N/A'
    else:
        df.at[0, 'post'] = data.loc[len(data) - 1, 'close']
    df.at[0, 'next day opening'] = next_day_opening(date, ticker)
    df.at[0, 'next day close'] = next_day_close(date, ticker)
    df.at[0, 'ten days close'] = ten_days_close(date, ticker)
    df.at[0, 'one month close'] = one_month_close(date, ticker)
    df.at[0,'status']='post_trade'
    return df