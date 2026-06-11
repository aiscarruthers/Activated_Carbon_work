
import pyodbc
import pandas as pd
from datetime import datetime
from datetime import timedelta
import matplotlib.pyplot as plt
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from scipy.signal import savgol_filter
import numpy as np
from numpy import polyfit
import csv


def fetch_tag(tag, duration, end, resolution=5):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with 5s intervals as default for the period asked for.
    """
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.47;PORT=10014")
    start = end - timedelta(hours = duration)
    end = end.strftime("%Y-%m-%d %H:%M:%S")
    start=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period is in tenths of seconds 
    sql = "select \"IP_TREND_TIME\" as \"TS\", \"IP_TREND_VALUE\" as \"VALUE\" from \"%s\" "\
            "where \"IP_TREND_TIME\" between TIMESTAMP'%s' and TIMESTAMP'%s'" % (tag ,start, end)
    cursor = con.cursor()
    result = cursor.execute(sql)
    rows = result.fetchall()
    cols = []
    for i in result.description:
        cols.append(i[0])
    #print(len(rows))
    data = pd.DataFrame(data=np.array(rows), columns=cols)
    data['VALUE'] = data['VALUE'].astype(float)
    data['VALUE'] = data['VALUE'].round(4)
 # Pandas DataFrame with your data to 4 decimal places!
    data.name = tag[5:10]
    return data

def savgol(x, wl=50, p=5):
    return savgol_filter(x, window_length=wl, polyorder=p)

def differ(Dataframe):
    "adds a numerical derivative column to data frame named 'diff'"
    # Dataframe['rolling mean'] = Dataframe['VALUE'].rolling(100, center=True).mean()
    Dataframe['savgol'] = savgol(Dataframe['VALUE'])
    tdelta = Dataframe['TS'][1] - Dataframe['TS'][0]
    rol = 6
    t = [i * tdelta / pd.to_timedelta(1, unit='H') for i in range(rol)]
    t = np.array(t)
    Dataframe['diff'] = Dataframe['savgol'].rolling(rol, center = True).apply(lambda x:polyfit(t, np.array(x), 1)[0])
    # Dataframe['diff'] = Dataframe['savgol'].diff(periods=4) / ( Dataframe['TS'].diff(periods=4) / pd.to_timedelta(1, unit='H') )
    Dataframe['diff'] = savgol(Dataframe['diff'], 40, 3)
    return

def zeroposneg(value, posneg):
    if value > 0 and posneg == 'pos':
        return value
    if value > 0 and posneg == 'neg':
        return 0
    if value < 0 and posneg == 'pos':
        return 0
    if value < 0 and posneg == 'neg':
        return value

def daily_perf(datetime):
    blender = 'PEL1.MIXR2.PAS.WI.01.PV'
    daily_perf = {}
    data = fetch_tag(blender,0.25,datetime)
    data['TS'] = pd.to_datetime(data['TS'], format='%d-%b-%y %H:%M:%S.%f')
    differ(data)
    data['feed rate'] = data['diff'].apply(lambda x:zeroposneg(x, 'pos'))
    dumprate = data[data['diff']<0]
    daily_perf['dump'] = dumprate['diff'].mean()
    feedrate = data['feed rate'].mean()
    daily_perf['feed'] = feedrate
    return daily_perf

def perf_hourly(start, end):
    data = {}
    # print(int((end-start).days*24)+1)
    for n in range(int((end-start).days*96)+1):
        # print(n)
        current_date = start + timedelta(hours=n*0.25)
        print(current_date)
        data[current_date] = daily_perf(current_date)
    return data

data = perf_hourly(datetime(year = 2022,month = 7,day = 11, hour = 17), datetime(year = 2022, month = 7,day = 15,hour = 17))

with open( "ELCD_rate_report_230711.csv", "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["date", "feed", "dump"])
    for date, rec in data.items():
        writer.writerow([date, rec['feed'], rec['dump']])


