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
import os


script_dir = os.path.dirname(os.path.realpath(__file__))

os.chdir(script_dir)

def fetch_tag(tag, duration, end, resolution=5):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with 10s intervals as default for the period asked for.
    """
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.144;PORT=10014")
    start = end - timedelta(days = duration)
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

def savgol(x, wl=70, p=5):
    return savgol_filter(x, window_length=wl, polyorder=p)

def differ(Dataframe):
    "adds a numerical derivative column to data frame named 'diff'"
    # Dataframe['rolling mean'] = Dataframe['VALUE'].rolling(100, center=True).mean()
    Dataframe['savgol'] = savgol(Dataframe['VALUE'])
    tdelta = Dataframe['TS'][1] - Dataframe['TS'][0]
    rol = 6
    t = [i * tdelta / pd.to_timedelta(1, unit='h') for i in range(rol)]
    t = np.array(t)
    Dataframe['diff'] = Dataframe['savgol'].rolling(rol, center = True).apply(lambda x:polyfit(t, np.array(x), 1)[0])
    # Dataframe['diff'] = Dataframe['savgol'].diff(periods=4) / ( Dataframe['TS'].diff(periods=4) / pd.to_timedelta(1, unit='H') )
    Dataframe['diff'] = savgol(Dataframe['diff'], 40, 3)
    return

def daily_perf(datetime):
    O2 = 'MBR1.TK1.WW.AI.02.PV'
    daily_perf = {}
    data = fetch_tag(O2,1,datetime)
    data['TS'] = pd.to_datetime(data['TS'], format='%d-%b-%y %H:%M:%S.%f')
    differ(data)
    peak = data[data['diff']>0]
    trough = data[data['diff']<0]
    peakdata = peak['diff']
    troughdata = trough['diff']
    delta = peakdata.mean()-troughdata.mean()
    daily_perf['peak'] = peakdata.max()
    daily_perf['trough'] = troughdata.min()
    daily_perf['ave_consumption'] = troughdata.mean()
    daily_perf['ave delta'] = delta
    return daily_perf

# print(daily_perf(datetime(year=2023, month=3, day=17, hour=7)))

def o2_perf_daily(start, end):
    data = {}
    for n in range(int((end-start).days)+1):
        
        current_date = start + timedelta(n)
        print(current_date)
        data[current_date] = daily_perf(current_date)
    return data

dt_format = "%Y-%m-%d %H"
start_date = None
while start_date==None:
    try:
        start_date = datetime.strptime(input("Start date (YYYY-MM-DD): ")+" 07",dt_format)
    except ValueError:
        print("Entry must be in the YYYY-MM-DD format")
        pass

end_date = None
while end_date==None:
    try:
        end_date =  datetime.strptime(input("End Date(YYYY-MM-DD)(Default = Today): ")+" 07" or datetime.today().strftime(dt_format)+ " 07",dt_format)
        pass
    except ValueError:
        print("Entry must be in the YYYY-MM-DD format")
        pass

data = o2_perf_daily(start_date, end_date)

with open( "O2_report.csv", "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["date","peak","trough", "Mean O2 consumption rate","mean delta"])
    for date, rec in data.items():
        writer.writerow([date, round(rec["peak"], 4), round(rec["trough"], 4),round(rec["ave_consumption"], 4),round(rec["ave delta"], 4)])


os.system("start EXCEL.EXE O2_report.csv")



