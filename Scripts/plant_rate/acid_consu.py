import pyodbc
import pandas as pd
from datetime import datetime
from datetime import timedelta
import matplotlib.pyplot as plt
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from scipy import integrate
import numpy as np
from numpy import polyfit
import time
import os


def fetch_tag(tag, duration, end, resolution=0.1):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with 5s intervals as default for the period asked for.
    """
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.47;PORT=10014")
    start = end - timedelta(days = duration)
    end = end.strftime("%Y-%m-%d %H:%M:%S")
    start=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period is in tenths of seconds 
    sql = "select TS, VALUE from HISTORY "\
            "where NAME='%s' "\
            "and PERIOD=%s*10 "\
            "and REQUEST=2 "\
            "and TS between TIMESTAMP'%s' and TIMESTAMP'%s'" % (tag, resolution ,start, end)
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

def fetch_tag_avg(tag, duration, end):
    """
    Takes a tag name, duration and end point as a datetime object and submits 
    a sql query returning the average value of the tag over that time period.
    """
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.47;PORT=10014")
    start = end - timedelta(days = duration)
    end = end.strftime("%Y-%m-%d %H:%M:%S")
    start=start.strftime("%Y-%m-%d %H:%M:%S")
    
    sql = "select AVG(VALUE) from HISTORY "\
    "where NAME='%s' "\
    "and REQUEST=2 "\
    "and TS between TIMESTAMP'%s' and TIMESTAMP'%s'" % (tag ,start, end)
    cursor = con.cursor()
    result = cursor.execute(sql)
    data = result.fetchall()
    return data[0][0]

def fetch_tag_stddev(tag, duration, end):
    """
    Takes a tag name, duration and end point as a datetime object and submits 
    a sql query returning the average value of the tag over that time period.
    """
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.47;PORT=10014")
    start = end - timedelta(days = duration)
    end = end.strftime("%Y-%m-%d %H:%M:%S")
    start=start.strftime("%Y-%m-%d %H:%M:%S")
    
    sql = "select STDDEV(VALUE) from HISTORY "\
    "where NAME='%s' "\
    "and REQUEST=2 "\
    "and TS between TIMESTAMP'%s' and TIMESTAMP'%s'" % (tag ,start, end)
    cursor = con.cursor()
    result = cursor.execute(sql)
    data = result.fetchall()
    return data[0][0]


def acid_balance(date, duration):
    """
    Takes a date and performs a volumetric balance over the day
    """
    dict = {}
    tags =  {   "F09":"WSH1.TK2.PPHO.FC.01.PV",
                "F11" : "TK1.TK1.PHO.FC.01.PV",
                "F28" : "TK1.TK2.PPHO.FC.01.PV",
                "F04" : "KLN1.MIXR1.PPHO.FC.01.PV",
                "F05" : "KLN1.MIXR2.PPHO.FC.01.PV",
                "F06" : "KLN1.MIXR3.PPHO.FC.01.PV",
                "L04" : "TK1.TK6.PPHO.LI.01.PV"}
    for i in tags:

        if i == "L04":
            data = fetch_tag(tags[i], duration, date)
            dict[i+"start"] = data['VALUE'][:100].mean()
            dict[i+"end"] =  data['VALUE'][-100:].mean()
        else:
            data = fetch_tag(tags[i], duration, date)
            data['TS']= pd.to_datetime(data['TS'], format='%d-%b-%y %H:%M:%S.%f')
            time_values = (data["TS"] - data["TS"].iloc[0]).dt.total_seconds().values
            time_values = time_values / 3600
            Integ = integrate.trapz(data["VALUE"], time_values)
            dict[i] = Integ
    
    IN = dict["F09"] + dict["F11"] - dict["F28"]
    OUT = dict["F04"] + dict["F05"] + dict["F06"]
    ACCU = dict["L04end"] - dict["L04start"]
    if OUT > 500:
        return IN - OUT - ACCU
    else:
        return 0

def bulk_acid_balance(start, no_days_previous, resolution):
    base = start
    tenth_of_days = range(no_days_previous*resolution, 0, -1)
    dates = [ i/resolution for i in tenth_of_days ] 
    date_list = [base - timedelta(days = x) for x in dates]
    delta = date_list[0] - date_list[1]
    results = []
    for j, i in enumerate(date_list):
        results.append(acid_balance(i, 0.1))
        print( "{} of {}".format(j, len(date_list)))
    dict = { 'TS': date_list, "VALUES": results}
    df = pd.DataFrame(dict)
    return df 

def plot_bulk_acid_results(start, no_days_previous):
    data = bulk_acid_balance(start, no_days_previous, 10)
    plt.plot(data['TS'],data['VALUES'])
    plt.title('Extra or Missing Acid')
    plt.ylabel("Volume of Acid returned more than going into mix")
    plt.xlabel("day")
    plt.grid()
    plt.show()
    plt.savefig(os.path.expanduser('~/Downloads/Acid_balance.png'))

def integrate_tag(tag, duration, end):
    data = fetch_tag(tag, duration, end)
    #print(data['TS'].iloc[0])
    #print(data['TS'].iloc[-1])
    data['TS']= pd.to_datetime(data['TS'], format='%d-%b-%y %H:%M:%S.%f')

    time_values = (data["TS"] - data["TS"].iloc[0]).dt.total_seconds().values
    #print(time_values[-1])
    time_values = time_values / 3600
    #print(time_values[-1])
    Integ = integrate.trapz(data["VALUE"], time_values)
    return Integ


tags = {"F04" : "KLN1.MIXR1.PPHO.FC.01.PV",
"F05" : "KLN1.MIXR2.PPHO.FC.01.PV",
"F06" : "KLN1.MIXR3.PPHO.FC.01.PV"}

plot_bulk_acid_results(datetime(year=2023, month=1, day=11, hour=7), 48)
# print(timedelta(days=0.29))

# data = fetch_tag("KLN1.MIXR1.PPHO.FC.01.PV", 0.1, datetime(year=2023, month=1, day=6, hour=0))
# print(len(data['TS']))
# data = fetch_tag("KLN1.MIXR1.PPHO.FC.01.PV", 0.5, datetime(year=2023, month=1, day=6, hour=0))
# print(len(data['TS']))


# for i in tags:
#     INT = integrate_tag( tags[i], 0.1, datetime(year=2023, month=1, day=6, hour=0))
#     print(INT/2.4)

