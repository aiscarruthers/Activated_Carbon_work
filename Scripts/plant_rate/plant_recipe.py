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
import re
import time
import os
import xlsxwriter

def convert_to_float(val):
    pattern = re.compile(r'^-?\d+(\.\d+)?$')
    if pattern.match(val):
        return float(val)
    else:
        return val
    
def fetch_tag(tag, duration, end, resolution=600):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with 10s intervals as default for the period asked for.
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
    data = pd.DataFrame(data=np.array(rows), columns=cols)
    data['VALUE'] = data['VALUE'].apply(convert_to_float)

    data['TS']= pd.to_datetime(data['TS'], format='%d-%b-%y %H:%M:%S.%f')
    # data['VALUE'] = data['VALUE'].fillna(data['VALUE'])
    # data['VALUE'] = data['VALUE'].round(0)
    # data['VALUE'] = data['VALUE'].astype(int)
 # Pandas DataFrame with your data!
    data.name = tag
    return data

def fetch_tag_stst(tag, start, end, resolution=5):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with 10s intervals as default for the period asked for.
    """
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.47;PORT=10014")
    end = end.strftime("%Y-%m-%d %H:%M:%S")
    start=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period is in tenths of seconds 
    sql = "select \"IP_TREND_TIME\" as TS, \"IP_TREND_VALUE\" as VALUE from \"%s\" "\
            " where TS between TIMESTAMP'%s' and TIMESTAMP'%s'" %(tag ,start, end)
    
    cursor = con.cursor()
    result = cursor.execute(sql)
    rows = result.fetchall()
    cols = []
    for i in result.description:
        cols.append(i[0])
    data = pd.DataFrame(data=np.array(rows), columns=cols)
    data['TS']= pd.to_datetime(data['TS'], format='%d-%b-%y %H:%M:%S.%f')
    data['VALUE'] = pd.to_numeric(data['VALUE'], errors='coerce')
    # data['VALUE'] = data['VALUE'].fillna(data['VALUE'])
    data['VALUE'] = data['VALUE'].round(0)
    data['VALUE'] = data['VALUE'].astype(int)
 # Pandas DataFrame with your data!
    data.name = tag
    return data


milling_recipes = {
    1:"C*SPF", 2:"C*SPF", 3:"C*SPF", 4:"C*SPF",
    5:"C*SP", 6:"C*SP", 7:"C*SP", 8:"C*SP",
    9:"C*1P", 10:"C*1P", 11:"C*1P", 12:"C*1P",
    13:"PUREFLOW C", 14:"PUREFLOW C", 15:"PUREFLOW C",
    16:"C*1", 17:"C*1", 18:"C*1", 19:"C*1", 20:"C*1", 21:"C*1", 22:"C*1", 23:"C*1", 24:"C*1",
    25:"C*1F", 26:"C*1F",
    27:"C*1FW",
    28:"C*1W",
    29:"LAC",
    30:"Bag splitter",
    31:"C*1W",
    32:"C*2F", 33:"C*2F", 34:"C*2F",
    35:"C*3SPF",
    36:"C*3", 37:"C*3", 38:"C*3", 39:"C*3", 40:"C*3",
    41:"C*3C",
    42:"C*4", 43:"C*4", 44:"C*4",
    45:"C*4W",
    46:"C*5",
    47:"GB1", 48:"GB1", 49:"GB1",
    50:"GB2",
    51:"GBG", 52:"GBG",
    53:"GBG1",
    54:"ADC/W35/D10",
    55:"ADC/W35", 56:"ADC/W35",
    57:"GL", 58:"GL", 59:"GL",
    60:"W20", 61:"W20", 62:"W20",
    63:"SXRO", 64:"SXRO",
    65:"KB-WJ/CAP SUPER", 66:"KB-WJ/CAP SUPER",
    67:"KB-G", 68:"KB-G",
    69:"CAUF",
    70:"MELCD",
    71:"KB/KBB",
    72:"GBSP",
    73:"CAP SUPER WET",
    74:"C*1C", 75:"C*1C",
    "Stopped":"Stopped", "Nan":"Stopped"
}

# data = fetch_tag('MILL2.MILB1.PAC.XI.01.DCS', 50, datetime(year=2023, month=4, day=15, hour=7), 600)

# data['CHANGE'] = data['VALUE'].ne(data['VALUE'].shift())
# data["GRADE"] = data["VALUE"].map(milling_recipes)
# print(data[data['CHANGE']])

# dat = fetch_tag('MIL2.UNIT.RUN.NUMBER', 50, datetime(year=2023, month=4, day=15, hour=7), 600)
# dat['CHANGE'] = data['VALUE'].ne(data['VALUE'].shift())
# print(dat[dat['CHANGE']])

# da = fetch_tag("0.UTIL0.RECIPE.NUM", 50, datetime(year=2023, month=4, day=15, hour=7), 60)
# da['CHANGE'] = da['VALUE'].ne(da['VALUE'].shift())
# print(da[da['CHANGE']])

def list_of_times(df):
    result = []
    start_index = None
    for i, row in df.iterrows():
        if row['CHANGE'] == 1 and start_index is None:
            j = row['TS']
            start_index = i
        elif row['VALUE'] == 0 and start_index is not None:
            k = row['TS']
            result.append([j,k])
            start_index = None
    return result

def print_changes(tag, duration, end):
    da = fetch_tag(tag, duration, end, 60)
    da['CHANGE'] = da['VALUE'].ne(da['VALUE'].shift())
    new = da[da['CHANGE']]
    print(da)
    res = list_of_times(new)
    # print(res)
    return

def print_changes_mill( duration, end):
    da = fetch_tag("MILL2.MILB1.PAC.XI.01.DCS", duration, end)
    da['CHANGE'] = da['VALUE'].ne(da['VALUE'].shift())
    da["GRADE"] = da["VALUE"].map(milling_recipes)
    print(da[da['CHANGE']])
    return

def print_changes_recipe(duration, end):
    recipies = {1:2, 2:3, 3:2, 4:3, 5:2, 6:3, 12:1}
    df = fetch_tag("0.UTIL0.RECIPE.NUM", duration, end)
    df['CHANGE'] = df['VALUE'].ne(df['VALUE'].shift())
    df["Kiln numbers"] = df["VALUE"].map(recipies)
    with pd.ExcelWriter("C:\\Users\\ACarruther\\recipie_changes.xlsx", datetime_format="yyyy-mm-dd HH:MM:SS.00")as writer:
        df.to_excel(writer)
    print(df)
    return

def running(duration, end):
    df = fetch_tag("KLN1.KLN3.UWAC.XS.01.PV", duration, end)
    print(df)
    with pd.ExcelWriter("C:\\Users\\ACarruther\\running.xlsx", datetime_format="yyyy-mm-dd HH:MM:SS.00")as writer:
        df.to_excel(writer)
    return

ending = datetime(year=2023, month=4, day=6)
# print_changes("MIL2.MILB1.PAC.XS.01.PV", 50, datetime(year=2023,month=4, day=17))
# print_changes_mill(120, datetime.now())
print_changes_recipe(30,ending)
running(30 , ending)



# def iseven(x):
#     if x % 2 == 0 and x!= 12:
#         return 1
#     else:
#          return 0

# data["2kilns"] = data['VALUE'].apply(lambda x: iseven(x))

# kiln = data["2kilns"]

# for i in range(52):
#     print(kiln[i*288:(i+1)*288].mean())
