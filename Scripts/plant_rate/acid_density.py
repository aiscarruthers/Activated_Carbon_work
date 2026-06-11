
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
import time
import os

df = pd.read_csv("C:\\Users\\ACarruther\\Documents\\Campaign Analysis\\Nov-2022_Evidence\\Lab_acid_density.csv")

df['Datetime'] = pd.to_datetime(df['Datetime'], format='%d/%m/%Y %H:%M')

def isformated(x):
    if x<2:
        return x*1000
    else:
         return x

df["Lab acid Density"] = df["Lab acid Density"].apply(lambda x: isformated(x))

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

new_den = []
for i in df['Datetime']:
    dens = fetch_tag_avg('TK1.TK3.PPHO.DI.01.PV', 0.021, i + timedelta(hours = 0.5))
    new_den.append(dens)
df['Aspen data'] = new_den
df['Datetime'] = df['Datetime'].dt.strftime("%Y/%m/%d %H:%M")

df.to_csv("C:\\Users\\ACarruther\\Documents\\Campaign Analysis\\Nov-2022_Evidence\\Lab_density_res.csv", date_format='%Y/%m/%d %H:%M')


