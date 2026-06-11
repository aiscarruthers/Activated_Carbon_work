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


def fetch_tag(tag, start, end, resolution=0.1):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with 5s intervals as default for the period asked for.
    """
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.47;PORT=10014")
    
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

def fetch_rec_delta(start, finish):
    recipie = "0.UTIL0.RECIPE.NUM"
    data = fetch_tag(recipie, start, finish)
    data['Delta'] = data["VALUES"].diff()
    for i , j in enumerate(data['Delta']):
        if j != 0:
            data['TS'] =
