
import pyodbc
import datetime as dt
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO

def fetch_tag(self, tag, start=None, duration=None, end=None, resolution=5):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with 5s intervals as default for the period asked for.
    """
    if start is None and duration is None and end is None:
        raise ValueError("You need to provie at least 2 points of referance in time to use this function")

    if start is None:
        if duration is None or end is None:
            raise ValueError("If you provide no start you must provide an end or duration")
        start = end - duration

    elif duration is None:
        if end is None:
            raise ValueError("If no duration is provided you must provide an end with a start")
        duration = start - end

    elif end is None:
        end = start + duration

    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.47;PORT=10014")
    # start = end - timedelta(days = duration)
    end = end.strftime("%Y-%m-%d %H:%M:%S")
    start=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period depends on the reporting from the instruments 

    sql = "select \"IP_TREND_TIME\" as TS, \"IP_TREND_VALUE\" as VALUE from \"%s\" "\
        " where TS between TIMESTAMP'%s' and TIMESTAMP'%s'" %(tag ,start, end)

    # data = pd.read_sql(sql,con) # Pandas DataFrame with your data!

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

    data.name = tag

    return data

def filter(df):
    df.loc[df['VALUE']< 100, 'VALUE'] = 0
    totals = df.loc[(df['VALUE']!=0) & (df['VALUE'].shift(1)==0)].copy()
    totals 