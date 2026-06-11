import pyodbc
import pandas as pd
from datetime import datetime
from datetime import timedelta
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker

import matplotlib.pyplot as plt
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from scipy.signal import savgol_filter
import numpy as np
from numpy import polyfit
import csv


def fetch_tag(tag, duration, end, resolution=5):
    """
    Takes a tag name, duration, and end point and submits a raw SQL query returning
    data with 10s intervals as default for the period asked for.
    """
    # Replace the following with your SQLAlchemy connection details
    connection_string = 'aiosqlite+pyodbc://GLdata?driver=AspenTech SQLplus'
    
    engine = create_engine(connection_string)
    Session = sessionmaker(bind=engine)
    session = Session()

    try:
        # Check the default schema
        default_schema_query = text("SELECT USER_NAME()")
        default_schema = session.execute(default_schema_query).scalar()
        print(f"{default_schema}")
        
        start = end - timedelta(days=duration)
        end = end.strftime("%Y-%m-%d %H:%M:%S")
        start = start.strftime("%Y-%m-%d %H:%M:%S")

        # time period is in tenths of seconds
        query = text(f"""
            SELECT "IP_TREND_TIME" AS "TS", "IP_TREND_VALUE" AS "VALUE"
            FROM "{tag}"
            WHERE "IP_TREND_TIME" BETWEEN '{start}' AND '{end};'
        """)

        # Execute the query and fetch the results into a Pandas DataFrame
        result = session.execute(query)
        rows = result.fetchall()
        cols = result.keys()
        data = pd.DataFrame(data=rows, columns=cols)

        # Clean up
        session.close()

        data['VALUE'] = data['VALUE'].astype(float)
        data['VALUE'] = data['VALUE'].round(4)
        data.name = tag[5:10]
        return data

    except Exception as e:
        print(f"Error: {e}")
    
    finally:
        session.close()


def savgol(x, wl=70, p=5):
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

data = o2_perf_daily(datetime(2019,1,10,7), datetime(2019,1,14,7))

with open( "O2_report.csv", "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["date","peak","trough", "Mean O2 consumption rate","mean delta"])
    for date, rec in data.items():
        writer.writerow([date, rec["peak"], rec["trough"],rec["ave_consumption"],rec["ave delta"]])




