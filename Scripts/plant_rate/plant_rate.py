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

def fetch_tag(tag, start=None, duration=None, end=None, resolution=5):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with 5s intervals as default for the period asked for.
    """
    if start is None and duration is None and end is None:
        raise ValueError("You need to provie at least 2 points of refrence in time to use this function")
    
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
    
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=GLAASPEN-P01;PORT=10014")
    # start = end - timedelta(days = duration)
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period depends on the reporting from the instruments 

    sql = "select \"IP_TREND_TIME\" as TS, \"IP_TREND_VALUE\" as VALUE from \"%s\" "\
            " where TS between TIMESTAMP'%s' and TIMESTAMP'%s'" %(tag ,start_str, end_str)

    # data = pd.read_sql(sql,con) # Pandas DataFrame with your data!
    
    cursor = con.cursor()
    result = cursor.execute(sql)
    rows = result.fetchall()
    cols = []
    for i in result.description:
        cols.append(i[0])
    
    if len(rows) > 0:
        data = pd.DataFrame(data=np.array(rows), columns=cols)
        data['VALUE'] = data['VALUE'].astype(float)
        data['VALUE'] = data['VALUE'].round(4)
    
    elif len(rows)==0:
        print("Warning: no changes detected in {} from {} to {} collecting interpolated data".format(tag, start_str, end_str))
        data = fetch_tag_detailed(tag, start=start, end=end, resolution=50)
    # Pandas DataFrame with your data to 4 decimal places!
    
    data.name = tag
    
    return data 

def fetch_tag_detailed(tag, start=None, duration=None, end=None, resolution=5):

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

    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=GLAASPEN-P01;PORT=10014")
    # start = end - timedelta(days = duration)
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period depends on the reporting from the instruments 

    sql = "select TS, VALUE from HISTORY "\
        " where NAME='%s' and PERIOD=%s*10 and TS between TIMESTAMP'%s' and TIMESTAMP'%s'" %(tag, resolution ,start_str, end_str)

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

def savgol(x, wl=70, p=5):
    return savgol_filter(x, window_length=wl, polyorder=p)

def differ(Dataframe):
    "adds a numerical derivative column to data frame named 'diff'"
    name_str = Dataframe.name
    Dataframe['TS'] = pd.to_datetime(Dataframe['TS'], format='%d-%b-%y %H:%M:%S.%f')
    Dataframe = Dataframe.set_index('TS')
    # print(Dataframe)
    # print(Dataframe[Dataframe.index.duplicated(keep=False)])
    Dataframe = Dataframe[~Dataframe.index.duplicated(keep='last')]
    # Dataframe = Dataframe.resample('5s').ffill()
    # Dataframe = Dataframe.interpolate(method='polynomial', order=3)
    # Dataframe = Dataframe.dropna()
    Dataframe = Dataframe.reset_index()
    Dataframe['savgol'] = savgol_filter(Dataframe['VALUE'],window_length=70 ,polyorder=2, deriv=0)
    tdelta = Dataframe['TS'][1] - Dataframe['TS'][0]
    # rol = 6
    # t = [i * tdelta / pd.to_timedelta(1, unit='H') for i in range(rol)]
    # t = np.array(t)
    # Dataframe['diff'] = Dataframe['savgol'].rolling(rol, center = True).apply(lambda x:polyfit(t, np.array(x), 1)[0])
    Dataframe['diff'] = savgol_filter(Dataframe['VALUE'], window_length=70, polyorder=2, deriv=1)
    Dataframe['diff'] = Dataframe['diff'] * 3600 / tdelta.seconds # this is a correction factor to convert the delta between points into a kg/hr,  
    Dataframe.name = name_str
    return Dataframe

def plot_smooth_data(data):
    data = differ(data)
    fig = make_subplots(rows = 2, cols = 1)
    fig.add_trace(go.Scatter(x=data['TS'].values, y=data['savgol'].values, name='Weight', mode='lines'), row = 1 ,  col = 1)
    fig.add_trace(go.Scatter(x=data['TS'].values, y=data['diff'].values, name= 'Gradient', mode='lines'), row = 2 ,  col = 1)
    print(f"{data.name} feed rate is: {data.loc[((data["diff"] > 100) & (data["diff"] < 2000)), "diff" ].mean()}")
    fig.update_layout(title = data.name)
    fig.show()

def plant_rate_daily(start, end):
    date_range = []
    delta = timedelta(days=1)
    while start <= end:
        date_range.append(start)
        start+= delta
    results = {}
    
    for i in date_range:
        print(i)
        res = {}
        silo_1 = fetch_tag('MIL2.SILO1.GAC.WI.01.PV', start=i,duration=timedelta(days=1))
        silo_1 = differ(silo_1)
        rate_502 = silo_1[silo_1['diff']>100]['diff'].mean()
        res["tk502"] = rate_502
        silo_2 = fetch_tag('MIL2.SILO2.GAC.WI.01.PV', start=i, duration=timedelta(days=1))
        silo_2 = differ(silo_2)
        rate_503 = silo_2[silo_2['diff']>100]['diff'].mean()
        res["tk503"] = rate_503
        plant_rate = rate_502+rate_503
        res["Plant"] = plant_rate
        w01 = fetch_tag('KLN1.MIXR1.SAW.WC.01.PV', start=i,duration=timedelta(days=1))
        w02 = fetch_tag('KLN1.MIXR2.SAW.WC.01.PV', start=i,duration=timedelta(days=1))
        w03 = fetch_tag('KLN1.MIXR3.SAW.WC.01.PV', start=i,duration=timedelta(days=1))
        feed1 = w01[w01['VALUE']>10]['VALUE'].mean()
        if np.isnan(feed1):
            feed1 = 0
        feed2 = w02[w02['VALUE']>10]['VALUE'].mean()
        if np.isnan(feed2):
            feed2 = 0
        feed3 = w03[w03['VALUE']>10]['VALUE'].mean()
        if np.isnan(feed3):
            feed3 = 0
        plant_feed = feed1 + feed2 + feed3
        res['W01'] = feed1
        res['W02'] = feed2
        res['W03'] = feed3
        res['Feed'] = plant_feed
        if plant_feed < 1:
            res['Yeild'] = "Nan"
        else:
            res['Yeild'] = plant_rate / plant_feed
        results[i] = res
    
    with open( "Plant_rate_report.csv", "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["date","TK-502","TK-503", "Total", "W01","W02","W03" , "Feed", "Yeild"])
        for date, rec in results.items():
            writer.writerow([date, rec["tk502"], rec["tk503"],rec["Plant"], rec["W01"], rec["W02"], rec["W03"], rec["Feed"], rec["Yeild"]])
    return

mixer = 'PEL1.MIXR1.PAS.WI.01.PV'
blender = 'PEL1.MIXR2.PAS.WI.01.PV'
t502 = 'MIL2.SILO1.PAC.WI.01.PV'
t503 = 'MIL2.SILO2.PAC.WI.01.PV'
tk502 = 'MIL2.SILO1.GAC.WI.01.PV'
tk503 = 'MIL2.SILO2.GAC.WI.01.PV'
O2 = 'MBR1.TK1.WW.AI.02.PV'


def ctrl_data_filter(df,threshhold):
    df.loc[df['VALUE']<threshhold, 'VALUE'] = 0
    totals = df.loc[(df['VALUE']!=0) & (df['VALUE'].shift(1)==0)]
    return totals

def plot_IMR(df):
    xbar = df['VALUE'].mean()
    df['MR'] = abs(df['VALUE']-df['VALUE'].shift(1))
    mrbar = df['MR'].mean()
    UCL = xbar + mrbar * 2.66
    LCL = xbar - mrbar * 2.66
    rUCL = mrbar * 3.27
    fig, ax = plt.subplots(nrows=2, sharex=True)
    ax1, ax2 = ax.flatten()
    ax1.plot(df['TS'], df['VALUE'], marker='o', linestyle='-', color='red',label='VALUE')
    ax1.set_title(df.name)
    ax1.axhline(xbar, color='green', label='xbar')
    ax1.axhline(UCL, color='orange', label='UCL')
    ax1.axhline(LCL, color='orange', label='LCL')
    ax2.plot(df['TS'], df['MR'], color='red', marker='o', linestyle='-',label='MR')
    ax2.axhline(mrbar, color='green', label='mrbar')
    ax2.axhline(rUCL, color='orange', label='red')
    ax1.set_xlabel('Index')
    ax1.set_ylabel('Value')
    ax2.set_ylabel('MR')
    ax1.legend(loc='upper left')
    ax2.legend(loc='upper right')
    plt.show()
    return 

def Downtime_report(days = [datetime.today()]):
    kilnC_tag = "KLN1.KLN3.UWAC.XS.01.PV"
    return

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

def recipe_changes(day=datetime.today()):
    R_plant_rec_tag = "PEL1.MIXR1.PAS.XI.04.DCS"
    recipies = fetch_tag(R_plant_rec_tag, duration=timedelta(days=1), end=datetime(year=day.year, month=day.month, day=day.day, hour=7, minute=0))
    recipies['CHANGE'] = recipies['VALUE'].ne(recipies['VALUE'].shift())
    chng = recipies[recipies['CHANGE']]
    res = list_of_times(chng)
    return res

def ELCD_control(day=datetime.today()):
    ELCD = {
        "osf_tot":["PEL1.MIXR1.OSF.FQ.01.PV",'{} osf addition to mix'],
        "wf_tot":["PEL1.MIXR1.WF.FQ.01.PV",'{} wf addition to mix'],
        "acid_tot":["PEL1.MIXR1.PPHO.FQ.01.PV",'{} acid addition to mix'],
            }
    data = []
    for i in ELCD:
        temp_1 = fetch_tag(ELCD[i][0], duration=timedelta(days=1), end=datetime(year=day.year, month=day.month, day=day.day, hour=7, minute=0))
        temp_1['TS'] = pd.to_datetime(temp_1['TS'], format='%d-%b-%y %H:%M:%S.%f')
        if i == "wf_tot":
            temp_2 = ctrl_data_filter(temp_1, 10)
        else:
            temp_2 = ctrl_data_filter(temp_1, 100)
        temp_2 = temp_2.iloc[::-1]
        temp_2 = temp_2.reset_index(drop=True)
        temp_2.name = ELCD[i][1].format((day-timedelta(days=1)).strftime("%Y-%B-%d"))
        data.append(temp_2)
    
    return data

# osf , wf, acid = ELCD_control(day=datetime(year=2023, month=6, day=5))
# osf , wf, acid = ELCD_control()

# plot_IMR(osf)
# plot_IMR(acid)


# def main():

#     # start = datetime.strptime(input("start time YY-mm-dd HH:MM-: "), "%Y-%m-%d %H:%M" )

#     data1 = fetch_tag_detailed(tk502, duration=timedelta(hours=24), end=datetime(year=2025, month=11, day=1, hour=9, minute=0))
#     data2 = fetch_tag_detailed(tk503, duration=timedelta(hours=24), end=datetime(year=2025, month=11, day=1, hour=9, minute=0))
#     data1.name = "TK502"
#     data2.name = "TK503"

#     # data3 = fetch_tag(O2, 2, datetime(year=2023, month=5, day=3, hour=7))

#     plot_smooth_data(data1)
#     plot_smooth_data(data2)

# if __name__ == "__main__":
#     main()


# plot_smooth_data(data3)
data1 = fetch_tag_detailed(tk502, duration=timedelta(hours=24), end=datetime(year=2025, month=11, day=1, hour=9, minute=0))
data2 = fetch_tag_detailed(tk503, duration=timedelta(hours=24), end=datetime(year=2025, month=11, day=1, hour=9, minute=0))
data1.name = "TK502"
data2.name = "TK503"

    # data3 = fetch_tag(O2, 2, datetime(year=2023, month=5, day=3, hour=7))

plot_smooth_data(data1)
plot_smooth_data(data2)

# plant_rate_daily(datetime(year=2023,month=3, day=6, hour=7), datetime(year=2023, month=4, day=4, hour=7, ))
#plt.plot(pos['TS'], pos['diff'])
#plt.plot(neg['TS'], neg['diff'])
#plt.show()
