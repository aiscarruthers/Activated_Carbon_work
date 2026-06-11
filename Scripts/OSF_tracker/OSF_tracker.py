import pyodbc
import pandas as pd
from datetime import datetime
from datetime import timedelta
import numpy as np
import os
def alphabet_enumerate(iterable):
    def get_label(index):
        alphabet = 'abcdefghijklmnopqrstuvwxyz'
        label = ''
        while index >= 0:
            label = alphabet[index % 26] + label
            index = index // 26 - 1
        return label

    for i, item in enumerate(iterable):
        yield (get_label(i), item)



def fetch_tag(tags, start=None, duration=None, end=None, resolution=5):
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
    
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.144;PORT=10014")
    # start = end - timedelta(days = duration)
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period depends on the reporting from the instruments 
    
    sql_2 = ""
    sql_3 = ""

    Desc_1 = tags
    
    
    for i,j in alphabet_enumerate(tags):
        if i == "a": #skip the fist element 
            Desc_1 = tags[j]["Desc"]
            tag_1 = j
            continue
        sql_2 += f",\n\t{i}.IP_TREND_VALUE \"{tags[j]["Desc"]}\""
        sql_3 += f"\nFULL JOIN \n\t\"{j}\" {i} USING (IP_TREND_TIME)"    
    
    sql_1 = "select \"IP_TREND_TIME\" as TS,\n"\
        f"\ta.IP_TREND_VALUE \"{Desc_1}\" "\
        f"{sql_2}"\
        f"\nfrom \n\t\"{tag_1}\" a"\
        f"{sql_3}"\
        f"\nwhere TS between TIMESTAMP'{start_str}' and TIMESTAMP'{end_str}'"\
        "\norder by TS;"
    
    # data = pd.read_sql(sql,con) # Pandas DataFrame with your data!
    # print(sql_1)
    cursor = con.cursor()
    result = cursor.execute(sql_1)
    rows = result.fetchall()
    cols = []

    for i in result.description:
        cols.append(i[0])
    
    if len(rows) > 0:
        data = pd.DataFrame(data=np.array(rows), columns=cols)
    #     data['VALUE'] = data['VALUE'].astype(float)
    #     data['VALUE'] = data['VALUE'].round(4)
    data.fillna(method='ffill',inplace=True)
    # elif len(rows)==0:
    #     print("Warning: no changes detected in {} from {} to {} collecting interpolated data".format(tag, start_str, end_str))
    #     data = fetch_tag_detailed(tag, start=start, end=end, resolution=50)
    # # Pandas DataFrame with your data to 4 decimal places!
    
    data.name = "OSF_data"

    # print(data)
    return data

def convert_periods(df, column):
    periods = []
    start_time = None
    current_state = None
    
    for i, row in df.iterrows():
        if row[column] != current_state:
            if current_state is not None:
                period_duration = row['TS'] - start_time
                if period_duration >= timedelta(seconds=300):
                    periods.append({'State': current_state, 'Start Time': start_time, 'Stop Time': row['TS']})
                else:
                    # Absorb short period into the previous period
                    if len(periods) > 0:
                        periods[-1]['Stop Time'] = row['TS']
            start_time = row['TS']
            current_state = row[column]
    
    # Append the last period
    if current_state is not None:
        periods.append({'State': current_state, 'Start Time': start_time, 'Stop Time': df.iloc[-1]['TS']})
    
    return pd.DataFrame(periods)
# Function to absorb adjacent periods into each other
def absorb_adjacent_periods(periods):
    absorbed_periods = []
    current_period = None
    
    for i, row in periods.iterrows():
        if current_period is None:
            current_period = row
        else:
            if row['State'] == current_period['State']:
                current_period['Stop Time'] = row['Stop Time']
            else:
                absorbed_periods.append(current_period)
                current_period = row
    
    if current_period is not None:
        absorbed_periods.append(current_period)
    
    return pd.DataFrame(absorbed_periods)
# Function to split periods crossing the 07:00 mark
def split_crossing_periods(periods):
    split_periods = []
    
    for i, row in periods.iterrows():
        start_time = row['Start Time']
        stop_time = row['Stop Time']
        
        while start_time < stop_time:
            next_split = start_time.replace(hour=7, minute=0) + pd.Timedelta(days=1)
            if next_split < stop_time:
                split_periods.append({'State': row['State'], 'Start Time': start_time, 'Stop Time': next_split})
                start_time = next_split
            else:
                split_periods.append({'State': row['State'], 'Start Time': start_time, 'Stop Time': stop_time})
                break
    
    return pd.DataFrame(split_periods)

osf_tag_dict = {
    "PEL1.DRYB1.PAS.XS.01.PV": {"Desc":"R-plant", "Area":"R-plant"},
    "KLN1.KLN3.UWAC.XS.01.PV": {"Desc":"Kiln C", "Area": "Activation"},
    "KLN1.KLN2.UWAC.XS.01.PV": {"Desc":"Kiln B","Area":"Activation"},
    "KLN1.KLN1.UWAC.XS.01.PV": {"Desc":"Kiln A","Area":"Activation"},
    "WSH1.WSH1.UWAC.XS.01.PV": {"Desc":"Prayon","Area": "Washing"},
    "DRYR1.BURN.FG.XI.01.PV": {"Desc":"Rotary Dryer","Area":"Drying"},
    "MIL2.MILB1.PAC.XS.01.PV": {"Desc":"Carbon Mill","Area":"Milling"},
    "MIL2.MIXR1.PAC.BV.04.PV": {"Desc":"Packaging A Blender","Area":"Packaging"},
    "MIL2.MIXR2.PAC.BV.04.PV": {"Desc":"Packaging B Blender","Area":"Packaging"}
}

start = datetime(year=2024,month=1,day=1,hour=7)
end =  datetime(year=2025,month=1,day=1,hour=7)
data = fetch_tag(tags = osf_tag_dict,start=start,end=end)

data["TS"] = pd.to_datetime(data["TS"])
result = pd.DataFrame()
for column in data.columns[1:]:
    column_periods = convert_periods(data,column)
    column_periods = absorb_adjacent_periods(column_periods)
    column_periods = split_crossing_periods(column_periods)
    column_periods["Asset"] = column
    result = pd.concat([result,column_periods],ignore_index=True)

print(result)

result.to_excel("osf_report.xlsx")
os.system("start EXCEL.EXE osf_report")