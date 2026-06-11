import os
import json 
from prod_scrape import WorkOrder
import pandas as pd
import pyodbc
import datetime as dt
import numpy as np

def fetch_saw_dust(start=None, duration=None, end=None):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with intervals prescribed by the IP21 server. It is important to 
    note that the intervals will vary depending on the tag being queried. 
    
    Parameters:
    tag(str): the aspen tag data is being requested for
    start(datetime object): this marks the start of the data being queried
    duration(timedelta object): this marks the duration of the data being queried
    end(datetime object): this marks the end of the data being queried
        
    Returns:
    Dataframe: A pandas Dataframe of the data being queried if the data exists

    Raises:
    Value Error: if no Data exists in the period being queried

    """
    tags = {"W01":"KLN1.MIXR1.SAW.WC.01.PV",
            "W02":"KLN1.MIXR2.SAW.WC.01.PV",
            "W03":"KLN1.MIXR3.SAW.WC.01.PV"}
    
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
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period depends on the reporting from the instruments 
    sql_string = f"""
        SELECT SUM(avg_value) AS SAWDUST
        FROM (
            SELECT AVG("IP_TREND_VALUE") AS avg_value
            FROM "{tags["W01"]}"
            WHERE "IP_TREND_TIME" BETWEEN TIMESTAMP'{start_str}' AND TIMESTAMP'{end_str}'
            UNION ALL
            SELECT AVG("IP_TREND_VALUE") AS avg_value
            FROM "{tags["W02"]}"
            WHERE "IP_TREND_TIME" BETWEEN TIMESTAMP'{start_str}' AND TIMESTAMP'{end_str}'
            UNION ALL
            SELECT AVG("IP_TREND_VALUE") AS avg_value
            FROM "{tags["W03"]}"
            WHERE "IP_TREND_TIME" BETWEEN TIMESTAMP'{start_str}' AND TIMESTAMP'{end_str}'
        ) subquery
        """
    
    cursor = con.cursor()
    result = cursor.execute(sql_string)
    rows = result.fetchall()
    value = rows[0][0]
    return value

def fetch_moisture(start=None, duration=None, end=None):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with intervals prescribed by the IP21 server. It is important to 
    note that the intervals will vary depending on the tag being queried. 
    
    Parameters:
    tag(str): the aspen tag data is being requested for
    start(datetime object): this marks the start of the data being queried
    duration(timedelta object): this marks the duration of the data being queried
    end(datetime object): this marks the end of the data being queried
        
    Returns:
    Dataframe: A pandas Dataframe of the data being queried if the data exists

    Raises:
    Value Error: if no Data exists in the period being queried

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
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")
    
    tag = "KLN1.MIXR2.SAW.AI.01.PV"
    
    # time period depends on the reporting from the instruments 

    sql = f"select AVG(\"IP_TREND_VALUE\") as VALUE from \"{tag}\" "\
        f" where \"IP_TREND_TIME\" between TIMESTAMP'{start_str}' and TIMESTAMP'{end_str}'"

    # data = pd.read_sql(sql,con) # Pandas DataFrame with your data!
    
    cursor = con.cursor()
    result = cursor.execute(sql)
    rows = result.fetchall()
    result = rows[0][0]
    
    return result


def sawdust_check(row):
    if type(row["Sawdust"]) is float: 
        start = row["start"].to_pydatetime()
        end = row["end"].to_pydatetime()
        time_delta = (end - start).total_seconds() / 3600
        value = fetch_saw_dust(start=start, end=end)
        try:
            total = value * time_delta
        except TypeError:
            print(f"value is {type(value)} type and time delta is {type(time_delta)} returning 0")
            return 0
        
        return total
    else:
        return 0

def moisture_check(row):
    if type(row["Sawdust"] is float):
        start = row["start"].to_pydatetime()
        end = row["end"].to_pydatetime()
        value = fetch_moisture(start=start, end=end)
        
        return value
    else:
        return 0
        

def main():
    json_directory = os.path.dirname(os.path.abspath(__file__))
    json_file_path = os.path.join(json_directory,'WO_data.json')
    with open(json_file_path, 'r') as json_file:
        data = json.load(json_file)
    
    new_list =[]
    for i in data:
        new_WO = {}
        print(data[i])
        if type(data[i]) is dict:
            for j in data[i]:
                if type(data[i][j]) is dict:
                    for k in data[i][j]:
                        new_WO[k] = data[i][j][k]
                else:
                    new_WO[j] = data[i][j]
        new_list.append(new_WO)
    for i in new_list:
        print(i)
    
    final_dump = pd.DataFrame(new_list)
    final_dump["start"] = pd.to_datetime(final_dump["start"])
    final_dump["end"] = pd.to_datetime(final_dump["end"])
    final_dump["Sawdust check"] = final_dump.apply(sawdust_check, axis=1)
    final_dump["Sawdust moisture"] = final_dump.apply(moisture_check, axis=1)
    
    print(final_dump)
    WO_excel_path = os.path.join(json_directory,"WO_data.xlsx")
    final_dump.to_excel(WO_excel_path, index=False)


if __name__=="__main__":
    main()




    

