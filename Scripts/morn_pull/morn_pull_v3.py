import pyodbc
import re
import numpy
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
import json
import xlsxwriter as xlw
import pylightxl as xl
from xlsxpandasformatter import FormattedWorksheet


def Tag_query(tag, command="AVG",start=None, duration=None, end=None, resolution=5):
    """
    Takes a tag name, command, duration and end point and submits a sql query 
    returning data summarising the tag over this duration.
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
    
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=GLAASPEN-P01;PORT=10014")  # 192.168.9.144
    # start = end - timedelta(days = duration)
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")

    # time period depends on the reporting from the instruments 
    if command not in ('AVG','STDDEV', 'VARIANCE', 'COUNT','SUM', 'MIN', 'MAX'):
        raise ValueError("""This function will only accept the following commands: 
                         COUNT, MIN, MAX, SUM, AVG, STDDEV, VARIANCE""")
    
    sql = "select %s(\"IP_TREND_VALUE\") as \"VALUE\" from \"%s\" "\
            " where \"IP_TREND_TIME\" between TIMESTAMP'%s' and TIMESTAMP'%s'" %(command, tag, start_str, end_str)

    # data = pd.read_sql(sql,con) # Pandas DataFrame with your data!
    
    cursor = con.cursor()
    result = cursor.execute(sql)
    rows = result.fetchall()
    # print(rows)
    cols = []
    for i in result.description:
        cols.append(i[0])
    
    if len(rows) > 0:
        data = pd.DataFrame(data=np.array(rows), columns=cols)
        data['VALUE'] = data['VALUE'].astype(float)
        data['VALUE'] = data['VALUE'].round(4)
        data.name = tag
    
    elif len(rows)==0:
        print("Warning: no changes detected in {} from {} to {} collecting interpolated data".format(tag, start_str, end_str))
    
    # Pandas DataFrame with your data to 4 decimal places!
    return data

def x_y_comparision(tag, pv_tag, op_tag, sp_tag, start=None, duration=None, end=None):
    """
    Takes a collection of Process Value, Output, and Setpoint tags, duration and
    end point and submits a sql query returning a table of PV, OP, and SP data
    with 1s intervals.
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
    
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=GLAASPEN-P01;PORT=10014" )# 192.168.9.144
    # start = end - timedelta(days = duration)
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")
    
    sql = 'select a.TS "Datetime", a.VALUE "PV", b.VALUE "OP", c.VALUE "SP"'\
        f'from (select "TS", "VALUE" from history where NAME=\'{pv_tag}\'and '\
        f'PERIOD=10 and TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\') a, '\
        f'(select "TS", "VALUE" from history where NAME=\'{op_tag}\' ' \
        f'and PERIOD=10 and TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\') b, '\
        f'(select "TS", "VALUE" from history where NAME=\'{sp_tag}\'and '\
        f'PERIOD=10 and STEPPED=1 and TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' ) c '\
        'where a.TS = b.TS and b.TS = c.TS '\
        f'and a.TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' '\
        f'and b.TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' '\
        f'and c.TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' '\
        'ORDER BY a.TS'

    # data = pd.read_sql(sql,con) # Pandas DataFrame with your data!
    
    cursor = con.cursor()
    result = cursor.execute(sql)
    rows = result.fetchall()
    # print(rows)
    cols = []
    for i in result.description:
        cols.append(i[0])
    
    if len(rows) > 0:
        data = pd.DataFrame(data=np.array(rows), columns=cols)
        data['PV'] = data['PV'].astype(float)
        data['PV'] = data['PV'].round(4)
        data['OP'] = data['OP'].astype(float)
        data['OP'] = data['OP'].round(4)
        data.name = tag
        return data
    
    elif len(rows)==0:
        print("Warning: no changes detected in {} from {} to {} collecting interpolated data".format(tag, start_str, end_str))
        return None
    # Pandas DataFrame with your data to 4 decimal places!

def SP_eval(tag, pv_tag, sp_tag, start=None, duration=None, end=None):
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
    
    con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=GLAASPEN-P01;PORT=10014")# 192.168.9.144
    # start = end - timedelta(days = duration)
    end_str = end.strftime("%Y-%m-%d %H:%M:%S")
    start_str=start.strftime("%Y-%m-%d %H:%M:%S")
    
    sql = 'select a.TS "Datetime", a.VALUE "PV", c.VALUE "SP"'\
        f'from (select "TS", "VALUE" from history where NAME=\'{pv_tag}\'and '\
        f'PERIOD=10 and TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\') a, '\
        f'(select "TS", "VALUE" from history where NAME=\'{sp_tag}\'and '\
        f'PERIOD=10 and STEPPED=1 and TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' ) c '\
        'where a.TS = c.TS '\
        f'and a.TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' '\
        f'and c.TS between TIMESTAMP\'{start_str}\' and TIMESTAMP\'{end_str}\' '\
        'ORDER BY a.TS'

    # data = pd.read_sql(sql,con) # Pandas DataFrame with your data!
    
    cursor = con.cursor()
    result = cursor.execute(sql)
    rows = result.fetchall()
    # print(rows)
    cols = []
    for i in result.description:
        cols.append(i[0])
    
    if len(rows) > 0:
        data = pd.DataFrame(data=np.array(rows), columns=cols)
        data['PV'] = data['PV'].astype(float)
        data['PV'] = data['PV'].round(4)
        # data['OP'] = data['OP'].astype(float)
        # data['OP'] = data['OP'].round(4)
        data.name = tag
        return data
    
    elif len(rows)==0:
        print("Warning: no changes detected in {} from {} to {} collecting interpolated data".format(tag, start_str, end_str))
        return None
    
    # Pandas DataFrame with your data to 4 decimal places!

class Tag_dict():
    def __init__(self, dictionary, start, end):
        self.dict = dictionary
        for ailius in self.dict:
            self.scp(ailius, start, end)
    def __iter__(self):
        for element in self.dict:
            yield element
        
    def scp(self, tag, start_time, end_time):
        tags = self.dict[tag]
        print (tag)
        if "OP" in tags:          
            data = x_y_comparision(tag, tags["PV"]["tag"], tags["OP"]["tag"], tags["SP"]["tag"], start=start_time, end=end_time)
            if isinstance(data, pd.DataFrame) == False:
                return 
            # print(data[~data["SP"].isna()]["SP"])
            data["SP"] = pd.to_numeric(data["SP"], errors="coerce").ffill()
            data["OP"] = pd.to_numeric(data["OP"], errors="coerce").ffill()
            data["AI"] = data["SP"] - data["PV"]
            data["AI"] = data["AI"].abs()
            area_index = data["AI"].mean() #* (end_time - start_time).total_seconds() / 3600

            corr = round(data["PV"].corr(data["OP"]), 2)
            # print(data["PV"])
            tags["PV"]["mean"] = round(data["PV"].mean(),2)#.round(2)
            tags["PV"]["std"] = round(data["PV"].std(),2)#.round(2)

            tags["OP"]["mean"] = round(data["OP"].mean(),2)
            tags["OP"]["std"] = round(data["OP"].std(),2)
            tags["string"] = f"{tags['PV']['mean']} ± {(tags['PV']['std'] * 3).round(2)}"

            print(f"\tPV = {tags['PV']['mean']} ± {(tags['PV']['std'] * 3).round(2)}")
            print(f"\tOP = {tags['OP']['mean']} ± {(tags['OP']['std'] * 3).round(2)}")
            print(f"\tCorrelation = {corr}")
            # print(data["AI"].mean())
            print(f"\tArea index = {round(area_index,2)}")

        elif all(x in tags for x in ("PV", "SP")):
            data = SP_eval(tag, tags["PV"]["tag"], tags["SP"]["tag"], start=start_time, end=end_time)
            if isinstance(data, pd.DataFrame) == False:
                return
            data["SP"] = pd.to_numeric(data["SP"], errors="coerce").ffill()
            data["AI"] = data["SP"] - data["PV"]
            data["AI"] = data["AI"].abs()
            area_index = data["AI"].mean() #* (end_time - start_time).total_seconds() / 3600
        
            tags["PV"]["mean"] = round(data["PV"].mean(), 2)
            tags["PV"]["std"] = round(data["PV"].std(),2)
            tags["string"] = f"{tags['PV']['mean']} ± {(tags['PV']['std'] * 3).round(2)}"
            
            print(f"\tPV = {tags["PV"]["mean"]} ± {tags["PV"]["std"] * 3}")
            # print(data["AI"].mean())
            print(f"\tArea index = {round(area_index,2)}")
        
        elif len(tags) == 2 and "PV" in tags:
            tags["PV"] ["mean"]= round(Tag_query(tags["PV"]["tag"], 'AVG', start=start_time, end=end_time,)['VALUE'][0],2)
            tags["PV"]["std"] = round(Tag_query(tags["PV"]["tag"], 'STDDEV', start=start_time, end=end_time,)['VALUE'][0],2)
            tags["string"] = f"{tags['PV']['mean']} ± {(tags['PV']['std'] * 3).round(2)}"
            print(f"\tPV = {tags["PV"]["mean"]} ± {(tags["PV"]["std"] * 3).round(2)}")

        elif len(tags) == 2 and "SP" in tags:
            tags["PV"] = {}
            tags["PV"]["mean"]= round( Tag_query(tags["SP"]["tag"], 'AVG', start=start_time, end=end_time,)["VALUE"][0],2)
            tags["PV"]["std"] = round(Tag_query(tags["SP"]["tag"], 'STDDEV', start=start_time, end=end_time,)["VALUE"][0],2)
            tags["string"] = f"{tags['PV']['mean']} ± {(tags['PV']['std'] * 3).round(2)}"
            print(f"\tPV = {tags["PV"]["mean"]} ± {(tags["PV"]["std"] * 3).round(2)}")

class Control():
    """
    Control objects structure the parameters required to perform control 
    calculations on Aspen Trends  
    """
    def __init__(self, key, string, mean, variance, target, USL, LSL, units):
        """
        Control objects need to be initialised with the following information:
        
        :param self: the object itself
        :param key: the key that indicates where the output of the control 
                    object should go this takes the form of 3-4 characters from 
                    the start of the trend description(F03 or L302 or FROM).
        :param string: this a string of mean +/- 3 sigma
        :param mean: the mean of the parameter being analysed
        :param variance: the variance of the parameter being analysed
        :param target: the control plan target for the parameter
        :param USL: the Upper Specification Limit for the parameter
        :param LSL: the Lower Specification Limit for the parameter
        :param units: the units of the parameter (eg, kg/hr, l/s)
        """
        self.key = key
        self.string = string
        self.mean = float(mean)
        self.var = float(variance)
        self.target = target
        self.usl = USL
        self.lsl = LSL
        self.units = units
        self.string = None
        self.pp = None
        self.ppu = None
        self.ppl = None
        self.ppk = None
        self.control_calc()
        self.stringify()
        self.values = None
        self.calculate()
        self.cell_format = None
        self.cell_format_key = None
        self.cell_colour()

    def control_calc(self):
        """
        Calculates the control parameters from the mean, the variance, and the 
        specification limits. If variance is 0 this will not generate control 
        values.
        """
        if None not in (self.usl, self.lsl):
            try:
                self.pp = (self.usl -self.lsl) / ( 6* self.var**0.5)
                self.ppu = (self.usl-self.mean) / ( 3* self.var**0.5)
                self.ppl = (self.mean - self.lsl) / ( 3* self.var** 0.5)
                self.ppk = min(self.ppu, self.ppl)
            except ZeroDivisionError:
                print(f"Variance of {self.key} is 0 meaning control parameters cannot be computed")
        else:
            print(f"{self.key} does not have specification limits... skipping this calculation")

    def stringify(self):
        """Generates the string used to present data"""
        self.string = f"{round(self.mean, 2)} ± {round(self.var, 2)} {self.units}"
    
    def calculate(self):
        """Generates a string without units"""
        self.values = f"{round(self.mean,2)} ± {round(self.var**0.5 * 3, 2)}"

    def cell_colour(self):
        """
        Determines the cell colour that should be used for this parameter 
        based on the control of the parameter. Green if the parameter is 
        controled with enough precision and centered on target. Orange if 
        the control is precise and not centered and Red if the parameter does 
        not have the required precision. 
        """
        if self.pp == None:
            self.cell_format = {"font_color":"black"}
            self.cell_format_key = '{"font_color":"black"}'

        elif self.ppk > 1.3 :
            self.cell_format = {"bg_color":"green", "font_color":"white"}
            self.cell_format_key = '{"bg_color":"green", "font_color":"white"}'
        elif self.pp > 1.3 :
            self.cell_format = {"bg_color":"orange", "font_color":"white"}
            self.cell_format_key = '{"bg_color":"orange", "font_color":"white"}'
        else:
            self.cell_format = {"bg_color":"red","font_color":"white"}
            self.cell_format_key = '{"bg_color":"red","font_color":"white"}'

def total_feed_rate(Control_dict):
    Total_feed_rate = 0
    Total_feed_rate_var = 0
    for i in ["W01", "W02", "W03"]:
        Total_feed_rate += Control_dict[i].mean
        Total_feed_rate_var += Control_dict[i].var
    string =  f"{round(Total_feed_rate,2)} ± {round((Total_feed_rate_var**0.5)*3, 2)}"   
    Control_dict["Total_feed"] = Control("Total_feed", string, Total_feed_rate, Total_feed_rate_var, None, None, None, "kg/hr")

def DEMA(x):
    """Polyfit to convert acid density to mass fraction (DEnsity(kg/m3 or g/l) to MAss fraction)"""
    return round(-527.057716 + 1.074723463 * x - 0.000801043 * x**2 + 2.95816E-07 * x**3 + (-4.2184E-11)* x ** 4, 4)/100

def Density_correlations(Control_dict, Control_params, recipe):
    
    D1 = 0.95502 * Control_dict["D01"].mean + 54.5339 # correlation updated using 2025 data from 1.0288 and -59.227
    D1_var = 0.95502 ** 2 * Control_dict["D01"].var
    Control_dict["D01_corr"] = Control("D01_corr", None, D1, D1_var, 
                                       Control_params[recipe]["D01_corr"]["x"], 
                                       Control_params[recipe]["D01_corr"]["USL"],
                                       Control_params[recipe]["D01_corr"]["LSL"], "kg/m3")

    D2 = Control_dict["D02"].mean * 0.826431 + 204.4444 #correlation updated using 2025 data from 0.8184 and 199.88
    D2_var = Control_dict["D02"].var * 0.826431 ** 2
    Control_dict["D02_corr"] = Control("D02_corr", None, D2, D2_var,
                                       None,
                                       None,
                                       None, "kg/m3")

def acid_ratio(Control_dict):
    D = Control_dict["D01_corr"].mean
    D_var = Control_dict["D01_corr"].var
    M = Control_dict["M01"].mean / 100
    M_var = Control_dict["M01"].var / 10000

    D_DEMA_var = D_var * ((DEMA(D+D_var**0.5) - DEMA(D-D_var**0.5))/(D_var **0.5)) ** 2 

    def kiln_acid_ratio(kiln):
        acid_key = str( 3 + kiln)
        kiln_key = str(kiln)
        ratio_key = "R0" + kiln_key
        W = Control_dict["W0"+ kiln_key].mean
        W_var = Control_dict["W0"+kiln_key].var
        F = Control_dict["F0"+ acid_key].mean / 1000 # L/hr to m3/hr
        F_var = Control_dict["F0"+acid_key].var / 1000000
        try:
            R = F * D * DEMA(D) / ( W * ( 1 - M ))
        except ZeroDivisionError:
            R = 0              
        # define partial derivatives for variance calculation
        def RparF(W, D=D,M=M):
            try:
                res = (D/(W*(1-M))) ** 2
            except ZeroDivisionError:
                res = 0
            return res
        
        def RparD(F, W, M=M):
            try:
                res = (F/(W*(1-M))) ** 2 
            except ZeroDivisionError:
                res = 0
            return res
        
        def RparW(F,W, D=D,M=M):
            try:
                res = (F*D/((1-M) * (W**2))) ** 2
            except ZeroDivisionError:
                res = 0
            return res
        
        def RparM(F, W, D=D,M=M):
            try:
                res = (F*D/(W*((1-M)**2))) ** 2
            except ZeroDivisionError:
                res = 0
            return res
        R_var = RparF(W) * F_var
        R_var += RparD(F,W) * D_DEMA_var
        R_var += RparW(F,W) * W_var 
        R_var += RparM(F,W) * M_var

        Control_dict[ratio_key] = Control(ratio_key, None, R, R_var, None, None, None, "ratio")

    for i in [1,2,3]:
        kiln_acid_ratio(i)

def acid_recovery(Control_dict):

    D02 = Control_dict["D02_corr"].mean # kg/m3
    D02_var = Control_dict["D02_corr"].var # kg/m3

    D2 = DEMA(D02)
    D2_var = (D02_var) * ((DEMA(D02 - D02_var ** 0.5 ) - DEMA(D02 + D02_var ** 0.5)) / (D02_var ** 0.5)) ** 2

    F09 = Control_dict["F09"].mean / 1000 # l/hr to m3/hr
    F09_var = Control_dict["F09"].var / 1000000

    F28 = Control_dict["F28"].mean / 1000 # m3/hr
    F28_var = Control_dict["F28"].var / 10000000 # m3/hr

    washed = F09 * D02 * D2 #kg/hr
    washed_var = (washed**2) * (F09_var/F09**2 + D02_var/D02**2 + D2_var/D2**2) #kg/hr

    bled = F28 * D02 * D2 #kg/hr
    bled_var = (bled**2) * (F28_var/F28**2 + D02_var/D02**2 + D2/D2_var**2) #kg/hr

    rec = washed - bled
    rec_var = washed_var + bled_var

    D01 = Control_dict ["D01_corr"].mean # kg/m3
    D01_var = Control_dict["D01_corr"].var # kg/m3

    D1 = DEMA(D01)
    D1_var = (D01_var**2) * ((DEMA(D01 - D01_var ** 0.5 ) - DEMA(D01 + D01_var ** 0.5)) / (D01_var ** 0.5)) ** 2
    
    F04 = Control_dict["F04"].mean / 1000 # m3/hr
    F04_var = Control_dict["F04"].var / 1000000
    F05 = Control_dict["F05"].mean / 1000
    F05_var = Control_dict["F05"].var / 1000000
    F06 = Control_dict["F06"].mean / 1000
    F06_var = Control_dict["F06"].var / 1000000

    flow = F04 + F05 + F06
    flow_var = F04_var + F05_var + F06_var

    used = (flow) * D01 * D1
    used_var = (used**2) * (flow_var / flow + D01_var/ D01 + D1_var/ D1)

    Acid_efficiency = rec / used
    Acid_efficiency_var = (Acid_efficiency**2) * (rec_var/rec**2 + used/used_var**2)

    Control_dict["Dry_Acid_recovery_efficieny"] = Control("Acid recovery_efficiency",
                                                       None, Acid_efficiency, 
                                                       Acid_efficiency_var, None,
                                                       None, None, "%")
    
    Control_dict["Dry_Acid_bled"] = Control("Dry_Acid_bled", None, bled, 
                                            bled_var, None, None, None, "kg/hr")
    
    Control_dict["Dry_Acid_eff"] = Control("Dry_Acid_eff", None, 
                                                Acid_efficiency, 
                                                Acid_efficiency_var, None, None,
                                                None, "")

def main():
    print("Current directory: " , os.getcwd())
    print(__file__)
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    with open("C-carbon_Aspen_tags.json") as f:
        Tags = json.load(f)
    with open("Recipe_parameters.json") as f:
        Params = json.load(f)
    
    # User input for report
    start_time = datetime.strptime(input("start (YY-mm-dd HH:MM): "), '%y-%m-%d %H:%M')
    end_time = datetime.strptime(input("end (YY-mm-dd HH:MM): "), '%y-%m-%d %H:%M')
    Recipe = str(input("Recipe? (1-6):"))
    acid_used = str(input("Fresh or int acid? (lower case fresh/int): "))
    neutralisation = str(input("Neutralisation grade?(A,G,N): "))
    percent_rate = int(input("Rate percent?(numbers only 1-100): "))/100

    # data = Tag_query("WSH1.TK2.PPHO.FC.01.PV", command="STDDEV", 
    #                  start=datetime(2023,1,1,7), end=datetime(2023,1,2,7))
    # print(data)

    Aspen_data = Tag_dict(Tags, start_time, end_time)
    Control_dicts = {}
    for i in Aspen_data:
        if i in ["F28", "F11 int", "F11 fresh"]:
            if i == "F11 int":
                j = "F11"
                Control_dicts[i] = Control(i, Aspen_data.dict[i]["string"], 
                                           Aspen_data.dict[i]["PV"]["mean"], 
                                           Aspen_data.dict[i]["PV"]["std"]**2, 
                                           Params[Recipe][j]["int"]["x"]*percent_rate, 
                                           Params[Recipe][j]["int"]["USL"]*percent_rate, 
                                           Params[Recipe][j]["int"]["LSL"]*percent_rate,
                                           Tags[i]["units"])
            elif i == "F11 fresh":
                j="F11"
                Control_dicts[i] = Control(i, Aspen_data.dict[i]["string"], 
                                           Aspen_data.dict[i]["PV"]["mean"], 
                                           Aspen_data.dict[i]["PV"]["std"]**2, 
                                           Params[Recipe][j]["fresh"]["x"]*percent_rate, 
                                           Params[Recipe][j]["fresh"]["USL"]*percent_rate, 
                                           Params[Recipe][j]["fresh"]["LSL"]*percent_rate,
                                           Tags[i]["units"])
            if i == "F28":
                j="F28"
                Control_dicts[i] = Control(i, Aspen_data.dict[i]["string"], 
                                           Aspen_data.dict[i]["PV"]["mean"], 
                                           Aspen_data.dict[i]["PV"]["std"]**2, 
                                           Params[Recipe][j][acid_used]["x"]*percent_rate, 
                                           Params[Recipe][j][acid_used]["USL"]*percent_rate, 
                                           Params[Recipe][j][acid_used]["LSL"]*percent_rate,
                                           Tags[i]["units"])
        
        elif i == "C01":
            Control_dicts[i] = Control(i, Aspen_data.dict[i]["string"],
                                       Aspen_data.dict[i]["PV"]["mean"],
                                       Aspen_data.dict[i]["PV"]["std"]**2,
                                       Params[Recipe][i][neutralisation]["x"],
                                       Params[Recipe][i][neutralisation]["USL"],
                                       Params[Recipe][i][neutralisation]["LSL"],
                                       Tags[i]["units"])
            
        elif i in ["T01", "T02", "T03", "T17", "T18", "T19", "V601", 
                   "V601 Demist", "D01", "P01", ]:
            # print(i)
            Control_dicts[i] = Control(i, Aspen_data.dict[i]["string"], 
                                       Aspen_data.dict[i]["PV"]["mean"], 
                                       Aspen_data.dict[i]["PV"]["std"]**2, 
                                       Params[Recipe][i]["x"], 
                                       Params[Recipe][i]["USL"], 
                                       Params[Recipe][i]["LSL"],
                                       Tags[i]["units"])
        
        elif i in Params[Recipe]:
            # print(i) 
            Control_dicts[i] = Control(i, Aspen_data.dict[i]["string"], 
                                       Aspen_data.dict[i]["PV"]["mean"], 
                                       Aspen_data.dict[i]["PV"]["std"]**2, 
                                       Params[Recipe][i]["x"]*percent_rate, 
                                       Params[Recipe][i]["USL"]*percent_rate, 
                                       Params[Recipe][i]["LSL"]*percent_rate,
                                       Tags[i]["units"])
        else:
            # print(i)
            Control_dicts[i] = Control(i, Aspen_data.dict[i]["string"], 
                                       Aspen_data.dict[i]["PV"]["mean"], 
                                       Aspen_data.dict[i]["PV"]["std"]**2,
                                       None, None, None,
                                       Tags[i]["units"])

    # Data Calculations
    total_feed_rate(Control_dicts)
    Density_correlations(Control_dicts, Params, Recipe)
    acid_ratio(Control_dicts)
    acid_recovery(Control_dicts)

    # writing formats to excel sheets
    workbook = xlw.Workbook("morning_report.xlsx")

    C_carbons_order = [
                        "Total_feed",
                        "W01","W02","W03", "Space_1", "F04", "F05", "F06", 
                        "Space_2", "Space_3",
                        "R01","R02","R03", "M01", # Mixing
                        "Space_4", "Space_5", "F03", "Space_6","T03","T19",
                        "Space_7", "Space_8", "F02", "Space_9", "T02","T18", 
                        "Space_10", "Space_11",
                        "F01", "Space_12", "T01", "T17","Space_13", # Kilning
                        "F16", "V601", "V601 Demist", "Space_14",
                        "Space_15", "D01", "D01_corr", "F11 fresh", "F11 int", 
                        "F28", "F09", "F07", "D02", "D02_corr",  
                        "Space_16", "F25", "F26", "F27", "F08","F30", "F31",
                        "C01","C02", "Space_17", # Washing and Acid control
                        "Space_18", "L302","T06", "T41B","P01", #Drying
                        "Space_19", "Space_20", "S51","S52","FROM", "P51", "F502" # Milling 
                        ]

    C_Carbons_Glossary = {
        "Total_feed":"Plant Feed rate",
        "W01":"W01 Wood to Kiln A","W02":"W02 Wood to Kiln B",
        "W03":"W03 Wood to Kiln C", "Space_1":None, 
        "F04":"F04 Acid to Kiln A", "F05":"F05 Acid to Kiln B", 
        "F06":"F06 Acid to Kiln C", "Space_2":None, 
        "Space_3":"Dry_Sawdust to Dry Acid Ratio",
        "R01":"R01 A Kiln Ratio","R02":"R02 B Kiln Ratio",
        "R03":"R03 C Kiln Ratio", "M01":"M01 Sawdust Moisture", # Mixing
        "Space_4":None, "Space_5":None, "F03":"F03 Gas Flow", # Kiln C 
        "Space_6":"T03 SP","T03":"T03 Activation Temp",
        "T19":"T19 Back box Temp", 
        "Space_7":None, "Space_8":None, "F02":"F02 Gas Flow", # Kiln B
        "Space_9":"T02 SP", "T02":"T02 Activation Temp",
        "T18":"T18 Backbox Temp", 
        "Space_10":None, "Space_11":None, "F01":"F01 Gas Flow",# Kiln A
        "Space_12":"T01 SP", "T01":"T01 Activation Temp",
        "T17":"T17 Backbox Temp","Space_13":None, 
        "F16":"F16 Abatement flow from Kilns", 
        "V601":"Kiln Scrubber Pressure drop", 
        "V601 Demist":"Demister PRessure drop", "Space_14":None, # Abatement 
        "Space_15":"D01 Target", "D01":"D01 Supply Acid Density", 
        "D01_corr":"D01 Estimated actual", 
        "F11 fresh":"F11 Fresh Acid addition", "F11 int":"F11 Intermediate Acid", "F28":"F28 Acid Bleed", 
        "F09":"F09 Recovered Acid", "F07":"F07 Digestor feed", 
        "D02":"D02 Recovered acid", "D02_corr":"D02 Estimated Actual", #Acid Control 
        "Space_16":None, "F25":"F25 Water Wash", 
        "F26":"F26 Water Wash", "F27":"F27 Water Wash", "F08":"F08 Total Water to Prayon",
        "F30":"F30 Additional Wash", "F31":"F31 Additional Wash",
        "C01":"C01 Caustic wash conductivity",
        "C02":"C02 Wash off conductivity", "Space_17":None, # Washing
        "Space_18":None, "L302":"L302 Dryer Feed", 
        "T06":"T06 Exhaust temperature", 
        "T41B":"T41 Bed temperature", "P01":"P01 Dryer Hood Pressure","Space_19":None, # Dryer 
        "Space_20":"Feeding rate", "S51":"S51 TK502 feed rate",
        "S52":"S52 TK503 feed rate","FROM":"S55 LAC feed rate", 
        "P51":"P51 Mill pressure", "F502":"F502 DeltP" #Milling
    }

    Cdf_grouping = {
        "Total_feed":"Mixing",
        "W01":"Mixing", "W02":"Mixing", "W03":"Mixing", "Space_1":"Mixing", 
        "F04":"Mixing", "F05":"Mixing", "F06":"Mixing", "Space_2":"Mixing", 
        "Space_3":"Mixing", "R01":"Mixing", "R02":"Mixing", "R03":"Mixing", 
        "M01":"Mixing","Space_4":"Mixing", #Mixing
        "Space_5":"KilnC", "F03":"KilnC", "Space_6":"KilnC","T03":"KilnC",
        "T19":"KilnC", "Space_7":"KilnC", #Kiln C
        "Space_8":"KilnB", "F02":"KilnB", "Space_9":"KilnB", "T02":"KilnB",
        "T18":"KilnB", "Space_10":"KilnB", # Kiln B 
        "Space_11":"KilnA", "F01":"KilnA", "Space_12":"KilnA", "T01":"KilnA", 
        "T17":"KilnA","Space_13":"KilnA", # Kiln A
        "F16":"Abatement", "V601":"Abatement", "V601 Demist":"Abatement", 
        "Space_14":"Abatement", #Abatement
        "Space_15":"Acid Mgmt", "D01":"Acid Mgmt", "D01_corr":"Acid Mgmt", 
        "F11 fresh":"Acid Mgmt", "F11 int":"Acid Mgmt","F28":"Acid Mgmt", "F09":"Acid Mgmt", 
        "F07":"Acid Mgmt", "D02":"Acid Mgmt", "D02_corr":"Acid Mgmt", 
        "Space_16":"Acid Mgmt", # Acid Mgmt
        "F25":"Washing", "F26":"Washing", "F27":"Washing", "F08":"Washing","F30":"Washing", 
        "F31":"Washing","C01":"Washing","C02":"Washing", "Space_17":"Washing", # Washing
        "Space_18":"Drying", "L302":"Drying","T06":"Drying", "T41B":"Drying","P01":"Drying",
        "Space_19":"Drying", #Drying
        "Space_20":"Milling", "S51":"Milling","S52":"Milling","FROM":"Milling",
        "P51":"Milling", "F502":"Milling" # Milling 

    }

    Cdf_index = [C_Carbons_Glossary[i] for i in C_carbons_order]

    Cdf_data = []
    for i in C_carbons_order:
        if "Space" in i:
            Cdf_data.append([None, None, None, None, None,
            None, None, Cdf_grouping[i]])
        else:
            Cdf_data.append(
                [C_Carbons_Glossary[i], Control_dicts[i].values, 
                Control_dicts[i].units, Control_dicts[i].usl, 
                Control_dicts[i].lsl, Control_dicts[i].pp, 
                Control_dicts[i].ppk, Cdf_grouping[i]])

    Cdf_columns = ["Name", "Value","Units", "USL", "LSL", "PP", "PPK", "Grouping",]
    Cdf = pd.DataFrame(Cdf_data,columns=Cdf_columns)

    Cdf["new group"] = Cdf["Grouping"].ne(Cdf["Grouping"].shift(1).bfill()).astype(int)
    
    mid_cells = Cdf.index[Cdf["new group"]== 0].to_list()
    bottom_cells = Cdf.index[Cdf["new group"]== 1].to_list()
    bottom_cells.append(69)
    mid_cells.append(68)    # final cells for the report format

    writer = pd.ExcelWriter("morning_report_v3.xlsx", engine='xlsxwriter')
    Cdf.to_excel(writer, index=True, sheet_name="report")
    wkbk = writer.book
    wksht = writer.sheets["report"]

    mid_format = {"bottom":1,"top":1,"left":2,"right":2}
    bottom_format = {"bottom":2,"top":1,"left":2,"right":2} # 

    mid = wkbk.add_format(mid_format)
    bottom = wkbk.add_format(bottom_format)
    
    for i in mid_cells:
        wksht.set_row(i,None,mid)
    for i in bottom_cells:
        wksht.set_row(i,None,bottom)

    Good = wkbk.add_format({"font_color":"#013301","bg_color":"#37fc37"})
    NotGood = wkbk.add_format({"font_color":"#333301","bg_color":"#fcfc37"})
    Bad = wkbk.add_format({"font_color":"#ffffff","bg_color":"#fc3737"})

    wksht.conditional_format('C2:C69',{'type':'formula',
                                       'criteria':'=AND(G2>=1.3,H2>=1.3)',
                                       'format':Good})
    wksht.conditional_format('C2:C69',{'type':'formula',
                                       'criteria':'=AND(G2>=1.3,H2<1.3)',
                                       'format':NotGood})
    wksht.conditional_format('C2:C69',{'type':'formula',
                                       'criteria':'=AND(G2<1.3,H2<1.3,H2<>"")',
                                       'format':Bad})

    formattedWorksheet = FormattedWorksheet(wksht, wkbk, Cdf, hasIndex=True)
    formattedWorksheet.format_cols(colFormatList=["Name","Value", "Units"], 
        colPatternFormatList={"border":True})    
    formattedWorksheet.format_add_separation_border_between_groups('Grouping')
    
    writer._save()

    os.system("start EXCEL.EXE morning_report_v3")


if __name__=="__main__":
    main()