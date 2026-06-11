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

class Production_day:
    def __init__(self, date):
        self.date = dt.datetime(year=date.year, month=date.month, day=date.day, hour=7, minute=0)
        self.up_times = None
        self.down_times = None
        self.production_periods = None 
        self.get_times()
        self.up_time_no = len(self.up_times)
        self.down_time_no = len(self.down_times)
        self.prune_uptimes()
        self.get_production_periods()
        self.document = None


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
    
    def fetch_tag_detailed(self, tag, start=None, duration=None, end=None, resolution=5):

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

        sql = "select TS, VALUE from HISTORY "\
            " where NAME='%s' and PERIOD=%s*10 and TS between TIMESTAMP'%s' and TIMESTAMP'%s'" %(tag, resolution ,start, end)

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
    
    def find_changes(self, df):
        df['CHANGE'] = df['VALUE'].ne(df['VALUE'].shift())
        result = []
        start_index = None
        for i, row in df.iterrows():
            if row['CHANGE'] == 1 and start_index is None:
                j = row['TS']
                m = row['VALUE']
                start_index = i
            elif row['CHANGE'] == 1 and start_index is not None:
                k = row['TS']
                result.append([m,j,k])
                j = row['TS']
                m = row['VALUE']
                start_index = i
            elif i == len(df.index)-1:
                k = row['TS']
                result.append([m,j,k])
        return result

    def get_times(self):
        kiln_C_drive_tag = "KLN1.KLN3.UWAC.XS.01.PV"
        drdata = self.fetch_tag_detailed(kiln_C_drive_tag, start = self.date, duration=dt.timedelta(days=1), resolution=5)
        changes = self.find_changes(drdata)
        str_down_times = []
        str_up_times = []
        for i in changes:
            if i[0] == 1:
                str_up_times.append(i[1:])
            else:
                str_down_times.append(i[1:])
        
        down_times = []
        for i in str_down_times:
            down_time = [dt.datetime.strptime(j, '%d-%b-%y %H:%M:%S.%f') for j in i]
            down_times.append(down_time)
        
        up_times = []
        for i in str_up_times:
            up_time = [dt.datetime.strptime(j, '%d-%b-%y %H:%M:%S.%f') for j in i]
            up_times.append(up_time)
        
        self.down_times = down_times
        self.up_times = up_times
        
        print(self.date.strftime("%Y-%m-%d"),len(down_times))

        return

    def get_production_periods(self):
        periods = []
        for i in self.up_times:
            Prod = Production_period(self.date, i[0], i[1])
            periods.append(Prod)
        self.production_periods = periods
        return

    def prune_uptimes(self):
        print(self.down_time_no)
        sig_periods = []
        for i in self.up_times:
            if abs((i[0]-i[1]).total_seconds()) > 300:
                sig_periods.append(i)
        
        self.up_times = sig_periods

    def ELCD_control(self):
        
        self.document = Document()
        sections = self.document.sections
        for section in sections:
            section.top_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
        
        self.document.add_heading('{} V901 I-MR control charts,\n{} plant downtime events'.format(self.date.strftime("%Y-%m-%d"), str(self.down_time_no)),0)
        
        print("len of production periods {}".format(len(self.production_periods)))

        for m , i in enumerate(self.production_periods):
            # print(m)
            for n, j in enumerate(i.production_runs):
                j.V_901_ctrl_chart()
                # print("number of production runs {}".format(len(i.production_runs)))
                if isinstance(j.V901_ctrl_figs, dict):
                    for k in j.V901_ctrl_figs:
                        self.document.add_picture(j.V901_ctrl_figs[k], width=Inches(7))
                        self.document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    # print("{} production run index".format(n))
                    if n < len(i.production_runs) - 1:
                        self.document.add_page_break()
            if m < len(self.production_periods) - 1:
                self.document.add_page_break()

        self.document.save("{}_ELCD_control_charts.docx".format(self.date.strftime("%Y-%m-%d")))
        
        plt.close()

class Production_period:
    def __init__(self, date, start, end):
        self.date = date
        self.starttime = start 
        self.endtime = end
        self.R_recipies = None
        self.main_recipies = None
        self.bento_recipies = None
        self.production_run_times = None
        self.production_runs = None
        self.get_recipe()
        self.get_setpoints()
        self.prune_production_runs()
        self.get_runs()
        # self.get_main_recipe()
        # self.get_bento_recipe()
        # self.reduce_finagling()
    
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
            data = self.fetch_tag_detailed(tag, start=start, end=end, resolution=50)
        # Pandas DataFrame with your data to 4 decimal places!
    
        data.name = tag
    
        return data

    def fetch_tag_detailed(self, tag, start=None, duration=None, end=None, resolution=5):

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
        end_str = end.strftime("%Y-%m-%d %H:%M:%S")
        start_str = start.strftime("%Y-%m-%d %H:%M:%S")

        # time period depends on the reporting from the instruments 

        sql = "select TS, VALUE from HISTORY "\
            " where NAME='%s' and PERIOD=%s*10 and TS between TIMESTAMP'%s' and TIMESTAMP'%s'" %(tag, resolution ,start_str, end_str)

        # data = pd.read_sql(sql,con) # Pandas DataFrame with your data!
    
        cursor = con.cursor()
        result = cursor.execute(sql)
        rows = result.fetchall()
        con.close()
        cols = []
        for i in result.description:
            cols.append(i[0])
    
        #print(len(rows))
        try:
            data = pd.DataFrame(data=np.array(rows), columns=cols)
        except ValueError:
            interval = end - start
            delT = interval/24
            new_start_times = [start + i * delT for i in range(24)]
            new_end_times = [start + (i+1)*delT for i in range(24)]
            new_intervals = [[new_start_times[i] ,new_end_times[i]] for i in range(len(new_start_times))]
            # print(new_intervals)
            dataframes = []
            for i in new_intervals:
                new_start_str = i[0].strftime("%Y-%m-%d %H:%M:%S")
                new_end_str = i[1].strftime("%Y-%m-%d %H:%M:%S")
                con = pyodbc.connect("Driver={AspenTech SQLplus};HOST=192.168.9.47;PORT=10014")
                cursor = con.cursor()
                sql = "select TS, VALUE from HISTORY"\
                    f" where NAME='{tag}' and PERIOD={resolution}*10 and TS between TIMESTAMP'{new_start_str}' and TIMESTAMP'{new_end_str}'"
                print(sql)
                result = cursor.execute(sql)
                rows = result.fetchall()
                con.close()
                print(rows)
                cols = []
                for j in result.description:
                    cols.append(j[0])
                try:
                    data = pd.DataFrame(data=np.array(rows), columns=cols)
                    dataframes.append(data)
                except ValueError:
                    continue
            data = pd.concat(dataframes, ignore_index=True)
        

        data['VALUE'] = data['VALUE'].astype(float)
        data['VALUE'] = data['VALUE'].round(4)
        
        # Pandas DataFrame with your data to 4 decimal places!
    
        data.name = tag
    
        return data

    def find_changes(self, df):
        df['CHANGE'] = df['VALUE'].ne(df['VALUE'].shift())
        result = []
        start_index = None
        for i, row in df.iterrows():
            if row['CHANGE'] == 1 and start_index is None:
                j = row['TS']
                m = row['VALUE']
                start_index = i
            elif row['CHANGE'] == 1 and start_index is not None:
                k = row['TS']
                result.append([m,j,k])
                j = row['TS']
                m = row['VALUE']
                start_index = i
            elif i == len(df.index)-1:
                k = row['TS']
                result.append([m,j,k])
        return result

    def get_times(self):
        kiln_C_drive_tag = "KLN1.KLN3.UWAC.XS.01.PV"
        drdata = self.fetch_tag_detailed(kiln_C_drive_tag, start = self.date, duration=dt.timedelta(days=1), resolution=5)
        changes = self.find_changes(drdata)
        # print(changes)
        str_down_times = []
        str_up_times = []
        for i in changes:
            if i[0] == 1:
                str_up_times.append(i[1:])
            else:
                str_down_times.append(i[1:])
        
        down_times = []
        for i in str_down_times:
            down_time = [dt.datetime.strptime(j, '%d-%b-%y %H:%M:%S.%f') for j in i]
            down_times.append(down_time)
        
        up_times = []
        for i in str_up_times:
            up_time = [dt.datetime.strptime(j, '%d-%b-%y %H:%M:%S.%f') for j in i]
            up_times.append(up_time)
        
        self.down_times = down_times
        self.up_times = up_times
        
        return

    def get_recipe_2(self):
        R_recipe_tag = "PEL1.MIXR1.PAS.XI.04.DCS"
        main_tag = "0.UTIL0.RECIPE.NUM"
        bento_tag = "PEL1.MIXR1.RECIPE"
        ELCD_acid_tag = "PEL1.MIXR1.PPHO.FQ.01.SP"
        R_plnt_df = self.fetch_tag_detailed(R_recipe_tag, start = self.starttime, end = self.endtime, resolution=1)
        main_df = self.fetch_tag_detailed(main_tag, start = self.starttime, end = self.endtime, resolution=1)
        bento_df = self.fetch_tag_detailed(bento_tag, start = self.starttime, end = self.endtime, resolution=1)
        R_plnt_df['VALUE'] = R_plnt_df['VALUE'].round(0)
        main_df['VALUE'] = main_df['VALUE'].round(0)
        bento_df['VALUE'] = bento_df['VALUE'].round(0) 
        R_plnt_df = R_plnt_df.rename(columns={'VALUE':'R_plant'})
        main_df = main_df.rename(columns={'VALUE':'Main'})
        bento_df =bento_df.rename(columns={'VALUE':'Bento'})
        recipe_df = pd.concat([main_df, R_plnt_df['R_plant'], bento_df['Bento']], axis=1)

        prev_df = recipe_df[['Main','R_plant','Bento']].shift(1)
        mask = ((recipe_df[['Main','R_plant','Bento']]!= prev_df).any(axis=1))
        changes = recipe_df[mask]
        changes = changes.reset_index(drop=True)
        # print(changes)
        changes['TS'] = pd.to_datetime(changes['TS'],format='%d-%b-%y %H:%M:%S.%f')
        
        
        runs = []
        for i , row in changes.iterrows():
            if i + 1 < len(changes):
                temp = [[row['Main'], row['R_plant'], row['Bento']], row['TS'].to_pydatetime(), changes.loc[i+1,'TS'].to_pydatetime()]
                runs.append(temp)
            elif i + 1 == len(changes):
                temp = [[row['Main'], row['R_plant'], row['Bento']], row['TS'].to_pydatetime(), self.endtime]
                runs.append(temp)
        # print
        self.production_run_times = runs
        return

    def get_recipe(self):
        R_recipe_tag = "PEL1.MIXR1.PAS.XI.04.DCS"
        main_tag = "0.UTIL0.RECIPE.NUM"
        bento_tag = "PEL1.MIXR1.RECIPE"


        R_plnt_df = self.fetch_tag_detailed(R_recipe_tag, start = self.starttime, end = self.endtime, resolution=1)
        main_df = self.fetch_tag_detailed(main_tag, start = self.starttime, end = self.endtime, resolution=1)
        bento_df = self.fetch_tag_detailed(bento_tag, start = self.starttime, end = self.endtime, resolution=1)
        R_plnt_df['VALUE'] = R_plnt_df['VALUE'].round(0)
        main_df['VALUE'] = main_df['VALUE'].round(0)
        bento_df['VALUE'] = bento_df['VALUE'].round(0) 
        R_plnt_df = R_plnt_df.rename(columns={'VALUE':'R_plant'})
        main_df = main_df.rename(columns={'VALUE':'Main'})
        bento_df =bento_df.rename(columns={'VALUE':'Bento'})
        recipe_df = pd.concat([main_df, R_plnt_df['R_plant'], bento_df['Bento']], axis=1)

        prev_df = recipe_df[['Main','R_plant','Bento']].shift(1)
        mask = ((recipe_df[['Main','R_plant','Bento']]!= prev_df).any(axis=1))
        changes = recipe_df[mask]
        # print(changes)
        changes = changes.reset_index(drop=True)
        changes['TS'] = pd.to_datetime(changes['TS'],format='%d-%b-%y %H:%M:%S.%f')
        
        runs = []
        for i , row in changes.iterrows():
            if i + 1 < len(changes):
                temp = [[row['Main'], row['R_plant'], row['Bento']], row['TS'].to_pydatetime(), changes.loc[i+1,'TS'].to_pydatetime()]
                runs.append(temp)
            elif i + 1 == len(changes):
                temp = [[row['Main'], row['R_plant'], row['Bento']], row['TS'].to_pydatetime(), self.endtime]
                runs.append(temp)
        
        self.production_run_times = runs
        
        return
    
    def get_setpoints(self):

        recipe_runs = []        

        def alan_filter():
            for i in self.production_run_times:
                try:
                    temp = self.fetch_tag_detailed("0.UTIL0.RECIPE.NUM", start= i[1],end = i[2])
                except ValueError:
                    delT = i[1]-i[2]
                    interval = delT / 24
                    starts = [i[1] + j * interval for j in range(24)]
                    ends = [i[1] + (j + 1) * interval for j in range(24)]
                    new_intervals = list(zip(starts, ends))
                    print(new_intervals)
                    for j, k in enumerate(new_intervals):
                        try:
                            temp2 = self.fetch_tag_detailed("",k[0],k[1])
                            end = k[1]
                        except ValueError:
                            if end:
                                i[2] = end
                                self.production_run_times.insert()

                            continue
            return

        def concat_sp_df(dict_of_tags, start, end):
            SP_df = {}

            for i in dict_of_tags:
                print(i)
                SP_df[i] = self.fetch_tag_detailed(dict_of_tags[i], start=start, end=end)
                SP_df[i] = SP_df[i].rename(columns={'VALUE':i})

            target_df_key = next(iter(dict_of_tags))
            target_df = SP_df[target_df_key]
            Inter = pd.DataFrame()

            for df_name, df in SP_df.items():
                if df_name != target_df_key:
                    Inter = pd.concat([Inter, df.iloc[:,-1]],axis=1)
            SPs = pd.concat([target_df ,Inter], axis=1)
            # print(self.starttime.strftime("%H:%M"))
            # print(SPs)
            return SPs

        def sp_changes(df, recipe):
            prev_values = df.iloc[:, 1:].shift(1)
            mask = ((df.iloc[:, 1:]!= prev_values).any(axis=1))
            # print(df)
            df['TS']= pd.to_datetime(df['TS'],format='%d-%b-%y %H:%M:%S.%f' )
            changed_rows = df[mask]
            changed_rows = changed_rows.reset_index(drop=True)
            # print(changed_rows)
            changed_rows['TS'] = pd.to_datetime(changed_rows['TS'],format='%d-%b-%y %H:%M:%S.%f')
            runs = []
            for i , row in changed_rows.iterrows():
                setpoints = [i for i in row[1:]]
                if i + 1 < len(changed_rows):
                    temp = [recipe+setpoints, row['TS'].to_pydatetime(), changed_rows.loc[i+1,'TS'].to_pydatetime()]
                    runs.append(temp)
                elif i + 1 == len(changed_rows) and (abs((row['TS'].to_pydatetime() - df['TS'].iloc[-1].to_pydatetime()).total_seconds()) > 300):
                    temp = [recipe+setpoints, row['TS'].to_pydatetime(), df['TS'].iloc[-1].to_pydatetime()]
                    runs.append(temp)
            
            return runs 


        ELCD_std ={"acid":"PEL1.MIXR1.PPHO.FQ.01.SP",
                   "osf":"PEL1.MIXR1.OSF.WC.01.SP" }
        
        ELCD_wf ={"acid":"PEL1.MIXR1.PPHO.FQ.01.SP",
                "osf":"PEL1.MIXR1.OSF.WC.01.SP",
                "wf":"PEL1.MIXR1.WF.WC.01.SP" }

        C_Carbons = {"f04":"KLN1.MIXR1.PPHO.FC.01.SP",
                     "f05":"KLN1.MIXR2.PPHO.FC.01.SP",
                     "f06":"KLN1.MIXR3.PPHO.FC.01.SP",
                     "w01":"KLN1.MIXR1.SAW.WC.01.SP" ,
                     "W02":"KLN1.MIXR2.SAW.WC.01.SP" ,
                     "W03":"KLN1.MIXR3.SAW.WC.01.SP" }
        
        # for i in self.production_run_times:
        #     print(i)
        # print("----")
        for i in self.production_run_times:

            # alan_filter()
            print(i)

            if  i[0][0]<= 6:
                SPs = concat_sp_df(C_Carbons, i[1],i[2])
                
            if i[0][0] > 6:
                if i[0][1] == 1:
                    SPs = concat_sp_df(ELCD_std, i[1],i[2])
                    
                elif i[0][1] == 2:
                    SPs = concat_sp_df(ELCD_wf, i[1], i[2])
            
            # try:
            runs = sp_changes(SPs, i[0])
            recipe_runs += runs
            
            # except: 
                # print("Unable to put SP data frames into runs of SPs")
        self.production_runs = recipe_runs
    
    def prune_production_runs(self):
        # for i in self.production_runs:
        #     print(i)
        # print("-------") 
        def is_within_margin(time1, time2, margin_mins):
            time_diff = time1 - time2
            margin = dt.timedelta(minutes=margin_mins)
            return time_diff <= margin
        # print(self.production_runs)
        filtered_runs = []
        for i in self.production_runs:
            if abs((i[2]-i[1]).total_seconds()) > 180:
                filtered_runs.append(i)
        
        concated_runs = []
        if len(filtered_runs)>1:
            current_m = filtered_runs[0]
            for i in range(1, len(filtered_runs)):
                current_e = filtered_runs[i]
                if current_e[0] == current_m[0] and is_within_margin(current_m[2], current_e[1], 180):
                    current_m[2] = current_e[2]
                else:
                    concated_runs.append(current_m)
                    current_m = current_e
        
            concated_runs.append(current_m)
            # for i in concated_runs:
                # print(i)

            self.production_runs = concated_runs
        
        return

    def get_runs(self):
        runs = []

        for i in self.production_runs:
            r = Production_run(date = self.date, start=i[1], end=i[2], recipe=i[0])
            runs.append(r)
        self.production_runs = runs
        return


class Production_run:
    def __init__(self, date, start, end, recipe):
        self.date = date
        self.starttime = start
        self. endtime = end
        self.recipe = recipe
        self.recipe_str = None
        self.get_recipe_str()
        self.conditions = None
        self.get_conditions()
        self.V901_ctrl_figs = None
        self.mixes_duration = None
        self.ELCD_mix_data = None
        # self.get_mixes_duration()
        # self.V_901_ctrl()
        # self.plot_IMR(self.conditions["F09"])
    
    def get_recipe_str(self):
        
        if len(self.recipe) == 5:
            self.recipe_str = "CNR115"
        elif len(self.recipe) == 6:
            if self.recipe[-1] < 70: 
                self.recipe_str = "CNR115LB"
            else:
                self.recipe_str = "CNR120"
        elif len(self.recipe) == 9:
            carbon = {
                1.0:"recipe2 2kilns",
                2.0:"recipe2",
                3.0:"recipe4 2kilns",
                4.0:"recipe4",
                5.0:"recipe6 2 kilns",
                6.0:"recipe6" 
                      }
            self.recipe_str = carbon[self.recipe[0]]

    def get_mixes_duration(self):

        def ctrl_data_filter(df, threshold):
            df.loc[df['VALUE']<threshold, 'VALUE'] = 0
            totals = df.loc[(df['VALUE']!=0) & (df['VALUE'].shift(1)==0)].copy()
            totals['MR'] = abs(totals['VALUE']-totals['VALUE'].shift(1))
            self.mixes_duration = len(totals)
            return totals
        
        ctrl = {}
        
        # print(self.starttime.strftime("%d %H:%M"), self.endtime.strftime("%d %H:%M"), self.recipe)
        if self.recipe[0] <= 6:
            print("there is no ELCD production during this period {} to {}".format(self.starttime.strftime("%Y-%m-%d %H:%M"), self.endtime.strftime("%Y-%b-%d %H:%M")))
            return        
        
        elif self.recipe[1] == 2:
            tags = ["osf_tot", "wf_tot", "acid_tot"]
            # print("has wf")
            for i in tags:
                if i == "wf_tot":
                    threshold = 20
                else:
                    threshold = 100
                temp = ctrl_data_filter(self.conditions[i], threshold=threshold)
        
        elif self.recipe[1] == 1:
            tags = ["osf_tot", "acid_tot"]
            # print("is std")
            for i in tags:
                temp = ctrl_data_filter(self.conditions[i], threshold=100)


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
    
        if len(rows)>= 1:
            data = pd.DataFrame(data=np.array(rows), columns=cols)
            data['VALUE'] = data['VALUE'].astype(float)
            data['VALUE'] = data['VALUE'].round(4)
            data['TS'] = pd.to_datetime(data['TS'],format='%d-%b-%y %H:%M:%S.%f')
        else:
            data = self.fetch_tag_detailed(tag=tag, start=start, end=end, resolution=1)
        
        # Pandas DataFrame with your data to 4 decimal places!
    
        data.name = start.strftime(" %Y-%m-%d %H:%M ")+ tag
    
        return data

    def fetch_tag_detailed(self, tag, start=None, duration=None, end=None, resolution=5):
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
        data['TS'] = pd.to_datetime(data['TS'],format='%d-%b-%y %H:%M:%S.%f')
        # Pandas DataFrame with your data to 4 decimal places!
    
        data.name = tag
    
        return data

    def get_conditions(self):
        KPVs = {
            "osf_tot":"PEL1.MIXR1.OSF.FQ.01.pv","wf_tot":"PEL1.MIXR1.WF.FQ.01.PV",
            "acid_tot":"PEL1.MIXR1.PPHO.FQ.01.PV", #ELCD mix
            "S93":"PEL1.PELL2.PAS.SC.01.PV","S94":"PEL1.PELL2.PAS.SC.01.PV", # pelletisers
            "T91":"PEL1.DRYB1.FG.TC.01.PV","T93":"PEL1.DRYB1.FG.TC.03.PV",
            "T94":"PEL1.DRYB1.FG.TC.04.PV","T95":"PEL1.DRYB1.FG.TC.05.PV",
            "T96":"PEL1.DRYB1.PAS.TI.01.PV","S94":"PEL1.DRYB1.PAS.SI.01.PV",
            "pres_z1":"PEL1.DRYB1.FG.PI.01.PV","F92":"PEL1.DRYB1.STM.FI.01.PV", # Preactivation 
            "W01":"KLN1.MIXR1.SAW.WC.01.PV", "W02":"KLN1.MIXR2.SAW.WC.01.PV", 
            "W03":"KLN1.MIXR3.SAW.WC.01.PV", "F04":"KLN1.MIXR1.PPHO.FC.01.PV",
            "F05":"KLN1.MIXR2.PPHO.FC.01.PV","F06":"KLN1.MIXR3.PPHO.FC.01.PV", # C-carbon Mix
            "T01":"KLN1.BED1.UWAC.TC.01.PV","T02":"KLN1.BED2.UWAC.TC.01.PV",
            "T03":"KLN1.BED3.UWAC.TC.01.PV","T17":"KLN1.VENT1.FG.TI.01.PV",
            "T18":"KLN1.VENT2.FG.TI.01.PV","T19":"KLN1.VENT3.FG.TI.01.PV",
            "F01":"KLN1.BURN1.NG.FI.01.PV","F02":"KLN1.BURN2.NG.FI.01.PV",
            "F03":"KLN1.BURN3.NG.FI.01.PV","F16":"KLN1.EXHA1.FG.FI.01.PV", # Activation 
            "F07":"WSH1.TK3.PPHO.FC.01.PV", "F09":"WSH1.TK2.PPHO.FC.01.PV",
            "F25":"WSH1.TK6.FW.FC.01.PV","F26":"WSH1.TK7.FW.FC.01.PV",
            "F27":"WSH1.TK7.FW.FC.02.PV","F30":"WSH1.TK8.DNAOH.FI.01.PV",
            "F31":"WSH1.TK8.DNAOH.FI.02.PV","P06":"WSH1.WSH1.FG.PI.02.PV", # Wash
            "C01":"WSH1.TK8.DNAOH.AC.01.PV","C02":"WSH1.WSH1.WW.AI.01.PV", # Wash performance
            "F28":"TK1.TK2.PPHO.FC.01.PV","F11_fresh":"TK1.TK1.FPHO.FC.01.PV",
            "F11_plant":"TK1.TK1.PHO.FC.01.PV","D01":"TK1.TK3.PPHO.DI.01.PV",
            "D02":"TK1.TK2.PPHO.DI.01.PV", # Acid recovery
            "L302":"DRYR1.FEED3.WAC.SC.02.SP","T06":"DRYR1.EXHA1.FG.TC.01.PV",
            "T41A":"DRYR1.BED1.GAC.TC.01.PV","T41B":"DRYR1.BED1.GAC.TC.02.PV", # Dryer
        }
        ELCD_KPVs = {
            "osf_tot":"PEL1.MIXR1.OSF.FQ.01.pv","wf_tot":"PEL1.MIXR1.WF.FQ.01.PV",
            "acid_tot":"PEL1.MIXR1.PPHO.FQ.01.PV", #ELCD mix
            "S93":"PEL1.PELL2.PAS.SC.01.PV","S94":"PEL1.PELL2.PAS.SC.01.PV", # pelletisers
            "T91":"PEL1.DRYB1.FG.TC.01.PV","T93":"PEL1.DRYB1.FG.TC.03.PV",
            "T94":"PEL1.DRYB1.FG.TC.04.PV","T95":"PEL1.DRYB1.FG.TC.05.PV",
            "T96":"PEL1.DRYB1.PAS.TI.01.PV","S94":"PEL1.DRYB1.PAS.SI.01.PV",
            "pres_z1":"PEL1.DRYB1.FG.PI.01.PV","F92":"PEL1.DRYB1.STM.FI.01.PV", # Preactivation 
            "T03":"KLN1.BED3.UWAC.TC.01.PV","T19":"KLN1.VENT3.FG.TI.01.PV",
            "F03":"KLN1.BURN3.NG.FI.01.PV","F16":"KLN1.EXHA1.FG.FI.01.PV", # Activation 
            "F07":"WSH1.TK3.PPHO.FC.01.PV", "F09":"WSH1.TK2.PPHO.FC.01.PV",
            "F25":"WSH1.TK6.FW.FC.01.PV","F26":"WSH1.TK7.FW.FC.01.PV",
            "F27":"WSH1.TK7.FW.FC.02.PV","F30":"WSH1.TK8.DNAOH.FI.01.PV",
            "F31":"WSH1.TK8.DNAOH.FI.02.PV","P06":"WSH1.WSH1.FG.PI.02.PV", # Wash
            "C01":"WSH1.TK8.DNAOH.AC.01.PV","C02":"WSH1.WSH1.WW.AI.01.PV", # Wash performance
            "F28":"TK1.TK2.PPHO.FC.01.PV","F11_fresh":"TK1.TK1.FPHO.FC.01.PV",
            "F11_plant":"TK1.TK1.PHO.FC.01.PV","D01":"TK1.TK3.PPHO.DI.01.PV",
            "D02":"TK1.TK2.PPHO.DI.01.PV", # Acid recovery
            "L302":"DRYR1.FEED3.WAC.SC.02.SP","T06":"DRYR1.EXHA1.FG.TC.01.PV",
            "T41A":"DRYR1.BED1.GAC.TC.01.PV","T41B":"DRYR1.BED1.GAC.TC.02.PV", # Dryer
        }
        Ccarbon_KPVs = {
            "W01":"KLN1.MIXR1.SAW.WC.01.PV", "W02":"KLN1.MIXR2.SAW.WC.01.PV", 
            "W03":"KLN1.MIXR3.SAW.WC.01.PV", "F04":"KLN1.MIXR1.PPHO.FC.01.PV",
            "F05":"KLN1.MIXR2.PPHO.FC.01.PV","F06":"KLN1.MIXR3.PPHO.FC.01.PV", # C-carbon Mix
            "T01":"KLN1.BED1.UWAC.TC.01.PV","T02":"KLN1.BED2.UWAC.TC.01.PV",
            "T03":"KLN1.BED3.UWAC.TC.01.PV","T17":"KLN1.VENT1.FG.TI.01.PV",
            "T18":"KLN1.VENT2.FG.TI.01.PV","T19":"KLN1.VENT3.FG.TI.01.PV",
            "F01":"KLN1.BURN1.NG.FI.01.PV","F02":"KLN1.BURN2.NG.FI.01.PV",
            "F03":"KLN1.BURN3.NG.FI.01.PV","F16":"KLN1.EXHA1.FG.FI.01.PV", # Activation 
            "F07":"WSH1.TK3.PPHO.FC.01.PV", "F09":"WSH1.TK2.PPHO.FC.01.PV",
            "F25":"WSH1.TK6.FW.FC.01.PV","F26":"WSH1.TK7.FW.FC.01.PV",
            "F27":"WSH1.TK7.FW.FC.02.PV","F30":"WSH1.TK8.DNAOH.FI.01.PV",
            "F31":"WSH1.TK8.DNAOH.FI.02.PV","P06":"WSH1.WSH1.FG.PI.02.PV", # Wash
            "C01":"WSH1.TK8.DNAOH.AC.01.PV","C02":"WSH1.WSH1.WW.AI.01.PV", # Wash performance
            "F28":"TK1.TK2.PPHO.FC.01.PV","F11_fresh":"TK1.TK1.FPHO.FC.01.PV",
            "F11_plant":"TK1.TK1.PHO.FC.01.PV","D01":"TK1.TK3.PPHO.DI.01.PV",
            "D02":"TK1.TK2.PPHO.DI.01.PV", # Acid recovery
            "L302":"DRYR1.FEED3.WAC.SC.02.SP","T06":"DRYR1.EXHA1.FG.TC.01.PV",
            "T41A":"DRYR1.BED1.GAC.TC.01.PV","T41B":"DRYR1.BED1.GAC.TC.02.PV", # Dryer
        }
        conditions = {}
               
        # print(self.starttime, self.endtime ,self.recipe)
        if self.recipe[0] <= 6:
            for i in Ccarbon_KPVs:
                # print(i)
                conditions[i] = self.fetch_tag(Ccarbon_KPVs[i],start=self.starttime,end=self.endtime)
                conditions[i].name = self.starttime.strftime(" %Y-%m-%d") + i 

        if self.recipe[0] > 6:
            for i in ELCD_KPVs:
                conditions[i] = self.fetch_tag(ELCD_KPVs[i],start=self.starttime, end=self.endtime)
                conditions[i].name = self.starttime.strftime(" %Y-%m-%d") + i 
        # print("complete")
        self.conditions = conditions

    def V_901_ctrl_chart(self):
        
        def ctrl_data_filter(df, threshold):
            df.loc[df['VALUE']<threshold, 'VALUE'] = 0
            totals = df.loc[(df['VALUE']!=0) & (df['VALUE'].shift(1)==0)].copy()
            totals['MR'] = abs(totals['VALUE']-totals['VALUE'].shift(1))
            self.mixes_duration = len(totals)
            return totals

        def plot_IMR(df, tag, SP):
            SMALL_SIZE = 6
            MEDIUM_SIZE = 8
            BIGGER_SIZE = 10

            plt.rc('font', size=SMALL_SIZE)          # controls default text sizes
            plt.rc('axes', titlesize=MEDIUM_SIZE)     # fontsize of the axes title
            plt.rc('axes', labelsize=MEDIUM_SIZE)    # fontsize of the x and y labels
            plt.rc('xtick', labelsize=SMALL_SIZE)    # fontsize of the tick labels
            plt.rc('ytick', labelsize=SMALL_SIZE)    # fontsize of the tick labels
            plt.rc('legend', fontsize=SMALL_SIZE)    # legend fontsize
            plt.rc('figure', titlesize=BIGGER_SIZE)  # fontsize of the figure title
            memfile = BytesIO()
            xbar = df['VALUE'].mean()
            mrbar = df['MR'].mean()
            UCL = xbar + mrbar * 2.66
            LCL = xbar - mrbar * 2.66
            rUCL = mrbar * 3.27
            sig = mrbar / 1.128
            sigma = df['VALUE'].std()
            
            if tag.split("_")[0] == "acid":     # acid specs
                if self.recipe[1] == 1:         # Standard specs
                    spec = {
                        "USL" : 490,
                        "LSL" : 450
                    }                
                elif self.recipe[1] == 2:       # CNR 115LB 120
                    spec = {
                        "USL" : 530,
                        "LSL" : 470
                        }
            elif tag.split("_")[0] == "wf":     # wf specs 
                if self.recipe[-1] > 70:        # CNR120
                    spec = {
                        "USL" : 90,
                        "LSL" : 100
                        }
                else: 
                    spec = {                    # CNR115LB
                        "USL" : 60,
                        "LSL" : 40
                        }
            elif tag.split("_")[0] == "osf":    # osf specs
                if self.recipe[1] == 1:         # standard 
                    spec = {
                        "USL" : 510,
                        "LSL" : 490
                        }
                elif self.recipe[1] == 2:
                    if self.recipe[-1] > 70:    # CNR120
                       spec = {
                            "USL": 310,
                            "LSL": 290
                            }
                    else:
                        spec = {                # CNR115LB
                            "USL":460,
                            "LSL":440
                        }

            Cp = round((spec["USL"] - spec["LSL"]) / (6 * sig), 2)
            Cpk = round(min([(spec["USL"] - xbar)/(3 * sig),(xbar - spec["LSL"])/(3 * sig)]), 2)
            Pp = round((spec["USL"]-spec["LSL"])/(6*sigma), 2)
            Ppk = round(min([(spec["USL"] - xbar)/(3*sigma),(xbar - spec["LSL"])/(3*sigma)]), 2)

            fig, ax = plt.subplots(nrows=2, sharex=True)
            fig.set_size_inches(7,2.4)
            ax1, ax2 = ax.flatten()
            ax1.plot(df['TS'], df['VALUE'], linestyle='-', marker='.',markersize=5, color='red',label='VALUE')
            
            if self.recipe[1]== 1:
                ax1.set_title("{} CNR115 {} addition, from {} to {}. {} mixes. Cp = {}, Cpk = {}, Pp = {}, Ppk = {}".format(self.starttime.strftime("%Y-%m-%d"), tag.split("_")[0],self.starttime.strftime("%H:%M"), self.endtime.strftime("%H:%M"), str(len(df)),Cp, Cpk, Pp, Ppk))
            elif self.recipe[1] == 2:
                if self.recipe[-1] > 70:
                    ax1.set_title("{} CNR120 {} addition, from {} to {}. {} mixes. Cp = {}, Cpk = {}, Pp = {}, Ppk = {}".format(self.starttime.strftime("%Y-%m-%d"), tag.split("_")[0],self.starttime.strftime("%H:%M"), self.endtime.strftime("%H:%M"), str(len(df)),Cp, Cpk, Pp, Ppk))
                else:
                    ax1.set_title("{} CNR115LB {} addition, from {} to {}. {} mixes. Cp = {}, Cpk = {}, Pp = {}, Ppk = {}".format(self.starttime.strftime("%Y-%m-%d"), tag.split("_")[0],self.starttime.strftime("%H:%M"), self.endtime.strftime("%H:%M"), str(len(df)),Cp, Cpk, Pp, Ppk))
            
            if tag.split("_")[0] != "acid":
                ax1.axhline(spec["USL"], color="lime", label="USL = {}".format(str(spec["USL"])))
                ax1.axhline(spec["LSL"], color="lime", label="USL = {}".format(str(spec["LSL"])))
            
            ax1.axhline(xbar, color='green', label='xbar = {}'.format(str(round(xbar, 2))))
            ax1.axhline(UCL, color='orange', linestyle='--', label='UCL = {}'.format(str(round(UCL, 2))))
            ax1.axhline(LCL, color='orange', linestyle='--', label='LCL = {}'.format(str(round(LCL, 2))))
            if isinstance(SP, float):
                ax1.axhline(SP, color='magenta', linestyle='dotted',label='SP = {}'.format(str(SP)))
            ax2.plot(df['TS'], df['MR'], color='red', marker='.',markersize=5,linestyle='-',label='MR')
            ax2.axhline(mrbar, color='green', label='MRbar = {}'.format(str(round(mrbar, 2))))
            ax2.axhline(rUCL, color='orange', label='MR UCL = {}'.format(str(round(rUCL, 2))))
            ax2.set_xlabel('Time')
            ax2.xaxis.set_major_formatter(mdates.DateFormatter("%d %H:%M"))
            ax1.set_ylabel('Value')
            ax2.set_ylabel('MR')
            
            plt.tight_layout()
            box1 = ax1.get_position()
            ax1.set_position([box1.x0, box1.y0, box1.width * 0.85, box1.height])
            box2 = ax2.get_position()
            ax2.set_position([box2.x0, box2.y0, box2.width * 0.85, box2.height])
            ax2.text(0.5,-0.1, "Cp = {}, Cpk = {}, Pp = {}, Ppk = {}".format(Cp, Cpk, Pp, Ppk),size=SMALL_SIZE, ha="center")

            ax1.legend(loc='center right', bbox_to_anchor=(1.18,0.5))
            ax2.legend(loc='center right', bbox_to_anchor=(1.18,0.5))
            # adjustFigAspect(fig, 2)
            # plt.show()
            # if self.recipe[1]== 1:
            #     plt.savefig("{}_CNR1115_ctrl_chart.png".format(self.starttime.strftime("%Y-%m-%d-%H")))
            # elif self.recipe[1] == 2:
            #     if self.recipe[-1] > 70:
            #         plt.savefig("{}_CNR120_ctrl_chart.png".format(self.starttime.strftime("%Y-%m-%d-%H")))
            #     else:
            #         plt.savefig("{}_CNR115LB_ctrl_chart.png".format(self.starttime.strftime("%Y-%m-%d-%H")))
            plt.savefig(memfile)

            return memfile
        
        ctrl = {}
        
        # print(self.starttime.strftime("%d %H:%M"), self.endtime.strftime("%d %H:%M"), self.recipe)
        if self.recipe[0] <= 6:
            print("there is no ELCD production during this period {} to {}".format(self.starttime.strftime("%Y-%m-%d %H:%M"), self.endtime.strftime("%Y-%b-%d %H:%M")))
            return        
        elif self.recipe[1] == 2:
            tags = ["osf_tot", "wf_tot", "acid_tot"]
            # print("has wf")
            for i in tags:
                if i == "wf_tot":
                    threshold = 20
                else:
                    threshold = 100
                temp = ctrl_data_filter(self.conditions[i], threshold=threshold)
                
                # for index, row in temp.iterrows():
                #     print( row["TS"], i," " , row["VALUE"])

                if i == "wf_tot":
                    ctrl[i] = plot_IMR(temp, i, self.recipe[5])
                elif i == "osf_tot":
                    ctrl[i] = plot_IMR(temp, i, self.recipe[4])
                elif i == "acid_tot":
                    ctrl[i] = plot_IMR(temp, i, self.recipe[3] )
            
            self.V901_ctrl_figs = ctrl
            
            return

        elif self.recipe[1] == 1:
            tags = ["osf_tot", "acid_tot"]
            # print("is std")
            for i in tags:
                temp = ctrl_data_filter(self.conditions[i], threshold=100)
                
                # for index, row in temp.iterrows():
                #     print( i," " ,row)
                
                if i == "osf_tot":
                    ctrl[i] = plot_IMR(temp, i, self.recipe[4])
                elif i == "acid_tot":
                    ctrl[i] = plot_IMR(temp, i, self.recipe[3] )
            
            self.V901_ctrl_figs = ctrl
            return         
        
        else:
            print("the recipe for ELCD is neither 1 nor 2")
            return

 #   def L204_ctrl_chart(self):        

for i in range(0,30,1):
    D = Production_day(date=dt.datetime(year=2023, month=8, day=20)+dt.timedelta(days=i))
    D.ELCD_control()
# print(D.production_periods[0].production_runs[0].recipe)

