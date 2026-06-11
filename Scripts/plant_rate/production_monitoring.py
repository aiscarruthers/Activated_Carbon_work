import pyodbc
import datetime as dt
import pandas as pd
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO
import openpyxl as xl
from scipy.signal import find_peaks

# path = "ELCD_mix_data.xlsx"
# book = xl.load_workbook(path)
# writer = pd.ExcelWriter("ELCD_mix_data.xlsx", engine = "openpyxl", mode='a')
# writer.sheets = {ws.title: ws for ws in book.worksheets}
mpl.rcParams.update({'figure.max_open_warning':0})


def fetch_tag(tag, start=None, duration=None, end=None):
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

    # time period depends on the reporting from the instruments 

    sql = f"select \"IP_TREND_TIME\" as TS, \"IP_TREND_VALUE\" as VALUE from \"{tag}\" "\
        f" where TS between TIMESTAMP'{start_str}' and TIMESTAMP'{end_str}'"

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
        con.close()

    else:
        data = fetch_tag_detailed(tag=tag, start=start, end=end, resolution=1)

    data.name = tag
    
    return data
    
def fetch_tag_detailed(tag, start=None, duration=None, end=None, resolution=5):
    """
    Takes a tag name, duration and end point and submits a sql query returning
    data with intervals prescribed by the user, default is set to every 5 
    seconds. It is important to note that the data returned is often 
    interpolated. This fact can become problematic when using this fucntion
    to collect data about binary systems (open/closed) recipe values. 

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

    # time period depends on the reporting from the instruments 

    sql = "select TS, VALUE from HISTORY "\
        f" where NAME='{tag}' and PERIOD={resolution}*10 and TS between TIMESTAMP'{start_str}' and TIMESTAMP'{end_str}'"

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

    con.close()
    
    return data
    
class Production_day:
    """
    A Production day is a way of organising the data collection from the plant 
    into days, with uptimes, downtimes and production periods.

    Attributes:

    Methods
    """

    def __init__(self, date):
        """
        Initialise a Production day object for a given date.

        Parameters:
        date (datetime object): the date of the production day

        Attributes:
        date (datetime object): the date of the production day
        up_times (list): list of [[start time, end time], ...] for the uptime periods
        down_times (list): list of [[start time, end time], ...] for the downtime periods
        production periods (list): list of production period objects
        moisture (float): moisture percentage for the OSF being used that day.
        up_time_no (int): number of uptimes
        down_time_no (int): number of downtimes in a day
        """
        self.date = dt.datetime(year=date.year, month=date.month, day=date.day, hour=7, minute=0)
        self.up_times = None
        self.down_times = None
        self.production_periods = None
        self.moisture = None
        # self.set_moisture()
        # self.get_times()
        # self.up_time_no = len(self.up_times)
        # self.down_time_no = len(self.down_times)
        # self.prune_uptimes(300)
        # self.get_production_periods()
        self.document = None
    
    def find_changes(self, df):
        """
        This method takes a dataframe and returns a list of changes in VALUE 
        Column of the data frame. in the form [VALUE, start of VALUE, end of VALUE]

        Parameters:
        df (Dataframe): a data frame with columns TS and VALUE, usually a data 
        frame of setpoints

        Returns:
        List: a list of the Values and their start and end times in 
        chronological order 
        """
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
        """
        This function uses the find changes fuction on the kiln C drive to 
        determine periods of up and down time in the main plant.

        Parameters:
        date (datetime object): this takes the day of the class it is within to
        collect data from kiln C

        Returns:
        self.uptimes: a list of times between which Kiln C is turning
        self.downtimes: a list of time between which Kiln C is not turning
        """
        kiln_C_drive_tag = "KLN1.KLN3.UWAC.XS.01.PV"
        drdata = fetch_tag_detailed(kiln_C_drive_tag, start = self.date, duration=dt.timedelta(days=1), resolution=5)
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
            down_time = [j.to_pydatetime() for j in i]
            down_times.append(down_time)
        
        up_times = []
        for i in str_up_times:
            up_time = [j.to_pydatetime() for j in i]
            up_times.append(up_time)
        
        self.down_times = down_times
        self.up_times = up_times
        
        # print(self.date.strftime("%Y-%m-%d"),len(down_times))

        return

    def set_moisture(self):

        self.moisture =  5 # float(input("What is the OSF moisture percentage for {}?: ".format(self.date.strftime("%Y-%m-%d"))))
    
    def get_production_periods(self):
        """
        This method loops through the self.uptimes list and creates Production
        periods using the start and endtimes in the list.

        Attributes:
        production periods (list): A list of production periods dated within 
        the production day
        """
        periods = []
        # print(self.up_times)
        for i in self.up_times:
            Prod = Production_period(date=self.date, start=i[0], end=i[1], moisture=self.moisture)
            periods.append(Prod)
        # print(periods)
        self.production_periods = periods
        return

    def prune_uptimes(self, cutoff):
        """
        This method looks through the up_time list and removes any elements 
        with an interval shorter than the cut off.

        Parameters:
        cutoff (float): the number of seconds that an uptime period needs to
        have for it to not be pruned. 

        Attributes:
        up_times (list): A now pruned list of uptimes longer than the cut off 
        time
        """
        # print(self.down_time_no)
        sig_periods = []
        for i in self.up_times:
            if abs((i[0]-i[1]).total_seconds()) > cutoff:
                sig_periods.append(i)
        
        self.up_times = sig_periods

    def print_ELCD_mixes(self):
        
        osf_tot_tag = "PEL1.MIXR1.OSF.FQ.01.PV"
        acid_tot_tag = "PEL1.MIXR1.PPHO.FQ.01.PV"
        wf_tot_tag = "PEL1.MIXR1.WF.FQ.01.PV"
        d01_tag ="TK1.TK3.PPHO.DI.01.PV"
        V901_tag = "PEL1.MIXR1.PAS.WI.01.PV"
        TK902_tag = "PEL1.SILO2.OSF.WI.01.PV"

        starttime = self.date
        endtime = self.date + dt.timedelta(days=1)

        osf = fetch_tag(osf_tot_tag, start=starttime, end=endtime)
        acid = fetch_tag(acid_tot_tag, start=starttime, end=endtime)
        wf= fetch_tag(wf_tot_tag, start=starttime, end=endtime)
        d01 = fetch_tag_detailed(d01_tag, start=starttime, end=endtime, resolution=10)
        V901 = fetch_tag_detailed(V901_tag,start=starttime, end=endtime, resolution=10)
        TK902 = fetch_tag_detailed(TK902_tag,start=starttime, end=endtime, resolution=10)

        def mix_filter(df, threshold):
            # set any value below threshold in data frame to 0
            df.loc[df['VALUE']<threshold, 'VALUE'] = 0
            # create a new dataframe only including rows where Value is not 0 and the next value is 0.
            totals = df.loc[(df['VALUE']!=0) & (df['VALUE'].shift(1)==0)].copy()
            # Create a moving range column for the new data frame
            totals['MR'] = abs(totals['VALUE']-totals['VALUE'].shift(1))
            return totals
        
        def v901_filter(df, threshold):
            nf = df.copy()
            nf['VALUE'] = nf['VALUE'].apply(lambda x: 0 if x < threshold else x)
            peaks, _ = find_peaks(nf['VALUE'], distance= 60)
            maxima_df = df.loc[peaks].copy()
            return maxima_df
    
        def v901_trough(df, threshold):
            nf = df.copy()
            # nf['VALUE'] = nf['VALUE'].apply(lambda x: 0 if x > threshold else x)
            # print(nf)
            invert = -nf['VALUE']
            # print(invert)
            peaks, _ = find_peaks(invert, distance=60)
            # print(peaks)
            minima_df = df.loc[peaks].copy()
            return minima_df

        osf_tot = mix_filter(osf, 50)
        acid_tot = mix_filter(acid, 50)
        wf_tot = mix_filter(wf, 20)
        
        V901_res = v901_filter(V901,1000)
        V901_min = v901_trough(V901, 600)

        TK902_max = v901_filter(TK902, 1100)
        TK902_min = v901_trough(TK902, 1100)

        TK902_min.reset_index
        TK902_max.reset_index

        V901_min.reset_index
        V901_res.reset_index
        osf_tot.reset_index
        acid_tot.reset_index
        wf_tot.reset_index

        osf_tot = osf_tot.rename(columns={"VALUE":"OSF"})
        acid_tot = acid_tot.rename(columns={"VALUE":"ACID"})
        wf_tot = wf_tot.rename(columns={"VALUE":"WF"})
        d01 = d01.rename(columns={"VALUE":"D01"})
        
        V901_res = V901_res.rename(columns={"VALUE":"Max V901 weight"})
        V901_min = V901_min.rename(columns={"VALUE":"Min V901 weight"})

        TK902_max = TK902_max.rename(columns={"VALUE":"Max TK902 weight"})
        TK902_min = TK902_min.rename(columns={"VALUE":"Min TK902 weight"})

        TK902_max = TK902_max.sort_values(by="TS", ascending=True)
        TK902_min = TK902_min.sort_values(by="TS", ascending=True)

        # print(V901_min)

        V901_res = V901_res.sort_values(by="TS", ascending=True)
        osf_tot = osf_tot.sort_values(by="TS", ascending=True)
        acid_tot = acid_tot.sort_values(by="TS", ascending=True)
        wf_tot = wf_tot.sort_values(by="TS", ascending=True)       
        d01 = d01.sort_values(by="TS",ascending=True)

        newdf = pd.merge_asof(osf_tot,acid_tot, on='TS', direction='nearest')
        newdf = pd.merge_asof(newdf, V901_res,on='TS', direction='forward')
        newdf = pd.merge_asof(newdf, V901_min, on="TS", direction="backward")

        if len(wf_tot)> 1:
            newdf = pd.merge_asof(newdf, wf_tot, on='TS', direction='nearest', tolerance=pd.Timedelta(value=10, unit="min"))

        newdf = pd.merge_asof(newdf, TK902_max, on="TS", direction="backward" )
        newdf = pd.merge_asof(newdf, TK902_min, on="TS", direction="backward" )

        # print(newdf)
        merged = pd.merge_asof(newdf, d01, on='TS', direction='nearest')
        print(merged)
        merged.to_csv(f"Mix_data/{self.date.strftime('%Y-%m-%d')}_mix_data.csv")

        # print(merged["OSF"].sum())

        # print(osf_tot)
        # print(acid_tot)
        # print(wf_tot)

    def ELCD_control(self):
        """
        This method will generate a set of control charts for the day for a set
        of production runs within production periods within this production day.
        """
        # It might be a better design desision to create the control charts at 
        # the level of the production day rather than at the level of the 
        # production run. But this will need to be completed later. 
        
        # initialise word document
        self.document = Document()
        
        # format document
        sections = self.document.sections
        for section in sections:
            section.top_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
        
        # Document Title
        self.document.add_heading('{} V901 I-MR control charts,\n{} plant downtime events'.format(self.date.strftime("%Y-%m-%d"), str(self.down_time_no)),0)
        ELCD_check = 0
        for i in self.production_periods:
            for j in i.production_runs:
                if "CNR" in j.recipe_str:
                    ELCD_check +=1
                
        if ELCD_check == 0:
            print("Not making any ELCD Charts since there is no ELCD production")
            return

        # Loop through production runs in production periods to generate and 
        # insert the control charts.
        for m , i in enumerate(self.production_periods):
            # print(m)
            
            # Generate the chart for each production run
            for n, j in enumerate(i.production_runs):
                j.V_901_ctrl_chart()
                # print("number of production runs {}".format(len(i.production_runs)))

                # Check that the attribute V901_ctrl_figs is a dictionary and not None.
                if isinstance(j.V901_ctrl_figs, dict):
                    for k in j.V901_ctrl_figs:

                        # Check if the figure in the dictionary is not None type
                        if not isinstance(j.V901_ctrl_figs[k], type(None)):
                            self.document.add_picture(j.V901_ctrl_figs[k], width=Inches(7))
                            self.document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    # print("{} production run index".format(n))

                    # Make a new page for each production run.
                    if n < len(i.production_runs) - 1:
                        self.document.add_page_break()
                # Close any matplot objects to save memory
                plt.close()
            # Make a new page for each production period
            if m < len(self.production_periods) - 1:
                self.document.add_page_break()

        # save the docomuent 
        self.document.save("{}_ELCD_control_charts.docx".format(self.date.strftime("%Y-%m-%d")))
    
    def Wood_control(self):
        """
        This method will generate control charts for the Acid mixing in the 
        C-carbon process.
        """

        # initialise word document
        self.document = Document()
        
        # format the document
        sections = self.document.sections
        for section in sections:
            section.top_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
        
        # Document Title
        self.document.add_heading('{} L205 20min Moving Average - MR control charts,\n{} plant downtime events'.format(self.date.strftime("%Y-%m-%d"), str(self.down_time_no)),0)
        
        # begin looping through the production runs in the production periods 
        # and generating and inserting the control chart plots into the word 
        # Document
        for m , i in enumerate(self.production_periods):
            for n, j in enumerate(i.production_runs):
                # Create a heading and charts for each production run
                self.document.add_heading("{} to {}".format(j.starttime.strftime("%d %H:%M"), j.endtime.strftime("%d %H:%M"),2))
                j.L205_ctrl_chart()
                # print("number of production runs {}".format(len(i.production_runs)))
                
                # Check if the attribute L205_ctrl_figs is a dictionary and not a None type
                if isinstance(j.L205_ctrl_figs, dict):
                    
                    for k in j.L205_ctrl_figs:
                        # Check if each fig exist before adding it to the word doc
                        if not isinstance(j.L205_ctrl_figs[k], type(None)):
                            self.document.add_picture(j.L205_ctrl_figs[k], width=Inches(7))
                            self.document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    # print("{} production run index".format(n))
            
                    # Create a new page for each production run
                    if n < len(i.production_runs) - 1:
                        self.document.add_page_break()
            
                # Close any left over plots to save memory
                plt.close()
            
            # Create new page for a new production period
            if m < len(self.production_periods) - 1:
                self.document.add_page_break()

        # Save the Document
        self.document.save("{}_Wood_control_charts.docx".format(self.date.strftime("%Y-%m-%d")))
        
        # Close any left over plots to save memory
        plt.close()

class Production_period:
    """
    Production periods are periods of uptime within a production day, there 
    can be many recipe changes within a production periods, but each recipe 
    setting is given it's own production run.

    Attributes:

    Methods:

    """

    def __init__(self, date, start, end, moisture):
        """
        This method initialises a Production period

        Parameters:
        date (Datetime): the date of the production period
        start (Datetime): the start of the production period to the second
        end (Datetime): the end of the production period to the second
        moisture (float): the estimated moisture content of the OSF

        Atributes:
        date (Datetime):
        starttime (Datetime):
        endtime (Datetime):
        R_recipies
        main_recipes
        bento_recipes
        production_run_times (List):
        production_runs (List): A list of the production run objects representing
        periods where there are no setpoint or recipe changes.
        moisture (float): Estimation of the moisture percentage of OSF
        """
        self.date = date
        self.starttime = start 
        self.endtime = end
        self.production_run_times = None
        self.production_runs = None
        self.moisture = moisture
        # print(self.starttime, self.endtime)
        self.get_recipe()
        self.get_setpoints()
        self.prune_production_runs()
        self.get_runs()
    
    def get_recipe(self):
        """
        This method uses the start and end times uptimes in the production day 
        to collect the recipes being run during a particular day. This method 
        collects the different recipe settings from the main plant and 
        organsies them into periods where no changes take place collecting 
        them into a list.

        Attributes:
        production_run_times (List): The is is a list of the form
        [[[Mainplant recipe, R_plant recipe, Bentonorit recipe], starttime, endtime],...]
        and is then used to create production runs. 
        """

        # Recipe tags
        R_recipe_tag = "PEL1.MIXR1.PAS.XI.04.DCS"
        main_tag = "0.UTIL0.RECIPE.NUM"
        bento_tag = "PEL1.MIXR1.RECIPE"
        
        # Generate Dataframes, we need to use the detailed form because these 
        # tags only update every 30 mins or when there is a change in value 
        R_plnt_df = fetch_tag_detailed(R_recipe_tag, start = self.starttime, end = self.endtime, resolution=1)
        main_df = fetch_tag_detailed(main_tag, start = self.starttime, end = self.endtime, resolution=1)
        bento_df = fetch_tag_detailed(bento_tag, start = self.starttime, end = self.endtime, resolution=1)
        
        # Values in these data frames need to be rounded because the fetch tag 
        # detailed interpolates the values between changes, rounding to 0 
        # decimal places is a simple enough an way of reversing the interpolation
        R_plnt_df['VALUE'] = R_plnt_df['VALUE'].round(0)
        main_df['VALUE'] = main_df['VALUE'].round(0)
        bento_df['VALUE'] = bento_df['VALUE'].round(0) 
        
        # Rename the columns of these data frames and collect them into a single dataframe
        R_plnt_df = R_plnt_df.rename(columns={'VALUE':'R_plant'})
        main_df = main_df.rename(columns={'VALUE':'Main'})
        bento_df =bento_df.rename(columns={'VALUE':'Bento'})
        recipe_df = pd.concat([main_df, R_plnt_df['R_plant'], bento_df['Bento']], axis=1)
        
        # Create an new dataframe where all the Values in the recipe column are
        # shifted down 1, and compare this Dataframe to the previous to find any
        # changes in recipe to create the "changes" dataframe
        prev_df = recipe_df[['Main','R_plant','Bento']].shift(1)
        mask = ((recipe_df[['Main','R_plant','Bento']]!= prev_df).any(axis=1))
        changes = recipe_df[mask]
        changes = changes.reset_index(drop=True)
        changes['TS'] = pd.to_datetime(changes['TS'],format='%d-%b-%y %H:%M:%S.%f')
        
        # Reformat the changes into the [[[main, R_plant, bento] start, end],...] format
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
        """
        This method collects the setpoints for a given recipe setting and 
        collectes them under the production runs attribute based on the 
        production run times attribute. 

        Returns:
        None
        """
        recipe_runs = []        


        def concat_sp_df(dict_of_tags, start, end):
            """
            Function takes a dictionary of setpoint tags and concatenates 
            them into a single dataframe

            Parameters:
            dict_of_tags (dict): dictionary of the important SPs
            start (datetime):
            end (datetime):

            Returns:
            SPs (dataframe): this is a data frame of the setpoints for a given process
            """
            SP_df = {}
            
            # For each tag in the dictionary of tags collect a data frame for the SP changes
            for i in dict_of_tags:
                
                # Set points if they have not been changed will only update 
                # after a period of hours. Therefore we collect SP data going 
                # back much further than the period we are interested in. 
                df = fetch_tag(dict_of_tags[i], start=start-dt.timedelta(days=1), end=end)
                
                # Convert the time data into datetime values so that we can 
                # perform some maths to them  
                df["TS"] = pd.to_datetime(df["TS"])
                
                # find the last value in the data frame
                last_index = df["TS"].idxmax()
                last_values = df.loc[last_index]
                
                # find the latest value before the start time
                closest_index = (df['TS'] < start).idxmax()
                closest_values = df.loc[closest_index]
                
                # We slice the dataframe so that it is only as large as the 
                # period of interest making the first row the SP it begins 
                # with and last row the final setpoint and the end time 
                sliced_df = df[(df["TS"] >= start)]
                new_row = pd.DataFrame({"TS":[start],"VALUE":[closest_values["VALUE"]]})
                last_row = pd.DataFrame({"TS":[end],"VALUE":[last_values["VALUE"]]})
                res_df = pd.concat([new_row, sliced_df, last_row], axis=0, ignore_index=True)

                # The dataframe is then sorted by time and added to the 
                # dictionary SP_df and the Value columns renamed for the tag
                res_df = res_df.sort_values(by="TS")
                SP_df[i] = res_df
                SP_df[i] = SP_df[i].rename(columns={'VALUE':i})
            
            # Pulls out the first dataframe key and dataframe (Not sure why I 
            # have written these 2 blocks the way I have, it is worth checking to 
            # see if this can be condensed into a single for loop)
            target_df_key = next(iter(dict_of_tags))
            target_df = SP_df[target_df_key]
            Inter = pd.DataFrame()

            # Concats the following dataframes into the Inter data frame
            for df_name, df in SP_df.items():
                if df_name != target_df_key:
                    Inter = pd.concat([Inter, df.iloc[:,-1]],axis=1)

            # Concats the first and the second data frame into the SPs which 
            # will be returned
            SPs = pd.concat([target_df ,Inter], axis=1)
            
            return SPs

        def sp_changes(df, recipe):
            """
            This function takes a dataframe of setpoints and a recipe settings
            to create a list of recipe and setpoints settings with their start 
            and end times.

            Parameters:
            df (Dataframe): A Dataframe of the Setpoints during a production period
            recipe (List): A list of the form [[main, R_plant,Bento],...]

            Returns:
            runs (list): A list of the form recipe [[[recipe values, setpoint values], start, end],...]
            """
            # Create a mask using a copy of the dataframe shifting the values by 1
            # to find any setpoint changes
            prev_values = df.iloc[:, 1:].shift(1)
            mask = ((df.iloc[:, 1:]!= prev_values).any(axis=1))
            df['TS']= pd.to_datetime(df['TS'],format='%d-%b-%y %H:%M:%S.%f' )
            changed_rows = df[mask]
            changed_rows = changed_rows.reset_index(drop=True)
            changed_rows['TS'] = pd.to_datetime(changed_rows['TS'],format='%d-%b-%y %H:%M:%S.%f')

            # Iterate through the dataframe of setpoint changes and create a 
            # list of runs appending the setpoint values to the recipe settings

            runs = []
            for i , row in changed_rows.iterrows():
                setpoints = {}
                for j in changed_rows.columns[1:]:
                    setpoints[j] = row[j]

                if i + 1 < len(changed_rows):  
                    rec = recipe.copy()
                    rec.append(setpoints)
                    temp = [rec, row['TS'].to_pydatetime(), changed_rows.loc[i+1,'TS'].to_pydatetime()]
                    runs.append(temp)

                elif i + 1 == len(changed_rows) and (abs((row['TS'].to_pydatetime() - df['TS'].iloc[-1].to_pydatetime()).total_seconds()) > 300):
                    rec = recipe.copy()
                    rec.append(setpoints)
                    temp = [rec, row['TS'].to_pydatetime(), df['TS'].iloc[-1].to_pydatetime()]
                    runs.append(temp)
            
            return runs 

        ELCD_std ={"acid_tot":"PEL1.MIXR1.PPHO.FQ.01.SP",
                   "osf_tot":"PEL1.MIXR1.OSF.WC.01.SP" }
        
        ELCD_wf ={"acid_tot":"PEL1.MIXR1.PPHO.FQ.01.SP",
                "osf_tot":"PEL1.MIXR1.OSF.WC.01.SP",
                "wf_tot":"PEL1.MIXR1.WF.WC.01.SP"}

        C_Carbons = {"F04":"KLN1.MIXR1.PPHO.FC.01.SP",
                     "F05":"KLN1.MIXR2.PPHO.FC.01.SP",
                     "F06":"KLN1.MIXR3.PPHO.FC.01.SP",
                     "W01":"KLN1.MIXR1.SAW.WC.01.SP" ,
                     "W02":"KLN1.MIXR2.SAW.WC.01.SP" ,
                     "W03":"KLN1.MIXR3.SAW.WC.01.SP" }

        # Collect the important setpoints for a given main plant recipe setting 
        for i in self.production_run_times:
            print(i)
            if  i[0][0]<= 6:
                SPs = concat_sp_df(C_Carbons, i[1],i[2])
                
            if i[0][0] > 6:
                if i[0][1] == 1:
                    SPs = concat_sp_df(ELCD_std, i[1],i[2])
                    
                elif i[0][1] == 2:
                    SPs = concat_sp_df(ELCD_wf, i[1], i[2])
            
            # Find the setpoint changes during the production period and add 
            # them to the runs list
            print(SPs)
            runs = sp_changes(SPs, i[0])
            recipe_runs += runs

        # Set the production runs to the recipe runs found   
        self.production_runs = recipe_runs
    
    def prune_production_runs(self):
        """
        This method is used to prune the list of production runs and remove 
        periods where there is no substancial amount of time spent in a 
        particular set point. This is neccessary as operators do not always 
        change to the setpoint they intend and can make changes in quick 
        succession. 

        Returns:
        None
        """
        def is_within_margin(time1, time2, margin_mins):
            """
            This function checks if 2 time values are within a certain margin
            
            Parameters:
            time1 (datetime): start time of an interval
            time2 (datetime): end time of an interval
            margin_mins (float): margin required to return True

            Returns:
            Boolean : True if the time intervals are greater than the margin in mins else false
            """
            
            time_diff = time1 - time2
            margin = dt.timedelta(minutes=margin_mins)
            return time_diff <= margin
        
        # Create a new list of runs and add the runs greater than 3 mins
        filtered_runs = []
        for i in self.production_runs:
            if abs((i[2]-i[1]).total_seconds()) > 180:
                filtered_runs.append(i)
        
        # if none are added end the routine
        if len(filtered_runs) == 0:
            
            return

        # Create a new list of runs and add the first element of the filtered runs
        concated_runs = []
        current_m = filtered_runs[0]
        
        # if the margin between runs is bellow the limit and the setpoints are all the same 
        # collapse these 2 runs into one.  Else go on to the next run.
        for i in range(1, len(filtered_runs)):
            current_e = filtered_runs[i]
            if current_e[0] == current_m[0] and is_within_margin(current_m[2], current_e[1], 30):
                current_m[2] = current_e[2]
            else:
                concated_runs.append(current_m)
                current_m = current_e
        # add the last run
        concated_runs.append(current_m)
        
        # Update the class
        self.production_runs = concated_runs
        
        return

    def get_runs(self):
        """
        This Method takes the production_runs in the class and uses them to 
        generate production run objects

        Returns:
        None
        """
        
        runs = []
        
        if len(self.production_runs)==0:
            print("Note there are no production runs of any significant length between {} and {}".format(self.starttime.strftime("%Y-%m-%d %H:%M", self.endtime.strftime("%Y-%m-%d %H:%M"))))
            return
        # print(self.production_runs)
        for i in self.production_runs:
            print(i)
            r = Production_run(date = self.date, start=i[1], end=i[2], recipe=i[0], moisture = self.moisture)
            runs.append(r)

        self.production_runs = runs
        return

class Production_run:
    """
    A Production run is a time frame where there are no setpoint or recipe 
    changes or periods of downtime, this allows us to exclude special causes 
    when programatically examining process data.

    Attributes:

    Methods:
    """
    def __init__(self, date, start, end, recipe, moisture):
        """
        Initialises a Production run for a given day, starttime endtime and 
        recipe. 

        Attributes:
        date (datetime):
        starttime (datetime):
        endtime (datetime):
        recipe (list):
        moisture (float):
        recipe_str (str):
        conditions (dict):
        V901_ctrl_figs (dict):
        L205_ctrl_figs (dict):
        # Bentonorit_ctrl_figs (dict):

        Returns:
        None
        """
        self.date = date
        self.starttime = start
        self.endtime = end
        self.recipe = recipe
        self.moisture = moisture
        self.recipe_str = None
        self.conditions = None
        self.get_recipe_str()
        self.get_conditions()
        self.V901_ctrl_figs = None
        self.L205_ctrl_figs = None
        self.mixes_duration = None
    
    def get_recipe_str(self):
        """ 
        Takes the recipe list of the class and generates the recipe string, 
        using the setpoint data to make an inferance. Updates the attribute 
        recipe_str

        Returns:
        None
        """
        # the wood flour setting is neccessary to determin the difference 
        # between CNR120 and CNR115LB

        
        if "wf_tot" in self.recipe[-1]:
            if self.recipe[-1]["wf_tot"] < 70: 
                self.recipe_str = "CNR115LB"
            else:
                self.recipe_str = "CNR120"
        
        elif "osf_tot" in self.recipe[-1]:
            self.recipe_str = "CNR115"
        
        elif len(self.recipe[-1]) == 6:
            carbon = {
                1.0:"recipe_2 2kilns",
                2.0:"recipe_2",
                3.0:"recipe_4 2kilns",
                4.0:"recipe_4",
                5.0:"recipe_6 2 kilns",
                6.0:"recipe_6" 
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
        
        else:
            tags = ["osf_tot", "wf_tot", "acid_tot"]
            # print("has wf")
            for i in self.recipe[-1]:
                if i == "wf_tot":
                    threshold = 20
                else:
                    threshold = 100
                temp = ctrl_data_filter(self.conditions[i], threshold=threshold)
        
        # elif self.recipe[1] == 1:
        #     tags = ["osf_tot", "acid_tot"]
        #     # print("is std")
        #     for i in :
        #         temp = ctrl_data_filter(self.conditions[i], threshold=100)
    
    def get_conditions(self):
        """ Within the Production run this looks at all the important KPVs of 
        the site for a given process"""
        # Note this does not consider the residence time of the R_plant or 
        # Main plant. Further work will be required to include delays to the 
        # periods observed in the rest of the plant so that we can plot the 
        # progress of a pallet moving through the whole of the plant.  
        
        # Dictionaries of the important KPVs of the site
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
            # "S93":"PEL1.PELL2.PAS.SC.01.PV","S94":"PEL1.PELL2.PAS.SC.01.PV", # pelletisers
            # "T91":"PEL1.DRYB1.FG.TC.01.PV","T93":"PEL1.DRYB1.FG.TC.03.PV",
            # "T94":"PEL1.DRYB1.FG.TC.04.PV","T95":"PEL1.DRYB1.FG.TC.05.PV",
            # "T96":"PEL1.DRYB1.PAS.TI.01.PV","S94":"PEL1.DRYB1.PAS.SI.01.PV",
            # "pres_z1":"PEL1.DRYB1.FG.PI.01.PV","F92":"PEL1.DRYB1.STM.FI.01.PV", # Preactivation 
            # "T03":"KLN1.BED3.UWAC.TC.01.PV","T19":"KLN1.VENT3.FG.TI.01.PV",
            # "F03":"KLN1.BURN3.NG.FI.01.PV","F16":"KLN1.EXHA1.FG.FI.01.PV", # Activation 
            # "F07":"WSH1.TK3.PPHO.FC.01.PV", "F09":"WSH1.TK2.PPHO.FC.01.PV",
            # "F25":"WSH1.TK6.FW.FC.01.PV","F26":"WSH1.TK7.FW.FC.01.PV",
            # "F27":"WSH1.TK7.FW.FC.02.PV","F30":"WSH1.TK8.DNAOH.FI.01.PV",
            # "F31":"WSH1.TK8.DNAOH.FI.02.PV","P06":"WSH1.WSH1.FG.PI.02.PV", # Wash
            "C01":"WSH1.TK8.DNAOH.AC.01.PV","C02":"WSH1.WSH1.WW.AI.01.PV", # Wash performance
            "F28":"TK1.TK2.PPHO.FC.01.PV","F11_fresh":"TK1.TK1.FPHO.FC.01.PV",
            "F11_plant":"TK1.TK1.PHO.FC.01.PV","D01":"TK1.TK3.PPHO.DI.01.PV",
            "D02":"TK1.TK2.PPHO.DI.01.PV", # Acid recovery
            # "L302":"DRYR1.FEED3.WAC.SC.02.SP","T06":"DRYR1.EXHA1.FG.TC.01.PV",
            # "T41A":"DRYR1.BED1.GAC.TC.01.PV","T41B":"DRYR1.BED1.GAC.TC.02.PV", # Dryer
        }
        Ccarbon_KPVs = {
            "W01":"KLN1.MIXR1.SAW.WC.01.PV", "W02":"KLN1.MIXR2.SAW.WC.01.PV", 
            "W03":"KLN1.MIXR3.SAW.WC.01.PV", "F04":"KLN1.MIXR1.PPHO.FC.01.PV",
            "F05":"KLN1.MIXR2.PPHO.FC.01.PV","F06":"KLN1.MIXR3.PPHO.FC.01.PV", 
            "M01":"KLN1.MIXR2.SAW.AI.01.PV",# C-carbon Mix
            # # "A Drive":"KLN1.KLN1.UWAC.XS.01.PV","B Drive":"KLN1.KLN2.UWAC.XS.01.PV",
            # # "C Drive":"KLN1.KLN3.UWAC.XS.01.PV", 
            # "T01":"KLN1.BED1.UWAC.TC.01.PV",
            # "T02":"KLN1.BED2.UWAC.TC.01.PV","T03":"KLN1.BED3.UWAC.TC.01.PV",
            # "T17":"KLN1.VENT1.FG.TI.01.PV", "T18":"KLN1.VENT2.FG.TI.01.PV",
            # "T19":"KLN1.VENT3.FG.TI.01.PV", "F01":"KLN1.BURN1.NG.FI.01.PV",
            # "F02":"KLN1.BURN2.NG.FI.01.PV","F03":"KLN1.BURN3.NG.FI.01.PV",
            # "F16":"KLN1.EXHA1.FG.FI.01.PV", # Activation 
            # "F07":"WSH1.TK3.PPHO.FC.01.PV", "F09":"WSH1.TK2.PPHO.FC.01.PV",
            # "F25":"WSH1.TK6.FW.FC.01.PV","F26":"WSH1.TK7.FW.FC.01.PV",
            # "F27":"WSH1.TK7.FW.FC.02.PV","F30":"WSH1.TK8.DNAOH.FI.01.PV",
            # "F31":"WSH1.TK8.DNAOH.FI.02.PV","P06":"WSH1.WSH1.FG.PI.02.PV", # Wash
            # "C01":"WSH1.TK8.DNAOH.AC.01.PV","C02":"WSH1.WSH1.WW.AI.01.PV", # Wash performance
            "F28":"TK1.TK2.PPHO.FC.01.PV","F11_fresh":"TK1.TK1.FPHO.FC.01.PV",
            "F11_plant":"TK1.TK1.PHO.FC.01.PV","D01":"TK1.TK3.PPHO.DI.01.PV",
            "D02":"TK1.TK2.PPHO.DI.01.PV", # Acid recovery
            # "L302":"DRYR1.FEED3.WAC.SC.02.SP","T06":"DRYR1.EXHA1.FG.TC.01.PV",
            # "T41A":"DRYR1.BED1.GAC.TC.01.PV","T41B":"DRYR1.BED1.GAC.TC.02.PV", # Dryer
        }
        
        conditions = {}
               
        # Collect the C-carbon process KPVs if the recipe_str is a C-carbon one
        if "recipe" in self.recipe_str:
            for i in Ccarbon_KPVs:
                
                if i in self.recipe[-1] or i in ["M01","D01"]:
                    conditions[i] = fetch_tag_detailed(Ccarbon_KPVs[i], start=self.starttime, end=self.endtime, resolution=1)
                else:
                    conditions[i] = fetch_tag(Ccarbon_KPVs[i],start=self.starttime,end=self.endtime)
                conditions[i].name = self.starttime.strftime(" %Y-%m-%d") + i 
        
        # Collect the ELCD process KPVs if the recipe_str is a ELCD_KPVs
        if "CNR" in self.recipe_str:
            for i in ELCD_KPVs:
                conditions[i] = fetch_tag(ELCD_KPVs[i],start=self.starttime, end=self.endtime)
                conditions[i].name = self.starttime.strftime(" %Y-%m-%d") + i 
        
        # Update the conditions attribute with a dictionary of the dataframes of all the important KPVs 
        self.conditions = conditions

    def V_901_ctrl_chart(self):
        """
        Generate a set of control charts for the production run if it is an 
        ELCD production run.
        """
        def ctrl_data_filter(df, threshold):
            """
            Function to select all the valuesbefore a 0 in the data frame
            
            Parameters:
            df (dataframe): data frame of the totaliser values
            threshold (float): the value bellow which all values are set to 0

            Returns:
            totals (dataframe) a data frame of the values before a 0 value. 
            """
            # set any value below threshold in data frame to 0
            df.loc[df['VALUE']<threshold, 'VALUE'] = 0
            # create a new dataframe only including rows where Value is not 0 and the next value is 0.
            totals = df.loc[(df['VALUE']!=0) & (df['VALUE'].shift(1)==0)].copy()
            # Create a moving range column for the new data frame
            totals['MR'] = abs(totals['VALUE']-totals['VALUE'].shift(1))
            # perhps mixes_duration is miss labelled since this value is really
            #  the number of mixes.
            self.mixes_duration = len(totals)
            return totals

        def plot_IMR(df, tag, SP):
            """
            This function takes a data frame, tag and set point and generates a
            control chart for the data.

            Returns:
            None

            """
            def filter_none_empty(data):
                if isinstance(data, pd.DataFrame):
                    return not data.empty
                return data is not None

            # Skips senarios where there is no data in the data frame
            if not filter_none_empty(df):
                # print("Either no data frame has been recieved or the dataframe is empty")
                return None          
            
            # Set-up graph format
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
            
            # Open BytesIO object to eventually be passes to word document
            memfile = BytesIO()

            # Calculate means and important control parameters
            xbar = df['VALUE'].mean()
            mrbar = df['MR'].mean()
            UCL = xbar + mrbar * 2.66
            LCL = xbar - mrbar * 2.66
            rUCL = mrbar * 3.27
            sig = mrbar / 1.128
            sigma = df['VALUE'].std()
            
            # System for determining the specification limits of each grade.
            # Given the problems with the Olive Stone flour circa 2023 May-Aug
            # and the much larger volume of acid that is required by this OSF
            # we have had serious deviations from the control plan which 
            # specifies these figures. Another way of delivering these values 
            # to the controlchart function may be neccessary
            # print(tag)

            if tag.split("_")[0] == "acid":     # acid specs
                if self.recipe_str == "CNR115":         # Standard specs
                    spec = {
                        "USL" : 490,
                        "LSL" : 450
                    }                
                elif self.recipe_str == "CNR115LB" or "CNR120":
                    spec = {
                        "USL" : 530,
                        "LSL" : 470
                        }
            elif tag.split("_")[0] == "wf":     # wf specs 
                # print(self.recipe_str)
                if self.recipe_str == "CNR120":   
                    spec = {
                        "USL" : 100,
                        "LSL" : 90
                        }
                elif self.recipe=="CNR115LB": 
                    spec = {              
                        "USL" : 60,
                        "LSL" : 40
                        }
            
            elif tag.split("_")[0] == "osf":    # osf specs
                if self.recipe_str == "CNR115": 
                    spec = {
                        "USL" : 510,
                        "LSL" : 490
                        }
                elif self.recipe_str == "CNR120":
                    spec = {
                        "USL": 310,
                        "LSL": 290
                        }
                elif self.recipe_str == "CNR115LB":
                        spec = {               
                            "USL":460,
                            "LSL":440
                        }
            
            # the specifications for the solid:acid ratio have bee calculated 
            # at the absolute maximum of the range using 3% to 10% moisture. 
            # These specification limits will need to be recalculated whenever 
            # a new olive stone with new acid requirments is recieved.
         
            elif tag.split("_")[0] == "ratio":  # solid:acid ratio specs
                if self.recipe_str== "CNR115":  
                    spec = {
                        "USL" : 1.223,
                        "LSL" : 0.932
                        }
                elif self.recipe_str == "CNR120":
                    spec = {
                        "USL": 1.704,
                        "LSL": 1.182
                        }
                elif self.recipe_str == "CNR115LB":
                    spec = {
                        "USL":1.349,
                        "LSL":0.954
                    }

            # # Calculate control indexes based off Specification limits
            Cp = round((spec["USL"] - spec["LSL"]) / (6 * sig), 2)
            Cpk = round(min([(spec["USL"] - xbar)/(3 * sig),(xbar - spec["LSL"])/(3 * sig)]), 2)
            # Pp = round((spec["USL"]-spec["LSL"])/(6*sigma), 2)
            # Ppk = round(min([(spec["USL"] - xbar)/(3*sigma),(xbar - spec["LSL"])/(3*sigma)]), 2)

            fig, ax = plt.subplots(nrows=2, sharex=True)
            fig.set_size_inches(7,1.8)
            ax1, ax2 = ax.flatten()
            ax1.plot(df['TS'], df['VALUE'], linestyle='-', marker='.',markersize=5, color='red',label='VALUE')

            date = self.starttime.strftime("%Y-%m-%d")
            start_str = self.starttime.strftime("%Y-%m-%d %H:%M")
            end_str = self.endtime.strftime("%Y-%m-%d %H:%M")
            name = tag.split("_")[0]
            duration = round((self.endtime - self.starttime).total_seconds()/3600,4)
            no_mixes = len(df)

            # Title for each plot based on calculated parameters
            # print(f"{self.recipe}, {name}, {start_str}, {end_str}, {duration}, {no_mixes}, {xbar}, {mrbar}, {sig}, {sigma}, {LCL}, {UCL}, {Cp}, {Cpk}, {Pp},{Ppk}")
            ax1.set_title(f"{date} {self.recipe_str} {name} control, from {start_str} to {end_str}. {no_mixes} mixes. Cp = {Cp}, Cpk = {Cpk}") #, Pp = {Pp}, Ppk = {Ppk}")
            
            # Create the Lines for an I chart
            # ax1.axhline(spec["USL"], color="#2D1CAB", linestyle="dotted",label="USL = {}".format(str(spec["USL"])))
            # ax1.axhline(spec["LSL"], color="#2D1CAB", linestyle="dotted",label="USL = {}".format(str(spec["LSL"])))
            ax1.axhline(xbar, color='green', label='xbar = {}'.format(str(round(xbar, 2))))
            ax1.axhline(UCL, color='orange', linestyle='--', label='UCL = {}'.format(str(round(UCL, 2))))
            ax1.axhline(LCL, color='orange', linestyle='--', label='LCL = {}'.format(str(round(LCL, 2))))
            print(tag.split("_")[0])
            print(type(SP))
            if (type(SP) is float or type(SP) is int) and tag.split("_")[0] != "ratio":
                ax1.axhline(SP, color='magenta', linestyle='solid',label=f'SP = {str(SP)}')
            
            # Create Lines for MR chart
            ax2.plot(df['TS'], df['MR'], color='red', marker='.',markersize=5,linestyle='-',label='MR')
            ax2.axhline(mrbar, color='green', label='MRbar = {}'.format(str(round(mrbar, 2))))
            ax2.axhline(rUCL, color='orange', label='MR UCL = {}'.format(str(round(rUCL, 2))))
            ax2.set_xlabel('Time')
            ax2.xaxis.set_major_formatter(mdates.DateFormatter("%d %H:%M"))
            ax1.set_ylabel('Value')
            ax2.set_ylabel('MR')

            # Adjust Layout and size for plotting in a word Document
            plt.tight_layout()
            box1 = ax1.get_position()
            ax1.set_position([box1.x0, box1.y0, box1.width * 0.85, box1.height])
            box2 = ax2.get_position()
            ax2.set_position([box2.x0, box2.y0, box2.width * 0.85, box2.height])

            # Position Legen
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
            
            # Save to BytesIO object for writing to word document
            plt.savefig(memfile)
            return memfile
        
        def dry_cell_dry_acid():
            
            def density_to_conc(x):
                return round(-527.057716 + 1.074723463 * x - 0.000801043 * x**2 + 2.95816E-07 * x**3 + (-4.2184E-11)* x ** 4, 4)

            if self.recipe[1] == 2:

                tags = ["osf_tot", "wf_tot", "acid_tot"]
                thresh = {
                    "osf_tot":100,
                    "wf_tot":30,
                    "acid_tot":100,
                }
            
            elif self.recipe[1] == 1:

                tags = ["osf_tot", "acid_tot"]
                thresh = {
                    "osf_tot":100,
                    "acid_tot":100,
                }
        
            data = [ctrl_data_filter(self.conditions[i],thresh[i]).reset_index() for i in tags]
            data[0] = data[0].drop('MR', axis=1)
            # print(data[0])
            concated_data = pd.concat([data[0]]+[df['VALUE'].rename(f'{tags[i+1]}') for i, df in enumerate(data[1:])], axis=1)
            concated_data.rename(columns={'VALUE':'osf_tot'}, inplace=True)
            # print(concated_data)
            concated_data = concated_data.dropna(axis=0)
            if concated_data.empty:
                # print("No ratio data from {} production period".format(self.starttime.strftime("%Y-%m-%d %H:%M")))
                return
            
            d01_lookup = self.conditions["D01"]
            
            for index, row in concated_data.iterrows():
                target_datetime = row['TS']
                d01_lookup['time_diff'] = abs(d01_lookup['TS'] - target_datetime)
                closest_row_index = d01_lookup['time_diff'].idxmin()
                closest_value = d01_lookup.loc[closest_row_index, 'VALUE']
                concated_data.at[index,'D01'] = closest_value

            for index, row in concated_data.iterrows():
                liquid_mass = row["acid_tot"] * row["D01"]
                conc = density_to_conc(row["D01"])
                # print(conc)
                dry_acid = liquid_mass * conc / 100000
                s = [row[i] for i in tags[:-1]]
                solids = sum(s) * (100 - self.moisture)/100
                Ac_S_ratio = dry_acid / solids

                concated_data.at[index, "VALUE"] = Ac_S_ratio
            
            concated_data['MR'] = abs(concated_data['VALUE']-concated_data['VALUE'].shift(1))
            
            if self.recipe[1] == 2:
                    if self.recipe[-1]["wf_tot"]>70:
                        concated_data['GRADE'] = "CNR120"
                    else:
                        concated_data['GRADE'] = "CNR115LB"
            
            elif self.recipe[1] == 1:
                concated_data['GRADE'] = "CNR115"

            # concated_data.to_csv(f"{self.starttime.strftime('%Y-%m-%d_%H-%M')} ELCD_mix_data.csv" ,mode='a',index = False, header = None)

            # print(concated_data)

            return concated_data[["TS", "VALUE", "MR"]]
        
        
        ctrl = {}
        
        # print(self.starttime.strftime("%d %H:%M"), self.endtime.strftime("%d %H:%M"), self.recipe)

        if self.recipe[0] <= 6:
            print("there is no ELCD production during this period {} to {}".format(self.starttime.strftime("%Y-%m-%d %H:%M"), self.endtime.strftime("%Y-%b-%d %H:%M")))
            return        

        elif self.recipe[1] == 2:
            A_s_ratio = dry_cell_dry_acid()
            tags = ["osf_tot", "wf_tot" , "acid_tot"]
            # print("has wf")
            for i in tags:
                if i == "wf_tot":
                    threshold = 20
                else:
                    threshold = 100
                
                temp = ctrl_data_filter(self.conditions[i], threshold=threshold)
                
                if i == "wf_tot":
                    print("Generating wf chart")
                    ctrl[i] = plot_IMR(temp, i, self.recipe[3][i])
                elif i == "osf_tot":
                    print("Generating osf chart")
                    ctrl[i] = plot_IMR(temp, i, self.recipe[3][i])
                elif i == "acid_tot":
                    print("Generating acid chart")
                    ctrl[i] = plot_IMR(temp, i, self.recipe[3][i] )
            
            if self.recipe[-1]["wf_tot"] > 70:
                print("Generating ratio chart")
                ctrl["ratio_tot"] = plot_IMR(A_s_ratio, "ratio_tot", 1.532)
            else:
                print("Generating ratio chart")
                ctrl["ratio_tot"] = plot_IMR(A_s_ratio, "ratio_tot", 1.112)

            self.V901_ctrl_figs = ctrl
            
            return

        elif self.recipe[1] == 1:
            A_s_ratio = dry_cell_dry_acid()
            tags = ["osf_tot", "acid_tot"]
            # print("is std")
            for i in tags:
                temp = ctrl_data_filter(self.conditions[i], threshold=100)
                if i == "osf_tot":
                    ctrl[i] = plot_IMR(temp, i, self.recipe[3])
                elif i == "acid_tot":
                    ctrl[i] = plot_IMR(temp, i, 470 )
            
            ctrl["ratio_tot"] = plot_IMR(A_s_ratio, "ratio_tot", 1.034)
            
            self.V901_ctrl_figs = ctrl
            return         
        
        else:
            # print("the recipe for ELCD is neither 1 nor 2")
            
            return

    def L205_ctrl_chart(self):

        def density_to_conc(x):
            return round(-527.057716 + 1.074723463 * x - 0.000801043 * x**2 + 2.95816E-07 * x**3 + (-4.2184E-11)* x ** 4, 4)
        
        def dry_acid_dry_wood(kiln):
           
            def dadw(row):
                if row["wood"]==0:
                    return 0
                else:
                    return (density_to_conc(row["D01"]) * row["acid"] * row["D01"] / ( row["wood"] * (1 - row["M01"]/100) * 100000 ))
            
            m01_lookup = self.conditions["M01"]
            d01_lookup = self.conditions["D01"]

            data = [self.conditions[kiln[i]] for i in kiln] + [ m01_lookup] + [d01_lookup]
            data[0].rename(columns={"VALUE":"acid"}, inplace=True)
            # data[1].rename(columns={"VALUE":"wood"}, inplace=True)
            concated_data = pd.concat([data[0]]+[df['VALUE'].rename(f'{df.name[-3:]}') for df in data[1:]], axis=1)
            
            wood_col = [col for col in concated_data.columns if col.startswith("W0")]
            
            for col in wood_col:
                concated_data.rename(columns={col:"wood"}, inplace=True)
            
            concated_data["ratio"] = concated_data.apply(dadw,axis=1)
            
            concated_data.set_index("TS", inplace=True)
            concated_data["VALUE"] = concated_data["ratio"].rolling(window='20T').mean()
            concated_data['MR'] = abs(concated_data['ratio']-concated_data['ratio'].shift(1))
            concated_data["STD"] = concated_data["ratio"].rolling(window="20T").std()
            
            concated_data["LCL"] = concated_data["VALUE"] - 3 * concated_data["STD"]
            concated_data["UCL"] = concated_data["VALUE"] + 3 * concated_data["STD"]

            concated_data = concated_data.reset_index()
            concated_data.name = "Phos cell Rto"

            return concated_data[["TS","VALUE","MR","LCL","UCL"]]

        def acid_wood_RMR(kiln):
            out = {}
            for i in kiln:
                data = self.conditions[kiln[i]]
                
                # data.rename(columns={"VALUE":"RAW"}, inplace=True)
                # print(data)
                data.set_index("TS", inplace=True)
                if "acid" in data.columns:
                    data["STD"] = data["acid"].rolling(window="20T").std()
                    data["VALUE"] = data["acid"].rolling('20min').mean()
                    data['MR'] = abs(data['acid']-data['acid'].shift(1))
                else:
                    data["STD"] = data["VALUE"].rolling(window="20T").std()
                    data['MR'] = abs(data['VALUE']-data['VALUE'].shift(1))
                    data["VALUE"] = data["VALUE"].rolling('20min').mean()

                data["LCL"] = data["VALUE"] - 3 * data["STD"]
                data["UCL"] = data["VALUE"] + 3 * data["STD"]
                data.reset_index(inplace=True)
                # print(data)
                out[i] = data
            return out

        def plot_rol_MR(df, kiln, name):
            
            def filter_none_empty(data):
                if isinstance(data, pd.DataFrame):
                    return not data.empty
                return data is not None

            if not filter_none_empty(df):
                # print("Either no data frame has been recieved or the dataframe is empty")
                return None          
            
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
            sigma = df['VALUE'].std()
            # UCL = xbar + sigma * 3
            # LCL = xbar - sigma * 3
            rUCL = mrbar * 3.27
            sig = mrbar / 1.128
            if df.name[-3:] in self.recipe[-1]:
                SP = self.recipe[-1][df.name[-3:]]
            else:
                k_sp_dict = { "A":["F04","W01"],
                             "B":["F05","W01"],
                             "C":["F06","W03"]}
                den_trgt = 1485
                conc_trgt = density_to_conc(den_trgt)
                SP = den_trgt * self.recipe[-1][k_sp_dict[kiln][0]]* conc_trgt/ ( self.recipe[-1][k_sp_dict[kiln][1]] * 0.9 * 100000 )
                SP= round(SP,3)

            # Cp = round((spec["USL"] - spec["LSL"]) / (6 * sig), 2)
            # Cpk = round(min([(spec["USL"] - xbar)/(3 * sig),(xbar - spec["LSL"])/(3 * sig)]), 2)
            # Pp = round((spec["USL"]-spec["LSL"])/(6*sigma), 2)
            # Ppk = round(min([(spec["USL"] - xbar)/(3*sigma),(xbar - spec["LSL"])/(3*sigma)]), 2)

            spec = {"USL": SP,
                  "LSL": SP
                  }
            Cp = 1
            Cpk = 1
            Pp = 1
            Ppk = 1

            fig, ax = plt.subplots(nrows=2, sharex=True)
            fig.set_size_inches(7,2.3)
            ax1, ax2 = ax.flatten()
            ax1.plot(df['TS'], df['VALUE'], linestyle='-',markersize=5, color='red',label='VALUE')
            
            if name == "tio":
                ax1.set_title("{} kiln ratio from {} to {}".format(kiln, self.starttime.strftime("%Y-%m-%d %H:%M"), self.endtime.strftime("%Y-%m-%d %H:%M")))
            else:
                ax1.set_title("{} kiln {} from {} to {}".format(kiln, name ,self.starttime.strftime("%Y-%m-%d %H:%M"), self.endtime.strftime("%Y-%m-%d %H:%M")))
            # if self.recipe[1]== 1:
            #     print("CNR115, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {},".format(tag.split("_")[0],self.starttime.strftime("%Y-%m-%d %H:%M"), self.endtime.strftime("%Y-%m-%d %H:%M"), round((self.endtime - self.starttime).total_seconds()/3600,4), str(len(df)), xbar, mrbar, sig, sigma, LCL, UCL,Cp, Cpk, Pp, Ppk ))
            #     ax1.set_title("{} CNR115 {} control, from {} to {}. {} mixes. Cp = {}, Cpk = {}, Pp = {}, Ppk = {}".format(self.starttime.strftime("%Y-%m-%d"), tag.split("_")[0],self.starttime.strftime("%H:%M"), self.endtime.strftime("%H:%M"), str(len(df)),Cp, Cpk, Pp, Ppk))
            # elif self.recipe[1] == 2:
            #     if self.recipe[-1]["wf"] > 70:
            #         print("CNR120, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {},".format(tag.split("_")[0],self.starttime.strftime("%Y-%m-%d %H:%M"), self.endtime.strftime("%Y-%m-%d %H:%M"), round((self.endtime - self.starttime).total_seconds()/3600,4), str(len(df)), xbar, mrbar, sig, sigma, LCL, UCL,Cp, Cpk, Pp, Ppk ))
            #         ax1.set_title("{} CNR120 {} control, from {} to {}. {} mixes. Cp = {}, Cpk = {}, Pp = {}, Ppk = {}".format(self.starttime.strftime("%Y-%m-%d"), tag.split("_")[0],self.starttime.strftime("%H:%M"), self.endtime.strftime("%H:%M"), str(len(df)),Cp, Cpk, Pp, Ppk))
            #     else:
            #         print("CNR115LB, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {},".format(tag.split("_")[0],self.starttime.strftime("%Y-%m-%d %H:%M"), self.endtime.strftime("%Y-%m-%d %H:%M"), round((self.endtime - self.starttime).total_seconds()/3600,4), str(len(df)), xbar, mrbar, sig, sigma, LCL, UCL,Cp, Cpk, Pp, Ppk ))
            #         ax1.set_title("{} CNR115LB {} control, from {} to {}. {} mixes. Cp = {}, Cpk = {}, Pp = {}, Ppk = {}".format(self.starttime.strftime("%Y-%m-%d"), tag.split("_")[0],self.starttime.strftime("%H:%M"), self.endtime.strftime("%H:%M"), str(len(df)),Cp, Cpk, Pp, Ppk))
            
            # if tag.split("_")[0] not in  ["acid", "ratio", "wf"]:
            #     ax1.axhline(spec["USL"], color="#2D1CAB", linestyle="dotted",label="USL = {}".format(str(spec["USL"])))
            #     ax1.axhline(spec["LSL"], color="#2D1CAB", linestyle="dotted",label="USL = {}".format(str(spec["LSL"])))
            
            ax1.axhline(xbar, color='green', label='xbar = {}'.format(str(round(xbar, 3))))
            # ax1.axhline(UCL, color='orange', linestyle='--', label='UCL = {}'.format(str(round(UCL, 2))))
            # ax1.axhline(LCL, color='orange', linestyle='--', label='LCL = {}'.format(str(round(LCL, 2))))
            ax1.axhline(SP, color='magenta', linestyle='solid',label='SP = {}'.format(str(SP)))
            ax1.plot(df["TS"], df["LCL"], color='orange', linestyle="-", label="LCL")
            ax1.plot(df["TS"], df["UCL"], color='orange', linestyle="-", label="UCL")
            ax2.plot(df['TS'], df['MR'], color='red',markersize=5,linestyle='-',label='MR')
            
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

        if self.recipe[0] > 6:
            # print("there is no C-carbon production during this period")
            return

        else:
            A_kiln = {"acid":"F04", "wood":"W01"}
            B_kiln = {"acid":"F05", "wood":"W02"}
            C_kiln = {"acid":"F06", "wood":"W03"}

            C_kiln_ratio = dry_acid_dry_wood(C_kiln)
            C_kiln_ratio.name = "acid_ratio"
            # print(C_kiln_ratio)
            C = acid_wood_RMR(C_kiln)
            C["ratio"] = C_kiln_ratio
            
            B_kiln_ratio = dry_acid_dry_wood(B_kiln)
            B_kiln_ratio.name = "acid_ratio"
            # print(B_kiln_ratio)
            B = acid_wood_RMR(B_kiln)
            B["ratio"] = B_kiln_ratio
            
            A_kiln_ratio = dry_acid_dry_wood(A_kiln)
            A_kiln_ratio.name = "acid_ratio"
            # print(A_kiln_ratio)
            A = acid_wood_RMR(A_kiln)
            A["ratio"] = A_kiln_ratio
            
            for i in C:
                ctrl["C " + i] = plot_rol_MR(C[i],"C",C[i].name[-3:])

            for i in B:
                ctrl["B " + i] = plot_rol_MR(B[i],"B",B[i].name[-3:])
            
            for i in A:
                ctrl["A " + i] = plot_rol_MR(A[i],"A",A[i].name[-3:])

            self.L205_ctrl_figs = ctrl
            return

def gen_date_range(start, end):
    dates = [] 
    cur = start
    while cur <= end:
        dates.append(cur)
        cur += dt.timedelta(days=1)
    return dates
dates = gen_date_range(dt.datetime(year=2023,month=5,day=3), dt.datetime(year=2023,month=5,day=4))

print("Here we go")
for i in dates:
    print(f"attempting to create {i.isoformat()}")
    D = Production_day(date=i)
    print("Day created, printing data")
    D.print_ELCD_mixes()
    print("finished plotting onto the next day")




# writer.save()
