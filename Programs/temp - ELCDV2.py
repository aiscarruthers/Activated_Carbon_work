# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import os

import pandas as pd

import datetime






def getFiles(a):
    # need my code to grab files
    
    return None
        
        

def historicDesc(filename):
    
    #this works for STD not wood flour
    wantedCols=['Actual Acid Addition to Mixer V901','C01 ADD WASH TK 309','D01 SUPPLY TANK DENSITY ','D02 ML TO RETURN TK702 ','D401 Dryer Moisture Result from S/Man','D401 Dryer pH Result from S/MAN ','D901 DRYER DISCHARGE TEMP','D901 DRYER EXHAUST TEMP','D901 DRYER STEAM FLOW F92','D901 DRYER ZONE 1 PRESSURE','D901 DRYER ZONE 1 TEMP T91','D901 DRYER ZONE 2 TEMP T92','D901 DRYER ZONE 3 TEMP T93','D901 DRYER ZONE 4 TEMP T94','D901 DRYER ZONE 5 TEMP T95','F03 NAT GAS TO KILN C','F07 1ST WASH TO DIG','F08  WATER TO PRAYON','F09 ML TO RETURN TANK','F10 ML TO DIGESTER','F11 Acid Consumed','F12 NAT GAS TO H401','F16 KILN GAS FROM V601','F17 DRY EX GAS TO STACK','F20 NAT GAS TO BOILERS','F25 WATER TO 4TH WASH','F26 WATER WASH 1ST BANK','F27 WATER WASH 2ND BANK','F28 PHOS ACID TRANSFER','F30 Additional Wash 1st Bank Flow ','F31 Additional Wash 2nd Bank Flow','Kiln C Venturi Water Flow','L04 BUFFER TK706','P01 DRYER HOOD PRESURE','P06 PRAYON WASH SUCTION','Pelletiser A Setpoint','Pelletiser B Setpoint','S01 PRAYON SPEED','T03 CARBON KILN C','T04 DRYER INLET GAS','T06 DRYER OUTLET  GAS','T19 KILN C BACK BOX TEMP','T41A CARBON DRYER BED TEMP','T41B CARBON DRYER BED TEMP']

    #filename='W:\\Process Data\\Daily Reports\\ELCD\\2020\\11. November\\ELCD 11.25.2020.xlsm'
    cols = ['DateTime', 'CNR(Std) Mixer Batch Counter', 'CNR(LB) Mixer Batch Counter', 'Actual Acid Addition to Mixer V901', 'Actual OSF Addition to Mixer V901', 'Pelletiser A Setpoint', 'Pelletiser B Setpoint', 'C01 ADD WASH TK 309', 'C02  WSH TO DRAIN TK309', 'KILNS KILN C DRIVE', 'DRY CARBON DRYER DRIVE', 'D901 DRYER CONVEYOR ', 'D401 Dryer Moisture Result from S/Man', 'D401 Dryer pH Result from S/MAN ', 'D01 SUPPLY TANK DENSITY ', 'D02 ML TO RETURN TK702 ', 'PRAYN PRAYON DRIVE', 'F03 NAT GAS TO KILN C', 'F10 ML TO DIGESTER', 'F07 1ST WASH TO DIG', 'F09 ML TO RETURN TANK', 'F08  WATER TO PRAYON', 'F25 WATER TO 4TH WASH', 'F26 WATER WASH 1ST BANK', 'F27 WATER WASH 2ND BANK', 'F30 Additional Wash 1st Bank Flow ', 'F31 Additional Wash 2nd Bank Flow', 'F12 NAT GAS TO H401', 'F16 KILN GAS FROM V601', 'F23 RW TO KILNGAS V601', 'Kiln C Venturi Water Flow', 'F21 VENTURI J603 WATER', 'F17 DRY EX GAS TO STACK', 'F11 Acid Consumed', 'F18 CAUSTIC EX TK704', 'F28 PHOS ACID TRANSFER', 'Towns Water Flowrate to TK804', 'River Water Flow', 'F20 NAT GAS TO BOILERS', 'V901 ACID TO V901 FLOW F91', 'D901 DRYER STEAM FLOW F92', 'F24 RW TO V601 OR V903', 'Current (I91) A903A', 'Current (I92) A903B', 'Current (I93) A904A', 'Current (I94) A904B', 'L903A Running Average (F_CV)', 'L903B RUNNING AVERAGE (F_CV)', 'L03 ML TANK', 'L02 1ST WASH TANK', 'L04 BUFFER TK706', 'L06 CAUSTIC TK704', 'L01 SUPPLY TK703 (F_CV)', 'P06 PRAYON WASH SUCTION', 'P01 DRYER HOOD PRESURE', 'D901 DRYER ZONE 1 PRESSURE', 'T903 PRESSURE P97', 'T904 PRESSURE P98', 'MENU RECIPE NUMBER', 'S01 PRAYON SPEED', 'D901 DRYER CONVEYOR SPEED', 'D901 UNDER ZONE 1 TC', 'D901 UNDER ZONE 2 TC', 'D901 UNDER ZONE 3 TC', 'D901 UNDER ZONE 4 TC', 'D901 UNDER ZONE 5 TC', 'T03 CARBON KILN C', 'T19 KILN C BACK BOX TEMP', 'T05 DRYER INCINERATOR', 'T41A CARBON DRYER BED TEMP', 'T41B CARBON DRYER BED TEMP', 'T04 DRYER INLET GAS', 'T06 DRYER OUTLET  GAS', 'T09 KILN GAS EXIT V601', 'D901 DRYER ZONE 1 TEMP T91', 'D901 DRYER ZONE 2 TEMP T92', 'D901 DRYER ZONE 3 TEMP T93', 'D901 DRYER ZONE 4 TEMP T94', 'D901 DRYER ZONE 5 TEMP T95', 'D901 DRYER EXHAUST TEMP', 'D901 DRYER DISCHARGE TEMP', 'Actual WF Addition to Mixer V901', 'WEIGHT IN TK103 4 - 20 mA', 'TK502_3 CARBON WEIGHT TK502', 'TK502_3 CARBON WEIGHT TK503', 'TK902 WEIGH HOPPER WEIGHT', 'V902 BLENDER WEIGHT W92 ']
    df = pd.read_excel(filename,skiprows=9,usecols='A:CI')
    df.columns = cols
    df.index=df.DateTime
    date=df.iloc[0,0].date()
    
    df=df.loc[(df['KILNS KILN C DRIVE'] >= 1)&(df['PRAYN PRAYON DRIVE'] >= 1) ]
    
    df=df.loc[df['T03 CARBON KILN C']>=440]
    df=df[wantedCols].copy()
    desc=df.describe()
    
    desc= desc.transpose()
    
    desc['Date']=date
    
    return desc

def dailyCheck():
    
    #NEED TO SELCT MODE AND AUTO FIND FILE PATH
    
    yesterday=pd.datetime.now() - datetime.timedelta(1)

    yesterdayELCD = 'ELCD ' + str(yesterday.month)+'.'+str(yesterday.day)+'.'+str(yesterday.year)
    
    filepath = 'W:\\Process Data\\Daily Reports\\ELCD\\2020\\12. December'+'\\'+yesterdayELCD+'.xlsm'
    
    
    yesterdaysStats = historicDesc(filepath)
    
    yesterdaysStats.to_excel('C:\\Users\\phester\\OneDrive - cabotcorp.com\\Desktop\\Files\\pythonFiles\\ELCD Daily\\Yesterday.xlsx')
    
    
    refs=pd.read_excel('C:\\Users\\phester\\OneDrive - cabotcorp.com\\Desktop\\Files\\pythonFiles\\ELCD Daily\\refs.xlsx',index_col=0)
    
    avrefs=refs.groupby(refs.index).mean()
    
    dif=100*(yesterdaysStats.drop(columns='Date')-avrefs)/avrefs
    
    dif.to_excel('C:\\Users\\phester\\OneDrive - cabotcorp.com\\Desktop\\Files\\pythonFiles\\ELCD Daily\\diff.xlsx')
    
    return dif
    
    
    
        
