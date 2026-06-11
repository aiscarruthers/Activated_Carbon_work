# -*- coding: utf-8 -*-
"""
Created on Thu Jul 18 11:43:21 2019

@author: PHester
"""

import os

import pandas as pd

os.chdir('W:\\Process Data\\Daily Reports\\MBR Plant\\2020')

files = os.listdir()


if 'Master' in files:
    files.remove('Master')
# selct files


def getMonth(fileName):
    
    os.chdir('W:\\Process Data\\Daily Reports\\MBR Plant\\2020\\'+fileName)
    print('W:\\Process Data\\Daily Reports\\MBR Plant\\2020\\'+fileName)
    files2 = os.listdir()
    
    if len(files2) == 0:
        
        return
        
    
    dfl=[]
    
    for file in files2:
        
        #if file[0] == '$':
           # continue
        df2 = pd.read_excel(file, sheet_name='Data')
        
        df3=df2.copy().drop([0,1,2,3,4,5,6])       
        cols= list(df2.iloc[4])
        cols[0]='Date'
        df3.columns=cols
        df3.index=df3['Date']
        df3.drop(columns='Date',inplace=True)
        df3=df3.apply(pd.to_numeric).resample('H',axis=0).mean()

        dfl.append(df3)
    
    return pd.concat(dfl)

    


allData = [getMonth(folder) for folder in files[:13]]
    
print('loop done')

allDataDf = pd.concat(allData)

allDataDf = allDataDf.apply(pd.to_numeric).resample('D',axis=0).mean()
    
    
    
os.chdir('C:\\Users\\phester\\OneDrive - cabotcorp.com\\Desktop')
    
allDataDf.to_excel('MBRAllData.xlsx')