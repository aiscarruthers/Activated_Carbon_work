# -*- coding: utf-8 -*-
"""
Created on Thu Jul 18 11:43:21 2019

@author: PHester
"""

import os

import pandas as pd

os.chdir('W:\\Process Data\\Daily Reports\\MBR Plant\\2020')

files = os.listdir()

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
        df2 = pd.read_excel(file, sheet_name='MBR Perf Chal')
        df2=df2[0:1]
        df2['Date']=file[4:14]
        dfl.append(df2)
    
    return pd.concat(dfl)

    


allData = [getMonth(folder) for folder in files[:12]]
    
print('loop done')

allDataDf = pd.concat(allData)
    
    
    
os.chdir('C:\\Users\\phester\\OneDrive - cabotcorp.com\\Desktop')
    
allDataDf.to_excel('MBRPerfCSafetoDelete.xlsx')