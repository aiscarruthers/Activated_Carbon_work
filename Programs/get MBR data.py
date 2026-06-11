# -*- coding: utf-8 -*-
"""
Created on Thu Jul 18 11:43:21 2019

@author: PHester
"""

import os

import pandas as pd

os.chdir('W:\\Process Data\\Daily Reports\\MBR Plant\\2020\\12. December')

files = os.listdir()

# selct files

files2 = files[:]

dfl=[]

for file in files2:
    print(file)
    if '$' in file:
        continue
    df2 = pd.read_excel(file, sheet_name='MBR Perf Chal')
    df2=df2[0:1]
    df2['Date']=file[4:14]
    dfl.append(df2)

df=pd.concat(dfl)


os.chdir('C:\\Users\\phester\\OneDrive - cabotcorp.com\\Desktop')

df.to_excel('MBRPerfCSafetoDelete.xlsx')