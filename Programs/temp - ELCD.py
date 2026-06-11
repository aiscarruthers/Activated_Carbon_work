# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import re
import os

import pandas as pd


# getFiles(2015)
# files = {'2. February':files['2. February']}

def getFiles(a):
    
    
    mode = 'ELCD'+'\\' #select mode by comenting out the one you dont want although im not looking at elcd yet

    year =str(a) # set this to pick year probally could od an input at some point

    os.chdir('W:\\Process Data\\Daily Reports\\'+mode+year) # directory location for wood

    folders = os.listdir()
    files =  {}
    for f in folders:
        os.chdir('W:\\Process Data\\Daily Reports\\'+mode+year+'\\'+f)
        listt = []
        for i in os.listdir():
            if '.xlsm' in i and '$' not in i:
                listt.append(i)
        files[f] = listt
    
    return files

        
        
            
                 
             

#os.chdir('.\\PyWork')


def historicDesc(dic):
    
    final = [] # place to put results
    
    for z in dic: # loop through the months
        
        #try:
            
        
            for i in dic[z]: # loop through each days files
                print(i)    
            
                regex = re.compile(r'(\d+).(\d+).(\d\d\d\d|\d\d)')
            
                year = regex.search(str(i))[3]
            
                if int(year) <= 1700:
                    year = int(year) + 2000
            
                mode = 'ELCD'+'\\' #select mode by comenting out the one you dont want although im not looking at elcd yet
                # df = pd.read_excel('W:\\Process Data\\Daily Reports\\ELCD\\2015\\2. February\\2.17.2015.xlsm',skiprows=6,usecols='A:CI')
                df  = pd.read_excel('W:\\Process Data\\Daily Reports\\'+mode+str(year)+'\\'+z+'\\'+i,skiprows=6,usecols='A:CI')
                df.columns = df.iloc[0] # had to change the column names 
                df = df[3:] # cuts away useless stuff

                df = df.reset_index(drop=True)
                


                
                df2 = df.loc[df['KILNS KILN C DRIVE'] >= 1  ] # filte out downtime
                 

                
                
                 
                
                desc = df2.iloc[:,1:].astype(float).describe().transpose() # Just need to add a date colum to this then poulate from the read in string
                desc = desc.assign(Date=regex.search(str(i))[1]+'.'+regex.search(str(i))[2]+'.'+str(year))
                
                desc['Date'] = pd.to_datetime(desc['Date'])
            #Need to work out somethig to join the dataframes together into one big table then maybe save to excel
                #print(desc) #troubleshooting only            
            
                final.append(desc)
        
    df2 = final[0]        
    for dfs in range(len(final)-1):
        df2 = df2.append(final[dfs+1])
        
            
    return df2
    #return final
        


#data.to_excel('C:\\Users\\phester\\PyWork\\2018rec2.xlsx')



    #tagnames = list(df2)
    
    
    #corr = df2.iloc[:,1:].astype(float).corr()    # Need to convert enteries to floating point numbers to do stats on them


# would like to run through all the stroed files in a folder and then save the description info to a data frame for each param by date (which is in the file name)
#for l in list:



# for testing only
#df  = pd.read_excel('W:\\Process Data\\Daily Reports\\'+'Wood Carbons\\2018'+'\\'+'1. January'+'\\'+'1.10.18.xlsm',skiprows=5)

#df = df[2:]

#df = df.reset_index(drop=True)
def autoEx(years): # feed in the year and a list of the recipes you want
    
    for year in years:
        files = getFiles(year)
        data = historicDesc(files)
        
        data.to_excel('C:\\Users\\phester\\PyWork\\'+str(year)+'ELCD.xlsx')
        
def autoEx2(years):
    data = pd.DataFrame()
    for year in years:
        files = getFiles(year)
        
        data = data.append(historicDesc(files))

    data.to_excel('C:\\Users\\phester\\PyWork\\'+str(''.join(years))+'ELCD.xlsx')

        
