# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import re
import os

import pandas as pd



def getFiles(a):
    
    
    mode = 'Wood Carbons'+'\\' #select mode by comenting out the one you dont want although im not looking at elcd yet

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


def historicDesc(dic,rec):
    
    final = [] # place to put results
    
    for z in dic: # loop through the months
        
        #try:
            
        
            for i in dic[z]: # loop through each days files
                print(i)    
            
                regex = re.compile(r'(\d+).(\d+).(\d\d\d\d|\d\d)')
            
                year = regex.search(str(i))[3]
            
                if int(year) <= 1700:
                    year = int(year) + 2000
            
                mode = 'Wood Carbons'+'\\' #select mode by comenting out the one you dont want although im not looking at elcd yet
                
                df  = pd.read_excel('W:\\Process Data\\Daily Reports\\'+mode+str(year)+'\\'+z+'\\'+i,skiprows=5,usecols='A:BE')

                df = df[2:]

                df = df.reset_index(drop=True)
                #print(df) # this is just for trouble shooting

# might want it to filter for full days with no downtime events or two kiln operation
                
                df2 = df.loc[df['Kiln C Running A03'] >= 1  ]
                #print(df2)#troubleshooting only  
# The next bit give sstats 
                df2 = df2.loc[df['Plant Recipe Number'] == rec  ] 
                
                #print(df2) #troubleshooting only  
                
                desc = df2.iloc[:,1:].astype(float).describe().transpose() # Just need to add a date colum to this then poulate from the read in string
                desc = desc.assign(Date=regex.search(str(i))[1]+'.'+regex.search(str(i))[2]+'.'+str(year))
                
                desc['Date'] = pd.to_datetime(desc['Date'])
            #Need to work out somethig to join the dataframes together into one big table then maybe save to excel
                #print(desc) #troubleshooting only            
            
                final.append(desc)
        #except:
           # continue
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
def autoEx(year,reclist): # feed in the year and a list of the recipes you want
    files = getFiles(year)
    for rec in reclist:
        data = historicDesc(files,rec)
        
        data.to_excel('C:\\Users\\phester\\PyWork\\'+str(year)+'rec'+str(rec)+'.xlsx')
        
        
