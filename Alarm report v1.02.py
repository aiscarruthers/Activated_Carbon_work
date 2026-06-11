

import os
import pandas as pd
import plotly_express as px

import win32com.client

from matplotlib import rcParams
rcParams.update({'figure.autolayout': True}) # https://stackoverflow.com/questions/6774086/why-is-my-xlabel-cut-off-in-my-matplotlib-plot/6776578#6776578


os.chdir('C:\\Users\\ACarruther\\bin\\Alarms')  


def cleanFile(fileName):
    
    file = open(fileName,encoding='ISO-8859-1') 
    text = file.readlines() 
    
    fix32 = []
    other = []
    clean = []
    ack = []
    
    # The text file contains data in diffrent formats not all relvant
    for line in text:
        # Copy operator fix32 actions into a seprate list.
        if 'Fix32' in line: 
            fix32.append(line)
        
        # Copy operator acknowledgements into a list, could perhaps throw these
        #  without the fix32s.
        elif 'acknowledged' in line: 
            ack.append(line)
        
        # Copy other operator acknowledgements into a list could perhaps throw 
        # these with the fix32s.
        elif 'NORITSV1::OPERATOR' in line: 
            other.append(line)
        
        elif 'NORITSV2::OPERATOR' in line: 
            other.append(line)
        
        elif 'OW2::BCSADMIN' in line: 
            other.append(line)

        # After the first 5 filters anything with this is the data i want.
        elif '[NORIT   ]' in line: 
            clean.append(line)
        
        # catch the unspecified stuff
        else:
            other.append(line)     
    
    output = open('cleantext.txt','w',encoding='ISO-8859-1')
    clean2 = ''.join(clean) # make it into one long string so you can write to a text file
    
    # Note: when this function is run in bulk there are occasions when it does 
    # not write to the txt file not sure what the root cause for this is.

    output.write(clean2) # do this so i can the read it back as a fixed width file
    df = readclean()
    return df

def readclean():
    
    # The colspec variable determins how the read_fwf() function interprets 
    # the txt file. If we are ever looking for a different format for line 
    # entries this variable will need to change.
    colspec = [(0,21),(33,63),(64,70),(71,85),(86,90),(90,126)]
    df= pd.read_fwf('cleantext.txt', header=None, parse_dates=[[0]], 
                    colspecs=colspec, dayfirst=True, encoding='ISO-8859-1')
    
    columnNames = ['DTime','Tag','Status', 'Value', 'Units', 'Desc'] 
    df.columns = columnNames
    
    return df


# This block grabs the top ten by count saves graph and also returns list of tags
def topTen(cleanDf):
    
    # this groups all the identical tag counts together using the query which also gives unique tag list
    notOkDf = cleanDf.query('Status != "OK"').iloc[:,0:2].groupby(['Tag'],as_index=False).count() 
    tTen = notOkDf.nlargest(10,'DTime').plot.bar(x='Tag',y='DTime',legend=None)
    tTen.set(ylabel='Count')
    tTen.get_figure().savefig('topTen')
     

    return list(notOkDf.nlargest(10,'DTime')['Tag']) #returns list of top ten tags



# this block here is getting the days alarm profile and saving it as a figure 
# doesnt return anything

def dailyT(cleanDf,timeStep='H'): # timeStep changes the sample frequency
    
    # this step is required for other blocks and actually alters the global df
    cleanDf.index=cleanDf.DTime 
    
    # numberT gives intervals of that many mins can change for H to get hours
    resampleCount = cleanDf.query('Status != "OK"').resample(timeStep).count() 
    
    dailyTrends = resampleCount.plot.line(y='Tag',legend=None)
    dailyTrends.set(xlabel='Time',ylabel='Count')
    dailyTrends.get_figure().savefig('dailyTrends')



def timeCheck(T,DF): # Put tag the data frame return describe for it. This works on the dataframe labled df
    #https://stackoverflow.com/questions/16777570/calculate-time-difference-between-pandas-dataframe-indices
    xxx = DF.query('Tag == @T') # the @T allows you to get the variable inside the equality

    xxx2 = xxx.index.to_series().diff().dt.seconds.div(60, fill_value=0)
    
    xxx = xxx.assign(Delta=xxx2)
    
    s1=xxx.query('Status == "OK"')['Delta'].describe() # Time in alarm
    
    s2=xxx.query('Status != "OK"')['Delta'].describe() # Time out of alarm
    
    
    return  [T,{'In':s1,'Out':s2}]

def autoTC(DF): # Iterates through all the tags in a data frame and gets the time stats
    a = DF.groupby('Tag').count() # i just do this to get a singular list of tags
    timeStats={}
    for i in a.index:
        timeStats[i] = timeCheck(i,DF)[1]
    return timeStats

def timeDf(x): # This function just joins all the series data into a data frame
    AlIn = pd.DataFrame()
    AlOut = pd.DataFrame()
    
    for i in x.keys():
        print(i)
        AlIn[i] = x[i]['In'] # Time stats for in alarm
        AlOut[i] = x[i]['Out'] # time stats for out of alarm
        
    return {'AlIn':AlIn,'AlOut':AlOut}


#This block gets the time in alarm stats, saves the grapgs and returns a list containing the tag lists

def longestStats(cleanDf): #need to feed it the clean data frame

    ts = autoTC(cleanDf) # collects all the time stats for each tag
    timeDfs = timeDf(ts) # turns all the time stats into one data frame so i can compare them
    
    qwsd= timeDfs['AlIn'].transpose() # transpose then can use these functions like mormal
    
    fiveLongest = qwsd.nlargest(5,'max').plot.bar(y='max',legend=None) # should be the graph of the top 10 maximum times spent in alarm
    fiveLongest.set(xlabel='Tag',ylabel='Time in alarm (Mins)')
    
    longestAverage = qwsd.nlargest(5,'mean').plot.bar(y='mean',legend=None) #same as above for mean
    longestAverage.set(xlabel='Tag',ylabel='Time in alarm (mins)')
    
    longestAverage.get_figure().savefig('longestAverage')
    tagsLA = list(qwsd.nlargest(5,'mean').index)
    
    fiveLongest.get_figure().savefig('fiveLongest')
    tagsFL = list(qwsd.nlargest(5,'max').index)
    
    return [tagsLA,tagsFL]



def descList(listTags,cleanDf): #Insert a list of tags + dataframe and it will give you a dcitionary of tag:desc
    
    tgs = {}
    listyboi=[]
    for t in listTags:
        # query for tag equal to the one in the for loop, get the unique 
        # descriptions and put into dic
        tgs[t] = list(cleanDf.query('Tag == @t')['Desc'].unique()) 
        
        listyboi.append(str(t)+' '+str(tgs[t])[2:-2])
    
        
        # so far has only retuned indivdual description but should put out a 
        # list if there is more than one descrip, not sure if this will break 
        # something else
    return listyboi



def emailReport(tgs): ##input the list of tag lists
    
    htmlStr ='''<!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><head>  <title></title>  <!--[if !mso]><!-- -->  <meta http-equiv="X-UA-Compatible" content="IE=edge">  <!--<![endif]--><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><style type="text/css">  #outlook a {{ padding: 0; }}  .ReadMsgBody {{ width: 100%; }}  .ExternalClass {{ width: 100%; }}  .ExternalClass * {{ line-height:100%; }}  body {{ margin: 0; padding: 0; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }}  table, td {{ border-collapse:collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; }}  img {{ border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; }}  p {{ display: block; margin: 13px 0; }}</style><!--[if !mso]><!--><style type="text/css">  @media only screen and (max-width:480px) {{    @-ms-viewport {{ width:320px; }}    @viewport {{ width:320px; }}  }}</style><!--<![endif]--><!--[if mso]><xml>  <o:OfficeDocumentSettings>    <o:AllowPNG/>    <o:PixelsPerInch>96</o:PixelsPerInch>  </o:OfficeDocumentSettings></xml><![endif]--><!--[if lte mso 11]><style type="text/css">  .outlook-group-fix {{    width:100% !important;  }}</style><![endif]--><!--[if !mso]><!-->    <link href="https://fonts.googleapis.com/css?family=Lato" rel="stylesheet" type="text/css"><link href="https://fonts.googleapis.com/css?family=Ubuntu:300,400,500,700" rel="stylesheet" type="text/css">    <style type="text/css">        @import url(https://fonts.googleapis.com/css?family=Lato);  @import url(https://fonts.googleapis.com/css?family=Ubuntu:300,400,500,700);    </style>  <!--<![endif]--><style type="text/css">  @media only screen and (min-width:480px) {{    .mj-column-per-100 {{ width:100%!important; }}.mj-column-per-50 {{ width:50%!important; }}  }}</style></head><body style="background: #FFFFFF;">    <div class="mj-container" style="background-color:#FFFFFF;"><!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="600" align="center" style="width:600px;">        <tr>          <td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;">      <![endif]--><table role="presentation" cellpadding="0" cellspacing="0" style="background:#49a6e8;font-size:0px;width:100%;" border="0"><tbody><tr><td><div style="margin:0px auto;max-width:600px;"><table role="presentation" cellpadding="0" cellspacing="0" style="font-size:0px;width:100%;" align="center" border="0"><tbody><tr><td style="text-align:center;vertical-align:top;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;"><!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0">        <tr>          <td style="vertical-align:top;width:600px;">      <![endif]--><div class="mj-column-per-100 outlook-group-fix" style="vertical-align:top;display:inline-block;direction:ltr;font-size:13px;text-align:left;width:100%;"><table role="presentation" cellpadding="0" cellspacing="0" style="vertical-align:top;" width="100%" border="0"><tbody><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 20px 0px 20px;" align="center"><div style="cursor:auto;color:#FFFFFF;font-family:Lato, Tahoma, sans-serif;font-size:14px;line-height:22px;text-align:center;"><h1 style="font-family: &apos;Cabin&apos;, sans-serif; color: #FFFFFF; font-size: 32px; line-height: 100%;">Alarm Managment Report</h1></div></td></tr><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 0px 0px 0px;" align="center"><table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-spacing:0px;" align="center" border="0"><tbody><tr><td style="width:486px;">
    <img alt height="auto" src="cid:dailyTrends.png" style="border:none;border-radius:0px;display:block;font-size:13px;outline:none;text-decoration:none;width:100%;height:auto;" width="486"></td></tr></tbody></table></td></tr><tr><td style="word-wrap:break-word;font-size:0px;padding:15px 15px 15px 15px;" align="center"><div style="cursor:auto;color:#000000;font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:11px;line-height:1.5;text-align:center;"><p><span style="color:#ffffff;"><strong><span style="font-size:14px;">Frequency over the last week</span></strong></span></p></div></td></tr><tr><td style="word-wrap:break-word;font-size:0px;"><div style="font-size:1px;line-height:50px;white-space:nowrap;">&#xA0;</div></td></tr></tbody></table></div><!--[if mso | IE]>      </td></tr></table>      <![endif]--></td></tr></tbody></table></div></td></tr></tbody></table><!--[if mso | IE]>      </td></tr></table>      <![endif]-->      <!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="600" align="center" style="width:600px;">        <tr>          <td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;">      <![endif]--><table role="presentation" cellpadding="0" cellspacing="0" style="font-size:0px;width:100%;" border="0"><tbody><tr><td><div style="margin:0px auto;max-width:600px;"><table role="presentation" cellpadding="0" cellspacing="0" style="font-size:0px;width:100%;" align="center" border="0"><tbody><tr><td style="text-align:center;vertical-align:top;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;"><!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0">        <tr>          <td style="vertical-align:top;width:600px;">      <![endif]--><div class="mj-column-per-100 outlook-group-fix" style="vertical-align:top;display:inline-block;direction:ltr;font-size:13px;text-align:left;width:100%;"><table role="presentation" cellpadding="0" cellspacing="0" style="vertical-align:top;" width="100%" border="0"><tbody><tr><td style="word-wrap:break-word;font-size:0px;padding:15px 15px 15px 15px;" align="center"><div style="cursor:auto;color:#000000;font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:11px;line-height:1.5;text-align:center;"><p><span style="font-size:14px;">Ten most frequent alarms for the last week</span></p></div></td></tr><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 0px 0px 0px;" align="center"><table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-spacing:0px;" align="center" border="0"><tbody><tr><td style="width:486px;">
        <img alt height="auto" src="cid:topTen.png" style="border:none;border-radius:0px;display:block;font-size:13px;outline:none;text-decoration:none;width:100%;height:auto;" width="486"></td></tr></tbody></table></td></tr></tbody></table></div><!--[if mso | IE]>      </td></tr></table>      <![endif]--></td></tr></tbody></table></div></td></tr></tbody></table><!--[if mso | IE]>      </td></tr></table>      <![endif]-->      <!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="600" align="center" style="width:600px;">        <tr>          <td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;">      <![endif]--><table role="presentation" cellpadding="0" cellspacing="0" style="font-size:0px;width:100%;" border="0"><tbody><tr><td><div style="margin:0px auto;max-width:600px;"><table role="presentation" cellpadding="0" cellspacing="0" style="font-size:0px;width:100%;" align="center" border="0"><tbody><tr><td style="text-align:center;vertical-align:top;direction:ltr;font-size:0px;padding:9px 0px 9px 0px;"><!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0">        <tr>          <td style="vertical-align:top;width:300px;">      <![endif]--><div class="mj-column-per-50 outlook-group-fix" style="vertical-align:top;display:inline-block;direction:ltr;font-size:13px;text-align:left;width:100%;"><table role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0"><tbody><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 0px 0px 0px;" align="center"><div style="cursor:auto;color:#000000;font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:11px;line-height:1.5;text-align:center;">
            <p><span style="font-size:14px;">|{TT1}|</span></p>
            <p><span style="font-size:14px;">|{TT2}|</span></p>
            <p><span style="font-size:14px;">|{TT3}|</span></p>
            <p><span style="font-size:14px;">|{TT4}|</span></p>
            <p><span style="font-size:14px;">|{TT5}|</span></p></div></td></tr></tbody></table></div><!--[if mso | IE]>      </td><td style="vertical-align:top;width:300px;">      <![endif]--><div class="mj-column-per-50 outlook-group-fix" style="vertical-align:top;display:inline-block;direction:ltr;font-size:13px;text-align:left;width:100%;"><table role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0"><tbody><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 0px 0px 0px;" align="center"><div style="cursor:auto;color:#000000;font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:11px;line-height:1.5;text-align:center;">
                <p><span style="font-size:14px;">|{TT6}|</span></p><p><span style="font-size:14px;">|{TT7}|</span></p>
                <p><span style="font-size:14px;">|{TT8}|</span></p>
                <p><span style="font-size:14px;">|{TT9}|</span></p>
                <p><span style="font-size:14px;">|{TT10}|</span></p></div></td></tr></tbody></table></div><!--[if mso | IE]>      </td></tr></table>      <![endif]--></td></tr></tbody></table></div></td></tr></tbody></table><!--[if mso | IE]>      </td></tr></table>      <![endif]-->      <!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="600" align="center" style="width:600px;">        <tr>          <td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;">      <![endif]--><table role="presentation" cellpadding="0" cellspacing="0" style="background:#59c4f4;font-size:0px;width:100%;" border="0"><tbody><tr><td><div style="margin:0px auto;max-width:600px;"><table role="presentation" cellpadding="0" cellspacing="0" style="font-size:0px;width:100%;" align="center" border="0"><tbody><tr><td style="text-align:center;vertical-align:top;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;"><!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0">        <tr>          <td style="vertical-align:top;width:600px;">      <![endif]--><div class="mj-column-per-100 outlook-group-fix" style="vertical-align:top;display:inline-block;direction:ltr;font-size:13px;text-align:left;width:100%;"><table role="presentation" cellpadding="0" cellspacing="0" style="vertical-align:top;" width="100%" border="0"><tbody><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 0px 0px 0px;" align="center"><div style="cursor:auto;color:#000000;font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:11px;line-height:1.5;text-align:center;"><p><span style="color:#ffffff;"><span style="font-size:14px;"><strong>&#xA0;Top five for maximum alarm duration</strong></span></span></p></div></td></tr><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 0px 0px 0px;" align="center"><table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-spacing:0px;" align="center" border="0"><tbody><tr><td style="width:486px;">
                    <img alt height="auto" src="cid:fiveLongest.png" style="border:none;border-radius:0px;display:block;font-size:13px;outline:none;text-decoration:none;width:100%;height:auto;" width="486"></td></tr></tbody></table></td></tr><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 0px 0px 0px;" align="center"><div style="cursor:auto;color:#000000;font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:11px;line-height:1.5;text-align:center;">
                        <p><span style="color:#ffffff;"><span style="font-size:14px;">|{TM1}|</span></span></p>
                        <p><span style="color:#ffffff;"><span style="font-size:14px;">|{TM2}|</span></span></p>
                        <p><span style="color:#ffffff;"><span style="font-size:14px;">|{TM3}|</span></span></p>
                        <p><span style="color:#ffffff;"><span style="font-size:14px;">|{TM4}|</span></span></p>
                        <p><span style="color:#ffffff;"><span style="font-size:14px;">|{TM5}|</span></span></p></div></td></tr></tbody></table></div><!--[if mso | IE]>      </td></tr></table>      <![endif]--></td></tr></tbody></table></div></td></tr></tbody></table><!--[if mso | IE]>      </td></tr></table>      <![endif]-->      <!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="600" align="center" style="width:600px;">        <tr>          <td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;">      <![endif]--><table role="presentation" cellpadding="0" cellspacing="0" style="font-size:0px;width:100%;" border="0"><tbody><tr><td><div style="margin:0px auto;max-width:600px;"><table role="presentation" cellpadding="0" cellspacing="0" style="font-size:0px;width:100%;" align="center" border="0"><tbody><tr><td style="text-align:center;vertical-align:top;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;"><!--[if mso | IE]>      <table role="presentation" border="0" cellpadding="0" cellspacing="0">        <tr>          <td style="vertical-align:top;width:600px;">      <![endif]--><div class="mj-column-per-100 outlook-group-fix" style="vertical-align:top;display:inline-block;direction:ltr;font-size:13px;text-align:left;width:100%;"><table role="presentation" cellpadding="0" cellspacing="0" style="vertical-align:top;" width="100%" border="0"><tbody><tr><td style="word-wrap:break-word;font-size:0px;padding:15px 15px 15px 15px;" align="center"><div style="cursor:auto;color:#000000;font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:11px;line-height:1.5;text-align:center;"><p><span style="font-size:14px;">Top five average time in alarm</span></p></div></td></tr><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 0px 0px 0px;" align="center"><table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-spacing:0px;" align="center" border="0"><tbody><tr><td style="width:486px;">
                            <img alt height="auto" src="cid:longestAverage.png" style="border:none;border-radius:0px;display:block;font-size:13px;outline:none;text-decoration:none;width:100%;height:auto;" width="486"></td></tr></tbody></table></td></tr><tr><td style="word-wrap:break-word;font-size:0px;padding:0px 0px 0px 0px;" align="center"><div style="cursor:auto;color:#000000;font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:11px;line-height:1.5;text-align:center;">
                                <p><span style="font-size:14px;">|{TA1}|</span></p>
                                <p><span style="font-size:14px;">|{TA2}|</span></p>
                                <p><span style="font-size:14px;">|{TA3}|</span></p>
                                <p><span style="font-size:14px;">|{TA4}|</span></p>
                                <p><span style="font-size:14px;">|{TA5}|</span></p></div></td></tr></tbody></table></div><!--[if mso | IE]>      </td></tr></table>      <![endif]--></td></tr></tbody></table></div></td></tr></tbody></table><!--[if mso | IE]>      </td></tr></table>      <![endif]--></div></body></html>'''.format(TT1=tgs[2][0],TT2=tgs[2][1],TT3=tgs[2][2],TT4=tgs[2][3],TT5=tgs[2][4],TT6=tgs[2][5],TT7=tgs[2][6],TT8=tgs[2][7],TT9=tgs[2][8],TT10=tgs[2][9],TM1=tgs[0][0],TM2=tgs[0][1],TM3=tgs[0][2],TM4=tgs[0][3],TM5=tgs[0][4],TA1=tgs[1][0],TA2=tgs[1][1],TA3=tgs[1][2],TA4=tgs[1][3],TA5=tgs[1][4])
                         
    return htmlStr     

def sendMail(html,fp='C:\\Users\\ACarruther\\bin\\Alarms'):

    o = win32com.client.Dispatch("Outlook.Application")
    
    Msg = o.CreateItem(0)
    Msg.Importance = 1
    Msg.Subject = 'Alarm Managment Report'
    Msg.Attachments.Add(fp+'\\dailyTrends.png')
    Msg.Attachments.Add(fp+'\\topTen.png')
    Msg.Attachments.Add(fp+'\\fiveLongest.png')
    Msg.Attachments.Add(fp+'\\longestAverage.png')
    Msg.Attachments.Add(fp+'\\cor.html')
    Msg.Attachments.Add(fp+'\\toptagsline.html')
    
    
    Msg.HTMLBody = html
    
    #Msg.To = 'andrew.carruthers@norit.com'
    #Msg.To = 'andrew.carruthers@norit.com;Alan.Thompson@norit.com'
    #Msg.To = 'andrew.carruthers@norit.com;Mairi.NicolWood@norit.com;Alan.Thompson@norit.com'
    #Msg.CC = STRING_CONTAINING_CC
    #Msg.BCC = STRING_CONTAINING_BCC
    
    #Msg.SentOnBehalfOfName = "ANOTHER_MAIL_BOX@DOMAIN.COM"
    #Msg.ReadReceiptRequested = True
    #Msg.OriginatorDeliveryReportRequested = True
    
    # Msg.Send()  
    
    Msg.Display()                    
    
    
def mainLoop(fileName):    
    
     
    df = cleanFile(fileName) #clean the file return a dataframe


    tagsTT = topTen(df) #Top ten stats

    dailyT(df) # daily alarm trends

    longTags = longestStats(df) 
  
    tagLists = [descList(longTags[0],df),descList(longTags[1],df),descList(tagsTT,df)] 
    # Get a list of dictionaries the tags are the keys and descriptions to go 
    # with the respective graphs
    # longest alarm stats 0 is longest ave 1 is max 2 is most frequent
  ################################################################## end of main loop
    emailHtml= emailReport(tagLists) 
     
    sendMail(emailHtml)

#mainLoop('180514.ALM') # this is just a place holder for now just to run the thing, i think i would call this on the file name and it would also include the email part

def findFiles(filePath='C:\\Users\\ACarruther\\bin\\Alarms\\Week'):
    
    os.chdir(filePath)
    files = os.listdir()
    dfs = []
    
    for file in files:
        if('.ALM' in file):    
            print(file)
            cleanDf = cleanFile(file)
            dfs.append(cleanDf) # this doesnt work properly as i need to add a date back in against them
    
    bigDf = pd.concat(dfs)
    
    return bigDf,filePath
        
def week():    
    ### This is the main report
     
    df,filePath = findFiles()
    tagsTT = topTen(df) #Top ten stats
    dailyT(df) # daily alarm trends

    longTags = longestStats(df) 
  
    tagLists = [descList(longTags[0],df),descList(longTags[1],df),descList(tagsTT,df)] 
    
    tagTraceList()

    emailHtml= emailReport(tagLists) 
     
    sendMail(emailHtml,fp=filePath)          
    
    
def longTerm():
    
    #This is to look at alarm frequeny over a long period broken into daily average
    
    fP='C:\\Users\\ACarruther\\bin\\Alarms\\Alarm_Files'
    
    df,z = findFiles(filePath=fP)
    
    dailyT(df,timeStep='D')
    
    resampleCount = df.query('Status != "OK"').resample('D').count()
    
    return resampleCount#.describe()
    


    
def tagTrace():
    
    #Any graphs from this currently overwrite other graphs!
    
    test = cleanFile('180515.ALM')
    
    # I used this to get the list of tags. just need to feed it a cleanfile DF
    testTags = topTen(test) 
    
    test.index=test.DTime
    
    # need to load these tags from a variable probably in a list form
    tag= test.query('Status !="OK"& Tag == "290504DEVAL"').resample('H').count() 
    
    dailyTrends = tag.plot.line(y='Tag',legend=None)
    dailyTrends.set(xlabel='Time',ylabel='Count')
    dailyTrends.get_figure().savefig('dailyTrends')
    
def tagTraceList():
    
    #correlation probally needs a resample time of a min to really mean anything
    
    cleanData = findFiles()[0]
    
    topTags = topTen(cleanData)
    
    cleanData.index=cleanData.DTime
    
    Dfs=[]
    
    for t in topTags:    
        
        dataFrame= cleanData.query('Status !="OK"& Tag == @t').resample('H').count() #do the resample on each tag
        
        series = dataFrame.DTime #drop all but one line
        
        Dfs.append(series) # add the result to the loop list
    
    topTagsDf=pd.concat(Dfs,axis=1) #concate the series
    
    topTagsDf.columns = topTags # name the tags correctly in the df
    
    topTagsDf.copy().iloc[:,0:6]
    
    cor=topTagsDf.corr() #correlation coeffs for the tags, could try the diffrent methods tho    
    
    g1= px.imshow(cor,color_continuous_scale='PRGn',color_continuous_midpoint=0)
    
    # g1.show()
    
    g1.write_html('C:\\Users\\ACarruther\\bin\\Alarms\\Week\\cor.html')
    
    g2=px.line(topTagsDf)
    
    g2.write_html('C:\\Users\\ACarruther\\bin\\Alarms\\Week\\toptagsline.html')
    
    descs=descList(topTags,cleanData)
    
    print(descs)
    
    

