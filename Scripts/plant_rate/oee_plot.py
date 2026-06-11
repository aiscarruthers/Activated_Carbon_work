import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime

import matplotlib
matplotlib.use("QT5agg")
matplotlib.rcParams['font.family'] = ['Calibri']
df = pd.read_csv("C:\\Users\\ACarruther\\OneDrive - cabotcorp.com\\Documents\\Campaign Analysis\\Nov-2022_Evidence\\OEE_data.csv")
pd.to_datetime(df['Datetime'])

def make_subplots(Dataframe):
    f, ((fpq, fpq_), (osf, osf_), (pri, pri_)) = plt.subplots(nrows=3, ncols=2, width_ratios=[5,1], squeeze=True)
    
    
    oeel = fpq.plot(Dataframe['Datetime'], Dataframe['OEE'], color='red', label='OEE' )
    fpql = fpq.plot(Dataframe['Datetime'], Dataframe['FPQ'], color='magenta', label='FPQ' )
    pril = pri.plot(Dataframe['Datetime'], Dataframe['PRI'], color= 'purple', label='PRI' )
    osfl = osf.plot(Dataframe['Datetime'], Dataframe['GOSF'], color='orange', label='OSF' )
    fpq.set_xticklabels([])
    # fpq.legend(loc='center', bbox_to_anchor=(0.5, 1.3), shadow=False, ncol=2)
    
    osf.set_xticklabels([])
    # pri.legend(loc='center', bbox_to_anchor=(0.5, 1.3), shadow=False, ncol=2)
  
    
    pri.plot(Dataframe['Datetime'], Dataframe['OEE'], color='red', label='_OEE')
    osf.plot(Dataframe['Datetime'], Dataframe['OEE'], color='red', label='_OEE')
    # osf.legend(loc='center', bbox_to_anchor=(0.5, 1.3), shadow=False, ncol=2)
    
    pri_.scatter(Dataframe['PRI'], Dataframe['OEE'], color='purple', label="_PRI",zorder=3)
    fpq_.scatter(Dataframe['FPQ'], Dataframe['OEE'], color='magenta', label="_FPQ", zorder=3)
    osf_.scatter(Dataframe['GOSF'], Dataframe['OEE'], color='orange', label="_GOSF",zorder=3)
    
    fpq.grid()
    fpq.margins(x=0)
    fpq_.grid(zorder=0)
    fpq_.set_title(' ',fontdict={'fontsize':18})
    
    osf.grid()
    osf.margins(x=0)
    osf_.grid(zorder=0)
    #osf_.set_title(' ',fontdict={'fontsize':18})
    pri.grid()
    pri.margins(x=0)
    pri_.grid(zorder=0)
    #pri_.set_title(' ',fontdict={'fontsize':18})

   

    f.legend(loc='center', bbox_to_anchor=(0.5, 0.978), shadow=False, ncol=4, fontsize = 18)
    plt.sca(pri)
    plt.xticks(rotation=45, rotation_mode='anchor', ha='right',fontsize=11)
    f.set_figheight(6)
    f.set_figwidth(16)
    f.tight_layout()
    f.savefig('C:\\Users\\ACarruther\\OEE_summary.png', transparent=True)

make_subplots(df)
