                            ###### PJM Transmission Service Requests Scrape ######

# ### ***BEFORE RUNNING***: COPY THE LAST LINE IN THE PRIOR DAYS TOTALS IN YESTERDAYS  PJM_Trans_Reqts_RAW_DATA.xlsx spreadsheet
# ***NOTE***: You may have to run the first cell of this script a second time (to enable manual login to PJM via your certificate) in order to download the spreadsheet. If needed, check your downloads folder to ensure the date name reflects tomorrows date, if not, run the cell a third time

### PJM Trans Service Requests data
# https://pjmoasis.pjm.com/OASIS/pages/secure/tsr-list.jsf

import requests
import os
import pandas as pd
import glob
import numpy as np
import win32com.client as win32
import time
from os import path
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By

############################################################################################################
username, password = 'Username', 'password' ### Input your PJM username/password to login to PJM
############################################################################################################

### Confingure for automatic login to PJM w/ Selenium
url = f'https://pjmoasis.pjm.com/OASIS/PJM/data/transstatus?TEMPLATE=TRANSSTATUS&OUTPUT_FORMAT=XLSX&PRIMARY_PROVIDER_CODE=PJM&PRIMARY_PROVIDER_DUNS=073647877&RETURN_TZ=EP&VERSION=3.3&POINT_OF_RECEIPT=PJM&TS_WINDOW=FIXED'
response = requests.get(url)  # Open API Web Address in url browser to download api link (this is a direct download)
#webbrowser.open(url)  import webbrowser

driver = webdriver.Edge() # Open an instnace of Edge to facilitate data collection
driver.get("https://pjmoasis.pjm.com/OASIS/PJM/data/transstatus?TEMPLATE=TRANSSTATUS&OUTPUT_FORMAT=XLSX&PRIMARY_PROVIDER_CODE=PJM&PRIMARY_PROVIDER_DUNS=073647877&RETURN_TZ=EP&VERSION=3.3&POINT_OF_RECEIPT=PJM&TS_WINDOW=FIXED")
time.sleep(5) #Wait for Edge to load PJM login page

driver.find_element(By.XPATH, "//*[@id='idToken1']").send_keys(username)
driver.find_element(By.XPATH, "//*[@id='idToken2']").send_keys(password)
driver.find_element(By.XPATH, "//*[@id='loginButton_0']").click() # Login

time.sleep(20) # Wait for excel file to download to give you time to login tp PJM

if response.status_code == 200:
    print("Succesful connection with PJM API.")
    print('-------------------------------')
    
    files = r"C:/Users/*/Downloads/*.xlsx"
    newest = sorted(glob.glob(files), key=os.path.getmtime, reverse=True)
    print(newest[0])   ### This is the file name you are pulling from    

    df = pd.read_excel(newest[0], sheet_name ='TRANSSTATUS')
    
elif response.status_code == 404:
    print("Unable to reach URL.")
else:
    print("Unable to connect API or retrieve data.")   
    
### Clean up and shape data
df = df[['CUSTOMER_CODE','POINT_OF_RECEIPT','POINT_OF_DELIVERY','CAPACITY_GRANTED','SERVICE_INCREMENT','TS_CLASS','START_TIME','STOP_TIME','STATUS','TIME_QUEUED','TS_PERIOD']]
df = df.loc[df['POINT_OF_RECEIPT'] == 'PJM']
Year = datetime.now().strftime("%Y")
Month = datetime.now().strftime("%m")
Day = datetime.now().strftime("%d")#(datetime.now() + timedelta(days=1)).strftime("%d")
HourStart = '00'
HourEnd = '00'
df = df.loc[df['POINT_OF_DELIVERY'].isin(['ALTE','AMIL','CIN','IPL','MEC','MECS','WEC'])]
df = df.loc[df['STATUS'] == 'CONFIRMED']
df['TransservSTART_Year'] = df['START_TIME'].str[:4]
df['TransservSTART_MONTH']= df['START_TIME'].str[4:6]
df['TransservSTART_DAY']= df['START_TIME'].str[6:8]
df['TransservSTART_HOUR']= df['START_TIME'].str[8:10].astype(int)
df['STARTTime_Zone']= df['START_TIME'].str[:-2]
df['TransservSTOP_Year'] = df['START_TIME'].str[:4]
df['TransservSTOP_MONTH']= df['START_TIME'].str[4:6]
df['TransservSTOP_DAY']= df['START_TIME'].str[6:8]
df['TransservSTOP_HOUR']= df['START_TIME'].str[8:10].astype(int)
df['STOPTime_Zone']= df['START_TIME'].str[:-2]

df['START_TIME'] = df.apply(lambda x:datetime.strptime("{0} {1} {2} {3} 00:00".format(x['TransservSTART_Year'], x['TransservSTART_MONTH'], x['TransservSTART_DAY'], x['TransservSTOP_HOUR']),                                                 "%Y %m %d %H %M:%S"),axis=1)

df['STOP_TIME'] = df.apply(lambda x:datetime.strptime("{0} {1} {2} {3} 00:00".format(x['TransservSTOP_Year'], x['TransservSTOP_MONTH'], x['TransservSTOP_DAY'], x['TransservSTOP_HOUR']),                                                 "%Y %m %d %H %M:%S"),axis=1)

### Classify date as Days of the week
df['WEEKEND'] = df['START_TIME'].dt.day_name() ### Classify the day as weekend
df['DAYOFWEEK'] = df['WEEKEND'].apply(lambda x: 'WEEKEND' if x == 'Saturday' or x == 'Sunday' else "WEEKDAY")

# DA = (df['START_TIME'] > f'{Year}-{Month}-{Day}') & (df['STOP_TIME'] <= f'{Year}-{Month}-{Day}')
# df = df.loc[DA]

### Select the columns we want to keep in our dataframe
df = df[['CUSTOMER_CODE','POINT_OF_DELIVERY','CAPACITY_GRANTED','SERVICE_INCREMENT','TS_CLASS', 'START_TIME', \
         'STOP_TIME', 'WEEKEND', 'DAYOFWEEK', 'TransservSTOP_HOUR','STATUS','TS_PERIOD','TransservSTART_HOUR', \
         'TransservSTOP_HOUR']] #,'POINT_OF_RECEIPT'  ,'TIME_QUEUED'  ,'OnOff_Peak'

### Write two df's to excel file
df.to_excel(path.join('T:\\Power Trading\\PJM_Trans_Service_Requests\\', 'PJM_Trans_Reqts_RAW_DATA.xlsx'), sheet_name='TSRs', index=False)
#### path to folder in T drive: path = 'T:\\Power Trading\\PJM_Trans_Service Requests\\PJM_Trans_Reqts.xlsx'

#######################################################################################################################
### Pivot Table output Summary file
df2=df

### Pull only DA Trans data for pivot report
### Only keep DA rows
Year = datetime.now().strftime("%Y")
Month = datetime.now().strftime("%m")
DA = (datetime.now() + timedelta(days=1)).strftime("%d")
ND = (datetime.now() + timedelta(days=2)).strftime("%d")
HourStart = '00'
HourEnd = '00'

start_date = f'{Year}-{Month}-{DA} {HourStart}:00:00'
end_date = f'{Year}-{Month}-{ND} {HourEnd}:00:00'
DayAhead = (df2['START_TIME'] >= start_date) & (df2['STOP_TIME'] <= end_date)
df2 = df2.loc[DayAhead]
df2.head(50)

df2 = np.round(pd.pivot_table(df2, values='CAPACITY_GRANTED', 
                                index=['SERVICE_INCREMENT','TS_CLASS','TS_PERIOD','CUSTOMER_CODE'], 
                                columns=['TransservSTART_HOUR'], 
                                aggfunc=[np.sum],
                                fill_value=0,
                                margins=True, margins_name='Total'),2)

### df2.to_excel(path.join('T:\\Power Trading\\PJM_Trans_Service_Requests\\', 'PJM_Trans_Reqts_DAY_AHEAD.xlsx'), sheet_name= f'DA_Trans_Reqts_{Year}_{Month}_{DA}')

#######################################################################################################################
### Monthly and Weekly Requests Pivot Table

df3 = df.loc[(df['SERVICE_INCREMENT'] == 'MONTHLY') | (df['SERVICE_INCREMENT'] == 'WEEKLY')]
# DayAhead = (df3['START_TIME'] >= start_date)
# df3 = df3.loc[DayAhead]

df3 = np.round(pd.pivot_table(df3, values='CAPACITY_GRANTED',
                                index=['SERVICE_INCREMENT','TS_CLASS','TS_PERIOD','CUSTOMER_CODE'],
                                columns=['TransservSTART_HOUR'],
                                aggfunc=[np.sum],
                                fill_value=0,
                                margins=True, margins_name='Total'),2)

### Write to excel file
#df3.to_excel(path.join('T:\\Power Trading\\PJM_Trans_Service_Requests\\', 'PJM_Trans_Reqts_DAY_AHEAD.xlsx'), sheet_name= f'DA_Trans_Reqts_{Year}_{Month}_{DA}', startrow=len(df2) + 2)

### Write both df's to an xlsx file
with pd.ExcelWriter(path.join('T:\\Power Trading\\PJM_Trans_Service_Requests\\', 'PJM_Trans_Reqts_DAY_AHEAD.xlsx')) as xlsx:
    df2.to_excel(xlsx, sheet_name= f'DA_Trans_Reqts_{Year}_{Month}_{DA}')
    df3.to_excel(xlsx, sheet_name=f'DA_Trans_Reqts_{Year}_{Month}_{DA}', startrow=len(df2) + 6)

#######################################################################################################################
### FINAL VIEW
time.sleep(5) # Wait for file to download and give you time to login
os.startfile('T:\\Power Trading\\PJM_Trans_Service_Requests\\PJM_Trans_Reqts_DAY_AHEAD.xlsx') # Launch Excel to edit

#######################################################################################################################
#######################################################################################################################

### Confirm email of script run and data UPDATE ###
### DO NOT RUN PRIOR TO EDITING SPREADSHEET FOR PROPER FORMAT ###
time.sleep(250) # Wait for file to download and give you time to login
now = datetime.now() # datetime object containing current date and time
dt_string = now.strftime("%d/%m/%Y %H:%M:%S") # Change Date format dd/mm/YY H:M:S

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

# construct email item object
mailItem = olApp.CreateItem(0)
mailItem.Subject = "PJM Transmission Requests Day Ahead data Excel File has been updated"
mailItem.BodyFormat = 1
mailItem.HTMLBody = "Refer to Folder: T:\Power Trading\PJM_Trans_Service_Requests for raw data and pivot table summary. \
                    <br> \
                    <br> \
                    Data updated at 8:30AM, 9:00AM, 9:25AM, and 9:50 Daily \
                    <br> \
                    <br> \
                    ***NOTE***: PJM API data is lagged ~15 minutes"

mailItem.Attachments.Add("T:\Power Trading\PJM_Trans_Service_Requests\PJM_Trans_Reqts_DAY_AHEAD.xlsx")
mailItem.To = "ethan.schultz@conocophillips.com; Ryan.M.Saccone@conocophillips.com; Rick.Burton@conocophillips.com; 24hourdesk@conocophillips.com"
# ***NOTE***: To ensure that you are sending out the right spreadsheet, you can comment out everyone on the "mailItem.To" line except for yourself to test the email
mailItem.Display()
mailItem.Save()
mailItem.Send()   #-> uncomment if you want the email to send automatically
