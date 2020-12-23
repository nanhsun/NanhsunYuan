import xlwings as xw
import win32com.client as win32
import numpy as np
import datetime
import os.path
import pandas as pd 
# 若檔案夾位置有更換，請改每個r""裡的檔案位置到新的檔案位置
def nparray_save_file(var):
    """
    Input: var --- nparray for saving
    Output: none
    """
    np.savetxt(r'\\TAFS\Public\Alan Yuan\Reflow\LogTime.csv', var, fmt = "%s", delimiter = ",", encoding = "utf-8-sig")
    return

def open_csv_file():
    """
    Input: accum --- integer to determine whether for accumulation or not
    Output: df --- dataframe from file
    """
    df = np.genfromtxt(r'\\TAFS\Public\Alan Yuan\Reflow\LogTime.csv',delimiter = ',',dtype = None,encoding = 'utf-8')

    return df

def open_xlsx_file(fileloc):
    """
    Input: none
    Output: df --- dataframe from RawData file
    """
    df = pd.read_excel(fileloc,dtype = 'str',encoding = 'utf-8')
    return df

def LightArray(timeframe):
    SRH01 = []
    SRR01 = []
    if timeframe <= 12:
        for i in range(timeframe):
            SRH01.append(chr(79+i)+'174')
            SRR01.append(chr(79+i)+'175')
    else:
        for i in range(12):
            SRH01.append(chr(79+i)+'174')
            SRR01.append(chr(79+i)+'175')
        for i in range(int(np.floor((timeframe - 12)/26))):
            for j in range(26):
                SRH01.append(chr(65+i) + chr(65+j)+'174')
                SRR01.append(chr(65+i) + chr(65+j)+'175')
        for j in range(int(timeframe - 12 - 26*np.floor((timeframe - 12)/26))):
            SRH01.append(chr(65+int(np.floor((timeframe - 12)/26))) + chr(65+j)+'174')
            SRR01.append(chr(65+int(np.floor((timeframe - 12)/26))) + chr(65+j)+'175')
    return SRH01,SRR01

def LightCount(SRH01,SRR01,df):
    count = 0
    for j in range(len(SRH01)):
        if df.sheets[0].range(SRH01[j]).color == (255,255,0) and df.sheets[0].range(SRR01[j]).color == (0,255,0):
            count = count + 1
    return count

def MailAlert(mails,Sub,Body):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mails
    mail.Subject = Sub
    mail.Body = Body
    mail.Send()
    return

if __name__ == "__main__":

    start = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    while os.path.exists(r'\\TAFS\Public\Alan Yuan\Reflow\LogTime.csv') == False:
        timelog = np.array([['Start','Finish Opening File','While func Start','While func End/Con func Start','Con func End/Email Start','Email func End','End']])
        nparray_save_file(timelog)
    
    timelog2 = open_csv_file()
    app = xw.App()
    print('Starting reflow machine detection')
    df = xw.Book(r'\\Cna0133566\oeu\OEU-FPC240.xlsm')
    openfile = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    df2 = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Reflow\Reflow File.xlsx')

    timeframe = int(np.round(pd.to_numeric(df2['Time (min)'][0])/4))
    timeframe2 = int(np.round(pd.to_numeric(df2['Time (min)'][1])/4))

    whilefuncstart = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    SRH01,SRR01 = LightArray(timeframe)
    SRH01_2,SRR01_2 = LightArray(timeframe2)

    count = LightCount(SRH01,SRR01,df)
    count2 = LightCount(SRH01_2,SRR01_2,df)
    whilefuncend = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    forfuncend = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    mails1 = str()
    mails2 = str()
    
    for i in range(df2['Emails for First Time'].dropna().shape[0]):
        emails = df2['Emails for First Time'][i] + ';'
        mails1 = mails1 + emails

    for i in range(df2['Emails for Second Time'].dropna().shape[0]):
        emails = df2['Emails for Second Time'][i] + ';'
        mails2 = mails2 + emails

    outlook1 = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root_folder = outlook1.Folders.Item(1)
    sent_items = root_folder.Folders['Sent Items']
    messages = sent_items.Items
    message = messages.GetLast()
    date = message.SentOn.strftime("%Y-%m-%d %H:%M:%S")
    
    if count == timeframe:
        if (message.Subject == df2['First Alert Subject'][0]) and date >= (datetime.datetime.now() - datetime.timedelta(minutes = int(pd.to_numeric(df2['First Alert Time (min)'][0])))).strftime("%Y-%m-%d %H:%M"):
            pass
        else:
            MailAlert(mails1,df2['First Alert Subject'][0],df2['First Alert Body'][0])

    if count2 == timeframe2:
        if (message.Subject == df2['Second Alert Subject'][0]) and date >= (datetime.datetime.now() - datetime.timedelta(minutes = int(pd.to_numeric(df2['Second Alert Time (min)'][0])))).strftime("%Y-%m-%d %H:%M"):
            pass
        else:
            MailAlert(mails2,df2['Second Alert Subject'][0],df2['Second Alert Body'][0])
        

    emailend = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    df.close()
    app.quit()
    print('Ending reflow machine detection')
    endtime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    times = np.array([[start,openfile,whilefuncstart,whilefuncend,forfuncend,emailend,endtime]])
    output = np.vstack((timelog2,times))
    nparray_save_file(output)
    app.kill()    