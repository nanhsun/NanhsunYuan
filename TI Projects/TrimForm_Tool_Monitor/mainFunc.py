import numpy as np
import pandas as pd
import cx_Oracle as cxO
import os.path
import matplotlib.pyplot as plt
import datetime
import matplotlib.dates as mdates
import win32com.client as win32
def open_htm_file(fileloc):
    """
    Input: none
    Output: df --- dataframe from file
    """
    with open(fileloc,'r') as f:
        dfs = pd.read_html(f.read())
    df = dfs[0]
    return df

def open_xlsx_file(fileloc,head = 0):
    """
    opens xlsx files
    Input: none
    Output: df --- dataframe from RawData file
    """
    if head == 0:
        df = pd.read_excel(fileloc,dtype = 'str',encoding = 'utf-8')
    elif head == 1:
        df = pd.read_excel(fileloc,dtype = 'str',encoding = 'utf-8', header=[0,1,2])
    return df

def dataframe_save_file(fileloc,var,parts = False):
    """
    save dataframe
    Input: fileloc --- file location
           var --- dataframe for saving
    Output: none
    """
    if parts == False:
        var.to_excel(fileloc,index = False, encoding = "utf-8-sig")
    elif parts == True:
        var.to_excel(fileloc, encoding = "utf-8-sig")
    return

def nparray_save_file(fileloc,var):
    """
    save numpy arrays
    Input: var --- nparray for saving
    Output: none
    """
    np.savetxt(fileloc, var, fmt = "%s", delimiter = ",", encoding = "utf-8-sig")
    return

def accumulate_file(csv,htm):
    """
    accumulate/update data
    Input: csv --- file to be updated (in .csv)
           htm --- update file (in .htm)
    Output: accumulated dataframe
    """
    htm.to_excel(r'\\TAFS\Public\Alan Yuan\Temp.xlsx',index = False,encoding = "utf-8-sig")
    htm2 = pd.read_excel(r'\\TAFS\Public\Alan Yuan\Temp.xlsx',dtype = 'str',encoding = 'utf-8')
    htm2.columns = htm2.iloc[0]
    htm2 = htm2.drop(htm2.index[0])
    htm2['Date'] = pd.to_datetime(htm2['Date'])
    htm2.to_excel(r'\\TAFS\Public\Alan Yuan\Temp.xlsx',index = False,encoding = "utf-8-sig")
    htm3 = pd.read_excel(r'\\TAFS\Public\Alan Yuan\Temp.xlsx',dtype = 'str',encoding = 'utf-8')
    csv_drop = csv.drop(columns = ["Tooling"])
    df = pd.concat([htm3,csv_drop])
    df.drop_duplicates(subset=None, keep="first", inplace=True)
    tooling = csv["Tooling"]
    df_final = pd.concat([df,tooling],axis = 1)
    num = df.shape[0] - csv.shape[0]
    df_final["Tooling"] = df_final["Tooling"].shift(periods = num)
    return df_final

def find_unique_max(df,fileloc):
    """
    find ma values of each timeframe
    Input: df --- dataframe for finding max value of each timeframe
    Output: max_value --- np.array that stores results
    """
    if fileloc == r'\\TAFS\Public\Alan Yuan\SO_TSSOP Files\AccumRawDataWithTooling.xlsx':
        max_value = np.array([['Date','Product','Machine#','Device','Lot','P/P','Tooling','Tip to Tip','Interlead Flash','Seating Height','Foot Angle']])
        dates_ = pd.unique(df['Date'].dropna())
        for date_ in dates_:
            date_df = df[df['Date'] == date_]
            mfw = np.array([[date_df.iat[0,0],date_df.iat[0,1],date_df.iat[0,2],date_df.iat[0,3],date_df.iat[0,4],date_df.iat[0,5],date_df.iat[0,16],pd.to_numeric(date_df.iloc[:,9]).dropna().max(),pd.to_numeric(date_df.iloc[:,10]).dropna().max(),pd.to_numeric(date_df.iloc[:,11]).dropna().max(),pd.to_numeric(date_df.iloc[:,12]).dropna().max()]])
            max_value = np.vstack((max_value,mfw))    
    elif fileloc == r'\\TAFS\Public\Alan Yuan\QFP Files\AccumRawDataWithTooling.xlsx':
        max_value = np.array([['Date','Product','Machine#','Device','Lot','P/P','Tooling','Tip to Tip X','Tip to Tip Y','Seating Height X','Foot Angle X','Lead Length X','Seating Height Y','Foot Angle Y','Lead Length Y']])
        dates_ = pd.unique(df['Date'].dropna())
        for date_ in dates_:
            date_df = df[df['Date'] == date_]
            mfw = np.array([[date_df.iat[0,0],date_df.iat[0,1],date_df.iat[0,2],date_df.iat[0,3],date_df.iat[0,4],date_df.iat[0,5],date_df.iat[0,20],pd.to_numeric(date_df.iloc[:,9]).dropna().max(),pd.to_numeric(date_df.iloc[:,10]).dropna().max(),pd.to_numeric(date_df.iloc[:,11]).dropna().max(),pd.to_numeric(date_df.iloc[:,12]).dropna().max(),pd.to_numeric(date_df.iloc[:,13]).dropna().max(),pd.to_numeric(date_df.iloc[:,14]).dropna().max(),pd.to_numeric(date_df.iloc[:,15]).dropna().max(),pd.to_numeric(date_df.iloc[:,16]).dropna().max()]])
            max_value = np.vstack((max_value,mfw))  
    return max_value

def SQLQuery(days):
    """
    SQL
    Input: days
    Output: dataframe extracted from SQL
    """
    print("Begin SQL query")
    connection = cxO.connect('SQL SERVER NAME HERE','SQL SERVER NAME HERE','PW HERE')
    query = '''
    Select lot, char_value, logout_dttm from lot_parm
    where facility in ('TAI')
    and lpt in ('6010','6100')
    and parm in ('1296')
    and logout_dttm > sysdate -''' + days
    print("Starting Query")
    df_ora = pd.read_sql(query, con=connection)
    print("Query finished")
    connection.close()
    return df_ora

def InsertTooling(df_ora,df_htm):
    """
    insert tooling number to dataframe
    Input: df_ora --- dataframe from SQL
           df_htm --- dataframe from htm
    Output: df_htm --- dataframe with additional Tooling column
    """
    lots_ = pd.unique(df_ora["LOT"])
    lots_htm = pd.unique(df_htm["Lot"])
    for lot in lots_:
        if lot in lots_htm:
            df = df_ora[df_ora["LOT"] == lot]
            df_htm.loc[df_htm["Lot"] == lot,"Tooling"]= df["CHAR_VALUE"].iloc[0]
    return df_htm

def UpdateFile(fileloc):
    '''
    wrapper func around accumulate_file(), SQLQuery(), and InsertTooling()
    Input: file location string
    Output: None
    '''
    firsttime = False
    if fileloc == r'\\TAFS\Public\Alan Yuan\SO_TSSOP Files\AccumRawDataWithTooling.xlsx':
        print('Extracting data from \\tafs\PQC\Assy\Dimension\SO_TSSOP_New.htm')
        df_raw_htm = open_htm_file(r"\\tafs\PQC\Assy\Dimension\SO_TSSOP_New.htm")
    elif fileloc == r'\\TAFS\Public\Alan Yuan\QFP Files\AccumRawDataWithTooling.xlsx':
        print('Extracting data from \\tafs\PQC\Assy\Dimension\QFP.htm')
        df_raw_htm = open_htm_file(r"\\tafs\PQC\Assy\Dimension\QFP.htm")

    while os.path.exists(fileloc) == False:
        print('Saving data to \\TAFS\Public\Alan Yuan\SO_TSSOP Files or \\TAFS\Public\Alan Yuan\QFP Files')
        dataframe_save_file(fileloc,df_raw_htm)
        print('File Saved')
        firsttime = True

    if firsttime == False:
        print('Start accumulating data')
        df_raw = open_xlsx_file(fileloc)
        df_araw = accumulate_file(df_raw,df_raw_htm)
        print('Accumulation complete')
    elif firsttime == True:
        df_araw = open_xlsx_file(fileloc)

    print('Start inserting tooling number')
    if firsttime == False:
        days = (pd.to_datetime(df_araw['Date']).iloc[1] - pd.to_datetime(df_raw['Date']).iloc[1]).days
        if days == 0 or days < 0:
            days = 2
        elif days > 0:
            days = days + 2
            pass
    elif firsttime == True:
        df_araw.columns = df_araw.iloc[0]
        df_araw = df_araw.drop(df_araw.index[0])
        df_araw['Date'] = pd.to_datetime(df_araw['Date'])
        days = (df_araw['Date'].iloc[1] - df_araw['Date'].iloc[-1]).days
        days = days + 1

    df_ora = SQLQuery(str(days))
    df_araw = InsertTooling(df_ora,df_araw)
    dataframe_save_file(fileloc,df_araw)
    print('Tooling number inserted')
    print('File updated and saved to \\TAFS\Public\Alan Yuan\SO_TSSOP Files or \\TAFS\Public\Alan Yuan\QFP Files')
    return

def FindMax(fileloc):
    '''
    wrapper func around Find_Unique_Max()
    Input: file location string
    Output: None
    '''
    print('Start finding max values for each time frame')
    df_araw = open_xlsx_file(fileloc)
    df_final = find_unique_max(df_araw,fileloc)    
    if fileloc == r'\\TAFS\Public\Alan Yuan\SO_TSSOP Files\AccumRawDataWithTooling.xlsx':
        nparray_save_file(r'\\TAFS\Public\Alan Yuan\SO_TSSOP Files\MaxValues.csv',df_final)
        df_final2 = pd.DataFrame(df_final)
        dataframe_save_file(r'\\TAFS\Public\Alan Yuan\SO_TSSOP Files\MaxValues.xlsx',df_final2)
        print('Saving results to csv and xlsx file to \\TAFS\Public\Alan Yuan\SO_TSSOP Files')
    elif fileloc == r'\\TAFS\Public\Alan Yuan\QFP Files\AccumRawDataWithTooling.xlsx':
        nparray_save_file(r'\\TAFS\Public\Alan Yuan\QFP Files\MaxValues.csv',df_final)
        df_final2 = pd.DataFrame(df_final)
        dataframe_save_file(r'\\TAFS\Public\Alan Yuan\QFP Files\MaxValues.xlsx',df_final2)
        print('Saving results to csv and xlsx file to \\TAFS\Public\Alan Yuan\QFP Files')
    print('Files saved')
    return

def DFGen():# 從SO_TSSOP,QFP的"MaxValues.xlsx"檔案中取資料出來    
    '''
    Separate SOIC, TSSOP, QFP data from one file
    Input: file location string
    Output: Two Dataframes
    '''
    df_SOT = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\SO_TSSOP Files\MaxValues.xlsx')
    df_QFP = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\QFP Files\MaxValues.xlsx')
    df_SOT.columns = df_SOT.iloc[0]
    df_SOT = df_SOT.drop([0])
    df_QFP.columns = df_QFP.iloc[0]
    df_QFP = df_QFP.drop([0])
    return df_SOT, df_QFP

def Plot(df,Category,Product,Pin,Tool,Perc):# 以所有資料繪圖
    '''
    Draw plots using all data
    Input: Dataframe (pd df), category name (str), product name (str), pin type (str), tooling number (int), percentage or not (int)
        Perc = 0 --- Value
        Perc = 1 --- Percentage
    Output: matplotlib plots
    '''
    df['Date'] = pd.to_datetime(df['Date'])
    if Perc == 0:
        if Product == 'SOIC':
            df_limits = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\SOIC_Limit.xlsx')
        elif Product == 'TSSOP':
            df_limits = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\TSSOP_Limit.xlsx')
        elif Product == 'QFP':
            df_limits = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\QFP_Limit.xlsx')
        df_limits = df_limits.set_index('Pin Type')
        if Pin == 'All' or Tool == 'All':
            if Pin == 'All' and Tool != 'All':
                pins = pd.unique(df['P/P'].dropna())
                for pin in pins:
                    plt.plot(df[df['P/P'] == pin]['Date'], pd.to_numeric(df[df['P/P'] == pin][Category]),label = pin)
            elif Pin != 'All' and Tool == 'All':
                tools = pd.unique(df['Tooling'].dropna())
                fig,ax = plt.subplots()
                for tool in tools:
                    plt.plot(df[df['Tooling'] == tool]['Date'], pd.to_numeric(df[df['Tooling'] == tool][Category]),label = tool)
                if Pin in df_limits.index:
                    ax.axhline(pd.to_numeric(df_limits.loc[Pin][Category + ' Upper']), color ='red', ls = '--')
                    ax.axhline(pd.to_numeric(df_limits.loc[Pin][Category + ' Lower']), color ='red', ls = '--')
                else:
                    print('No upper and lower limit data for '+Pin)
            elif Pin == 'All' and Tool == 'All':
                plt.plot(df['Date'],pd.to_numeric(df[Category]),label = 'All')
            pass
        else:
            fig,ax = plt.subplots()
            x = mdates.date2num(pd.to_datetime(df['Date']))
            z4 = np.polyfit(x, pd.to_numeric(df[Category]),1)
            p4 = np.poly1d(z4)
            xx = np.linspace(x.min(), x.max(), 100)
            dd = mdates.num2date(xx)
            plt.plot(dd, p4(xx), '-.m',lw = 1)
            plt.plot(df['Date'],pd.to_numeric(df[Category]),label = Pin)
            if Pin in df_limits.index:
                    ax.axhline(pd.to_numeric(df_limits.loc[Pin][Category + ' Upper']), color ='red', ls = '--')
                    ax.axhline(pd.to_numeric(df_limits.loc[Pin][Category + ' Lower']), color ='red', ls = '--')
            else:
                print('No upper and lower limit data for '+Pin)
            plt.ylim(pd.to_numeric(df_limits.loc[Pin][Category + ' Lower'])-2,pd.to_numeric(df_limits.loc[Pin][Category + ' Upper'])+2)
    elif Perc == 1:
        if Pin == 'All' or Tool == 'All':
            if Pin == 'All' and Tool != 'All':
                pass
            elif Pin != 'All' and Tool == 'All':
                tools = pd.unique(df['Tooling'].dropna())
                fig,ax = plt.subplots()
                for tool in tools:
                    plt.plot(df[df['Tooling'] == tool]['Date'],pd.to_numeric(df[df['Tooling'] == tool][Category]),label = tool)
                ax.axhline(100, color ='red', ls = '--')
                pass
            elif Pin == 'All' and Tool == 'All':
                pass
        else:
            fig,ax = plt.subplots()
            x = mdates.date2num(pd.to_datetime(df['Date']))
            z4 = np.polyfit(x, pd.to_numeric(df[Category]),1)
            p4 = np.poly1d(z4)
            xx = np.linspace(x.min(), x.max(), 100)
            dd = mdates.num2date(xx)
            plt.plot(dd, p4(xx), '-.m',lw = 1)
            plt.plot(df['Date'],pd.to_numeric(df[Category]),label = Pin)
            ax.axhline(100, color ='red', ls = '--')
        plt.ylim(0, 120)
    plt.xlabel('Date and Time')
    plt.ylabel(Category)
    plt.title(Product + ' ' + Tool)
    plt.legend()
    plt.gcf().autofmt_xdate()
    plt.show(block=False)
    return 

def GeneratePlot(Product,DataMode,Category,Tool,Pin):# 這函式負責跑Plot()
    '''
    wrapper func around Plot()
    Input: product name (str), datamode between percentage and value (str), category (str), tooling num (int), pin type (str)
    Output: None
    '''
    df = None
    df_SOT,df_QFP = DFGen()
    if DataMode == 'Value':
        if Pin == 'All' or Tool == 'All':
            if Product == 'SOIC' or Product == 'TSSOP':
                if Pin == 'All' and Tool != 'All':
                    df = df_SOT[df_SOT['Product'] == Product][df_SOT['Tooling'] == Tool]
                    Plot(df,Category,Product,Pin,Tool,0)
                elif Pin != 'All' and Tool == 'All':
                    df = df_SOT[df_SOT['Product'] == Product][df_SOT['P/P'] == Pin]
                    Plot(df,Category,Product,Pin,Tool,0)
                elif Pin == 'All' and Tool == 'All':
                    df = df_SOT[df_SOT['Product'] == Product]
                    Plot(df,Category,Product,Pin,Tool,0)
            elif Product == 'QFP':
                if Pin == 'All' and Tool != 'All':
                    df = df_QFP[df_QFP['Product'] == Product][df_QFP['Tooling'] == Tool]
                    Plot(df,Category,Product,Pin,Tool,0)
                elif Pin != 'All' and Tool == 'All':
                    df = df_QFP[df_QFP['Product'] == Product][df_QFP['P/P'] == Pin]
                    Plot(df,Category,Product,Pin,Tool,0)
                elif Pin == 'All' and Tool == 'All':
                    df = df_QFP[df_QFP['Product'] == Product]
                    Plot(df,Category,Product,Pin,Tool,0)
        else:
            if Product == 'SOIC' or Product == 'TSSOP':
                df = df_SOT[df_SOT['Product'] == Product][df_SOT['P/P'] == Pin][df_SOT['Tooling'] == Tool]
                if df.shape[0] > 2:
                    Plot(df,Category,Product,Pin,Tool,0)
                else:
                    print('Data insufficient')
            elif Product == 'QFP':
                df = df_QFP[df_QFP['P/P'] == Pin][df_QFP['Tooling'] == Tool]
                if df.shape[0] > 2:
                    Plot(df,Category,Product,Pin,Tool,0)
                else:
                    print('Data insufficient')
    elif DataMode == 'Percentage':
        if Product == 'SOIC':
            df_limits = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\SOIC_Limit.xlsx')
        elif Product == 'TSSOP':
            df_limits = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\TSSOP_Limit.xlsx')
        elif Product == 'QFP':
            df_limits = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\QFP_Limit.xlsx')
        df_limits = df_limits.set_index('Pin Type')
        df_range = pd.to_numeric(df_limits[Category + ' Upper']) - pd.to_numeric(df_limits[Category + ' Lower'])
        if Pin == 'All' or Tool == 'All':
            if Product == 'TSSOP' or Product == 'SOIC':
                if Pin == 'All' and Tool != 'All':
                    pins = pd.unique(df_SOT[df_SOT['Tooling'] == Tool]['P/P'].dropna())
                    fig,ax = plt.subplots()
                    for pin in pins:
                        if pin in df_limits.index:
                            df_SOT[Category + ' Percentage'] = 100*(pd.to_numeric(df_SOT[df_SOT['Product'] == Product][df_SOT['P/P'] == pin][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][pin]))/df_range[pin]
                            df = df_SOT.dropna(subset = [Category + ' Percentage'])
                            df = df[df['Tooling'] == Tool]
                            df['Date'] = pd.to_datetime(df['Date'])
                            plt.plot(df['Date'], pd.to_numeric(df[Category + ' Percentage']),label = pin)
                    ax.axhline(100, color ='red', ls = '--')
                    if df is not None:
                        Plot(df,Category + ' Percentage',Product,Pin,Tool,1)
                    else:
                        print('Selected P/P is incorrect/does not exist')
                elif Pin != 'All' and Tool == 'All':
                    if Pin in df_limits.index:
                        df_SOT[Category + ' Percentage'] = 100*(pd.to_numeric(df_SOT[df_SOT['Product'] == Product][df_SOT['P/P'] == Pin][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][Pin]))/df_range[Pin]
                        df = df_SOT.dropna(subset = [Category + ' Percentage'])
                        if df.shape[0] > 2:
                            Plot(df,Category + ' Percentage',Product,Pin,Tool,1)
                        else:
                            print('Data is insufficient')
                    else:
                        print('Selected P/P is incorrect/does not exist')
                elif Pin == 'All' and Tool == 'All':
                    pins = pd.unique(df_SOT['P/P'].dropna())
                    fig,ax = plt.subplots()
                    for pin in pins:
                        if pin in df_limits.index:
                            df_SOT[Category + ' Percentage'] = 100*(pd.to_numeric(df_SOT[df_SOT['Product'] == Product][df_SOT['P/P'] == pin][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][pin]))/df_range[pin]
                            df = df_SOT.dropna(subset = [Category + ' Percentage'])
                            df['Date'] = pd.to_datetime(df['Date'])
                            plt.plot(df['Date'], pd.to_numeric(df[Category + ' Percentage']),label = pin + 'All Tooling')
                    ax.axhline(100, color ='red', ls = '--')
                    if df is not None:
                        Plot(df,Category + ' Percentage',Product,Pin,Tool,1)
                    else:
                        print('Selected P/P is incorrect/does not exist')
            elif Product == 'QFP':
                if Pin == 'All' and Tool != 'All':
                    pins = pd.unique(df_QFP[df_QFP['Tooling'] == Tool]['P/P'].dropna())
                    fig,ax = plt.subplots()
                    for pin in pins:
                        if pin in df_limits.index:
                            df_QFP[Category + ' Percentage'] = 100*(pd.to_numeric(df_QFP[df_QFP['Product'] == Product][df_QFP['P/P'] == pin][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][pin]))/df_range[pin]
                            df = df_QFP.dropna(subset = [Category + ' Percentage'])
                            df = df[df['Tooling'] == Tool]
                            df['Date'] = pd.to_datetime(df['Date'])                            
                            plt.plot(df['Date'], pd.to_numeric(df[Category + ' Percentage']),label = pin)
                    ax.axhline(100, color ='red', ls = '--')                            
                    if df is not None:
                        Plot(df,Category + ' Percentage',Product,Pin,Tool,1)
                    else:
                        print('Selected P/P is incorrect/does not exist')
                elif Pin != 'All' and Tool == 'All':
                    if Pin in df_limits.index:
                        df_QFP[Category + ' Percentage'] = 100*(pd.to_numeric(df_QFP[df_QFP['P/P'] == Pin][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][Pin]))/df_range[Pin]
                        df = df_QFP.dropna(subset = [Category + ' Percentage'])
                        if df.shape[0] > 2:
                            Plot(df,Category + ' Percentage',Product,Pin,Tool,1)
                        else:
                            print('Data is insufficient')
                    else:
                        print('Selected P/P is incorrect/does not exist')
                elif Pin == 'All' and Tool == 'All':
                    pins = pd.unique(df_QFP['P/P'].dropna())
                    fig,ax = plt.subplots()
                    for pin in pins:
                        if pin in df_limits.index:
                            df_QFP[Category + ' Percentage'] = 100*(pd.to_numeric(df_QFP[df_QFP['Product'] == Product][df_QFP['P/P'] == pin][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][pin]))/df_range[pin]
                            df = df_QFP.dropna(subset = [Category + ' Percentage'])
                            df['Date'] = pd.to_datetime(df['Date'])
                            plt.plot(df['Date'], pd.to_numeric(df[Category + ' Percentage']),label = pin + 'All Tooling')
                    ax.axhline(100, color ='red', ls = '--')
                    if df is not None:
                        Plot(df,Category + ' Percentage',Product,Pin,Tool,1)
                    else:
                        print('Selected P/P is incorrect/does not exist')
                pass
            pass
        else:            
            if Product == 'SOIC' or Product == 'TSSOP':
                if Pin in df_limits.index:
                    df_SOT[Category + ' Percentage'] = 100*(pd.to_numeric(df_SOT[df_SOT['Product'] == Product][df_SOT['P/P'] == Pin][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][Pin]))/df_range[Pin]
                    df = df_SOT.dropna(subset = [Category + ' Percentage'])
                    df = df[df['Tooling'] == Tool]
                    if df.shape[0] > 2:
                        Plot(df,Category + ' Percentage',Product,Pin,Tool,1)
                    else:
                        print('Data is insufficient')
                else:
                    print('Selected P/P is incorrect/does not exist')
            elif Product == 'QFP':
                if Pin in df_limits.index:
                    df_QFP[Category + ' Percentage'] = 100*(pd.to_numeric(df_QFP[df_QFP['P/P'] == Pin][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][Pin]))/df_range[Pin]
                    df = df_QFP.dropna(subset = [Category + ' Percentage'])
                    df = df[df['Tooling'] == Tool]
                    if df.shape[0] > 2:
                        Plot(df,Category + ' Percentage',Product,Pin,Tool,1)
                    else:
                        print('Data is insufficient')
                else:
                    print('Selected P/P is incorrect/does not exist')
        pass
    return

def Generate7Plot(Product,Category,Tool,df,threshold,Days):# 以使用者選擇的天數及門檻繪圖
    '''
    Generate plots with a given timeframe in days and a chosen threshold to determine whether the tool is near
    its life cycle end
    Input: product name (str), category name (str), tooling num (int), dataframe, threshold (str), days (int)
    Output: plots
    '''
    threshold = int(threshold)
    ### FOR MAIL ###
    df_mail = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\Email Recipients.xlsx')
    mails = str()
    for i in range(df_mail['Email Recipients'].dropna().shape[0]):
        emails = df_mail['Email Recipients'][i] + ';'
        mails = mails + emails
    ################
    if Tool == 'All':
        bad_tools = []
        tools = pd.unique(df['Tooling'])
        col = 4
        row = int(np.ceil(tools.shape[0]/col))
        i = 1
        fig,ax = plt.subplots(row, col,squeeze=False)
        ax = ax.T.flatten()
        for tool in tools:            
            plt.subplot(row,col,i)           
            pins = pd.unique(df[df['Tooling'] == tool]['P/P'])
            if df[df['Tooling'] == tool].shape[0] > 1 and df[df['Tooling'] == tool][pd.to_numeric(df[df['Tooling'] == tool][Category + ' Percentage']) >= threshold].shape[0] >= 0.5*df[df['Tooling'] == tool].shape[0]:
                plt.plot(df[df['Tooling'] == tool]['Date'], df[df['Tooling'] == tool][Category + ' Percentage'],label = tool,color = 'red',linewidth = 2)
                bad_tools.append(tool)
            else:
                plt.plot(df[df['Tooling'] == tool]['Date'], df[df['Tooling'] == tool][Category + ' Percentage'],label = tool)
            plt.title(Product + ' ' + tool + ' ' + pins[0])
            plt.ylim(0, 120)
            i = i+1
        firsttime = True
        if bad_tools != []:
            for tool in bad_tools:
                df2 = df[df['Tooling'] == tool]
                if firsttime == True:
                    df_bad = df2
                    firsttime = False
                elif firsttime == False:
                    df_bad = pd.concat([df_bad,df2])
        plt.gcf().autofmt_xdate()
        fig.canvas.set_window_title(Category + ' Percentage')
        plt.tight_layout()
        fig.savefig(r'\\TAFS\Public\Alan Yuan\SOTSSOP and QFP files\BadTools'+ Category + r'.png')
        plt.show(block=False)
        if bad_tools != []:
            toolstr = '\n'.join(bad_tools)
            fig.savefig(r'\\TAFS\Public\Alan Yuan\SOTSSOP and QFP files\BadTools'+ Category + r'.png')
            dataframe_save_file(r'\\TAFS\Public\Alan Yuan\SOTSSOP and QFP files\BadTools' + Category + r'.xlsx',df_bad)
            Mail(mails,threshold,toolstr,Category,Days,r'\\TAFS\Public\Alan Yuan\SOTSSOP and QFP files\BadTools' + Category + r'.xlsx',r'\\TAFS\Public\Alan Yuan\SOTSSOP and QFP files\BadTools'+ Category + r'.png')

    else:
        if df is None:
            print('Insufficient Data')
        else:
            df['Date'] = pd.to_datetime(df['Date'])
            fig,ax = plt.subplots()
            x = mdates.date2num(pd.to_datetime(df['Date']))
            z4 = np.polyfit(x, pd.to_numeric(df[Category + ' Percentage']),1)
            p4 = np.poly1d(z4)
            xx = np.linspace(x.min(), x.max(), 100)
            dd = mdates.num2date(xx)
            plt.plot(dd, p4(xx), '-.m',lw = 1)
            plt.plot(df['Date'],df[Category + ' Percentage'])            
            ax.axhline(100, color ='red', ls = '--')
            ax.axhline(threshold, color ='green', ls = '--',lw = 0.5)
            plt.xlabel('Date and Time')
            plt.ylabel(Category + ' Percentage')
            if df[pd.to_numeric(df[Category + ' Percentage']) >= threshold].shape[0] >= 0.5*df.shape[0]:
                fig.text(0.4, 0.9,'ALERT',fontsize = 40,color = 'red')
            if Category == 'Tip to Tip' or Category == 'Tip to Tip X' or Category == 'Tip to Tip Y':
                fig.canvas.set_window_title(Product + ' ' + Tool + ' Lead Cut Punch/Die')
            elif Category == 'Seating Height' or Category == 'Seating Height X' or Category == 'Seating Height Y':
                fig.canvas.set_window_title(Product + ' ' + Tool + ' Final Form punch/die or Stroke motor bearing')
            elif Category == 'Foot Angle' or Category == 'Foot Angle X' or Category == 'Foot Angle Y':
                fig.canvas.set_window_title(Product + ' ' + Tool + ' Pre Form Punch/Die')
            elif Category == 'Interlead Flash':
                fig.canvas.set_window_title(Product + ' ' + Tool + ' Cutting plate or De gate punch')
            elif Category == 'Lead Length X' or Category == 'Lead Length Y':
                fig.canvas.set_window_title(Product + ' ' + Tool + ' Pre Form Punch/Die')
            else:
                fig.canvas.set_window_title(Product + ' ' + Tool)
            plt.ylim(0, 120)
            plt.gcf().autofmt_xdate()
            plt.show(block=False)
    return

def PercTable(Product,Category,Tool,df_SOT,df_QFP,firsttime,df_limits,df_range,Days):
    '''
    Generate dataframe for tables with a given timeframe in days and a chosen threshold to determine whether the tool is near
    its life cycle end. Table is in percentage.
    Input: product name (str), category name (str), tooling num (int), dataframe, dataframe, 
    first time (int), limits (int), range (int0), days (int)
    Output: dataframe
    '''
    df = None
    data = []
    header_list = []
    if Tool == 'All':
        if Product == 'SOIC' or Product == 'TSSOP':
            tools = pd.unique(df_SOT[df_SOT['Product'] == Product]['Tooling'].dropna())
            for tool in tools:
                pins = pd.unique(df_SOT[df_SOT['Product'] == Product][df_SOT['Tooling'] == tool]['P/P'].dropna())
                for pin in pins:
                    if pin in df_limits.index:
                        df_SOT[Category + ' Percentage'] = np.round(100*(pd.to_numeric(df_SOT[df_SOT['Product'] == Product][df_SOT['P/P'] == pin][df_SOT['Tooling'] == tool][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][pin]))/df_range[pin],decimals = 2)
                        if firsttime == True:
                            df = df_SOT.dropna(subset = [Category + ' Percentage'])
                            firsttime = False
                        else:
                            temp = df_SOT.dropna(subset = [Category + ' Percentage'])
                            df = pd.concat([df,temp])
                    else:
                        pass
                    if df is None:
                        pass
                    else:
                        df['Date'] = pd.to_datetime(df['Date'])
                        mask = df['Date'] > (datetime.datetime.today() - datetime.timedelta(days = int(Days)))
                        df = df.loc[mask].sort_values(by = 'Date', ascending=False)
            df_tools = pd.unique(df['Tooling'].dropna())
            firsttime = True
            for tool in df_tools:
                if firsttime == True:
                    df_2 = df[df['Tooling'] == tool]
                    firsttime = False
                else:
                    temp = df[df['Tooling'] == tool]
                    df_2 = pd.concat([df_2,temp])
            tooling = df_2['Tooling']
            df_2.drop(labels = ['Tooling','Product','Device','Lot','Tip to Tip','Interlead Flash','Seating Height','Foot Angle'],axis = 1, inplace = True)
            df_2.insert(0,'Tooling',tooling)
            df_final = df_2
            data = df_final.values.tolist()
            header_list = df_final.columns.tolist()
        elif Product == 'QFP':
            tools = pd.unique(df_QFP['Tooling'].dropna())
            for tool in tools:
                pins = pd.unique(df_QFP[df_QFP['Tooling'] == tool]['P/P'].dropna())
                for pin in pins:
                    if pin in df_limits.index:
                        df_QFP[Category + ' Percentage'] = np.round(100*(pd.to_numeric(df_QFP[df_QFP['Product'] == Product][df_QFP['P/P'] == pin][df_QFP['Tooling'] == tool][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][pin]))/df_range[pin],decimals=2)
                        if firsttime == True:
                            df = df_QFP.dropna(subset = [Category + ' Percentage'])
                            firsttime = False
                        else:
                            temp = df_QFP.dropna(subset = [Category + ' Percentage'])
                            df = pd.concat([df,temp])
                    else:
                        pass
                    if df is None:
                        pass
                    else:
                        df['Date'] = pd.to_datetime(df['Date'])
                        mask = df['Date'] > (datetime.datetime.today() - datetime.timedelta(days = int(Days)))
                        df = df.loc[mask].sort_values(by = 'Date', ascending=False)
            df_tools = pd.unique(df['Tooling'].dropna())
            firsttime = True
            for tool in df_tools:
                if firsttime == True:
                    df_2 = df[df['Tooling'] == tool]
                    firsttime = False
                else:
                    temp = df[df['Tooling'] == tool]
                    df_2 = pd.concat([df_2,temp])
            tooling = df_2['Tooling']
            df_2.drop(labels = ['Tooling','Product','Device','Lot','Tip to Tip X','Tip to Tip Y','Seating Height X','Seating Height Y','Foot Angle X','Foot Angle Y','Lead Length X','Lead Length Y'],axis = 1, inplace = True)
            df_2.insert(0,'Tooling',tooling)
            df_final = df_2
            data = df_final.values.tolist()
            header_list = df_final.columns.tolist()
    else:
        if Product == 'SOIC' or Product == 'TSSOP':        
            pins = pd.unique(df_SOT[df_SOT['Product'] == Product][df_SOT['Tooling'] == Tool]['P/P'].dropna())
            for pin in pins:
                if pin in df_limits.index:
                    df_SOT[Category + ' Percentage'] = np.round(100*(pd.to_numeric(df_SOT[df_SOT['Product'] == Product][df_SOT['P/P'] == pin][df_SOT['Tooling'] == Tool][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][pin]))/df_range[pin],decimals=2)
                    if firsttime == True:
                        df = df_SOT.dropna(subset = [Category + ' Percentage'])
                        firsttime = False
                    else:
                        temp = df_SOT.dropna(subset = [Category + ' Percentage'])
                        df = pd.concat([df,temp])
                else:
                    pass
            if df is None:
                pass
            else:
                tooling = df['Tooling']
                df.drop(labels = ['Tooling','Product','Device','Lot','Tip to Tip','Interlead Flash','Seating Height','Foot Angle'],axis = 1, inplace = True)
                df.insert(0,'Tooling',tooling)
                df['Date'] = pd.to_datetime(df['Date'])
        elif Product == 'QFP':
            pins = pd.unique(df_QFP[df_QFP['Tooling'] == Tool]['P/P'].dropna())
            for pin in pins:
                if pin in df_limits.index:
                    df_QFP[Category + ' Percentage'] = np.round(100*(pd.to_numeric(df_QFP[df_QFP['Product'] == Product][df_QFP['P/P'] == pin][df_QFP['Tooling'] == Tool][Category],errors = 'coerce') - pd.to_numeric(df_limits[Category + ' Lower'][pin]))/df_range[pin],decimals=2)
                    if firsttime == True:
                        df = df_QFP.dropna(subset = [Category + ' Percentage'])
                        firsttime = False
                    else:
                        temp = df_QFP.dropna(subset = [Category + ' Percentage'])
                        df = pd.concat([df,temp])
                else:
                    pass
            if df is None:
                pass
            else:
                tooling = df['Tooling']
                df.drop(labels = ['Tooling','Product','Device','Lot','Tip to Tip X','Tip to Tip Y','Seating Height X','Seating Height Y','Foot Angle X','Foot Angle Y','Lead Length X','Lead Length Y'],axis = 1, inplace = True)
                df.insert(0,'Tooling',tooling)
                df['Date'] = pd.to_datetime(df['Date'])
        if df is None:
            df_final = None
        else:
            mask = df['Date'] > (datetime.datetime.today() - datetime.timedelta(days = int(Days)))
            df_final = df.loc[mask].sort_values(by = 'Date', ascending=False)
            data = df_final.values.tolist()
            header_list = df_final.columns.tolist()
    return df_final,data,header_list

def RowColor(indexlist,df_tools):
    '''
    Color the tables to make tables look clearer
    Input: indexlist (list), tools (list)
    Output: tuple
    '''
    list_ = []
    for i in range(df_tools.shape[0]):
        if i%2 == 0:
            for j in range(len(indexlist[i])):
                list_.append((indexlist[i][j],'#000000'))
        else:
            for j in range(len(indexlist[i])):
                list_.append((indexlist[i][j],'#212F3C'))
    l = tuple(list_)
    return l

def PercAverage(df,Category):
    ave = None
    if df is None:
        pass
    else:
        tools = pd.unique(df['Tooling'])
        ave = np.array([['Tooling',Category + ' Percentage Average']])
        for tool in tools:
            temp = np.array([[tool, np.round(pd.to_numeric(df[df['Tooling'] == tool][Category + ' Percentage'].mean()),decimals = 2)]])
            ave = np.vstack([ave,temp])
        ave = pd.DataFrame(ave)
        ave.columns = ave.iloc[0]
        ave = ave.drop([0])
    return ave

def RawPercAve(Product,df):
    if Product == 'SOIC' or Product == 'TSSOP':
        tools = pd.unique(df['Tooling'])
        ave = np.array([['Tooling','Tip to Tip Average','Tip to Tip % Average','Interlead Flash Average','Interlead Flash % Average','Seating Height Average','Seating Height % Average','Foot Angle Average','Foot Angle % Average']])
        for tool in tools:
            firsttime = True
            for i in range(7,15):
                temp1 = np.array([[np.round(pd.to_numeric(df[df['Tooling'] == tool][df.columns[i]]).mean(skipna = True),decimals = 2)]])
                if firsttime == True:
                    row = np.array([[tool]])
                    row = np.hstack([row,temp1])
                    firsttime = False
                elif firsttime == False:
                    row = np.hstack([row,temp1])
            ave = np.vstack([ave,row])
    elif Product == 'QFP':
        tools = pd.unique(df['Tooling'])
        ave = np.array([['Tooling','Tip to Tip X Average','Tip to Tip X % Average','Tip to Tip Y Average','Tip to Tip Y % Average','Seating Height X Average','Seating Height X % Average','Foot Angle X Average','Foot Angle X % Average','Lead Length X Average','Lead Length X % Average','Seating Height Y Average','Seating Height Y % Average','Foot Angle Y Average','Foot Angle Y % Average','Lead Length Y Average','Lead Length Y % Average']])
        for tool in tools:
            firsttime = True
            for i in range(7,23):
                temp1 = np.array([[np.round(pd.to_numeric(df[df['Tooling'] == tool][df.columns[i]]).mean(skipna = True),decimals = 2)]])
                if firsttime == True:
                    row = np.array([[tool]])
                    row = np.hstack([row,temp1])
                    firsttime = False
                elif firsttime == False:
                    row = np.hstack([row,temp1])
            ave = np.vstack([ave,row])
    ave = pd.DataFrame(ave)
    ave.columns = ave.iloc[0]
    ave = ave.drop([0])
    return ave

def RawTable(Product,Tool,Days,full):    
    df_SO,df_TS,df_QFP = FullTable()
    firsttime = True
    if full == True:
        if Tool == 'All':
            if Product == 'SOIC':
                tools = pd.unique(df_SO['Tooling'].dropna())
                for tool in tools:
                    df_tool = df_SO[df_SO['Tooling'] == tool]
                    if firsttime == True:
                        df_SOF = df_tool
                        firsttime = False
                    else:
                        df_SOF = pd.concat([df_SOF,df_tool])
                return df_SOF
            elif Product == 'TSSOP':
                tools = pd.unique(df_TS['Tooling'].dropna())
                for tool in tools:
                    df_tool = df_TS[df_TS['Tooling'] == tool]
                    if firsttime == True:
                        df_TSF = df_tool
                        firsttime = False
                    else:
                        df_TSF = pd.concat([df_TSF,df_tool])
                return df_TSF
            elif Product == 'QFP':
                tools = pd.unique(df_QFP['Tooling'].dropna())
                for tool in tools:
                    df_tool = df_QFP[df_QFP['Tooling'] == tool]
                    if firsttime == True:
                        df_QFPF = df_tool
                        firsttime = False
                    else:
                        df_QFPF = pd.concat([df_QFPF,df_tool])
                return df_QFPF
        else:
            if Product == 'SOIC':
                df_SO = df_SO[df_SO['Tooling'] == Tool]
                return df_SO
            elif Product == 'TSSOP':
                df_TS = df_TS[df_TS['Tooling'] == Tool]
                return df_TS
            elif Product == 'QFP':
                df_QFP = df_QFP[df_QFP['Tooling'] == Tool]
                return df_QFP
        pass
    elif full == False:
        mask1 = df_SO['Date'] > (datetime.datetime.today() - datetime.timedelta(days = int(Days)))
        mask2 = df_TS['Date'] > (datetime.datetime.today() - datetime.timedelta(days = int(Days)))
        mask3 = df_QFP['Date'] > (datetime.datetime.today() - datetime.timedelta(days = int(Days)))
        df_SO = df_SO.loc[mask1].sort_values(by = 'Date', ascending=False)
        df_TS = df_TS.loc[mask2].sort_values(by = 'Date', ascending=False)
        df_QFP = df_QFP.loc[mask3].sort_values(by = 'Date', ascending=False)
        if Tool == 'All':
            if Product == 'SOIC':
                tools = pd.unique(df_SO['Tooling'].dropna())
                for tool in tools:
                    df_tool = df_SO[df_SO['Tooling'] == tool]
                    if firsttime == True:
                        df_SOF = df_tool
                        firsttime = False
                    else:
                        df_SOF = pd.concat([df_SOF,df_tool])
                return df_SOF
            elif Product == 'TSSOP':
                tools = pd.unique(df_TS['Tooling'].dropna())
                for tool in tools:
                    df_tool = df_TS[df_TS['Tooling'] == tool]
                    if firsttime == True:
                        df_TSF = df_tool
                        firsttime = False
                    else:
                        df_TSF = pd.concat([df_TSF,df_tool])
                return df_TSF
            elif Product == 'QFP':
                tools = pd.unique(df_QFP['Tooling'].dropna())
                for tool in tools:
                    df_tool = df_QFP[df_QFP['Tooling'] == tool]
                    if firsttime == True:
                        df_QFPF = df_tool
                        firsttime = False
                    else:
                        df_QFPF = pd.concat([df_QFPF,df_tool])
                return df_QFPF
        else:
            if Product == 'SOIC':
                df_SO = df_SO[df_SO['Tooling'] == Tool]
                return df_SO
            elif Product == 'TSSOP':
                df_TS = df_TS[df_TS['Tooling'] == Tool]
                return df_TS
            elif Product == 'QFP':
                df_QFP = df_QFP[df_QFP['Tooling'] == Tool]
                return df_QFP
            pass

    return

def FullTable():
    df_SOT,df_QFP = DFGen()
    df_limits_SO = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\SOIC_Limit.xlsx')
    df_limits_TS = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\TSSOP_Limit.xlsx')
    df_limits_QF = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\QFP_Limit.xlsx')

    df_limits_SO = df_limits_SO.set_index('Pin Type')
    df_limits_TS = df_limits_TS.set_index('Pin Type')
    df_limits_QF = df_limits_QF.set_index('Pin Type')

    df_SO = df_SOT[df_SOT['Product'] == 'SOIC']
    df_TS = df_SOT[df_SOT['Product'] == 'TSSOP']
    
    pins_SO = pd.unique(df_SO['P/P'].dropna())
    pins_TS = pd.unique(df_TS['P/P'].dropna())
    pins_QF = pd.unique(df_QFP['P/P'].dropna())    
    
    df_ranges_SO = []
    df_ranges_TS = []
    df_ranges_QF = []
    for i in range(7,11):
        df_ranges_SO.append(pd.to_numeric(df_limits_SO[df_SOT.columns[i] + ' Upper']) - pd.to_numeric(df_limits_SO[df_SOT.columns[i] + ' Lower']))
        df_ranges_TS.append(pd.to_numeric(df_limits_TS[df_SOT.columns[i] + ' Upper']) - pd.to_numeric(df_limits_TS[df_SOT.columns[i] + ' Lower']))
    
    for i in range(7,15):
        df_ranges_QF.append(pd.to_numeric(df_limits_QF[df_QFP.columns[i] + ' Upper']) - pd.to_numeric(df_limits_QF[df_QFP.columns[i] + ' Lower']))

    for pin in pins_SO:
        if pin in df_limits_SO.index:
            for i in range(7,11):
                df_SO_perc = np.round(100*(pd.to_numeric(df_SO[df_SO['P/P'] == pin][df_SO.columns[i]],errors = 'coerce') - pd.to_numeric(df_limits_SO[df_SO.columns[i] + ' Lower'][pin]))/df_ranges_SO[i-7][pin],decimals=2)
                df_SO.loc[df_SO['P/P'] == pin,df_SO.columns[i] + ' Percentage'] = df_SO_perc
    j = 0
    for i in range(11,15):        
        perc = df_SO[df_SO.columns[i]]
        df_SO.drop(labels = df_SO.columns[i],axis = 1, inplace = True)
        df_SO.insert(i-3+j,df_SO.columns[i-4+j]+'%',perc)
        j = j+1
    df_SO = df_SO.dropna(subset = [df_SO.columns[i-3]])
    df_SO['Date'] = pd.to_datetime(df_SO['Date'])
    df_SO = df_SO.sort_values(by = 'Date', ascending=False)

    for pin in pins_TS:
        if pin in df_limits_TS.index:
            for i in range(7,11):
                df_TS_perc = np.round(100*(pd.to_numeric(df_TS[df_TS['P/P'] == pin][df_TS.columns[i]],errors = 'coerce') - pd.to_numeric(df_limits_TS[df_TS.columns[i] + ' Lower'][pin]))/df_ranges_TS[i-7][pin],decimals=2)
                df_TS.loc[df_TS['P/P'] == pin,df_TS.columns[i] + ' Percentage'] = df_TS_perc
    j = 0
    for i in range(11,15):        
        perc = df_TS[df_TS.columns[i]]
        df_TS.drop(labels = df_TS.columns[i],axis = 1, inplace = True)
        df_TS.insert(i-3+j,df_TS.columns[i-4+j]+'%',perc)
        j = j+1
    df_TS = df_TS.dropna(subset = [df_TS.columns[i-3]])
    df_TS['Date'] = pd.to_datetime(df_TS['Date'])
    df_TS = df_TS.sort_values(by = 'Date', ascending=False)

    for pin in pins_QF:
        if pin in df_limits_QF.index:
            for i in range(7,15):
                df_QFP_perc = np.round(100*(pd.to_numeric(df_QFP[df_QFP['P/P'] == pin][df_QFP.columns[i]],errors = 'coerce') - pd.to_numeric(df_limits_QF[df_QFP.columns[i] + ' Lower'][pin]))/df_ranges_QF[i-7][pin],decimals=2)
                df_QFP.loc[df_QFP['P/P'] == pin,df_QFP.columns[i] + ' Percentage'] = df_QFP_perc
    j = 0
    for i in range(15,23):        
        perc = df_QFP[df_QFP.columns[i]]
        df_QFP.drop(labels = df_QFP.columns[i],axis = 1, inplace = True)
        df_QFP.insert(i-7+j,df_QFP.columns[i-8+j]+'%',perc)
        j = j+1
    df_QFP = df_QFP.dropna(subset = [df_QFP.columns[i-7]])
    df_QFP['Date'] = pd.to_datetime(df_QFP['Date'])
    df_QFP = df_QFP.sort_values(by = 'Date', ascending=False)
    return df_SO,df_TS,df_QFP

def FindQuant(df_pro,df_map,df_empty,df_PTS,df_SAP,map_pins,map_tools,index1):    
    firsttime = True
    pins = pd.unique(df_pro['P/P'].dropna())
    array = None
    for pin in pins:        
        if pin not in map_pins:
            pass
        else:
            tools= pd.unique(df_pro[df_pro['P/P'] == pin]['Tooling'].dropna())
            for tool in tools:
                if tool not in map_tools:
                    pass
                else:
                    temp = []
                    temp.append(pin)
                    temp.append(tool)
                    j = 2
                    for i in index1:
                        while df_map.columns.tolist()[j][0] == i:
                            ID = df_map[df_map['PKG'] == pin][df_map['Tooling'] == tool][i][df_map.columns.tolist()[j][1]][df_map.columns.tolist()[j][2]].values[0]
                            if j%2 == 0:
                                ##### PTS ID #####
                                '''
                                以下是PTS的code:
                                if ID in pd.unique(df_PTS['Material'].dropna()):
                                    amount = df_PTS[df_PTS['Material'] == ID]['Total Stock'].values[0]
                                else:
                                    amount = 'No Matching ID'
                                記得把 amount = 'No Data' 刪掉
                                '''
                                amount = 'No Data'
                                pass
                            elif j%2 != 0:
                                ##### SAP ID #####
                                
                                if ID in pd.unique(df_SAP['Material'].dropna()):
                                    amount = df_SAP[df_SAP['Material'] == ID]['Total Stock'].values[0]
                                else:
                                    amount = 'No Matching ID'
                            temp.append(amount)
                            j = j+1
                            if j == len(df_map.columns.tolist()):
                                break
                    if firsttime == True:
                        array = np.array(temp)
                        firsttime = False
                    elif firsttime == False:
                        array = np.vstack([array,temp])
    if array is not None:
        df_final = pd.DataFrame(np.atleast_2d(array))
        df_final.columns = pd.MultiIndex.from_tuples(tuple(df_empty.columns.tolist()))
    elif array is None:
        df_final = None
    return df_final

def MapGen():
    df_map = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\TrimForm Part Map\Map.xlsx',head = 1)
    df_empty = open_xlsx_file(r'\\TAFS\Public\Alan Yuan\TrimForm Part Map\ToSave.xlsx',head = 1)
    df_map = df_map.rename(columns=lambda x: x if not 'Unnamed' in str(x) else '')
    df_empty = df_empty.rename(columns=lambda x: x if not 'Unnamed' in str(x) else '')
    return df_map,df_empty

def FullParts():
    df_map,df_empty = MapGen()
    df_SAP = open_xlsx_file(r'http://wplnet.sc.ti.com/reports/PartCatalogsTITL.xlsx')
    df_PTS = None
    ##### For PTS data #####
    # df_PTS = open_xlsx_file(INSERT PTS FILE HERE)
    ##### PTS data code end #####
    
    df_SOT, df_QFP = DFGen()
    df_SO = df_SOT[df_SOT['Product'] == 'SOIC']
    df_TS = df_SOT[df_SOT['Product'] == 'TSSOP']

    index1 = []
    for i in range(2,len(df_map.columns.tolist())):
        index1.append(df_map.columns.tolist()[i][0])
    index1 = list(dict.fromkeys(index1))
    map_pins = pd.unique(df_map['PKG'].dropna())
    map_tools = pd.unique(df_map['Tooling'].dropna())

    dfs = [df_SO,df_TS,df_QFP]
    df_parts = []
    for df in dfs:
        df_parts.append(FindQuant(df,df_map,df_empty,df_PTS,df_SAP,map_pins,map_tools,index1))
    firsttime = True
    for i in range(len(df_parts)-1):
        if df_parts[i] is None:
            pass
        elif df_parts[i] is not None:
            if firsttime == True:
                df_final = pd.concat([df_parts[i],df_parts[i+1]])
                firsttime = False
            elif firsttime == False:
                df_final = pd.concat([df_final,df_parts[i+1]])    
    df_final.drop_duplicates(keep = 'first',inplace = True)

    return df_final

def Mail(mails,threshold,tools,Category,Days,File,pic):
    attachments = []
    attachments.append(File)
    attachments.append(pic)
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    for attachment in attachments:
        mail.Attachments.Add(attachment)
    mail.To = mails
    mail.Subject = 'Tools Exceed ' + str(threshold) + '%' + ' on ' + Category + ' for the past ' + Days + ' days'
    mail.Body = 'The following tools exceed the chosen threshold ' + str(threshold) + '%' + ' for ' + Category + ' for the past '+ Days + ' days:\n' + tools
    mail.Send()
    return