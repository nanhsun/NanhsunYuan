# 這檔案是主檔案。執行檔是由這裡跑
import mainFunc as mf # 引入 mainFunc.py
import PySimpleGUI as sg 
import pandas as pd
# 若檔案夾位置有更換，請改每個r""裡的檔案位置到新的檔案位置
def table(Product,Category,Tool,Days): # 這是用來產生"數據百分比"表格的函式 此函式也會計算"百分比平均"。註：有Days變數的函式，皆是要使用者輸入天數
    firsttime = True
    if Product == 'SOIC':
        df_limits = mf.open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\SOIC_Limit.xlsx')
    elif Product == 'TSSOP':
        df_limits = mf.open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\TSSOP_Limit.xlsx')
    elif Product == 'QFP':
        df_limits = mf.open_xlsx_file(r'\\TAFS\Public\Alan Yuan\Tables\QFP_Limit.xlsx')
    df_SOT,df_QFP = mf.DFGen()
    df_limits = df_limits.set_index('Pin Type')
    df_range = pd.to_numeric(df_limits[Category + ' Upper']) - pd.to_numeric(df_limits[Category + ' Lower'])

    df_final,data,header_list = mf.PercTable(Product,Category,Tool,df_SOT,df_QFP,firsttime,df_limits,df_range,Days)
    df_average = mf.PercAverage(df_final,Category)

    if Tool == 'All':
        df_final.index = range(df_final.shape[0])
        df_tools = pd.unique(df_final['Tooling'])
        indexlist = []
        for tool in df_tools:
            indexlist.append(tuple(df_final[df_final['Tooling'] == tool].index.tolist()))
        indexlist = tuple(indexlist)
        l = mf.RowColor(indexlist,df_tools)

    if data != []:
        if Tool == 'All':
            layout = [
                [sg.Table(values=data,
                        headings=header_list,
                        auto_size_columns=True,
                        row_colors = l,
                        num_rows=min(25, len(data)),
                        key = '_TABLE_')],
                [sg.Button('Find Average')]
            ]
        else:
            layout = [
                [sg.Table(values=data,
                        headings=header_list,
                        auto_size_columns=True,
                        background_color='black',
                        num_rows=min(25, len(data)),
                        key = '_TABLE_')],
                [sg.Button('Find Average')]
            ]
    elif data == []:
        layout = [
            [sg.Text('Insufficient Data')]
        ]

    return sg.Window(Product + ' ' + Category, layout, grab_anywhere=False, finalize=True),df_final,df_average

def tableAverage(df): # 這是用來產生"數據百分比平均"表格的函式 "百分比平均"由table()函式得到
    if df is None:
        layout = [
            [sg.Text('Insufficient Data')]
        ]
    else:
        header_list2 = df.columns.tolist()
        data = df.values.tolist()
        layout = [
                    [sg.Table(values=data,
                            headings=header_list2,
                            auto_size_columns=True,
                            num_rows=min(25, len(df)),
                            alternating_row_color='black',
                            background_color= '#212F3C')]
                ]
    return sg.Window('Average', layout, grab_anywhere=False, finalize=True)

def tableRaw(Product,Tool,Days,full,ave): # 這是用來產生"完整數據"表格的函式 (包含確切數據、LOT、Device...等等)
    df = mf.RawTable(Product,Tool,Days,full)
    if ave == True:
        df = mf.RawPercAve(Product,df)
    header_list = df.columns.tolist()
    data = df.values.tolist()
    if Tool == 'All':
        df.index = range(df.shape[0])
        df_tools = pd.unique(df['Tooling'])
        indexlist = []
        for tool in df_tools:
            indexlist.append(tuple(df[df['Tooling'] == tool].index.tolist()))
        indexlist = tuple(indexlist)
        l = mf.RowColor(indexlist,df_tools)    
        layout = [  
                    [sg.Table(values=data,
                            headings=header_list,
                            auto_size_columns=False,
                            num_rows=min(25, len(df)),
                            row_colors = l)],
                    [sg.Text('Save file as:'), sg.InputText(size=(100,1),default_text=r'\\TAFS\Public\Alan Yuan\SOTSSOP and QFP files\Data.xlsx',key = '_SAVE_')],
                    [sg.Button('Save File')]
                ]
    else:
        layout = [  
                    [sg.Table(values=data,
                            headings=header_list,
                            auto_size_columns=False,
                            num_rows=min(25, len(df)),
                            background_color= 'black')],
                    [sg.Text('Save file as:'), sg.InputText(size=(100,1),default_text=r'\\TAFS\Public\Alan Yuan\SOTSSOP and QFP files\Data.xlsx',key = '_SAVE_')],
                    [sg.Button('Save File')]
                ]
    return sg.Window(Product + ' Raw Max Value Data', layout, grab_anywhere=False, finalize=True),df

def PartsTable(PP = None,Tool = None,Category = None):# 這是用來產生"零件數目"表格的函式
    df = mf.FullParts()
    if PP is None and Tool is None and Category is None:
        pass
    else:
        if PP == 'P/P' or Tool == 'Tooling' or Category == 'Category':
            layout = [
                [sg.Text('Please choose a P/P, Tooling, and Category')]
            ]
            return sg.Window('Test', layout, grab_anywhere=False, finalize=True),df
        else:
            df = df[df['PKG'] == PP][df['Tooling'] == Tool]
            df = df[['PKG','Tooling',Category]]

    layout = [  
                [sg.Text('Save file as:'), sg.InputText(size=(100,1),default_text=r'C:\Users\a0489097\Desktop\Test1.xlsx',key = '_SAVE_')],
                [sg.Button('Save Parts to File')]
            ]
    return sg.Window('Test', layout, grab_anywhere=False, finalize=True),df

def main_window():# 這是用來產生"主要UI"的函式
    col1 = [
        [sg.Button('Update SO TSSOP')],
        [sg.Button('Update QFP')],
        [sg.Button('Find SO TSSOP Max Value')],
        [sg.Button('Find QFP Max Value')],
        [sg.Button('Clear Output Window')]
    ]
    col2 = [
        [sg.Combo(['SOIC','TSSOP','QFP'],default_value = 'Product',size = (10,5),key = 'C1'), sg.Combo(['Value','Percentage'],default_value = 'Data Mode',size = (10,5),key = 'C2')],
        [sg.Combo('',default_value = 'Category',size = (15,5),key = 'C3'),sg.Button('UPDATE BY PRODUCT',font = ('Helvetica',8))],
        [sg.Combo([''],default_value = 'Tooling',size = (15,5),key = 'C4'),sg.Button('UPDATE BY P/P',font = ('Helvetica',8))],
        [sg.Combo([''],default_value = 'P/P',size = (15,5),key = 'C5'),sg.Button('UPDATE BY TOOLING',font = ('Helvetica',8))],
        [sg.Button('Generate Full Plot'),sg.Button('Show Full Raw Data'),sg.Button('Full Average Raw Data')],
        [sg.Text('Input desired timeframe (in days): '), sg.InputText(size=(5,1),key='_IN_'),sg.Text('Input desired threshold (in %): '), sg.InputText(size=(3,1),key='_THRES_')],
        [sg.Button('Generate Plots and Tables'), sg.Button('Generate Raw Data'),sg.Button('Average Raw Data')],
        [sg.Button('Find Parts'),sg.Button('Find All Parts')]
    ]
    # ,sg.Output(size=(85,7), background_color='black', text_color='white',key = '_output_')
    layout = [
        [sg.Column(col1)],
        [sg.Column(col2),sg.Image(r'\\TAFS\Public\Alan Yuan\TI Logo.png')]
    ]
    return sg.Window('PROJECT',layout, finalize=True)

if __name__ == "__main__":# 此處是程式會跑的部分。以上函式由這邊用。
    cat1 = ['Tip to Tip','Seating Height','Interlead Flash','Foot Angle']
    cat2 = ['Tip to Tip X','Tip to Tip Y','Seating Height X','Seating Height Y','Foot Angle X','Foot Angle Y','Lead Length X','Lead Length Y']
    window1,window2,window3,window4 = main_window(), None, None, None
    while True:
        window, event, values = sg.read_all_windows()
        if event in (sg.WIN_CLOSED, 'Exit'):
            window.close()
            if window == window2:       # if closing win 2, mark as closed
                window2 = None
            elif window == window1:     # if closing win 1, exit program
                break
            elif window == window3:
                window3 = None
            elif window == window4:
                window4 = None
        elif event == 'Update SO TSSOP':
            mf.UpdateFile(r'\\TAFS\Public\Alan Yuan\SO_TSSOP Files\AccumRawDataWithTooling.xlsx')
        elif event == 'Update QFP':
            mf.UpdateFile(r'\\TAFS\Public\Alan Yuan\QFP Files\AccumRawDataWithTooling.xlsx')
        elif event == 'Find SO TSSOP Max Value':
            mf.FindMax(r'\\TAFS\Public\Alan Yuan\SO_TSSOP Files\AccumRawDataWithTooling.xlsx')
        elif event == 'Find QFP Max Value':
            mf.FindMax(r'\\TAFS\Public\Alan Yuan\QFP Files\AccumRawDataWithTooling.xlsx')
        elif event == 'Clear Output Window':
            window.FindElement('_output_').Update('')
        elif event == 'UPDATE BY PRODUCT':
            df_SOT,df_QFP = mf.DFGen()
            if values['C1'] == 'SOIC':
                tool_SOIC = pd.unique(df_SOT[df_SOT['Product'] == 'SOIC']['Tooling'].dropna())
                tool_SOIC = tool_SOIC.tolist()
                tool_SOIC.insert(0,'All')
                pp_SOIC = pd.unique(df_SOT[df_SOT['Product'] == 'SOIC']['P/P'].dropna())
                pp_SOIC = pp_SOIC.tolist()
                pp_SOIC.insert(0,'All')
                window.Element('C3').update(values = cat1)
                window.Element('C4').update(values = tool_SOIC)
                window.Element('C5').update(values = pp_SOIC)
            elif values['C1'] == 'TSSOP':
                tool_TSSOP = pd.unique(df_SOT[df_SOT['Product'] == 'TSSOP']['Tooling'].dropna())
                tool_TSSOP = tool_TSSOP.tolist()
                tool_TSSOP.insert(0,'All')
                pp_TSSOP = pd.unique(df_SOT[df_SOT['Product'] == 'TSSOP']['P/P'].dropna())
                pp_TSSOP = pp_TSSOP.tolist()
                pp_TSSOP.insert(0,'All')
                window.Element('C3').update(values = cat1)
                window.Element('C4').update(values = tool_TSSOP)
                window.Element('C5').update(values = pp_TSSOP)
            elif values['C1'] == 'QFP':
                tool_QFP = pd.unique(df_QFP['Tooling'].dropna())
                tool_QFP = tool_QFP.tolist()
                tool_QFP.insert(0,'All')
                pp_QFP = pd.unique(df_QFP['P/P'].dropna())
                pp_QFP = pp_QFP.tolist()
                pp_QFP.insert(0,'All')
                window.Element('C3').update(values = cat2)
                window.Element('C4').update(values = tool_QFP)
                window.Element('C5').update(values = pp_QFP)
            else:
                print('Choose a product first')
        elif event == 'UPDATE BY TOOLING':
            df_SOT,df_QFP = mf.DFGen()
            if values['C1'] == 'SOIC' or values['C1'] == 'TSSOP':
                if values['C4'] != 'All':
                    df_tool = pd.unique(df_SOT[df_SOT['Product'] == values['C1']][df_SOT['Tooling'] == values['C4']]['P/P'].dropna())
                    df_tool = df_tool.tolist()
                    df_tool.insert(0,'All')
                    window.Element('C5').update(values = df_tool)
                else:
                    df_tool = pd.unique(df_SOT[df_SOT['Product'] == values['C1']]['P/P'].dropna())
                    df_tool = df_tool.tolist()
                    df_tool.insert(0,'All')
                    window.Element('C5').update(values = df_tool)
            elif values['C1'] == 'QFP':
                if values['C4'] != 'All':
                    df_tool = pd.unique(df_QFP[df_QFP['Tooling'] == values['C4']]['P/P'].dropna())
                    df_tool = df_tool.tolist()
                    df_tool.insert(0,'All')
                    window.Element('C5').update(values = df_tool)
                else:
                    df_tool = pd.unique(df_SOT[df_SOT['Product'] == values['C1']]['P/P'].dropna())
                    df_tool = df_tool.tolist()
                    df_tool.insert(0,'All')
                    window.Element('C5').update(values = df_tool)
            else:
                print('Choose a tooling first')
        elif event == 'UPDATE BY P/P':
            df_SOT,df_QFP = mf.DFGen()
            if values['C1'] == 'SOIC' or values['C1'] == 'TSSOP':
                if values['C5'] != 'All':
                    df_tool = pd.unique(df_SOT[df_SOT['Product'] == values['C1']][df_SOT['P/P'] == values['C5']]['Tooling'].dropna())
                    df_tool = df_tool.tolist()
                    df_tool.insert(0,'All')
                    window.Element('C4').update(values = df_tool)
                else:
                    df_tool = pd.unique(df_SOT[df_SOT['Product'] == values['C1']]['Tooling'].dropna())
                    df_tool = df_tool.tolist()
                    df_tool.insert(0,'All')
                    window.Element('C4').update(values = df_tool)
            elif values['C1'] == 'QFP':
                if values['C4'] != 'All':
                    df_tool = pd.unique(df_QFP[df_QFP['P/P'] == values['C5']]['Tooling'].dropna())
                    df_tool = df_tool.tolist()
                    df_tool.insert(0,'All')
                    window.Element('C4').update(values = df_tool)
                else:
                    df_tool = pd.unique(df_SOT[df_SOT['Product'] == values['C1']]['Tooling'].dropna())
                    df_tool = df_tool.tolist()
                    df_tool.insert(0,'All')
                    window.Element('C4').update(values = df_tool)
            else:
                print('Choose a P/P first')
        elif event == 'Generate Full Plot':
            mf.GeneratePlot(values['C1'],values['C2'],values['C3'],values['C4'],values['C5'])
        elif event == 'Generate Plots and Tables':
            window2,df_final,df_average= table(values['C1'],values['C3'],values['C4'],values['_IN_'])
            mf.Generate7Plot(values['C1'],values['C3'],values['C4'],df_final,values['_THRES_'],values['_IN_'])
        elif event == 'Find Average':
            window3 = tableAverage(df_average)
        elif event == 'Show Full Raw Data':
            window4,df = tableRaw(values['C1'],values['C4'],values['_IN_'],True,False)
        elif event == 'Generate Raw Data':
            window4,df = tableRaw(values['C1'],values['C4'],values['_IN_'],False,False)
        elif event == 'Full Average Raw Data':
            window4,df = tableRaw(values['C1'],values['C4'],values['_IN_'],True,True)
        elif event == 'Average Raw Data':
            window4,df = tableRaw(values['C1'],values['C4'],values['_IN_'],False,True)
        elif event == 'Save File':
            mf.dataframe_save_file(r'' + values['_SAVE_'],df)
        elif event == 'Find Parts':
            windows4,df = PartsTable(PP = values['C5'],Tool = values['C4'],Category=values['C3'] )
        elif event == 'Find All Parts':
            window4,df = PartsTable()
        elif event == 'Save Parts to File':
            mf.dataframe_save_file(r'' + values['_SAVE_'],df,parts = True)
    window.close()