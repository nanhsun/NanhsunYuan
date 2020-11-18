from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait 
import pandas as pd
from datetime import datetime 
import calendar
import numpy as np
# 若檔案夾位置有更換，請改每個r""裡的檔案位置到新的檔案位置
def nparray_save_file(var):
    """
    Saves numpy array into any file
    """
    fileloc = pd.read_csv(r"\\TAFS\Public\Alan Yuan\Xray\FileLoc.txt").columns.format()[0]
    np.savetxt(fileloc, var, fmt = "%s", delimiter = ",", encoding = "utf-8-sig")
    return

if __name__ == "__main__":
    driver = webdriver.Chrome(executable_path=r"\\TAFS\Public\Alan Yuan\Xray\chromedriver.exe")
    page = driver.get("https://sps16.itg.ti.com/sites/TITLAssyOps/Engr/_layouts/15/start.aspx#/Lists/XRay%20Rejected%20Lot%20Tracking%20System/AllItems.aspx")
    table = WebDriverWait(driver, 15).until(lambda d: d.find_element_by_id("{4BC469D3-E8CE-42E8-8A51-9737F2043ABA}-{6E8A2AAA-DDC5-41BC-BD7D-054275BB9FB7}"))
    content = table.get_attribute('outerHTML')
    df = pd.read_html(content)[0]
    driver.quit()
    month = input('Input a month (in numbers): ')
    df['Date'] = df['Date'].map(lambda x: datetime.strptime(str(x), '%m/%d/%Y'))
    df_month = df[df['Date'].dt.month == int(month)]
    DefectModes = pd.unique(df_month['Defect Mode ．'])
    Count = np.array([[calendar.month_name[int(month)],'Num of Occurrences','Engr Comment','Percentage','Disposition','Percentage','Root Cause','Percentage','Improvement Action','Percentage']])
    for DMode in DefectModes:
        Mode_df = df_month[df_month['Defect Mode ．']==DMode]
        Amount = Mode_df.shape[0]
        EC = Mode_df['Engr Comment ．'].notna().sum()
        ECP = Mode_df['Engr Comment ．'].notna().sum()/Mode_df['Engr Comment ．'].shape[0]*100
        Dis = Mode_df['Disposition　．'].notna().sum()
        DisP = Mode_df['Disposition　．'].notna().sum()/Mode_df['Disposition　．'].shape[0]*100
        RC = Mode_df['Root Cause'].notna().sum()
        RCP = Mode_df['Root Cause'].notna().sum()/Mode_df['Root Cause'].shape[0]*100
        IA = Mode_df['Improvement Action'].notna().sum()
        IAP = Mode_df['Improvement Action'].notna().sum()/Mode_df['Improvement Action'].shape[0]*100
        temp = np.array([[DMode,Amount,EC,str(ECP)+'%',Dis,str(DisP)+'%',RC,str(RCP)+'%',IA,str(IAP)+'%']])
        Count = np.vstack((Count,temp))
    nparray_save_file(Count)