import xlrd
import pyautogui
import webbrowser
import time
from datetime import datetime
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium import Options

# options = Options()
# options.add_argument('--headless')
# options.add_argument('--profile-directory=Default') 
# browser = webdriver.Chrome(options=options,executable_path='chromedriver.exe')

# options.add_argument('--profile-directory=Default') 

import pandas as pd
# driver = webdriver.Chrome('C:\\Users\\Nilesh\\Desktop\\online_bot\\drivers\\chromedriver.exe')
# pyautogui.hotkey('ctrl', 'w')
time.sleep(4)
def join():
    time.sleep(5)
    pyautogui.hotkey('ctrl', 'd')
    pyautogui.hotkey('ctrl', 'e')

    time.sleep(3)
    # print("jsdadsad")
    # driver.find_element(By.XPATH,'//a[@class="NPEfkd RveJvd snByac"]').click()
    #print(type(elem))
    #elem.click()
    pyautogui.click(1328, 586)


path="D:\\online_bot.xls"
data=pd.read_excel(path)
#print(data.value[0,1])
#print(data)
x=datetime.today().isoweekday()
wb=xlrd.open_workbook(path)
sheet=wb.sheet_by_index(0)

timear=data.columns.values.tolist()
#print(timear)

format='%H:%M'
datetime_str1=pd.to_datetime(timear[1])
datetime_str2=pd.to_datetime(timear[2])
timedelta=(datetime_str2-datetime_str1)
print(datetime_str1)
print(timedelta.total_seconds())



link=sheet.cell_value(x,1)

datetime_str1=pd.to_datetime(timear[1])
datetime_str2=pd.to_datetime(timear[2])

# while True:
#     now = datetime.now()
#     if (now.strftime("%H%M%S") >= "071200" and now.strftime("%H%M%S") <= "082600"):
#         print("dsa")
#         time.sleep(923)
#         break
#     print("dks")
#     print(now)

while True:
    now = datetime.now()
    if (now>=datetime_str1 and now <= datetime_str2):
        webbrowser.get("C:/Program Files/Google/Chrome/Application/chrome.exe %s").open_new(link)
        join()
        print("first class")
        timedelta=(datetime_str2-datetime_str1)
        print("Slept for ",timedelta)
        time.sleep(timedelta.total_seconds()-5)
        print("Wokeup")
        break
    time.sleep(4)



for i in range(3,sheet.ncols,2):
    print(i)
    datetime_str1=pd.to_datetime(timear[i])
    datetime_str2=pd.to_datetime(timear[i+1])
    timedelta=(datetime_str2-datetime_str1)
    # print(datetime_str2)
    # print(datetime_str1)
    link=sheet.cell_value(x,i)
    while True:
        now = datetime.now()
        if (now>=datetime_str1 and now <= datetime_str2):
            pyautogui.hotkey('ctrl', 'w')
            webbrowser.get("C:/Program Files/Google/Chrome/Application/chrome.exe %s").open_new(link)
            join()  # calling join function
            time.sleep(timedelta.total_seconds()-5)
            break
        time.sleep(4)