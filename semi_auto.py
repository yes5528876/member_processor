from tracemalloc import start
from unicodedata import name
import pandas as pd
import numpy as np
# import matplotlib.pyplot as plt
from collections import Counter
import os,sys
import string
import csv
import glob
# import requests
import webbrowser
from tkinter import *   # from tkinter import Tk for Python 3.x
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import *
import warnings
from bs4 import BeautifulSoup
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Border,Side,Alignment
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re
import pyautogui
from pyautogui import *
import keyboard
import threading
warnings.filterwarnings("ignore")


path_of_source_excel=output_path=itri_id=itri_password=民國年=""
df=df_temp=log_DataFrame=log_DataFrame=pd.DataFrame()
index=0
last_index=0
row=None
driver=None
auto_enable=False

def key_listener():
    print("按下F1切換自動/半自動模式")
    global auto_enable
    while True:
        if keyboard.is_pressed('f1'):
            auto_enable=not auto_enable
            if auto_enable:
                print("自動查詢已預定，查詢完目前的人後電腦將自動查詢")
            else:
                print("手動查詢已預定，請等待目前查詢進度結束後，即可按下方向鍵(down)查詢下個人")
            time.sleep(1)
        time.sleep(0.01)

# 開啟新線程
t = threading.Thread(target=key_listener)
t.start()

def run():

    global output_path,民國年

    localtime=time.localtime()
    output_path = time.strftime("輸出/%m月%d日%H點%M分%S秒輸出結果/", localtime)
    year=time.strftime("%Y",localtime)
    民國年=str(int(year)-1911)
    os.makedirs(output_path)


    # print("從上次中斷地方開始查詢?  1.Yes   0.No")
    # recovery=input("請輸入數字後按Enter:")
    # while(recovery!='1' and recovery!='0'):
    #     print("輸入錯誤，請重新輸入0或1")

    recovery='0'

    if(recovery=='1'):
        ask_id_and_pwd()
        load_json()
        login_process()
        search_all()
    if(recovery=='0'):
        ask_id_and_pwd()
        ask_file_path()
        login_process()
        search_all()

    print("名單製造完成，請至 "+output_path+" 資料夾查詢")
    

def load_json():
    global path_of_source_excel,output_path,last_index,df,df_temp
    with open("user_input.json", "r") as file:
        user_input = json.load(file)
    path_of_source_excel=user_input['path_of_source_excel']
    output_path=user_input['output_path']
    last_index=user_input['last_index']
    df = pd.read_excel(path_of_source_excel)
    df_temp=pd.read_excel(output_path+'temp.xlsx')


def save_json():
    global output_path,path_of_source_excel,index
    data = {
        "path_of_source_excel": path_of_source_excel,
        "output_path": output_path,
        "last_index": index
    }
    with open("user_input.json", "w") as file:
        json.dump(data, file)


def ask_id_and_pwd():
    print('請在小框框輸入帳號密碼')
    root = Tk()
    root.withdraw
    root.title('Login ITRI')
    def button_event():
        global itri_id,itri_password
        itri_id=myentry.get()
        itri_password=myentry2.get()
        root.quit()
    mylabel = Label(root, text='Name:')
    mylabel.grid(row=0, column=0)
    myentry = Entry(root)
    myentry.grid(row=0, column=1)

    mylabel2 = Label(root, text='Password:')
    mylabel2.grid(row=1, column=0)
    myentry2 = Entry(root)
    myentry2.grid(row=1, column=1)

    mybutton = Button(root, text='Send',command=button_event)
    mybutton.grid(row=2, column=1)

    root.mainloop()


def ask_file_path():
    global path_of_source_excel,df
    root1 = Tk()
    root1.withdraw()
    path_of_source_excel = filedialog.askopenfilename()
    df = pd.read_excel(path_of_source_excel)


def login_process():
    global driver,itri_id,itri_password
    try:
        driver=webdriver.Chrome('.\chromedriver.exe')
    except Exception as e:
        print(e)
        print("chromedriver無法開啟!  請將上方英文內容拍照傳給開發人員")
        sys.exit(0)
    # driver.implicitly_wait(20)
    driver.get("https://empfinder.itri.org.tw/WebPage/ED_QueryIndex.aspx")
    time.sleep(3)
    # driver.find_element_by_id("idToken1").send_keys(itri_id)
    driver.find_element(By.ID,"idToken1").send_keys(itri_id)
    # driver.find_element_by_id("idToken2").send_keys(itri_password)
    driver.find_element(By.ID,"idToken2").send_keys(itri_password)
    # driver.find_element_by_id("loginButton_0").click()
    driver.find_element(By.ID,"loginButton_0").click()
    time.sleep(3)

def search_all():
    global log_DataFrame,index,row,df,last_index,auto_enable
    init_log={"事件":[]}
    log_DataFrame=pd.DataFrame(init_log)
    for index,row in df.iterrows():
        if(index<last_index):
            continue
        if(not auto_enable):
            while(not keyboard.is_pressed('down')):
                time.sleep(0.01)
        search_one()
        save_json()

def search_one():
    global driver,df,log_DataFrame,index,row,df_temp
    print(str(index)+'/'+str(len(df.axes[0])))
    try:
        driver.get("https://empfinder.itri.org.tw/WebPage/ED_QueryIndex.aspx")
        time.sleep(3)
        # driver.find_element_by_id("wuc_queryConditions_tbx_empno").click()
        driver.find_element(By.ID,"wuc_queryConditions_tbx_empno").click()
        # driver.find_element_by_id("wuc_queryConditions_tbx_empno").clear()
        driver.find_element(By.ID,"wuc_queryConditions_tbx_empno").clear()
        # driver.find_element_by_id("wuc_queryConditions_tbx_empno").send_keys(str(df['工號'][index])) #str(df['工號'][index])
        driver.find_element(By.ID,"wuc_queryConditions_tbx_empno").send_keys(str(df['工號'][index]))
        # driver.find_element_by_id("tbempno").send_keys("A60206")
        # driver.find_element_by_id("btn_search").click()
        driver.find_element(By.ID,"btn_search").click()
        time.sleep(3)
    except Exception as e: 
        print("Driver Error Occur. Restart the driver")
        print(e)
        login_process()
        search_one()

    try:
        r = driver.page_source
        soup = BeautifulSoup(r, 'html.parser')
    except Exception as e: #UnexpectedAlertPresentException
        print(e)
        print(df['姓名'][index]+"  已離職  ")
        log_DataFrame=log_DataFrame.append({"事件":df['姓名'][index]+"  已離職  " },ignore_index=True)
        # df=df.drop(index=index,axis=0)
        return
    extract=soup.find_all("td")

    # if(len(extract)==0):   #len(extract)==0 means he's quit
    #     print(df['姓名'][index]+"  已離職  ")
    #     log_DataFrame=log_DataFrame.append({"事件":df['姓名'][index]+"  已離職  " },ignore_index=True)
    #     df=df.drop(index=index,axis=0)
    #     continue

    #----------------The resign detection function-------------------
    try:
        office=extract[4].getText().replace('\n','')
    except Exception as e:        
        print(e)
        print(df['姓名'][index]+"  已離職  ")
        log_DataFrame=log_DataFrame.append({"事件":df['姓名'][index]+"  已離職  " },ignore_index=True)
        # df=df.drop(index=index,axis=0)
        return
    #----------------End of the resign detection function-------------------
    if(df['辦公室'][index]!=office):
        print(df['姓名'][index]+"  辦公室更改為  "+ office )
        log_DataFrame=log_DataFrame.append({"事件":df['姓名'][index]+"  辦公室更改為  "+ office },ignore_index=True)
        # df['辦公室'][index]=office
        row['辦公室']=office

    unit = extract[0].getText().strip()
    if(df['單位'][index]!=unit):
        print(df['姓名'][index]+"  調轉單位到  "+ unit )
        log_DataFrame=log_DataFrame.append({"事件":df['姓名'][index]+"  調轉單位到  "+ unit  },ignore_index=True)
        # df['單位'][index]=unit
        row['單位']=unit
    df_temp=df_temp.append(row)
    df_temp.to_excel(output_path+'temp'+'.xlsx', index=False)
    log_DataFrame.to_excel(output_path+'temp'+'_log.xlsx', index=False)


if __name__ == '__main__':
    run()