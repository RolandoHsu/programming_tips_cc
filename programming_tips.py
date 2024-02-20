#!/usr/bin/env python
# coding: utf-8

# In[1]:


## import the library
# Basic 
import datetime

start = datetime.datetime.now()
import os # check the working directory
from datetime import date
from datetime import timedelta
import time
import pandas as pd
import numpy as np
import random

# website crawler
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys # 用來執行 Keys.ENTER
from bs4 import BeautifulSoup
import requests
requests.packages.urllib3.disable_warnings() ## 禁用安全请求警告 https://blog.csdn.net/qq_42739440/article/details/90754558
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException ## 偵測是否出現 time out 的訊息
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium import webdriver

## 不顯示爬蟲頁面
chrome_options = Options()
# chrome_options.add_argument("--headless")
chrome_options.add_argument('--disable-software-rasterizer') # 
chrome_options.add_argument('--no-sandbox') # 使用最高權限 解決DevToolsActivePort檔案不存在的報錯
chrome_options.add_argument('blink-settings=imagesEnabled=false') # 不載入圖片,提升速度
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('lang=zh_CN.UTF-8')

import logging

## 抓取電腦使用者
import getpass

import sys
import traceback
import win32gui
import win32con


# In[2]:


class create_programming_tips_tool():
    def __init__(self):
        
        self.path = 'C:/Users/esb21774/Desktop/bodun/17.FUN/programming_tips'
        self.programming_tips_hist = self.path+'/programming_tips_hist.xlsx'
        self.programming_topic = self.path+'./programming_topic.xlsx'
        self.genie_account = 'OANT\esb21774'
        self.genie_password = 'password'
        self.api_key = 'personal_api_key'
        self.chat_sn = '815'
    
    def get_prompt(self):
    
        ## 先篩選最近一次選擇的主題以及工具，做排除，避免取回的都一樣
        programming_tips_hist = pd.read_excel(self.programming_tips_hist).sort_values('datetime', ascending = False).reset_index(drop = True)
        programming_tips_hist = programming_tips_hist.tail(1)[['tool', 'type']]

        ## 透過excel隨機篩選需要的主題
        programming_topic = pd.read_excel(self.programming_topic)
        programming_topic = programming_topic[(programming_topic['tool']!=programming_tips_hist['tool'].values[0])&(programming_topic['type']!=programming_tips_hist['type'].values[0])]
        random_rows = programming_topic.sample(n=1).reset_index(drop = True)

        ## 產製本次的prompt
        prompt = f"#zh-tw 你現在是一個{random_rows['tool'].values[0]}大師，請提供一個{random_rows['type'].values[0]}的使用技巧，此技巧可與基礎使用技巧、效能精進、程式碼優化、使用必須注意事項有關，請將回覆限制在100字內"

        return [prompt, random_rows]
    
    def get_programming_tips(self, prompt):
    
        driver = webdriver.Chrome(self.get_chrome_driver(), options = chrome_options)
        driver.set_page_load_timeout(10)
        driver.implicitly_wait(10)

        ## 設定要探訪的網址 
        url = '企業版CHATGPT網址'
        driver.implicitly_wait(5)
        driver.get(url)

        ## 選擇玉山銀行
        driver.find_element(by = By.XPATH, value = '//*[@id="bySelection"]/div[2]/img').click()
        driver.implicitly_wait(2)

        ## 輸入帳密
        selectA = driver.find_element(by = By.XPATH, value = '//*[@id="userNameInput"]')
        selectA.send_keys(self.genie_account)
        time.sleep(0.5)
        selectB = driver.find_element(by = By.XPATH, value = '//*[@id="passwordInput"]')
        selectB.send_keys(self.genie_password)
        driver.find_element(by = By.XPATH, value = '//*[@id="submitButton"]').click()
        driver.implicitly_wait(3)

        # 點選確認使用手冊
        driver.find_element(by = By.CLASS_NAME, value = 'uiStyle.btnGreen.sizeS.relative').send_keys(Keys.ENTER)
        time.sleep(0.5)

        ## 輸入prompt
        driver.find_element(by = By.CLASS_NAME, value = 'el-textarea__inner').click()
        selectC = driver.find_element(by = By.CLASS_NAME, value = 'el-textarea__inner')
        selectC.send_keys(prompt[0])
        driver.implicitly_wait(2)

        ## 開始生成文字
        driver.find_element(by = By.XPATH, value = '//*[@id="app"]/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div/button').send_keys(Keys.TAB)
        driver.find_element(by = By.XPATH, value = '//*[@id="app"]/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div/button').send_keys(Keys.ENTER)
        time.sleep(5)

        try:
            ## 擷取 頁面所有 html 資訊 
            html_allinformation = driver.page_source
            soup_allinformation = BeautifulSoup(html_allinformation, "html.parser")
            ## 抓生成的文字
            allinformation = soup_allinformation.find(class_= "prose prose-slate max-w-none text-base text-slate-600 leading-relaxed").text
        except:

            time.sleep(5)       
            ## 擷取 頁面所有 html 資訊 
            html_allinformation = driver.page_source
            soup_allinformation = BeautifulSoup(html_allinformation, "html.parser")
            ## 抓生成的文字
            allinformation = soup_allinformation.find(class_= "prose prose-slate max-w-none text-base text-slate-600 leading-relaxed").text
            driver.close()

        allinformation = allinformation.replace("\n", "")



        ## 傳送team+
        information = []
        information.append('📣📣📣程式小助手📣📣📣')
        information.append(f"📗今日主題: {prompt[1]['tool'][0]} / {prompt[1]['type'][0]}")
        information.append('➖➖➖➖➖➖➖➖➖➖➖')
        information.append(f"💬 {allinformation}")

        url = "公司通訊軟體api網址"

        OutputDataSet = ("\n".join(str(x) for x in information))

        myobj = {
            "account": getpass.getuser().upper(),
            "api_key": self.api_key,
            "chat_sn": self.chat_sn,
            "content_type": "1",
            "msg_content": OutputDataSet,
                }

        x = requests.post(url, data = myobj)

        ## 將最新的紀錄存入excel中
        new_row = pd.DataFrame({
            'datetime': datetime.datetime.now(),
            'tool': prompt[1]['tool'][0],
            'type': prompt[1]['type'][0],
            'prompt': prompt[0],
            'response': allinformation}, 
            index=[0])

        programming_tips_hist = pd.read_excel(self.programming_tips_hist)
        programming_tips_hist = programming_tips_hist.append(new_row).reset_index(drop = True)
        programming_tips_hist.to_excel(self.programming_tips_hist, index=False)

        return allinformation


# In[3]:


def main():
    try:
        programming_tips_tool = create_programming_tips_tool()
        prompt = programming_tips_tool.get_prompt()
        programming_tips_tool.get_programming_tips(prompt)
    except Exception as e:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        lines = traceback.format_exception(exc_type, exc_value, exc_traceback)
        line = f"測試"+'\n'+''.join('!! ' + line for line in lines)
        for i in line.split('\n'):
            print(i)
        a = input('確認上述log')


# In[4]:


if __name__ == '__main__':
    main()


# In[ ]:




