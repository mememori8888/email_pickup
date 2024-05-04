# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException,ElementClickInterceptedException,StaleElementReferenceException,TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common import exceptions
from selenium.webdriver.support.select import Select
import math
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome import service as fs
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from subprocess import CREATE_NO_WINDOW
import pandas as pd
import sys
import json
import re
import os
from collections import Counter
from csv import writer
import csv

import random
from webdriver_manager.chrome import ChromeDriverManager
import threading
#ランダム数の作成
randomC = random.uniform(1,7)

##chormeのオプションを指定
options = webdriver.ChromeOptions()
# options.add_argument("--headless")# ヘッドレスで起動するオプションを指定
options.page_load_strategy = 'eager'

options.add_argument("--incognito")
# options.add_argument("--no-startup-window")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1200,1200")
options.add_argument("--no-sandbox")
options.add_argument("--enable-javascript")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument("--enable-webgl")
options.add_argument('--enable-accelerated-2d-canvas')
options.add_argument("--renderer-process-limit=5")


desiredcapabilities = DesiredCapabilities.CHROME.copy()
desiredcapabilities['platform'] = "MAC"
desiredcapabilities['version'] = "106.0.5249.61"
desiredcapabilities['javascriptEnabled'] = True
path = os.getcwd()
CHROMEDRIVER = path + r'\chromedriver.exe'
new_driver = ChromeDriverManager().install()
chrome_service = fs.Service(executable_path=new_driver)
chrome_service.creationflags = CREATE_NO_WINDOW

driver = webdriver.Chrome(desired_capabilities=options.to_capabilities(),options=options,service=chrome_service)
driver.implicitly_wait(randomC)
#csvからcountact_urlを頂く

filename = r'C:\Users\user\Downloads\email.xlsx'
resultfile = r'C:\Users\user\Desktop\python\pickup_email.csv'


df = pd.read_excel(filename,sheet_name = 0, header=None)

url_list = df.loc[:,6].to_list()
name_list = df.loc[:,0].to_list()
address_list = df.loc[:,]
print(len(url_list))
count = len(url_list)

for i in range(0,count,1):
  
    #各パラメータ格納用リスト
    html = '-'
    get_email = '-'
    domain_url = '-'
    domain_list = []
    param_list = []
    url = url_list[i]
    name = name_list[i]
    get_email_text = ''
    email_page_text = ''
    
    #contact list
    contact_list = ['問合せフォーム']
     #email list
    email_list = []
    email_page_list = []
    email_column_name = 'メールアドレス'
    email_page_column_name = 'メールアドレスURL'
    
    # print(url)
    # print(len(url))
    print(name)
    print(url)
    print('{}番目のURLです'.format(i))
    #-の場合
    if url == '-' or url == '':
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
            writer = csv.writer(f)
            writer.writerow([name,'URLなし'])
            f.close()
        time.sleep(randomC)
        print('URLなし')
        continue
    elif 'fc2' in str(url):
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
            writer = csv.writer(f)
            writer.writerow([name,'fc2'])
            f.close()
        time.sleep(randomC)
        print('fc2')
        continue
    
    elif 'gogo.gs' in str(url):
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
            writer = csv.writer(f)
            writer.writerow([name,'gogo.gs'])
            f.close()
        time.sleep(randomC)
        print('fc2')
        continue
    
    else:

        
        try:    
            driver.get(url)
            time.sleep(randomC)
        except:
            with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
                writer = csv.writer(f)
                writer.writerow([name,'アクセス不可'])
                f.close()
            print('アクセス不可')
            continue
        
    try:
        html = driver.find_element(By.TAG_NAME,'html').text
    except:
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
            writer = csv.writer(f)
            writer.writerow([name,'アクセス不可'])
            f.close()

            print('アクセス不可')
            continue
    # print(html)
    
    #タグからリンクを取り出して、独自ドメインを取得して、独自ドメインが含まれているURLのテキストを取得して、メールアドレスがあるかしらべる。また問合せのテキストが含まれていたら、そのURLを問合せ用として取得する。
    link_list = []
    link = driver.find_elements(By.TAG_NAME,'a')
    for param in link:
        try:
            link_param = param.get_attribute('href')
            link_list.append(link_param)
        except:
            pass
    
    count_link_list = len(link_list)
    print('link_listの数は{}'.format(count_link_list))      
    print(link_list)
    #ドメイン取得
    domain = url.split('/')
    domain = domain[2].replace('www.','')
    print(domain)
    
    if 'mapion' in domain:
      
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
              writer = csv.writer(f)
              writer.writerow([name,url,'公式webなし'])
              f.close()
        
        continue
     
    elif 'navi' in domain:
      
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
              writer = csv.writer(f)
              writer.writerow([name,url,'公式webなし'])
              f.close()
        
        continue
      
    elif 'xn--zcklx7evic7044c1qeqrozh7c.com' in domain:
      
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
              writer = csv.writer(f)
              writer.writerow([name,url,'公式webなし'])
              f.close()
        
        continue
    
    elif 'biz-maps.com' in domain:
      
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
              writer = csv.writer(f)
              writer.writerow([name,url,'公式webなし'])
              f.close()
        
        continue
      
    elif 'nikkei' in domain:
      
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
              writer = csv.writer(f)
              writer.writerow([name,url,'公式webなし'])
              f.close()
        
        continue
    
    elif 'regus-office' in domain:
      
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
              writer = csv.writer(f)
              writer.writerow([name,url,'公式webなし'])
              f.close()
        
        continue
    
    elif 'ameblo' in domain:
      
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
              writer = csv.writer(f)
              writer.writerow([name,url,'公式webなし'])
              f.close()
        
        continue
    
    elif 'wikipedia' in domain:
      
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
              writer = csv.writer(f)
              writer.writerow([name,url,'公式webなし'])
              f.close()
        
        continue
    
    elif 'el.e-shops.jp' in domain:
      
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
              writer = csv.writer(f)
              writer.writerow([name,url,'公式webなし'])
              f.close()
        
        continue
    
    elif '.ekiten' in domain:
    
        with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
                writer = csv.writer(f)
                writer.writerow([name,url,'公式webなし'])
                f.close()
        
        continue
    
    
    else:
       pass
    # 独自ドメインが含まれているURLのテキストを取得して
   
    for domain_url in link_list:
        try:
            if str(domain) in str(domain_url):
                print('domainに引っかかってok')
                driver.get(domain_url)
            elif 'google' in str(domain_url):
                continue
            else:
                continue
        except:
            continue
        
        
        #html取得
        try:
            html = driver.find_element(By.TAG_NAME,'html').text
        except:
            html = '取得不可'
            continue
        
        #@の有無確認
        
        # email = [s for s in html if '@' in s]
        # count_email = len(email)
        
        if '@' in html:
            print('emailがありました。')
            
            email_page_list.append(domain_url)
            
            # #正規表現設定
            email = r'[0-9a-zA-Z\_\-\.]+@[0-9a-zA-Z\_\-\.]+\.[0-9a-zA-Z\_\-\.]+'
            email_pattern = re.compile(email)
            try:
                result = email_pattern.search(html)
                get_email = result.group()
                print(get_email)
                
                email_list.append(str(get_email))
                email_page_list.append(domain_url)
                
                
              
            except:
                get_email = '取得できなかった'
                print(get_email)
                email_list.append(str(get_email))
                email_page_list.append(domain_url)
              

        else:
            pass
                
        
        ##問合せのテキスト含まれていたら、そのURL取得
        if '送信' in html or '確認' in html or 'submit'.casefold() in html:
            contact_url = driver.current_url
            if 'contact' in contact_url or 'inquiry' in contact_url or 'form' in contact_url:
                print('コンタクトURLは{}'.format(contact_url))
                contact_list.append(contact_url)
            else:
                pass
        else:
            contact_url = '-'
            print('コンタクトURLは{}'.format(contact_url))
    
            
        
    # 重複の削除
    email_list = list(dict.fromkeys(email_list))
    email_page_list = list(dict.fromkeys(email_page_list))
    
    for param in email_list:
        
        get_email = '{},'.format(str(param))
        get_email_text = get_email_text + get_email
        
    for param in email_page_list:
        email_page = '{},'.format(str(param))
        email_page_text = email_page_text + email_page
    
    # print(contact_list)
    print(get_email_text)
    print(email_page_text)
    # param_list.extend(email_list)
    # param_list.extend(email_page_list)
    param_list = [name,url]
    param_list.append(email_column_name)
    param_list.append(get_email_text)
    param_list.append(email_page_column_name)
    param_list.append(email_page_text)
    #email_listとcontact_listをparam_listにいれて、csvに出力
    #重複の削除
    contact_list = list(dict.fromkeys(contact_list)) 
    contact_count = len(contact_list)
    if contact_count == 1:
        contact_list = ['問合せフォーム','-']
        param_list.extend(contact_list)
    else:
        param_list.extend(contact_list) 
              
    print(param_list)
        
    with open(resultfile,mode='a',encoding='cp932',newline='', errors='replace') as f:
        writer = csv.writer(f)
        writer.writerow(param_list)
        f.close()
    
    continue
        
        
