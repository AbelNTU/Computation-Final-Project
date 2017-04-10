
# coding: utf-8
# 3/27
# 主題：
可能主題：用資訊預測每一季股票
    資料：(1)財務報表
            a.資產負債表(v)
            b.
         (2)新聞、政策
         (3)大盤走向
         (4)每日交易
         (5)大戶狀況
         (6)匯率
         (7)銀行利率
         (8)除權息
        *(9)每月營收
    股票分析常用3種分析模式(1)基本面(2)技術面(3)籌碼面
                (1)基本面：
                    優點：巴菲特再用(誤)、能夠反映股價與公司實際價值、
                    缺點：每一季才有新的財報，往往跟不上股票價值變化的速度
                    適合的人：
                (2)技術面：
                    優點：
                    缺點：
                    適合的人：
                (3)籌碼面：
                    優點：
                    缺點：
                    適合的人：
                       
 原因： 一張股票在現在常被當成一種商品在買賣，很少人真的把作為股東這件事當成一回事，所以也就造成很多人著重在於技術面分析而忽視了基本面，
 所以我們將盡可能地蒐集所有影響或可判斷一個公司價值的資訊，並運用現在的機器學習法()去預測股票的走向，目標是
 (1)預測準度
 (2)實際的看股票的走向是否與財務報表給出的資訊有正相關
 (3)比較各項資訊的重要程度
 
 
 Topic2:抓出財務報表有問題的公司
 原因：如果公司負責人為了吸引資金而作假財務報表，常會讓投資人損失慘重。我們將依照各個產業不同，分析產業在每一季大概的趨勢，並且挑出偏離值
 作為投資人可以更深入了解的參考
 
 
# 
# |考試|Date|我的預期|Done?|宗穎預期|Done?|Extra|Problem|Solve?|
# |---|----|---------|-----|-----|-------|------|
# ||3/20||-||-|-|-|-|
# ||3/27||-||-|-|-|-|
# ||4/03|蒐集全部資料|-|改善資料+寫程式畫圖|-|-|-|-|
# ||4/10|找分類器|-||-|-|-|-|
# |代導、機率期中考|4/17|Debug|-|還可以取哪些資訊|-|-|-|-|
# ||4/24|測試多家數據|-|看數據，想理由|-|-|-|-|
# |密碼學期中|5/01|Debug|-|看數據想理由|-|-|-|-|
# ||5/08|預測數據|-|如果結果不準，調整|-|-|-|-|
# ||5/15|Debug|-|在不準，再調整|-|-|-|-|
# ||5/22|選定幾家公司|-|取幾家公司(Why?)|-|-|-|-|
# ||5/29|檢討修改|-|最後調整|-|-|-|-|
# ||6/05|再跑一次|-|把結果及可能原因整理|-|-|-|-|
# |代導期末考|6/12|做ptt加影片|-|做ptt加影片|-|-|-|-|
# |一堆期末考|6/19|放暑假囉!|-|撈個毛?|-|-|-|-|

# # i = 模式(3=資產負債，4=綜合損益，5=現金流量，6=每月營收)

# ## 模組

# In[14]:

from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import selenium
import time
import re
import csv
from pathlib import Path
import os
from pandas import DataFrame
import shutil


# In[20]:

#加速瀏覽
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException 


# # 瀏覽器

# In[3]:

driver = webdriver.Chrome('Desktop/ntumath/chromedriver')


# # 防止爬蟲中斷，檢查table出現了沒

# In[40]:

def table_is_present(xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True


# # 抓三報表

# In[4]:

def scrath(number,i):
    driver.find_element_by_id('co_id').send_keys(number)
    for year in range(102,106,1):
        driver.find_element_by_id('year').send_keys(year)
        for season in range(4):
            try:
                driver.find_element_by_id('season').send_keys(season+1)
                driver.find_element_by_xpath("//input[@type='button' and @value=' 查詢 ']").click()
                table_path = '//*[@id="table01"]/center/table[2]'
                try:
                    WebDriverWait(driver,3).until(lambda driver: driver.find_element_by_xpath(table_path))
                except TimeoutException:
                    print("網頁載入過久！")
                bf = BeautifulSoup(driver.page_source,'html.parser')
                table = bf.find_all('table',class_="hasBorder")[0]
                rows = table.find_all('tr')
                csvfiles = open("Desktop/ntumath/computation/"+number+"/"+number+str(year)+"."+str(season+1)+".csv",
                                'wt',newline='',encoding='utf-8')
                writer = csv.writer(csvfiles)
                try:
                    for row in range(len(rows)):
                        csvRow = []
                        if row <= 2:
                            pass
                        else:
                            for cell in rows[row].find_all(['td','th']):
                                csvRow.append(cell.get_text().strip())
                        if row == 3 :
                            if i != 5:
                                csvRow = csvRow[1:3]
                            else:
                                csvRow = csvRow[1:2]
                        else:
                            if i != 5:
                                csvRow = csvRow[:3]
                            else:
                                csvRow = csvRow[:2]
                        writer.writerow(csvRow)
                finally:
                    csvfiles.close()
            except:
                pass
        driver.find_element_by_id('year').clear()


# # 抓每月營收

# In[5]:

def crawl_earn_per_month(number):
    [year_new,season_new] = get_start(number,6)
    url = 'http://mops.twse.com.tw/mops/web/t05st10_ifrs'
    driver.get(url)
    Select(driver.find_element_by_id('isnew')).select_by_index(1)
    driver.find_element_by_id('co_id').clear()
    driver.find_element_by_id('co_id').send_keys(number)
    driver.find_element_by_id('year').clear()
    for year in range(year_new,106,1):
        driver.find_element_by_id('year').send_keys(year)
        for month in range(1,13,1):
            try:
                if month != 11:
                    driver.find_element_by_id('month').send_keys(month)
                else:
                    driver.find_element_by_id('month').send_keys(1)
                driver.find_element_by_xpath("//input[@type='button' and @value=' 查詢 ']").click()
                
                time.sleep(3)
                bf = BeautifulSoup(driver.page_source,'html.parser')
                table = bf.find_all('table',class_="hasBorder")[0]
                rows = table.find_all('tr')
                csvfiles = open("Desktop/ntumath/computation/"+number+"/"+number+str(year)+"."+str(month)+".csv",
                                'wt',newline='',encoding='utf-8')
                writer = csv.writer(csvfiles)
                try:
                    for row in range(len(rows)):
                        csvRow = []
                        for cell in rows[row].find_all(['td','th']):
                            csvRow.append(cell.get_text().strip())
                        if row == 0:
                            writer.writerow([])
                            csvRow = csvRow[1:2]
                            writer.writerow(csvRow)
                        elif row not in [1,5]:
                            pass
                        else:
                            writer.writerow(csvRow)
                finally:
                    csvfiles.close()
            except:
                pass
            time.sleep(0.2)
        driver.find_element_by_id('year').clear()


# # 取得開始的年、季

# In[6]:

def get_start(number,i):
    for year in range(102,106,1):
        if i <= 5:
            for season in range(1,5,1):
                try:
                    my_file = Path("Desktop/ntumath/computation/"+number+"/"+number+str(year)+"."+str(season)+".csv")
                    if my_file.is_file():
                        year_new = year
                        season_new = season
                        return [year_new,season_new]
                    else:
                        pass
                except:
                    pass
        elif i == 6:
            for month in range(1,13,1):
                try:
                    my_file = Path("Desktop/ntumath/computation/"+number+"/"+number+str(year)+"."+str(month)+".csv")
                    if my_file.is_file():
                        year_new = year
                        month_new = month
                        return [year_new,month_new]
                    else:
                        pass
                except:
                    pass
            


# # 合併

# In[7]:

def combine_1(number,i):
    if i <= 5:
        [year_final,season_final] = get_start(number,i)
        [year_new,season_new] = get_start(number,i)
        data = pd.read_csv("Desktop/ntumath/computation/"+number+"/"+number+str(year_new)+"."+str(season_new)+".csv")
    elif i == 6:
        [year_final,month_final] = get_start(number,i)
        [year_new,month_new] = get_start(number,i)
        data = pd.read_csv("Desktop/ntumath/computation/"+number+"/"+number+str(year_new)+"."+str(month_new)+".csv")
    data = data[~data.index.duplicated(keep='last')]
    for year in range(year_new,106,1):
        try:
            if i <= 5:
                if year == year_new and season_new != 4:
                    for season in range(season_new+1,5,1):
                        data_right = pd.read_csv("Desktop/ntumath/computation/"+number+"/"+number+str(year)+"."+str(season)+".csv")
                        data_right = data_right[~data_right.index.duplicated(keep='last')]
                        data = pd.merge(data,data_right,right_index=True,left_index=True,how='outer',sort=False)
                        season_final += 1
                else:
                    if season_new == 4:
                        pass
                    else:
                        for season in range(1,5,1):
                            data_right = pd.read_csv("Desktop/ntumath/computation/"+number+"/"+number+str(year)+"."+str(season)+".csv")
                            data_right = data_right[~data_right.index.duplicated(keep='last')]
                            data = pd.merge(data,data_right,right_index=True,left_index=True,how='outer',sort=False)
                            season_final += 1
                if year != 105:
                    year_final += 1
                    season_final = 0
            elif i == 6:
                if year == year_new and month_new != 12:
                    for month in range(month_new+1,13,1):
                        data_right = pd.read_csv("Desktop/ntumath/computation/"+number+"/"+number+str(year)+"."+str(month)+".csv")
                        data = pd.merge(data,data_right,right_index=True,left_index=True,how='outer',sort=False)
                        month_final += 1
                else:
                    if month_new == 12:
                        pass
                    else:
                        for month in range(1,13,1):
                            data_right = pd.read_csv("Desktop/ntumath/computation/"+number+"/"+number+str(year)+"."+str(month)+".csv")
                            data = pd.merge(data,data_right,right_index=True,left_index=True,how='outer',sort=False)
                            month_final += 1
                if year != 105:
                    year_final += 1
                    month_final = 0
        except:
            break
        data = data.dropna(how='all')
    if i <= 5:
        data.to_csv("Desktop/ntumath/computation/"+number+"/"+number+"("+str(year_new)+"."+str(season_new)+"~"+str(year_final)+
                        "."+str(season_final)+")"+str(i)+".csv")
    elif i == 6:
        data.to_csv("Desktop/ntumath/computation/"+number+"/"+number+"("+str(year_new)+"."+str(month_new)+"~"+str(year_final)+
                        "."+str(month_final)+")"+'EPM'+".csv")


# # 創資料夾+合併

# In[8]:

def catch_csv(company_number,i):
    company_number = str(company_number)
    Path_folder =  Path("Desktop/ntumath/computation/"+company_number)
    if Path_folder.is_dir():
        pass
    else:
        os.makedirs("Desktop/ntumath/computation/"+company_number)
    if i <= 5:
        scrath(company_number,i)
    elif i == 6:
        crawl_earn_per_month(company_number)
    combine_1(company_number,i)


# # 爬資料(三報表+每月營收)

# In[9]:

def crawl_number(number):
    for i in range(3,7,1):
        if i <= 5:
            url="http://mops.twse.com.tw/mops/web/t164sb0"+str(i)
            driver.get(url)
            Select(driver.find_element_by_id('isnew')).select_by_index(1)
            driver.find_element_by_id('co_id').clear()
            driver.find_element_by_id('year').clear()
            driver.find_element_by_id('season').send_keys(0)
            catch_csv(number,i)
            time.sleep(2)
        if i == 6:
            driver.find_element_by_id('year').clear()
            driver.find_element_by_id('co_id').clear()
            catch_csv(number,i)
        


# # 刪除不需要的檔案

# In[10]:

def clear(number):
    for year in range(102,106,1):
        for month in range(1,13,1):
            try:
                os.remove('Desktop/ntumath/computation/'+str(number)+'/'+str(number)+str(year)+'.'+str(month)+'.csv')
            except:
                pass


# In[11]:

index = pd.read_csv('Desktop/ntumath/computation/index.csv',index_col=0)
category = index.T.index


# # 各種分類公司開抓

# In[12]:

def category_to_crawl():
    index = pd.read_csv('Desktop/ntumath/computation/index.csv',index_col=0)
    category = index.T.index
    for i in range(1,len(category)+1,1):
        cate = category[i]
        path_of_folder = Path('Desktop/ntumath/computation/'+cate)
        if path_of_folder.is_dir():
            pass
        else:
            os.makedirs("Desktop/ntumath/computation/"+cate)
        schedule = list(map(int,list(index.iloc[:,i].dropna())))
        try:
            for company_num in schedule:
                crawl_number(company_num)
                clear(company_num)
                shutil.move('Desktop/ntumath/computation/'+str(company_num),'Desktop/ntumath/computation/'+cate+'/'+str(company_num))
        except:
            pass


# # 抓取每月平均價格

# In[57]:

def crawl_month_stock(number):
    ##[start_year, start_month] = get_start(number,6)
    [start_year, start_month] = [102,1]
    url = 'http://www.tse.com.tw/ch/trading/exchange/STOCK_DAY_AVG/STOCK_DAY_AVGMAIN.php'
    path = '//*[@id="main-content"]/table'
    driver.get(url)
    driver.find_element_by_name('CO_ID').send_keys(number)
    #select_year = Select(driver.find_element_by_name('query_year'))
    #select_month = Select(driver.find_element_by_name('query_month'))
    ##假設都能抓到105年12月
    csvfiles = open("Desktop/ntumath/computation/每月股價/"+str(number)+"股價("+str(start_year)+'.'+str(start_month)+
                    '~105.12).csv',
                    'wt',newline='',encoding='utf-8')
    writer = csv.writer(csvfiles)
    for year in range(start_year,106,1):
        driver.find_element_by_name('query_year').send_keys(year)
        for month in range(1,13,1):
            if month != 11:
                driver.find_element_by_name('query_month').send_keys(month)
            else:
                driver.find_element_by_name('query_month').send_keys(1)
            driver.find_element_by_name('query-button').click()
            while table_is_present(path) != True:
                time.sleep(1)
                driver.find_element_by_name('query-button').click()
            bf = BeautifulSoup(driver.page_source,'html.parser')
            table = bf.find_all('table',border="1")[0]
            price_per_month = table.find_all('td')[-1].get_text()
            if month == 1:
                rows = [str(year)+'年1月',price_per_month]
            else:
                rows = [str(month)+'月',price_per_month]
            writer.writerow(rows)
            time.sleep(1)
    csvfiles.close()


# In[13]:

category_to_crawl()


# # TEST region

# In[45]:

driver = webdriver.Chrome('Desktop/ntumath/chromedriver')


# In[46]:

driver.get('http://www.tse.com.tw/ch/trading/exchange/STOCK_DAY_AVG/STOCK_DAY_AVGMAIN.php')
select = Select(driver.find_element_by_name('query_month'))

select.select_by_visible_text('3')


# In[47]:

select.select_by_visible_text('4')


# In[58]:

crawl_month_stock(2330)


# In[ ]:



