from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
 
import re # regex 

import xlsxwriter # import xlsxwriter module

options = Options()
options.add_argument("--disable-notifications")
options.add_argument('--headless') # Chrome無頭騎士模式
 
chrome = webdriver.Chrome('./broswerDrivers/chromedriver.exe', chrome_options=options)
chrome.get("https://structurednotes-announce.tdcc.com.tw/Snoteanc/apps/bas/BAS210.jsp")

########################################################################
########################################################################
########################################################################
"""
輸入查詢條件
"""
# 商品國內發行日
iDateStart = chrome.find_element_by_id("iDateStart")
iDateEnd = chrome.find_element_by_id("iDateEnd")
searchBtn = chrome.find_element_by_xpath('//input[@value="查詢"]') # 查詢按鈕

iDateStart.send_keys('2021/09/01')
iDateEnd.send_keys('2021/09/30')

searchBtn.click()

########################################################################
########################################################################
########################################################################

myList = [] # 空列表(存放最後要寫進excel的資料)

totalPageBtn = chrome.find_element_by_xpath("//*[contains(@src,'/Snoteanc/images/lp.gif')]") # 至最末頁的按鈕
maxPageNum = re.findall("[0-9]+", totalPageBtn.get_attribute("onclick"))[0]
print(" >>> maxPageNum : {} ".format(maxPageNum))

currentPageNum = 1
while (currentPageNum <= int(maxPageNum)):

    soup = BeautifulSoup(chrome.page_source, 'html.parser')
    dataTable = soup.find('table', { 'width': '100%', 'bordercolor': '#D9B55C' }) # 目標TABLE
    # print(" >>> dataTable : {}".format(dataTable))

    trRows = dataTable.find_all('tr')
    print(" >>> Page: {} - Max trRows number (不含表頭): 共 {} 列 ".format(currentPageNum, len(trRows) - 2))

    for i in range(len(trRows)):
        if i >= 2: # 跳過表頭
            perRow = trRows[i]
            tds = perRow.find_all('td')

            perDict = {
                'col-0' : tds[0].getText().strip(), # 1.商品代號
                'col-1' : tds[1].getText().strip(), # 2.商品名稱
                'col-2' : tds[3].getText().strip(), # 4.終止生效日
                'col-3' : tds[2].getText().strip(), # 3.計價幣別
                'col-4' : tds[4].getText().strip(), # 5.連結標的類別
                'col-5' : tds[5].getText().strip(), # 6.計價貨幣本金保本率
                'col-6' : tds[6].getText().strip(), # 7.投資人類別
                'col-7' : tds[7].getText().strip(), # 8.商品國內發行日
                'col-8' : tds[8].getText().strip(), # 9.發行機構
                'col-9' : tds[9].getText().strip(), # 10.總代理
                'col-10' : tds[10].getText().strip() # 11. 受託或銷售機構
            }

            myList.append(perDict)

    nextPageBtn = chrome.find_element_by_xpath("//*[contains(@src,'/Snoteanc/images/np.gif')]") # 下一頁按鈕
    nextPageBtn.click() # to下一頁
    currentPageNum += 1

########################################################################
########################################################################
########################################################################
"""
寫進Excel
"""
workbook = xlsxwriter.Workbook('./Example.xlsx')
worksheet = workbook.add_worksheet()
rr = 0 # row

for i in range(len(myList)):
    # print(myList[i])
    cc = 0 # col
    worksheet.write(rr, cc, myList[i].get("col-0"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-1"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-2"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-3"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-4"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-5"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-6"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-7"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-8"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-9"));  cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-10")); cc += 1;
    rr += 1

workbook.close()
chrome.quit() # close broswer