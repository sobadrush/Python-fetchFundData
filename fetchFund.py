from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
 
import codecs
import re
 
# import xlsxwriter module
import xlsxwriter

options = Options()
options.add_argument("--disable-notifications")
# options.add_argument('--headless') # Chrome無頭騎士模式
 
chrome = webdriver.Chrome('./broswerDrivers/chromedriver.exe', chrome_options=options)
chrome.get("https://structurednotes-announce.tdcc.com.tw/Snoteanc/apps/bas/BAS210.jsp")

# 商品國內發行日
iDateStart = chrome.find_element_by_id("iDateStart")
iDateEnd = chrome.find_element_by_id("iDateEnd")
searchBtn = chrome.find_element_by_xpath('//input[@value="查詢"]') # 查詢按鈕

iDateStart.send_keys('2021/09/01')
iDateEnd.send_keys('2021/09/30')

searchBtn.click()

# time.sleep(1.5)

myList = [] ## 空列表(存放最後要寫進excel的資料)

totalPageBtn = chrome.find_element_by_xpath("//*[contains(@src,'/Snoteanc/images/lp.gif')]") # 至最末頁的按鈕
maxPageNum = re.findall("[0-9]+", totalPageBtn.get_attribute("onclick"))[0]
print(" >>> maxPageNum : {} ".format(maxPageNum))

currentPageNum = 1
while (currentPageNum <= int(maxPageNum)):
    # file = codecs.open("test.txt", "w", "utf-8")

    soup = BeautifulSoup(chrome.page_source, 'html.parser')
    dataTable = soup.find('table', { 'width': '100%', 'bordercolor': '#D9B55C' }) # 目標TABLE
    # print(" >>> dataTable : {}".format(dataTable))

    trRows = dataTable.find_all('tr')
    print(" >>> Page: {} - Max trRows number (不含表頭): 共 {} 列 ".format(currentPageNum, len(trRows) - 2))

    for i in range(len(trRows)):
        if i >= 2:
            perRow = trRows[i]
            tds = perRow.find_all('td')
            # print(tds[0].getText().strip(), end = ' ') # 1.商品代號
            # print(tds[1].getText().strip(), end = ' ') # 2.商品名稱
            # print(tds[3].getText().strip(), end = ' ') # 4.終.strip()止生效日
            # print(tds[2].getText().strip(), end = ' ') # 3.計價幣別
            # print(tds[4].getText().strip(), end = ' ') # 5.連結標的類別
            # print(tds[5].getText().strip(), end = ' ') # 6.計價貨幣本金保本率
            # print(tds[6].getText().strip(), end = ' ') # 7.投資人類別	
            # print(tds[7].getText().strip(), end = ' ') # 8.商品國內發行日
            # print(tds[8].getText().strip(), end = ' ') # 9.發行機構
            # print(tds[9].getText().strip(), end = ' ') # 10.總代理
            # print(tds[10].getText().strip()) # 11.受託或銷售機構

            perDict = {
                'col-0' : tds[0].getText().strip(),
                'col-1' : tds[1].getText().strip(),
                'col-2' : tds[3].getText().strip(),
                'col-3' : tds[2].getText().strip(),
                'col-4' : tds[4].getText().strip(),
                'col-5' : tds[5].getText().strip(),
                'col-6' : tds[6].getText().strip(),
                'col-7' : tds[7].getText().strip(),
                'col-8' : tds[8].getText().strip(),
                'col-9' : tds[9].getText().strip(),
                'col-10' : tds[10].getText().strip()
            }

            myList.append(perDict)

    # time.sleep(1.5)
    nextPageBtn = chrome.find_element_by_xpath("//*[contains(@src,'/Snoteanc/images/np.gif')]") # 下一頁按鈕
    nextPageBtn.click() # to下一頁
    currentPageNum += 1

            # file.write(str(row))
        # file.close()

"""
Writing Excel
"""
workbook = xlsxwriter.Workbook('./Example.xlsx')
worksheet = workbook.add_worksheet()
rr = 0 # row

for i in range(len(myList)):
    # print(myList[i])
    cc = 0 # col
    worksheet.write(rr, cc, myList[i].get("col-0")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-1")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-2")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-3")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-4")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-5")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-6")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-7")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-8")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-9")); cc += 1;
    worksheet.write(rr, cc, myList[i].get("col-10")); cc += 1;
    rr += 1

workbook.close()
chrome.quit() # close broswer