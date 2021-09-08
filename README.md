
# Python-fetchFundData
**這是一個用Python爬基金資料&輸出Excel的專案**

## 使用套件
1. webdriver : **selenium** (參考 [安裝Selenium及Webdriver](https://www.learncodewithmike.com/2020/05/python-selenium-scraper.html))
2. parser : **BeautifulSoup**
3. regex : **re**
4. excel : **xlsxwriter**

## 使用方式

* 安裝python虛擬環境 virtualenv
    * $ ***pip3 install virtualenv***
    
* 建立python虛擬環境(envGG為自訂的環境名稱)
    * $ ***virtualenv.exe envGG***
    
* 進入envGG虛擬環境(cmd執行activate.bat)
    * $ ***.\envGG\Scripts\activate.bat***
    
* 使用 requirements.txt 進行dependency安裝 (envGG)表示已在虛擬環境)
    * ***(envGG)*** $ ***pip install -r .\requirements.txt***
    * 可觀察到 envGG/ 中的 lib/site-packages 目錄下安裝了所需的lib

* 執行程式
    * ***(envGG)*** $ ***py fetchFund_simple.py***

* 結果
    * 會在同目錄產生 **Example.xlsx** 內含所有基金資料


## 其他

1. 若有增加lib, 導出dependency :  **pip freeze > requirements.txt**
2. **options.add_argument('--headless')** # Chrome無頭騎士模式 → 執行時是否蹦出瀏覽器(debug可用)
3. **fetchFund.py** → 沒整理過的版本
    可用 file = codecs.open("test.txt", "w", "utf-8") 寫出txt測試 or debug
5. **fetchFund_simple.py** → 整理過的精簡版

## 注意事項 & 修改方向
<ul>
    <li style="color:red">目前「查詢條件」僅以「商品國內發行日」寫死進行測試，可自行增加</li>
    <li style="color:red">
        可嘗試使用PyQt5套件建立Windows視窗應用程式接收USER輸入資料
        (<a href="https://dotblogs.com.tw/zeus83157/2018/06/23/011440">Python建立圖形使用者介面的神兵利器-PyQt5)</a>
    </li>
</ul>
