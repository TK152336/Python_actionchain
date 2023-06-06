# 匯入套件
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from time import sleep
import os
import json
import csv
from pprint import pprint
import requests as req
import re

# 啟動瀏覽器工具的選項
my_options = webdriver.ChromeOptions()
# my_options.add_argument("--headless")                #不開啟實體瀏覽器背景執行
my_options.add_argument("--start-maximized")         #最大化視窗
my_options.add_argument("--incognito")               #開啟無痕模式
my_options.add_argument("--disable-popup-blocking") #禁用彈出攔截
my_options.add_argument("--disable-notifications")  #取消 chrome 推播通知
my_options.add_argument("--lang=zh-TW")  #設定為正體中文

# 使用 Chrome 的 WebDriver
# 自動取得 Chrome 的 WebDriver
driver = webdriver.Chrome(
    options = my_options,
    service = Service(ChromeDriverManager().install())
)

# 走訪網頁
driver.get("https://24h.pchome.com.tw/")

# 放置商品資訊的地方
listData = [["商品名稱","商品價格","24hr到貨"]]


def search():
    txtInput = driver.find_element (By.CSS_SELECTOR ,'input.c-siteSearchInput')
    txtInput.send_keys('外接硬碟')

    # 等待一下
    sleep(1)
    
    # 按鈕選擇器
    cssSelectorBtn = "div.c-siteSearchBtn>button.gtmClickV2.btn.btn--sm"
    # 取得按鈕元素
    btn = driver.find_element(By.CSS_SELECTOR, cssSelectorBtn)
    
    # 按下按鈕
    btn.click()
    

# 滾動頁面
def scroll():
    innerHeight = 0
    offset = 0
    count = 0
    limit = 3
    
    while count <= limit:
        # 每次移動高度
        offset = driver.execute_script(
            'return window.document.documentElement.scrollHeight;'
        )
        # 捲軸往下滑動 auto 是直接到底 smooth是慢滑
        driver.execute_script(f'''
            window.scrollTo({{top: {offset}, behavior: 'smooth' }});
            ''')
        
        # 強制等待，此時若有新元素生成，瀏覽器內部高度會自動增加
        sleep(3)
    
        # 透過執行 js 語法來取得捲動後的當前總高度
        innerHeight = driver.execute_script(
            'return window.document.documentElement.scrollHeight;'
        )
        
        # 經過計算，如果滾動距離(offset)大於等於視窗內部總高度(innerHeight)，代表已經到底了
        # if offset == innerHeight:
        #     count += 1
        
        # 蒐集一定資料即可，捲動超過一定的距離，就結束程式
        if innerHeight >= 6000:
            break

# 取得頁面我要的元素資訊
def collectinfo():
    items = driver.find_elements(By.CSS_SELECTOR, 'dl.col3f')
    
    for a in items:
        try:

            title = a.find_element(By.CSS_SELECTOR, 'h5.prod_name a').get_attribute('innerText') #今日發現 可以如此接著寫下去
            # titlename = title.get_attribute('innerText')
            # titlename = title.text
            # print(title)
            # listData.append({
            #         '商品名稱':title
            #     })            

            price = a.find_element(By.CSS_SELECTOR, 'span.price span.value').get_attribute('innerText')
            # print(price)
            # print(type(price))
            # listData.append({
            #         '商品價格':price
            #     })            

            hrs = a.find_element(By.CSS_SELECTOR, 'span.ico.label-24h-new').get_attribute('innerText')
            # print(hrs)
            # listData.append({
            #         '24hr到貨': hrs
            #     })            
            
            listData.append([title, price, hrs])
        
        except NoSuchElementException as e:
            # print('沒有24H到貨')
            listData.append([title, price, '無'])         
            continue
            
    # pprint(listData)

def savecsv():
    
    # 建立儲存商品資料的資料夾
    folderPath = 'HW_PChome_actionchain'
    if not os.path.exists(folderPath):
        os.makedirs(folderPath) #預設新增資料夾在平行的位置
    
    with open(f"{folderPath}/pchome0201.csv", "w", encoding='utf-8') as file:
        # file.write( csv.dumps(listData, ensure_ascii=False, indent=4))
        writer = csv.writer(file)
        writer.writerows(listData)

# 關閉瀏覽器
def close():
    driver.quit()


import pyautogui
from time import sleep, time
import keyboard
from IPython.display import clear_output

def action():
    # 設定每一個動作，都暫停若干秒
    pyautogui.PAUSE = 3.75

    # 1.點選工作列excel [1480,1055]
    pyautogui.click(1430,1055)

    # 2.開啟空白活頁簿 [345,260]
    pyautogui.click(345,260)

    # 3.資料[360,75]
    pyautogui.click(360,75)

    # 4.>> 從文字/CSV [130,115]
    pyautogui.click(130,115)

    # 5 找到檔案並打開
    # pyautogui.click(670,240)
    pyautogui.moveTo(354,390)
    # pyautogui.scroll(-3)
    pyautogui.dragTo(354, 440, 2)
    pyautogui.click(290,605)
    pyautogui.doubleClick(460,530)
    
    # 5.[670,240]
    # C:\\Users\\terra\\BDSE29_學習資料\\網路爬蟲\\python_web_scraping-master\\HW_PChome_actionchain
    # pyautogui.click(670,240)
    # pyautogui.typewrite('C:\\Users\\terra\\BDSE29_學習資料\\網路爬蟲\\python_web_scraping-master\\HW_PChome_actionchain',interval=0.75)
    # pyautogui.press('enter')

    # 6.找到檔案 點擊兩下 打開檔案 [375,200]
    # pyautogui.click(510,410)
    pyautogui.doubleClick(510,410)

    # 7.按下載入[1220,865]
    pyautogui.click(1220,865)
    
    #8.ctrl+q 變更美美字型
    pyautogui.hotkey('ctrl', 'q')
    
    # 9.商品價格[870,370]>> 大到小[750,430]
    pyautogui.click(870,370)
    pyautogui.click(750,430)
    
    # 10.24hr到貨[980,370]>> Z到A排序[990,435]
    pyautogui.click(980,370)
    pyautogui.click(990,435)
    
    # 11.ctrl+S 存檔
    pyautogui.hotkey('ctrl', 's')

    # 12.輸入檔名 
    pyautogui.click(1610,1050)
    # pyautogui.press('shiftleft')
    pyautogui.typewrite('pchomeHDinfo',interval=0.75)
    
    # 13.按下enter
    pyautogui.press('enter')

# 養成好習慣 上面定義函數 下面就呼叫函數並執行 一個function一個動作  
if __name__ == '__main__':
    search()
    scroll()
    collectinfo()
    savecsv()
    close()

if __name__ == '__main__':
    action()




























