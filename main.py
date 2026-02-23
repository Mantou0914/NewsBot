from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import os
from linebot.v3.messaging import (
    BroadcastRequest,
    Configuration,
    ApiClient,
    MessagingApi,
    TextMessage,
)
# --- 設定區 ---
CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN', 'DsrXA3UcWQkYEO9aG831xPQIxhyHWFEJYuPLxDGIIoAf1Vs28fxRSPnGFygnH1Us9mBRo/wlUR81yJxJmQAVJgxUMLyb3dmekgVanCSEwiwWx/0DAlCNgl36rbxM/5gkRnyXQQ7P0KDLzKQ/PLtGawdB04t89/1O/w1cDnyilFU=')
ID_FILE = 'last_id.txt'  # 用來儲存最後一筆公告 ID 的檔案
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)

# 1. 讀取上次存的 ID
last_id = ""
if os.path.exists(ID_FILE):
    with open(ID_FILE, 'r') as f:
        last_id = f.read().strip()
temp = int(last_id)

def get_announcements(temp):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    announcements = []
    try:
        driver.get("https://www.ahs.nccu.edu.tw/home")
        time.sleep(5) # 等待載入

        # 檢查是否有討厭的彈窗
        try:
            alert = driver.switch_to.alert
            print(f"偵測到網頁彈窗: {alert.text}，正在嘗試關閉...")
            alert.accept() # 按下「確定」關閉彈窗
        except NoAlertPresentException:
            pass # 沒有彈窗，正常執行

        soup = BeautifulSoup(driver.page_source, 'lxml')
        rows = soup.find_all('tr', class_='tcontent')
        
        # 從HTML原始碼抓取公告標題
        for row in rows:
            cols = row.find_all('td')
            if len(cols) >= 3:
                link_tag = cols[2].find('a', id='content_href')
                nid = link_tag.get('nid')
                # 選取最新公告的ID(取nid最大者)
                if int(nid) > temp:
                    temp = int(nid)
                title = link_tag.get('title')
                announcements.append({'id': nid, 'title': title})
    except UnexpectedAlertPresentException as e:
        print(f"發生意外彈窗，程式跳過本次執行: {e}")
        return [], last_id # 發生錯誤時回傳空清單，避免程式崩潰
    except Exception as e:
        print(f"發生其他錯誤: {e}")
        return [], last_id
    finally:
        driver.quit()
    return announcements, temp

# --- 主程式邏輯 ---
# 2. 抓取公告
all_news, temp = get_announcements(temp)

if all_news:
    # 3. 找出比上次更新的公告 (假設第一筆是最新的)
    new_posts_content = []
    for post in all_news:
        if int(post['id']) > int(last_id):
            post_link = f"https://www.ahs.nccu.edu.tw/p/406-1000-{post['id']}.php"
            formatted_post = f"📌 {post['title']}\n🔗 連結：{post_link}"
            new_posts_content.append(formatted_post)
    
    # 4. 發送通知
    if new_posts_content:
        print(f"發現 {len(new_posts_content)} 則新公告！")

        combined_message = "📢 偵測到新公告！\n\n" + "\n\n".join(new_posts_content)
        
        # v3 新版發送邏輯：使用 ApiClient
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            
            try:
                # 建立發送請求
                broadcast_request = BroadcastRequest(
                    messages=[TextMessage(text=combined_message)]
                )
                
                # 執行推播
                line_bot_api.broadcast(broadcast_request)
        
                # 5. 更新最後記錄
                with open(ID_FILE, 'w') as f:
                    f.write(str(temp))
        
                print(f"成功發送 {len(new_posts_content)} 則新公告，ID 已更新為 {temp}")

            except Exception as e:
                print(f"發送失敗: {e}")
    else:
        print("目前沒有新公告。")