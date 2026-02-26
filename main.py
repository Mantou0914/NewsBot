from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import UnexpectedAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import os
import json
import gspread
from google.oauth2.service_account import Credentials
from linebot.v3.messaging import (
    BroadcastRequest,
    Configuration,
    ApiClient,
    MessagingApi,
    TextMessage,
    MulticastRequest
)
# --- 設定區 ---
CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
ID_FILE = 'last_id.txt'  # 用來儲存最後一筆公告 ID 的檔案
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)

# 讀取上次存的 ID
last_id = ""
if os.path.exists(ID_FILE):
    with open(ID_FILE, 'r') as f:
        last_id = f.read().strip()
temp = int(last_id)

def get_users_from_sheets():
    try:
        # 從 GitHub Secrets 抓取 JSON
        service_account_info = json.loads(os.getenv("GOOGLE_SHEETS_JSON"))
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
        client = gspread.authorize(creds)
        
        # 開啟試算表 (請確保名稱正確)
        sheet = client.open("NewsBot_Users").sheet1
        return sheet.get_all_records()
    except Exception as e:
        print(f"❌ Google Sheets 讀取失敗: {e}")
        return []

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

users = get_users_from_sheets()
grade7_ids = [u['UserID'] for u in users if str(u['Grade']) == '7']
grade8_ids = [u['UserID'] for u in users if str(u['Grade']) == '8']
grade9_ids = [u['UserID'] for u in users if str(u['Grade']) == '9']
grade10_ids = [u['UserID'] for u in users if str(u['Grade']) == '10']
grade11_ids = [u['UserID'] for u in users if str(u['Grade']) == '11']
grade12_ids = [u['UserID'] for u in users if str(u['Grade']) == '12']
all_user_ids = [u['UserID'] for u in users] # 所有人

# --- 主程式邏輯 ---
# 抓取公告
all_news, temp = get_announcements(temp)

if all_news:
    # 找出比上次更新的公告
    new_posts_content = []
    new_post_id = []
    for post in all_news:
        if int(post['id']) > int(last_id):
            if post['id'] not in new_post_id:
                post_link = f"https://www.ahs.nccu.edu.tw/ischool/public/news_view/show.php?nid={post['id']}"
                formatted_post = f"📌 {post['title']}\n🔗 連結：{post_link}"
                new_posts_content.append(formatted_post)
                new_post_id.append(post['id'])
    
    # 發送通知
    if new_posts_content:
        print(f"發現 {len(new_posts_content)} 則新公告！")

        combined_message = "📢 偵測到新公告！\n\n" + "\n\n".join(new_posts_content)
        
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            
        try:
            # 建立廣播請求
            broadcast_request = BroadcastRequest(
                messages=[TextMessage(text=combined_message)]
            )
            
            # 廣播
            line_bot_api.broadcast(broadcast_request)
    
            # 更新最後公告ID
            with open(ID_FILE, 'w') as f:
                f.write(str(temp))
    
            print(f"成功發送 {len(new_posts_content)} 則新公告，ID 已更新為 {temp}")

        except Exception as e:
            print(f"發送失敗: {e}")
    else:
        print("目前沒有新公告。")