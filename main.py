from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import UnexpectedAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import re
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

# 讀取使用者資料
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

def categorize_news(title):
    if re.search(r"教師|老師|學務創新人員", title):
        return "teacher"
    elif re.search(r"國一", title):
        return "G7"
    elif re.search(r"國二", title):
        return "G8"
    elif re.search(r"國三", title):
        return "G9"
    elif re.search(r"高一", title):
        return "G10"
    elif re.search(r"高二", title):
        return "G11"
    elif re.search(r"高三", title):
        return "G12"
    elif re.search(r"國中", title) and not re.search(r"高國中", title):
        return "junior high"
    elif re.search(r"高中", title) and not re.search(r"國高中", title):
        return "senior high"
    elif re.search(r"國高中|高國中", title) and not re.search(r"美國|全國", title):
        return "the whole school"
    elif re.search(r"獎學金|獎助學金", title):
        return "scholarships or grants"
    elif re.search(r"大學", title)and not re.search(r"附屬", title):
        return "college"
    else:
        return "activities"


users = get_users_from_sheets()
general_ids = [u['UserID'] for u in users if 'gen' in str(u['Category'])]
grade7_ids = [u['UserID'] for u in users if '7' in str(u['Category'])]
grade8_ids = [u['UserID'] for u in users if '8' in str(u['Category'])]
grade9_ids = [u['UserID'] for u in users if '9' in str(u['Category'])]
grade10_ids = [u['UserID'] for u in users if '10' in str(u['Category'])]
grade11_ids = [u['UserID'] for u in users if '11' in str(u['Category'])]
grade12_ids = [u['UserID'] for u in users if '12' in str(u['Category'])]
teachers_ids = [u['UserID'] for u in users if 'tea' in str(u['Category'])]
college_informations_ids = [u['UserID'] for u in users if 'col' in str(u['Category'])]
activities_informations_ids = [u['UserID'] for u in users if 'act' in str(u['Category'])]
scholarships_grants_ids = [u['UserID'] for u in users if 'mon' in str(u['Category'])]

ids_list = [grade7_ids, grade8_ids, grade9_ids, grade10_ids, grade11_ids, grade12_ids, teachers_ids, college_informations_ids, college_informations_ids, activities_informations_ids, scholarships_grants_ids]

# --- 主程式邏輯 ---
# 抓取公告
all_news, temp = get_announcements(temp)

if all_news:
    # 找出新的公告
    new_posts_content = []
    new_posts_ids = []
    new_posts_message = []
    for post in all_news:
        if int(post['id']) > int(last_id):
            if post['id'] not in new_posts_ids:
                post['link'] = f"https://www.ahs.nccu.edu.tw/ischool/public/news_view/show.php?nid={post['id']}"
                new_posts_content.append(post)
                formatted_post = f"📌 {post['title']}\n🔗 連結：{post['link']}"
                new_posts_message.append(formatted_post)
                new_posts_ids.append(post['id'])

    # 發送通知
    if new_posts_content:
        print(f"發現 {len(new_posts_message)} 則新公告！")

        G7_message = []
        G8_message = []
        G9_message = []
        G10_message = []
        G11_message = []
        G12_message = []
        teacher_message = []
        college_message = []
        activities_message = []
        scholarships_grants_message = []

        for post in new_posts_content:
            category = categorize_news(post['title'])
            formatted_post = f"📌 {post['title']}\n🔗 連結：{post['link']}"
            print(f"已成功分類{category}")
            if category == "teacher":
                teacher_message.append(formatted_post)
            elif category == "G7":
                G7_message.append(formatted_post)
            elif category == "G8":
                G8_message.append(formatted_post)
            elif category == "G9":
                G9_message.append(formatted_post)
            elif category == "G10":
                G10_message.append(formatted_post)
            elif category == "G11":
                G11_message.append(formatted_post)
            elif category == "G12":
                G12_message.append(formatted_post)
            elif category == "junior high":
                G7_message.append(formatted_post)
                G8_message.append(formatted_post)
                G9_message.append(formatted_post)
            elif category == "senior high":
                G10_message.append(formatted_post)
                G11_message.append(formatted_post)
                G12_message.append(formatted_post)
            elif category == "the whole school":
                G7_message.append(formatted_post)
                G8_message.append(formatted_post)
                G9_message.append(formatted_post)
                G10_message.append(formatted_post)
                G11_message.append(formatted_post)
                G12_message.append(formatted_post)
            elif category == "scholarships or grants":
                scholarships_grants_message.append(formatted_post)
            elif category == "college":
                college_message.append(formatted_post)
            elif category == "activities":
                activities_message.append(formatted_post)

        general_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(new_posts_message)
        G7_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(G7_message)
        G8_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(G8_message)
        G9_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(G9_message)
        G10_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(G10_message)
        G11_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(G11_message)
        G12_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(G12_message)
        teacher_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(teacher_message)
        college_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(college_message)
        activities_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(activities_message)
        scholarships_grants_summary = "📢 偵測到新公告！\n\n" + "\n\n".join(scholarships_grants_message)

        summary_list = [G7_summary, G8_summary, G9_summary, G10_summary, G11_summary, G12_summary, teacher_summary, college_summary, activities_summary, scholarships_grants_summary, general_summary]


        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            
        try:
            for i in range(11):
                if ids_list[i] and summary_list[i]:
                    line_bot_api.multicast(
                        MulticastRequest(
                            to=ids_list[i],
                            messages=[TextMessage(text=summary_list[i])]
                        )
                    )

            # 更新最後公告ID
            with open(ID_FILE, 'w') as f:
                f.write(str(temp))
    
            print(f"ID 已更新為 {temp}")

        except Exception as e:
            print(f"發送失敗: {e}")
    else:
        print("目前沒有新公告。")