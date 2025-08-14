# config.py
# 全局設定檔案，請在此填寫共用參數

import os
from dotenv import load_dotenv

# 載入環境變數
load_dotenv()

# 從環境變數獲取敏感資訊
USERNAME = os.getenv('USERNAME', '')  # 從 .env 文件讀取
PASSWORD = os.getenv('PASSWORD', '')  # 從 .env 文件讀取  
BASE_URL = os.getenv('BASE_URL', 'https://staging.tronclass.com')  # 從 .env 文件讀取
SEARCH_TERM = os.getenv('SEARCH_TERM', '')  # 機構搜索關鍵字

# 其他系統設定
COOKIE = 'session=V2-168-9c3515c8-9366-4917-8829-9cc9157cf66e.ODE0ODc.1755229507109.NhTYWQPToTGbSvh-gR6ZmwkFPcI; _ga_ZCC2R3ZYVG=GS2.1.s1755143094$o1$g1$t1755143103$j51$l0$h0; samesite=strict; _ga=GA1.1.534984174.1755143094; samesite=strict; warning:verification_email=show; lang=zh-TW'  # 自動登入獲取
SLEEP_SECONDS = 0.1  # 每次請求間隔，避免被擋
LOGIN_URL = f'{BASE_URL}/login'  # 登入網址
COURSE_ID = 10000  # 預設的課程 ID
MODULE_ID = 10000  # 預設的章節 ID

# 活動類型映射
ACTIVITY_TYPE_MAPPING = {
    '線上連結': 'web_link',
    '影音連結': 'web_link',
    '影音教材_影音連結': 'online_video',
    '參考檔案_圖片': 'material',
    '參考檔案_PDF': 'material',
    '影音教材_影片': 'video',
    '影音教材_音訊': 'audio'  # 尚未驗證支持
}

# 支援的活動類型
SUPPORTED_ACTIVITY_TYPES = [
    '線上連結',
    '影音連結',
    '影音教材_影音連結',
    '參考檔案_圖片',
    '參考檔案_PDF',
    '影音教材_影片',
    '影音教材_音訊'  # 尚未驗證支持
]

def get_api_urls():
    """獲取 API URLs"""
    return {
        'COURSE_CREATE_URL': f'{BASE_URL}/api/course',
        'MODULE_CREATE_URL': f'{BASE_URL}/api/course/{COURSE_ID}/module',
        'SYLLABUS_CREATE_URL': f'{BASE_URL}/api/syllabus',
        'ACTIVITY_CREATE_URL': f'{BASE_URL}/api/courses/{COURSE_ID}/activities',
        'MATERIAL_UPLOAD_URL': f'{BASE_URL}/api/uploads',
        'MATERIAL_CREATE_URL': f'{BASE_URL}/api/material'
    }