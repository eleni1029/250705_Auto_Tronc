# config_template.py
# 設定範本檔案 - 請複製此檔案為 config.py 並填入您的實際資訊

# 基本登入資訊
USERNAME = 'your_email@example.com'  # 您的帳號
PASSWORD = 'your_password'  # 您的密碼
LOGIN_URL = 'https://wg.tronclass.com/login'  # 登入網址

# 預設課程和章節 ID
COURSE_ID = 16387  # 預設課程 ID
MODULE_ID = 28652  # 預設章節 ID

# Cookie 和請求設定
COOKIE = ''  # 自動登入獲取，請留空
SLEEP_SECONDS = 1  # 每次請求間隔，避免被擋

# API 基礎 URLs
BASE_URL = 'https://wg.tronclass.com'

# 動態 API URLs（會根據 COURSE_ID 自動生成）
def get_api_urls():
    """獲取所有 API URLs"""
    return {
        'COURSE_CREATE_URL': f'{BASE_URL}/api/course',
        'MODULE_CREATE_URL': f'{BASE_URL}/api/course/{COURSE_ID}/module',
        'SYLLABUS_CREATE_URL': f'{BASE_URL}/api/syllabus',
        'ACTIVITY_CREATE_URL': f'{BASE_URL}/api/course/{COURSE_ID}/activity',
        'UPLOAD_URL': f'{BASE_URL}/api/upload',
        'MATERIAL_CREATE_URL': f'{BASE_URL}/api/course/{COURSE_ID}/material'
    }

# 學習活動類型映射
ACTIVITY_TYPE_MAPPING = {
    '線上連結': 'web_link',
    '影音連結': 'web_link',
    '參考檔案_圖片': 'material',
    '參考檔案_PDF': 'material'
}

# 支援的學習活動類型（用於用戶選擇）
SUPPORTED_ACTIVITY_TYPES = ['線上連結', '影音連結', '參考檔案_圖片', '參考檔案_PDF']