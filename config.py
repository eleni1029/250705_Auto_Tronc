# config.py
# 全局設定檔案，請在此填寫共用參數
COOKIE = '_ga_SG0N0692X5=GS2.1.s1753785258$o1$g1$t1753785264$j54$l0$h0; warning:verification_email=show; session=V2-1-51c9d711-ee11-476c-b53a-fc51e2956a66.Mjc4NjI.1753792464360.CeyF-65_HBPirIsjZxuyq5zprx4; _ga=GA1.1.35448435.1753785258; lang=zh-TW'  # 自動登入獲取
SLEEP_SECONDS = 0.1  # 每次請求間隔，避免被擋
USERNAME = 'eleni1029@gmail.com'  # 你的帳號
PASSWORD = 'eleni1029'  # 你的密碼
BASE_URL = 'https://staging.tronclass.com'  # 基礎網址
LOGIN_URL = f'{BASE_URL}/login'  # 登入網址
COURSE_ID = 16401  # 預設的課程 ID
MODULE_ID = 28739  # 預設的章節 ID

# 活動類型映射
ACTIVITY_TYPE_MAPPING = {
    '線上連結': 'web_link',
    '影音連結': 'web_link',
    '影音教材_影音連結': 'online_video',
    '參考檔案_圖片': 'material',
    '參考檔案_PDF': 'material',
    '影音教材_影片': 'video',
    '影音教材_音訊': 'audio'
}

# 支援的活動類型
SUPPORTED_ACTIVITY_TYPES = [
    '線上連結',
    '影音連結',
    '影音教材_影音連結',
    '參考檔案_圖片',
    '參考檔案_PDF',
    '影音教材_影片',
    '影音教材_音訊'
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