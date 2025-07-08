import os
import requests
import json
from datetime import datetime
import urllib3

# 抑制 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from datetime import date

def create_course(cookie_string: str, url: str, course_name: str, start_date: str | None = None) -> dict:
    """
    建立課程並回傳課程資訊
    參數:
        cookie_string: 來自瀏覽器的 Cookie 字串
        url: 創建課程 API 的完整 URL（例如 https://example.com/api/course）
        course_name: 課程名稱
        start_date: 開課日期，格式為 yyyy-mm-dd
    回傳:
        dict，包含課程名稱與課程 ID（若失敗則回傳錯誤訊息）
    """

    # 將 Cookie 字串轉為字典
    cookies = dict(item.split("=", 1) for item in cookie_string.split("; "))

    # 標頭（可根據實際需求擴充）
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": url.split("/api")[0],
        "Referer": url.replace("/api/course", "/course/add"),
        "User-Agent": "Mozilla/5.0",
        "X-Requested-With": "XMLHttpRequest"
    }

    # 如果沒有提供開課日期，預設為今天
    if start_date is None:
        start_date = date.today().strftime('%Y-%m-%d')
    
    # 建課資料（固定結構）
    course_data = {
        "semester_id": 0,
        "semester": {"id": 0, "sort": -1},
        "compulsory": False,
        "course_type": 1,
        "academic_year_id": 0,
        "academic_year": {"id": 0, "sort": -1},
        "course_template": -1,
        "disable_nav_enrollments": True,
        "name": course_name,
        "start_date": start_date,
        "instructor_ids": [],
        "description": "",
        "published": True,
        "enrollment_type": "open"
    }

    try:
        response = requests.post(url, headers=headers, cookies=cookies, json=course_data, verify=False)
    except Exception as e:
        return {"success": False, "error": str(e)}

    if response.status_code == 201:
        # 嘗試從 headers 抓課程 ID
        course_id = None
        location = response.headers.get("Location", "")
        if location.startswith("/course/"):
            course_id = location.split("/")[-1]
        
        # 嘗試從回應內容中獲取課程 ID
        try:
            response_data = response.json()
            course_id = response_data.get("id") or response_data.get("course_id")
        except:
            pass

        return {
            "success": True,
            "course_id": course_id,
            "course_name": course_name
        }
    else:
        return {
            "success": False,
            "status_code": response.status_code,
            "error": response.text
        }
