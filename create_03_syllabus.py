import requests
import urllib3

# 抑制 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def create_syllabus(cookie_string: str, url: str, module_id: int, summary: str, course_id: int | None = None, sort: int = 1) -> dict:
    """
    建立單元（syllabus）並回傳 JSON 結果
    參數:
        cookie_string: 瀏覽器登入 Cookie 字串
        url: 單元建立 API 的完整 URL（如 https://xxx/api/syllabus）
        module_id: 章節 ID（單元所屬章節）
        summary: 單元簡述或標題
        sort: 單元排序（預設 1）
    回傳:
        dict，包含單元 ID、名稱、所屬章節 ID，或錯誤資訊
    """

    # Cookie 轉換
    cookies = dict(item.split("=", 1) for item in cookie_string.split("; "))

    # 標頭
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": url.split("/api")[0],
        "Referer": url.replace("/api", ""),
        "User-Agent": "Mozilla/5.0"
    }

    # 最小必要資料
    data = {
        "module_id": module_id,
        "teaching_type": "face_to_face",
        "summary": summary,
        "sort": sort
    }
    
    # 如果有提供課程ID，則加入
    if course_id is not None:
        data["course_id"] = course_id

    try:
        response = requests.post(url, headers=headers, cookies=cookies, json=data, verify=False)
    except Exception as e:
        return {"success": False, "error": str(e)}

    if response.status_code == 201:
        try:
            res_json = response.json()
            return {
                "success": True,
                "syllabus_id": res_json.get("id"),
                "summary": res_json.get("summary"),
                "module_id": res_json.get("module_id")
            }
        except Exception:
            return {"success": False, "error": "回傳格式錯誤", "raw": response.text}
    else:
        return {
            "success": False,
            "status_code": response.status_code,
            "error": response.text
        }
