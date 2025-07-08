import requests
import urllib3

# 抑制 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def create_module(cookie_string: str, url: str, module_name: str, course_id: int | None = None, sort: int = 1) -> dict:
    """
    建立章節（module）並回傳 JSON 結果
    參數:
        cookie_string: 登入後的 Cookie 字串
        url: 章節建立 API 的完整 URL（如 https://xxx/api/course/16390/module）
        module_name: 章節名稱
        sort: 排序順序（整數）
    回傳:
        dict，包含是否成功、章節 ID、章節名稱、課程 ID，或錯誤資訊
    """

    # Cookie 轉換
    cookies = dict(item.split("=", 1) for item in cookie_string.split("; "))

    # 標頭
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": url.split("/api")[0],
        "Referer": url.replace("/api/course", "/course"),
        "User-Agent": "Mozilla/5.0"
    }

    # 傳送資料（使用最小欄位）
    data = {
        "name": module_name,
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
                "course_id": res_json.get("course_id"),
                "module_id": res_json.get("id"),
                "module_name": res_json.get("name")
            }
        except Exception as e:
            return {
                "success": False,
                "error": "回傳內容格式錯誤",
                "raw": response.text
            }
    else:
        return {
            "success": False,
            "status_code": response.status_code,
            "error": response.text
        }
