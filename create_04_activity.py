import requests
import os, json
from datetime import datetime
import urllib3
from config import BASE_URL, get_api_urls

# 抑制 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def _build_headers(cookie_string: str, url: str):
    cookies = dict(item.split("=", 1) for item in cookie_string.split("; "))
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": url.split("/api")[0],
        "Referer": url.replace("/api/course", "/course"),
        "User-Agent": "Mozilla/5.0",
        "X-Requested-With": "XMLHttpRequest"
    }
    return headers, cookies


def create_link_activity(cookie_string: str, url: str, title: str, link_url: str,
                         module_id: int | None = None, syllabus_id: int | None = None, sort: int = 1) -> dict:
    """
    建立「線上連結」學習活動
    """
    headers, cookies = _build_headers(cookie_string, url)

    payload = {
        "type": "web_link",
        "announce_score_type": 2,
        "announce_score_time": None,
        "score_percentage": 0,
        "week": 0,
        "teaching_method": 0,
        "knowledge_node_ids": [],
        "publishStatusClass": "publish-type-button unpublished",
        "publishTypeText": "未發布",
        "isPublishScheduled": False,
        "publishScheduledDuration": "發布時間: Invalid date ～ 無截止",
        "title": title,
        "link": link_url,
        "published": True,
        "completion_criterion": {
            "activity_completion_criterion_type_id": 2,
            "value": 0
        },
        "using_phase": "unspecified",
        "teaching_model": "online",
        "prerequisites": [],
        "sort": sort
    }
    # 傳值邏輯 - 同時傳送 module_id 和 syllabus_id（如果有的話）
    if module_id is not None:
        payload["module_id"] = module_id
    if syllabus_id is not None:
        payload["syllabus_id"] = syllabus_id

    try:
        response = requests.post(url, headers=headers, cookies=cookies, json=payload, verify=False)
    except Exception as e:
        return {"success": False, "error": str(e), "request_payload": payload}

    if response.status_code == 201:
        try:
            res = response.json()
            return {
                "success": True,
                "activity_id": res.get("id"),
                "activity_title": res.get("title"),
                "module_id": res.get("module_id"),
                "syllabus_id": res.get("syllabus_id", None)
            }
        except Exception:
            return {"success": False, "error": "無法解析 response", "raw": response.text, "request_payload": payload}
    else:
        return {"success": False, "status_code": response.status_code, "error": response.text, "request_payload": payload}


def create_reference_activity(cookie_string: str, url: str, title: str, module_id: int | None = None,
                              syllabus_id: int | None = None, description: str = " ",
                              upload_id: int | None = None, upload_name: str = "", sort: int = 1) -> dict:
    """
    建立「參考資料」學習活動
    參數:
        upload_id: 上傳檔案的 ID
        upload_name: 上傳檔案的名稱
    """
    headers, cookies = _build_headers(cookie_string, url)

    # 構建 upload_references
    upload_references = []
    uploads = []
    if upload_id and upload_name:
        upload_references.append({
            "upload_id": upload_id,
            "allow_download": True,
            "cc_license_name": "private",
            "reference_id": 0,
            "name": upload_name,
            "origin_allow_download": False  # 預設為 false
        })
        uploads = [upload_id]

    payload = {
        "type": "material",
        "description": description,
        "uploads": uploads,
        "other_resources": [],
        "publishStatusClass": "publish-type-button unpublished",
        "publishTypeText": "未發布",
        "isPublishScheduled": False,
        "publishScheduledDuration": "發布時間: Invalid date ～ 無截止",
        "title": title,
        "upload_references": upload_references,
        "material_type": "Material",
        "published": True,
        "completion_criterion": {
            "activity_completion_criterion_type_id": 4,  # 修改為 4
            "value": 0
        },
        "using_phase": "unspecified",
        "teaching_model": "online",
        "prerequisites": [],
        "submit_times": 1,
        "announce_answer_and_explanation": False,
        "sort": sort
    }
    # 傳值邏輯 - 同時傳送 module_id 和 syllabus_id（如果有的話）
    if module_id is not None:
        payload["module_id"] = module_id
    if syllabus_id is not None:
        payload["syllabus_id"] = syllabus_id

    try:
        response = requests.post(url, headers=headers, cookies=cookies, json=payload, verify=False)
    except Exception as e:
        return {"success": False, "error": str(e), "request_payload": payload}

    if response.status_code == 201:
        try:
            res = response.json()
            return {
                "success": True,
                "activity_id": res.get("id"),
                "activity_title": res.get("title"),
                "module_id": res.get("module_id"),
                "syllabus_id": res.get("syllabus_id", None)
            }
        except Exception:
            return {"success": False, "error": "無法解析 response", "raw": response.text, "request_payload": payload}
    else:
        return {"success": False, "status_code": response.status_code, "error": response.text, "request_payload": payload}


def create_video_activity(cookie_string: str, url: str, title: str, upload_id: int,
                         upload_name: str, module_id: int | None = None, 
                         syllabus_id: int | None = None, sort: int = 1,
                         completion_criterion_value: int = 80, submit_times: int = 1) -> dict:
    """
    建立「影音教材_影片」學習活動（使用上傳檔案）
    
    Args:
        cookie_string: Cookie字串
        url: API URL
        title: 活動標題
        upload_id: 上傳檔案的ID (注意：不再使用video_url參數)
        upload_name: 上傳檔案的名稱（包含副檔名，如：7-3-2.mp4）
        module_id: 章節ID
        syllabus_id: 單元ID
        sort: 排序
        completion_criterion_value: 完成條件值（預設80%）
        submit_times: 提交次數（預設1）
    """
    headers, cookies = _build_headers(cookie_string, url)

    # 處理預設值，避免NaN問題
    if completion_criterion_value is None or str(completion_criterion_value).lower() == 'nan':
        completion_criterion_value = 80
    if submit_times is None or str(submit_times).lower() == 'nan':
        submit_times = 1

    payload = {
        "type": "online_video",  # 修正：使用online_video而不是video
        "week": 0,
        "teaching_method": 0,
        "knowledge_node_ids": [],
        "publishStatusClass": "publish-type-button unpublished",
        "publishTypeText": "未發布",
        "isPublishScheduled": False,
        "publishScheduledDuration": "發布時間: Invalid date ～ 無截止",
        "title": title,
        "uploads": [upload_id],  # 修正：使用uploads而不是video_url
        "upload_references": [{  # 修正：新增upload_references
            "upload_id": upload_id,
            "cc_license_name": "private",
            "reference_id": 0,
            "origin_allow_download": True,
            "name": upload_name
        }],
        "allow_download": False,  # 新增：是否允許下載
        "allow_forward_seeking": True,  # 新增：是否允許快進
        "pause_when_leaving_window": True,  # 新增：離開視窗時暫停
        "published": True,
        "completion_criterion": {
            "activity_completion_criterion_type_id": 12,  # 修正：使用12而不是2
            "value": completion_criterion_value
        },
        "using_phase": "unspecified",
        "teaching_model": "online",
        "prerequisites": [],
        "submit_times": submit_times,
        "announce_answer_and_explanation": False,
        "sort": sort
    }
    
    # 傳值邏輯 - 同時傳送 module_id 和 syllabus_id（如果有的話）
    if module_id is not None:
        payload["module_id"] = module_id
    if syllabus_id is not None:
        payload["syllabus_id"] = syllabus_id

    try:
        response = requests.post(url, headers=headers, cookies=cookies, json=payload, verify=False)
    except Exception as e:
        return {"success": False, "error": str(e), "request_payload": payload}

    if response.status_code == 201:
        try:
            res = response.json()
            return {
                "success": True,
                "activity_id": res.get("id"),
                "activity_title": res.get("title"),
                "module_id": res.get("module_id"),
                "syllabus_id": res.get("syllabus_id", None)
            }
        except Exception:
            return {"success": False, "error": "無法解析 response", "raw": response.text, "request_payload": payload}
    else:
        return {"success": False, "status_code": response.status_code, "error": response.text, "request_payload": payload}


def create_audio_activity(cookie_string: str, url: str, title: str, upload_id: int,
                         upload_name: str, module_id: int | None = None, 
                         syllabus_id: int | None = None, sort: int = 1,
                         completion_criterion_value: int = 80, submit_times: int = 1) -> dict:
    """
    建立「影音教材_音訊」學習活動（使用上傳檔案）
    
    Args:
        cookie_string: Cookie字串
        url: API URL  
        title: 活動標題
        upload_id: 上傳檔案的ID (注意：不再使用audio_url參數)
        upload_name: 上傳檔案的名稱（包含副檔名，如：audio.mp3）
        module_id: 章節ID
        syllabus_id: 單元ID
        sort: 排序
        completion_criterion_value: 完成條件值（預設80%）
        submit_times: 提交次數（預設1）
    """
    headers, cookies = _build_headers(cookie_string, url)

    # 處理預設值，避免NaN問題
    if completion_criterion_value is None or str(completion_criterion_value).lower() == 'nan':
        completion_criterion_value = 80
    if submit_times is None or str(submit_times).lower() == 'nan':
        submit_times = 1

    payload = {
        "type": "online_audio",  # 音訊類型
        "week": 0,
        "teaching_method": 0,
        "knowledge_node_ids": [],
        "publishStatusClass": "publish-type-button unpublished", 
        "publishTypeText": "未發布",
        "isPublishScheduled": False,
        "publishScheduledDuration": "發布時間: Invalid date ～ 無截止",
        "title": title,
        "uploads": [upload_id],
        "upload_references": [{
            "upload_id": upload_id,
            "cc_license_name": "private",
            "reference_id": 0,
            "origin_allow_download": True,
            "name": upload_name
        }],
        "allow_download": False,
        "published": True,
        "completion_criterion": {
            "activity_completion_criterion_type_id": 12,
            "value": completion_criterion_value
        },
        "using_phase": "unspecified",
        "teaching_model": "online",
        "prerequisites": [],
        "submit_times": submit_times,
        "announce_answer_and_explanation": False,
        "sort": sort
    }
    
    # 傳值邏輯
    if module_id is not None:
        payload["module_id"] = module_id
    if syllabus_id is not None:
        payload["syllabus_id"] = syllabus_id

    try:
        response = requests.post(url, headers=headers, cookies=cookies, json=payload, verify=False)
    except Exception as e:
        return {"success": False, "error": str(e), "request_payload": payload}

    if response.status_code == 201:
        try:
            res = response.json()
            return {
                "success": True,
                "activity_id": res.get("id"),
                "activity_title": res.get("title"),
                "module_id": res.get("module_id"),
                "syllabus_id": res.get("syllabus_id", None)
            }
        except Exception:
            return {"success": False, "error": "無法解析 response", "raw": response.text, "request_payload": payload}
    else:
        return {"success": False, "status_code": response.status_code, "error": response.text, "request_payload": payload}


def create_video_link_activity(cookie_string: str, url: str, title: str, video_link_url: str,
                              module_id: int | None = None, syllabus_id: int | None = None, sort: int = 1) -> dict:
    """
    建立「影音教材_影音連結」學習活動
    """
    headers, cookies = _build_headers(cookie_string, url)

    payload = {
        "type": "video_link",
        "title": title,
        "video_link_url": video_link_url,
        "published": True,
        "sort": sort
    }
    if module_id is not None:
        payload["module_id"] = module_id
    if syllabus_id is not None:
        payload["syllabus_id"] = syllabus_id

    try:
        response = requests.post(url, headers=headers, cookies=cookies, json=payload, verify=False)
    except Exception as e:
        return {"success": False, "error": str(e), "request_payload": payload}

    if response.status_code == 201:
        try:
            res = response.json()
            return {
                "success": True,
                "activity_id": res.get("id"),
                "activity_title": res.get("title"),
                "module_id": res.get("module_id"),
                "syllabus_id": res.get("syllabus_id", None)
            }
        except Exception:
            return {"success": False, "error": "無法解析 response", "raw": response.text, "request_payload": payload}
    else:
        return {"success": False, "status_code": response.status_code, "error": response.text, "request_payload": payload}


def create_online_video_activity(cookie_string: str, url: str, title: str, link: str,
                                 module_id: int | None = None, syllabus_id: int | None = None, sort: int = 1, description: str = "", completion_criterion_value=None, submit_times=None) -> dict:
    """
    建立「影音教材_影音連結」學習活動（type: online_video）
    """
    headers, cookies = _build_headers(cookie_string, url)

    # 預設值處理
    if completion_criterion_value is None or completion_criterion_value == '' or str(completion_criterion_value).lower() == 'nan':
        completion_criterion_value = 80
    if submit_times is None or submit_times == '' or str(submit_times).lower() == 'nan':
        submit_times = 1

    payload = {
        "type": "online_video",
        "week": 0,
        "teaching_method": 0,
        "title": title,
        "link": link,  # 影片網址
        "knowledge_node_ids": [],
        "publishStatusClass": "publish-type-button unpublished",
        "publishTypeText": "已發布",
        "published": True,
        "isPublishScheduled": False,
        "publishScheduledDate": None,
        "publishScheduledTime": None,
        "description": description,
        "attachments": [],
        "uploads": [],
        "other_resources": [],
        "evaluation_items": [],
        "homework": None,
        "sort": sort,
        "completion_criterion": {"type": "progress", "value": completion_criterion_value},
        "submit_times": submit_times
    }
    if module_id is not None:
        payload["module_id"] = module_id
    if syllabus_id is not None:
        payload["syllabus_id"] = syllabus_id

    try:
        response = requests.post(url, headers=headers, cookies=cookies, json=payload, verify=False)
    except Exception as e:
        return {"success": False, "error": str(e), "request_payload": payload}

    if response.status_code == 201:
        try:
            res = response.json()
            return {
                "success": True,
                "activity_id": res.get("id"),
                "activity_title": res.get("title"),
                "module_id": res.get("module_id"),
                "syllabus_id": res.get("syllabus_id", None)
            }
        except Exception:
            return {"success": False, "error": "無法解析 response", "raw": response.text, "request_payload": payload}
    else:
        return {"success": False, "status_code": response.status_code, "error": response.text, "request_payload": payload}


# 若 log_error 不存在則補上
def log_error(operation_type, item_name, request_params, response_data, error_msg=None):
    try:
        log_dir = "log"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"{log_dir}/{operation_type}_error_{timestamp}.json"
        log_data = {
            "timestamp": datetime.now().isoformat(),
            "operation_type": operation_type,
            "item_name": item_name,
            "request_params": request_params,
            "response_data": response_data,
            "error_msg": error_msg
        }
        with open(log_filename, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)
        print(f"📝 錯誤日誌已記錄: {log_filename}")
    except Exception as e:
        print(f"❌ 記錄錯誤日誌失敗: {e}")

# 主程式測試區塊
if __name__ == "__main__":
    print("🧪 學習活動建立測試")
    print("=" * 50)
    try:
        from config import COOKIE, COURSE_ID, MODULE_ID, get_api_urls
        api_urls = get_api_urls()
    except ImportError as e:
        print(f"❌ 無法導入 config.py: {e}")
        print("請確認 config.py 檔案存在且包含必要的設定")
        exit(1)
    test_params = {
        "cookie_string": COOKIE,
        "url": f"{BASE_URL}/api/courses/{COURSE_ID}/activities",
        "title": "測試學習活動",
        "module_id": MODULE_ID
    }
    
    # 測試不同的課程和單元組合
    test_cases = [
        {
            "name": "測試案例1 - 使用 config.py 的設定",
            "course_id": COURSE_ID,
            "module_id": MODULE_ID,
            "syllabus_id": None,
            "title": "test",
            "link_url": "https://www.google.com"
        },
        {
            "name": "測試案例2 - 使用實際數據的課程ID",
            "course_id": 16402,
            "module_id": 28750,
            "syllabus_id": 12784,
            "title": "1-1-1 從有趣的實例談起",
            "link_url": f"{BASE_URL}/api/uploads/scorm/22?sco=content/w01/1-1-1_p.html"
        },
        {
            "name": "測試案例3 - 只使用章節ID，不用單元ID",
            "course_id": 16402,
            "module_id": 28750,
            "syllabus_id": None,
            "title": "測試活動 - 只用章節",
            "link_url": "https://www.google.com"
        },
        {
            "name": "測試案例4 - 測試錯誤的單元ID",
            "course_id": 16403,
            "module_id": 28760,
            "syllabus_id": 12809,
            "title": "2-0 從疑惑中開始學習",
            "link_url": f"{BASE_URL}/api/uploads/scorm/22?sco=content/w02/2-0_p.html"
        },
        {
            "name": "測試案例5 - 只使用章節ID，不用錯誤的單元ID",
            "course_id": 16403,
            "module_id": 28760,
            "syllabus_id": None,
            "title": "2-0 從疑惑中開始學習 (只用章節)",
            "link_url": f"{BASE_URL}/api/uploads/scorm/22?sco=content/w02/2-0_p.html"
        }
    ]
    
    print(f"📋 測試參數:")
    print(f"  - 課程ID: {COURSE_ID}")
    print(f"  - 章節ID: {MODULE_ID}")
    print(f"  - API URL: {test_params['url']}")
    print(f"  - 活動標題: {test_params['title']}")
    
    # 執行所有測試案例
    for i, test_case in enumerate(test_cases, 1):
        print(f"\n🧪 {test_case['name']}")
        print(f"  - 課程ID: {test_case['course_id']}")
        print(f"  - 章節ID: {test_case['module_id']}")
        print(f"  - 單元ID: {test_case['syllabus_id']}")
        
        # 構建測試 URL
        test_url = f"{BASE_URL}/api/courses/{test_case['course_id']}/activities"
        
        # 測試線上連結活動
        print(f"🔗 測試線上連結活動...")
        link_result = create_link_activity(
            cookie_string=test_params["cookie_string"],
            url=test_url,
            title=test_case["title"],
            link_url=test_case["link_url"],
            module_id=test_case["module_id"],
            syllabus_id=test_case["syllabus_id"]
        )
        
        print(f"📊 線上連結活動結果:")
        print(f"  - 成功: {link_result.get('success', False)}")
        if link_result.get('success', False):
            print(f"  - 活動ID: {link_result.get('activity_id', 'N/A')}")
            print(f"  - 活動標題: {link_result.get('activity_title', 'N/A')}")
        else:
            print(f"  - 錯誤: {link_result.get('error', '未知錯誤')}")
            print(f"  - 狀態碼: {link_result.get('status_code', 'N/A')}")
            # 自動寫入 log
            log_error(
                operation_type="activity",
                item_name=test_case["title"],
                request_params={
                    "title": test_case["title"],
                    "link_url": test_case["link_url"],
                    "module_id": test_case["module_id"],
                    "syllabus_id": test_case["syllabus_id"],
                    "url": test_url,
                    "activity_type": "web_link",
                    "test_case": test_case["name"]
                },
                response_data=link_result,
                error_msg=link_result.get('error', '未知錯誤')
            )