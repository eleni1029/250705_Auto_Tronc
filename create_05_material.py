import os
import requests
import urllib3
from config import get_api_urls

# 抑制 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def upload_material(cookie_string, filename, parent_id=0, file_type="resource"):
    """
    新版 TronClass 上傳資源：
    1. POST /api/uploads 取得 upload_url 和 material_id
    2. PUT 檔案到 upload_url
    """
    file_path = os.path.abspath(filename)
    file_size = os.path.getsize(file_path)
    
    # 從 config 獲取 API URLs
    api_urls = get_api_urls()
    upload_url = api_urls['MATERIAL_UPLOAD_URL']
    
    # 將 Cookie 字串轉為字典
    cookies = dict(item.split("=", 1) for item in cookie_string.split("; "))
    print(f"DEBUG: cookies = {cookies}")
    
    # 步驟一：通知後端
    payload = {
        "name": os.path.basename(filename),
        "size": file_size,
        "parent_id": parent_id,
        "is_scorm": False,
        "is_wmpkg": False,
        "source": "",
        "embed_material_type": "",
        "is_marked_attachment": False
    }
    upload_info = requests.post(
        upload_url,
        json=payload,
        cookies=cookies,
        verify=False
    )
    print("DEBUG: upload_info.status_code =", upload_info.status_code)
    print("DEBUG: upload_info.text =", upload_info.text)
    if upload_info.status_code not in [200, 201]:
        return {"success": False, "step": "notify", "error": upload_info.text}
    upload_info = upload_info.json()
    upload_url = upload_info["upload_url"]
    material_id = upload_info["id"]
    
    # 步驟二：POST 檔案內容
    with open(file_path, "rb") as f:
        response = requests.post(
            upload_url,
            files={'file': (os.path.basename(file_path), f)},
            verify=False
        )
    print("DEBUG: post response.status_code =", response.status_code)
    print("DEBUG: post response.text =", response.text)
    if response.status_code != 200:
        return {"success": False, "step": "upload", "error": response.text, "material_id": material_id}
    return {"success": True, "material_id": material_id}

# 主程式測試區塊
if __name__ == "__main__":
    print("📚 TronClass 新版材料上傳工具測試")
    print("=" * 50)
    try:
        from config import COOKIE
        # 使用完整的 COOKIE 字串
        COOKIE_STRING = COOKIE
    except Exception as e:
        print(f"❌ 取得 COOKIE 失敗: {e}")
        exit(1)
    
    # 測試檔案
    test_filename = "test_file.txt"
    test_content = "這是一個新版 API 測試檔案。"
    with open(test_filename, "w", encoding="utf-8") as f:
        f.write(test_content)
    print(f"✅ 測試檔案已建立: {test_filename}")
    
    # 執行上傳
    result = upload_material(COOKIE_STRING, test_filename)
    if result["success"]:
        print(f"✅ 上傳成功，material_id: {result['material_id']}")
    else:
        print(f"❌ 上傳失敗，步驟: {result.get('step')}，錯誤: {result.get('error')}")
        if "material_id" in result:
            print(f"  - material_id: {result['material_id']}")
    
    # 清理
    try:
        os.remove(test_filename)
        print(f"🧹 測試檔案已清理: {test_filename}")
    except:
        pass
    print("\n🎉 測試完成!")
    print("=" * 50)
    print("新版 TronClass 上傳流程：\n1. POST /api/uploads 取得 upload_url 和 material_id\n2. POST 檔案到 upload_url\n")


