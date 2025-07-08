import os
import glob
import pandas as pd
import re
from datetime import datetime
import time
import json

# 導入配置和創建函數
from config import (
    COOKIE, COURSE_ID, MODULE_ID, SLEEP_SECONDS, 
    get_api_urls, ACTIVITY_TYPE_MAPPING, SUPPORTED_ACTIVITY_TYPES
)
from create_01_course import create_course
from create_02_module import create_module
from create_03_syllabus import create_syllabus
from create_04_activity import create_link_activity, create_reference_activity, create_video_activity, create_audio_activity, create_online_video_activity
from create_05_material import upload_material as upload_and_create_material
from tronc_login import login_and_get_cookie, update_config

def log_error(operation_type, item_name, request_params, response_data, error_msg=None):
    """
    記錄錯誤日誌到 log 資料夾
    
    Args:
        operation_type: 操作類型 (course, module, syllabus, activity, resource)
        item_name: 項目名稱
        request_params: 請求參數
        response_data: API回應資料
        error_msg: 錯誤訊息
    """
    try:
        # 確保 log 資料夾存在
        log_dir = "log"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # 生成時間戳
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 生成日誌檔案名稱
        log_filename = f"{log_dir}/{operation_type}_error_{timestamp}.json"
        
        # 準備日誌資料
        log_data = {
            "timestamp": datetime.now().isoformat(),
            "operation_type": operation_type,
            "item_name": item_name,
            "request_params": request_params,
            "response_data": response_data,
            "error_msg": error_msg
        }
        
        # 寫入日誌檔案
        with open(log_filename, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)
        
        print(f"📝 錯誤日誌已記錄: {log_filename}")
        
    except Exception as e:
        print(f"❌ 記錄錯誤日誌失敗: {e}")

class TronClassCreator:
    def __init__(self):
        self.cookie_string = COOKIE
        self.excel_file = None
        self.result_df = None
        self.resource_df = None
        self.api_urls = get_api_urls()
        self.error_action_policy = 'ask'  # 'ask', 'skip', 'retry'
        self.skipped_items = []  # 記錄略過明細
        self.failed_items = []   # 記錄失敗明細
    
    def check_and_update_cookie(self):
        """檢查並更新 Cookie - 只在需要時才重新登入"""
        print("🔍 檢查 Cookie 狀態...")
        
        # 先嘗試使用現有的 cookie 進行一個簡單的 API 測試
        if self.test_cookie_validity():
            print("✅ 現有 Cookie 仍然有效，無需重新登入")
            return True
        
        # 如果測試失敗，則重新登入
        print("⚠️  Cookie 已過期或無效，開始自動登入...")
        result = login_and_get_cookie()
        
        if result:
            cookie_string, modules = result
            self.cookie_string = cookie_string
            print("✅ 自動登入成功，Cookie 已更新")
            
            # 更新配置文件
            update_config(cookie_string, modules)
            
            # 重新載入 API URLs（可能課程ID有變化）
            self.api_urls = get_api_urls()
            return True
        else:
            print("❌ 自動登入失敗，請檢查帳號密碼設定")
            return False
    
    def test_cookie_validity(self):
        """測試當前 Cookie 是否有效"""
        try:
            # 使用一個簡單的 API 調用來測試 cookie 有效性
            # 這裡使用課程列表 API 作為測試
            import requests
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            
            # 將 Cookie 字串轉為字典
            cookies = dict(item.split("=", 1) for item in self.cookie_string.split("; "))
            
            headers = {
                "Accept": "application/json, text/plain, */*",
                "User-Agent": "Mozilla/5.0",
                "X-Requested-With": "XMLHttpRequest"
            }
            
            # 使用課程列表 API 進行測試
            test_url = "https://wg.tronclass.com/api/course"
            response = requests.get(test_url, headers=headers, cookies=cookies, verify=False, timeout=10)
            
            # 如果狀態碼是 200 或 201，表示 cookie 有效
            if response.status_code in [200, 201]:
                return True
            # 如果狀態碼是 401 或 403，表示認證失敗
            elif response.status_code in [401, 403]:
                return False
            # 其他狀態碼可能是其他問題，我們假設 cookie 仍然有效
            else:
                return True
                
        except Exception as e:
            print(f"⚠️  Cookie 測試時發生錯誤: {e}")
            # 如果測試失敗，我們假設需要重新登入
            return False
    
    def get_extracted_files(self):
        """獲取 to_be_executed 目錄中所有 extracted 檔案，按時間戳排序"""
        pattern = os.path.join('to_be_executed', '*extracted*.xlsx')
        files = glob.glob(pattern)
        
        # 過濾掉暫存檔案（以 ~$ 開頭的檔案）
        files = [f for f in files if not os.path.basename(f).startswith('~$')]
        
        if not files:
            print("❌ 在 to_be_executed 目錄中找不到 extracted 檔案")
            return []
        
        # 根據檔案名中的時間戳排序（最新的在前）
        files.sort(key=self.extract_timestamp_for_sorting, reverse=True)
        return files
    
    def extract_timestamp_for_sorting(self, filename):
        """從檔案名中提取時間戳用於排序"""
        match = re.search(r'extracted_(\d{8}_\d{6})', filename)
        if match:
            timestamp_str = match.group(1)
            # 將時間戳轉換為可比較的格式 (YYYYMMDD_HHMMSS)
            return timestamp_str
        # 如果沒有時間戳，使用檔案修改時間作為備用
        return datetime.fromtimestamp(os.path.getmtime(filename)).strftime('%Y%m%d_%H%M%S')
    
    def extract_timestamp(self, filename):
        """從檔案名中提取時間戳用於顯示"""
        match = re.search(r'extracted_(\d{8}_\d{6})', filename)
        if match:
            timestamp_str = match.group(1)
            # 格式化時間戳顯示 (YYYY-MM-DD HH:MM:SS)
            try:
                dt = datetime.strptime(timestamp_str, '%Y%m%d_%H%M%S')
                return dt.strftime('%Y-%m-%d %H:%M:%S')
            except ValueError:
                return timestamp_str
        return "unknown"
    
    def select_file(self):
        """讓用戶選擇檔案"""
        files = self.get_extracted_files()
        if not files:
            return None
            
        print("\n📁 找到以下 extracted 檔案（按時間戳排序，最新的在前）：")
        for i, file in enumerate(files, 1):
            filename = os.path.basename(file)
            print(f"{i}. {filename}")
        
        while True:
            try:
                choice = input(f"\n請選擇檔案 (1-{len(files)})，或按 Enter 選擇最新的: ").strip()
                if not choice:
                    return files[0]  # 預設選擇最新的
                
                choice_num = int(choice)
                if 1 <= choice_num <= len(files):
                    return files[choice_num - 1]
                else:
                    print(f"❌ 請輸入 1-{len(files)} 之間的數字")
            except ValueError:
                print("❌ 請輸入有效的數字")
    
    def load_data(self):
        """載入 Excel 數據"""
        try:
            self.result_df = pd.read_excel(self.excel_file, sheet_name='Result')
            self.resource_df = pd.read_excel(self.excel_file, sheet_name='Resource')
            
            # 確保 ID 欄位是整數類型
            id_columns = ['ID', '所屬課程ID', '所屬章節ID', '所屬單元ID', '資源ID']
            for col in id_columns:
                if col in self.result_df.columns:
                    # 將非空值轉換為整數，空值保持為 NaN
                    numeric_col = pd.to_numeric(self.result_df[col], errors='coerce')
                    self.result_df[col] = numeric_col
                if col in self.resource_df.columns:
                    numeric_col = pd.to_numeric(self.resource_df[col], errors='coerce')
                    self.resource_df[col] = numeric_col
            
            print(f"✅ 已載入數據: Result({len(self.result_df)}行), Resource({len(self.resource_df)}行)")
            return True
        except Exception as e:
            print(f"❌ 載入數據失敗: {e}")
            return False
    
    def select_operation(self):
        """讓用戶選擇操作"""
        operations = [
            "建立文件內所有元素",
            "建立所有課程", 
            "建立所有章節",
            "建立所有單元",
            "建立所有學習活動",
            "建立特定類型學習活動",
            "建立所有資源"
        ]
        
        print("\n📋 請選擇要進行的操作：")
        for i, op in enumerate(operations, 1):
            print(f"{i}. {op}")
        
        while True:
            try:
                choice = input(f"\n請選擇操作 (1-{len(operations)}): ").strip()
                choice_num = int(choice)
                if 1 <= choice_num <= len(operations):
                    selected_op = operations[choice_num - 1]
                    
                    # 如果是特定類型學習活動，需要進一步選擇類型
                    if selected_op == "建立特定類型學習活動":
                        activity_type = self.select_activity_type()
                        if activity_type:
                            return selected_op, activity_type
                        else:
                            continue
                    
                    return selected_op, None
                else:
                    print(f"❌ 請輸入 1-{len(operations)} 之間的數字")
            except ValueError:
                print("❌ 請輸入有效的數字")
    
    def select_activity_type(self):
        """選擇學習活動類型"""
        print("\n📝 請選擇學習活動類型：")
        for i, activity_type in enumerate(SUPPORTED_ACTIVITY_TYPES, 1):
            print(f"{i}. {activity_type}")
        
        while True:
            try:
                choice = input(f"\n請選擇類型 (1-{len(SUPPORTED_ACTIVITY_TYPES)}): ").strip()
                choice_num = int(choice)
                if 1 <= choice_num <= len(SUPPORTED_ACTIVITY_TYPES):
                    return SUPPORTED_ACTIVITY_TYPES[choice_num - 1]
                else:
                    print(f"❌ 請輸入 1-{len(SUPPORTED_ACTIVITY_TYPES)} 之間的數字")
            except ValueError:
                print("❌ 請輸入有效的數字")
    
    def analyze_operation(self, operation, activity_type=None):
        """分析即將執行的操作並顯示統計"""
        stats = {}
        
        if operation == "建立所有課程":
            courses = self.result_df[
                (self.result_df['類型'] == '課程') & 
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            stats['課程'] = list(courses['名稱'])
            
        elif operation == "建立所有章節":
            modules = self.result_df[
                (self.result_df['類型'] == '章節') & 
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            stats['章節'] = list(modules['名稱'])
            
        elif operation == "建立所有單元":
            syllabi = self.result_df[
                (self.result_df['類型'] == '單元') & 
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            stats['單元'] = list(syllabi['名稱'])
            
        elif operation == "建立所有學習活動":
            activities = self.result_df[
                (self.result_df['類型'] == '學習活動') & 
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            # 按類型分組
            for act_type in SUPPORTED_ACTIVITY_TYPES:
                type_activities = activities[activities['學習活動類型'] == act_type]
                if not type_activities.empty:
                    stats[f'學習活動-{act_type}'] = list(type_activities['名稱'])
                    
        elif operation == "建立特定類型學習活動":
            activities = self.result_df[
                (self.result_df['類型'] == '學習活動') & 
                (self.result_df['學習活動類型'] == activity_type) &
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            stats[f'學習活動-{activity_type}'] = list(activities['名稱'])
            
        elif operation == "建立所有資源":
            resources = self.resource_df[
                self.resource_df['資源ID'].isna() | (self.resource_df['資源ID'] == '')
            ]
            stats['資源'] = list(resources['檔案名稱'])
            
        elif operation == "建立文件內所有元素":
            # 統計所有類型
            resources = self.resource_df[
                self.resource_df['資源ID'].isna() | (self.resource_df['資源ID'] == '')
            ]
            if not resources.empty:
                stats['資源'] = list(resources['檔案名稱'])
                
            for item_type in ['課程', '章節', '單元']:
                items = self.result_df[
                    (self.result_df['類型'] == item_type) & 
                    (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
                ]
                if not items.empty:
                    stats[item_type] = list(items['名稱'])
            
            activities = self.result_df[
                (self.result_df['類型'] == '學習活動') & 
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            for act_type in SUPPORTED_ACTIVITY_TYPES:
                type_activities = activities[activities['學習活動類型'] == act_type]
                if not type_activities.empty:
                    stats[f'學習活動-{act_type}'] = list(type_activities['名稱'])
        
        return stats
    
    def confirm_operation(self, stats):
        """確認操作"""
        if not stats:
            print("❌ 沒有找到需要建立的項目")
            return False
            
        print(f"\n📊 即將建立以下項目：")
        total_count = 0
        for item_type, items in stats.items():
            print(f"  {item_type}: {len(items)} 個")
            total_count += len(items)
            
        print(f"\n📝 詳細清單：")
        for item_type, items in stats.items():
            print(f"\n{item_type}:")
            for i, item in enumerate(items[:5], 1):  # 只顯示前5個
                print(f"  {i}. {item}")
            if len(items) > 5:
                print(f"  ... 等共 {len(items)} 個")
        
        print(f"\n總計: {total_count} 個項目")
        
        while True:
            confirm = input(f"\n確認開始建立？(y/n) [預設: y]: ").strip().lower()
            if confirm in ['y', 'yes', '是', '']:
                # 新增詢問預設錯誤處理策略
                print("\n⚠️  如果遇到無法成功建立的資源或學習活動，預設操作？")
                print("  1. 每次詢問 (預設)")
                print("  2. 一律略過")
                print("  3. 一律重試")
                while True:
                    policy = input("請選擇 (1=每次詢問, 2=一律略過, 3=一律重試) [預設: 1]: ").strip()
                    if policy in ['', '1']:
                        self.error_action_policy = 'ask'
                        break
                    elif policy == '2':
                        self.error_action_policy = 'skip'
                        break
                    elif policy == '3':
                        self.error_action_policy = 'retry'
                        break
                    else:
                        print("請輸入 1、2、3 或直接 Enter")
                return True
            elif confirm in ['n', 'no', '否']:
                return False
            else:
                print("請輸入 y 或 n，或直接按 Enter 確認")

    def save_excel(self):
        """保存 Excel 檔案"""
        try:
            with pd.ExcelWriter(self.excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                self.result_df.to_excel(writer, sheet_name='Result', index=False)
                self.resource_df.to_excel(writer, sheet_name='Resource', index=False)
            print("✅ Excel 檔案已更新")
            return True
        except Exception as e:
            print(f"❌ 保存 Excel 檔案失敗: {e}")
            return False
    
    def update_result_id(self, row_index, new_id, status="success"):
        """更新 Result 表中的 ID"""
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        if status == "success":
            # 確保 ID 是整數
            if new_id is not None:
                new_id = int(new_id)
            self.result_df.loc[row_index, 'ID'] = new_id
        else:
            self.result_df.loc[row_index, 'ID'] = "創建失敗，用戶已略過"
            
        # 確保最後修改時間欄位是字串類型
        if '最後修改時間' not in self.result_df.columns:
            self.result_df['最後修改時間'] = ''
        self.result_df.loc[row_index, '最後修改時間'] = str(current_time)
        
        # 如果成功，更新相關項目的所屬ID
        if status == "success" and new_id is not None:
            item_name = self.result_df.loc[row_index, '名稱']
            item_type = self.result_df.loc[row_index, '類型']
            
            if item_type == '課程':
                # 更新所有 所屬課程 為此名稱的項目
                mask = self.result_df['所屬課程'] == item_name
                self.result_df.loc[mask, '所屬課程ID'] = new_id
                
            elif item_type == '章節':
                # 更新所有 所屬章節 為此名稱的項目
                mask = self.result_df['所屬章節'] == item_name
                self.result_df.loc[mask, '所屬章節ID'] = new_id
                
            elif item_type == '單元':
                # 更新所有 所屬單元 為此名稱的項目
                mask = self.result_df['所屬單元'] == item_name
                self.result_df.loc[mask, '所屬單元ID'] = new_id
    
    def update_resource_id(self, row_index, new_id, status="success"):
        """更新 Resource 表中的 ID"""
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        if status == "success":
            # 確保 ID 是整數
            if new_id is not None:
                new_id = int(new_id)
            self.resource_df.loc[row_index, '資源ID'] = new_id
            
            # 同時更新 Result 表中相同檔案路徑的項目
            file_path = self.resource_df.loc[row_index, '檔案路徑']
            mask = self.result_df['檔案路徑'] == file_path
            self.result_df.loc[mask, '資源ID'] = new_id
        else:
            self.resource_df.loc[row_index, '資源ID'] = "創建失敗，用戶已略過"
            
        # 確保最後修改時間欄位是字串類型
        if '最後修改時間' not in self.resource_df.columns:
            self.resource_df['最後修改時間'] = ''
        self.resource_df.loc[row_index, '最後修改時間'] = str(current_time)
    
    def check_missing_ids(self, operation):
        """檢查缺失的 ID 並提供預設值選項"""
        need_course_id = operation in ["建立所有章節", "建立所有單元", "建立所有學習活動"]
        need_module_id = operation in ["建立所有單元", "建立所有學習活動"]
        
        if need_course_id:
            # 檢查是否有空的課程ID
            if operation == "建立所有章節":
                items = self.result_df[
                    (self.result_df['類型'] == '章節') & 
                    (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
                ]
            elif operation == "建立所有單元":
                items = self.result_df[
                    (self.result_df['類型'] == '單元') & 
                    (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
                ]
            else:  # 學習活動
                items = self.result_df[
                    (self.result_df['類型'] == '學習活動') & 
                    (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
                ]
            
            empty_course_ids = items[items['所屬課程ID'].isna() | (items['所屬課程ID'] == '')]
            if not empty_course_ids.empty:
                print(f"\n⚠️  發現 {len(empty_course_ids)} 個項目沒有所屬課程ID")
                use_default = input(f"是否使用預設課程ID ({COURSE_ID})？(y/n) [預設: y]: ").strip().lower()
                if use_default in ['y', 'yes', '是', '']:  # 空字串表示按 Enter
                    self.result_df.loc[empty_course_ids.index, '所屬課程ID'] = int(COURSE_ID)
                    # 確保更新後的欄位是數值類型
                    self.result_df['所屬課程ID'] = pd.to_numeric(self.result_df['所屬課程ID'], errors='coerce')
                    print(f"✅ 已設定預設課程ID: {COURSE_ID}")
                else:
                    print("❌ 取消操作")
                    return False
        
        if need_module_id:
            # 檢查是否有空的章節ID
            if operation == "建立所有單元":
                items = self.result_df[
                    (self.result_df['類型'] == '單元') & 
                    (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
                ]
            else:  # 學習活動
                items = self.result_df[
                    (self.result_df['類型'] == '學習活動') & 
                    (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
                ]
            
            empty_module_ids = items[items['所屬章節ID'].isna() | (items['所屬章節ID'] == '')]
            if not empty_module_ids.empty:
                print(f"\n⚠️  發現 {len(empty_module_ids)} 個項目沒有所屬章節ID")
                use_default = input(f"是否使用預設章節ID ({MODULE_ID})？(y/n) [預設: y]: ").strip().lower()
                if use_default in ['y', 'yes', '是', '']:  # 空字串表示按 Enter
                    self.result_df.loc[empty_module_ids.index, '所屬章節ID'] = int(MODULE_ID)
                    # 確保更新後的欄位是數值類型
                    self.result_df['所屬章節ID'] = pd.to_numeric(self.result_df['所屬章節ID'], errors='coerce')
                    print(f"✅ 已設定預設章節ID: {MODULE_ID}")
                else:
                    print("❌ 取消操作")
                    return False
        
        return True
    
    def handle_error(self, item_name, error_msg):
        """處理錯誤"""
        # 檢查是否為認證錯誤
        if self.is_authentication_error(error_msg):
            print(f"\n🔐 檢測到認證錯誤，嘗試重新登入...")
            if self.check_and_update_cookie():
                print(f"✅ 重新登入成功，請重新執行操作")
                return "retry"
            else:
                print(f"❌ 重新登入失敗")
                return False
        print(f"\n❌ 建立 '{item_name}' 時發生錯誤: {error_msg}")
        # 根據 policy 決定自動略過/重試/詢問
        if self.error_action_policy == 'skip':
            print("[自動略過]")
            self.skipped_items.append({'item': item_name, 'reason': error_msg})
            return True
        elif self.error_action_policy == 'retry':
            print("[自動重試]")
            return "retry"
        else:
            while True:
                choice = input("是否略過此項目繼續執行？(y=略過繼續, n=終止程序) [預設: y]: ").strip().lower()
                if choice in ['y', 'yes', '是', '']:
                    self.skipped_items.append({'item': item_name, 'reason': error_msg})
                    return True
                elif choice in ['n', 'no', '否']:
                    self.failed_items.append({'item': item_name, 'reason': error_msg})
                    return False
                else:
                    print("請輸入 y 或 n，或直接按 Enter 略過")
    
    def is_authentication_error(self, error_msg):
        """檢查是否為認證相關錯誤"""
        auth_keywords = [
            'unauthorized', '401', '403', 'forbidden', 'authentication',
            'login', 'session', 'cookie', 'token', '認證', '登入', '會話'
        ]
        
        error_lower = error_msg.lower()
        return any(keyword in error_lower for keyword in auth_keywords)
    
    def create_single_course(self, row_index):
        """建立單一課程"""
        row = self.result_df.loc[row_index]
        course_name = row['名稱']
        
        print(f"📝 正在建立課程: {course_name}")
        
        # 準備請求參數
        request_params = {
            "course_name": course_name,
            "url": self.api_urls['COURSE_CREATE_URL']
        }
        
        try:
            result = create_course(
                cookie_string=self.cookie_string,
                url=self.api_urls['COURSE_CREATE_URL'],
                course_name=course_name
            )
            
            if result['success']:
                course_id = result['course_id']
                print(f"✅ 課程建立成功: {course_name} (ID: {int(course_id)})")
                self.update_result_id(row_index, course_id)
                return True
            else:
                error_msg = result.get('error', '未知錯誤')
                # 記錄錯誤日誌
                log_error(
                    operation_type="course",
                    item_name=course_name,
                    request_params=request_params,
                    response_data=result,
                    error_msg=error_msg
                )
                
                error_action = self.handle_error(course_name, error_msg)
                if error_action == "retry":
                    # 重新登入成功，重試當前操作
                    return self.create_single_course(row_index)
                elif error_action:
                    self.update_result_id(row_index, None, "failed")
                    return True
                else:
                    return False
                    
        except Exception as e:
            # 記錄例外錯誤日誌
            log_error(
                operation_type="course",
                item_name=course_name,
                request_params=request_params,
                response_data={"exception": str(e)},
                error_msg=str(e)
            )
            
            error_action = self.handle_error(course_name, str(e))
            if error_action == "retry":
                # 重新登入成功，重試當前操作
                return self.create_single_course(row_index)
            elif error_action:
                self.update_result_id(row_index, None, "failed")
                return True
            else:
                return False
    
    def create_single_module(self, row_index):
        """建立單一章節"""
        row = self.result_df.loc[row_index]
        module_name = row['名稱']
        course_id = row['所屬課程ID']
        
        print(f"📝 正在建立章節: {module_name} (課程ID: {int(course_id) if pd.notna(course_id) and course_id != '' else 'None'})")
        
        # 動態構建章節建立的 API URL
        module_url = f"https://wg.tronclass.com/api/course/{int(course_id)}/module"
        
        # 準備請求參數
        request_params = {
            "module_name": module_name,
            "course_id": int(course_id) if pd.notna(course_id) and course_id != '' else None,
            "url": module_url
        }
        
        try:
            result = create_module(
                cookie_string=self.cookie_string,
                url=module_url,
                module_name=module_name,
                course_id=int(course_id) if pd.notna(course_id) and course_id != '' else None
            )
            
            if result['success']:
                module_id = result['module_id']
                print(f"✅ 章節建立成功: {module_name} (ID: {int(module_id)})")
                self.update_result_id(row_index, module_id)
                return True
            else:
                error_msg = result.get('error', '未知錯誤')
                # 記錄錯誤日誌
                log_error(
                    operation_type="module",
                    item_name=module_name,
                    request_params=request_params,
                    response_data=result,
                    error_msg=error_msg
                )
                
                error_action = self.handle_error(module_name, error_msg)
                if error_action == "retry":
                    # 重新登入成功，重試當前操作
                    return self.create_single_module(row_index)
                elif error_action:
                    self.update_result_id(row_index, None, "failed")
                    return True
                else:
                    return False
                    
        except Exception as e:
            # 記錄例外錯誤日誌
            log_error(
                operation_type="module",
                item_name=module_name,
                request_params=request_params,
                response_data={"exception": str(e)},
                error_msg=str(e)
            )
            
            error_action = self.handle_error(module_name, str(e))
            if error_action == "retry":
                # 重新登入成功，重試當前操作
                return self.create_single_module(row_index)
            elif error_action:
                self.update_result_id(row_index, None, "failed")
                return True
            else:
                return False
    
    def create_single_syllabus(self, row_index):
        """建立單一單元"""
        row = self.result_df.loc[row_index]
        summary = row['名稱']
        module_id = row['所屬章節ID']
        course_id = row['所屬課程ID']
        
        print(f"📝 正在建立單元: {summary} (章節ID: {int(module_id) if pd.notna(module_id) and module_id != '' else 'None'})")
        
        # 準備請求參數
        request_params = {
            "summary": summary,
            "module_id": int(module_id) if pd.notna(module_id) and module_id != '' else None,
            "course_id": int(course_id) if pd.notna(course_id) and course_id != '' else None,
            "url": self.api_urls['SYLLABUS_CREATE_URL']
        }
        
        try:
            result = create_syllabus(
                cookie_string=self.cookie_string,
                url=self.api_urls['SYLLABUS_CREATE_URL'],
                module_id=int(module_id) if pd.notna(module_id) and module_id != '' else None,
                summary=summary,
                course_id=int(course_id) if pd.notna(course_id) and course_id != '' else None
            )
            
            if result['success']:
                syllabus_id = result['syllabus_id']
                print(f"✅ 單元建立成功: {summary} (ID: {int(syllabus_id)})")
                self.update_result_id(row_index, syllabus_id)
                return True
            else:
                error_msg = result.get('error', '未知錯誤')
                # 記錄錯誤日誌
                log_error(
                    operation_type="syllabus",
                    item_name=summary,
                    request_params=request_params,
                    response_data=result,
                    error_msg=error_msg
                )
                
                error_action = self.handle_error(summary, error_msg)
                if error_action == "retry":
                    # 重新登入成功，重試當前操作
                    return self.create_single_syllabus(row_index)
                elif error_action:
                    self.update_result_id(row_index, None, "failed")
                    return True
                else:
                    return False
                    
        except Exception as e:
            # 記錄例外錯誤日誌
            log_error(
                operation_type="syllabus",
                item_name=summary,
                request_params=request_params,
                response_data={"exception": str(e)},
                error_msg=str(e)
            )
            
            error_action = self.handle_error(summary, str(e))
            if error_action == "retry":
                # 重新登入成功，重試當前操作
                return self.create_single_syllabus(row_index)
            elif error_action:
                self.update_result_id(row_index, None, "failed")
                return True
            else:
                return False
    
    def create_single_activity(self, row_index):
        """建立單一學習活動"""
        row = self.result_df.loc[row_index]
        title = row['名稱']
        activity_type = row['學習活動類型']
        module_id = row['所屬章節ID']
        syllabus_id = row['所屬單元ID']
        course_id = row['所屬課程ID']  # 新增：獲取課程ID
        
        # 轉換活動類型
        api_type = ACTIVITY_TYPE_MAPPING.get(activity_type)
        if not api_type:
            print(f"❌ 不支援的學習活動類型: {activity_type}")
            return False
        
        print(f"📝 正在建立學習活動: {title} (類型: {activity_type})")
        
        # 檢查單元ID是否有效
        valid_syllabus_id = None
        if pd.notna(syllabus_id) and syllabus_id != '' and syllabus_id is not None:
            valid_syllabus_id = int(syllabus_id)
        
        # 詳細日誌記錄
        print(f"🔍 參數詳情:")
        print(f"  - 課程ID: {int(course_id) if pd.notna(course_id) and course_id != '' else 'None'}")
        print(f"  - 章節ID: {int(module_id) if pd.notna(module_id) and module_id != '' else 'None'}")
        print(f"  - 單元ID: {valid_syllabus_id if valid_syllabus_id else 'None (無效或空值)'}")
        
        try:
            # 動態構建學習活動建立的 API URL
            activity_url = f"https://wg.tronclass.com/api/courses/{int(course_id)}/activities"
            print(f"  - API URL: {activity_url}")
            
            if api_type == 'web_link':
                # 線上連結活動（包括影音連結）
                link_url = row['網址路徑']
                print(f"  - 連結網址: {link_url}")
                
                # 準備請求參數
                request_params = {
                    "title": title,
                    "link_url": link_url,
                    "module_id": int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    "syllabus_id": valid_syllabus_id,
                    "url": activity_url,
                    "activity_type": "web_link"
                }
                
                result = create_link_activity(
                    cookie_string=self.cookie_string,
                    url=activity_url,
                    title=title,
                    link_url=link_url,
                    module_id=int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    syllabus_id=valid_syllabus_id
                )
            elif api_type == 'online_video':
                # 影音教材_影音連結
                link = row['網址路徑']
                print(f"  - 影音連結網址: {link}")
                request_params = {
                    "title": title,
                    "link": link,
                    "module_id": int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    "syllabus_id": valid_syllabus_id,
                    "url": activity_url,
                    "activity_type": "online_video"
                }
                result = create_online_video_activity(
                    cookie_string=self.cookie_string,
                    url=activity_url,
                    title=title,
                    link=link,
                    module_id=int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    syllabus_id=valid_syllabus_id
                )
            elif api_type == 'video':
                # 影音教材_影片活動
                video_url = row['網址路徑']
                print(f"  - 影片網址: {video_url}")
                
                # 準備請求參數
                request_params = {
                    "title": title,
                    "video_url": video_url,
                    "module_id": int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    "syllabus_id": valid_syllabus_id,
                    "url": activity_url,
                    "activity_type": "video"
                }
                
                result = create_video_activity(
                    cookie_string=self.cookie_string,
                    url=activity_url,
                    title=title,
                    video_url=video_url,
                    module_id=int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    syllabus_id=valid_syllabus_id
                )
            elif api_type == 'audio':
                # 影音教材_音訊活動
                audio_url = row['網址路徑']
                print(f"  - 音訊網址: {audio_url}")
                
                # 準備請求參數
                request_params = {
                    "title": title,
                    "audio_url": audio_url,
                    "module_id": int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    "syllabus_id": valid_syllabus_id,
                    "url": activity_url,
                    "activity_type": "audio"
                }
                
                result = create_audio_activity(
                    cookie_string=self.cookie_string,
                    url=activity_url,
                    title=title,
                    audio_url=audio_url,
                    module_id=int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    syllabus_id=valid_syllabus_id
                )
            else:
                # 參考資料活動（material 類型）
                upload_id = row['資源ID']
                file_path = row['檔案路徑']
                upload_name = os.path.basename(file_path) if file_path else ""
                
                print(f"  - 資源ID: {int(upload_id) if pd.notna(upload_id) and upload_id != '' else 'None'}")
                print(f"  - 檔案路徑: {file_path}")
                print(f"  - 檔案名稱: {upload_name}")
                
                # 準備請求參數
                request_params = {
                    "title": title,
                    "module_id": int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    "syllabus_id": valid_syllabus_id,
                    "upload_id": int(upload_id) if pd.notna(upload_id) and upload_id != '' else None,
                    "upload_name": upload_name,
                    "url": activity_url,
                    "activity_type": "material"
                }
                
                result = create_reference_activity(
                    cookie_string=self.cookie_string,
                    url=activity_url,
                    title=title,
                    module_id=int(module_id) if pd.notna(module_id) and module_id != '' else None,
                    syllabus_id=valid_syllabus_id,
                    upload_id=int(upload_id) if pd.notna(upload_id) and upload_id != '' else None,
                    upload_name=upload_name
                )
            
            # 記錄API回應
            print(f"🔍 API回應:")
            print(f"  - 成功: {result.get('success', False)}")
            if not result.get('success', False):
                print(f"  - 錯誤: {result.get('error', '未知錯誤')}")
                print(f"  - 狀態碼: {result.get('status_code', 'N/A')}")
                if 'raw' in result:
                    print(f"  - 原始回應: {result['raw']}")
                
                # 記錄錯誤日誌
                log_error(
                    operation_type="activity",
                    item_name=title,
                    request_params=request_params,
                    response_data=result,
                    error_msg=result.get('error', '未知錯誤')
                )
            else:
                print(f"  - 活動ID: {result.get('activity_id', 'N/A')}")
                print(f"  - 課程ID: {result.get('course_id', 'N/A')}")
                print(f"  - 章節ID: {result.get('module_id', 'N/A')}")
                print(f"  - 單元ID: {result.get('syllabus_id', 'N/A')}")
            
            if result['success']:
                activity_id = result['activity_id']
                print(f"✅ 學習活動建立成功: {title} (ID: {int(activity_id)})")
                self.update_result_id(row_index, activity_id)
                return True
            else:
                error_msg = result.get('error', '未知錯誤')
                
                error_action = self.handle_error(title, error_msg)
                if error_action == "retry":
                    # 重新登入成功，重試當前操作
                    return self.create_single_activity(row_index)
                elif error_action:
                    self.update_result_id(row_index, None, "failed")
                    return True
                else:
                    return False
                    
        except Exception as e:
            print(f"🔍 例外詳情: {str(e)}")
            # 記錄例外錯誤日誌
            log_error(
                operation_type="activity",
                item_name=title,
                request_params=request_params if 'request_params' in locals() else {"error": "無法獲取請求參數"},
                response_data={"exception": str(e)},
                error_msg=str(e)
            )
            
            error_action = self.handle_error(title, str(e))
            if error_action == "retry":
                # 重新登入成功，重試當前操作
                return self.create_single_activity(row_index)
            elif error_action:
                self.update_result_id(row_index, None, "failed")
                return True
            else:
                return False
    
    def create_single_resource(self, row_index):
        """建立單一資源"""
        row = self.resource_df.loc[row_index]
        title = row['檔案名稱']
        file_path = row['檔案路徑']
        
        print(f"📝 正在建立資源: {title}")
        
        # 從完整 cookie 字串中提取 session
        import re
        m = re.search(r'session=([^;]+)', self.cookie_string)
        if not m:
            print(f"❌ 無法從 cookie 中提取 session")
            return False
        session_cookie = m.group(1)
        print(f"DEBUG: 提取的 session_cookie = {session_cookie[:20]}...")  # 只顯示前20個字符
        
        # 準備請求參數
        request_params = {
            "title": title,
            "file_path": file_path
        }
        
        try:
            result = upload_and_create_material(
                cookie_string=self.cookie_string,
                filename=file_path,
                parent_id=0,
                file_type="resource"
            )
            
            if result['success']:
                material_id = result['material_id']
                print(f"✅ 資源建立成功: {title} (ID: {int(material_id)})")
                self.update_resource_id(row_index, material_id)
                return True
            else:
                error_msg = result.get('error', '未知錯誤')
                # 記錄錯誤日誌
                log_error(
                    operation_type="resource",
                    item_name=title,
                    request_params=request_params,
                    response_data=result,
                    error_msg=error_msg
                )
                
                error_action = self.handle_error(title, error_msg)
                if error_action == "retry":
                    # 重新登入成功，重試當前操作
                    return self.create_single_resource(row_index)
                elif error_action:
                    self.update_resource_id(row_index, None, "failed")
                    return True
                else:
                    return False
                    
        except Exception as e:
            # 記錄例外錯誤日誌
            log_error(
                operation_type="resource",
                item_name=title,
                request_params=request_params,
                response_data={"exception": str(e)},
                error_msg=str(e)
            )
            
            error_action = self.handle_error(title, str(e))
            if error_action == "retry":
                # 重新登入成功，重試當前操作
                return self.create_single_resource(row_index)
            elif error_action:
                self.update_resource_id(row_index, None, "failed")
                return True
            else:
                return False
    
    def execute_operation(self, operation, activity_type=None):
        """執行選定的操作"""
        success_count = 0
        total_count = 0
        
        if operation == "建立所有資源":
            # 建立所有資源
            resources = self.resource_df[
                self.resource_df['資源ID'].isna() | (self.resource_df['資源ID'] == '')
            ]
            
            for idx in resources.index:
                total_count += 1
                if self.create_single_resource(idx):
                    success_count += 1
                    time.sleep(SLEEP_SECONDS)
                    # 每次創建後保存
                    self.save_excel()
                else:
                    break  # 用戶選擇終止
                    
        elif operation == "建立所有課程":
            # 建立所有課程
            courses = self.result_df[
                (self.result_df['類型'] == '課程') & 
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            
            for idx in courses.index:
                total_count += 1
                if self.create_single_course(idx):
                    success_count += 1
                    time.sleep(SLEEP_SECONDS)
                    self.save_excel()
                else:
                    break
                    
        elif operation == "建立所有章節":
            # 建立所有章節
            modules = self.result_df[
                (self.result_df['類型'] == '章節') & 
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            
            for idx in modules.index:
                total_count += 1
                if self.create_single_module(idx):
                    success_count += 1
                    time.sleep(SLEEP_SECONDS)
                    self.save_excel()
                else:
                    break
                    
        elif operation == "建立所有單元":
            # 建立所有單元
            syllabi = self.result_df[
                (self.result_df['類型'] == '單元') & 
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            
            for idx in syllabi.index:
                total_count += 1
                if self.create_single_syllabus(idx):
                    success_count += 1
                    time.sleep(SLEEP_SECONDS)
                    self.save_excel()
                else:
                    break
                    
        elif operation == "建立所有學習活動":
            # 先檢查參考檔案類型的活動是否需要上傳資源
            activities = self.result_df[
                (self.result_df['類型'] == '學習活動') & 
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            
            # 檢查參考檔案活動（需要上傳資源的活動）
            reference_activities = activities[
                activities['學習活動類型'].isin(['參考檔案_圖片', '參考檔案_PDF'])
            ]
            
            if not reference_activities.empty:
                # 檢查是否有需要上傳的資源
                missing_resources = []
                for idx, row in reference_activities.iterrows():
                    file_path = row['檔案路徑']
                    if pd.isna(row['資源ID']) or row['資源ID'] == '':
                        # 檢查 resource 表中是否已有此檔案的ID
                        resource_match = self.resource_df[
                            (self.resource_df['檔案路徑'] == file_path) & 
                            (~self.resource_df['資源ID'].isna()) & 
                            (self.resource_df['資源ID'] != '')
                        ]
                        
                        if resource_match.empty:
                            missing_resources.append((file_path, row['名稱']))
                
                if missing_resources:
                    print(f"\n📤 發現需要先上傳的資源：")
                    for i, (file_path, activity_name) in enumerate(missing_resources, 1):
                        print(f"  {i}. {os.path.basename(file_path)} (用於活動: {activity_name})")
                    
                    confirm_upload = input(f"\n確認先上傳這 {len(missing_resources)} 個資源？(y/n) [預設: y]: ").strip().lower()
                    if confirm_upload not in ['y', 'yes', '是', '']:  # 空字串表示按 Enter
                        print("❌ 取消操作")
                        return 0, 0
                    
                    # 上傳缺失的資源
                    for file_path, activity_name in missing_resources:
                        # 檢查是否已存在於 resource 表中
                        existing = self.resource_df[self.resource_df['檔案路徑'] == file_path]
                        if not existing.empty:
                            # 更新現有記錄
                            resource_idx = existing.index[0]
                        else:
                            # 添加新記錄
                            new_row = {
                                '檔案名稱': os.path.splitext(os.path.basename(file_path))[0],
                                '檔案路徑': file_path,
                                '資源ID': '',
                                '最後修改時間': '',
                                '來源Sheet': 'auto_generated'
                            }
                            resource_idx = len(self.resource_df)
                            self.resource_df.loc[resource_idx] = new_row
                        
                        if self.create_single_resource(resource_idx):
                            time.sleep(SLEEP_SECONDS)
                            self.save_excel()
                        else:
                            print("❌ 資源上傳失敗，終止操作")
                            return 0, 0
            
            # 建立學習活動
            for idx in activities.index:
                total_count += 1
                if self.create_single_activity(idx):
                    success_count += 1
                    time.sleep(SLEEP_SECONDS)
                    self.save_excel()
                else:
                    break
                    
        elif operation == "建立特定類型學習活動":
            # 建立特定類型學習活動
            activities = self.result_df[
                (self.result_df['類型'] == '學習活動') & 
                (self.result_df['學習活動類型'] == activity_type) &
                (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
            ]
            
            for idx in activities.index:
                total_count += 1
                if self.create_single_activity(idx):
                    success_count += 1
                    time.sleep(SLEEP_SECONDS)
                    self.save_excel()
                else:
                    break
                    
        elif operation == "建立文件內所有元素":
            # 先建立所有資源
            resources = self.resource_df[
                self.resource_df['資源ID'].isna() | (self.resource_df['資源ID'] == '')
            ]
            
            print(f"\n🔄 第1步：建立所有資源 ({len(resources)} 個)")
            for idx in resources.index:
                total_count += 1
                if self.create_single_resource(idx):
                    success_count += 1
                    time.sleep(SLEEP_SECONDS)
                    self.save_excel()
                else:
                    return success_count, total_count
            
            # 按順序建立課程結構元素
            structure_types = ['課程', '章節', '單元', '學習活動']
            
            for i, item_type in enumerate(structure_types, 2):
                items = self.result_df[
                    (self.result_df['類型'] == item_type) & 
                    (self.result_df['ID'].isna() | (self.result_df['ID'] == ''))
                ]
                
                if items.empty:
                    continue
                    
                print(f"\n🔄 第{i}步：建立所有{item_type} ({len(items)} 個)")
                
                for idx in items.index:
                    total_count += 1
                    
                    if item_type == '課程':
                        success = self.create_single_course(idx)
                    elif item_type == '章節':
                        success = self.create_single_module(idx)
                    elif item_type == '單元':
                        success = self.create_single_syllabus(idx)
                    else:  # 學習活動
                        success = self.create_single_activity(idx)
                    
                    if success:
                        success_count += 1
                        time.sleep(SLEEP_SECONDS)
                        self.save_excel()
                    else:
                        return success_count, total_count
        
        return success_count, total_count
    
    def run(self):
        """主要執行流程"""
        print("🚀 TronClass 自動建立工具")
        print("=" * 50)
        
        # 0. 檢查並更新 Cookie
        if not self.check_and_update_cookie():
            print("❌ Cookie 檢查失敗，無法繼續")
            return
        
        # 1. 選擇檔案
        self.excel_file = self.select_file()
        if not self.excel_file:
            print("❌ 沒有選擇檔案")
            return
        
        print(f"✅ 已選擇檔案: {os.path.basename(self.excel_file)}")
        
        # 2. 載入數據
        if not self.load_data():
            return
        
        # 3. 選擇操作
        operation, activity_type = self.select_operation()
        print(f"✅ 已選擇操作: {operation}")
        if activity_type:
            print(f"   活動類型: {activity_type}")
        
        # 4. 檢查缺失的ID
        if not self.check_missing_ids(operation):
            return
        
        # 5. 分析操作並確認
        stats = self.analyze_operation(operation, activity_type)
        if not self.confirm_operation(stats):
            print("❌ 用戶取消操作")
            return
        
        # 6. 執行操作
        print(f"\n🔄 開始執行: {operation}")
        print("=" * 30)
        
        start_time = datetime.now()
        success_count, total_count = self.execute_operation(operation, activity_type)
        end_time = datetime.now()
        
        # 7. 顯示結果
        duration = end_time - start_time
        print(f"\n🎉 執行完成！")
        print(f"📊 成功: {success_count}/{total_count}")
        print(f"⏱️  耗時: {duration}")
        print(f"📁 結果已保存到: {self.excel_file}")
        
        if success_count < total_count:
            print(f"⚠️  有 {total_count - success_count} 個項目失敗或被略過")
        # 匯總 log 寫入
        summary = {
            'timestamp': datetime.now().isoformat(),
            'operation': operation,
            'success_count': success_count,
            'total_count': total_count,
            'skipped_count': len(self.skipped_items),
            'failed_count': len(self.failed_items),
            'skipped_items': self.skipped_items,
            'failed_items': self.failed_items
        }
        log_dir = "log"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        log_filename = f"{log_dir}/import_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(log_filename, 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        print(f"📝 匯總結果已寫入: {log_filename}")

def main():
    """主函數"""
    print("🔍 正在初始化 TronClass 自動建立工具...")
    
    try:
        creator = TronClassCreator()
        creator.run()
    except ImportError as e:
        print(f"❌ 導入模組失敗: {e}")
        print("請確認以下檔案存在且正確:")
        print("  - config.py (從 config_template.py 複製)")
        print("  - create_01_course.py")
        print("  - create_02_module.py") 
        print("  - create_03_syllabus.py")
        print("  - create_04_activity.py")
        print("  - create_05_material.py")
        print("  - tronc_login.py")
    except Exception as e:
        print(f"❌ 程序執行失敗: {e}")

if __name__ == "__main__":
    main()