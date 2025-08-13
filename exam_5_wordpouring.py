#!/usr/bin/env python3
"""
exam_5_wordpouring.py - Word資源批量上傳工具
功能：
1. 選擇 exam_04_docx_todolist 中的 Excel 文件（按時間排序，過濾暫存文件）
2. 逐行讀取Excel，創建未上傳的Word資源
3. 取得資源ID後更新Excel文件
4. 即時更新Excel文件
"""

import os
import glob
import requests
import time
import logging
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("正在安裝 pandas...")
    os.system("pip3 install pandas openpyxl")
    import pandas as pd

# 導入子模組
from create_05_material import upload_material
from config import BASE_URL, COOKIE
from tronc_login import login_and_get_cookie, update_config, setup_driver
from sub_library_creator import create_library
from sub_word_import import word_import_and_convert
from sub_identify_save import identify_and_save_questions
from sub_return_library import return_to_library_page

def convert_chinese_numbers_to_digits(text):
    """將中文數字轉換為阿拉伯數字以便自然排序"""
    if not isinstance(text, str):
        return str(text)
    
    # 中文數字對應表
    chinese_digits = {
        '零': '0', '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
        '六': '6', '七': '7', '八': '8', '九': '9', '十': '10',
        '○': '0', '壹': '1', '貳': '2', '參': '3', '肆': '4', '伍': '5',
        '陸': '6', '柒': '7', '捌': '8', '玖': '9', '拾': '10'
    }
    
    result = text
    
    # 處理複雜的十位數字模式（如：十一、二十三、九十九等）
    import re
    
    # 處理 "十X" 模式（如：十一、十二...十九）
    def replace_shi_x(match):
        x = match.group(1)
        if x in chinese_digits:
            return '1' + chinese_digits[x]
        return match.group(0)
    
    result = re.sub(r'十([一二三四五六七八九壹貳參肆伍陸柒捌玖])', replace_shi_x, result)
    
    # 處理 "X十Y" 模式（如：二十一、三十五、九十九等）
    def replace_x_shi_y(match):
        x = match.group(1)
        y = match.group(2) if match.group(2) else ''
        x_digit = chinese_digits.get(x, x)
        y_digit = chinese_digits.get(y, y) if y else '0'
        if y:
            return x_digit + y_digit
        else:
            return x_digit + '0'
    
    result = re.sub(r'([一二三四五六七八九壹貳參肆伍陸柒捌玖])十([一二三四五六七八九壹貳參肆伍陸柒捌玖]?)', replace_x_shi_y, result)
    
    # 處理單獨的 "十"
    result = result.replace('十', '10')
    
    # 處理其他單個中文數字
    for chinese, digit in chinese_digits.items():
        result = result.replace(chinese, digit)
    
    return result

class WordResourceTool:
    def __init__(self):
        self.base_url = BASE_URL
        self.current_driver = None  # 保持selenium會話
        self.cookie_string = COOKIE
        self.last_cookie_update = None  # 記錄最後一次cookie更新時間
        self.setup_logging()
    
    def setup_logging(self):
        """設置log記錄功能"""
        log_dir = 'log'
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_filename = f'wordresource_{timestamp}.log'
        log_path = os.path.join(log_dir, log_filename)
        
        self.logger = logging.getLogger('WordResourceTool')
        self.logger.setLevel(logging.DEBUG)
        self.logger.handlers.clear()
        
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        detailed_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        simple_formatter = logging.Formatter('%(levelname)s - %(message)s')
        
        file_handler.setFormatter(detailed_formatter)
        console_handler.setFormatter(simple_formatter)
        
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
        self.logger.info(f"Log檔案: {log_path}")
        self.logger.info("=== Word資源批量上傳工具開始 ===\n")
    
    def test_cookie_validity(self):
        """測試當前 Cookie 是否有效"""
        try:
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
            test_url = f"{self.base_url}/api/course"
            response = requests.get(test_url, headers=headers, cookies=cookies, verify=False, timeout=10)
            
            self.logger.info(f"Cookie測試 - 狀態碼: {response.status_code}")
            self.logger.info(f"Cookie測試 - 響應頭: {dict(response.headers)}")
            
            # 記錄響應內容（前200字符）
            response_text = response.text[:200] + "..." if len(response.text) > 200 else response.text
            self.logger.info(f"Cookie測試 - 響應內容: {response_text}")
            
            # 如果狀態碼是 200 或 201，表示 cookie 有效
            if response.status_code in [200, 201]:
                return True
            # 如果狀態碼是 401、403 或 500，表示認證失敗或服務器錯誤
            elif response.status_code in [401, 403, 500]:
                return False
            # 其他狀態碼（如404等）可能是endpoint問題，嘗試其他測試方式
            else:
                # 嘗試另一個endpoint進行測試
                return self.test_cookie_alternative_endpoint()
                
        except Exception as e:
            self.logger.error(f"Cookie測試時發生錯誤: {e}")
            return False
    
    def test_cookie_alternative_endpoint(self):
        """使用備選endpoint測試Cookie有效性"""
        try:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            
            # 將 Cookie 字串轉為字典
            cookies = dict(item.split("=", 1) for item in self.cookie_string.split("; "))
            
            headers = {
                "Accept": "application/json, text/plain, */*",
                "User-Agent": "Mozilla/5.0",
                "X-Requested-With": "XMLHttpRequest"
            }
            
            # 嘗試用戶資源endpoint
            alternative_endpoints = [
                f"{self.base_url}/api/user/resources",
                f"{self.base_url}/api/user/profile",
                f"{self.base_url}/api/user"
            ]
            
            for endpoint in alternative_endpoints:
                try:
                    self.logger.info(f"測試備選endpoint: {endpoint}")
                    response = requests.get(endpoint, headers=headers, cookies=cookies, verify=False, timeout=10)
                    
                    self.logger.info(f"備選endpoint測試 - 狀態碼: {response.status_code}")
                    
                    if response.status_code in [200, 201]:
                        self.logger.info("✅ 備選endpoint測試通過，Cookie有效")
                        return True
                    elif response.status_code in [401, 403]:
                        self.logger.info("❌ 備選endpoint測試失敗，Cookie無效")
                        return False
                        
                except requests.RequestException as e:
                    self.logger.debug(f"備選endpoint {endpoint} 測試失敗: {e}")
                    continue
            
            # 如果所有endpoint都無法確認，假設cookie無效
            self.logger.warning("所有備選endpoint測試都無法確認Cookie狀態，假設無效")
            return False
                
        except Exception as e:
            self.logger.error(f"備選endpoint Cookie測試時發生錯誤: {e}")
            return False
    
    def refresh_cookie(self):
        """自動登入並刷新Cookie"""
        print("🔄 Cookie已過期，正在自動重新登入...")
        self.logger.info("開始自動登入流程")
        
        try:
            result = login_and_get_cookie()
            
            if result:
                cookie_string, modules = result
                self.cookie_string = cookie_string
                self.last_cookie_update = datetime.now()  # 記錄更新時間
                self.logger.info("自動登入成功，Cookie已更新")
                print("✅ 自動登入成功，Cookie已更新")
                
                # 更新配置文件
                update_config(cookie_string, modules)
                
                # tronc_login.py中已經驗證了登入是否成功（通過檢查URL不包含'login'）
                # 如果能成功返回cookie和modules，就說明登入成功
                # 不再進行額外的API驗證，因為可能會因為500狀態碼誤判
                self.logger.info("✅ 登入流程成功完成，相信tronc_login的驗證結果")
                print("✅ 登入驗證成功")
                return True
            else:
                self.logger.error("自動登入失敗")
                print("❌ 自動登入失敗")
                return False
                
        except Exception as e:
            self.logger.exception(f"自動登入過程發生錯誤: {e}")
            print(f"❌ 自動登入過程發生錯誤: {e}")
            return False
    
    def check_and_refresh_cookie(self):
        """檢查Cookie並在需要時自動刷新"""
        # 如果cookie是最近更新的（5分鐘內），直接認為有效
        if self.last_cookie_update:
            from datetime import datetime, timedelta
            if datetime.now() - self.last_cookie_update < timedelta(minutes=5):
                self.logger.info("Cookie是最近更新的，直接認為有效")
                print("✅ Cookie是最近更新的，跳過驗證")
                return True
        
        # 嘗試主要endpoint測試
        self.logger.info("開始測試Cookie有效性...")
        if self.test_cookie_validity():
            self.logger.info("主要endpoint測試通過")
            return True
        
        # 如果主要endpoint失敗，嘗試備選endpoint
        self.logger.info("主要endpoint測試失敗，嘗試備選endpoint...")
        print("⚠️ 主要endpoint測試失敗，嘗試備選endpoint...")
        if self.test_cookie_alternative_endpoint():
            self.logger.info("備選endpoint驗證通過，Cookie有效")
            print("✅ 備選endpoint驗證通過，Cookie有效")
            return True
        
        # 所有API驗證都失敗，進行自動登入刷新
        # 但刷新成功後不再進行API驗證（避免500狀態碼誤判）
        self.logger.info("所有endpoint驗證失敗，嘗試自動刷新Cookie...")
        print("⚠️ Cookie可能無效，正在嘗試自動刷新...")
        return self.refresh_cookie()
    
    def select_operation(self):
        """讓用戶選擇要執行的操作"""
        operations = [
            "上傳所有資源",
            "上傳所有題庫", 
            "生成最新exam_todolist"
        ]
        
        print("\n📋 請選擇要進行的操作：")
        for i, op in enumerate(operations, 1):
            print(f"{i}. {op}")
        
        while True:
            try:
                print(f"\n請選擇操作 (1-{len(operations)}): ", end="", flush=True)
                choice = input().strip()
                if not choice:
                    print("⚠️ 請輸入有效值")
                    continue
                    
                choice_num = int(choice)
                if 1 <= choice_num <= len(operations):
                    selected_op = operations[choice_num - 1]
                    print(f"✅ 已選擇操作: {selected_op}")
                    return choice_num
                else:
                    print(f"❌ 請輸入 1-{len(operations)} 之間的數字")
            except ValueError:
                print("❌ 請輸入有效的數字")
            except KeyboardInterrupt:
                print("\n❌ 用戶取消操作")
                return None
    
    def select_excel_file(self):
        """選擇要處理的Excel文件，按時間戳排序，過濾暫存文件"""
        todolist_dir = "exam_04_docx_todolist"
        if not os.path.exists(todolist_dir):
            print(f"找不到 {todolist_dir} 資料夾")
            return None
        
        # 尋找Excel文件，過濾~開頭的暫存文件
        all_files = glob.glob(os.path.join(todolist_dir, "*.xlsx"))
        excel_files = [f for f in all_files if not os.path.basename(f).startswith('~')]
        
        if not excel_files:
            print(f"在 {todolist_dir} 中找不到任何有效的 Excel 文件")
            return None
        
        # 按修改時間排序（最新到最舊）
        excel_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        
        print(f"找到以下 Excel 文件（按時間排序）:")
        for i, file in enumerate(excel_files, 1):
            filename = os.path.basename(file)
            mtime = datetime.fromtimestamp(os.path.getmtime(file)).strftime("%Y-%m-%d %H:%M:%S")
            print(f"{i}. {filename} ({mtime})")
        
        while True:
            try:
                choice = input(f"請選擇要處理的文件 (1-{len(excel_files)}, 輸入0使用預設第一項): ").strip()
                if not choice:
                    print("取消操作")
                    return None
                
                choice_num = int(choice)
                if choice_num == 0:
                    selected_file = excel_files[0]
                    print(f"使用預設文件: {os.path.basename(selected_file)}")
                    return selected_file
                elif 1 <= choice_num <= len(excel_files):
                    selected_file = excel_files[choice_num - 1]
                    print(f"選擇的文件: {os.path.basename(selected_file)}")
                    return selected_file
                else:
                    print(f"請輸入 0-{len(excel_files)} 之間的數字")
            except ValueError:
                print("請輸入有效的數字")
            except KeyboardInterrupt:
                print("\n取消操作")
                return None
    
    def build_docx_path(self, folder_name, docx_filename):
        """構建Word文件的完整路徑"""
        base_dir = "exam_03_wordsplitter"
        if folder_name:
            docx_path = os.path.join(base_dir, folder_name, docx_filename)
        else:
            docx_path = os.path.join(base_dir, docx_filename)
        return docx_path
    
    def create_resource(self, docx_file_path, title):
        """創建資源"""
        self.logger.info(f"開始創建資源: 文件={docx_file_path}, 標題={title}")
        
        try:
            # 檢查文件是否存在
            if not os.path.exists(docx_file_path):
                self.logger.error(f"文件不存在: {docx_file_path}")
                return None, None
            
            # 檢查並刷新Cookie
            if not self.check_and_refresh_cookie():
                self.logger.error("Cookie驗證和刷新均失敗")
                print("❌ 認證失敗，請檢查登入憑證")
                return None, None
            
            # 使用真實的API上傳資源
            self.logger.info(f"正在上傳文件: {os.path.basename(docx_file_path)}")
            self.logger.info(f"文件大小: {os.path.getsize(docx_file_path)} 字節")
            
            result = upload_material(self.cookie_string, docx_file_path, parent_id=0, file_type="resource")
            
            self.logger.info(f"上傳結果: {result}")
            
            if result["success"]:
                resource_id = str(result["material_id"])
                create_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                self.logger.info(f"資源上傳成功: ID={resource_id}")
                print(f"✅ 資源創建成功: ID={resource_id}")
                return resource_id, create_time
            else:
                error_msg = result.get("error", "未知錯誤")
                step = result.get("step", "未知步驟")
                self.logger.error(f"資源上傳失敗: 步驟={step}, 錯誤={error_msg}")
                self.logger.error(f"完整錯誤信息: {result}")
                print(f"❌ 資源創建失敗: 步驟={step}, 錯誤={error_msg}")
                return None, None
                
        except Exception as e:
            self.logger.exception(f"創建資源時發生異常: {e}")
            print(f"❌ 創建資源時發生異常: {e}")
            return None, None
    
    def generate_new_todolist(self):
        """生成最新的exam_todolist"""
        print("📝 正在生成最新的docx_todolist...")
        self.logger.info("開始生成最新docx_todolist")
        
        try:
            # 掃描exam_03_wordsplitter目錄中的所有Word文件
            base_dir = Path("exam_03_wordsplitter")
            if not base_dir.exists():
                print(f"❌ 找不到 {base_dir} 目錄")
                self.logger.error(f"目錄不存在: {base_dir}")
                return False
            
            word_files_with_folder = []
            
            self.logger.info(f"開始掃描目錄: {base_dir}")
            
            # 遞歸查找所有Word文件（排除original開頭的文件）
            for root, _, files in os.walk(str(base_dir)):
                root_path = Path(root)
                self.logger.info(f"掃描資料夾: {root_path}")
                
                for file in files:
                    if (file.lower().endswith(('.docx', '.doc')) and 
                        not file.startswith('~$') and 
                        not file.startswith('original_')):
                        
                        word_file_path = root_path / file
                        
                        # 獲取相對於base_dir的資料夾名稱
                        try:
                            relative_folder = root_path.relative_to(base_dir)
                            folder_name = str(relative_folder) if str(relative_folder) != '.' else ""
                            self.logger.info(f"找到文件: {file}, 資料夾: {folder_name}")
                        except ValueError as e:
                            self.logger.warning(f"計算相對路徑失敗: {e}")
                            folder_name = ""
                        
                        word_files_with_folder.append((word_file_path, folder_name))
            
            if not word_files_with_folder:
                print("❌ 在exam_03_wordsplitter中找不到任何Word文件")
                return False
            
            print(f"✅ 找到 {len(word_files_with_folder)} 個Word文件")
            self.logger.info(f"總共找到 {len(word_files_with_folder)} 個Word文件")
            
            # 創建輸出目錄
            output_dir = Path("exam_04_docx_todolist")
            try:
                if not output_dir.exists():
                    output_dir.mkdir(parents=True, exist_ok=True)
                    print(f"✅ 已創建目錄: {output_dir}")
                    self.logger.info(f"創建輸出目錄: {output_dir}")
                else:
                    self.logger.info(f"輸出目錄已存在: {output_dir}")
            except Exception as e:
                print(f"❌ 創建輸出目錄失敗: {e}")
                self.logger.error(f"創建輸出目錄失敗: {e}")
                return False
            
            # 提取所有Word文件資訊
            data_list = []
            for word_file_path, folder_name in word_files_with_folder:
                try:
                    filename = word_file_path.name
                    title = filename.replace('.docx', '').replace('.doc', '')
                    
                    docx_info = {
                        '資料夾': folder_name,
                        'docx文件名': filename,
                        '標題名稱': title,
                        '題庫資源ID': "",  # 空白
                        '資源創建時間': "",  # 空白
                        '題庫編號': "",  # 空白
                        '題庫創建時間': "",  # 空白
                        '題庫標題': title,  # 同標題名稱
                        '標題修改時間': "",  # 空白
                        '識別完成保存時間': ""  # 空白
                    }
                    data_list.append(docx_info)
                    self.logger.info(f"處理文件成功: {filename} (資料夾: {folder_name})")
                    
                except Exception as e:
                    self.logger.error(f"處理文件時發生錯誤 {word_file_path}: {e}")
                    print(f"⚠️ 跳過文件: {word_file_path} (錯誤: {e})")
                    continue
            
            if not data_list:
                print("❌ 無法提取任何有效資料")
                return False
            
            # 創建DataFrame並進行自然排序
            try:
                import pandas as pd
                import natsort
            except ImportError:
                print("正在安裝所需套件...")
                os.system("pip3 install pandas openpyxl natsort")
                import pandas as pd
                import natsort
            
            df = pd.DataFrame(data_list)
            
            # 按標題名稱自然排序（支援中文數字）
            if not df.empty and '標題名稱' in df.columns:
                # 先轉換中文數字為阿拉伯數字
                converted_titles = df['標題名稱'].astype(str).apply(convert_chinese_numbers_to_digits)
                # 使用轉換後的標題進行自然排序
                sorted_indices = natsort.index_natsorted(converted_titles)
                df = df.iloc[sorted_indices].reset_index(drop=True)
                print("✅ 已按標題名稱進行自然排序")
            
            # 生成時間戳檔案名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_filename = f"docx_todolist_{timestamp}.xlsx"
            excel_path = output_dir / excel_filename
            
            # 儲存為Excel檔案
            try:
                df.to_excel(excel_path, index=False, engine='openpyxl')
                print(f"✅ 已生成最新todolist: {excel_path}")
                print(f"📊 包含 {len(data_list)} 筆資料")
                self.logger.info(f"成功儲存Excel文件: {excel_path}")
                
                # 顯示前5筆資料作為預覽
                print("\n📋 預覽前5筆資料:")
                print(df.head().to_string(index=False))
                
                self.logger.info(f"成功生成todolist: {excel_path}")
                return True
                
            except Exception as e:
                print(f"❌ 儲存Excel文件失敗: {e}")
                self.logger.error(f"儲存Excel文件失敗: {e}")
                return False
            
        except Exception as e:
            self.logger.exception(f"生成todolist時發生錯誤: {e}")
            print(f"❌ 生成todolist時發生錯誤: {e}")
            return False
    
    def upload_all_resources(self):
        """上傳所有資源（原有的主要功能）"""
        print("📤 開始上傳所有資源...")
        
        # 選擇Excel文件
        excel_path = self.select_excel_file()
        if not excel_path:
            return False
        
        # 處理Excel文件
        print(f"\n開始處理文件: {os.path.basename(excel_path)}")
        return self.process_excel(excel_path)
    
    def upload_all_libraries(self):
        """上傳所有題庫（第二個功能）"""
        print("📚 開始上傳所有題庫...")
        
        # 選擇Excel文件
        excel_path = self.select_excel_file()
        if not excel_path:
            return False
        
        # 第一步：先自動完成所有資源上傳
        print(f"\n🔄 步驟1: 先確保所有資源已上傳...")
        if not self.ensure_all_resources_uploaded(excel_path):
            print("❌ 資源上傳未完成，無法繼續題庫上傳")
            return False
        
        # 第二步：開始題庫上傳流程
        print(f"\n🔄 步驟2: 開始處理Word題庫上傳: {os.path.basename(excel_path)}")
        return self.process_library_excel(excel_path)
    
    def ensure_all_resources_uploaded(self, excel_path):
        """確保所有資源都已上傳，如果沒有則自動上傳"""
        try:
            # 讀取Excel文件
            df = pd.read_excel(excel_path)
            
            # 檢查題庫資源ID欄位
            if '題庫資源ID' not in df.columns:
                print("❌ Excel文件缺少'題庫資源ID'欄位")
                return False
            
            # 檢查哪些資源需要上傳（沒有資源ID的）
            need_upload = df[pd.isna(df['題庫資源ID']) | (df['題庫資源ID'] == '')]
            
            if len(need_upload) == 0:
                print("✅ 所有資源都已有ID，跳過資源上傳步驟")
                return True
            
            print(f"⚠️ 發現 {len(need_upload)} 個資源尚未上傳，開始自動上傳...")
            
            # 調用現有的資源上傳邏輯
            success = self.process_excel(excel_path)
            
            if success:
                print("✅ 所有資源上傳完成")
                return True
            else:
                print("❌ 資源上傳過程中發生錯誤")
                return False
                
        except Exception as e:
            self.logger.exception(f"檢查資源上傳狀態時發生錯誤: {e}")
            print(f"❌ 檢查資源上傳狀態時發生錯誤: {e}")
            return False
    
    def setup_driver_with_cookies(self):
        """為新建的driver設置cookies並訪問首頁"""
        try:
            # 先訪問主頁
            self.logger.info("設置瀏覽器cookies...")
            self.current_driver.get(self.base_url)
            
            # 將cookie字符串解析並添加到driver
            cookies = dict(item.split("=", 1) for item in self.cookie_string.split("; "))
            
            for name, value in cookies.items():
                try:
                    self.current_driver.add_cookie({
                        'name': name,
                        'value': value,
                        'domain': '.tronclass.com'  # 設置適當的域
                    })
                except Exception as e:
                    self.logger.debug(f"添加cookie失敗 {name}: {e}")
            
            # 重新載入頁面使cookies生效
            self.current_driver.refresh()
            time.sleep(3)
            
            self.logger.info("✅ 瀏覽器cookies設置完成")
            return True
            
        except Exception as e:
            self.logger.exception(f"設置瀏覽器cookies時發生錯誤: {e}")
            return False
    
    def perform_login(self, driver):
        """執行登入流程"""
        self.logger.info("=== 開始登入流程 ===")
        
        try:
            from config import USERNAME, PASSWORD, LOGIN_URL, SEARCH_TERM
            
            # 訪問登入頁面
            self.logger.info(f"訪問登入頁面: {LOGIN_URL}")
            driver.get(LOGIN_URL)
            
            # 等待頁面加載
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.webdriver.common.by import By
            from selenium.webdriver.common.keys import Keys
            from selenium.common.exceptions import TimeoutException
            
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script('return document.readyState') == 'complete'
            )
            
            # 搜索機構（如果有設定）
            if SEARCH_TERM:
                self.logger.info(f"搜索機構: {SEARCH_TERM}")
                try:
                    search_field = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="text"]'))
                    )
                    search_field.clear()
                    search_field.send_keys(SEARCH_TERM)
                    time.sleep(2)
                    
                    # 嘗試選擇搜索結果
                    org_selected = False
                    try:
                        results = driver.find_elements(By.CSS_SELECTOR, '.dropdown-item, .result-item, ul li')
                        for result in results:
                            if SEARCH_TERM in result.text:
                                self.logger.info(f"找到匹配結果: {result.text}")
                                result.click()
                                org_selected = True
                                break
                    except:
                        self.logger.info("未找到或無法選擇搜索結果，繼續登入")
                    
                    if org_selected:
                        self.logger.info("已選擇機構，等待頁面重新載入...")
                        time.sleep(5)
                        
                        WebDriverWait(driver, 15).until(
                            lambda d: d.execute_script('return document.readyState') == 'complete'
                        )
                        
                except:
                    self.logger.info("搜索機構失敗，直接登入")
            
            # 填寫用戶名
            self.logger.info("尋找用戶名輸入欄位...")
            username_selectors = [
                '#username', '[name="username"]', '[name="email"]', 
                'input[type="text"]', 'input[placeholder*="帳號"]',
                'input[placeholder*="用戶"]', 'input[placeholder*="使用者"]'
            ]
            
            username_field = None
            for selector in username_selectors:
                try:
                    username_field = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    self.logger.info(f"找到用戶名欄位: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if not username_field:
                self.logger.error("找不到用戶名輸入欄位")
                return False
            
            username_field.clear()
            username_field.send_keys(USERNAME)
            
            # 填寫密碼
            self.logger.info("尋找密碼輸入欄位...")
            password_selectors = [
                '#password', '[name="password"]', 'input[type="password"]',
                'input[placeholder*="密碼"]', 'input[placeholder*="密码"]'
            ]
            
            password_field = None
            for selector in password_selectors:
                try:
                    password_field = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    self.logger.info(f"找到密碼欄位: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if not password_field:
                self.logger.error("找不到密碼輸入欄位")
                return False
            
            password_field.clear()
            password_field.send_keys(PASSWORD)
            
            # 提交登入
            self.logger.info("尋找提交按鈕...")
            submit_selectors = [
                'button[type="submit"]', 'input[type="submit"]', 
                '.login-btn', '.btn-login', 'button:contains("登入")',
                'button:contains("登录")', 'button:contains("Login")',
                'button[class*="submit"]', 'button[class*="login"]'
            ]
            
            submit_button = None
            for selector in submit_selectors:
                try:
                    submit_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    self.logger.info(f"找到提交按鈕: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if submit_button:
                submit_button.click()
                self.logger.info("已點擊提交按鈕")
            else:
                self.logger.info("未找到提交按鈕，嘗試按Enter鍵")
                password_field.send_keys(Keys.RETURN)
            
            # 等待登入處理
            self.logger.info("等待登入處理...")
            time.sleep(5)
            
            current_url = driver.current_url
            self.logger.info(f"登入後URL: {current_url}")
            
            # 訪問首頁確保登入狀態
            driver.get(self.base_url)
            WebDriverWait(driver, 15).until(
                lambda d: d.execute_script('return document.readyState') == 'complete'
            )
            time.sleep(3)
            
            self.logger.info("✅ 登入流程完成")
            return True
            
        except Exception as e:
            self.logger.exception(f"登入流程發生錯誤: {e}")
            return False
    
    def process_library_excel(self, excel_path):
        """處理Excel文件，逐行執行Word題庫上傳流程"""
        try:
            # 讀取Excel文件
            df = pd.read_excel(excel_path)
            print(f"讀取到 {len(df)} 筆資料")
            
            # 確認所需欄位存在
            required_columns = ['題庫資源ID', '題庫標題', '題庫編號', '題庫創建時間']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                print(f"Excel文件缺少必要欄位: {missing_columns}")
                print(f"現有欄位: {list(df.columns)}")
                return False
            
            # 檢查並刷新Cookie
            print("🔍 檢查認證狀態...")
            if not self.check_and_refresh_cookie():
                print("❌ 認證失敗，請檢查網路連接和登入憑證")
                return False
            print("✅ 認證有效")
            
            # 初始化selenium driver
            if not self.current_driver:
                self.current_driver = setup_driver()
                if not self.current_driver:
                    print("❌ 無法啟動瀏覽器")
                    return False
                
                # 只有在新建driver時才需要設置cookies和訪問首頁
                self.setup_driver_with_cookies()
            else:
                print("✅ 使用現有瀏覽器會話")
            
            # 逐行處理
            for index, row in df.iterrows():
                print(f"\n=== 處理第 {index + 1} 行: {row.get('標題名稱', '未知')} ===")
                
                # 檢查是否有資源ID
                if pd.isna(row['題庫資源ID']) or not row['題庫資源ID']:
                    print(f"❌ 第 {index + 1} 行缺少資源ID，跳過")
                    continue
                
                upload_id = row['題庫資源ID']
                
                # 檢查是否已有題庫編號
                if pd.isna(row['題庫編號']):
                    # 創建新的測驗題庫
                    result = create_library(self.current_driver, row.get('題庫標題'), self.logger)
                    
                    if result:
                        lib_id, create_time = result
                        
                        # 即時更新Excel
                        df.at[index, '題庫編號'] = str(lib_id)  # 確保為字符串
                        df.at[index, '題庫創建時間'] = str(create_time)  # 確保為字符串
                        df.to_excel(excel_path, index=False)
                        
                        print(f"✅ 題庫創建成功: ID={lib_id}")
                        self.logger.info(f"題庫創建成功: ID={lib_id}")
                        
                        # 修改題庫標題
                        if row.get('題庫標題'):
                            print(f"🔄 修改題庫標題: {row['題庫標題']}")
                            modify_time = self.update_library_title(lib_id, row['題庫標題'])
                            if modify_time:
                                df.at[index, '標題修改時間'] = str(modify_time)  # 確保為字符串
                                df.to_excel(excel_path, index=False)
                                print(f"✅ 標題修改完成")
                            else:
                                print(f"⚠️ 標題修改失敗，但繼續流程")
                        
                    else:
                        print(f"❌ 題庫創建失敗")
                        continue
                else:
                    lib_id = str(int(float(row['題庫編號'])))  # 轉換為字符串格式的整數
                    print(f"使用現有題庫ID: {lib_id}")
                    
                    # 檢查是否需要修改標題
                    if pd.isna(row['標題修改時間']) and row.get('題庫標題'):
                        print(f"🔄 修改題庫標題: {row['題庫標題']}")
                        modify_time = self.update_library_title(lib_id, row['題庫標題'])
                        if modify_time:
                            df.at[index, '標題修改時間'] = str(modify_time)  # 確保為字符串
                            df.to_excel(excel_path, index=False)
                            print(f"✅ 標題修改完成")
                        else:
                            print(f"⚠️ 標題修改失敗，但繼續流程")
                
                # 執行Word匯入和AI轉換
                if word_import_and_convert(self.current_driver, lib_id, upload_id, self.logger):
                    # 執行識別和儲存
                    if identify_and_save_questions(self.current_driver, self.logger):
                        # 更新識別完成保存時間
                        save_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        df.at[index, '識別完成保存時間'] = str(save_time)  # 確保為字符串
                        df.to_excel(excel_path, index=False)
                        
                        # 返回題庫列表
                        if return_to_library_page(self.current_driver, self.logger):
                            print(f"✅ 第 {index + 1} 行處理完成")
                        else:
                            print(f"⚠️ 第 {index + 1} 行處理完成，但返回失敗")
                    else:
                        print(f"❌ 第 {index + 1} 行識別儲存失敗")
                else:
                    print(f"❌ 第 {index + 1} 行Word匯入失敗")
                
                # 短暫延遲
                time.sleep(2)
            
            print(f"\n🎉 處理完成！共處理 {len(df)} 筆資料")
            return True
            
        except Exception as e:
            self.logger.exception(f"處理Excel文件時發生錯誤: {e}")
            print(f"處理Excel文件時發生錯誤: {e}")
            return False
        finally:
            # 關閉selenium會話
            if self.current_driver:
                try:
                    self.current_driver.quit()
                    self.logger.info("已關閉selenium會話")
                except:
                    pass
    
    def update_library_title(self, lib_id, new_title):
        """修改題庫標題"""
        self.logger.info(f"開始修改題庫標題: ID={lib_id}, 標題={new_title}")
        
        if not self.current_driver:
            self.logger.error("沒有可用的瀏覽器會話，無法獲取認證信息")
            print("❌ 沒有可用的瀏覽器會話")
            return None
        
        url = f"{self.base_url}/api/subject-libs/{lib_id}"
        headers = {
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        
        payload = {"title": new_title}
        
        try:
            # 從selenium driver獲取cookies
            selenium_cookies = self.current_driver.get_cookies()
            self.logger.info(f"獲取到 {len(selenium_cookies)} 個cookies")
            
            # 轉換為requests格式
            cookies = {}
            for cookie in selenium_cookies:
                cookies[cookie['name']] = cookie['value']
            
            # 記錄主要的認證相關cookies
            auth_cookies = ['PHPSESSID', 'sessionid', 'csrftoken', 'auth_token', 'access_token']
            for cookie_name in auth_cookies:
                if cookie_name in cookies:
                    self.logger.info(f"找到認證cookie: {cookie_name}")
            
            # 發送API請求
            self.logger.info(f"發送PUT請求到: {url}")
            self.logger.info(f"請求payload: {payload}")
            
            resp = requests.put(
                url, 
                headers=headers, 
                json=payload, 
                cookies=cookies,
                timeout=30
            )
            
            self.logger.info(f"API回應狀態碼: {resp.status_code}")
            self.logger.info(f"API回應Headers: {dict(resp.headers)}")
            
            # 記錄回應內容（前500字符）
            response_text = resp.text
            if len(response_text) > 500:
                self.logger.info(f"API回應內容: {response_text[:500]}...")
            else:
                self.logger.info(f"API回應內容: {response_text}")
            
            # 檢查回應是否為JSON
            try:
                response_json = resp.json()
                self.logger.info(f"API回應JSON: {response_json}")
                
                if resp.status_code == 200:
                    modify_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.logger.info(f"標題修改成功: {modify_time}")
                    print(f"✅ 標題修改成功: {new_title}")
                    return modify_time
                else:
                    self.logger.error(f"API返回錯誤狀態: {resp.status_code}")
                    print(f"❌ 標題修改失敗: {resp.status_code}")
                    return None
                    
            except ValueError:
                # 不是JSON回應，可能是HTML（登入頁面）
                if "doctype html" in response_text.lower() or "<html" in response_text.lower():
                    self.logger.error("API返回HTML頁面，可能是登入頁面，認證失敗")
                    print("❌ 認證失敗，API返回登入頁面")
                    return None
                elif resp.status_code == 200:
                    # 有些API可能返回純文本成功信息
                    self.logger.info("API可能成功（非JSON回應）")
                    modify_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    print(f"✅ 標題修改可能成功: {new_title}")
                    return modify_time
                else:
                    self.logger.error(f"API返回非JSON回應: {resp.status_code}")
                    print(f"❌ 標題修改失敗: {resp.status_code}")
                    return None
        
        except Exception as e:
            self.logger.exception(f"修改標題過程發生錯誤: {e}")
            print(f"❌ 修改標題過程發生錯誤: {e}")
            return None
    
    def is_on_subject_lib_page(self, driver):
        """檢查當前是否在 /subject-lib/ 頁面"""
        try:
            current_url = driver.current_url
            return '/subject-lib/' in current_url
        except:
            return False
    
    def process_excel(self, excel_path):
        """處理Excel文件，逐行創建資源"""
        try:
            # 讀取Excel文件
            df = pd.read_excel(excel_path)
            print(f"讀取到 {len(df)} 筆資料")
            
            # 確認所需欄位存在
            required_columns = ['資料夾', 'docx文件名', '標題名稱', '題庫資源ID', '資源創建時間']
            for col in required_columns:
                if col not in df.columns:
                    print(f"Excel文件缺少必要欄位: {col}")
                    return False
            
            # 檢查是否所有資源都已經有ID
            has_resource_id = df['題庫資源ID'].notna() & (df['題庫資源ID'] != '')
            if has_resource_id.all():
                print("✅ 所有資源都已經有題庫資源ID，程序結束")
                self.logger.info("所有資源都已創建，無需處理")
                return True
            
            # 逐行處理
            for index, row in df.iterrows():
                print(f"\n=== 處理第 {index + 1} 行: {row['docx文件名']} ===")
                
                # 檢查是否已有題庫資源ID和創建時間
                if pd.isna(row['題庫資源ID']) or pd.isna(row['資源創建時間']) or row['題庫資源ID'] == '':
                    # 構建Word文件路徑
                    docx_file_path = self.build_docx_path(row['資料夾'], row['docx文件名'])
                    if not docx_file_path or not os.path.exists(docx_file_path):
                        print(f"找不到Word文件: {docx_file_path}")
                        continue
                    
                    print(f"準備創建資源: {docx_file_path}")
                    print(f"資源標題: {row['標題名稱']}")
                    
                    # 創建資源
                    try:
                        resource_id, create_time = self.create_resource(docx_file_path, row['標題名稱'])
                        
                        if resource_id:
                            # 即時更新Excel (确保数据类型兼容)
                            df.at[index, '題庫資源ID'] = str(resource_id)
                            df.at[index, '資源創建時間'] = str(create_time)
                            df.to_excel(excel_path, index=False)
                            
                            print(f"✅ 資源創建成功: ID={resource_id}, 創建時間={create_time}")
                            self.logger.info(f"Excel已更新: 資源ID={resource_id}, 創建時間={create_time}")
                            
                            # 短暫延遲
                            time.sleep(2)
                        else:
                            print(f"❌ 資源創建失敗: {row['docx文件名']}")
                            
                            # 詢問用戶是否繼續
                            try:
                                choice = input("資源創建失敗，是否繼續下一個？(y/N): ").strip().lower()
                                if choice not in ['y', 'yes', '是']:
                                    print("用戶選擇停止處理")
                                    return False
                            except KeyboardInterrupt:
                                print("\n用戶中斷操作")
                                return False
                            continue
                            
                    except Exception as e:
                        print(f"❌ 創建資源時發生錯誤: {e}")
                        self.logger.error(f"創建資源失敗: {e}")
                        
                        # 詢問用戶是否繼續
                        try:
                            choice = input("創建資源時發生錯誤，是否繼續下一個？(y/N): ").strip().lower()
                            if choice not in ['y', 'yes', '是']:
                                print("用戶選擇停止處理")
                                return False
                        except KeyboardInterrupt:
                            print("\n用戶中斷操作")
                            return False
                        continue
                else:
                    resource_id = row['題庫資源ID']
                    print(f"資源已存在，ID: {resource_id}")
            
            print(f"\n🎉 處理完成！共處理 {len(df)} 筆資料")
            return True
            
        except Exception as e:
            self.logger.exception(f"處理Excel文件時發生錯誤: {e}")
            print(f"處理Excel文件時發生錯誤: {e}")
            return False
        finally:
            # 關閉selenium會話
            if self.current_driver:
                try:
                    self.current_driver.quit()
                    self.logger.info("已關閉selenium會話")
                except:
                    pass
    
    def run(self):
        """主要執行流程"""
        print("=== Word資源批量上傳工具 ===")
        
        # 1. 選擇操作
        operation = self.select_operation()
        if operation is None:
            return
        
        # 2. 根據選擇執行不同操作
        if operation == 1:  # 上傳所有資源
            # 需要檢查Cookie認證
            print("🔍 檢查認證狀態...")
            if not self.check_and_refresh_cookie():
                print("❌ 認證失敗，請檢查網路連接和登入憑證")
                return
            print("✅ 認證有效")
            
            success = self.upload_all_resources()
            if success:
                print("\n✅ 所有資源上傳完成！")
            else:
                print("\n❌ 資源上傳過程中發生錯誤")
                
        elif operation == 2:  # 上傳所有題庫
            # 需要檢查Cookie認證
            print("🔍 檢查認證狀態...")
            if not self.check_and_refresh_cookie():
                print("❌ 認證失敗，請檢查網路連接和登入憑證")
                return
            print("✅ 認證有效")
            
            success = self.upload_all_libraries()
            if success:
                print("\n✅ 所有題庫上傳完成！")
            else:
                print("\n❌ 題庫上傳過程中發生錯誤")
                
        elif operation == 3:  # 生成最新exam_todolist
            # 生成todolist不需要網路認證
            success = self.generate_new_todolist()
            if success:
                print("\n✅ 最新todolist生成完成！")
            else:
                print("\n❌ todolist生成過程中發生錯誤")

def main():
    tool = WordResourceTool()
    tool.run()

if __name__ == "__main__":
    main()