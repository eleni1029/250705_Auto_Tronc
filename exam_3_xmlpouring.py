#!/usr/bin/env python3
"""
exam_3_xmlpouring.py - XML題庫批量上傳工具（重構版）
功能：
1. 選擇 exam_02_xml_todolist 中的 Excel 文件（按時間排序，過濾暫存文件）
2. 逐行讀取Excel，上傳未創建的XML題庫
3. 修改題庫標題
4. 即時更新Excel文件
"""

import os
import glob
import requests
import time
import logging
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("正在安裝 pandas...")
    os.system("pip3 install pandas openpyxl")
    import pandas as pd

# 導入子模組
from sub_exam_login_upload import upload_xml
from sub_exam_keep import upload_next_xml
from config import BASE_URL, COOKIE
from tronc_login import login_and_get_cookie, update_config

class XMLPouringTool:
    def __init__(self):
        self.base_url = BASE_URL
        self.current_driver = None  # 保持selenium會話
        self.cookie_string = COOKIE
        self.setup_logging()
    
    def setup_logging(self):
        """設置log記錄功能"""
        log_dir = 'log'
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_filename = f'xmlpouring_{timestamp}.log'
        log_path = os.path.join(log_dir, log_filename)
        
        self.logger = logging.getLogger('XMLPouringTool')
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
        self.logger.info("=== XML題庫批量上傳工具開始 ===\n")
    
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
            self.logger.error(f"Cookie測試時發生錯誤: {e}")
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
                self.logger.info("自動登入成功，Cookie已更新")
                print("✅ 自動登入成功，Cookie已更新")
                
                # 更新配置文件
                update_config(cookie_string, modules)
                
                # 驗證新Cookie是否有效
                if self.test_cookie_validity():
                    self.logger.info("新Cookie驗證成功")
                    print("✅ 新Cookie驗證成功")
                    return True
                else:
                    self.logger.error("新Cookie驗證失敗")
                    print("❌ 新Cookie驗證失敗")
                    return False
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
        if self.test_cookie_validity():
            return True
        
        # Cookie無效，嘗試自動刷新
        print("⚠️ Cookie無效，正在嘗試自動刷新...")
        return self.refresh_cookie()
    
    def select_excel_file(self):
        """選擇要處理的Excel文件，按時間戳排序，過濾暫存文件"""
        todolist_dir = "exam_02_xml_todolist"
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
    
    def build_xml_path(self, folder_name, xml_filename):
        """構建XML文件的完整路徑"""
        base_dir = "exam_01_02_merged_projects"
        if folder_name:
            xml_path = os.path.join(base_dir, folder_name, xml_filename)
        else:
            xml_path = os.path.join(base_dir, xml_filename)
        return xml_path
    
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
        """處理Excel文件，逐行上傳和修改標題"""
        try:
            # 讀取Excel文件
            df = pd.read_excel(excel_path)
            print(f"讀取到 {len(df)} 筆資料")
            
            # 確認所需欄位存在
            required_columns = ['資料夾', 'xml文件名', '題庫標題', '題庫編號', '題庫創建時間', '標題修改時間']
            for col in required_columns:
                if col not in df.columns:
                    print(f"Excel文件缺少必要欄位: {col}")
                    return False
            
            # 逐行處理
            for index, row in df.iterrows():
                print(f"\n=== 處理第 {index + 1} 行: {row['xml文件名']} ===")
                
                # 檢查是否已有題庫編號和創建時間
                if pd.isna(row['題庫編號']) or pd.isna(row['題庫創建時間']):
                    # 構建XML文件路徑
                    xml_file_path = self.build_xml_path(row['資料夾'], row['xml文件名'])
                    if not xml_file_path or not os.path.exists(xml_file_path):
                        print(f"找不到XML文件: {xml_file_path}")
                        continue
                    
                    print(f"準備上傳: {xml_file_path}")
                    
                    # 檢查是否有現有的selenium會話
                    lib_id = None
                    if self.current_driver and self.is_on_subject_lib_page(self.current_driver):
                        self.logger.info("使用現有selenium會話繼續上傳...")
                        try:
                            lib_id = upload_next_xml(self.current_driver, xml_file_path)
                        except Exception as e:
                            self.logger.error(f"使用現有會話上傳失敗: {e}")
                            lib_id = None
                    
                    # 如果沒有現有會話或上傳失敗，使用完整登入流程
                    if not lib_id:
                        self.logger.info("使用完整登入流程上傳...")
                        lib_id, new_driver = upload_xml(xml_file_path)
                        if new_driver:
                            # 如果有舊的driver，先關閉
                            if self.current_driver:
                                try:
                                    self.current_driver.quit()
                                except:
                                    pass
                            self.current_driver = new_driver
                    
                    if lib_id:
                        create_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        # 即時更新Excel
                        df.at[index, '題庫編號'] = lib_id
                        df.at[index, '題庫創建時間'] = create_time
                        df.to_excel(excel_path, index=False)
                        
                        print(f"✅ 上傳成功: ID={lib_id}, 創建時間={create_time}")
                        self.logger.info(f"Excel已更新: ID={lib_id}, 創建時間={create_time}")
                        
                        # 短暫延遲
                        time.sleep(2)
                    else:
                        print(f"❌ 上傳失敗: {row['xml文件名']}")
                        
                        # 詢問用戶是否繼續
                        try:
                            choice = input("上傳失敗，是否繼續下一個？(y/N): ").strip().lower()
                            if choice not in ['y', 'yes', '是']:
                                print("用戶選擇停止處理")
                                return False
                        except KeyboardInterrupt:
                            print("\n用戶中斷操作")
                            return False
                        continue
                else:
                    lib_id = row['題庫編號']
                    print(f"題庫已存在，ID: {lib_id}")
                
                # 檢查是否需要修改標題
                if pd.isna(row['標題修改時間']) and not pd.isna(row['題庫標題']) and lib_id:
                    print(f"準備修改標題: {row['題庫標題']}")
                    
                    modify_time = self.update_library_title(lib_id, row['題庫標題'])
                    if modify_time:
                        # 即時更新Excel
                        df.at[index, '標題修改時間'] = modify_time
                        df.to_excel(excel_path, index=False)
                        print(f"✅ 標題修改成功: {modify_time}")
                        self.logger.info(f"Excel已更新: 標題修改時間={modify_time}")
                        
                        # 短暫延遲
                        time.sleep(1)
                    else:
                        print("❌ 標題修改失敗，但繼續處理下一個")
                else:
                    print("標題已修改或無標題內容，跳過")
            
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
        print("=== XML題庫批量上傳工具 ===")
        
        # 0. 檢查並刷新Cookie
        print("🔍 檢查認證狀態...")
        if not self.check_and_refresh_cookie():
            print("❌ 認證失敗，請檢查網路連接和登入憑證")
            return
        print("✅ 認證有效")
        
        # 1. 選擇Excel文件
        excel_path = self.select_excel_file()
        if not excel_path:
            return
        
        # 2. 處理Excel文件
        print(f"\n開始處理文件: {os.path.basename(excel_path)}")
        success = self.process_excel(excel_path)
        
        if success:
            print("\n✅ 所有操作完成！")
        else:
            print("\n❌ 處理過程中發生錯誤")

def main():
    tool = XMLPouringTool()
    tool.run()

if __name__ == "__main__":
    main()