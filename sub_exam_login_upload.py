#!/usr/bin/env python3
"""
sub_exam_login_upload.py - 子模組：登入並上傳XML檔案
執行完整的登入和上傳流程，返回題庫ID
"""

import os
import time
import logging
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# 導入配置和登入功能
from config import BASE_URL
from tronc_login import login_and_get_cookie, setup_driver

class SeleniumLoginUpload:
    def __init__(self):
        self.cookies = None
        self.base_url = BASE_URL
        self.setup_logging()
    
    def setup_logging(self):
        """設置log記錄功能"""
        log_dir = 'log'
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_filename = f'selenium_test_{timestamp}.log'
        log_path = os.path.join(log_dir, log_filename)
        
        self.logger = logging.getLogger('SeleniumTest')
        self.logger.setLevel(logging.DEBUG)
        self.logger.handlers.clear()
        
        # 文件handler
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # 控制台handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
        self.logger.info(f"Log檔案: {log_path}")
        self.logger.info("=== Selenium登入上傳流程開始 ===\n")
    
    def parse_cookie_string(self, cookie_str):
        """解析cookie字符串為cookie字典"""
        cookies = {}
        for item in cookie_str.split('; '):
            if '=' in item:
                key, value = item.split('=', 1)
                cookies[key] = value
        return cookies
    
    def save_page_info(self, driver, step_name):
        """保存頁面資訊到log"""
        try:
            current_url = driver.current_url
            page_title = driver.title
            page_source = driver.page_source
            
            self.logger.info(f"=== {step_name} 頁面資訊 ===")
            self.logger.info(f"URL: {current_url}")
            self.logger.info(f"標題: {page_title}")
            self.logger.debug(f"頁面源碼長度: {len(page_source)} 字元")
            
            # 尋找相關元素並記錄
            self.log_relevant_elements(driver, step_name)
            
        except Exception as e:
            self.logger.error(f"保存{step_name}頁面資訊時出錯: {e}")
    
    def log_relevant_elements(self, driver, step_name):
        """記錄頁面中相關的元素"""
        try:
            # 根據步驟尋找相關元素
            if "主頁" in step_name:
                # 尋找導航相關元素
                nav_elements = driver.find_elements(By.XPATH, "//nav//a | //div[@class='nav']//a | //ul[@class='nav']//a")
                self.logger.info(f"找到 {len(nav_elements)} 個導航元素:")
                for i, elem in enumerate(nav_elements[:10]):  # 只記錄前10個
                    try:
                        text = elem.text.strip()
                        href = elem.get_attribute('href')
                        self.logger.info(f"  {i+1}. 文字: '{text}', 連結: '{href}'")
                    except:
                        pass
            
            elif "資源" in step_name:
                # 尋找資源相關元素
                resource_elements = driver.find_elements(By.XPATH, "//*[contains(text(), '資源') or contains(text(), '題庫')]")
                self.logger.info(f"找到 {len(resource_elements)} 個資源相關元素:")
                for i, elem in enumerate(resource_elements[:10]):
                    try:
                        text = elem.text.strip()
                        tag = elem.tag_name
                        self.logger.info(f"  {i+1}. 標籤: {tag}, 文字: '{text}'")
                    except:
                        pass
            
            elif "新增" in step_name:
                # 尋找按鈕和上傳相關元素
                button_elements = driver.find_elements(By.XPATH, "//button | //a[@class*='button'] | //span[contains(text(), '新增')]")
                self.logger.info(f"找到 {len(button_elements)} 個按鈕元素:")
                for i, elem in enumerate(button_elements[:15]):
                    try:
                        text = elem.text.strip()
                        classes = elem.get_attribute('class')
                        self.logger.info(f"  {i+1}. 文字: '{text}', 類別: '{classes}'")
                    except:
                        pass
                
                # 尋找檔案上傳相關元素
                upload_elements = driver.find_elements(By.XPATH, "//input[@type='file'] | //div[@ngf-drop] | //*[contains(@class, 'upload')]")
                self.logger.info(f"找到 {len(upload_elements)} 個上傳相關元素:")
                for i, elem in enumerate(upload_elements):
                    try:
                        tag = elem.tag_name
                        classes = elem.get_attribute('class')
                        ngf_drop = elem.get_attribute('ngf-drop')
                        self.logger.info(f"  {i+1}. 標籤: {tag}, 類別: '{classes}', ngf-drop: '{ngf_drop}'")
                    except:
                        pass
        
        except Exception as e:
            self.logger.error(f"記錄相關元素時出錯: {e}")
    
    def upload_xml_file(self, xml_file_path):
        """上傳XML檔案並返回題庫ID"""
        self.logger.info(f"=== 開始上傳XML檔案: {os.path.basename(xml_file_path)} ===")
        
        if not os.path.exists(xml_file_path):
            self.logger.error(f"XML檔案不存在: {xml_file_path}")
            return None
        
        driver = None
        try:
            # === 步驟1-3: 設置瀏覽器並完成登入 ===
            self.logger.info("步驟1: 設置瀏覽器...")
            driver = setup_driver()
            if not driver:
                self.logger.error("無法啟動瀏覽器")
                return False
            
            self.logger.info("步驟2: 訪問首頁並登入...")
            # 直接使用tronc_login的邏輯，但在我們的瀏覽器中執行
            from config import USERNAME, PASSWORD, LOGIN_URL, SEARCH_TERM
            
            # 訪問登入頁面
            self.logger.info(f"訪問登入頁面: {LOGIN_URL}")
            driver.get(LOGIN_URL)
            
            # 等待頁面加載
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script('return document.readyState') == 'complete'
            )
            
            self.save_page_info(driver, "登入頁面")
            
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
                    
                    # 如果選擇了機構，需要等待頁面重新載入
                    if org_selected:
                        self.logger.info("已選擇機構，等待頁面重新載入...")
                        time.sleep(5)  # 等待頁面跳轉
                        
                        # 等待新頁面完全載入
                        WebDriverWait(driver, 15).until(
                            lambda d: d.execute_script('return document.readyState') == 'complete'
                        )
                        
                        current_url = driver.current_url
                        self.logger.info(f"機構選擇後URL: {current_url}")
                        self.save_page_info(driver, "選擇機構後")
                        
                except:
                    self.logger.info("搜索機構失敗，直接登入")
            
            # 填寫用戶名 - 增加更多等待時間和選擇器
            self.logger.info("尋找用戶名輸入欄位...")
            username_field = None
            username_selectors = [
                '#username', 
                '[name="username"]', 
                '[name="email"]', 
                'input[type="text"]',
                'input[placeholder*="帳號"]',
                'input[placeholder*="用戶"]',
                'input[placeholder*="使用者"]'
            ]
            
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
                self.save_page_info(driver, "找不到用戶名欄位")
                return None, driver
            username_field.clear()
            username_field.send_keys(USERNAME)
            
            # 填寫密碼 - 使用相同的改進邏輯
            self.logger.info("尋找密碼輸入欄位...")
            password_field = None
            password_selectors = [
                '#password',
                '[name="password"]', 
                'input[type="password"]',
                'input[placeholder*="密碼"]',
                'input[placeholder*="密码"]'
            ]
            
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
                self.save_page_info(driver, "找不到密碼欄位")
                return None, driver
            
            password_field.clear()
            password_field.send_keys(PASSWORD)
            
            # 提交登入 - 改進提交按鈕尋找邏輯
            self.logger.info("尋找提交按鈕...")
            submit_button = None
            submit_selectors = [
                'button[type="submit"]', 
                'input[type="submit"]', 
                '.login-btn',
                '.btn-login',
                'button:contains("登入")',
                'button:contains("登录")',
                'button:contains("Login")',
                'button[class*="submit"]',
                'button[class*="login"]'
            ]
            
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
            
            # 檢查是否成功登入（不在登入頁面）
            if 'login' in current_url.lower():
                self.logger.warning("可能仍在登入頁面，但繼續流程...")
            
            self.save_page_info(driver, "登入後")
            
            # === 步驟3: 訪問首頁確保登入狀態 ===
            self.logger.info("步驟3: 訪問首頁確保登入狀態...")
            driver.get(self.base_url)
            WebDriverWait(driver, 15).until(
                lambda d: d.execute_script('return document.readyState') == 'complete'
            )
            time.sleep(3)
            
            self.save_page_info(driver, "首頁載入")
            
            # === 步驟4-5: 直接訪問個人題庫 ===
            self.logger.info("步驟4-5: 直接訪問個人題庫...")
            
            # 直接訪問個人題庫URL
            personal_lib_url = f"{self.base_url}/user/resources/subject-libs#/?parent_id=0&pageIndex=1"
            self.logger.info(f"直接訪問個人題庫URL: {personal_lib_url}")
            driver.get(personal_lib_url)
            WebDriverWait(driver, 15).until(
                lambda d: d.execute_script('return document.readyState') == 'complete'
            )
            time.sleep(3)
            
            self.save_page_info(driver, "個人題庫頁面")
            
            # === 步驟6: hover "新增" 按鈕，點擊 "上傳智慧大師題庫" ===
            self.logger.info("步驟6: 尋找'新增'按鈕並hover...")
            
            new_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='新增'] | //a[contains(text(), '新增')] | //button[contains(text(), '新增')]"))
            )
            
            # 使用 ActionChains hover
            actions = ActionChains(driver)
            actions.move_to_element(new_button).perform()
            self.logger.info("已hover'新增'按鈕")
            time.sleep(2)
            
            self.save_page_info(driver, "hover新增按鈕後")
            
            # 點擊 "上傳智慧大師題庫"
            self.logger.info("尋找'上傳智慧大師題庫'選項...")
            upload_option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='上傳智慧大師題庫'] | //a[contains(text(), '上傳智慧大師題庫')]"))
            )
            upload_option.click()
            self.logger.info("已點擊'上傳智慧大師題庫'")
            time.sleep(3)
            
            self.save_page_info(driver, "點擊上傳智慧大師題庫後")
            
            # === 步驟7: 找到有ngf-drop屬性的div，選擇檔案 ===
            self.logger.info("步驟7: 尋找ngf-drop區域...")
            
            # 先記錄所有可能的上傳相關元素
            all_upload_elements = driver.find_elements(By.XPATH, "//*[@ngf-drop or @type='file' or contains(@class, 'upload')]")
            self.logger.info(f"找到 {len(all_upload_elements)} 個上傳相關元素:")
            for i, elem in enumerate(all_upload_elements):
                try:
                    tag = elem.tag_name
                    classes = elem.get_attribute('class')
                    ngf_drop = elem.get_attribute('ngf-drop')
                    elem_type = elem.get_attribute('type')
                    self.logger.info(f"  元素{i+1}: 標籤={tag}, 類別='{classes}', ngf-drop='{ngf_drop}', type='{elem_type}'")
                except Exception as e:
                    self.logger.error(f"  元素{i+1}: 讀取失敗 - {e}")
            
            # 嘗試找到指定的ngf-drop元素
            ngf_drop_element = None
            try:
                ngf_drop_element = driver.find_element(By.XPATH, "//div[@ngf-drop='onFileSelect($files, maxUploadSize)']")
                self.logger.info("找到指定的ngf-drop元素")
            except NoSuchElementException:
                self.logger.warning("找不到指定的ngf-drop元素，嘗試其他方法...")
            
            # 如果找不到，嘗試其他方法
            if not ngf_drop_element:
                try:
                    ngf_drop_element = driver.find_element(By.XPATH, "//div[@ngf-drop]")
                    self.logger.info("找到通用ngf-drop元素")
                except NoSuchElementException:
                    self.logger.error("找不到任何ngf-drop元素")
            
            # 嘗試找到檔案輸入元素
            file_input = None
            file_selectors = [
                "input[type='file']",
                "//input[@type='file']",
                "//div[@ngf-drop]//input[@type='file']",
                "//input[@ngf-select]"
            ]
            
            for selector in file_selectors:
                try:
                    if selector.startswith('//'):
                        elements = driver.find_elements(By.XPATH, selector)
                    else:
                        elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    
                    for element in elements:
                        if element.get_attribute('type') == 'file':
                            file_input = element
                            self.logger.info(f"找到檔案輸入元素: {selector}")
                            break
                    
                    if file_input:
                        break
                except Exception as e:
                    self.logger.debug(f"尋找檔案輸入元素失敗 {selector}: {e}")
            
            # 直接使用XML檔案，不需要壓縮
            xml_path = os.path.abspath(xml_file_path)
            self.logger.info(f"準備上傳XML檔案: {xml_path}")
            
            # 嘗試上傳檔案
            if file_input:
                self.logger.info(f"正在上傳XML檔案: {xml_path}")
                file_input.send_keys(xml_path)
                self.logger.info("XML檔案上傳指令已發送")
                time.sleep(5)
            elif ngf_drop_element:
                self.logger.info("嘗試點擊ngf-drop元素...")
                ngf_drop_element.click()
                time.sleep(2)
                
                # 再次尋找檔案輸入元素
                for selector in file_selectors:
                    try:
                        if selector.startswith('//'):
                            file_input = driver.find_element(By.XPATH, selector)
                        else:
                            file_input = driver.find_element(By.CSS_SELECTOR, selector)
                        
                        if file_input:
                            self.logger.info(f"點擊後找到檔案輸入元素: {selector}")
                            file_input.send_keys(xml_path)
                            self.logger.info("XML檔案上傳指令已發送")
                            time.sleep(5)
                            break
                    except NoSuchElementException:
                        continue
            else:
                self.logger.error("無法找到檔案上傳方式")
                return False
            
            self.save_page_info(driver, "檔案上傳後")
            
            # === 步驟8: 點擊"確定匯入"按鈕 ===
            self.logger.info("步驟8: 等待並尋找'確定匯入'按鈕...")
            
            # 等待上傳成功
            confirm_button = None
            confirm_selectors = [
                "//a[@ng-if='ui.uploadSucceeded' and @ng-click='uploadXml()']",
                "//a[@ng-click='uploadXml()' and contains(@class, 'button-green')]",
                "//button[@ng-click='uploadXml()']",
                "//a[contains(text(), '確定匯入')]",
                "//button[contains(text(), '確定匯入')]"
            ]
            
            # 等待確定匯入按鈕出現
            for selector in confirm_selectors:
                try:
                    confirm_button = WebDriverWait(driver, 30).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    self.logger.info(f"找到'確定匯入'按鈕: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if not confirm_button:
                self.save_page_info(driver, "找不到確定匯入按鈕")
                self.logger.error("找不到'確定匯入'按鈕")
                return False
            
            self.logger.info("點擊'確定匯入'按鈕...")
            confirm_button.click()
            time.sleep(3)
            
            # === 步驟9: 等待頁面跳轉或成功指示 ===
            self.logger.info("步驟9: 等待頁面跳轉或成功指示...")
            
            # 記錄當前URL
            current_url_before = driver.current_url
            self.logger.info(f"點擊確定匯入前URL: {current_url_before}")
            
            # 等待多種成功指示：URL變化、成功信息、或特定元素
            start_time = time.time()
            timeout = 30  # 降低超時時間
            final_url = current_url_before
            success_detected = False
            
            while time.time() - start_time < timeout:
                current_url = driver.current_url
                
                # 方法1: 檢查URL是否包含subject-lib並且有ID
                import re
                if re.search(r'/subject-lib/\d+', current_url):
                    self.logger.info(f"檢測到題庫URL格式: {current_url}")
                    final_url = current_url
                    success_detected = True
                    break
                
                # 方法2: 檢查是否有成功信息
                try:
                    success_indicators = [
                        "//*[contains(text(), '匯入成功')]",
                        "//*[contains(text(), '上傳成功')]", 
                        "//*[contains(text(), '建立成功')]",
                        "//*[contains(text(), '完成')]",
                        "//div[@class='alert-success']",
                        "//div[contains(@class, 'success')]"
                    ]
                    
                    for indicator in success_indicators:
                        success_elements = driver.find_elements(By.XPATH, indicator)
                        if success_elements:
                            self.logger.info(f"找到成功指示: {indicator}")
                            for elem in success_elements:
                                try:
                                    self.logger.info(f"成功信息: {elem.text}")
                                except:
                                    pass
                            success_detected = True
                            final_url = current_url
                            break
                    
                    if success_detected:
                        break
                except:
                    pass
                
                # 方法3: 檢查URL變化（但不要求特定格式）
                if current_url != current_url_before:
                    self.logger.info(f"檢測到URL變化: {current_url}")
                    final_url = current_url
                    # 再等待2秒看是否還會繼續跳轉
                    time.sleep(2)
                    final_check_url = driver.current_url
                    if final_check_url != current_url:
                        final_url = final_check_url
                        self.logger.info(f"URL繼續變化到: {final_url}")
                    success_detected = True
                    break
                
                self.logger.debug(f"等待跳轉中... 當前URL: {current_url}")
                time.sleep(1)
            
            if not success_detected:
                # 超時後最終檢查
                final_url = driver.current_url
                self.logger.warning(f"等待超時，最終URL: {final_url}")
                self.save_page_info(driver, "等待超時後")
                
                # 檢查頁面是否已經在正確位置
                if re.search(r'/subject-lib/\d+', final_url):
                    self.logger.info("雖然等待超時，但URL格式正確，可能已成功")
                    success_detected = True
            
            if success_detected:
                self.logger.info(f"✓ 成功檢測到頁面跳轉或成功指示 - 最終URL: {final_url}")
            else:
                self.logger.warning(f"⚠ 未明確檢測到成功指示，但繼續嘗試提取ID - 最終URL: {final_url}")
            
            # === 步驟10: 從URL提取題庫ID ===
            self.logger.info("步驟10: 從URL提取題庫ID...")
            
            import re
            lib_id = None
            
            # 嘗試多種URL模式來提取ID
            patterns = [
                r'/subject-lib/(\d+)/edit',      # 編輯頁面格式
                r'/subject-lib/(\d+)/',          # 一般題庫頁面格式
                r'/subject-lib/(\d+)$',          # 結尾是ID的格式
                r'lib_id=(\d+)',                 # 參數格式
                r'id=(\d+)'                      # 通用ID參數格式
            ]
            
            for pattern in patterns:
                match = re.search(pattern, final_url)
                if match:
                    lib_id = match.group(1)
                    self.logger.info(f"✅ 使用模式 '{pattern}' 提取到題庫ID: {lib_id}")
                    break
            
            if lib_id:
                self.logger.info(f"✅ 上傳成功! 題庫ID: {lib_id}")
                return lib_id, driver  # 返回ID和driver
            else:
                self.logger.error(f"無法從URL提取ID: {final_url}")
                
                # 嘗試其他方法：檢查頁面元素中是否有ID信息
                try:
                    id_elements = driver.find_elements(By.XPATH, "//*[contains(@href, 'subject-lib/') or contains(@data-id, '') or contains(@value, '')]")
                    for elem in id_elements:
                        try:
                            href = elem.get_attribute('href')
                            if href:
                                for pattern in patterns:
                                    match = re.search(pattern, href)
                                    if match:
                                        lib_id = match.group(1)
                                        self.logger.info(f"✅ 從頁面元素提取到題庫ID: {lib_id}")
                                        return lib_id, driver
                        except:
                            continue
                except:
                    pass
                
                self.logger.warning("無法提取題庫ID，但可能上傳仍然成功")
                return None, driver
        
        except TimeoutException as e:
            self.logger.error(f"Selenium操作超時: {e}")
            if driver:
                self.save_page_info(driver, "操作超時時")
            return None, None
        except Exception as e:
            self.logger.exception(f"上傳流程發生錯誤: {e}")
            if driver:
                self.save_page_info(driver, "發生錯誤時")
            return None, None
        # 注意：不在finally中關閉driver，讓調用方決定是否關閉

def upload_xml(xml_file_path):
    """外部調用接口：上傳XML檔案並返回(題庫ID, driver)"""
    uploader = SeleniumLoginUpload()
    return uploader.upload_xml_file(xml_file_path)

def main():
    """測試用主函數"""
    test_xml = "exam_01_02_merged_projects/6-1 遊程規劃/6-1 遊程規劃_q08.xml"
    lib_id, driver = upload_xml(test_xml)
    
    if lib_id:
        print(f"✅ 上傳成功，題庫ID: {lib_id}")
    else:
        print("❌ 上傳失敗")
    
    # 關閉瀏覽器
    if driver:
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    main()