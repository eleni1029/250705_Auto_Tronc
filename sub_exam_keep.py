#!/usr/bin/env python3
"""
sub_exam_keep.py - 子模組：在現有頁面繼續上傳
當已經在 /subject-lib/ 頁面時，返回題庫列表並繼續上傳下一個檔案
"""

import os
import time
import logging
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class SeleniumKeepUpload:
    def __init__(self, driver):
        self.driver = driver
        self.setup_logging()
    
    def setup_logging(self):
        """設置log記錄功能"""
        log_dir = 'log'
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_filename = f'selenium_keep_{timestamp}.log'
        log_path = os.path.join(log_dir, log_filename)
        
        self.logger = logging.getLogger('SeleniumKeep')
        self.logger.setLevel(logging.DEBUG)
        self.logger.handlers.clear()
        
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
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
        self.logger.info("=== Selenium繼續上傳流程開始 ===\n")
    
    def upload_next_xml(self, xml_file_path):
        """在現有頁面繼續上傳XML檔案並返回題庫ID"""
        self.logger.info(f"=== 繼續上傳XML檔案: {os.path.basename(xml_file_path)} ===")
        
        if not os.path.exists(xml_file_path):
            self.logger.error(f"XML檔案不存在: {xml_file_path}")
            return None
        
        try:
            # === 步驟1: 點擊"返回題庫" ===
            self.logger.info("步驟1: 尋找並點擊'返回題庫'...")
            
            back_selectors = [
                "//a[contains(text(), '返回題庫')]",
                "//button[contains(text(), '返回題庫')]",
                "//span[contains(text(), '返回題庫')]",
                "//a[contains(@href, 'subject-libs')]"
            ]
            
            back_button = None
            for selector in back_selectors:
                try:
                    back_button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    self.logger.info(f"找到'返回題庫'按鈕: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if not back_button:
                self.logger.error("找不到'返回題庫'按鈕")
                return None
            
            back_button.click()
            self.logger.info("已點擊'返回題庫'")
            time.sleep(3)
            
            # === 步驟2: 點擊"返回上一層" ===
            self.logger.info("步驟2: 尋找並點擊'返回上一層'...")
            
            back_up_selectors = [
                "//a[contains(text(), '返回上一層')]",
                "//button[contains(text(), '返回上一層')]",
                "//span[contains(text(), '返回上一層')]",
                "//a[contains(text(), '上一層')]"
            ]
            
            back_up_button = None
            for selector in back_up_selectors:
                try:
                    back_up_button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    self.logger.info(f"找到'返回上一層'按鈕: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if back_up_button:
                back_up_button.click()
                self.logger.info("已點擊'返回上一層'")
                time.sleep(3)
            else:
                self.logger.info("未找到'返回上一層'按鈕，可能已在頂層")
            
            # === 步驟3: hover "新增" 按鈕，點擊 "上傳智慧大師題庫" ===
            self.logger.info("步驟3: 尋找'新增'按鈕並hover...")
            
            new_button = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='新增'] | //a[contains(text(), '新增')] | //button[contains(text(), '新增')]"))
            )
            
            # 使用 ActionChains hover
            actions = ActionChains(self.driver)
            actions.move_to_element(new_button).perform()
            self.logger.info("已hover'新增'按鈕")
            time.sleep(2)
            
            # 點擊 "上傳智慧大師題庫"
            self.logger.info("尋找'上傳智慧大師題庫'選項...")
            upload_option = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='上傳智慧大師題庫'] | //a[contains(text(), '上傳智慧大師題庫')]"))
            )
            upload_option.click()
            self.logger.info("已點擊'上傳智慧大師題庫'")
            time.sleep(3)
            
            # === 步驟4: 找到有ngf-drop屬性的div，選擇檔案 ===
            self.logger.info("步驟4: 尋找ngf-drop區域...")
            
            # 嘗試找到指定的ngf-drop元素
            ngf_drop_element = None
            try:
                ngf_drop_element = self.driver.find_element(By.XPATH, "//div[@ngf-drop='onFileSelect($files, maxUploadSize)']")
                self.logger.info("找到指定的ngf-drop元素")
            except NoSuchElementException:
                self.logger.warning("找不到指定的ngf-drop元素，嘗試其他方法...")
            
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
                        elements = self.driver.find_elements(By.XPATH, selector)
                    else:
                        elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    
                    for element in elements:
                        if element.get_attribute('type') == 'file':
                            file_input = element
                            self.logger.info(f"找到檔案輸入元素: {selector}")
                            break
                    
                    if file_input:
                        break
                except Exception as e:
                    self.logger.debug(f"尋找檔案輸入元素失敗 {selector}: {e}")
            
            # 直接使用XML檔案
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
                            file_input = self.driver.find_element(By.XPATH, selector)
                        else:
                            file_input = self.driver.find_element(By.CSS_SELECTOR, selector)
                        
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
                return None
            
            # === 步驟5: 點擊"確定匯入"按鈕 ===
            self.logger.info("步驟5: 等待並尋找'確定匯入'按鈕...")
            
            confirm_selectors = [
                "//a[@ng-if='ui.uploadSucceeded' and @ng-click='uploadXml()']",
                "//button[@ng-click='uploadXml()']",
                "//a[contains(text(), '確定匯入')]",
                "//button[contains(text(), '確定匯入')]"
            ]
            
            confirm_button = None
            for selector in confirm_selectors:
                try:
                    confirm_button = WebDriverWait(self.driver, 30).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    self.logger.info(f"找到'確定匯入'按鈕: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if not confirm_button:
                self.logger.error("找不到'確定匯入'按鈕")
                return None
            
            self.logger.info("點擊'確定匯入'按鈕...")
            confirm_button.click()
            time.sleep(3)
            
            # === 步驟6: 等待頁面跳轉或成功指示 ===
            self.logger.info("步驟6: 等待頁面跳轉或成功指示...")
            
            # 記錄當前URL
            current_url_before = self.driver.current_url
            self.logger.info(f"點擊確定匯入前URL: {current_url_before}")
            
            # 等待多種成功指示：URL變化、成功信息、或特定元素
            start_time = time.time()
            timeout = 30  # 降低超時時間
            final_url = current_url_before
            success_detected = False
            
            while time.time() - start_time < timeout:
                current_url = self.driver.current_url
                
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
                        success_elements = self.driver.find_elements(By.XPATH, indicator)
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
                    final_check_url = self.driver.current_url
                    if final_check_url != current_url:
                        final_url = final_check_url
                        self.logger.info(f"URL繼續變化到: {final_url}")
                    success_detected = True
                    break
                
                self.logger.debug(f"等待跳轉中... 當前URL: {current_url}")
                time.sleep(1)
            
            if not success_detected:
                # 超時後最終檢查
                final_url = self.driver.current_url
                self.logger.warning(f"等待超時，最終URL: {final_url}")
                
                # 檢查頁面是否已經在正確位置
                if re.search(r'/subject-lib/\d+', final_url):
                    self.logger.info("雖然等待超時，但URL格式正確，可能已成功")
                    success_detected = True
            
            if success_detected:
                self.logger.info(f"✓ 成功檢測到頁面跳轉或成功指示 - 最終URL: {final_url}")
            else:
                self.logger.warning(f"⚠ 未明確檢測到成功指示，但繼續嘗試提取ID - 最終URL: {final_url}")
            
            # === 步驟7: 從URL提取題庫ID ===
            self.logger.info("步驟7: 從URL提取題庫ID...")
            
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
                self.logger.info(f"✅ 繼續上傳成功! 題庫ID: {lib_id}")
                return lib_id
            else:
                self.logger.error(f"無法從URL提取ID: {final_url}")
                
                # 嘗試其他方法：檢查頁面元素中是否有ID信息
                try:
                    id_elements = self.driver.find_elements(By.XPATH, "//*[contains(@href, 'subject-lib/') or contains(@data-id, '') or contains(@value, '')]")
                    for elem in id_elements:
                        try:
                            href = elem.get_attribute('href')
                            if href:
                                for pattern in patterns:
                                    match = re.search(pattern, href)
                                    if match:
                                        lib_id = match.group(1)
                                        self.logger.info(f"✅ 從頁面元素提取到題庫ID: {lib_id}")
                                        return lib_id
                        except:
                            continue
                except:
                    pass
                
                self.logger.warning("無法提取題庫ID，但可能上傳仍然成功")
                return None
        
        except TimeoutException as e:
            self.logger.error(f"Selenium操作超時: {e}")
            return None
        except Exception as e:
            self.logger.exception(f"繼續上傳流程發生錯誤: {e}")
            return None

def upload_next_xml(driver, xml_file_path):
    """外部調用接口：在現有頁面繼續上傳XML檔案"""
    keeper = SeleniumKeepUpload(driver)
    return keeper.upload_next_xml(xml_file_path)