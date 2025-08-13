#!/usr/bin/env python3
"""
sub_word_import.py - Word匯入和AI轉換子模組
功能：
1. hover 匯入試題，點擊 Word 匯入
2. 確認頁面跳轉到 /subject-lib/{id}/import?mode=word
3. 調用AI轉換API並等待完成
4. 等待解析完成的DOM元素出現
"""

import time
import logging
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class WordImporter:
    def __init__(self, logger=None):
        self.logger = logger or self._setup_default_logger()
    
    def _setup_default_logger(self):
        """設置默認logger"""
        logger = logging.getLogger('WordImporter')
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        return logger
    
    def navigate_to_library_page(self, driver, lib_id):
        """導航到指定題庫頁面"""
        try:
            self.logger.info(f"導航到題庫ID: {lib_id}")
            
            # 方法1: 直接訪問題庫URL
            current_url = driver.current_url
            self.logger.info(f"當前頁面URL: {current_url}")
            
            # 如果已經在正確的題庫頁面，直接返回
            if f'/subject-lib/{lib_id}' in current_url:
                self.logger.info("✅ 已在正確的題庫頁面")
                return True
            
            # 嘗試直接訪問題庫頁面（使用正確的URL格式）
            from config import BASE_URL
            library_url = f"{BASE_URL}/subject-lib/{lib_id}/edit#/"
            self.logger.info(f"嘗試直接訪問題庫頁面: {library_url}")
            driver.get(library_url)
            time.sleep(5)  # 增加等待時間讓頁面完全載入
            
            # 檢查頁面是否載入成功
            current_url_after = driver.current_url
            self.logger.info(f"載入後頁面URL: {current_url_after}")
            
            # 檢查URL是否包含正確的題庫ID
            if f'/subject-lib/{lib_id}' in current_url_after:
                self.logger.info("✅ 成功導航到題庫頁面")
                return True
            
            # 方法2: 通過導航菜單到達題庫頁面
            self.logger.info("直接訪問失敗，嘗試通過導航菜單...")
            return self.navigate_via_menu(driver, lib_id)
            
        except Exception as e:
            self.logger.exception(f"導航到題庫頁面時發生錯誤: {e}")
            return False
    
    def navigate_via_menu(self, driver, lib_id):
        """通過導航菜單到達題庫頁面"""
        try:
            # 點擊 我的主頁
            self.logger.info("通過菜單導航: 點擊 我的主頁...")
            homepage_selectors = [
                "//a[contains(text(), '我的主頁')]",
                "//span[contains(text(), '我的主頁')]",
                "//li[contains(text(), '我的主頁')]"
            ]
            
            for selector in homepage_selectors:
                try:
                    homepage_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    homepage_elem.click()
                    self.logger.info("已點擊 我的主頁")
                    break
                except TimeoutException:
                    continue
            else:
                self.logger.warning("找不到 我的主頁 按鈕")
            
            time.sleep(3)
            
            # 點擊 我的資源
            self.logger.info("通過菜單導航: 點擊 我的資源...")
            resource_selectors = [
                "//a[contains(text(), '我的資源')]",
                "//span[contains(text(), '我的資源')]",
                "//li[contains(text(), '我的資源')]"
            ]
            
            for selector in resource_selectors:
                try:
                    resource_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    resource_elem.click()
                    self.logger.info("已點擊 我的資源")
                    break
                except TimeoutException:
                    continue
            else:
                self.logger.warning("找不到 我的資源 按鈕")
            
            time.sleep(3)
            
            # 點擊 個人題庫
            self.logger.info("通過菜單導航: 點擊 個人題庫...")
            library_selectors = [
                "//a[contains(text(), '個人題庫')]",
                "//span[contains(text(), '個人題庫')]",
                "//li[contains(text(), '個人題庫')]"
            ]
            
            for selector in library_selectors:
                try:
                    library_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    library_elem.click()
                    self.logger.info("已點擊 個人題庫")
                    break
                except TimeoutException:
                    continue
            else:
                self.logger.warning("找不到 個人題庫 按鈕")
            
            time.sleep(5)
            
            # 在題庫列表中找到指定ID的題庫並點擊
            self.logger.info(f"在題庫列表中查找題庫ID: {lib_id}")
            
            # 嘗試多種選擇器來找到題庫
            library_link_selectors = [
                f"//a[contains(@href, '/subject-lib/{lib_id}')]",
                f"//a[contains(@href, '/subject-lib/{lib_id}/')]",
                f"//td[text()='{lib_id}']//following-sibling::td//a",
                f"//*[contains(text(), '{lib_id}')]//ancestor::tr//a"
            ]
            
            for selector in library_link_selectors:
                try:
                    library_link = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    library_link.click()
                    self.logger.info(f"已點擊題庫ID {lib_id} 的連結")
                    time.sleep(3)
                    
                    # 檢查是否成功進入題庫頁面
                    current_url = driver.current_url
                    self.logger.info(f"點擊題庫連結後URL: {current_url}")
                    
                    # 如果不是edit頁面，嘗試導航到edit頁面
                    if f'/subject-lib/{lib_id}' in current_url:
                        if '/edit' not in current_url:
                            edit_url = f"{current_url.rstrip('/')}/edit#/"
                            self.logger.info(f"導航到edit頁面: {edit_url}")
                            driver.get(edit_url)
                            time.sleep(3)
                        self.logger.info("✅ 成功通過菜單導航到題庫頁面")
                        return True
                    break
                except TimeoutException:
                    continue
            
            self.logger.warning("通過菜單導航失敗")
            return False
            
        except Exception as e:
            self.logger.exception(f"通過菜單導航時發生錯誤: {e}")
            return False
    
    def word_import_and_ai_convert(self, driver, lib_id, upload_id):
        """執行Word匯入和AI轉換流程"""
        self.logger.info(f"=== 開始Word匯入和AI轉換 (題庫ID: {lib_id}) ===")
        
        try:
            # 策略1: 直接跳轉到Word匯入頁面（最高效）
            self.logger.info("步驟1: 嘗試直接跳轉到Word匯入頁面...")
            if self.try_direct_word_import_page(driver, lib_id, upload_id):
                return True
            
            # 策略2: 如果直接跳轉失敗，使用原有的hover流程
            self.logger.info("步驟2: 直接跳轉失敗，嘗試hover流程...")
            return self.try_hover_word_import_flow(driver, lib_id, upload_id)
                
        except Exception as e:
            self.logger.exception(f"Word匯入和AI轉換時發生錯誤: {e}")
            return False
    
    def try_direct_word_import_page(self, driver, lib_id, upload_id):
        """嘗試直接跳轉到Word匯入頁面"""
        try:
            from config import BASE_URL
            direct_url = f"{BASE_URL}/subject-lib/{lib_id}/import?mode=word#/"
            
            self.logger.info(f"直接跳轉到: {direct_url}")
            driver.get(direct_url)
            time.sleep(5)  # 等待頁面完全載入
            
            # 驗證URL是否正確
            current_url = driver.current_url
            self.logger.info(f"跳轉後當前URL: {current_url}")
            
            expected_pattern = f"/subject-lib/{lib_id}/import\\?mode=word"
            if not re.search(expected_pattern, current_url):
                self.logger.warning(f"URL不符合預期模式: {expected_pattern}")
                # 但仍然繼續嘗試，可能URL格式有變化
            
            # 直接嘗試第二次Word匯入點擊
            self.logger.info("在直接跳轉的頁面嘗試點擊Word匯入...")
            return self.perform_second_word_import_click(driver, lib_id, upload_id)
            
        except Exception as e:
            self.logger.exception(f"直接跳轉Word匯入頁面時發生錯誤: {e}")
            return False
    
    def try_hover_word_import_flow(self, driver, lib_id, upload_id):
        """使用hover流程進行Word匯入"""
        try:
            # 先導航到正確的題庫頁面
            self.logger.info("導航到題庫頁面...")
            if not self.navigate_to_library_page(driver, lib_id):
                self.logger.error("無法導航到題庫頁面")
                return False
            
            # hover 匯入試題
            self.logger.info("hover 匯入試題...")
            import_selectors = [
                "//span[contains(text(), '匯入試題')]",
                "//a[contains(text(), '匯入試題')]",
                "//button[contains(text(), '匯入試題')]"
            ]
            
            import_elem = None
            for selector in import_selectors:
                try:
                    import_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if import_elem:
                actions = ActionChains(driver)
                actions.move_to_element(import_elem).perform()
                self.logger.info("已hover 匯入試題")
                time.sleep(2)
            else:
                self.logger.warning("找不到 匯入試題 元素，繼續嘗試")
            
            # 點擊 Word 匯入（第一次點擊，導致頁面跳轉）
            self.logger.info("第一次點擊 Word 匯入（頁面跳轉）...")
            word_import_selectors = [
                "//span[contains(text(), 'Word 匯入')]",
                "//a[contains(text(), 'Word 匯入')]",
                "//li[contains(text(), 'Word 匯入')]"
            ]
            
            for selector in word_import_selectors:
                try:
                    word_import_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    word_import_elem.click()
                    self.logger.info("已第一次點擊 Word 匯入")
                    break
                except TimeoutException:
                    continue
            else:
                self.logger.error("找不到 Word 匯入 選項")
                return False
            
            time.sleep(3)
            
            # 檢查是否跳轉到匯入頁面
            current_url = driver.current_url
            self.logger.info(f"第一次點擊後URL: {current_url}")
            
            expected_pattern = f"/subject-lib/{lib_id}/import\\?mode=word"
            if re.search(expected_pattern, current_url):
                self.logger.info("✅ 成功跳轉到Word匯入頁面")
                return self.perform_second_word_import_click(driver, lib_id, upload_id)
            else:
                self.logger.error(f"hover流程跳轉失敗: {current_url}")
                return False
                
        except Exception as e:
            self.logger.exception(f"hover Word匯入流程時發生錯誤: {e}")
            return False
    
    def perform_second_word_import_click(self, driver, lib_id, upload_id):
        """執行第二次Word匯入點擊和後續流程"""
        try:
            self.logger.info("=== 執行第二次Word匯入點擊 ===")
            
            # 等待頁面載入完成
            time.sleep(3)
            
            # 記錄點擊前的頁面狀態
            initial_popup_count = self.count_popup_elements(driver)
            self.logger.info(f"點擊前彈窗元素數量: {initial_popup_count}")
            
            # 使用更全面的選擇器列表，包括特定屬性
            comprehensive_selectors = [
                # 基於用戶提到的特定屬性
                "//*[@data-v-04cb9283 and contains(text(), 'Word 匯入')]",
                "//*[contains(@data-v, '04cb9283')]//span[contains(text(), 'Word')]",
                "//*[@data-v-04cb9283]//span[contains(text(), 'Word')]",
                
                # CSS選擇器轉換為XPath
                "//*[contains(@class, 'word-import') or contains(@id, 'word')]//span[contains(text(), 'Word')]",
                
                # 原有選擇器
                "//span[contains(text(), 'Word 匯入')]",
                "//a[contains(text(), 'Word 匯入')]",
                "//li[contains(text(), 'Word 匯入')]",
                "//button[contains(text(), 'Word 匯入')]",
                
                # 更具體的選擇器
                "//div[contains(text(), 'Word 匯入')]",
                "//*[text()='Word 匯入']",
                "//*[contains(text(), 'Word') and contains(text(), '匯入')]"
            ]
            
            second_click_success = self.perform_enhanced_click(driver, comprehensive_selectors)
            
            if not second_click_success:
                # 如果常規點擊失敗，嘗試JavaScript點擊
                self.logger.info("常規點擊失敗，嘗試JavaScript點擊...")
                second_click_success = self.try_javascript_click(driver, comprehensive_selectors)
            
            if not second_click_success:
                self.logger.error("所有點擊方式都失敗了")
                return False
            
            # 等待頁面反應
            time.sleep(3)
            
            # 驗證點擊是否有效果
            self.logger.info("驗證第二次點擊的效果...")
            click_effective = self.verify_second_click_effect(driver, initial_popup_count)
            
            if not click_effective:
                self.logger.error("第二次點擊沒有產生預期的頁面交互效果")
                return False
            
            # 檢查是否有彈窗出現
            self.logger.info("檢查彈窗是否正確顯示...")
            popup_found = self.check_word_import_popup(driver)
            
            if not popup_found:
                self.logger.error("驗證點擊有效但沒有找到預期的彈窗")
                return False
            
            # 點擊"智慧 Word 匯入"
            self.logger.info("點擊智慧 Word 匯入...")
            smart_import_success = self.click_smart_word_import(driver)
            
            if not smart_import_success:
                self.logger.error("❌ 點擊智慧 Word 匯入失敗")
                return False
            
            time.sleep(3)
            
            # 調用AI轉換API
            self.logger.info("調用AI轉換API...")
            return self.call_ai_convert_and_wait(driver, upload_id, lib_id)
            
        except Exception as e:
            self.logger.exception(f"執行第二次Word匯入點擊時發生錯誤: {e}")
            return False
    
    def perform_enhanced_click(self, driver, selectors):
        """增強的點擊方法，嘗試多種點擊策略"""
        try:
            for selector in selectors:
                self.logger.info(f"嘗試選擇器: {selector}")
                
                try:
                    # 方法1: 標準WebDriver點擊
                    elements = driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        if elem.is_displayed() and elem.is_enabled():
                            elem_text = elem.text
                            elem_tag = elem.tag_name
                            elem_attrs = driver.execute_script(
                                "var items = {}; "
                                "for (index = 0; index < arguments[0].attributes.length; ++index) { "
                                "items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; "
                                "return items;", elem
                            )
                            
                            self.logger.info(f"找到元素: 標籤={elem_tag}, 文字='{elem_text}'")
                            self.logger.info(f"元素屬性: {elem_attrs}")
                            
                            try:
                                # 嘗試滾動到元素
                                driver.execute_script("arguments[0].scrollIntoView();", elem)
                                time.sleep(1)
                                
                                # 嘗試點擊
                                elem.click()
                                self.logger.info(f"✅ 成功點擊元素: {selector}")
                                time.sleep(2)
                                
                                # 立即檢查是否有反應
                                if self.quick_interaction_check(driver):
                                    return True
                                    
                            except Exception as click_err:
                                self.logger.warning(f"標準點擊失敗: {click_err}")
                                
                                # 方法2: ActionChains點擊
                                try:
                                    actions = ActionChains(driver)
                                    actions.move_to_element(elem).click().perform()
                                    self.logger.info(f"✅ ActionChains點擊成功: {selector}")
                                    time.sleep(2)
                                    
                                    if self.quick_interaction_check(driver):
                                        return True
                                        
                                except Exception as action_err:
                                    self.logger.warning(f"ActionChains點擊失敗: {action_err}")
                                    continue
                            
                except TimeoutException:
                    self.logger.debug(f"選擇器無匹配元素: {selector}")
                    continue
                except Exception as e:
                    self.logger.debug(f"選擇器處理錯誤 {selector}: {e}")
                    continue
            
            return False
            
        except Exception as e:
            self.logger.exception(f"增強點擊方法發生錯誤: {e}")
            return False
    
    def try_javascript_click(self, driver, selectors):
        """使用JavaScript進行點擊"""
        try:
            self.logger.info("嘗試JavaScript點擊方式...")
            
            for selector in selectors:
                try:
                    # 轉換XPath為可用的元素查找
                    js_code = f"""
                    var xpath = "{selector}";
                    var result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                    var element = result.singleNodeValue;
                    
                    if (element && element.offsetParent !== null) {{
                        element.scrollIntoView();
                        element.click();
                        return true;
                    }}
                    return false;
                    """
                    
                    success = driver.execute_script(js_code)
                    if success:
                        self.logger.info(f"✅ JavaScript點擊成功: {selector}")
                        time.sleep(3)
                        
                        if self.quick_interaction_check(driver):
                            return True
                            
                except Exception as e:
                    self.logger.debug(f"JavaScript點擊失敗 {selector}: {e}")
                    continue
            
            # 嘗試通用的JavaScript點擊所有包含"Word 匯入"的元素
            generic_js = """
            var elements = document.querySelectorAll('*');
            for (var i = 0; i < elements.length; i++) {
                var element = elements[i];
                if (element.textContent && element.textContent.includes('Word 匯入') && 
                    element.offsetParent !== null && 
                    element.style.display !== 'none') {
                    console.log('Found Word Import element:', element);
                    element.click();
                    return true;
                }
            }
            return false;
            """
            
            success = driver.execute_script(generic_js)
            if success:
                self.logger.info("✅ 通用JavaScript點擊成功")
                time.sleep(3)
                return self.quick_interaction_check(driver)
            
            return False
            
        except Exception as e:
            self.logger.exception(f"JavaScript點擊方法發生錯誤: {e}")
            return False
    
    def quick_interaction_check(self, driver):
        """快速檢查點擊是否有交互反應"""
        try:
            # 檢查是否出現新的可見元素
            interactive_indicators = [
                "//div[contains(@class, 'modal') and contains(@style, 'display: block')]",
                "//div[contains(@class, 'modal') and not(contains(@style, 'display: none'))]",
                "//button[contains(text(), '智慧')]",
                "//span[contains(text(), '剩餘')]",
                "//*[contains(text(), 'credit')]"
            ]
            
            for selector in interactive_indicators:
                elements = driver.find_elements(By.XPATH, selector)
                visible_elements = [e for e in elements if e.is_displayed()]
                if visible_elements:
                    self.logger.info(f"✅ 檢測到交互反應: {selector}")
                    return True
            
            return False
            
        except Exception as e:
            self.logger.debug(f"快速交互檢查錯誤: {e}")
            return False
    
    def count_popup_elements(self, driver):
        """計算當前頁面的彈窗元素數量"""
        try:
            popup_selectors = [
                "//div[contains(@class, 'modal')]",
                "//div[contains(@class, 'popup')]", 
                "//div[contains(@class, 'dialog')]",
                "//div[contains(@class, 'ant-modal')]"
            ]
            
            total_count = 0
            for selector in popup_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    visible_count = sum(1 for elem in elements if elem.is_displayed())
                    total_count += visible_count
                    self.logger.debug(f"選擇器 {selector}: {visible_count} 個可見元素")
                except Exception:
                    continue
            
            return total_count
            
        except Exception as e:
            self.logger.debug(f"計算彈窗元素數量時發生錯誤: {e}")
            return 0
    
    def verify_second_click_effect(self, driver, initial_popup_count):
        """驗證第二次點擊是否有效果（彈窗數量變化或其他交互）"""
        try:
            # 檢查彈窗數量是否有變化
            current_popup_count = self.count_popup_elements(driver)
            self.logger.info(f"點擊後彈窗元素數量: {current_popup_count}")
            
            if current_popup_count > initial_popup_count:
                self.logger.info("✅ 檢測到新彈窗出現，點擊有效")
                return True
            
            # 檢查是否有新的可見元素出現（彈窗內容）
            interactive_selectors = [
                "//button[contains(text(), '智慧')]",
                "//span[contains(text(), '智慧')]",
                "//div[contains(text(), '剩餘')]",
                "//div[contains(text(), 'credit')]",
                "//button[contains(@class, 'ant-btn')]//span[contains(text(), 'Word')]"
            ]
            
            for selector in interactive_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    visible_elements = [elem for elem in elements if elem.is_displayed()]
                    if visible_elements:
                        self.logger.info(f"✅ 檢測到交互元素出現: {selector} ({len(visible_elements)}個)")
                        return True
                except Exception:
                    continue
            
            # 檢查URL是否有變化（某些情況下點擊可能導致URL變化）
            current_url = driver.current_url
            self.logger.info(f"驗證時當前URL: {current_url}")
            
            # 檢查頁面文字是否有新內容
            try:
                page_text = driver.find_element(By.TAG_NAME, "body").text
                credit_keywords = ["credit", "額度", "剩餘", "remaining", "智慧 Word 匯入", "Word 匯入"]
                
                found_keywords = [keyword for keyword in credit_keywords if keyword in page_text]
                if found_keywords:
                    self.logger.info(f"✅ 檢測到相關關鍵字: {found_keywords}")
                    return True
            except Exception:
                pass
            
            self.logger.warning("沒有檢測到明顯的點擊效果")
            return False
            
        except Exception as e:
            self.logger.exception(f"驗證點擊效果時發生錯誤: {e}")
            return False
    
    def try_alternative_second_click(self, driver, original_selectors, already_tried):
        """嘗試其他可能的第二次點擊方式"""
        try:
            self.logger.info("嘗試其他方式進行第二次點擊...")
            
            # 嘗試其他可能的選擇器
            alternative_selectors = [
                "//button[contains(text(), 'Word')]",
                "//a[contains(@title, 'Word')]",
                "//div[contains(text(), 'Word 匯入')]",
                "//span[contains(@class, 'anticon')]//following-sibling::span[contains(text(), 'Word')]",
                "//li[contains(@role, 'menuitem')]//span[contains(text(), 'Word')]"
            ]
            
            # 移除已經嘗試過的選擇器
            if already_tried:
                alternative_selectors = [s for s in alternative_selectors if s != already_tried]
            
            for selector in alternative_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        if elem.is_displayed() and elem.is_enabled():
                            self.logger.info(f"嘗試點擊替代元素: {selector}")
                            elem.click()
                            time.sleep(3)
                            
                            # 檢查是否有效果
                            popup_found = self.check_word_import_popup(driver)
                            if popup_found:
                                self.logger.info("✅ 替代點擊方式成功")
                                return True
                except Exception as e:
                    self.logger.debug(f"替代選擇器失敗 {selector}: {e}")
                    continue
            
            self.logger.error("所有替代點擊方式都失敗了")
            return False
            
        except Exception as e:
            self.logger.exception(f"嘗試替代點擊時發生錯誤: {e}")
            return False
    
    def check_word_import_popup(self, driver):
        """檢查Word匯入彈窗是否正確顯示，包含用戶信用額度信息"""
        try:
            # 檢查彈窗相關元素
            popup_selectors = [
                "//div[contains(@class, 'modal')]",
                "//div[contains(@class, 'popup')]", 
                "//div[contains(@class, 'dialog')]",
                "//div[contains(@class, 'ant-modal')]"
            ]
            
            popup_found = False
            visible_popup = None
            
            for selector in popup_selectors:
                try:
                    popup_elements = driver.find_elements(By.XPATH, selector)
                    for popup_elem in popup_elements:
                        if popup_elem.is_displayed():
                            self.logger.info(f"找到可見彈窗元素: {selector}")
                            popup_found = True
                            visible_popup = popup_elem
                            break
                    if popup_found:
                        break
                except Exception:
                    continue
            
            if not popup_found:
                self.logger.info("沒有找到可見的彈窗元素")
                return False
            
            # 檢查彈窗內容是否包含用戶信用額度信息
            credit_keywords = ["credit", "額度", "剩餘", "remaining", "智慧", "匯入"]
            
            # 優先檢查彈窗內的文字
            popup_text = visible_popup.text if visible_popup else ""
            page_text = driver.find_element(By.TAG_NAME, "body").text
            
            combined_text = popup_text + " " + page_text
            
            found_keywords = [keyword for keyword in credit_keywords if keyword in combined_text]
            if found_keywords:
                self.logger.info(f"✅ 檢測到包含信用額度信息的彈窗，關鍵字: {found_keywords}")
                return True
            else:
                self.logger.info("彈窗不包含預期的信用額度信息")
                self.logger.debug(f"彈窗文字: {popup_text[:200]}...")
                return False
                
        except Exception as e:
            self.logger.exception(f"檢查Word匯入彈窗時發生錯誤: {e}")
            return False
    
    def click_smart_word_import(self, driver):
        """點擊智慧Word匯入按鈕"""
        try:
            smart_import_selectors = [
                "//span[contains(text(), '智慧 Word 匯入')]",
                "//button[contains(text(), '智慧 Word 匯入')]",
                "//a[contains(text(), '智慧 Word 匯入')]",
                "//span[contains(text(), '智慧Word匯入')]",
                "//button[contains(text(), '智慧Word匯入')]",
                "//span[contains(text(), 'Word 匯入')]",  # 備選
                "//button[contains(text(), 'Word 匯入')]"  # 備選
            ]
            
            for selector in smart_import_selectors:
                try:
                    smart_import_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    smart_import_elem.click()
                    self.logger.info(f"已點擊智慧Word匯入按鈕: {selector}")
                    return True
                except TimeoutException:
                    continue
            
            self.logger.error("找不到智慧Word匯入按鈕")
            return False
            
        except Exception as e:
            self.logger.exception(f"點擊智慧Word匯入時發生錯誤: {e}")
            return False
    
    def handle_word_import_page(self, driver, upload_id, lib_id):
        """處理Word匯入頁面的邏輯"""
        try:
            self.logger.info("處理Word匯入頁面...")
            
            # 在匯入頁面再次尋找並點擊Word匯入按鈕
            word_import_selectors = [
                "//span[contains(text(), 'Word 匯入')]",
                "//button[contains(text(), 'Word 匯入')]",
                "//a[contains(text(), 'Word 匯入')]"
            ]
            
            for selector in word_import_selectors:
                try:
                    word_import_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    word_import_elem.click()
                    self.logger.info("在匯入頁面點擊了Word匯入")
                    break
                except TimeoutException:
                    continue
            else:
                self.logger.warning("在匯入頁面找不到Word匯入按鈕，直接調用API")
            
            time.sleep(2)
            
            return self.call_ai_convert_and_wait(driver, upload_id, lib_id)
            
        except Exception as e:
            self.logger.exception(f"處理Word匯入頁面時發生錯誤: {e}")
            return False
    
    def call_ai_convert_and_wait(self, driver, upload_id, lib_id):
        """調用AI轉換API並等待完成，然後等待前端顯示結果"""
        try:
            # 調用AI轉換API（新版本會等待SSE流完成）
            self.logger.info("調用改進版AI轉換API...")
            ai_convert_result = self.call_ai_convert_api(driver, upload_id, "library", lib_id)
            
            if ai_convert_result:
                self.logger.info("✅ AI轉換和SSE流處理完成")
                
                # 等待前端顯示AI轉換結果
                self.logger.info("等待前端顯示AI轉換結果...")
                frontend_ready = self.wait_for_frontend_results(driver)
                
                if frontend_ready:
                    self.logger.info("✅ 前端已顯示轉換結果")
                    
                    # 尋找並點擊"識別"按鈕
                    self.logger.info("尋找並點擊識別按鈕...")
                    identify_success = self.click_identify_button(driver)
                    
                    if identify_success:
                        self.logger.info("✅ 已點擊識別按鈕")
                        return True
                    else:
                        self.logger.error("❌ 無法找到或點擊識別按鈕")
                        return False
                else:
                    self.logger.error("❌ 前端未顯示轉換結果")
                    return False
            else:
                self.logger.error("❌ AI轉換失敗")
                return False
                
        except Exception as e:
            self.logger.exception(f"調用AI轉換並等待時發生錯誤: {e}")
            return False
    
    def wait_for_frontend_results(self, driver):
        """等待前端顯示AI轉換結果"""
        try:
            self.logger.info("等待前端顯示轉換結果，最多等待60秒...")
            
            # 檢查前端結果顯示的指標
            result_indicators = [
                # 識別按鈕出現
                "//button[contains(text(), '識別')]",
                "//span[contains(text(), '識別')]",
                "//a[contains(text(), '識別')]",
                
                # 轉換結果相關文字
                "//div[contains(text(), '轉換完成')]",
                "//div[contains(text(), '解析完成')]",
                "//div[contains(text(), '匯入完成')]",
                
                # 問題/題目相關元素
                "//div[contains(text(), '題目')]",
                "//div[contains(text(), '問題')]",
                "//div[contains(text(), '選項')]",
                
                # 表格或列表顯示結果
                "//table[contains(@class, 'result')]",
                "//div[contains(@class, 'question')]",
                "//ul[contains(@class, 'import')]"
            ]
            
            wait = WebDriverWait(driver, 60, poll_frequency=2.0)
            
            for indicator in result_indicators:
                try:
                    element = wait.until(
                        EC.presence_of_element_located((By.XPATH, indicator))
                    )
                    if element.is_displayed():
                        self.logger.info(f"✅ 檢測到前端結果指標: {indicator}")
                        return True
                except TimeoutException:
                    continue
            
            # 如果沒有找到明確指標，檢查頁面是否有變化
            self.logger.info("沒有找到明確指標，檢查頁面內容變化...")
            time.sleep(5)  # 額外等待
            
            try:
                page_text = driver.find_element(By.TAG_NAME, "body").text
                result_keywords = ["識別", "轉換", "完成", "題目", "選項", "匯入", "結果"]
                
                found_keywords = [keyword for keyword in result_keywords if keyword in page_text]
                if found_keywords:
                    self.logger.info(f"✅ 頁面包含結果相關關鍵字: {found_keywords}")
                    return True
                else:
                    self.logger.warning("頁面沒有顯示預期的結果內容")
                    return False
                    
            except Exception as e:
                self.logger.error(f"檢查頁面內容時發生錯誤: {e}")
                return False
                
        except Exception as e:
            self.logger.exception(f"等待前端結果時發生錯誤: {e}")
            return False
    
    def click_identify_button(self, driver):
        """尋找並點擊識別按鈕"""
        try:
            identify_selectors = [
                # 基於用戶提到的特定屬性
                "//*[@data-v-04cb9283 and contains(text(), '識別')]",
                "//*[contains(@data-v, '04cb9283')]//span[contains(text(), '識別')]",
                "//*[@data-v-04cb9283]//button[contains(text(), '識別')]",
                
                # 標準選擇器
                "//button[contains(text(), '識別')]",
                "//span[contains(text(), '識別')]",
                "//a[contains(text(), '識別')]",
                "//input[@type='button' and contains(@value, '識別')]",
                "//div[contains(text(), '識別') and contains(@class, 'btn')]",
                "//*[@role='button' and contains(text(), '識別')]"
            ]
            
            for selector in identify_selectors:
                try:
                    identify_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    
                    # 記錄找到的元素信息
                    elem_text = identify_elem.text
                    elem_tag = identify_elem.tag_name
                    self.logger.info(f"找到識別按鈕: {selector} - 標籤={elem_tag}, 文字='{elem_text}'")
                    
                    # 嘗試滾動到元素
                    driver.execute_script("arguments[0].scrollIntoView();", identify_elem)
                    time.sleep(1)
                    
                    # 點擊識別按鈕
                    identify_elem.click()
                    self.logger.info(f"✅ 已點擊識別按鈕")
                    
                    # 等待點擊反應
                    time.sleep(2)
                    return True
                    
                except TimeoutException:
                    continue
                except Exception as e:
                    self.logger.debug(f"點擊識別按鈕失敗 {selector}: {e}")
                    continue
            
            # 如果標準點擊失敗，嘗試JavaScript點擊
            self.logger.info("標準點擊失敗，嘗試JavaScript點擊識別按鈕...")
            js_code = """
            var elements = document.querySelectorAll('*');
            for (var i = 0; i < elements.length; i++) {
                var element = elements[i];
                if (element.textContent && element.textContent.includes('識別') && 
                    element.offsetParent !== null) {
                    console.log('找到識別元素:', element);
                    element.click();
                    return true;
                }
            }
            return false;
            """
            
            success = driver.execute_script(js_code)
            if success:
                self.logger.info("✅ JavaScript點擊識別按鈕成功")
                time.sleep(2)
                return True
            
            self.logger.error("所有方式都無法找到或點擊識別按鈕")
            return False
            
        except Exception as e:
            self.logger.exception(f"點擊識別按鈕時發生錯誤: {e}")
            return False
    
    def call_ai_convert_api(self, driver, upload_id, belong, belong_id):
        """在瀏覽器環境中調用AI轉換API（改進版，支持SSE流處理）"""
        self.logger.info(f"調用AI轉換API: upload_id={upload_id}, belong={belong}, belong_id={belong_id}")
        
        try:
            import uuid
            marker = f"selenium-{uuid.uuid4()}"  # 用於標記這次請求
            
            js_code = f"""
            const done = arguments[0];
            (async () => {{
              try {{
                const resp = await fetch('/api/data-import/ai-convert', {{
                  method: 'POST',
                  headers: {{
                    'Content-Type': 'application/json',
                    'X-From-Selenium': '{marker}'   // 指紋標頭（稽核用）
                  }},
                  body: JSON.stringify({{
                    upload_id: {upload_id},
                    belong: '{belong}',
                    belong_id: {belong_id}
                  }})
                }});

                // 不是 2xx 直接回 Selenium
                if (!resp.ok) {{
                  done({{
                    ok: false,
                    status: resp.status,
                    statusText: resp.statusText,
                    marker: '{marker}'
                  }});
                  return;
                }}

                // 嘗試判斷是否為 SSE
                const ctype = resp.headers.get('content-type') || '';
                if (ctype.includes('text/event-stream')) {{
                  // 逐塊讀取 SSE，收集最後一筆或完整內容
                  const reader = resp.body.getReader();
                  const decoder = new TextDecoder();
                  let full = '';
                  let lastEvent = null;
                  let eventCount = 0;

                  while (true) {{
                    const {{ value, done: streamDone }} = await reader.read();
                    if (streamDone) break;
                    const chunk = decoder.decode(value, {{stream: true}});
                    full += chunk;

                    // 粗略解析 SSE 格式：以 \\n\\n 分段，抓最後一段的 data
                    const parts = full.split('\\n\\n');
                    for (let part of parts) {{
                      if (part.startsWith('data:')) {{
                        const raw = part.replace(/^data:\\s*/, '').trim();
                        if (raw && raw !== '[DONE]') {{
                          lastEvent = raw;
                          eventCount++;
                        }}
                      }}
                    }}
                  }}
                  
                  done({{
                    ok: true,
                    type: 'sse',
                    marker: '{marker}',
                    lastEvent: lastEvent,
                    eventCount: eventCount,
                    streamLength: full.length
                  }});
                  return;
                }} else {{
                  // 一般 JSON 回應
                  const data = await resp.json().catch(() => null);
                  done({{
                    ok: true,
                    type: 'json',
                    marker: '{marker}',
                    data
                  }});
                  return;
                }}
              }} catch (err) {{
                done({{
                  ok: false,
                  error: String(err),
                  marker: '{marker}'
                }});
              }}
            }})();
            """
            
            # 設置較長的超時時間，因為AI轉換可能需要時間
            driver.set_script_timeout(120)  # 2分鐘超時
            
            result = driver.execute_async_script(js_code)
            self.logger.info(f"AI轉換API響應: {result}")
            
            # 檢查結果
            if result.get('ok', False):
                if result.get('type') == 'sse':
                    self.logger.info(f"✅ SSE流處理完成，事件數量: {result.get('eventCount', 0)}")
                    self.logger.info(f"最後事件: {result.get('lastEvent', 'None')}")
                    return True
                elif result.get('type') == 'json':
                    self.logger.info(f"✅ JSON響應處理完成: {result.get('data')}")
                    return True
                else:
                    self.logger.info("✅ API調用成功")
                    return True
            else:
                error_info = result.get('error', result.get('statusText', '未知錯誤'))
                status = result.get('status', 'unknown')
                self.logger.error(f"❌ AI轉換API失敗: 狀態={status}, 錯誤={error_info}")
                return False
            
        except Exception as e:
            self.logger.exception(f"調用AI轉換API時發生錯誤: {e}")
            return False
        finally:
            # 恢復默認超時時間
            try:
                driver.set_script_timeout(30)
            except:
                pass

def word_import_and_convert(driver, lib_id, upload_id, logger=None):
    """外部調用接口：執行Word匯入和AI轉換"""
    importer = WordImporter(logger)
    return importer.word_import_and_ai_convert(driver, lib_id, upload_id)

if __name__ == "__main__":
    # 測試用代碼
    print("這是Word匯入子模組，請通過其他程序調用")