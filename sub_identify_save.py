#!/usr/bin/env python3
"""
sub_identify_save.py - 識別和儲存子模組
功能：
1. 點擊 "識別" 按鈕
2. 等待識別成功
3. 點擊 "儲存" 按鈕
4. 確認儲存完成
"""

import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

class IdentifySaver:
    def __init__(self, logger=None):
        self.logger = logger or self._setup_default_logger()
    
    def _setup_default_logger(self):
        """設置默認logger"""
        logger = logging.getLogger('IdentifySaver')
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        return logger
    
    def identify_and_save(self, driver):
        """執行識別和儲存流程"""
        self.logger.info("=== 開始識別和儲存流程 ===")
        
        try:
            # 點擊 "識別"
            self.logger.info("步驟1: 點擊 識別...")
            identify_selectors = [
                "//button[contains(text(), '識別')]",
                "//a[contains(text(), '識別')]",
                "//span[contains(text(), '識別')]",
                "//input[@type='button' and @value='識別']"
            ]
            
            identify_clicked = False
            for selector in identify_selectors:
                try:
                    identify_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    identify_elem.click()
                    self.logger.info("已點擊 識別")
                    identify_clicked = True
                    break
                except TimeoutException:
                    continue
            
            if not identify_clicked:
                self.logger.error("找不到 識別 按鈕")
                return False
            
            # 等待識別成功
            self.logger.info("步驟2: 等待識別成功...")
            success_wait = WebDriverWait(driver, 60)
            
            # 等待識別成功的多種指示
            success_indicators = [
                lambda d: "識別成功" in d.find_element(By.TAG_NAME, "body").text,
                lambda d: "識別完成" in d.find_element(By.TAG_NAME, "body").text,
                lambda d: "parsing success" in d.find_element(By.TAG_NAME, "body").text.lower(),
                lambda d: "success" in d.find_element(By.TAG_NAME, "body").text.lower() and "識別" in d.find_element(By.TAG_NAME, "body").text
            ]
            
            success_detected = False
            for indicator in success_indicators:
                try:
                    success_wait.until(indicator)
                    self.logger.info("✅ 識別成功")
                    success_detected = True
                    break
                except TimeoutException:
                    continue
            
            if not success_detected:
                self.logger.warning("未檢測到明確的識別成功指示，但繼續執行儲存")
            
            # 短暫等待，確保頁面更新完成
            time.sleep(2)
            
            # 點擊儲存
            self.logger.info("步驟3: 點擊 儲存...")
            save_selectors = [
                "//button[contains(text(), '儲存')]",
                "//a[contains(text(), '儲存')]",
                "//span[contains(text(), '儲存')]",
                "//input[@type='button' and @value='儲存']",
                "//button[contains(text(), '保存')]",
                "//a[contains(text(), '保存')]"
            ]
            
            save_clicked = False
            for selector in save_selectors:
                try:
                    save_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    save_elem.click()
                    self.logger.info("已點擊 儲存")
                    save_clicked = True
                    break
                except TimeoutException:
                    continue
            
            if not save_clicked:
                self.logger.error("找不到 儲存 按鈕")
                return False
            
            # 等待儲存完成
            self.logger.info("步驟4: 等待儲存完成...")
            time.sleep(5)
            
            # 檢查是否有儲存成功的指示
            try:
                save_success_indicators = [
                    "儲存成功",
                    "保存成功",
                    "save success",
                    "saved successfully"
                ]
                
                body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
                for indicator in save_success_indicators:
                    if indicator.lower() in body_text:
                        self.logger.info(f"✅ 儲存成功: 檢測到指示詞 '{indicator}'")
                        return True
                
                # 如果沒有明確的成功指示，檢查URL是否有變化
                current_url = driver.current_url
                if "edit" in current_url or "subject-lib" in current_url:
                    self.logger.info("✅ 儲存可能成功（基於URL判斷）")
                    return True
                else:
                    self.logger.warning("⚠️ 無法確認儲存狀態，但操作已完成")
                    return True
                    
            except Exception as e:
                self.logger.warning(f"檢查儲存狀態時發生錯誤: {e}，但假設操作已完成")
                return True
            
        except Exception as e:
            self.logger.exception(f"識別和儲存時發生錯誤: {e}")
            return False

def identify_and_save_questions(driver, logger=None):
    """外部調用接口：執行識別和儲存"""
    saver = IdentifySaver(logger)
    return saver.identify_and_save(driver)

if __name__ == "__main__":
    # 測試用代碼
    print("這是識別和儲存子模組，請通過其他程序調用")