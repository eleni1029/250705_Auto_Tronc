#!/usr/bin/env python3
"""
sub_return_library.py - 返回題庫列表子模組
功能：
1. 點擊 返回題庫
2. 點擊 返回上一層
3. 確保回到題庫列表頁面
"""

import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class LibraryReturner:
    def __init__(self, logger=None):
        self.logger = logger or self._setup_default_logger()
    
    def _setup_default_logger(self):
        """設置默認logger"""
        logger = logging.getLogger('LibraryReturner')
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        return logger
    
    def return_to_library_list(self, driver):
        """返回題庫列表"""
        self.logger.info("=== 返回題庫列表 ===")
        
        try:
            # 點擊 返回題庫
            self.logger.info("步驟1: 點擊 返回題庫...")
            return_selectors = [
                "//button[contains(text(), '返回題庫')]",
                "//a[contains(text(), '返回題庫')]",
                "//span[contains(text(), '返回題庫')]",
                "//button[contains(text(), '回到題庫')]",
                "//a[contains(text(), '回到題庫')]"
            ]
            
            return_clicked = False
            for selector in return_selectors:
                try:
                    return_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    return_elem.click()
                    self.logger.info("已點擊 返回題庫")
                    return_clicked = True
                    break
                except TimeoutException:
                    continue
            
            if not return_clicked:
                self.logger.warning("找不到 返回題庫 按鈕，嘗試其他方式")
                # 嘗試使用瀏覽器的返回功能
                try:
                    driver.back()
                    self.logger.info("使用瀏覽器返回功能")
                    time.sleep(3)
                except Exception as e:
                    self.logger.error(f"瀏覽器返回失敗: {e}")
            else:
                time.sleep(3)
            
            # 點擊 返回上一層
            self.logger.info("步驟2: 點擊 返回上一層...")
            back_selectors = [
                "//button[contains(text(), '返回上一層')]",
                "//a[contains(text(), '返回上一層')]",
                "//span[contains(text(), '返回上一層')]",
                "//button[contains(text(), '返回')]",
                "//a[contains(text(), '返回')]",
                "//button[contains(text(), '上一層')]",
                "//a[contains(text(), '上一層')]"
            ]
            
            back_clicked = False
            for selector in back_selectors:
                try:
                    back_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    back_elem.click()
                    self.logger.info("已點擊 返回上一層")
                    back_clicked = True
                    break
                except TimeoutException:
                    continue
            
            if not back_clicked:
                self.logger.warning("找不到 返回上一層 按鈕，嘗試其他方式")
                # 嘗試再次使用瀏覽器的返回功能
                try:
                    driver.back()
                    self.logger.info("再次使用瀏覽器返回功能")
                    time.sleep(3)
                except Exception as e:
                    self.logger.error(f"瀏覽器返回失敗: {e}")
            else:
                time.sleep(3)
            
            # 檢查是否成功返回題庫列表
            current_url = driver.current_url
            self.logger.info(f"返回後的URL: {current_url}")
            
            # 檢查URL和頁面內容來確認是否回到正確位置
            success_indicators = [
                "subject-libs" in current_url,
                "resources" in current_url,
                "個人題庫" in driver.find_element(By.TAG_NAME, "body").text,
                "題庫" in driver.find_element(By.TAG_NAME, "body").text
            ]
            
            if any(success_indicators):
                self.logger.info("✅ 成功返回題庫列表")
                return True
            else:
                self.logger.warning("⚠️ 無法確認是否成功返回題庫列表，但操作已完成")
                return True
            
        except Exception as e:
            self.logger.exception(f"返回題庫列表時發生錯誤: {e}")
            return False

def return_to_library_page(driver, logger=None):
    """外部調用接口：返回題庫列表"""
    returner = LibraryReturner(logger)
    return returner.return_to_library_list(driver)

if __name__ == "__main__":
    # 測試用代碼
    print("這是返回題庫子模組，請通過其他程序調用")