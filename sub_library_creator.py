#!/usr/bin/env python3
"""
sub_library_creator.py - 題庫創建子模組
功能：
1. 參考 exam_3_xmlpouring.py 的流程創建測驗題庫
2. 點擊 我的主頁 -> 我的資源 -> 個人題庫 -> hover 新增 -> 點擊 測驗題庫
3. 獲取題庫ID並返回
"""

import time
import logging
import re
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class LibraryCreator:
    def __init__(self, logger=None):
        self.logger = logger or self._setup_default_logger()
    
    def _setup_default_logger(self):
        """設置默認logger"""
        logger = logging.getLogger('LibraryCreator')
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        return logger
    
    def create_test_library(self, driver, title=None):
        """創建測驗題庫，參考exam_3_xmlpouring.py的流程"""
        self.logger.info("=== 開始創建測驗題庫 ===")
        
        try:
            # 點擊 我的主頁
            self.logger.info("步驟1: 點擊 我的主頁...")
            homepage_selectors = [
                "//a[contains(text(), '我的主頁')]",
                "//span[contains(text(), '我的主頁')]",
                "//div[contains(text(), '我的主頁')]",
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
                self.logger.error("找不到 我的主頁 按鈕")
                return None
            
            time.sleep(3)
            
            # 點擊 我的資源
            self.logger.info("步驟2: 點擊 我的資源...")
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
                self.logger.error("找不到 我的資源 按鈕")
                return None
            
            time.sleep(3)
            
            # 點擊 個人題庫
            self.logger.info("步驟3: 點擊 個人題庫...")
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
                self.logger.error("找不到 個人題庫 按鈕")
                return None
            
            time.sleep(3)
            
            # hover 新增
            self.logger.info("步驟4: hover 新增...")
            new_button_selectors = [
                "//span[text()='新增']",
                "//a[contains(text(), '新增')]",
                "//button[contains(text(), '新增')]"
            ]
            
            new_button = None
            for selector in new_button_selectors:
                try:
                    new_button = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if not new_button:
                self.logger.error("找不到 新增 按鈕")
                return None
            
            actions = ActionChains(driver)
            actions.move_to_element(new_button).perform()
            self.logger.info("已hover 新增按鈕")
            time.sleep(2)
            
            # 點擊 測驗題庫
            self.logger.info("步驟5: 點擊 測驗題庫...")
            test_library_selectors = [
                "//span[contains(text(), '測驗題庫')]",
                "//a[contains(text(), '測驗題庫')]",
                "//li[contains(text(), '測驗題庫')]"
            ]
            
            for selector in test_library_selectors:
                try:
                    test_library_elem = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    test_library_elem.click()
                    self.logger.info("已點擊 測驗題庫")
                    break
                except TimeoutException:
                    continue
            else:
                self.logger.error("找不到 測驗題庫 選項")
                return None
            
            time.sleep(5)
            
            # 獲取題庫ID
            current_url = driver.current_url
            self.logger.info(f"頁面跳轉後URL: {current_url}")
            
            # 使用多種模式來匹配題庫ID
            patterns = [
                r'/subject-lib/(\d+)/edit',
                r'/subject-lib/(\d+)/',
                r'/subject-lib/(\d+)$',
                r'/subject-lib/(\d+)#'
            ]
            
            lib_id = None
            for pattern in patterns:
                match = re.search(pattern, current_url)
                if match:
                    lib_id = match.group(1)
                    self.logger.info(f"✅ 使用模式 '{pattern}' 獲取到題庫ID: {lib_id}")
                    break
            
            if lib_id:
                create_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.logger.info(f"✅ 題庫創建成功: ID={lib_id}, 時間={create_time}")
                return lib_id, create_time
            else:
                self.logger.error(f"無法從URL提取題庫ID: {current_url}")
                return None
            
        except Exception as e:
            self.logger.exception(f"創建測驗題庫時發生錯誤: {e}")
            return None

def create_library(driver, title=None, logger=None):
    """外部調用接口：創建題庫並返回(題庫ID, 創建時間)"""
    creator = LibraryCreator(logger)
    return creator.create_test_library(driver, title)

if __name__ == "__main__":
    # 測試用代碼
    print("這是題庫創建子模組，請通過其他程序調用")