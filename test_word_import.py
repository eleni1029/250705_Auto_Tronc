#!/usr/bin/env python3
"""
test_word_import.py - 測試新的Word匯入流程
用於測試處理彈窗和智慧Word匯入的複雜邏輯
"""

import logging
from datetime import datetime
from tronc_login import setup_driver
from config import BASE_URL
from sub_word_import import WordImporter

def setup_test_logging():
    """設置測試日誌"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(f'test_word_import_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log', encoding='utf-8')
        ]
    )
    return logging.getLogger('TestWordImport')

def test_word_import_flow():
    """測試Word匯入流程"""
    logger = setup_test_logging()
    logger.info("=== 開始測試Word匯入流程 ===")
    
    # 測試參數
    lib_id = "356"  # 使用用戶提到的題庫ID
    upload_id = "12523"  # 使用用戶提到的上傳ID
    
    driver = None
    try:
        # 設置瀏覽器
        logger.info("初始化瀏覽器...")
        driver = setup_driver()
        if not driver:
            logger.error("❌ 無法啟動瀏覽器")
            return False
        
        # 先訪問基礎頁面
        driver.get(BASE_URL)
        
        # 設置cookies（這裡需要有效的cookie）
        from config import COOKIE
        cookies = dict(item.split("=", 1) for item in COOKIE.split("; "))
        for name, value in cookies.items():
            try:
                driver.add_cookie({
                    'name': name,
                    'value': value,
                    'domain': '.tronclass.com'
                })
            except Exception as e:
                logger.debug(f"添加cookie失敗 {name}: {e}")
        
        # 刷新頁面讓cookies生效
        driver.refresh()
        import time
        time.sleep(3)
        
        logger.info(f"測試參數: 題庫ID={lib_id}, 上傳ID={upload_id}")
        
        # 創建WordImporter實例並測試
        word_importer = WordImporter(logger)
        
        # 測試導航到題庫頁面
        logger.info("測試1: 導航到題庫頁面...")
        nav_success = word_importer.navigate_to_library_page(driver, lib_id)
        if not nav_success:
            logger.error("❌ 導航到題庫頁面失敗")
            return False
        logger.info("✅ 導航到題庫頁面成功")
        
        # 測試Word匯入和AI轉換流程（修正版）
        logger.info("測試2: Word匯入和AI轉換（修正版流程）...")
        logger.info("預期流程: 點擊Word匯入 → 頁面跳轉 → 再次點擊Word匯入 → 彈窗 → 智慧Word匯入")
        import_success = word_importer.word_import_and_ai_convert(driver, lib_id, upload_id)
        
        if import_success:
            logger.info("✅ Word匯入流程測試成功")
            return True
        else:
            logger.error("❌ Word匯入流程測試失敗")
            return False
            
    except Exception as e:
        logger.exception(f"測試過程中發生錯誤: {e}")
        return False
    finally:
        if driver:
            try:
                driver.quit()
                logger.info("已關閉瀏覽器")
            except:
                pass

def test_popup_detection():
    """測試彈窗檢測功能"""
    logger = setup_test_logging()
    logger.info("=== 測試彈窗檢測功能 ===")
    
    driver = None
    try:
        # 設置瀏覽器
        driver = setup_driver()
        if not driver:
            logger.error("❌ 無法啟動瀏覽器")
            return False
        
        # 創建WordImporter實例
        word_importer = WordImporter(logger)
        
        # 創建一個簡單的測試頁面來模擬彈窗
        test_html = """
        <html>
        <body>
            <div class="ant-modal">
                <div>用戶信用額度: 1967 剩餘</div>
                <button>智慧 Word 匯入</button>
            </div>
        </body>
        </html>
        """
        
        # 使用data URL載入測試頁面
        data_url = f"data:text/html;charset=utf-8,{test_html}"
        driver.get(data_url)
        
        import time
        time.sleep(2)
        
        # 測試彈窗檢測
        logger.info("測試彈窗檢測...")
        popup_found = word_importer.check_word_import_popup(driver)
        
        if popup_found:
            logger.info("✅ 彈窗檢測成功")
            
            # 測試點擊智慧Word匯入
            logger.info("測試點擊智慧Word匯入...")
            smart_click_success = word_importer.click_smart_word_import(driver)
            
            if smart_click_success:
                logger.info("✅ 智慧Word匯入點擊成功")
                return True
            else:
                logger.error("❌ 智慧Word匯入點擊失敗")
                return False
        else:
            logger.error("❌ 彈窗檢測失敗")
            return False
            
    except Exception as e:
        logger.exception(f"彈窗檢測測試中發生錯誤: {e}")
        return False
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass

def main():
    """主測試函數"""
    print("Word匯入流程測試工具")
    print("=" * 40)
    
    try:
        print("自動運行完整流程測試...")
        success = test_word_import_flow()
        
        if success:
            print("\n🎉 測試成功完成！")
        else:
            print("\n❌ 測試失敗，請檢查日誌文件")
            
    except KeyboardInterrupt:
        print("\n用戶中斷測試")
    except Exception as e:
        print(f"\n❌ 測試過程中發生錯誤: {e}")

if __name__ == "__main__":
    main()