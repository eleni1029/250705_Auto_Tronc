from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import datetime
import json
from config import USERNAME, PASSWORD, LOGIN_URL, COURSE_ID, BASE_URL, SEARCH_TERM

def setup_driver():
    """設置並返回 Chrome WebDriver"""
    chrome_options = Options()
    # 註解掉 headless 以便看到瀏覽器操作
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
    except Exception as e:
        print(f'無法啟動 ChromeDriver: {e}')
        return None

def save_debug_info(driver, step, success=True):
    """保存調試信息"""
    # 如果成功，不保存 log
    if success:
        print(f'✅ {step} 成功，跳過保存 log')
        return None
    
    # 只在失敗時保存 log
    log_dir = 'log'
    os.makedirs(log_dir, exist_ok=True)
    now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    
    status = 'error'
    filename = f'{step}_{status}_{now}.html'
    filepath = os.path.join(log_dir, filename)
    
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f'<!-- URL: {driver.current_url} -->\n')
            f.write(f'<!-- Title: {driver.title} -->\n')
            f.write(driver.page_source)
        print(f'❌ 調試信息已保存: {filepath}')
        return filepath
    except Exception as e:
        print(f'保存調試信息失敗: {e}')
        return None

def login_and_get_cookie():
    """登入並獲取 cookie"""
    driver = setup_driver()
    if not driver:
        return None
    
    try:
        print(f'正在訪問登入頁面: {LOGIN_URL}')
        driver.get(LOGIN_URL)
        
        # 等待頁面加載
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
        
        print(f'登入頁面標題: {driver.title}')
        # 登入頁面載入成功，不需要保存 log
        
        # 首先搜索並選擇機構
        if SEARCH_TERM:
            print(f'正在搜索機構: {SEARCH_TERM}')
            
            # 查找搜索輸入框
            search_selectors = [
                'input[type="text"]',
                'input[placeholder*="搜索"]',
                'input[placeholder*="search"]',
                '.search-input',
                '#search',
                '[name="search"]'
            ]
            
            search_field = None
            for selector in search_selectors:
                try:
                    search_field = WebDriverWait(driver, 2).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    print(f'找到搜索欄位: {selector}')
                    break
                except TimeoutException:
                    continue
            
            if search_field:
                # 輸入搜索關鍵字
                search_field.clear()
                search_field.send_keys(SEARCH_TERM)
                time.sleep(1)
                
                # 等待搜索結果出現
                try:
                    # 查找搜索結果列表
                    result_selectors = [
                        '.search-result',
                        '.dropdown-item',
                        '.result-item',
                        'ul li',
                        '[role="option"]'
                    ]
                    
                    result_found = False
                    for selector in result_selectors:
                        try:
                            results = WebDriverWait(driver, 3).until(
                                EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector))
                            )
                            
                            # 查找包含搜索關鍵字的結果
                            for result in results:
                                if SEARCH_TERM in result.text:
                                    print(f'找到匹配結果: {result.text}')
                                    driver.execute_script("arguments[0].click();", result)
                                    result_found = True
                                    break
                            
                            if result_found:
                                break
                                
                        except TimeoutException:
                            continue
                    
                    if result_found:
                        print('✅ 成功選擇機構')
                        time.sleep(2)  # 等待頁面加載
                    else:
                        print('⚠️  未找到匹配的搜索結果，繼續嘗試登入')
                        
                except Exception as e:
                    print(f'搜索結果處理出錯: {e}')
                    
            else:
                print('⚠️  未找到搜索輸入框，直接進行登入')
        
        # 查找並填入用戶名
        username_selectors = ['#username', '[name="username"]', '[name="email"]', 'input[type="text"]']
        username_field = None
        
        for selector in username_selectors:
            try:
                username_field = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                print(f'找到用戶名欄位: {selector}')
                break
            except TimeoutException:
                continue
        
        if not username_field:
            print('❌ 找不到用戶名輸入框')
            save_debug_info(driver, 'username_not_found', False)
            return None
        
        # 查找並填入密碼
        password_selectors = ['#password', '[name="password"]', 'input[type="password"]']
        password_field = None
        
        for selector in password_selectors:
            try:
                password_field = driver.find_element(By.CSS_SELECTOR, selector)
                print(f'找到密碼欄位: {selector}')
                break
            except NoSuchElementException:
                continue
        
        if not password_field:
            print('❌ 找不到密碼輸入框')
            save_debug_info(driver, 'password_not_found', False)
            return None
        
        # 填入登入信息
        print('正在填入登入信息...')
        username_field.clear()
        username_field.send_keys(USERNAME)
        time.sleep(0.5)
        
        password_field.clear()
        password_field.send_keys(PASSWORD)
        time.sleep(0.5)
        
        # 提交表單
        submit_selectors = [
            'button[type="submit"]', 
            'input[type="submit"]', 
            '.login-btn',
            '.btn-login',
            'button:contains("登入")',
            'button:contains("Login")'
        ]
        
        submitted = False
        for selector in submit_selectors:
            try:
                submit_button = driver.find_element(By.CSS_SELECTOR, selector)
                print(f'找到提交按鈕: {selector}')
                driver.execute_script("arguments[0].click();", submit_button)
                submitted = True
                break
            except NoSuchElementException:
                continue
        
        if not submitted:
            print('未找到提交按鈕，嘗試按 Enter 鍵')
            password_field.send_keys(Keys.RETURN)
        
        # 等待登入處理
        print('等待登入處理...')
        time.sleep(5)
        
        current_url = driver.current_url
        print(f'登入後 URL: {current_url}')
        
        # 如果還在登入頁面，檢查是否有錯誤
        if 'login' in current_url.lower():
            save_debug_info(driver, 'still_on_login_page', False)
            
            # 檢查錯誤信息
            error_elements = driver.find_elements(By.CSS_SELECTOR, '.error, .alert, .message')
            for elem in error_elements:
                error_text = elem.text.strip()
                if error_text:
                    print(f'錯誤信息: {error_text}')
            
            print('⚠️  可能登入失敗，但繼續嘗試訪問課程頁面...')
        
        # 嘗試訪問課程頁面
        course_url = f'{BASE_URL}/course/{COURSE_ID}/content'
        print(f'正在訪問課程頁面: {course_url}')
        driver.get(course_url)
        
        # 等待課程頁面加載
        time.sleep(8)
        
        final_url = driver.current_url
        page_title = driver.title
        print(f'最終 URL: {final_url}')
        print(f'頁面標題: {page_title}')
        
        # 最終頁面載入成功，不需要保存 log
        
        # 簡單檢查是否成功：不在登入頁面就算成功
        if 'login' not in final_url.lower():
            print('✅ 成功訪問課程頁面！')
            
            # 獲取所有 cookies
            cookies = driver.get_cookies()
            cookie_string = '; '.join([f"{c['name']}={c['value']}" for c in cookies])
            
            print(f'獲取到 {len(cookies)} 個 cookies')
            
            # 嘗試解析模組
            modules = []
            try:
                module_elements = driver.find_elements(By.CSS_SELECTOR, '[data-id]')
                for elem in module_elements:
                    module_id = elem.get_attribute('data-id')
                    name = elem.text.strip()
                    if module_id and name and len(name) > 2:  # 過濾太短的文字
                        modules.append({'id': module_id, 'name': name})
                        print(f'找到模組: {module_id} - {name}')
                
                if modules:
                    with open('module_list.json', 'w', encoding='utf-8') as f:
                        json.dump(modules, f, ensure_ascii=False, indent=2)
                    print(f'✅ 保存了 {len(modules)} 個模組到 module_list.json')
                
            except Exception as e:
                print(f'解析模組時出錯: {e}')
            
            return cookie_string, modules
        else:
            print('❌ 仍在登入頁面，登入可能失敗')
            return None
            
    except Exception as e:
        print(f'登入過程出錯: {e}')
        save_debug_info(driver, 'login_error', False)
        return None
    finally:
        driver.quit()

def update_config(cookie_string, modules):
    """更新配置文件"""
    try:
        with open('config.py', 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        with open('config.py', 'w', encoding='utf-8') as f:
            for line in lines:
                if line.strip().startswith('COOKIE'):
                    f.write(f"COOKIE = '{cookie_string}'  # 自動登入獲取\n")
                elif line.strip().startswith('MODULE_ID') and modules:
                    f.write(f"MODULE_ID = {modules[0]['id']}  # 預設章節 ID\n")
                else:
                    f.write(line)
        
        print('✅ 配置文件已更新')
        
    except Exception as e:
        print(f'更新配置文件時出錯: {e}')

def main():
    """主函數"""
    print('=== TronClass 簡化登入工具 ===')
    print(f'用戶名: {USERNAME}')
    print(f'課程 ID: {COURSE_ID}')
    print('=' * 40)
    
    result = login_and_get_cookie()
    
    if result:
        cookie_string, modules = result
        print('\n🎉 登入成功！')
        print(f'Cookie 長度: {len(cookie_string)} 字元')
        print(f'找到 {len(modules)} 個模組')
        
        update_config(cookie_string, modules)
        
        print('\n接下來您可以：')
        print('1. 運行 python start_tronc.py')
        print('2. 或直接運行其他腳本（create_syllabus.py 等）')
        
    else:
        print('\n❌ 登入失敗')
        print('請檢查 log/ 目錄中的 HTML 文件來診斷問題')

if __name__ == "__main__":
    main()