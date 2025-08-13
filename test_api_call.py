#!/usr/bin/env python3
"""
test_api_call.py - 測試改進的AI轉換API調用方法
用於測試新的SSE流處理機制
"""

import json
import uuid
from tronc_login import setup_driver
from config import BASE_URL

def test_ai_convert_api():
    """測試AI轉換API調用"""
    print("=== 測試改進的AI轉換API調用 ===")
    
    # 測試參數
    upload_id = 12523  # 使用實際的upload_id
    belong = "library"
    belong_id = 356
    
    driver = setup_driver()
    if not driver:
        print("❌ 無法啟動瀏覽器")
        return False
    
    try:
        # 導航到匯入頁面
        test_url = f"{BASE_URL}/subject-lib/{belong_id}/import?mode=word"
        print(f"訪問測試頁面: {test_url}")
        driver.get(test_url)
        
        # 設置cookies
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
                print(f"添加cookie失敗 {name}: {e}")
        
        driver.refresh()
        import time
        time.sleep(3)
        
        # 生成唯一標記
        marker = f"selenium-{uuid.uuid4()}"
        print(f"使用標記: {marker}")
        
        # 構建JavaScript代碼
        js_code = f"""
        const done = arguments[0];
        (async () => {{
          try {{
            console.log('開始AI轉換API調用...');
            const resp = await fetch('/api/data-import/ai-convert', {{
              method: 'POST',
              headers: {{
                'Content-Type': 'application/json',
                'X-From-Selenium': '{marker}'   // 指紋標頭
              }},
              body: JSON.stringify({{
                upload_id: {upload_id},
                belong: '{belong}',
                belong_id: {belong_id}
              }})
            }});

            console.log('API響應狀態:', resp.status);

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

            // 檢查響應類型
            const ctype = resp.headers.get('content-type') || '';
            console.log('響應Content-Type:', ctype);
            
            if (ctype.includes('text/event-stream')) {{
              console.log('處理SSE流...');
              const reader = resp.body.getReader();
              const decoder = new TextDecoder();
              let full = '';
              let lastEvent = null;
              let eventCount = 0;
              let events = [];

              while (true) {{
                const {{ value, done: streamDone }} = await reader.read();
                if (streamDone) break;
                const chunk = decoder.decode(value, {{stream: true}});
                full += chunk;

                // 解析SSE格式
                const parts = full.split('\\n\\n');
                for (let part of parts) {{
                  if (part.startsWith('data:')) {{
                    const raw = part.replace(/^data:\\s*/, '').trim();
                    if (raw && raw !== '[DONE]') {{
                      lastEvent = raw;
                      eventCount++;
                      events.push(raw);
                      console.log('SSE事件:', raw);
                    }}
                  }}
                }}
              }}
              
              console.log('SSE流處理完成，總事件數:', eventCount);
              done({{
                ok: true,
                type: 'sse',
                marker: '{marker}',
                lastEvent: lastEvent,
                eventCount: eventCount,
                streamLength: full.length,
                firstFewEvents: events.slice(0, 3)  // 返回前3個事件作為樣本
              }});
              return;
            }} else {{
              console.log('處理JSON響應...');
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
            console.error('API調用錯誤:', err);
            done({{
              ok: false,
              error: String(err),
              marker: '{marker}'
            }});
          }}
        }})();
        """
        
        print("執行API調用...")
        
        # 設置超時時間
        driver.set_script_timeout(120)  # 2分鐘
        
        # 執行異步JavaScript
        result = driver.execute_async_script(js_code)
        
        print("=== API調用結果 ===")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        
        # 分析結果
        if result.get('ok'):
            print("✅ API調用成功")
            if result.get('type') == 'sse':
                print(f"📡 SSE流處理完成:")
                print(f"   - 事件數量: {result.get('eventCount', 0)}")
                print(f"   - 流大小: {result.get('streamLength', 0)} bytes")
                print(f"   - 最後事件: {result.get('lastEvent', 'None')}")
                if result.get('firstFewEvents'):
                    print("   - 前幾個事件:")
                    for i, event in enumerate(result.get('firstFewEvents', []), 1):
                        print(f"     {i}. {event[:100]}..." if len(event) > 100 else f"     {i}. {event}")
            elif result.get('type') == 'json':
                print(f"📄 JSON響應: {result.get('data')}")
            
            # API成功後，等待前端顯示結果
            print("\n⏳ API成功，等待前端顯示轉換結果...")
            frontend_ready = wait_for_frontend_results(driver)
            
            if frontend_ready:
                print("✅ 前端已顯示轉換結果")
                
                # 尋找識別按鈕
                print("🔍 尋找識別按鈕...")
                identify_found = find_identify_button(driver)
                
                if identify_found:
                    print("✅ 找到識別按鈕")
                    # 不自動點擊，讓用戶手動操作
                    print("💡 請手動點擊識別按鈕進行下一步操作")
                    input("按Enter鍵關閉瀏覽器...")
                else:
                    print("⚠️ 沒有找到識別按鈕")
                    input("按Enter鍵關閉瀏覽器...")
                
                return True
            else:
                print("❌ 前端未顯示轉換結果")
                input("按Enter鍵關閉瀏覽器...")
                return False
                
        else:
            print("❌ API調用失敗")
            print(f"錯誤: {result.get('error', result.get('statusText'))}")
            print(f"狀態: {result.get('status')}")
            return False
            
    except Exception as e:
        print(f"測試過程中發生錯誤: {e}")
        return False
    finally:
        driver.quit()

def wait_for_frontend_results(driver):
    """等待前端顯示AI轉換結果"""
    try:
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.by import By
        from selenium.common.exceptions import TimeoutException
        
        print("等待前端顯示轉換結果，最多等待60秒...")
        
        # 檢查前端結果顯示的指標
        result_indicators = [
            "//button[contains(text(), '識別')]",
            "//span[contains(text(), '識別')]",
            "//div[contains(text(), '轉換完成')]",
            "//div[contains(text(), '題目')]",
            "//div[contains(text(), '選項')]"
        ]
        
        wait = WebDriverWait(driver, 60, poll_frequency=3.0)
        
        for indicator in result_indicators:
            try:
                element = wait.until(
                    EC.presence_of_element_located((By.XPATH, indicator))
                )
                if element.is_displayed():
                    print(f"✅ 檢測到前端結果指標: {indicator}")
                    return True
            except TimeoutException:
                continue
        
        # 檢查頁面內容
        import time
        time.sleep(5)
        try:
            page_text = driver.find_element(By.TAG_NAME, "body").text
            result_keywords = ["識別", "轉換", "完成", "題目", "選項"]
            
            found_keywords = [keyword for keyword in result_keywords if keyword in page_text]
            if found_keywords:
                print(f"✅ 頁面包含結果關鍵字: {found_keywords}")
                return True
        except:
            pass
        
        print("⚠️ 沒有檢測到前端結果顯示")
        return False
        
    except Exception as e:
        print(f"等待前端結果時發生錯誤: {e}")
        return False

def find_identify_button(driver):
    """尋找識別按鈕"""
    try:
        from selenium.webdriver.common.by import By
        
        identify_selectors = [
            # 基於用戶提到的特定屬性
            "//*[@data-v-04cb9283 and contains(text(), '識別')]",
            "//*[contains(@data-v, '04cb9283')]//span[contains(text(), '識別')]",
            "//*[@data-v-04cb9283]//button[contains(text(), '識別')]",
            
            # 標準選擇器
            "//button[contains(text(), '識別')]",
            "//span[contains(text(), '識別')]",
            "//a[contains(text(), '識別')]",
            "//input[@type='button' and contains(@value, '識別')]"
        ]
        
        for selector in identify_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for elem in elements:
                    if elem.is_displayed() and elem.is_enabled():
                        print(f"找到識別按鈕: {selector} - 文字='{elem.text}'")
                        return True
            except:
                continue
        
        print("沒有找到識別按鈕")
        return False
        
    except Exception as e:
        print(f"尋找識別按鈕時發生錯誤: {e}")
        return False

def main():
    """主函數"""
    print("AI轉換API測試工具")
    print("=" * 40)
    
    try:
        success = test_ai_convert_api()
        
        if success:
            print("\n🎉 API測試成功完成！")
        else:
            print("\n❌ API測試失敗")
            
    except KeyboardInterrupt:
        print("\n用戶中斷測試")
    except Exception as e:
        print(f"\n❌ 測試過程中發生錯誤: {e}")

if __name__ == "__main__":
    main()