"""
MP4/MP3檔案路徑識別模組
負責從HTML檔案中提取m3u8路徑，查找對應的mp4/mp3檔案，並更新活動類型和路徑
包含完整的日誌記錄功能
"""

import os
import re
import logging
from urllib.parse import urlparse
from pathlib import Path
from datetime import datetime

# 設置日誌
def setup_logger():
    """設置日誌記錄器"""
    logger = logging.getLogger('mp4_identifier')
    if not logger.handlers:
        # 創建日誌檔案名（包含時間戳）
        log_filename = f"mp4_identifier_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path = os.path.join('to_be_executed', log_filename)
        
        # 確保目錄存在
        os.makedirs('to_be_executed', exist_ok=True)
        
        # 設置日誌格式
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        
        # 文件處理器
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.INFO)
        
        # 控制台處理器
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.setLevel(logging.WARNING)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        logger.setLevel(logging.INFO)
    
    return logger

# 初始化日誌記錄器
logger = setup_logger()

# 配置項
CONFIG = {
    # 路徑層數配置（從第幾層開始查找，向上無限制）
    'PATH_LEVELS': [2],  # 從第2層開始查找，例如：merged_projects/1-1 生活化的資料科學-OK/
    
    # m3u8匹配模式配置
    'M3U8_PATTERNS': [
        r'["\']([^"\']*\.m3u8[^"\']*)["\']',  # 標準引號包圍
        r'src\s*=\s*["\']([^"\']*\.m3u8[^"\']*)["\']',  # src屬性
        r'url\s*:\s*["\']([^"\']*\.m3u8[^"\']*)["\']',  # url屬性
        r'([http|https]://[^\s]*\.m3u8[^\s]*)',  # 直接URL匹配
    ],
    
    # mp4檔案擴展名配置（不區分大小寫）
    'MP4_EXTENSIONS': ['.mp4', '.avi', '.mov'],
    
    # mp3檔案擴展名配置（暫時為空）
    'MP3_EXTENSIONS': [],
    
    # HTML檔案擴展名配置（不區分大小寫）
    'HTML_EXTENSIONS': ['.html', '.htm']
}

def extract_filename_from_m3u8_url(m3u8_url):
    """從m3u8 URL中提取檔案名（不含副檔名）"""
    try:
        # 解析URL
        parsed_url = urlparse(m3u8_url)
        
        # 取得路徑部分
        path = parsed_url.path
        
        # 提取檔案名
        filename = os.path.basename(path)
        
        # 移除.m3u8副檔名
        if filename.lower().endswith('.m3u8'):
            filename = filename[:-5]  # 移除.m3u8
        
        return filename
    except Exception as e:
        print(f"  ⚠️  解析m3u8 URL時出錯: {e}")
        return ""

def find_m3u8_in_html(html_file_path):
    """在HTML檔案中查找m3u8路徑"""
    try:
        # 嘗試不同編碼讀取檔案
        for encoding in ['utf-8', 'gbk', 'big5', 'cp1252']:
            try:
                with open(html_file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        else:
            print(f"  ⚠️  無法讀取HTML檔案: {html_file_path}")
            return []
        
        # 使用配置的模式查找m3u8
        m3u8_urls = []
        for pattern in CONFIG['M3U8_PATTERNS']:
            matches = re.findall(pattern, content, re.IGNORECASE)
            for match in matches:
                if match not in m3u8_urls:
                    m3u8_urls.append(match)
        
        return m3u8_urls
        
    except Exception as e:
        print(f"  ⚠️  讀取HTML檔案時出錯: {e}")
        return []

def get_search_base_path(system_path, level):
    """根據指定層數獲取搜索基礎路徑（從第level層開始，向上無限制）"""
    try:
        path_parts = Path(system_path).parts
        if len(path_parts) >= level:
            # 從第level層開始的路徑，例如：level=2 → merged_projects/1-1 生活化的資料科學-OK/
            return os.path.join(*path_parts[:level])
        else:
            # 如果層數不足，返回整個父目錄
            return str(Path(system_path).parent)
    except Exception:
        return str(Path(system_path).parent)

def find_media_file(base_path, filename):
    """在基礎路徑中查找媒體檔案（不區分大小寫）"""
    found_files = []
    
    try:
        # 遍歷所有層級配置
        for level in CONFIG['PATH_LEVELS']:
            search_path = get_search_base_path(base_path, level)
            
            if not os.path.exists(search_path):
                continue
                
            print(f"  📁 在路徑中搜索: {search_path}")
            
            # 遞歸搜索檔案
            for root, dirs, files in os.walk(search_path):
                for file in files:
                    file_base = os.path.splitext(file)[0]
                    file_ext = os.path.splitext(file)[1].lower()  # 轉小寫比較
                    
                    # 檢查檔案名是否匹配（不區分大小寫）
                    if file_base.lower() == filename.lower():
                        full_path = os.path.join(root, file)
                        
                        # 檢查是否為mp4類型（不區分大小寫）
                        mp4_extensions_lower = [ext.lower() for ext in CONFIG['MP4_EXTENSIONS']]
                        if file_ext in mp4_extensions_lower:
                            found_files.append({
                                'type': 'mp4',
                                'path': full_path,
                                'level': level
                            })
                        
                        # 檢查是否為mp3類型（不區分大小寫）
                        mp3_extensions_lower = [ext.lower() for ext in CONFIG['MP3_EXTENSIONS']]
                        if file_ext in mp3_extensions_lower:
                            found_files.append({
                                'type': 'mp3', 
                                'path': full_path,
                                'level': level
                            })
    
    except Exception as e:
        print(f"  ⚠️  搜索媒體檔案時出錯: {e}")
    
    return found_files

def process_html_for_media(system_path):
    """
    處理HTML檔案查找媒體檔案
    
    Args:
        system_path: HTML檔案的系統路徑
        
    Returns:
        dict: {
            'found': bool,
            'type': 'mp4'/'mp3'/None,
            'new_activity_type': str,
            'file_path': str,
            'error_message': str
        }
    """
    result = {
        'found': False,
        'type': None,
        'new_activity_type': '',
        'file_path': '',
        'error_message': ''
    }
    
    # 檢查是否為HTML檔案
    file_ext = os.path.splitext(system_path)[1].lower()
    html_extensions_lower = [ext.lower() for ext in CONFIG['HTML_EXTENSIONS']]
    if file_ext not in html_extensions_lower:
        return result
    
    # 檢查檔案是否存在
    if not os.path.exists(system_path):
        error_msg = f"HTML檔案不存在: {system_path}"
        print(f"  ⚠️  {error_msg}")
        logger.warning(error_msg)
        result['error_message'] = error_msg
        return result
    
    print(f"  🔍 處理HTML檔案: {os.path.basename(system_path)}")
    logger.info(f"開始處理HTML檔案: {system_path}")
    
    # 1. 在HTML中查找m3u8路徑
    m3u8_urls = find_m3u8_in_html(system_path)
    if not m3u8_urls:
        error_msg = f"未找到m3u8路徑: {system_path}"
        print(f"  ℹ️  未找到m3u8路徑")
        logger.info(error_msg)
        result['error_message'] = error_msg
        return result
    
    print(f"  ✅ 找到 {len(m3u8_urls)} 個m3u8路徑")
    logger.info(f"找到 {len(m3u8_urls)} 個m3u8路徑: {m3u8_urls}")
    
    # 2. 處理每個m3u8 URL
    for m3u8_url in m3u8_urls:
        print(f"  🎬 處理m3u8: {m3u8_url}")
        
        # 提取檔案名
        filename = extract_filename_from_m3u8_url(m3u8_url)
        if not filename:
            error_msg = f"無法從m3u8 URL提取檔案名: {m3u8_url}"
            logger.warning(error_msg)
            continue
            
        print(f"  📂 提取檔案名: {filename}")
        logger.info(f"提取檔案名: {filename}")
        
        # 3. 查找對應的媒體檔案
        found_files = find_media_file(system_path, filename)
        
        if found_files:
            # 取第一個找到的檔案（優先級：層數小的優先）
            best_file = sorted(found_files, key=lambda x: x['level'])[0]
            
            result['found'] = True
            result['type'] = best_file['type']
            result['file_path'] = best_file['path']
            
            # 設定新的活動類型
            if best_file['type'] == 'mp4':
                result['new_activity_type'] = '影音教材_影片'
            elif best_file['type'] == 'mp3':
                result['new_activity_type'] = '影音教材_音訊'
            
            success_msg = f"找到媒體檔案: {best_file['path']} (類型: {best_file['type']})"
            print(f"  🎉 {success_msg}")
            logger.info(f"成功 - {success_msg}")
            return result
        else:
            # 記錄未找到檔案的詳細信息
            error_msg = f"未找到對應媒體檔案 - HTML: {system_path}, 檔案名: {filename}, m3u8: {m3u8_url}"
            logger.warning(error_msg)
    
    # 所有m3u8都處理完仍未找到檔案
    final_error = f"處理完所有m3u8路徑後未找到媒體檔案: {system_path}"
    print(f"  ❌ 未找到對應的媒體檔案")
    logger.warning(final_error)
    result['error_message'] = final_error
    return result

def should_process_activity(activity_type, system_path):
    """
    判斷是否需要處理該學習活動
    
    Args:
        activity_type: 活動類型
        system_path: 系統路徑
        
    Returns:
        bool: 是否需要處理
    """
    # 檢查是否有系統路徑
    if not system_path or system_path.strip() == '' or str(system_path) == 'nan':
        return False
    
    # 檢查檔案副檔名（不區分大小寫）
    file_ext = os.path.splitext(system_path)[1].lower()
    html_extensions_lower = [ext.lower() for ext in CONFIG['HTML_EXTENSIONS']]
    return file_ext in html_extensions_lower

def process_activity_for_media(activity_type, system_path):
    """
    為學習活動處理媒體檔案識別
    
    Args:
        activity_type: 原始活動類型
        system_path: 系統路徑
        
    Returns:
        dict: {
            'should_update': bool,
            'new_activity_type': str,
            'new_file_path': str,
            'use_web_path': bool,
            'fallback_to_online': bool,  # 新增：是否需要回退到線上連結
            'error_message': str         # 新增：錯誤訊息
        }
    """
    default_result = {
        'should_update': False,
        'new_activity_type': activity_type,
        'new_file_path': '',
        'use_web_path': True,  # 預設使用網址路徑
        'fallback_to_online': False,
        'error_message': ''
    }
    
    # 檢查是否需要處理
    if not should_process_activity(activity_type, system_path):
        return default_result
    
    # 處理HTML檔案
    media_result = process_html_for_media(system_path)
    
    if media_result['found']:
        # 成功找到媒體檔案
        return {
            'should_update': True,
            'new_activity_type': media_result['new_activity_type'],
            'new_file_path': media_result['file_path'],
            'use_web_path': False,  # 使用檔案路徑
            'fallback_to_online': False,
            'error_message': ''
        }
    else:
        # 未找到媒體檔案，需要回退到線上連結
        logger.info(f"媒體檔案識別失敗，回退到線上連結模式: {system_path}")
        return {
            'should_update': True,
            'new_activity_type': '線上連結',  # 回退到線上連結
            'new_file_path': '',
            'use_web_path': True,  # 使用網址路徑
            'fallback_to_online': True,  # 標記為回退模式
            'error_message': media_result.get('error_message', '未找到對應媒體檔案')
        }