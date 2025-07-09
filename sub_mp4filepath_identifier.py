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
        
        # 修正：將日誌保存到 log 目錄中
        log_dir = Path("log")
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / log_filename
        
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
        
        # 記錄日誌檔案位置
        logger.info(f"=== MP4/MP3檔案路徑識別模組 日誌開始 ===")
        logger.info(f"日誌檔案位置: {log_path.absolute()}")
    
    return logger

# 初始化日誌記錄器
logger = setup_logger()

# 配置項
CONFIG = {
    # 路徑層數配置（從第幾層開始查找，向上無限制）
    'PATH_LEVELS': [2],  # 從第2層開始查找，例如：merged_projects/1-1 生活化的資料科學-OK/
    
    # m3u8匹配模式配置（改進版，避免提取到不完整的URL）
    'M3U8_PATTERNS': [
        r'(?:src|url)\s*[=:]\s*["\']([^"\']*\.m3u8)["\']',  # src或url屬性
        r'["\']([http|https]://[^"\']*\.m3u8)["\']',  # 完整HTTP URL
        r'["\']([^"\']*\.m3u8)["\']',  # 一般引號包圍
        r'(https?://[^\s<>"\']*\.m3u8)',  # 不被引號包圍的完整URL
    ],
    
    # mp4檔案擴展名配置（不區分大小寫）
    'MP4_EXTENSIONS': ['.mp4', '.avi', '.mov'],
    
    # mp3檔案擴展名配置（暫時為空）
    'MP3_EXTENSIONS': [],
    
    # HTML檔案擴展名配置（不區分大小寫）
    'HTML_EXTENSIONS': ['.html', '.htm']
}

def extract_filename_from_m3u8_url(m3u8_url):
    """從m3u8 URL中提取MP4檔案名（不含副檔名）"""
    try:
        # 解析URL
        parsed_url = urlparse(m3u8_url)
        
        # 取得路徑部分
        path = parsed_url.path
        
        # 查找路徑中的.mp4部分
        # 例如：/vod/_definst_/710058/800k/710058_01.mp4/playlist.m3u8
        # 我們要提取：710058_01
        
        # 方法1：使用正則表達式查找 .mp4 前的檔案名
        mp4_pattern = r'/([^/]+)\.mp4/'
        mp4_match = re.search(mp4_pattern, path)
        if mp4_match:
            filename = mp4_match.group(1)
            logger.info(f"使用正則表達式提取檔案名: {filename}")
            return filename
        
        # 方法2：如果正則失效，嘗試分割路徑查找包含.mp4的部分
        path_parts = path.split('/')
        for part in path_parts:
            if '.mp4' in part and not part.endswith('.m3u8'):
                # 移除.mp4副檔名
                filename = part.replace('.mp4', '')
                logger.info(f"使用路徑分割提取檔案名: {filename}")
                return filename
        
        # 方法3：如果以上都失效，提取URL中最後一個有意義的檔案名部分
        # 移除 /playlist.m3u8 等後綴
        clean_path = path.replace('/playlist.m3u8', '').replace('.m3u8', '')
        filename = os.path.basename(clean_path)
        if filename:
            logger.info(f"使用清理路徑提取檔案名: {filename}")
            return filename
            
        logger.warning(f"無法從URL提取檔案名: {m3u8_url}")
        return ""
        
    except Exception as e:
        logger.error(f"解析m3u8 URL時出錯: {e}, URL: {m3u8_url}")
        return ""

def find_m3u8_in_html(html_file_path):
    """在HTML檔案中查找m3u8路徑"""
    try:
        # 嘗試不同編碼讀取檔案
        for encoding in ['utf-8', 'gbk', 'big5', 'cp1252']:
            try:
                with open(html_file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                logger.info(f"成功使用 {encoding} 編碼讀取HTML檔案: {html_file_path}")
                break
            except UnicodeDecodeError:
                continue
        else:
            error_msg = f"無法讀取HTML檔案（嘗試多種編碼失敗）: {html_file_path}"
            print(f"  ⚠️  {error_msg}")
            logger.error(error_msg)
            return []
        
        # 使用配置的模式查找m3u8
        m3u8_urls = []
        for pattern in CONFIG['M3U8_PATTERNS']:
            matches = re.findall(pattern, content, re.IGNORECASE)
            for match in matches:
                if match not in m3u8_urls:
                    m3u8_urls.append(match)
                    logger.debug(f"使用模式 '{pattern}' 找到m3u8: {match}")
        
        logger.info(f"在HTML檔案中找到 {len(m3u8_urls)} 個m3u8路徑: {m3u8_urls}")
        return m3u8_urls
        
    except Exception as e:
        error_msg = f"讀取HTML檔案時出錯: {e}, 檔案: {html_file_path}"
        print(f"  ⚠️  {error_msg}")
        logger.error(error_msg)
        return []

def get_search_base_path(system_path, level):
    """根據指定層數獲取搜索基礎路徑（從第level層開始，向上無限制）"""
    try:
        path_parts = Path(system_path).parts
        if len(path_parts) >= level:
            # 從第level層開始的路徑，例如：level=2 → merged_projects/1-1 生活化的資料科學-OK/
            base_path = os.path.join(*path_parts[:level])
            logger.debug(f"層級 {level} 搜索路徑: {base_path}")
            return base_path
        else:
            # 如果層數不足，返回整個父目錄
            parent_path = str(Path(system_path).parent)
            logger.debug(f"層數不足，使用父目錄: {parent_path}")
            return parent_path
    except Exception as e:
        fallback_path = str(Path(system_path).parent)
        logger.warning(f"獲取搜索路徑時出錯: {e}, 使用備用路徑: {fallback_path}")
        return fallback_path

def find_media_file(base_path, filename):
    """在基礎路徑中查找媒體檔案（不區分大小寫）"""
    found_files = []
    
    try:
        # 遍歷所有層級配置
        for level in CONFIG['PATH_LEVELS']:
            search_path = get_search_base_path(base_path, level)
            
            if not os.path.exists(search_path):
                logger.warning(f"搜索路徑不存在: {search_path}")
                continue
                
            print(f"  📁 在路徑中搜索: {search_path}")
            logger.info(f"開始在路徑中搜索媒體檔案: {search_path}, 目標檔案名: {filename}")
            
            # 遞歸搜索檔案
            files_checked = 0
            for root, dirs, files in os.walk(search_path):
                for file in files:
                    files_checked += 1
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
                            logger.info(f"找到MP4檔案: {full_path}")
                        
                        # 檢查是否為mp3類型（不區分大小寫）
                        mp3_extensions_lower = [ext.lower() for ext in CONFIG['MP3_EXTENSIONS']]
                        if file_ext in mp3_extensions_lower:
                            found_files.append({
                                'type': 'mp3', 
                                'path': full_path,
                                'level': level
                            })
                            logger.info(f"找到MP3檔案: {full_path}")
            
            logger.info(f"層級 {level} 搜索完成，檢查了 {files_checked} 個檔案，找到 {len(found_files)} 個匹配檔案")
    
    except Exception as e:
        error_msg = f"搜索媒體檔案時出錯: {e}"
        print(f"  ⚠️  {error_msg}")
        logger.error(error_msg)
    
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
    
    logger.info(f"=== 開始處理HTML檔案 ===")
    logger.info(f"系統路徑: {system_path}")
    
    # 檢查是否為HTML檔案
    file_ext = os.path.splitext(system_path)[1].lower()
    html_extensions_lower = [ext.lower() for ext in CONFIG['HTML_EXTENSIONS']]
    if file_ext not in html_extensions_lower:
        logger.info(f"檔案不是HTML格式，跳過處理: {file_ext}")
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
    for i, m3u8_url in enumerate(m3u8_urls, 1):
        print(f"  🎬 處理m3u8 ({i}/{len(m3u8_urls)}): {m3u8_url}")
        logger.info(f"處理第 {i} 個m3u8 URL: {m3u8_url}")
        
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
            
            success_msg = f"成功找到媒體檔案 - 檔案名: {filename}.{best_file['type']}, 路徑: {best_file['path']}"
            print(f"  🎉 {success_msg}")
            logger.info(success_msg)
            logger.info(f"=== HTML檔案處理完成（成功）===")
            return result
        else:
            # 記錄未找到檔案的詳細信息
            error_msg = f"未找到媒體檔案 - 搜索檔案名: {filename}.mp4/.mp3, HTML: {system_path}, m3u8: {m3u8_url}"
            print(f"  ❌ 未找到檔案: {filename}.mp4/.mp3")
            logger.warning(error_msg)
    
    # 所有m3u8都處理完仍未找到檔案
    final_error = f"處理完所有m3u8路徑後未找到媒體檔案: {system_path}"
    print(f"  ❌ 未找到對應的媒體檔案")
    logger.warning(final_error)
    logger.info(f"=== HTML檔案處理完成（失敗）===")
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
        logger.debug(f"跳過處理：系統路徑為空 - 活動類型: {activity_type}")
        return False
    
    # 檢查檔案副檔名（不區分大小寫）
    file_ext = os.path.splitext(system_path)[1].lower()
    html_extensions_lower = [ext.lower() for ext in CONFIG['HTML_EXTENSIONS']]
    should_process = file_ext in html_extensions_lower
    
    logger.debug(f"活動處理判斷 - 類型: {activity_type}, 路徑: {system_path}, 副檔名: {file_ext}, 需要處理: {should_process}")
    return should_process

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
    
    logger.info(f"處理學習活動媒體識別 - 活動類型: {activity_type}, 系統路徑: {system_path}")
    
    # 檢查是否需要處理
    if not should_process_activity(activity_type, system_path):
        logger.info(f"學習活動不需要媒體檔案處理，返回預設結果")
        return default_result
    
    # 處理HTML檔案
    media_result = process_html_for_media(system_path)
    
    if media_result['found']:
        # 成功找到媒體檔案
        result = {
            'should_update': True,
            'new_activity_type': media_result['new_activity_type'],
            'new_file_path': media_result['file_path'],
            'use_web_path': False,  # 使用檔案路徑
            'fallback_to_online': False,
            'error_message': ''
        }
        logger.info(f"媒體檔案識別成功 - 新活動類型: {result['new_activity_type']}, 檔案路徑: {result['new_file_path']}")
        return result
    else:
        # 未找到媒體檔案，需要回退到線上連結
        result = {
            'should_update': True,
            'new_activity_type': '線上連結',  # 回退到線上連結
            'new_file_path': '',
            'use_web_path': True,  # 使用網址路徑
            'fallback_to_online': True,  # 標記為回退模式
            'error_message': media_result.get('error_message', '未找到對應媒體檔案')
        }
        logger.info(f"媒體檔案識別失敗，回退到線上連結模式: {system_path}, 錯誤: {result['error_message']}")
        return result