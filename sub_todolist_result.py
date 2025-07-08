"""
Result 資料處理模組 - 改寫版本
負責處理課程、章節、單元、學習活動的結構化資料生成
包含完善的父級查找邏輯和#Yes特殊處理
整合MP4/MP3檔案路徑識別功能
"""

import os
import pandas as pd
from sub_mp4filepath_identifier import process_activity_for_media

def find_parent_value(df, current_row, current_col, target_cols):
    """向左上方查找父級值"""
    for target_col in target_cols:
        if target_col < current_col and target_col >= 0:
            for row_idx in range(current_row, -1, -1):
                if target_col < len(df.columns):
                    value = str(df.iloc[row_idx, target_col]).strip()
                    if value and value != 'nan' and value != '' and not value.startswith('單元') and value not in ['課程名稱', '章節', '學習活動', '類型', '路徑', '系統路徑']:
                        return value, target_col
    return '', -1

def find_parent_value_upward(df, current_row, target_cols):
    """向上查找指定列的值"""
    for target_col in target_cols:
        if target_col >= 0:
            for row_idx in range(current_row, -1, -1):
                if target_col < len(df.columns):
                    value = str(df.iloc[row_idx, target_col]).strip()
                    if value and value != 'nan' and value != '' and not value.startswith('單元') and value not in ['課程名稱', '章節', '學習活動', '類型', '路徑', '系統路徑']:
                        return value, target_col
    return '', -1

def find_course_for_activity(df, current_row, header_info):
    """為學習活動查找所屬課程"""
    course_value, _ = find_parent_value_upward(df, current_row, [header_info['course_col']])
    return course_value

def determine_activity_parent(df, current_row, header_info, course_name, result_data):
    """
    學習活動父級判定邏輯：
    1. 如果上一行為課程 → 創建"章節N"作為父級
    2. 如果上一行為章節或單元 → 該章節或單元為父級  
    3. 如果上一行為"含章節單元的學習活動" → 該活動的章節或單元為父級
    4. 如果上一行為"純學習活動" → 繼續向上查找直到找到章節或單元
    """
    parent_chapter = ''
    parent_unit = ''
    
    # 向上查找，確定父級類型和內容
    for search_row in range(current_row - 1, -1, -1):
        # 檢查該行是否有課程
        if header_info['course_col'] >= 0 and header_info['course_col'] < len(df.columns):
            course_val = str(df.iloc[search_row, header_info['course_col']]).strip()
            if course_val and course_val != 'nan' and course_val != '':
                # 上一行是課程，需要創建自動章節
                return create_auto_chapter(course_name, result_data), ''
        
        # 檢查該行是否有章節
        if header_info['chapter_col'] >= 0 and header_info['chapter_col'] < len(df.columns):
            chapter_val = str(df.iloc[search_row, header_info['chapter_col']]).strip()
            if chapter_val and chapter_val != 'nan' and chapter_val != '':
                # 找到章節，這就是父級
                return chapter_val, ''
        
        # 檢查該行是否有單元
        for unit_col in header_info['unit_cols']:
            if unit_col < len(df.columns):
                unit_val = str(df.iloc[search_row, unit_col]).strip()
                if unit_val and unit_val != 'nan' and unit_val != '':
                    # 找到單元，需要找到該單元的章節作為父級
                    unit_chapter, _ = find_parent_value_upward(df, search_row, [header_info['chapter_col']])
                    return unit_chapter, unit_val
        
        # 檢查該行是否為學習活動（含章節單元的學習活動）
        if header_info['activity_col'] >= 0 and header_info['activity_col'] < len(df.columns):
            activity_val = str(df.iloc[search_row, header_info['activity_col']]).strip()
            if activity_val and activity_val != 'nan' and activity_val != '':
                # 這是學習活動行，需要查找該活動在result_data中的父級關係
                activity_parent = find_activity_parent_in_results(activity_val, result_data)
                if activity_parent:
                    return activity_parent['chapter'], activity_parent['unit']
                # 如果沒找到，繼續向上查找
                continue
    
    # 如果都沒找到，創建自動章節
    return create_auto_chapter(course_name, result_data), ''

def find_activity_parent_in_results(activity_name, result_data):
    """在result_data中查找學習活動的父級關係"""
    for item in reversed(result_data):  # 從後往前找最近的
        if item['類型'] == '學習活動' and item['名稱'] == activity_name:
            return {
                'chapter': item['所屬章節'],
                'unit': item['所屬單元']
            }
    return None

# 全域章節計數器
course_chapter_counter = {}

def create_auto_chapter(course_name, result_data):
    """為課程創建自動章節"""
    global course_chapter_counter
    
    if course_name not in course_chapter_counter:
        course_chapter_counter[course_name] = 1
    else:
        course_chapter_counter[course_name] += 1
    
    auto_chapter_name = f"章節{course_chapter_counter[course_name]}"
    
    # 添加章節到result_data
    result_data.append({
        '類型': '章節',
        '名稱': auto_chapter_name,
        'ID': '',
        '所屬課程': course_name,
        '所屬課程ID': '',
        '所屬章節': '',
        '所屬章節ID': '',
        '所屬單元': '',
        '所屬單元ID': '',
        '學習活動類型': '',
        '網址路徑': '',
        '檔案路徑': '',
        '資源ID': '',
        '最後修改時間': '',
        '來源Sheet': ''
    })
    
    return auto_chapter_name

def process_sheet_data(df, sheet_name, header_info):
    """處理單個 sheet 的資料"""
    global course_chapter_counter
    course_chapter_counter = {}  # 重置計數器
    
    result_data = []
    created_chapters = set()
    
    if header_info['header_row'] == -1:
        print(f"  ⚠️  Sheet '{sheet_name}' 沒有找到有效的表頭")
        return result_data
    
    # 從表頭行的下一行開始處理資料
    start_row = header_info['header_row'] + 1
    
    for row_idx in range(start_row, len(df)):
        row = df.iloc[row_idx]
        
        # 1. 處理課程
        if header_info['course_col'] >= 0 and header_info['course_col'] < len(row):
            course_name = str(row.iloc[header_info['course_col']]).strip()
            if course_name and course_name != 'nan' and course_name != '':
                result_data.append({
                    '類型': '課程',
                    '名稱': course_name,
                    'ID': '',
                    '所屬課程': '',
                    '所屬課程ID': '',
                    '所屬章節': '',
                    '所屬章節ID': '',
                    '所屬單元': '',
                    '所屬單元ID': '',
                    '學習活動類型': '',
                    '網址路徑': '',
                    '檔案路徑': '',
                    '資源ID': '',
                    '最後修改時間': '',
                    '來源Sheet': sheet_name
                })
        
        # 2. 處理章節
        if header_info['chapter_col'] >= 0 and header_info['chapter_col'] < len(row):
            chapter_name = str(row.iloc[header_info['chapter_col']]).strip()
            if chapter_name and chapter_name != 'nan' and chapter_name != '':
                # 查找所屬課程
                parent_course = find_course_for_activity(df, row_idx, header_info)
                
                # 避免重複產生同名章節
                if (parent_course, chapter_name) not in created_chapters:
                    result_data.append({
                        '類型': '章節',
                        '名稱': chapter_name,
                        'ID': '',
                        '所屬課程': parent_course,
                        '所屬課程ID': '',
                        '所屬章節': '',
                        '所屬章節ID': '',
                        '所屬單元': '',
                        '所屬單元ID': '',
                        '學習活動類型': '',
                        '網址路徑': '',
                        '檔案路徑': '',
                        '資源ID': '',
                        '最後修改時間': '',
                        '來源Sheet': sheet_name
                    })
                    created_chapters.add((parent_course, chapter_name))
        
        # 3. 處理單元
        for unit_col in header_info['unit_cols']:
            if unit_col < len(row):
                unit_name = str(row.iloc[unit_col]).strip()
                if unit_name and unit_name != 'nan' and unit_name != '':
                    # 查找所屬章節和課程
                    parent_chapter, _ = find_parent_value(df, row_idx, unit_col, [header_info['chapter_col']])
                    parent_course = find_course_for_activity(df, row_idx, header_info)
                    
                    result_data.append({
                        '類型': '單元',
                        '名稱': unit_name,
                        'ID': '',
                        '所屬課程': parent_course,
                        '所屬課程ID': '',
                        '所屬章節': parent_chapter,
                        '所屬章節ID': '',
                        '所屬單元': '',
                        '所屬單元ID': '',
                        '學習活動類型': '',
                        '網址路徑': '',
                        '檔案路徑': '',
                        '資源ID': '',
                        '最後修改時間': '',
                        '來源Sheet': sheet_name
                    })
        
        # 4. 處理學習活動
        activity_name = ''
        activity_col_valid = header_info['activity_col'] >= 0 and header_info['activity_col'] < len(row)
        if activity_col_valid:
            activity_name = str(row.iloc[header_info['activity_col']]).strip()
        
        # 如果學習活動欄位為空，但章節欄位有值，且該行有其他活動資訊，則使用章節欄位作為學習活動
        if (not activity_name or activity_name == 'nan' or activity_name == '') and header_info['chapter_col'] >= 0:
            chapter_value = str(row.iloc[header_info['chapter_col']]).strip()
            # 檢查該行是否有活動相關資訊
            has_activity_info = False
            if header_info['type_col'] >= 0 and header_info['type_col'] < len(row):
                type_value = str(row.iloc[header_info['type_col']]).strip()
                if type_value and type_value != 'nan' and type_value != '':
                    has_activity_info = True
            if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
                path_value = str(row.iloc[header_info['path_col']]).strip()
                if path_value and path_value != 'nan' and path_value != '':
                    has_activity_info = True
            
            if has_activity_info and chapter_value and chapter_value != 'nan' and chapter_value != '' and chapter_value != '章節':
                activity_name = chapter_value
        
        if activity_name and activity_name != 'nan' and activity_name != '':
            # 查找所屬課程
            parent_course = find_course_for_activity(df, row_idx, header_info)
            
            # === #Yes 邏輯 ===
            if activity_name == "#Yes":
                # 查找章節或單元欄位的值
                target_parent_name = ''
                target_parent_type = ''
                
                # 先檢查章節欄位
                if header_info['chapter_col'] >= 0 and header_info['chapter_col'] < len(row):
                    chapter_name = str(row.iloc[header_info['chapter_col']]).strip()
                    if chapter_name and chapter_name != 'nan' and chapter_name != '' and chapter_name != '章節':
                        target_parent_name = chapter_name
                        target_parent_type = '章節'
                
                # 如果章節欄位沒有值，檢查單元欄位
                if not target_parent_name:
                    for unit_col in header_info['unit_cols']:
                        if unit_col < len(row):
                            unit_name = str(row.iloc[unit_col]).strip()
                            if unit_name and unit_name != 'nan' and unit_name != '':
                                target_parent_name = unit_name
                                target_parent_type = '單元'
                                break
                
                if target_parent_name:
                    # 1. 先創建父級（章節或單元）
                    if target_parent_type == '章節':
                        if (parent_course, target_parent_name) not in created_chapters:
                            result_data.append({
                                '類型': '章節',
                                '名稱': target_parent_name,
                                'ID': '',
                                '所屬課程': parent_course,
                                '所屬課程ID': '',
                                '所屬章節': '',
                                '所屬章節ID': '',
                                '所屬單元': '',
                                '所屬單元ID': '',
                                '學習活動類型': '',
                                '網址路徑': '',
                                '檔案路徑': '',
                                '資源ID': '',
                                '最後修改時間': '',
                                '來源Sheet': sheet_name
                            })
                            created_chapters.add((parent_course, target_parent_name))
                        parent_chapter = target_parent_name
                        parent_unit = ''
                    else:  # 單元
                        parent_chapter, _ = find_parent_value_upward(df, row_idx, [header_info['chapter_col']])
                        result_data.append({
                            '類型': '單元',
                            '名稱': target_parent_name,
                            'ID': '',
                            '所屬課程': parent_course,
                            '所屬課程ID': '',
                            '所屬章節': parent_chapter,
                            '所屬章節ID': '',
                            '所屬單元': '',
                            '所屬單元ID': '',
                            '學習活動類型': '',
                            '網址路徑': '',
                            '檔案路徑': '',
                            '資源ID': '',
                            '最後修改時間': '',
                            '來源Sheet': sheet_name
                        })
                        parent_unit = target_parent_name
                    
                    # 2. 再創建學習活動，名稱為父級名稱
                    final_activity_name = target_parent_name
                else:
                    # 如果沒有找到目標父級，跳過這個#Yes
                    continue
            
            # === 普通學習活動邏輯 ===
            else:
                # 使用新的父級判定邏輯
                parent_chapter, parent_unit = determine_activity_parent(df, row_idx, header_info, parent_course, result_data)
                final_activity_name = activity_name
            
            # 獲取活動類型和路徑
            activity_type = ''
            web_path = ''
            file_path = ''
            system_path = ''
            
            # 獲取系統路徑（用於MP4/MP3識別）
            if header_info['system_path_col'] >= 0 and header_info['system_path_col'] < len(row):
                system_path = str(row.iloc[header_info['system_path_col']]).strip()
                if system_path == 'nan':
                    system_path = ''
            
            if header_info['type_col'] >= 0 and header_info['type_col'] < len(row):
                activity_type = str(row.iloc[header_info['type_col']]).strip()
                if activity_type == 'nan':
                    activity_type = ''
            
            # === 新增：MP4/MP3檔案識別處理 ===
            media_result = process_activity_for_media(activity_type, system_path)
            
            if media_result['should_update']:
                # 更新活動類型
                activity_type = media_result['new_activity_type']
                
                if media_result['fallback_to_online']:
                    # 回退到線上連結模式：使用網址路徑
                    if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
                        path_value = str(row.iloc[header_info['path_col']]).strip()
                        if path_value and path_value != 'nan':
                            web_path = path_value
                    file_path = ''
                    print(f"  ⚠️  媒體檔案識別失敗，回退到線上連結: {final_activity_name}")
                    print(f"      錯誤訊息: {media_result['error_message']}")
                elif not media_result['use_web_path']:
                    # 成功找到媒體檔案：使用檔案路徑
                    file_path = media_result['new_file_path']
                    web_path = ''
                    print(f"  🎬 媒體檔案識別成功: {final_activity_name} → {activity_type}")
                else:
                    # 按原邏輯處理路徑
                    if activity_type in ['線上連結', '影音連結']:
                        if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
                            path_value = str(row.iloc[header_info['path_col']]).strip()
                            if path_value and path_value != 'nan':
                                web_path = path_value
                    elif activity_type in ['參考檔案_圖片', '參考檔案_PDF']:
                        if system_path:
                            file_path = system_path
            else:
                # === 原有的路徑處理邏輯 ===
                if activity_type in ['線上連結', '影音連結']:
                    if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
                        path_value = str(row.iloc[header_info['path_col']]).strip()
                        if path_value and path_value != 'nan':
                            web_path = path_value
                elif activity_type in ['參考檔案_圖片', '參考檔案_PDF']:
                    if system_path:
                        file_path = system_path
            
            # 確保最後設定來源Sheet
            if 'final_activity_name' in locals():
                result_data.append({
                    '類型': '學習活動',
                    '名稱': final_activity_name,
                    'ID': '',
                    '所屬課程': parent_course,
                    '所屬課程ID': '',
                    '所屬章節': parent_chapter,
                    '所屬章節ID': '',
                    '所屬單元': parent_unit,
                    '所屬單元ID': '',
                    '學習活動類型': activity_type,
                    '網址路徑': web_path,
                    '檔案路徑': file_path,
                    '資源ID': '',
                    '最後修改時間': '',
                    '來源Sheet': sheet_name
                })
    
    return result_data