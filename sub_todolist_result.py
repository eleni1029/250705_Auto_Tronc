"""
Result 資料處理模組
負責處理課程、章節、單元、學習活動的結構化資料生成
包含所有相關的工具函數，方便迭代和調整層級結構邏輯
"""

import os
import pandas as pd

def find_parent_value(df, current_row, current_col, target_cols):
    """向左上方查找父級值"""
    # 向左查找第一個有值的目標欄位
    for target_col in target_cols:
        if target_col < current_col and target_col >= 0:
            # 向上查找該欄位的最近非空值
            for row_idx in range(current_row, -1, -1):
                if target_col < len(df.columns):
                    value = str(df.iloc[row_idx, target_col]).strip()
                    # 排除欄位名稱本身（如"單元1"、"課程名稱"等）
                    if value and value != 'nan' and value != '' and not value.startswith('單元') and value not in ['課程名稱', '章節', '學習活動', '類型', '路徑', '系統路徑']:
                        return value, target_col
    return '', -1

def find_parent_value_in_chapter(df, current_row, current_col, target_cols, chapter_name, chapter_col):
    """在指定章節範圍內向左上方查找父級值"""
    if not chapter_name:
        # 如果沒有章節名稱，回退到普通查找
        return find_parent_value(df, current_row, current_col, target_cols)
    
    # 向左查找第一個有值的目標欄位
    for target_col in target_cols:
        if target_col < current_col and target_col >= 0:
            # 向上查找該欄位的最近非空值，但確保在相同章節內
            for row_idx in range(current_row, -1, -1):
                if target_col < len(df.columns) and chapter_col < len(df.columns):
                    # 檢查當前行的章節是否與目標章節相同
                    current_chapter = str(df.iloc[row_idx, chapter_col]).strip()
                    if current_chapter == chapter_name:
                        # 在相同章節內查找目標值
                        value = str(df.iloc[row_idx, target_col]).strip()
                        # 排除欄位名稱本身
                        if value and value != 'nan' and value != '' and not value.startswith('單元') and value not in ['課程名稱', '章節', '學習活動', '類型', '路徑', '系統路徑']:
                            return value, target_col
                    elif current_chapter and current_chapter != 'nan' and current_chapter != '':
                        # 如果遇到不同的章節，停止查找
                        break
    return '', -1

def find_parent_value_upward(df, current_row, target_cols):
    """向上查找指定列的值"""
    for target_col in target_cols:
        if target_col >= 0:
            # 向上查找該欄位的最近非空值
            for row_idx in range(current_row, -1, -1):
                if target_col < len(df.columns):
                    value = str(df.iloc[row_idx, target_col]).strip()
                    # 排除欄位名稱本身
                    if value and value != 'nan' and value != '' and not value.startswith('單元') and value not in ['課程名稱', '章節', '學習活動', '類型', '路徑', '系統路徑']:
                        return value, target_col
    return '', -1

def find_parent_in_previous_row(df, current_row, header_info):
    """向上查找第一個非學習活動行，然後向左找到第一個有值的單元或章節"""
    if current_row <= 0:
        return '', '', -1
    
    # 向上查找，跳過所有學習活動行
    target_row = current_row - 1
    while target_row >= 0:
        # 檢查當前行是否為學習活動
        if header_info['activity_col'] < len(df.columns):
            activity_value = str(df.iloc[target_row, header_info['activity_col']]).strip()
            if activity_value and activity_value != 'nan' and activity_value != '':
                # 這是學習活動行，繼續向上查找
                target_row -= 1
                continue
        
        # 找到非學習活動行，開始向左查找
        break
    
    if target_row < 0:
        return '', '', -1
    
    # 獲取目標行
    target_row_data = df.iloc[target_row]
    
    # 從學習活動列開始向左查找
    start_col = header_info['activity_col']
    
    # 先查找單元（優先級更高）
    for unit_col in header_info['unit_cols']:
        if unit_col < start_col and unit_col < len(target_row_data):
            unit_value = str(target_row_data.iloc[unit_col]).strip()
            if unit_value and unit_value != 'nan' and unit_value != '':
                # 找到單元，再向上查找章節
                parent_chapter, _ = find_parent_value_upward(df, target_row, [header_info['chapter_col']])
                return unit_value, parent_chapter, unit_col
    
    # 如果沒有找到單元，查找章節
    if header_info['chapter_col'] < start_col and header_info['chapter_col'] < len(target_row_data):
        chapter_value = str(target_row_data.iloc[header_info['chapter_col']]).strip()
        if chapter_value and chapter_value != 'nan' and chapter_value != '':
            return '', chapter_value, -1
    
    return '', '', -1

def process_sheet_data(df, sheet_name, header_info):
    """處理單個 sheet 的資料"""
    result_data = []
    
    if header_info['header_row'] == -1:
        print(f"  ⚠️  Sheet '{sheet_name}' 沒有找到有效的表頭")
        return result_data
    
    # 從表頭行的下一行開始處理資料
    start_row = header_info['header_row'] + 1
    
    # 用於追蹤每個課程的章節編號
    course_chapter_counter = {}
    # 用於追蹤每個課程當前正在使用的章節
    course_current_chapter = {}
    created_chapters = set()
    
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
                parent_course, _ = find_parent_value(df, row_idx, header_info['chapter_col'], [header_info['course_col']])
                # 重置該課程的當前章節狀態，因為遇到了新的章節
                if parent_course and parent_course in course_current_chapter:
                    del course_current_chapter[parent_course]
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
                    # 查找所屬章節（向左查找章節欄位，然後向上查找）
                    parent_chapter, _ = find_parent_value(df, row_idx, unit_col, [header_info['chapter_col']])
                    # 查找所屬課程
                    parent_course, _ = find_parent_value(df, row_idx, unit_col, [header_info['course_col']])
                    
                    # 重置該課程的當前章節狀態，因為遇到了新的單元
                    if parent_course and parent_course in course_current_chapter:
                        del course_current_chapter[parent_course]
                    
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
        
        # 如果學習活動欄位為空，但章節欄位有值，且該行有其他活動資訊（如類型、路徑），則使用章節欄位作為學習活動
        if (not activity_name or activity_name == 'nan' or activity_name == '') and header_info['chapter_col'] >= 0:
            chapter_value = str(row.iloc[header_info['chapter_col']]).strip()
            # 檢查該行是否有活動相關資訊（類型、路徑等）
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
                parent_course, _ = find_parent_value(df, row_idx, header_info['activity_col'], [header_info['course_col']])
                
                # 查找父級單元和章節
                parent_unit = ''
                parent_chapter = ''
                unit_col = -1
                
                # 首先嘗試在當前行查找單元
                for ucol in header_info['unit_cols']:
                    if ucol < len(row):
                        unit_value = str(row.iloc[ucol]).strip()
                        if unit_value and unit_value != 'nan' and unit_value != '':
                            parent_unit = unit_value
                            unit_col = ucol
                            break
                
                # 如果當前行沒有單元，向上找一行，然後向左找到第一個有值的單元或章節
                if not parent_unit:
                    parent_unit, parent_chapter, unit_col = find_parent_in_previous_row(df, row_idx, header_info)
                
                # === 邏輯1: #Yes 邏輯 ===
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
                        final_activity_name = activity_name
                
                # === 邏輯2: 普通學習活動 ===
                else:
                    # 查找父級章節和單元
                    if not parent_chapter:
                        for row_idx_up in range(row_idx - 1, -1, -1):
                            if header_info['chapter_col'] < len(df.columns):
                                chapter_value = str(df.iloc[row_idx_up, header_info['chapter_col']]).strip()
                                if (chapter_value and chapter_value != 'nan' and chapter_value != '' and 
                                    chapter_value != '#Yes' and chapter_value != '章節' and 
                                    chapter_value not in ['課程名稱', '章節', '學習活動', '類型', '路徑', '系統路徑']):
                                    parent_chapter = chapter_value
                                    break
                    
                    # 如果沒有父級章節且沒有父級單元，且父級是課程，則動態創建章節
                    if parent_course and not parent_chapter and not parent_unit:
                        if parent_course not in course_current_chapter:
                            if parent_course not in course_chapter_counter:
                                course_chapter_counter[parent_course] = 1
                            else:
                                course_chapter_counter[parent_course] += 1
                            auto_chapter_name = f"章節{course_chapter_counter[parent_course]}"
                            course_current_chapter[parent_course] = auto_chapter_name
                            result_data.append({
                                '類型': '章節',
                                '名稱': auto_chapter_name,
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
                        parent_chapter = course_current_chapter[parent_course]
                    
                    final_activity_name = activity_name

                # 獲取活動類型和路徑
                activity_type = ''
                web_path = ''
                file_path = ''
                if header_info['type_col'] >= 0 and header_info['type_col'] < len(row):
                    activity_type = str(row.iloc[header_info['type_col']]).strip()
                    if activity_type == 'nan':
                        activity_type = ''
                if activity_type in ['線上連結', '影音連結']:
                    if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
                        path_value = str(row.iloc[header_info['path_col']]).strip()
                        if path_value and path_value != 'nan':
                            web_path = path_value
                elif activity_type in ['參考檔案_圖片', '參考檔案_PDF']:
                    if header_info['system_path_col'] >= 0 and header_info['system_path_col'] < len(row):
                        path_value = str(row.iloc[header_info['system_path_col']]).strip()
                        if path_value and path_value != 'nan':
                            file_path = path_value
                
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