import os
import glob
import pandas as pd
from datetime import datetime
import re

def get_analyzed_files():
    """獲取 to_be_executed 目錄中所有 analyzed 檔案，按時間排序"""
    pattern = os.path.join('to_be_executed', 'course_structures_analyzed_*.xlsx')
    files = glob.glob(pattern)
    
    if not files:
        print("❌ 在 to_be_executed 目錄中找不到 analyzed 檔案")
        return []
    
    # 按修改時間排序（最新的在前）
    files.sort(key=os.path.getmtime, reverse=True)
    return files

def extract_timestamp(filename):
    """從檔案名中提取時間戳"""
    match = re.search(r'analyzed_(\d{8}_\d{6})', filename)
    if match:
        return match.group(1)
    return "unknown"

def select_file(files):
    """讓用戶選擇檔案"""
    print("\n📁 找到以下 analyzed 檔案（按時間排序，最新的在前）：")
    for i, file in enumerate(files, 1):
        timestamp = extract_timestamp(file)
        print(f"{i}. {os.path.basename(file)} (時間戳: {timestamp})")
    
    while True:
        try:
            choice = input(f"\n請選擇檔案 (1-{len(files)})，或按 Enter 選擇最新的: ").strip()
            if not choice:
                return files[0]  # 預設選擇最新的
            
            choice_num = int(choice)
            if 1 <= choice_num <= len(files):
                return files[choice_num - 1]
            else:
                print(f"❌ 請輸入 1-{len(files)} 之間的數字")
        except ValueError:
            print("❌ 請輸入有效的數字")

def get_sheet_names(file_path):
    """獲取 Excel 檔案中的所有 sheet 名稱"""
    try:
        excel_file = pd.ExcelFile(file_path)
        return excel_file.sheet_names
    except Exception as e:
        print(f"❌ 讀取 Excel 檔案時出錯: {e}")
        return []

def select_sheet(sheet_names):
    """讓用戶選擇 sheet"""
    print(f"\n📋 檔案包含以下 {len(sheet_names)} 個 sheet:")
    for i, sheet in enumerate(sheet_names, 1):
        print(f"{i}. {sheet}")
    
    while True:
        try:
            choice = input(f"\n請選擇 sheet (1-{len(sheet_names)})，或輸入 'all' 選擇全部: ").strip().lower()
            
            if choice == 'all':
                return sheet_names
            
            choice_num = int(choice)
            if 1 <= choice_num <= len(sheet_names):
                return [sheet_names[choice_num - 1]]
            else:
                print(f"❌ 請輸入 1-{len(sheet_names)} 之間的數字，或輸入 'all'")
        except ValueError:
            print("❌ 請輸入有效的數字或 'all'")

def find_header_positions(df):
    """動態查找表頭位置"""
    header_info = {
        'header_row': -1,
        'course_col': -1,
        'chapter_col': -1,
        'unit_cols': [],
        'activity_col': -1,
        'type_col': -1,
        'path_col': -1,
        'system_path_col': -1  # 新增系統路徑欄位
    }
    
    # 查找表頭行（包含"課程名稱"的行）
    for row_idx in range(min(15, len(df))):
        row_values = df.iloc[row_idx].astype(str).str.strip()
        if '課程名稱' in row_values.values:
            header_info['header_row'] = row_idx
            break
    
    if header_info['header_row'] == -1:
        print("❌ 找不到包含'課程名稱'的表頭行")
        return header_info
    
    # 查找各欄位位置
    header_row = df.iloc[header_info['header_row']].astype(str).str.strip()
    
    for col_idx, col_value in enumerate(header_row):
        if col_value == '課程名稱':
            header_info['course_col'] = col_idx
        elif col_value == '章節':
            header_info['chapter_col'] = col_idx
        elif col_value.startswith('單元'):
            header_info['unit_cols'].append(col_idx)
        elif col_value == '學習活動':
            header_info['activity_col'] = col_idx
    
    # 查找類型、路徑和系統路徑欄位（可能在其他行）
    for row_idx in range(min(15, len(df))):
        row_values = df.iloc[row_idx].astype(str).str.strip()
        for col_idx, col_value in enumerate(row_values):
            if col_value == '類型' and header_info['type_col'] == -1:
                header_info['type_col'] = col_idx
            elif col_value == '路徑' and header_info['path_col'] == -1:
                header_info['path_col'] = col_idx
            elif col_value == '系統路徑' and header_info['system_path_col'] == -1:
                header_info['system_path_col'] = col_idx
    
    print(f"  ✅ 表頭資訊: 表頭行={header_info['header_row']+1}, 課程列={header_info['course_col']+1}, "
          f"章節列={header_info['chapter_col']+1}, 單元列={[c+1 for c in header_info['unit_cols']]}, "
          f"活動列={header_info['activity_col']+1}, 類型列={header_info['type_col']+1}, "
          f"路徑列={header_info['path_col']+1}, 系統路徑列={header_info['system_path_col']+1}")
    
    return header_info

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

def extract_filename_from_path(file_path):
    """從檔案路徑中提取檔案名（不含副檔名）"""
    if not file_path or str(file_path).strip() == '' or str(file_path) == 'nan':
        return ''
    
    # 移除路徑部分，只保留檔案名
    filename = os.path.basename(str(file_path))
    
    # 移除副檔名
    name_without_ext = os.path.splitext(filename)[0]
    
    return name_without_ext

def process_sheet_data(df, sheet_name, header_info):
    """處理單個 sheet 的資料"""
    result_data = []
    resource_data = []
    
    if header_info['header_row'] == -1:
        print(f"  ⚠️  Sheet '{sheet_name}' 沒有找到有效的表頭")
        return result_data, resource_data
    
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
                # 5. 如果有檔案路徑，添加到資源 sheet
                if file_path:
                    resource_title = extract_filename_from_path(file_path)
                    resource_data.append({
                        '檔案名稱': resource_title,
                        '檔案路徑': file_path,
                        '資源ID': '',
                        '最後修改時間': '',
                        '來源Sheet': sheet_name
                    })
    
    return result_data, resource_data

def create_extracted_excel(source_file, selected_sheets, timestamp):
    """創建提取後的 Excel 檔案"""
    output_filename = os.path.join('to_be_executed', f"todolist_extracted_{timestamp}.xlsx")
    
    all_result_data = []
    all_resource_data = []
    
    # 處理每個選中的 sheet
    for sheet_name in selected_sheets:
        print(f"\n📋 正在處理 sheet: {sheet_name}")
        
        try:
            # 讀取 sheet 資料
            df = pd.read_excel(source_file, sheet_name=sheet_name, header=None)
            
            # 查找表頭位置
            header_info = find_header_positions(df)
            
            # 處理資料
            result_data, resource_data = process_sheet_data(df, sheet_name, header_info)
            
            all_result_data.extend(result_data)
            all_resource_data.extend(resource_data)
            
            print(f"  ✅ 處理完成: {len(result_data)} 條記錄，{len(resource_data)} 個資源")
            
        except Exception as e:
            print(f"  ❌ 處理 sheet {sheet_name} 時出錯: {e}")
    
    # 創建 Excel 檔案
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        # 1. 複製原始資料到 Ori_document sheet
        print(f"\n📄 正在複製原始資料...")
        all_ori_data = []
        
        for sheet_name in selected_sheets:
            try:
                # 使用 header=None 讀取，保持原始格式
                df = pd.read_excel(source_file, sheet_name=sheet_name, header=None)
                # 添加來源標記列
                df['原始Sheet'] = sheet_name
                all_ori_data.append(df)
            except Exception as e:
                print(f"  ❌ 複製 sheet {sheet_name} 時出錯: {e}")
        
        if all_ori_data:
            # 垂直合併，保持原始結構
            combined_df = pd.concat(all_ori_data, ignore_index=True, sort=False)
            combined_df.to_excel(writer, sheet_name='Ori_document', index=False, header=False)
            print(f"  ✅ 已保存原始資料 ({len(combined_df)} 行)")
        
        # 2. 保存 Result sheet
        if all_result_data:
            result_df = pd.DataFrame(all_result_data)
            result_df.to_excel(writer, sheet_name='Result', index=False)
            print(f"  ✅ 已保存 Result 資料 ({len(all_result_data)} 條記錄)")
        else:
            # 創建空的 Result sheet
            result_columns = ['類型', '名稱', 'ID', '所屬課程', '所屬課程ID', '所屬章節', '所屬章節ID', 
                            '所屬單元', '所屬單元ID', '學習活動類型', '網址路徑', '檔案路徑', '資源ID', '最後修改時間', '來源Sheet']
            result_df = pd.DataFrame(columns=pd.Index(result_columns))
            result_df.to_excel(writer, sheet_name='Result', index=False)
            print(f"  ✅ 已創建空的 Result sheet")
        
        # 3. 保存 Resource sheet
        if all_resource_data:
            resource_df = pd.DataFrame(all_resource_data)
            resource_df.to_excel(writer, sheet_name='Resource', index=False)
            print(f"  ✅ 已保存 Resource 資料 ({len(all_resource_data)} 個資源)")
        else:
            # 創建空的 Resource sheet 但包含說明行
            resource_df = pd.DataFrame([['檔案名稱', '檔案路徑', '資源ID', '最後修改時間', '來源Sheet']], 
                                     columns=pd.Index(['檔案名稱', '檔案路徑', '資源ID', '最後修改時間', '來源Sheet']))
            resource_df.to_excel(writer, sheet_name='Resource', index=False)
            print(f"  ✅ 已創建空的 Resource sheet")
    
    print(f"\n🎉 提取完成！檔案已生成: {output_filename}")
    return output_filename

def main():
    """主函數"""
    print("🚀 TronClass 課程結構資料提取器")
    print("=" * 50)
    
    # 1. 獲取 analyzed 檔案
    files = get_analyzed_files()
    if not files:
        return
    
    # 2. 讓用戶選擇檔案
    selected_file = select_file(files)
    print(f"\n✅ 已選擇檔案: {os.path.basename(selected_file)}")
    
    # 3. 獲取 sheet 名稱
    sheet_names = get_sheet_names(selected_file)
    if not sheet_names:
        return
    
    # 4. 讓用戶選擇 sheet
    selected_sheets = select_sheet(sheet_names)
    print(f"\n✅ 已選擇 sheet: {', '.join(selected_sheets)}")
    
    # 5. 提取時間戳
    timestamp = extract_timestamp(selected_file)
    
    # 6. 創建提取後的 Excel 檔案
    output_file = create_extracted_excel(selected_file, selected_sheets, timestamp)
    
    print(f"\n📊 生成的檔案包含以下 sheet:")
    print("  1. Ori_document - 原始資料")
    print("  2. Result - 提取的結構化資料")
    print("  3. Resource - 需要上傳的資源檔案清單")
    
    print(f"\n💡 說明:")
    print("  ✅ 已自動提取課程、章節、單元、學習活動的層級關係")
    print("  ✅ 已自動識別線上連結和檔案路徑")
    print("  ✅ 已自動生成資源上傳清單")
    print("  ✅ 支援動態欄位位置識別")
    print("  ✅ 支援多種檔案結構")

if __name__ == "__main__":
    main()
