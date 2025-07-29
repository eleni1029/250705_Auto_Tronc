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

def find_course_for_activity(df, current_row, header_info, result_data):
    """為學習活動查找所屬課程（返回唯一名稱）"""
    original_course_value, _ = find_parent_value_upward(df, current_row, [header_info['course_col']])
    
    if not original_course_value:
        return ''
    
    # 從result_data中找到對應的唯一課程名稱
    for item in reversed(result_data):  # 從後往前找最近的
        if item['類型'] == '課程' and item['名稱'].startswith(original_course_value):
            # 如果是exact match或者是編號版本
            if item['名稱'] == original_course_value or item['名稱'].startswith(f"{original_course_value}_"):
                return item['名稱']
    
    return original_course_value

def find_chapter_for_activity(df, current_row, header_info, result_data, parent_course):
    """為學習活動查找所屬章節（返回唯一名稱）"""
    original_chapter_value, _ = find_parent_value_upward(df, current_row, [header_info['chapter_col']])
    
    if not original_chapter_value:
        return ''
    
    # 從result_data中找到對應的唯一章節名稱
    for item in reversed(result_data):  # 從後往前找最近的
        if (item['類型'] == '章節' and 
            item['所屬課程'] == parent_course and 
            item['名稱'].startswith(original_chapter_value)):
            # 如果是exact match或者是編號版本
            if item['名稱'] == original_chapter_value or item['名稱'].startswith(f"{original_chapter_value}_"):
                return item['名稱']
    
    return original_chapter_value


def find_unit_for_activity(df, current_row, header_info, result_data, parent_course, parent_chapter, original_unit_name):
    """為學習活動查找所屬單元（返回唯一名稱）"""
    # 從result_data中找到對應的唯一單元名稱
    for item in reversed(result_data):  # 從後往前找最近的
        if (item['類型'] == '單元' and 
            item['所屬課程'] == parent_course and 
            item['所屬章節'] == parent_chapter and
            item['名稱'].startswith(original_unit_name)):
            # 如果是exact match或者是編號版本
            if item['名稱'] == original_unit_name or item['名稱'].startswith(f"{original_unit_name}_"):
                return item['名稱']
    
    # 如果在result_data中沒找到，使用duplicate_manager生成唯一名稱
    return duplicate_manager.get_unique_unit_name(parent_course, parent_chapter, original_unit_name)

def determine_activity_parent(df, current_row, header_info, course_name, result_data):
    """
    學習活動父級判定邏輯（修復版本）：
    向上查找第一個非學習活動的結構行（單元或章節），找到就停止
    """
    
    for search_row in range(current_row - 1, -1, -1):
        row = df.iloc[search_row]
        
        # 檢查該行是否有學習活動信息（如果有，跳過繼續向上查找）
        if has_activity_info_from_row(row, header_info):
            continue
        
        # 檢查該行是否有單元信息
        unit_name = get_unit_name_from_row(row, header_info, search_row)
        chapter_name = get_chapter_name_from_row(row, header_info, search_row)
        
        # 如果同時有章節和單元信息，記錄異常但選擇單元
        if unit_name and chapter_name:
            print(f"  ⚠️  異常：第{search_row+1}行同時有章節({chapter_name})和單元({unit_name})信息，選擇單元")
        
        # 優先返回單元（如果有）
        if unit_name:
            # 找到單元，需要找到該單元的唯一章節名稱
            unit_chapter = find_chapter_for_activity(df, search_row, header_info, result_data, course_name)
            # 找到唯一單元名稱
            unique_unit = find_unit_for_activity(df, search_row, header_info, result_data, course_name, unit_chapter, unit_name)
            return unit_chapter, unique_unit
        
        # 其次返回章節（如果有）
        if chapter_name:
            unique_chapter = find_chapter_for_activity(df, search_row, header_info, result_data, course_name)
            return unique_chapter, ''
        
        # 檢查是否到達課程邊界（停止查找）
        if header_info['course_col'] >= 0 and header_info['course_col'] < len(row):
            course_val = str(row.iloc[header_info['course_col']]).strip()
            if course_val and course_val != 'nan' and course_val != '':
                # 遇到課程行，停止查找
                break
    
    # 如果向上查找沒有找到任何結構，使用兜底邏輯
    # 從 result_data 中查找最近創建的章節
    recent_chapter = find_recent_chapter_in_results(course_name, result_data)
    if recent_chapter:
        return recent_chapter, ''
    
    # 如果都沒找到，創建自動章節
    return create_auto_chapter(course_name, result_data), ''

def find_recent_chapter_in_results(course_name, result_data):
    """從 result_data 中查找最近創建的同課程章節"""
    for item in reversed(result_data):  # 從後往前找最近的
        if (item['類型'] == '章節' and 
            item['所屬課程'] == course_name):
            return item['名稱']
    return None

def is_pure_chapter_row(row, header_info):
    """檢查該行是否為純章節行（沒有學習活動信息）"""
    # 檢查是否有活動類型或路徑信息
    has_activity_info = False
    
    # 檢查類型欄位
    if header_info['type_col'] >= 0 and header_info['type_col'] < len(row):
        type_value = str(row.iloc[header_info['type_col']]).strip()
        if type_value and type_value != 'nan' and type_value != '':
            has_activity_info = True
    
    # 檢查路徑欄位
    if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
        path_value = str(row.iloc[header_info['path_col']]).strip()
        if path_value and path_value != 'nan' and path_value != '':
            has_activity_info = True
    
    # 檢查系統路徑欄位
    if header_info['system_path_col'] >= 0 and header_info['system_path_col'] < len(row):
        system_path_value = str(row.iloc[header_info['system_path_col']]).strip()
        if system_path_value and system_path_value != 'nan' and system_path_value != '':
            has_activity_info = True
    
    # 檢查學習活動欄位
    if header_info['activity_col'] >= 0 and header_info['activity_col'] < len(row):
        activity_value = str(row.iloc[header_info['activity_col']]).strip()
        if activity_value and activity_value != 'nan' and activity_value != '':
            has_activity_info = True
    
    # 如果沒有學習活動相關信息，則為純章節行
    return not has_activity_info

def is_pure_unit_row(row, header_info):
    """檢查該行是否為純單元行（沒有學習活動信息）"""
    # 與 is_pure_chapter_row 邏輯相同
    return is_pure_chapter_row(row, header_info)

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

# 全域名稱重複管理器
class DuplicateNameManager:
    def __init__(self):
        # 課程名稱計數器：{original_name: count}
        self.course_counters = {}
        # 課程名稱映射：{original_name: [numbered_name1, numbered_name2, ...]}
        self.course_mapping = {}
        
        # 章節名稱計數器：{course_name: {chapter_name: count}}
        self.chapter_counters = {}
        # 章節名稱映射：{course_name: {original_name: [numbered_name1, ...]}}
        self.chapter_mapping = {}
        
        # 單元名稱計數器：{(course_name, chapter_name): {unit_name: count}}
        self.unit_counters = {}
        # 單元名稱映射：{(course_name, chapter_name): {original_name: [numbered_name1, ...]}}
        self.unit_mapping = {}
        
        # 學習活動名稱計數器：{(course_name, chapter_name, unit_name): {activity_name: count}}
        self.activity_counters = {}
        # 學習活動名稱映射：{(course_name, chapter_name, unit_name): {original_name: [numbered_name1, ...]}}
        self.activity_mapping = {}
    
    def get_unique_course_name(self, original_name):
        """獲取唯一的課程名稱"""
        if original_name not in self.course_counters:
            self.course_counters[original_name] = 0
            self.course_mapping[original_name] = []
        
        self.course_counters[original_name] += 1
        count = self.course_counters[original_name]
        
        if count == 1:
            # 第一次出現，使用原名
            unique_name = original_name
        else:
            # 重複出現，加編號
            unique_name = f"{original_name}_{count-1}"
        
        self.course_mapping[original_name].append(unique_name)
        return unique_name
    
    def get_unique_chapter_name(self, course_name, original_chapter_name):
        """獲取唯一的章節名稱"""
        if course_name not in self.chapter_counters:
            self.chapter_counters[course_name] = {}
            self.chapter_mapping[course_name] = {}
        
        if original_chapter_name not in self.chapter_counters[course_name]:
            self.chapter_counters[course_name][original_chapter_name] = 0
            self.chapter_mapping[course_name][original_chapter_name] = []
        
        self.chapter_counters[course_name][original_chapter_name] += 1
        count = self.chapter_counters[course_name][original_chapter_name]
        
        if count == 1:
            # 第一次出現，使用原名
            unique_name = original_chapter_name
        else:
            # 重複出現，加編號
            unique_name = f"{original_chapter_name}_{count-1}"
        
        self.chapter_mapping[course_name][original_chapter_name].append(unique_name)
        return unique_name
    
    def get_unique_unit_name(self, course_name, chapter_name, original_unit_name):
        """獲取唯一的單元名稱"""
        key = (course_name, chapter_name)
        
        if key not in self.unit_counters:
            self.unit_counters[key] = {}
            self.unit_mapping[key] = {}
        
        if original_unit_name not in self.unit_counters[key]:
            self.unit_counters[key][original_unit_name] = 0
            self.unit_mapping[key][original_unit_name] = []
        
        self.unit_counters[key][original_unit_name] += 1
        count = self.unit_counters[key][original_unit_name]
        
        if count == 1:
            # 第一次出現，使用原名
            unique_name = original_unit_name
        else:
            # 重複出現，加編號
            unique_name = f"{original_unit_name}_{count-1}"
        
        self.unit_mapping[key][original_unit_name].append(unique_name)
        return unique_name
    
    def get_unique_activity_name(self, course_name, chapter_name, unit_name, original_activity_name):
        """獲取唯一的學習活動名稱"""
        # 處理無單元的情況
        unit_name = unit_name or ''
        key = (course_name, chapter_name, unit_name)
        
        if key not in self.activity_counters:
            self.activity_counters[key] = {}
            self.activity_mapping[key] = {}
        
        if original_activity_name not in self.activity_counters[key]:
            self.activity_counters[key][original_activity_name] = 0
            self.activity_mapping[key][original_activity_name] = []
        
        self.activity_counters[key][original_activity_name] += 1
        count = self.activity_counters[key][original_activity_name]
        
        if count == 1:
            # 第一次出現，使用原名
            unique_name = original_activity_name
        else:
            # 重複出現，加編號
            unique_name = f"{original_activity_name}_{count-1}"
        
        self.activity_mapping[key][original_activity_name].append(unique_name)
        return unique_name

# 全域重複名稱管理器實例
duplicate_manager = DuplicateNameManager()

def create_auto_chapter(course_name, result_data):
    """為課程創建自動章節"""
    global course_chapter_counter
    
    # 使用全域計數器（跨所有sheets）
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

def determine_row_record_type(df, row_idx, row, header_info, sheet_name, result_data):
    """
    根據優先級判定一行數據應該產生哪種類型的記錄
    優先級順序：學習活動 > 單元 > 章節 > 課程
    """
    
    # 檢查是否有學習活動信息（最高優先級）
    if has_activity_info(row, header_info):
        return create_activity_record(df, row_idx, row, header_info, sheet_name, result_data)
    
    # 檢查是否有單元信息
    unit_name = get_unit_name(row, header_info)
    if unit_name:
        return create_unit_record(df, row_idx, row, header_info, sheet_name, result_data, unit_name)
    
    # 檢查是否有章節信息
    chapter_name = get_chapter_name(row, header_info)
    if chapter_name:
        return create_chapter_record(df, row_idx, row, header_info, sheet_name, result_data, chapter_name)
    
    # 檢查是否有課程信息（最低優先級）
    course_name = get_course_name(row, header_info)
    if course_name:
        return create_course_record(df, row_idx, row, header_info, sheet_name, result_data, course_name)
    
    return None

def has_activity_info(row, header_info):
    """檢查該行是否有學習活動相關信息"""
    # 檢查學習活動欄位
    if header_info['activity_col'] >= 0 and header_info['activity_col'] < len(row):
        activity_name = str(row.iloc[header_info['activity_col']]).strip()
        if activity_name and activity_name != 'nan' and activity_name != '':
            return True
    
    # 檢查是否有活動類型或路徑信息（即使學習活動欄位為空）
    has_type = False
    has_path = False
    
    if header_info['type_col'] >= 0 and header_info['type_col'] < len(row):
        type_value = str(row.iloc[header_info['type_col']]).strip()
        if type_value and type_value != 'nan' and type_value != '':
            has_type = True
    
    if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
        path_value = str(row.iloc[header_info['path_col']]).strip()
        if path_value and path_value != 'nan' and path_value != '':
            has_path = True
    
    # 如果有類型或路徑信息，且章節欄位有值，則視為學習活動
    if (has_type or has_path) and header_info['chapter_col'] >= 0:
        chapter_value = str(row.iloc[header_info['chapter_col']]).strip()
        if chapter_value and chapter_value != 'nan' and chapter_value != '' and chapter_value != '章節':
            return True
    
    return False

def has_activity_info_from_row(row, header_info):
    """檢查該行（pandas Series）是否有學習活動相關信息"""
    # 檢查學習活動欄位
    if header_info['activity_col'] >= 0 and header_info['activity_col'] < len(row):
        activity_name = str(row.iloc[header_info['activity_col']]).strip()
        if activity_name and activity_name != 'nan' and activity_name != '':
            return True
    
    # 檢查是否有活動類型或路徑信息（即使學習活動欄位為空）
    has_type = False
    has_path = False
    
    if header_info['type_col'] >= 0 and header_info['type_col'] < len(row):
        type_value = str(row.iloc[header_info['type_col']]).strip()
        if type_value and type_value != 'nan' and type_value != '':
            has_type = True
    
    if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
        path_value = str(row.iloc[header_info['path_col']]).strip()
        if path_value and path_value != 'nan' and path_value != '':
            has_path = True
    
    # 如果有類型或路徑信息，且章節欄位有值，則視為學習活動
    if (has_type or has_path) and header_info['chapter_col'] >= 0:
        chapter_value = str(row.iloc[header_info['chapter_col']]).strip()
        if chapter_value and chapter_value != 'nan' and chapter_value != '' and chapter_value != '章節':
            return True
    
    return False

def get_unit_name(row, header_info):
    """獲取單元名稱"""
    for unit_col in header_info['unit_cols']:
        if unit_col < len(row):
            unit_name = str(row.iloc[unit_col]).strip()
            if unit_name and unit_name != 'nan' and unit_name != '':
                return unit_name
    return None

def get_unit_name_from_row(row, header_info, row_idx):
    """獲取該行的單元名稱，處理多個單元欄位的情況"""
    found_units = []
    unit_name = None
    
    for unit_col in header_info['unit_cols']:
        if unit_col < len(row):
            unit_val = str(row.iloc[unit_col]).strip()
            if unit_val and unit_val != 'nan' and unit_val != '':
                found_units.append(unit_val)
                unit_name = unit_val  # 取最後一個
    
    # 如果有多個單元欄位有值，記錄異常
    if len(found_units) > 1:
        print(f"  ⚠️  異常：第{row_idx+1}行有多個單元欄位有值({found_units})，選擇最後一個: {unit_name}")
    
    return unit_name

def get_chapter_name_from_row(row, header_info, row_idx):
    """獲取該行的章節名稱"""
    if header_info['chapter_col'] >= 0 and header_info['chapter_col'] < len(row):
        chapter_name = str(row.iloc[header_info['chapter_col']]).strip()
        if chapter_name and chapter_name != 'nan' and chapter_name != '' and chapter_name != '章節':
            return chapter_name
    return None

def get_chapter_name(row, header_info):
    """獲取章節名稱"""
    if header_info['chapter_col'] >= 0 and header_info['chapter_col'] < len(row):
        chapter_name = str(row.iloc[header_info['chapter_col']]).strip()
        if chapter_name and chapter_name != 'nan' and chapter_name != '' and chapter_name != '章節':
            return chapter_name
    return None

def get_course_name(row, header_info):
    """獲取課程名稱"""
    if header_info['course_col'] >= 0 and header_info['course_col'] < len(row):
        course_name = str(row.iloc[header_info['course_col']]).strip()
        if course_name and course_name != 'nan' and course_name != '':
            return course_name
    return None

def create_course_record(df, row_idx, row, header_info, sheet_name, result_data, course_name):
    """創建課程記錄"""
    unique_course_name = duplicate_manager.get_unique_course_name(course_name)
    
    return {
        '類型': '課程',
        '名稱': unique_course_name,
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
    }

def create_chapter_record(df, row_idx, row, header_info, sheet_name, result_data, chapter_name):
    """創建章節記錄"""
    parent_course = find_course_for_activity(df, row_idx, header_info, result_data)
    
    if not parent_course:
        return None
        
    unique_chapter_name = duplicate_manager.get_unique_chapter_name(parent_course, chapter_name)
    
    return {
        '類型': '章節',
        '名稱': unique_chapter_name,
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
    }

def create_unit_record(df, row_idx, row, header_info, sheet_name, result_data, unit_name):
    """創建單元記錄"""
    parent_course = find_course_for_activity(df, row_idx, header_info, result_data)
    parent_chapter = find_chapter_for_activity(df, row_idx, header_info, result_data, parent_course)
    
    if not parent_course or not parent_chapter:
        return None
        
    unique_unit_name = duplicate_manager.get_unique_unit_name(parent_course, parent_chapter, unit_name)
    
    return {
        '類型': '單元',
        '名稱': unique_unit_name,
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
    }

def create_yes_records(df, row_idx, row, header_info, sheet_name, result_data, 
                      parent_course, target_parent_name, target_parent_type):
    """
    處理 #Yes 邏輯：先創建父級記錄，再創建學習活動記錄
    返回記錄列表：[parent_record, activity_record]
    """
    records = []
    
    if target_parent_type == '章節':
        # 1. 創建章節記錄
        unique_chapter_name = duplicate_manager.get_unique_chapter_name(parent_course, target_parent_name)
        
        chapter_record = {
            '類型': '章節',
            '名稱': unique_chapter_name,
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
        }
        records.append(chapter_record)
        
        # 2. 創建學習活動記錄，名稱為章節名稱
        parent_chapter = unique_chapter_name
        parent_unit = ''
        final_activity_name = unique_chapter_name
        
    else:  # 單元
        # 1. 先找到所屬章節
        parent_chapter = find_chapter_for_activity(df, row_idx, header_info, result_data, parent_course)
        
        # 2. 創建單元記錄
        unique_unit_name = duplicate_manager.get_unique_unit_name(parent_course, parent_chapter, target_parent_name)
        
        unit_record = {
            '類型': '單元',
            '名稱': unique_unit_name,
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
        }
        records.append(unit_record)
        
        # 3. 創建學習活動記錄，名稱為單元名稱
        parent_unit = unique_unit_name
        final_activity_name = unique_unit_name
    
    # 創建學習活動記錄的通用邏輯
    activity_record = create_activity_record_content(
        row, header_info, sheet_name, parent_course, parent_chapter, parent_unit, final_activity_name
    )
    
    if activity_record:
        records.append(activity_record)
    
    return records

def create_activity_record_content(row, header_info, sheet_name, parent_course, parent_chapter, parent_unit, activity_name):
    """創建學習活動記錄的內容部分（路徑、類型等）"""
    
    # 對學習活動名稱進行重複檢測
    unique_activity_name = duplicate_manager.get_unique_activity_name(
        parent_course, parent_chapter, parent_unit, activity_name
    )
    
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
    from sub_mp4filepath_identifier import process_activity_for_media
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
            print(f"  ⚠️  媒體檔案識別失敗，回退到線上連結: {unique_activity_name}")
            print(f"      錯誤訊息: {media_result['error_message']}")
        elif not media_result['use_web_path']:
            # 成功找到媒體檔案：使用檔案路徑
            file_path = media_result['new_file_path']
            web_path = ''
            print(f"  🎬 媒體檔案識別成功: {unique_activity_name} → {activity_type}")
        else:
            # 按原邏輯處理路徑
            if activity_type in ['線上連結', '影音教材_影音連結']:
                if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
                    path_value = str(row.iloc[header_info['path_col']]).strip()
                    if path_value and path_value != 'nan':
                        web_path = path_value
            elif activity_type in ['參考檔案_圖片', '參考檔案_PDF', '影音教材_影片', '影音教材_音訊']:
                if system_path:
                    file_path = system_path
                elif header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
                    path_value = str(row.iloc[header_info['path_col']]).strip()
                    if path_value and path_value != 'nan':
                        file_path = path_value
    else:
        # === 修正後的路徑處理邏輯 ===
        if activity_type in ['線上連結', '影音教材_影音連結']:
            # 需要網址路徑的類型：從 '路徑' 欄位取值填入 '網址路徑'
            if header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
                path_value = str(row.iloc[header_info['path_col']]).strip()
                if path_value and path_value != 'nan':
                    web_path = path_value
        elif activity_type in ['參考檔案_圖片', '參考檔案_PDF', '影音教材_影片', '影音教材_音訊']:
            # 需要檔案路徑的類型：從 '系統路徑' 或 '路徑' 欄位取值填入 '檔案路徑'
            if system_path:
                # 優先使用系統路徑
                file_path = system_path
            elif header_info['path_col'] >= 0 and header_info['path_col'] < len(row):
                # 如果沒有系統路徑，使用路徑欄位
                path_value = str(row.iloc[header_info['path_col']]).strip()
                if path_value and path_value != 'nan':
                    file_path = path_value
    
    return {
        '類型': '學習活動',
        '名稱': unique_activity_name,
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
    }

def create_activity_record(df, row_idx, row, header_info, sheet_name, result_data):
    """創建學習活動記錄"""
    # 獲取學習活動名稱
    activity_name = ''
    if header_info['activity_col'] >= 0 and header_info['activity_col'] < len(row):
        activity_name = str(row.iloc[header_info['activity_col']]).strip()
    
    # 如果學習活動欄位為空，但章節欄位有值，且該行有其他活動資訊，則使用章節欄位作為學習活動
    if (not activity_name or activity_name == 'nan' or activity_name == '') and header_info['chapter_col'] >= 0:
        chapter_value = str(row.iloc[header_info['chapter_col']]).strip()
        if chapter_value and chapter_value != 'nan' and chapter_value != '' and chapter_value != '章節':
            activity_name = chapter_value
    
    if not activity_name or activity_name == 'nan' or activity_name == '':
        return None
    
    # 查找所屬課程
    parent_course = find_course_for_activity(df, row_idx, header_info, result_data)
    if not parent_course:
        return None
    
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
            # #Yes 邏輯：需要先創建父級記錄，再創建學習活動記錄
            # 標記這是 #Yes 邏輯，需要創建多個記錄
            return create_yes_records(df, row_idx, row, header_info, sheet_name, result_data, 
                                    parent_course, target_parent_name, target_parent_type)
        else:
            # 如果沒有找到目標父級，跳過這個#Yes
            return None
    
    # === 普通學習活動邏輯 ===
    else:
        # 使用新的父級判定邏輯
        parent_chapter, parent_unit = determine_activity_parent(df, row_idx, header_info, parent_course, result_data)
        
        
        final_activity_name = activity_name
    
    # 使用通用函數創建學習活動記錄
    return create_activity_record_content(
        row, header_info, sheet_name, parent_course, parent_chapter, parent_unit, final_activity_name
    )

def process_sheet_data(df, sheet_name, header_info):
    """處理單個 sheet 的資料"""
    global course_chapter_counter
    # 注意：不再重置 duplicate_manager，它應該在所有 sheets 處理開始前重置一次
    # 這樣才能正確處理跨 sheet 的重複名稱
    # course_chapter_counter 保持全域性，不需要為每個sheet重置
    
    result_data = []
    
    if header_info['header_row'] == -1:
        print(f"  ⚠️  Sheet '{sheet_name}' 沒有找到有效的表頭")
        return result_data
    
    # 從表頭行的下一行開始處理資料
    start_row = header_info['header_row'] + 1
    
    for row_idx in range(start_row, len(df)):
        row = df.iloc[row_idx]
        
        # 新的優先級處理邏輯：一行數據只產生一種類型的記錄
        # 優先級順序：學習活動 > 單元 > 章節 > 課程
        # 注意：#Yes 邏輯可能返回多個記錄
        row_records = determine_row_record_type(df, row_idx, row, header_info, sheet_name, result_data)
        
        if row_records:
            # 支持單個記錄或多個記錄（#Yes 邏輯）
            if isinstance(row_records, list):
                result_data.extend(row_records)
            else:
                result_data.append(row_records)
    
    return result_data