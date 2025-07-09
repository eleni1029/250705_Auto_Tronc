"""
Resource 資料處理模組 - 修正版本
負責從 result_data 中提取需要上傳的檔案資源
修正邏輯：相同檔案路徑的資源只生成一筆記錄，避免重複上傳
"""

import os

def extract_filename_from_path(file_path):
    """從檔案路徑中提取檔案名（不含副檔名）"""
    if not file_path or str(file_path).strip() == '' or str(file_path) == 'nan':
        return ''
    
    # 移除路徑部分，只保留檔案名
    filename = os.path.basename(str(file_path))
    
    # 移除副檔名
    name_without_ext = os.path.splitext(filename)[0]
    
    return name_without_ext

def extract_resources_from_result(result_data):
    """
    從 result_data 中提取資源檔案清單 - 去重版本
    相同檔案路徑的資源只生成一筆記錄，避免重複上傳
    """
    resource_data = []
    seen_file_paths = set()  # 用於追蹤已處理的檔案路徑
    course_file_mapping = {}  # 記錄每個課程中出現的檔案路徑
    
    # 第一步：統計每個檔案路徑在哪些課程中出現
    for item in result_data:
        if item['類型'] == '學習活動' and item['檔案路徑']:
            file_path = item['檔案路徑'].strip()
            if file_path and file_path != 'nan':
                course_name = item['所屬課程']
                
                if file_path not in course_file_mapping:
                    course_file_mapping[file_path] = {
                        'courses': set(),
                        'first_occurrence': item  # 記錄第一次出現的項目
                    }
                course_file_mapping[file_path]['courses'].add(course_name)
    
    # 第二步：為每個唯一的檔案路徑生成一筆資源記錄
    for file_path, info in course_file_mapping.items():
        first_item = info['first_occurrence']
        courses = info['courses']
        
        resource_title = extract_filename_from_path(file_path)
        
        # 生成來源Sheet資訊（如果跨多個課程，標註多個來源）
        if len(courses) == 1:
            source_sheet = first_item['來源Sheet']
        else:
            # 多個課程使用同一檔案
            source_sheet = f"{first_item['來源Sheet']} (跨{len(courses)}個課程)"
        
        resource_data.append({
            '檔案名稱': resource_title,
            '檔案路徑': file_path,
            '資源ID': '',
            '最後修改時間': '',
            '來源Sheet': source_sheet,
            '引用課程數': len(courses),  # 新增：記錄有多少個課程引用此資源
            '引用課程列表': ', '.join(sorted(courses))  # 新增：記錄引用的課程列表
        })
        
        print(f"  📁 資源: {resource_title} (路徑: {file_path})")
        if len(courses) > 1:
            print(f"      ⚠️  此資源被 {len(courses)} 個課程引用: {', '.join(sorted(courses))}")
    
    print(f"\n📊 資源去重統計:")
    print(f"  - 唯一檔案路徑數: {len(resource_data)}")
    
    # 統計跨課程共用的檔案
    cross_course_files = [r for r in resource_data if r['引用課程數'] > 1]
    if cross_course_files:
        print(f"  - 跨課程共用檔案: {len(cross_course_files)} 個")
        for resource in cross_course_files:
            print(f"    • {resource['檔案名稱']}: {resource['引用課程列表']}")
    else:
        print(f"  - 跨課程共用檔案: 0 個")
    
    return resource_data

def get_resource_statistics(result_data):
    """
    獲取資源統計資訊，用於調試和報告
    """
    stats = {
        'total_activities_with_files': 0,
        'unique_file_paths': 0,
        'cross_course_files': 0,
        'file_path_usage': {}  # 檔案路徑 -> 使用次數
    }
    
    file_path_count = {}
    
    for item in result_data:
        if item['類型'] == '學習活動' and item['檔案路徑']:
            file_path = item['檔案路徑'].strip()
            if file_path and file_path != 'nan':
                stats['total_activities_with_files'] += 1
                
                if file_path not in file_path_count:
                    file_path_count[file_path] = 0
                file_path_count[file_path] += 1
    
    stats['unique_file_paths'] = len(file_path_count)
    stats['cross_course_files'] = len([count for count in file_path_count.values() if count > 1])
    stats['file_path_usage'] = file_path_count
    
    return stats