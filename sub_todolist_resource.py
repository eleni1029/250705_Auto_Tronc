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
    file_path_info = {}  # 記錄每個檔案路徑的資訊
    
    # 第一步：統計每個檔案路徑的引用情況
    for item in result_data:
        if item['類型'] == '學習活動' and item['檔案路徑']:
            file_path = item['檔案路徑'].strip()
            if file_path and file_path != 'nan':
                if file_path not in file_path_info:
                    file_path_info[file_path] = {
                        'first_occurrence': item,  # 記錄第一次出現的項目
                        'reference_count': 0,      # 引用次數
                        'referencing_activities': []  # 引用的學習活動列表
                    }
                
                file_path_info[file_path]['reference_count'] += 1
                file_path_info[file_path]['referencing_activities'].append(item['名稱'])
    
    # 第二步：為每個唯一的檔案路徑生成一筆資源記錄
    for file_path, info in file_path_info.items():
        first_item = info['first_occurrence']
        reference_count = info['reference_count']
        referencing_activities = info['referencing_activities']
        
        resource_title = extract_filename_from_path(file_path)
        
        resource_data.append({
            '檔案名稱': resource_title,
            '檔案路徑': file_path,
            '資源ID': '',
            '最後修改時間': '',
            '來源Sheet': first_item['來源Sheet'],
            '引用學習活動數': reference_count,  # 記錄引用次數
            '引用活動列表': ', '.join(referencing_activities)  # 記錄引用的活動列表
        })
        
        print(f"  📁 資源: {resource_title} (路徑: {file_path})")
        if reference_count > 1:
            print(f"      🔄 此資源被 {reference_count} 個學習活動引用")
    
    print(f"\n📊 資源去重統計:")
    print(f"  - 唯一檔案路徑數: {len(resource_data)}")
    
    # 統計重複引用的檔案（同一檔案路徑被多個學習活動引用）
    multiple_reference_files = [r for r in resource_data if r['引用學習活動數'] > 1]
    if multiple_reference_files:
        print(f"  - 重複引用的檔案: {len(multiple_reference_files)} 個")
        print(f"    💡 這些檔案只會上傳一次，但被多個學習活動引用")
        # 顯示前幾個重複引用的檔案作為示例
        for i, resource in enumerate(multiple_reference_files[:3]):
            print(f"    • {resource['檔案名稱']}: 被{resource['引用學習活動數']}個活動引用")
        if len(multiple_reference_files) > 3:
            print(f"    • ... 等共{len(multiple_reference_files)}個重複引用檔案")
    else:
        print(f"  - 重複引用的檔案: 0 個")
    
    return resource_data

def get_resource_statistics(result_data):
    """
    獲取資源統計資訊，用於調試和報告
    """
    stats = {
        'total_activities_with_files': 0,
        'unique_file_paths': 0,
        'multiple_reference_files': 0,
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
    stats['multiple_reference_files'] = len([count for count in file_path_count.values() if count > 1])
    stats['file_path_usage'] = file_path_count
    
    return stats