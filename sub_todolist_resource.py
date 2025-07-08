"""
Resource 資料處理模組
負責從 result_data 中提取需要上傳的檔案資源
邏輯簡單：遍歷 result_data，找到有檔案路徑的學習活動，提取檔案資訊
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
    """從 result_data 中提取資源檔案清單"""
    resource_data = []
    
    for item in result_data:
        # 只處理學習活動類型且有檔案路徑的項目
        if item['類型'] == '學習活動' and item['檔案路徑']:
            file_path = item['檔案路徑'].strip()
            if file_path and file_path != 'nan':
                resource_title = extract_filename_from_path(file_path)
                resource_data.append({
                    '檔案名稱': resource_title,
                    '檔案路徑': file_path,
                    '資源ID': '',
                    '最後修改時間': '',
                    '來源Sheet': item['來源Sheet']
                })
    
    return resource_data