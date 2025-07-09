import os
import glob
import pandas as pd
from datetime import datetime
import re

from sub_todolist_result import process_sheet_data
from sub_todolist_resource import extract_resources_from_result, get_resource_statistics

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

def create_extracted_excel(source_file, selected_sheets, timestamp):
    """創建提取後的 Excel 檔案"""
    output_filename = os.path.join('to_be_executed', f"todolist_extracted_{timestamp}.xlsx")
    
    all_result_data = []
    
    # 處理每個選中的 sheet
    for sheet_name in selected_sheets:
        print(f"\n📋 正在處理 sheet: {sheet_name}")
        
        try:
            # 讀取 sheet 資料
            df = pd.read_excel(source_file, sheet_name=sheet_name, header=None)
            
            # 查找表頭位置
            header_info = find_header_positions(df)
            
            # 使用 sub_todolist_result 處理資料
            result_data = process_sheet_data(df, sheet_name, header_info)
            
            all_result_data.extend(result_data)
            
            print(f"  ✅ 處理完成: {len(result_data)} 條記錄")
            
        except Exception as e:
            print(f"  ❌ 處理 sheet {sheet_name} 時出錯: {e}")
    
    # 使用 sub_todolist_resource 從 result_data 中提取資源（去重版本）
    print(f"\n📦 正在提取資源檔案清單...")
    all_resource_data = extract_resources_from_result(all_result_data)
    
    # 獲取資源統計資訊
    resource_stats = get_resource_statistics(all_result_data)
    
    # 顯示統計資訊
    print(f"\n📊 資源處理統計:")
    print(f"  - 有檔案路徑的學習活動: {resource_stats['total_activities_with_files']} 個")
    print(f"  - 唯一檔案路徑: {resource_stats['unique_file_paths']} 個")
    print(f"  - 生成資源記錄: {len(all_resource_data)} 筆")
    
    if resource_stats['cross_course_files'] > 0:
        print(f"  - 跨課程共用檔案: {resource_stats['cross_course_files']} 個")
        print(f"    💡 這些檔案只會上傳一次，但可被多個學習活動引用")
    
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
            print(f"  ✅ 已保存 Resource 資料 ({len(all_resource_data)} 個唯一資源)")
        else:
            # 創建空的 Resource sheet 但包含標題行
            resource_columns = ['檔案名稱', '檔案路徑', '資源ID', '最後修改時間', '來源Sheet', '引用課程數', '引用課程列表']
            resource_df = pd.DataFrame(columns=pd.Index(resource_columns))
            resource_df.to_excel(writer, sheet_name='Resource', index=False)
            print(f"  ✅ 已創建空的 Resource sheet")
    
    print(f"\n🎉 提取完成！檔案已生成: {output_filename}")
    
    # 顯示重複資源節省的資訊
    if resource_stats['total_activities_with_files'] > resource_stats['unique_file_paths']:
        saved_count = resource_stats['total_activities_with_files'] - resource_stats['unique_file_paths']
        print(f"\n💾 資源去重效果:")
        print(f"  - 原本需要: {resource_stats['total_activities_with_files']} 個資源上傳")
        print(f"  - 去重後需要: {resource_stats['unique_file_paths']} 個資源上傳")
        print(f"  - 節省上傳: {saved_count} 個重複資源")
    
    return output_filename

def main():
    """主函數"""
    print("🚀 TronClass 課程結構資料提取器 - 資源去重版本")
    print("✨ 新功能：相同檔案路徑的資源只會生成一筆記錄")
    print("=" * 55)
    
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
    print("  3. Resource - 需要上傳的資源檔案清單（已去重）")
    
    print(f"\n💡 改進說明:")
    print("  ✅ 已自動提取課程、章節、單元、學習活動的層級關係")
    print("  ✅ 已自動識別線上連結和檔案路徑")
    print("  ✅ 已自動生成資源上傳清單（相同檔案路徑去重）")
    print("  ✅ 支援動態欄位位置識別")
    print("  ✅ 支援多種檔案結構")
    print("  🆕 Resource 表新增欄位：引用課程數、引用課程列表")
    print("  🆕 相同檔案路徑的資源只生成一筆記錄，避免重複上傳")

if __name__ == "__main__":
    main()