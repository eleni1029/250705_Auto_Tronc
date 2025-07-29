#!/usr/bin/env python3
"""
詳細的 Excel 檔案檢查工具
專門檢查課綱和模組關聯關係，以及數據完整性
"""

import pandas as pd
import os

def check_syllabus_module_relationship(df):
    """檢查課綱和模組的關聯關係"""
    print("\n🔗 檢查課綱(syllabus)和模組(module)關聯關係:")
    print("-" * 50)
    
    # 查找目標活動
    target_activity = df[df['名稱'].str.contains('課程導論與評量方式', na=False)]
    
    if target_activity.empty:
        print("❌ 未找到目標活動")
        return
    
    activity = target_activity.iloc[0]
    print(f"📍 目標活動: {activity['名稱']}")
    print(f"  所屬課程: {activity['所屬課程']}")
    print(f"  所屬課程ID (syllabus_id): {activity['所屬課程ID']}")
    print(f"  所屬章節: {activity['所屬章節']}")
    print(f"  所屬章節ID: {activity['所屬章節ID']}")
    print(f"  所屬單元: {activity['所屬單元']}")
    print(f"  所屬單元ID (module_id): {activity['所屬單元ID']}")
    print(f"  資源ID (upload_id): {activity['資源ID']}")
    
    # 檢查期望的數值
    expected_values = {
        'syllabus_id': 2141,
        'module_id': 2061,  
        'upload_id': 8390
    }
    
    actual_values = {
        'syllabus_id': activity['所屬課程ID'],
        'module_id': activity['所屬章節ID'],  # 根據輸出，章節ID對應module_id
        'upload_id': activity['資源ID']
    }
    
    print(f"\n🔍 數值匹配檢查:")
    for key in expected_values:
        expected = expected_values[key]
        actual = actual_values[key]
        
        try:
            actual_int = int(float(actual)) if pd.notna(actual) else None
            if actual_int == expected:
                print(f"  ✅ {key}: {expected} ✓")
            else:
                print(f"  ❌ {key}: 期望 {expected}, 實際 {actual_int}")
        except:
            print(f"  ❌ {key}: 期望 {expected}, 實際 {actual} (無法轉換為數值)")

def check_file_structure_issues(df):
    """檢查文件結構問題"""
    print("\n📋 檢查文件整體結構:")
    print("-" * 30)
    
    # 檢查空行
    completely_empty_rows = df.isnull().all(axis=1).sum()
    print(f"  完全空行數: {completely_empty_rows}")
    
    # 檢查關鍵欄位的空值情況
    key_columns = ['類型', '名稱', '所屬課程']
    for col in key_columns:
        if col in df.columns:
            null_count = df[col].isnull().sum()
            empty_string_count = (df[col] == '').sum()
            total_empty = null_count + empty_string_count
            print(f"  {col} 空值: {total_empty}/{len(df)} ({total_empty/len(df)*100:.1f}%)")
    
    # 檢查ID欄位的有效性
    id_columns = ['ID', '所屬課程ID', '所屬章節ID', '所屬單元ID', '資源ID']
    print(f"\n  ID欄位有效性檢查:")
    for col in id_columns:
        if col in df.columns:
            non_null = df[col].notna()
            non_empty = df[col] != ''
            valid_mask = non_null & non_empty
            valid_count = valid_mask.sum()
            
            # 檢查能轉換為數值的數量
            numeric_count = 0
            if valid_count > 0:
                valid_values = df.loc[valid_mask, col]
                for val in valid_values:
                    try:
                        float(val)
                        numeric_count += 1
                    except:
                        pass
            
            print(f"    {col}: {valid_count}/{len(df)} 非空, {numeric_count} 個有效數值")

def check_activity_types(df):
    """檢查活動類型分佈"""
    print("\n🎯 檢查學習活動類型分佈:")
    print("-" * 30)
    
    activities = df[df['類型'] == '學習活動']
    if not activities.empty and '學習活動類型' in activities.columns:
        activity_types = activities['學習活動類型'].value_counts()
        print(f"  學習活動總數: {len(activities)}")
        print(f"  活動類型分佈:")
        for activity_type, count in activity_types.head(10).items():
            print(f"    {activity_type}: {count}")
        
        # 檢查影音相關的活動
        video_related = activities[activities['學習活動類型'].str.contains('影音|video|Video', na=False)]
        print(f"\n  影音相關活動: {len(video_related)}")
        
        # 檢查online_video類型
        online_video = activities[activities['學習活動類型'].str.contains('online_video', na=False)]
        print(f"  online_video類型: {len(online_video)}")

def check_file_paths(df):
    """檢查檔案路徑"""
    print("\n📁 檢查檔案路徑:")
    print("-" * 20)
    
    activities = df[df['類型'] == '學習活動']
    if not activities.empty:
        # 檢查有檔案路徑的活動
        has_file_path = activities['檔案路徑'].notna() & (activities['檔案路徑'] != '')
        file_path_count = has_file_path.sum()
        print(f"  有檔案路徑的學習活動: {file_path_count}/{len(activities)}")
        
        # 檢查有網址路徑的活動
        has_web_path = activities['網址路徑'].notna() & (activities['網址路徑'] != '')
        web_path_count = has_web_path.sum()
        print(f"  有網址路徑的學習活動: {web_path_count}/{len(activities)}")
        
        # 檢查包含特定檔案名的路徑
        target_files = ['0-2.mp4', '.mp4', '.html']
        for target in target_files:
            matching = activities['檔案路徑'].str.contains(target, na=False).sum()
            print(f"  包含 '{target}' 的檔案路徑: {matching}")

def check_specific_course_structure(df):
    """檢查特定課程的結構"""
    print("\n🏗️ 檢查 Organization 課程結構:")
    print("-" * 35)
    
    org_course = df[df['所屬課程'] == 'Organization']
    if not org_course.empty:
        print(f"  Organization 課程相關記錄: {len(org_course)}")
        
        # 按類型統計
        type_counts = org_course['類型'].value_counts()
        for type_name, count in type_counts.items():
            print(f"    {type_name}: {count}")
        
        # 檢查章節
        chapters = org_course[org_course['類型'] == '章節']['名稱'].unique()
        print(f"\n  章節列表:")
        for chapter in chapters:
            chapter_data = org_course[org_course['所屬章節'] == chapter]
            activities_count = len(chapter_data[chapter_data['類型'] == '學習活動'])
            units_count = len(chapter_data[chapter_data['類型'] == '單元'])
            print(f"    {chapter}: {activities_count} 活動, {units_count} 單元")

def main():
    """主函數"""
    excel_path = "/Users/dominic/250708_Auto_Tronc/to_be_executed/todolist_extracted_20250722_103809.xlsx"
    
    print("🔍 詳細 Excel 檔案檢查工具")
    print("專門檢查課綱和模組關聯關係")
    print("=" * 60)
    
    if not os.path.exists(excel_path):
        print(f"❌ 檔案不存在: {excel_path}")
        return
    
    try:
        # 讀取 Result sheet
        df = pd.read_excel(excel_path, sheet_name='Result')
        print(f"✅ 成功讀取 Result sheet: {df.shape}")
        
        # 執行各項檢查
        check_syllabus_module_relationship(df)
        check_file_structure_issues(df)
        check_activity_types(df)
        check_file_paths(df)
        check_specific_course_structure(df)
        
    except Exception as e:
        print(f"❌ 處理檔案時出錯: {e}")
    
    print("\n" + "=" * 60)
    print("✅ 詳細檢查完成")

if __name__ == "__main__":
    main()