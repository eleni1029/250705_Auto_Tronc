#!/usr/bin/env python3
"""
Excel 檔案分析工具
用於檢查 todolist_extracted Excel 檔案中的問題
"""

import pandas as pd
import os
import sys

def analyze_excel_file(excel_path):
    """分析 Excel 檔案並檢查問題"""
    print(f"📋 分析檔案: {os.path.basename(excel_path)}")
    print("=" * 60)
    
    if not os.path.exists(excel_path):
        print(f"❌ 檔案不存在: {excel_path}")
        return
    
    try:
        # 讀取所有 sheets
        excel_file = pd.ExcelFile(excel_path)
        print(f"📊 檔案包含 {len(excel_file.sheet_names)} 個 sheet: {', '.join(excel_file.sheet_names)}")
        
        # 分析每個 sheet
        for sheet_name in excel_file.sheet_names:
            print(f"\n📋 分析 Sheet: {sheet_name}")
            print("-" * 40)
            
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            
            if sheet_name == 'Result':
                analyze_result_sheet(df)
            elif sheet_name == 'Resource':
                analyze_resource_sheet(df)
            elif sheet_name == 'Ori_document':
                analyze_ori_document_sheet(df)
            else:
                print(f"  ℹ️  未知的 sheet 類型，跳過詳細分析")
                print(f"  📏 形狀: {df.shape}")
    
    except Exception as e:
        print(f"❌ 讀取檔案時出錯: {e}")

def analyze_result_sheet(df):
    """分析 Result sheet"""
    print(f"  📏 數據形狀: {df.shape}")
    print(f"  📊 欄位: {list(df.columns)}")
    
    # 檢查空行
    empty_rows = df.isnull().all(axis=1).sum()
    if empty_rows > 0:
        print(f"  ⚠️  發現 {empty_rows} 個空行")
    
    # 檢查各類型的數量
    if '類型' in df.columns:
        type_counts = df['類型'].value_counts()
        print(f"  📈 類型統計:")
        for type_name, count in type_counts.items():
            print(f"    - {type_name}: {count}")
    
    # 查找特定活動 "課程導論與評量方式"
    print(f"\n🔍 查找活動 '課程導論與評量方式':")
    target_activity = "課程導論與評量方式"
    
    if '名稱' in df.columns:
        matching_rows = df[df['名稱'].str.contains(target_activity, na=False)]
        
        if not matching_rows.empty:
            print(f"  ✅ 找到 {len(matching_rows)} 個匹配的記錄:")
            
            for idx, row in matching_rows.iterrows():
                print(f"\n  📍 行號: {idx + 2}")  # +2 因為 Excel 從1開始且有標題行
                print(f"    類型: {row.get('類型', 'N/A')}")
                print(f"    名稱: {row.get('名稱', 'N/A')}")
                print(f"    ID: {row.get('ID', 'N/A')}")
                print(f"    所屬課程: {row.get('所屬課程', 'N/A')}")
                print(f"    所屬課程ID: {row.get('所屬課程ID', 'N/A')}")
                print(f"    所屬章節: {row.get('所屬章節', 'N/A')}")
                print(f"    所屬章節ID: {row.get('所屬章節ID', 'N/A')}")
                print(f"    所屬單元: {row.get('所屬單元', 'N/A')}")
                print(f"    所屬單元ID: {row.get('所屬單元ID', 'N/A')}")
                print(f"    學習活動類型: {row.get('學習活動類型', 'N/A')}")
                print(f"    網址路徑: {row.get('網址路徑', 'N/A')}")
                print(f"    檔案路徑: {row.get('檔案路徑', 'N/A')}")
                print(f"    資源ID: {row.get('資源ID', 'N/A')}")
                print(f"    最後修改時間: {row.get('最後修改時間', 'N/A')}")
                print(f"    來源Sheet: {row.get('來源Sheet', 'N/A')}")
                
                # 檢查關鍵信息
                check_activity_info(row)
        else:
            print(f"  ❌ 未找到名稱包含 '{target_activity}' 的記錄")
    
    # 檢查 ID 欄位的有效性
    check_id_fields(df)

def check_activity_info(row):
    """檢查單個活動的關鍵信息"""
    print(f"  🔍 檢查結果:")
    
    # 檢查活動類型
    activity_type = row.get('學習活動類型', '')
    if activity_type:
        if 'video' in activity_type.lower() or '影片' in activity_type or '音訊' in activity_type:
            print(f"    ✅ 活動類型設定正確: {activity_type}")
        else:
            print(f"    ⚠️  活動類型可能不正確: {activity_type}")
    else:
        print(f"    ❌ 活動類型為空")
    
    # 檢查檔案路徑
    file_path = row.get('檔案路徑', '')
    if file_path:
        if '0-2.mp4' in file_path:
            print(f"    ✅ 檔案路徑包含期望的檔案: 0-2.mp4")
        else:
            print(f"    ⚠️  檔案路徑不包含期望的檔案 '0-2.mp4': {file_path}")
        
        # 檢查檔案是否存在
        if os.path.exists(file_path):
            print(f"    ✅ 檔案路徑存在")
        else:
            print(f"    ❌ 檔案路徑不存在: {file_path}")
    else:
        print(f"    ❌ 檔案路徑為空")
    
    # 檢查 ID 相關字段
    resource_id = row.get('資源ID', '')
    if resource_id:
        try:
            resource_id_num = int(resource_id)
            if resource_id_num == 8390:
                print(f"    ✅ 資源ID匹配: {resource_id}")
            else:
                print(f"    ⚠️  資源ID不匹配期望值 8390: {resource_id}")
        except:
            print(f"    ❌ 資源ID不是有效數值: {resource_id}")
    else:
        print(f"    ⚠️  資源ID為空")

def check_id_fields(df):
    """檢查各種 ID 欄位的有效性"""
    print(f"\n  🔍 檢查 ID 欄位有效性:")
    
    id_columns = ['ID', '所屬課程ID', '所屬章節ID', '所屬單元ID', '資源ID']
    
    for col in id_columns:
        if col in df.columns:
            # 統計非空值
            non_empty = df[col].notna().sum()
            total = len(df)
            
            # 檢查數值有效性
            numeric_valid = 0
            if non_empty > 0:
                for val in df[col].dropna():
                    try:
                        if val != '' and str(val).strip() != '':
                            int(val)
                            numeric_valid += 1
                    except:
                        pass
            
            print(f"    {col}: {non_empty}/{total} 非空, {numeric_valid} 個有效數值")

def analyze_resource_sheet(df):
    """分析 Resource sheet"""
    print(f"  📏 數據形狀: {df.shape}")
    print(f"  📊 欄位: {list(df.columns)}")
    
    # 查找包含 0-2.mp4 的資源
    print(f"\n  🔍 查找包含 '0-2.mp4' 的資源:")
    
    if '檔案路徑' in df.columns:
        matching_resources = df[df['檔案路徑'].str.contains('0-2.mp4', na=False)]
        
        if not matching_resources.empty:
            print(f"    ✅ 找到 {len(matching_resources)} 個匹配的資源:")
            
            for idx, row in matching_resources.iterrows():
                print(f"\n    📍 行號: {idx + 2}")
                print(f"      檔案名稱: {row.get('檔案名稱', 'N/A')}")
                print(f"      檔案路徑: {row.get('檔案路徑', 'N/A')}")
                print(f"      資源ID: {row.get('資源ID', 'N/A')}")
                print(f"      引用學習活動數: {row.get('引用學習活動數', 'N/A')}")
                print(f"      引用活動列表: {row.get('引用活動列表', 'N/A')}")
        else:
            print(f"    ❌ 未找到包含 '0-2.mp4' 的資源")
    
    # 檢查資源ID的分配
    if '資源ID' in df.columns:
        non_empty_ids = df['資源ID'].dropna()
        if len(non_empty_ids) > 0:
            print(f"\n  📊 資源ID統計:")
            print(f"    已分配ID的資源: {len(non_empty_ids)}")
            try:
                # 嘗試轉換為數值類型計算範圍
                numeric_ids = pd.to_numeric(non_empty_ids, errors='coerce').dropna()
                if len(numeric_ids) > 0:
                    print(f"    ID範圍: {int(numeric_ids.min())} - {int(numeric_ids.max())}")
                else:
                    print(f"    ID範圍: 無法計算（非數值ID）")
            except Exception as e:
                print(f"    ID範圍: 計算錯誤 - {e}")

def analyze_ori_document_sheet(df):
    """分析 Ori_document sheet"""
    print(f"  📏 數據形狀: {df.shape}")
    print(f"  ℹ️  這是原始文檔，包含原始格式的數據")

def main():
    """主函數"""
    excel_path = "/Users/dominic/250708_Auto_Tronc/6_todolist/todolist_extracted_20250722_103809.xlsx"
    
    print("🔍 Excel 檔案分析器")
    print("檢查 todolist_extracted 檔案中的潛在問題")
    print("=" * 60)
    
    analyze_excel_file(excel_path)
    
    print("\n" + "=" * 60)
    print("✅ 分析完成")

if __name__ == "__main__":
    main()