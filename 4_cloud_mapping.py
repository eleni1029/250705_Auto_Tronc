import json
import pandas as pd
from pathlib import Path
import os

def generate_excel_from_path_mappings(json_file_path, output_excel_path):
    """
    從 path_mappings.json 生成 Excel 文件
    
    Args:
        json_file_path (str): path_mappings.json 文件路徑
        output_excel_path (str): 輸出的 Excel 文件路徑
    """
    
    try:
        # 讀取 JSON 文件
        with open(json_file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
        
        # 準備數據列表
        excel_data = []
        
        # 處理每個項目
        for item in data:
            source_dir = item.get('source_directory_relative', '')
            xml_path = item.get('xml_relative_path', '')
            
            # 提取第一個斜線前的內容作為「名稱」
            if '/' in source_dir:
                name = source_dir.split('/')[0]
            else:
                name = source_dir
            
            # 創建行數據 - 調整欄位順序：名稱、資源庫路徑、資料夾路徑、原始 manifest.xml 路徑
            row_data = {
                '名稱': name,
                '資源庫路徑': '',  # 留空
                '資料夾路徑': f"merged_projects/{source_dir}",
                '原始 imsmanifest.xml 路徑': f"merged_projects/{xml_path}"
            }
            
            excel_data.append(row_data)
        
        # 創建 DataFrame
        df = pd.DataFrame(excel_data)
        
        # 按照「名稱」欄位進行正序排序
        df = df.sort_values(by='名稱', ascending=True).reset_index(drop=True)
        
        # 確保輸出目錄存在
        output_dir = os.path.dirname(output_excel_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 寫入 Excel 文件
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='課程清單', index=False)
            
            # 獲取工作表以進行格式調整
            worksheet = writer.sheets['課程清單']
            
            # 調整欄寬
            worksheet.column_dimensions['A'].width = 25  # 名稱欄
            worksheet.column_dimensions['B'].width = 30  # 資源庫路徑欄
            worksheet.column_dimensions['C'].width = 60  # 資料夾路徑欄
            worksheet.column_dimensions['D'].width = 70  # 原始 imsmanifest.xml 路徑欄
        
        print(f"✅ Excel 文件已成功生成：{output_excel_path}")
        print(f"📊 共處理 {len(excel_data)} 筆記錄（已按名稱排序）")
        
        # 顯示前幾行數據預覽
        print("\n📋 數據預覽（前3行，已排序）：")
        print(df.head(3).to_string(index=False))
        
        return True
        
    except FileNotFoundError:
        print(f"❌ 錯誤：找不到文件 {json_file_path}")
        return False
    
    except json.JSONDecodeError:
        print(f"❌ 錯誤：JSON 文件格式不正確 {json_file_path}")
        return False
    
    except Exception as e:
        print(f"❌ 處理過程中發生錯誤：{str(e)}")
        return False

def main():
    """主函數"""
    
    # 設定文件路徑
    json_file_path = "manifest_structures/path_mappings.json"
    output_excel_path = "4_資源庫路徑_補充.xlsx"
    
    # 如果指定路徑不存在，嘗試當前目錄下的路徑
    if not os.path.exists(json_file_path):
        alternative_path = "path_mappings.json"
        if os.path.exists(alternative_path):
            json_file_path = alternative_path
            print(f"📁 使用當前目錄的文件：{json_file_path}")
        else:
            print(f"❌ 找不到 JSON 文件，請確認路徑：")
            print(f"   - {json_file_path}")
            print(f"   - {alternative_path}")
            return
    
    # 執行轉換
    print("🔄 開始處理 path_mappings.json...")
    success = generate_excel_from_path_mappings(json_file_path, output_excel_path)
    
    if success:
        print(f"\n✅ 處理完成！Excel 文件位於：{os.path.abspath(output_excel_path)}")
    else:
        print("\n❌ 處理失敗，請檢查錯誤信息")

if __name__ == "__main__":
    # 檢查必要的套件
    try:
        import pandas as pd
        import openpyxl
    except ImportError as e:
        print("❌ 缺少必要的 Python 套件，請執行以下命令安裝：")
        print("pip install pandas openpyxl")
        exit(1)
    
    main()