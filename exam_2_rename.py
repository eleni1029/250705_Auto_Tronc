#!/usr/bin/env python3
import os
import glob
import xml.etree.ElementTree as ET
import re
from datetime import datetime
try:
    import pandas as pd
except ImportError:
    print("正在安裝 pandas...")
    os.system("pip3 install pandas openpyxl")
    import pandas as pd

try:
    import natsort
except ImportError:
    print("正在安裝 natsort...")
    os.system("pip3 install natsort")
    import natsort

def convert_chinese_numbers_to_digits(text):
    """將中文數字轉換為阿拉伯數字以便自然排序"""
    if not isinstance(text, str):
        return str(text)
    
    # 中文數字對應表
    chinese_digits = {
        '零': '0', '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
        '六': '6', '七': '7', '八': '8', '九': '9', '十': '10',
        '〇': '0', '壹': '1', '貳': '2', '參': '3', '肆': '4', '伍': '5',
        '陸': '6', '柒': '7', '捌': '8', '玖': '9', '拾': '10'
    }
    
    result = text
    
    # 處理複雜的十位數字模式（如：十一、二十三、九十九等）
    import re
    
    # 處理 "十X" 模式（如：十一、十二...十九）
    def replace_shi_x(match):
        x = match.group(1)
        if x in chinese_digits:
            return '1' + chinese_digits[x]
        return match.group(0)
    
    result = re.sub(r'十([一二三四五六七八九壹貳參肆伍陸柒捌玖])', replace_shi_x, result)
    
    # 處理 "X十Y" 模式（如：二十一、三十五、九十九等）
    def replace_x_shi_y(match):
        x = match.group(1)
        y = match.group(2) if match.group(2) else ''
        x_digit = chinese_digits.get(x, x)
        y_digit = chinese_digits.get(y, y) if y else '0'
        if y:
            return x_digit + y_digit
        else:
            return x_digit + '0'
    
    result = re.sub(r'([一二三四五六七八九壹貳參肆伍陸柒捌玖])十([一二三四五六七八九壹貳參肆伍陸柒捌玖]?)', replace_x_shi_y, result)
    
    # 處理單獨的 "十"
    result = result.replace('十', '10')
    
    # 處理其他單個中文數字
    for chinese, digit in chinese_digits.items():
        result = result.replace(chinese, digit)
    
    return result

def find_exam_folders():
    """尋找所有 exam_01_ 開頭的資料夾"""
    folders = []
    for item in glob.glob("exam_01_*"):
        if os.path.isdir(item):
            folders.append(item)
    return sorted(folders)

def select_folder(folders):
    """讓用戶選擇要處理的資料夾"""
    if not folders:
        print("找不到任何 exam_01_ 開頭的資料夾")
        return None
    
    print("找到以下 exam_01_ 資料夾:")
    for i, folder in enumerate(folders, 1):
        print(f"{i}. {folder}")
    
    while True:
        try:
            choice = input(f"請選擇要處理的資料夾 (1-{len(folders)}): ").strip()
            if not choice:
                print("取消操作")
                return None
            
            choice_num = int(choice)
            if 1 <= choice_num <= len(folders):
                return folders[choice_num - 1]
            else:
                print(f"請輸入 1 到 {len(folders)} 之間的數字")
        except ValueError:
            print("請輸入有效的數字")
        except KeyboardInterrupt:
            print("\n取消操作")
            return None

def find_xml_files(folder_path):
    """在資料夾中尋找所有 .xml 檔案，並回傳檔案及其所在的第二層資料夾，按照資料夾和檔案名稱排序"""
    xml_files_with_subfolder = []
    
    # 先取得所有子資料夾，並按名稱排序
    subfolders = []
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isdir(item_path):
            subfolders.append(item)
    subfolders.sort()  # 按資料夾名稱排序
    
    # 依照排序後的資料夾順序處理
    for subfolder in subfolders:
        subfolder_path = os.path.join(folder_path, subfolder)
        xml_files_in_folder = []
        
        # 在每個子資料夾中找XML檔案
        for root, _, files in os.walk(subfolder_path):
            for file in files:
                if file.lower().endswith('.xml'):
                    xml_path = os.path.join(root, file)
                    xml_files_in_folder.append(xml_path)
        
        # 按檔案名稱排序
        xml_files_in_folder.sort()
        
        # 加入到結果列表
        for xml_path in xml_files_in_folder:
            xml_files_with_subfolder.append((xml_path, subfolder))
    
    # 處理根目錄中的XML檔案（如果有的話）
    root_xml_files = []
    for file in os.listdir(folder_path):
        if file.lower().endswith('.xml'):
            xml_path = os.path.join(folder_path, file)
            root_xml_files.append(xml_path)
    
    root_xml_files.sort()  # 按檔案名稱排序
    for xml_path in root_xml_files:
        xml_files_with_subfolder.append((xml_path, None))
    
    return xml_files_with_subfolder

def select_operation():
    """讓用戶選擇要執行的操作"""
    print("\n請選擇要執行的操作:")
    print("1. 重新命名檔案")
    print("2. 更新題庫標題")
    print("3. 創建待辦列表")
    
    while True:
        try:
            choice = input("請選擇操作 (1-3): ").strip()
            if not choice:
                print("取消操作")
                return None
            
            choice_num = int(choice)
            if choice_num in [1, 2, 3]:
                return choice_num
            else:
                print("請輸入 1, 2 或 3")
        except ValueError:
            print("請輸入有效的數字")
        except KeyboardInterrupt:
            print("\n取消操作")
            return None

def rename_xml_files(xml_files_with_subfolder):
    """重新命名 XML 檔案，在名稱前加上第二層資料夾名稱"""
    if not xml_files_with_subfolder:
        print("在選擇的資料夾中找不到任何 .xml 檔案")
        return
    
    print(f"\n找到 {len(xml_files_with_subfolder)} 個 XML 檔案，預計重新命名:")
    for xml_file, subfolder in xml_files_with_subfolder:
        filename = os.path.basename(xml_file)
        if not subfolder:
            print(f"  - {filename} -> 跳過 (不在子資料夾中)")
        elif filename.startswith(f"{subfolder}_"):
            print(f"  - {filename} -> 跳過 (已有前綴)")
        else:
            new_filename = f"{subfolder}_{filename}"
            print(f"  - {filename} -> {new_filename}")
    
    confirm = input(f"\n確認執行重新命名？(y/N): ").strip().lower()
    if confirm not in ['y', 'yes', '是']:
        print("取消重新命名操作")
        return
    
    renamed_count = 0
    for xml_file, subfolder in xml_files_with_subfolder:
        try:
            directory = os.path.dirname(xml_file)
            filename = os.path.basename(xml_file)
            
            # 如果沒有第二層資料夾，跳過
            if not subfolder:
                continue
            
            # 檢查檔案名稱是否已經有前綴
            if filename.startswith(f"{subfolder}_"):
                continue
            
            new_filename = f"{subfolder}_{filename}"
            new_path = os.path.join(directory, new_filename)
            
            os.rename(xml_file, new_path)
            print(f"重新命名: {filename} -> {new_filename}")
            renamed_count += 1
            
        except Exception as e:
            print(f"重新命名 {xml_file} 時發生錯誤: {e}")
    
    print(f"\n完成! 成功重新命名 {renamed_count} 個檔案")

def parse_serialized_title(title_text):
    """解析序列化的標題文本，提取Big5內容"""
    if not title_text:
        return None
    
    # 使用正則表達式匹配 s:4:"Big5";s:長度:"內容"
    match = re.search(r's:4:"Big5";s:\d+:"([^"]*)"', title_text)
    if match:
        return match.group(1)
    return None

def create_serialized_title(big5_title):
    """創建新的序列化標題格式"""
    big5_len = len(big5_title.encode('utf-8'))
    return f'a:5:{{s:4:"Big5";s:{big5_len}:"{big5_title}";s:6:"GB2312";s:19:"COPY_COPY_undefined";s:2:"en";s:0:"";s:6:"EUC-JP";s:0:"";s:11:"user_define";s:0:"";}}'

def update_xml_titles(xml_files_with_subfolder):
    """更新 XML 檔案中的題庫標題"""
    if not xml_files_with_subfolder:
        print("在選擇的資料夾中找不到任何 .xml 檔案")
        return
    
    print(f"\n找到 {len(xml_files_with_subfolder)} 個 XML 檔案，預計更新標題:")
    
    # 先檢查哪些檔案需要更新
    files_to_update = []
    for xml_file, subfolder in xml_files_with_subfolder:
        filename = os.path.basename(xml_file)
        if not subfolder:
            print(f"  - {filename} -> 跳過 (不在子資料夾中)")
            continue
            
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            
            # 尋找 wm:title 元素
            ns = {'wm': 'http://www.sun.net.tw/WisdomMaster'}
            title_elem = root.find('.//wm:title', ns)
            
            if title_elem is not None and title_elem.text:
                # 解析序列化的標題
                current_title = parse_serialized_title(title_elem.text)
                if current_title:
                    if current_title.startswith(f"{subfolder}_"):
                        print(f"  - {filename} -> 跳過 (標題已有前綴)")
                    else:
                        new_title = f"{subfolder}_{current_title}"
                        new_serialized = create_serialized_title(new_title)
                        print(f"  - {filename} -> {current_title} -> {new_title}")
                        files_to_update.append((xml_file, subfolder, title_elem, new_serialized, tree))
                else:
                    print(f"  - {filename} -> 跳過 (無法解析標題)")
            else:
                print(f"  - {filename} -> 跳過 (找不到標題)")
                
        except Exception as e:
            print(f"  - {filename} -> 跳過 (解析錯誤: {e})")
    
    if not files_to_update:
        print("\n沒有檔案需要更新")
        return
    
    confirm = input(f"\n確認更新 {len(files_to_update)} 個檔案的標題？(y/N): ").strip().lower()
    if confirm not in ['y', 'yes', '是']:
        print("取消更新操作")
        return
    
    updated_count = 0
    for xml_file, subfolder, title_elem, new_serialized, tree in files_to_update:
        try:
            title_elem.text = new_serialized
            
            # 儲存檔案
            tree.write(xml_file, encoding='utf-8', xml_declaration=True)
            
            filename = os.path.basename(xml_file)
            print(f"更新標題: {filename}")
            updated_count += 1
            
        except Exception as e:
            print(f"更新 {xml_file} 時發生錯誤: {e}")
    
    print(f"\n完成! 成功更新 {updated_count} 個檔案的標題")

def extract_xml_info(xml_file, subfolder):
    """從XML檔案提取所需資訊"""
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        # 提取標題
        ns = {'wm': 'http://www.sun.net.tw/WisdomMaster'}
        title_elem = root.find('.//wm:title', ns)
        title = ""
        if title_elem is not None and title_elem.text:
            parsed_title = parse_serialized_title(title_elem.text)
            title = parsed_title if parsed_title else ""
        
        # 提取檔案名稱
        filename = os.path.basename(xml_file)
        
        return {
            '資料夾': subfolder if subfolder else "",
            'xml文件名': filename,
            '標題名稱': title,
            '題庫編號': "",  # 空白
            '題庫創建時間': "",  # 空白
            '題庫標題': title,  # 同標題名稱
            '標題修改時間': ""  # 空白
        }
        
    except Exception as e:
        print(f"讀取 {xml_file} 時發生錯誤: {e}")
        return None

def create_todolist(xml_files_with_subfolder):
    """創建待辦列表Excel檔案"""
    if not xml_files_with_subfolder:
        print("在選擇的資料夾中找不到任何 .xml 檔案")
        return
    
    print(f"\n正在處理 {len(xml_files_with_subfolder)} 個 XML 檔案...")
    
    # 創建輸出資料夾
    output_dir = "exam_02_xml_todolist"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"已創建資料夾: {output_dir}")
    
    # 提取所有XML檔案資訊
    data_list = []
    for xml_file, subfolder in xml_files_with_subfolder:
        xml_info = extract_xml_info(xml_file, subfolder)
        if xml_info:
            data_list.append(xml_info)
    
    if not data_list:
        print("無法提取任何有效資料")
        return
    
    # 創建DataFrame
    df = pd.DataFrame(data_list)
    
    # 按標題名稱自然排序（支援中文數字）
    if not df.empty and '標題名稱' in df.columns:
        # 先轉換中文數字為阿拉伯數字
        converted_titles = df['標題名稱'].astype(str).apply(convert_chinese_numbers_to_digits)
        # 使用轉換後的標題進行自然排序
        sorted_indices = natsort.index_natsorted(converted_titles)
        df = df.iloc[sorted_indices].reset_index(drop=True)
        print("已按標題名稱進行自然排序（包含中文數字轉換）")
    
    # 生成時間戳檔案名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"exam_todolist_{timestamp}.xlsx"
    excel_path = os.path.join(output_dir, excel_filename)
    
    # 儲存為Excel檔案
    try:
        df.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"\n已創建待辦列表: {excel_path}")
        print(f"包含 {len(data_list)} 筆資料")
        
        # 顯示前5筆資料作為預覽
        print("\n預覽前5筆資料:")
        print(df.head().to_string(index=False))
        
    except Exception as e:
        print(f"儲存Excel檔案時發生錯誤: {e}")

def main():
    print("=== XML 檔案處理工具 ===")
    
    # 尋找 exam_01_ 資料夾
    folders = find_exam_folders()
    
    # 讓用戶選擇資料夾
    selected_folder = select_folder(folders)
    if not selected_folder:
        return
    
    print(f"\n選擇的資料夾: {selected_folder}")
    
    # 讓用戶選擇操作
    operation = select_operation()
    if not operation:
        return
    
    # 尋找 XML 檔案
    xml_files_with_subfolder = find_xml_files(selected_folder)
    
    # 執行選擇的操作
    if operation == 1:
        rename_xml_files(xml_files_with_subfolder)
    elif operation == 2:
        update_xml_titles(xml_files_with_subfolder)
    elif operation == 3:
        create_todolist(xml_files_with_subfolder)

if __name__ == "__main__":
    main()