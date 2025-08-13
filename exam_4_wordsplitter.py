#!/usr/bin/env python3
"""
進階Word檔案處理器
處理根資料夾的複製、Word檔案重命名和章節拆分
"""

import os
import shutil
import docx
import re
import logging
from datetime import datetime
from pathlib import Path
from sub_word_splitter import WordSplitter

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

class AdvancedWordProcessor:
    def __init__(self):
        self.base_dir = Path.cwd()
        self.source_dir = self.base_dir / "exam_01_02_merged_projects"
        self.target_dir = self.base_dir / "exam_03_wordsplitter"
        self.log_dir = self.base_dir / "log"
        
        # 設置logging
        self.setup_logging()
        
        # 初始化Word拆分器
        self.word_splitter = WordSplitter(self.logger)
        
    def setup_logging(self):
        """設置logging系統"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = self.log_dir / f"advanced_word_processing_{timestamp}.log"
        
        # 確保log目錄存在
        self.log_dir.mkdir(exist_ok=True)
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def find_root_folders_with_word_files(self):
        """找到包含Word檔案的根資料夾"""
        root_folders = []
        
        if not self.source_dir.exists():
            self.logger.error(f"源資料夾不存在: {self.source_dir}")
            return root_folders
        
        self.logger.info(f"掃描根資料夾: {self.source_dir}")
        
        # 只掃描第一層資料夾
        for item in self.source_dir.iterdir():
            if item.is_dir():
                word_files = self.find_word_files_in_folder(item)
                if word_files:
                    folder_info = {
                        'path': item,
                        'name': item.name,
                        'word_files': word_files
                    }
                    root_folders.append(folder_info)
                    self.logger.info(f"找到根資料夾: {item.name} (包含 {len(word_files)} 個Word檔案)")
        
        return root_folders
    
    def find_word_files_in_folder(self, folder_path):
        """遞歸查找資料夾中的所有Word檔案"""
        word_files = []
        
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith(('.docx', '.doc')) and not file.startswith('~$'):
                    word_files.append(Path(root) / file)
        
        return word_files
    
    def copy_root_folder(self, folder_info):
        """複製根資料夾到目標位置"""
        source_path = folder_info['path']
        target_path = self.target_dir / folder_info['name']
        
        try:
            # 如果目標資料夾已存在，先刪除
            if target_path.exists():
                shutil.rmtree(target_path)
                self.logger.info(f"刪除已存在的目標資料夾: {target_path}")
            
            # 複製整個資料夾
            shutil.copytree(source_path, target_path)
            self.logger.info(f"成功複製根資料夾: {folder_info['name']}")
            
            # 更新folder_info中的路徑
            folder_info['target_path'] = target_path
            
            # 更新word_files路徑
            new_word_files = []
            for old_path in folder_info['word_files']:
                relative_path = old_path.relative_to(source_path)
                new_path = target_path / relative_path
                new_word_files.append(new_path)
            
            folder_info['word_files'] = new_word_files
            
            return True
            
        except Exception as e:
            self.logger.error(f"複製根資料夾失敗 {folder_info['name']}: {e}")
            return False
    
    def process_single_word_file(self, folder_info):
        """處理只有一個Word檔案的情況"""
        if len(folder_info['word_files']) != 1:
            return
        
        word_file = folder_info['word_files'][0]
        folder_name = folder_info['name']
        
        # 步驟1: 重命名Word檔案為根資料夾名稱
        new_name = f"{folder_name}.docx"
        new_path = word_file.parent / new_name
        
        try:
            # 如果新名稱與舊名稱不同，才進行重命名
            if word_file.name != new_name:
                word_file.rename(new_path)
                self.logger.info(f"重命名檔案: {word_file.name} -> {new_name}")
                word_file = new_path
                folder_info['word_files'][0] = word_file
        
        except Exception as e:
            self.logger.error(f"重命名檔案失敗: {e}")
            return
        
        # 步驟2: 分析是否需要拆分
        analysis = self.word_splitter.analyze_document(word_file)
        
        self.logger.info(f"章節分析結果 - {word_file.name}:")
        self.logger.info(f"  包含章節: {analysis['has_chapters']}")
        self.logger.info(f"  章節數量: {analysis['chapter_count']}")
        self.logger.info(f"  需要拆分: {analysis['needs_splitting']}")
        
        # 步驟3: 如果需要拆分，執行拆分
        if analysis['needs_splitting']:
            success = self.word_splitter.split_document_by_chapters(
                word_file, 
                word_file.parent, 
                folder_name
            )
            
            if success:
                # 拆分成功後，可以選擇刪除原檔案或重命名
                original_backup = word_file.parent / f"original_{word_file.name}"
                word_file.rename(original_backup)
                self.logger.info(f"拆分成功，原檔案重命名為: {original_backup.name}")
        else:
            self.logger.info(f"不需要拆分: {word_file.name}")
    
    def process_multiple_word_files(self, folder_info):
        """處理多個Word檔案的情況"""
        if len(folder_info['word_files']) <= 1:
            return
        
        folder_name = folder_info['name']
        processed_files = []
        
        for word_file in folder_info['word_files']:
            try:
                # 分析是否需要拆分
                analysis = self.word_splitter.analyze_document(word_file)
                
                if analysis['needs_splitting']:
                    # 需要拆分的檔案
                    self.logger.info(f"檔案需要拆分: {word_file.name}")
                    base_name = f"{folder_name}_{word_file.stem}"
                    
                    success = self.word_splitter.split_document_by_chapters(
                        word_file,
                        word_file.parent,
                        base_name
                    )
                    
                    if success:
                        # 拆分成功後重命名原檔案
                        original_backup = word_file.parent / f"original_{word_file.name}"
                        word_file.rename(original_backup)
                        self.logger.info(f"拆分成功，原檔案重命名為: {original_backup.name}")
                        processed_files.append(original_backup)
                    else:
                        processed_files.append(word_file)
                else:
                    # 不需要拆分，直接重命名
                    new_name = f"{folder_name}_{word_file.name}"
                    new_path = word_file.parent / new_name
                    
                    if word_file.name != new_name:
                        word_file.rename(new_path)
                        self.logger.info(f"重命名檔案: {word_file.name} -> {new_name}")
                        processed_files.append(new_path)
                    else:
                        processed_files.append(word_file)
            
            except Exception as e:
                self.logger.error(f"處理檔案失敗 {word_file.name}: {e}")
                processed_files.append(word_file)
        
        folder_info['word_files'] = processed_files
    
    def generate_processing_report(self, processed_folders):
        """生成處理報告"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = self.log_dir / f"advanced_processing_report_{timestamp}.txt"
        
        total_folders = len(processed_folders)
        total_word_files = sum(len(folder['word_files']) for folder in processed_folders)
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("進階Word檔案處理報告\n")
            f.write("=" * 50 + "\n")
            f.write(f"處理時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"處理根資料夾數: {total_folders}\n")
            f.write(f"總Word檔案數: {total_word_files}\n\n")
            
            for folder in processed_folders:
                f.write(f"根資料夾: {folder['name']}\n")
                f.write(f"  Word檔案數量: {len(folder['word_files'])}\n")
                f.write(f"  處理類型: {'單一檔案處理' if len(folder['word_files']) <= 1 else '多檔案處理'}\n")
                f.write(f"  檔案列表:\n")
                
                for word_file in folder['word_files']:
                    relative_path = word_file.relative_to(folder['target_path'])
                    f.write(f"    - {relative_path}\n")
                f.write("\n")
            
        self.logger.info(f"處理報告已生成: {report_path}")
    
    def find_all_word_files_in_target(self):
        """在exam_03_wordsplitter中找到所有Word檔案，過濾original開頭的檔案"""
        word_files_with_folder = []
        
        if not self.target_dir.exists():
            self.logger.warning(f"目標資料夾不存在: {self.target_dir}")
            return word_files_with_folder
        
        for root, dirs, files in os.walk(self.target_dir):
            root_path = Path(root)
            for file in files:
                if file.lower().endswith(('.docx', '.doc')) and not file.startswith('~$') and not file.startswith('original_'):
                    word_file_path = root_path / file
                    # 獲取相對於target_dir的資料夾名稱
                    try:
                        relative_folder = root_path.relative_to(self.target_dir)
                        folder_name = str(relative_folder) if relative_folder != Path('.') else ""
                    except ValueError:
                        folder_name = ""
                    
                    word_files_with_folder.append((word_file_path, folder_name))
        
        return word_files_with_folder
    
    def extract_docx_info(self, word_file_path, folder_name):
        """從Word檔案提取所需資訊"""
        try:
            filename = word_file_path.name
            
            return {
                '資料夾': folder_name,
                'docx文件名': filename,
                '標題名稱': filename.replace('.docx', '').replace('.doc', ''),
                '題庫資源ID': "",  # 空白
                '資源創建時間': "",  # 空白
                '題庫編號': "",  # 空白
                '題庫創建時間': "",  # 空白
                '題庫標題': filename.replace('.docx', '').replace('.doc', ''),  # 同標題名稱
                '標題修改時間': "",  # 空白
                '識別完成保存時間': ""  # 空白
            }
            
        except Exception as e:
            self.logger.error(f"讀取 {word_file_path} 時發生錯誤: {e}")
            return None
    
    def create_docx_todolist(self):
        """創建docx檔案的待辦列表Excel檔案"""
        self.logger.info("開始創建docx todolist")
        
        # 找到所有Word檔案
        word_files_with_folder = self.find_all_word_files_in_target()
        
        if not word_files_with_folder:
            self.logger.info("在exam_03_wordsplitter中找不到任何Word檔案")
            return
        
        self.logger.info(f"找到 {len(word_files_with_folder)} 個Word檔案")
        
        # 創建輸出資料夾
        output_dir = self.base_dir / "exam_04_docx_todolist"
        if not output_dir.exists():
            output_dir.mkdir(parents=True)
            self.logger.info(f"已創建資料夾: {output_dir}")
        
        # 提取所有Word檔案資訊
        data_list = []
        for word_file_path, folder_name in word_files_with_folder:
            docx_info = self.extract_docx_info(word_file_path, folder_name)
            if docx_info:
                data_list.append(docx_info)
        
        if not data_list:
            self.logger.info("無法提取任何有效資料")
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
            self.logger.info("已按標題名稱進行自然排序（包含中文數字轉換）")
        
        # 生成時間戳檔案名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"docx_todolist_{timestamp}.xlsx"
        excel_path = output_dir / excel_filename
        
        # 儲存為Excel檔案
        try:
            df.to_excel(excel_path, index=False, engine='openpyxl')
            self.logger.info(f"已創建Word檔案待辦列表: {excel_path}")
            self.logger.info(f"包含 {len(data_list)} 筆資料")
            
            print(f"\n=== 創建Word檔案待辦列表完成 ===")
            print(f"檔案路徑: {excel_path}")
            print(f"檔案數量: {len(data_list)}")
            
            # 顯示前5筆資料作為預覽
            print("\n預覽前5筆資料:")
            print(df.head().to_string(index=False))
            
        except Exception as e:
            self.logger.error(f"儲存Excel檔案時發生錯誤: {e}")
    
    def process_all_folders(self):
        """處理所有根資料夾"""
        self.logger.info("開始進階Word檔案處理流程")
        
        # 步驟1: 找到包含Word檔案的根資料夾
        root_folders = self.find_root_folders_with_word_files()
        if not root_folders:
            self.logger.info("未找到包含Word檔案的根資料夾")
            return
        
        # 確保目標資料夾存在
        self.target_dir.mkdir(exist_ok=True)
        
        processed_folders = []
        
        for folder_info in root_folders:
            self.logger.info(f"\n處理根資料夾: {folder_info['name']}")
            
            # 步驟2: 複製根資料夾
            if self.copy_root_folder(folder_info):
                
                # 步驟3: 根據Word檔案數量決定處理方式
                if len(folder_info['word_files']) == 1:
                    self.logger.info("檢測到單一Word檔案，執行單一檔案處理流程")
                    self.process_single_word_file(folder_info)
                else:
                    self.logger.info(f"檢測到 {len(folder_info['word_files'])} 個Word檔案，執行多檔案處理流程")
                    self.process_multiple_word_files(folder_info)
                
                processed_folders.append(folder_info)
            else:
                self.logger.error(f"跳過根資料夾: {folder_info['name']} (複製失敗)")
        
        # 步驟4: 生成處理報告
        if processed_folders:
            self.generate_processing_report(processed_folders)
        
        # 步驟5: 創建docx todolist
        self.create_docx_todolist()
        
        self.logger.info("進階Word檔案處理流程完成")

def main():
    processor = AdvancedWordProcessor()
    processor.process_all_folders()

if __name__ == "__main__":
    main()