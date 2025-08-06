#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
試題驗證流程工具
從指定資料夾提取包含試題內容的專案資料夾（選擇性複製）
"""

import os
import re
import shutil
import logging
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Set


class ExamExtractor:
    """試題提取器類別"""
    
    def __init__(self, source_dir: str = "02_merged_projects"):
        """
        初始化試題提取器
        
        Args:
            source_dir: 來源目錄路徑（預設: "02_merged_projects"）
        """
        self.source_dir = Path(source_dir)
        self.target_dir = None
        self.stats = {
            'folders_scanned': 0,
            'exam_folders_found': 0,
            'folders_without_exams': 0,
            'files_copied': 0,
            'folders_copied': 0,
            'errors': 0
        }
        
        # 設定日誌
        self._setup_logging()
        
        # 試題關鍵字
        self.exam_keywords = ["題庫", "評量", "試題", "測驗"]
        self.exam_extensions = [".txt", ".doc", ".docx", ".pdf", ".xls", ".xlsx"]
        
        self.logger.info(f"試題提取器初始化完成 - 來源目錄: {self.source_dir}")
    
    def _setup_logging(self):
        """設定日誌系統"""
        # 確保 log 資料夾存在
        log_dir = Path("log")
        log_dir.mkdir(exist_ok=True)
        
        log_filename = log_dir / f"exam_extract_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        # 建立日誌格式
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def select_source_directory(self) -> str:
        """
        讓用戶選擇來源目錄
        
        Returns:
            選擇的目錄路徑
        """
        print(f"目前預設來源目錄: {self.source_dir}")
        print("請選擇來源目錄:")
        print("0 - 使用預設目錄")
        print("1 - 選擇其他目錄")
        
        while True:
            choice = input("請輸入選項 (0/1): ").strip()
            
            if choice == '0':
                print(f"使用預設來源目錄: {self.source_dir}")
                break
            elif choice == '1':
                new_source = input("請輸入來源目錄路徑: ").strip()
                if new_source and Path(new_source).exists():
                    self.source_dir = Path(new_source)
                    self.logger.info(f"來源目錄已更改為: {self.source_dir}")
                    print(f"來源目錄已設定為: {self.source_dir}")
                    break
                else:
                    print("無效的目錄路徑，請重新選擇")
            else:
                print("無效選項，請輸入 0 或 1")
        
        return str(self.source_dir)
    
    def set_target_directory(self, target_name: str = None) -> str:
        """
        設定目標目錄
        
        Args:
            target_name: 目標目錄名稱，預設為 exam_01_[來源目錄名稱]
            
        Returns:
            目標目錄路徑
        """
        if target_name is None:
            target_name = f'exam_01_{self.source_dir.name}'
        
        self.target_dir = Path(target_name)
        
        # 如果目標目錄已存在，詢問是否覆蓋
        if self.target_dir.exists():
            print(f"目標目錄 '{self.target_dir}' 已存在")
            print("請選擇處理方式:")
            print("0 - 覆蓋現有目錄")
            print("1 - 創建新的時間戳目錄")
            
            while True:
                choice = input("請輸入選項 (0/1): ").strip()
                
                if choice == '0':
                    shutil.rmtree(self.target_dir)
                    self.logger.info(f"已刪除現有目標目錄: {self.target_dir}")
                    print(f"已覆蓋目標目錄: {self.target_dir}")
                    break
                elif choice == '1':
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    original_name = target_name if target_name else f'exam_01_{self.source_dir.name}'
                    self.target_dir = Path(f"{original_name}_{timestamp}")
                    self.logger.info(f"目標目錄已更改為: {self.target_dir}")
                    print(f"目標目錄已設定為: {self.target_dir}")
                    break
                else:
                    print("無效選項，請輸入 0 或 1")
        
        # 創建目標目錄
        self.target_dir.mkdir(exist_ok=True)
        self.logger.info(f"目標目錄已創建: {self.target_dir}")
        
        return str(self.target_dir)
    
    def create_target_folder_structure(self):
        """根據來源目錄的第一層資料夾創建目標資料夾結構"""
        if not self.source_dir.exists():
            raise ValueError(f"來源目錄不存在: {self.source_dir}")
        
        self.logger.info("開始創建目標資料夾結構...")
        
        # 獲取第一層資料夾
        first_level_folders = [
            item for item in self.source_dir.iterdir()
            if item.is_dir()
        ]
        
        # 在目標目錄中創建對應的空資料夾
        for folder in first_level_folders:
            target_folder = self.target_dir / folder.name
            target_folder.mkdir(exist_ok=True)
            self.logger.info(f"已創建目標資料夾: {target_folder}")
        
        self.logger.info(f"目標資料夾結構創建完成，共創建 {len(first_level_folders)} 個資料夾")
        return first_level_folders
    
    def is_scorm_xml_with_exam(self, xml_path: Path) -> bool:
        """
        檢查 XML 檔案是否為 SCORM 格式且包含試題類型
        
        Args:
            xml_path: XML 檔案路徑
            
        Returns:
            是否為符合條件的 SCORM XML
        """
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # 檢查是否為 SCORM manifest
            if 'manifest' not in root.tag.lower():
                return False
            
            # 檢查是否包含試題相關的內容
            xml_content = ET.tostring(root, encoding='unicode', method='xml').lower()
            
            # 檢查試題相關關鍵字（更嚴格的檢測）
            exam_indicators = [
                'assessment', 'quiz', 'test', 'exam', 'examination',
                '試題', '測驗', '評量', '考試'
            ]
            
            # 排除一般性詞彙，避免誤判
            excluded_patterns = [
                '問題', 'question',  # 太泛泛，容易誤判
                '練習', 'practice'   # 練習不等於試題
            ]
            
            for indicator in exam_indicators:
                if indicator in xml_content:
                    # 額外檢查是否真的是試題相關，而不是一般性描述
                    # 檢查是否在試題相關的結構化內容中
                    if self._is_truly_exam_content(xml_content, indicator):
                        self.logger.info(f"找到 SCORM 試題 XML: {xml_path} (關鍵字: {indicator})")
                        return True
            
            return False
            
        except ET.ParseError as e:
            self.logger.warning(f"XML 解析錯誤 {xml_path}: {e}")
            return False
        except Exception as e:
            self.logger.error(f"檢查 XML 檔案時發生錯誤 {xml_path}: {e}")
            return False
    
    def _is_truly_exam_content(self, xml_content: str, indicator: str) -> bool:
        """
        更嚴格地檢查是否真的是試題內容
        
        Args:
            xml_content: XML 內容（小寫）
            indicator: 找到的關鍵字
            
        Returns:
            是否真的是試題內容
        """
        # 檢查是否在試題相關的 XML 結構中
        exam_context_patterns = [
            'adlcp:datafromlms',  # SCORM 數據交互
            'adlcp:masteryscore',  # 掌握分數
            'adlcp:prerequisites',  # 前置要求
            'objectives',  # 學習目標（試題相關）
            'interactions',  # 互動（試題相關）
            'correct_responses',  # 正確答案
            'learner_response',  # 學習者回應
            'score'  # 分數
        ]
        
        # 如果關鍵字出現在試題相關的結構中，才認為是真正的試題
        for pattern in exam_context_patterns:
            if pattern in xml_content and indicator in xml_content:
                context_start = max(0, xml_content.find(pattern) - 200)
                context_end = min(len(xml_content), xml_content.find(pattern) + 200)
                context = xml_content[context_start:context_end]
                
                if indicator in context:
                    return True
        
        # 如果只是在標題中出現，可能不是真正的試題
        # 進一步檢查是否在適當的上下文中
        return False
    
    def _is_in_exam_folder(self, xml_path: Path) -> bool:
        """
        檢查 XML 檔案是否位於明確的試題資料夾中
        
        Args:
            xml_path: XML 檔案路徑
            
        Returns:
            是否位於試題資料夾中
        """
        # 檢查 XML 檔案的路徑中是否包含明確的試題資料夾名稱
        path_parts = [part.lower() for part in xml_path.parts]
        
        exam_folder_names = [
            'exam', 'exams', 'quiz', 'test', 'assessment', 'evaluation',
            '試題', '測驗', '評量', '考試', '題庫', '自我評量'
        ]
        
        # 檢查路徑中是否有任何部分是試題資料夾
        for part in path_parts:
            if any(exam_name in part for exam_name in exam_folder_names):
                return True
        
        return False
    
    
    def copy_files_only(self, source_folder: Path, target_folder: Path):
        """
        複製來源資料夾中除了資料夾以外的所有文件到目標資料夾
        （只複製該資料夾直接包含的文件，不遞迴進入子資料夾）
        
        Args:
            source_folder: 來源資料夾路徑
            target_folder: 目標資料夾路徑
        """
        try:
            # 確保目標資料夾存在
            target_folder.mkdir(parents=True, exist_ok=True)
            
            files_copied_count = 0
            
            # 只複製該資料夾中的直接文件，不進入子資料夾
            for item in source_folder.iterdir():
                if item.is_file():
                    target_file = target_folder / item.name
                    
                    # 如果檔案已存在，先刪除舊的
                    if target_file.exists():
                        target_file.unlink()
                    
                    # 複製檔案
                    shutil.copy2(item, target_file)
                    files_copied_count += 1
                    self.stats['files_copied'] += 1
                    self.logger.debug(f"複製檔案: {item} -> {target_file}")
            
            self.logger.info(f"已複製 {files_copied_count} 個檔案（僅該資料夾中的文件）: {source_folder} -> {target_folder}")
            
        except Exception as e:
            self.logger.error(f"複製檔案時發生錯誤: {e}")
            self.stats['errors'] += 1
    
    def copy_folder_structure(self, source_item: Path, source_root: Path, target_root: Path):
        """
        按照結構複製資料夾到目標資料夾（遇到相同檔案直接覆蓋）
        
        Args:
            source_item: 要複製的來源項目路徑
            source_root: 來源根目錄
            target_root: 目標根目錄
        """
        try:
            # 計算相對路徑
            rel_path = source_item.relative_to(source_root)
            target_item = target_root / rel_path
            
            if source_item.is_file():
                # 確保目標檔案的父目錄存在
                target_item.parent.mkdir(parents=True, exist_ok=True)
                
                # 如果檔案已存在，直接覆蓋
                if target_item.exists():
                    target_item.unlink()
                    self.logger.debug(f"覆蓋現有檔案: {target_item}")
                
                # 複製檔案
                shutil.copy2(source_item, target_item)
                self.stats['files_copied'] += 1
                self.logger.debug(f"複製檔案: {source_item} -> {target_item}")
                
            elif source_item.is_dir():
                # 複製整個資料夾結構（允許覆蓋）
                if target_item.exists():
                    shutil.rmtree(target_item)
                    self.logger.debug(f"移除現有資料夾: {target_item}")
                
                shutil.copytree(source_item, target_item, dirs_exist_ok=True)
                self.stats['folders_copied'] += 1
                
                # 計算複製的檔案數量
                for _ in target_item.rglob("*"):
                    if _.is_file():
                        self.stats['files_copied'] += 1
                
                self.logger.debug(f"複製資料夾: {source_item} -> {target_item}")
            
            self.logger.info(f"已複製結構: {rel_path}")
            
        except Exception as e:
            self.logger.error(f"複製結構時發生錯誤: {e}")
            self.stats['errors'] += 1
    
    
    def rename_folder_without_exams(self, folder_path: Path):
        """
        為沒有試題內容的資料夾名稱添加 '_無題庫' 後綴
        
        Args:
            folder_path: 資料夾路徑
        """
        try:
            new_name = f"{folder_path.name}_無題庫"
            new_path = folder_path.parent / new_name
            
            if not new_path.exists():
                folder_path.rename(new_path)
                self.logger.info(f"資料夾已重新命名: {folder_path} -> {new_path}")
            else:
                self.logger.warning(f"目標路徑已存在，無法重新命名: {new_path}")
                
        except Exception as e:
            self.logger.error(f"重新命名資料夾時發生錯誤: {e}")
            self.stats['errors'] += 1
    
    def extract_exams(self):
        """執行試題提取流程 - 分別執行三個條件，各自獨立處理"""
        if not self.target_dir:
            raise ValueError("目標目錄未設定，請先調用 set_target_directory()")
        
        self.logger.info("開始執行試題提取流程（三條件分離執行模式）...")
        
        # 創建目標資料夾結構
        first_level_folders = self.create_target_folder_structure()
        
        # 掃描每個第一層資料夾
        for source_folder in first_level_folders:
            self.stats['folders_scanned'] += 1
            folder_name = source_folder.name
            target_folder = self.target_dir / folder_name
            
            self.logger.info(f"處理資料夾: {folder_name}")
            
            # 分別檢查三個條件
            condition_a_triggered = self.process_condition_a(source_folder, target_folder)
            condition_b_triggered = self.process_condition_b(source_folder, target_folder)  
            condition_c_triggered = self.process_condition_c(source_folder, target_folder)
            
            # 判斷是否有任何條件被觸發
            if condition_a_triggered or condition_b_triggered or condition_c_triggered:
                self.stats['exam_folders_found'] += 1
                self.logger.info(f"✅ {folder_name}: 找到試題內容")
                
                # 顯示觸發的條件
                triggered_conditions = []
                if condition_a_triggered:
                    triggered_conditions.append("條件a(SCORM XML)")
                if condition_b_triggered:
                    triggered_conditions.append("條件b(exam資料夾)")
                if condition_c_triggered:
                    triggered_conditions.append("條件c(試題文件/資料夾)")
                
                self.logger.info(f"   觸發條件: {', '.join(triggered_conditions)}")
            else:
                # 標記為無題庫
                self.rename_folder_without_exams(target_folder)
                self.stats['folders_without_exams'] += 1
                self.logger.info(f"❌ {folder_name}: 未找到試題內容，已標記為_無題庫")
    
    def process_condition_a(self, source_folder: Path, target_folder: Path) -> bool:
        """
        條件 a: 處理 SCORM 格式的試題 XML 檔案
        複製所在資料夾除了資料夾以外的所有文件到目標資料夾
        
        Returns:
            是否觸發此條件
        """
        triggered = False
        scorm_xmls = []
        
        # 尋找 SCORM 試題 XML
        for xml_file in source_folder.rglob("*.xml"):
            # 只檢查位於明確試題資料夾中的 XML 檔案
            if self._is_in_exam_folder(xml_file) and self.is_scorm_xml_with_exam(xml_file):
                scorm_xmls.append(xml_file)
        
        if scorm_xmls:
            triggered = True
            self.logger.info(f"條件 a 觸發: 找到 {len(scorm_xmls)} 個位於試題資料夾中的 SCORM 試題 XML")
            
            for scorm_xml in scorm_xmls:
                xml_folder = scorm_xml.parent
                self.logger.info(f"處理 SCORM XML: {scorm_xml}")
                self.logger.info(f"複製資料夾中的文件: {xml_folder.relative_to(source_folder)}")
                self.copy_files_only(xml_folder, target_folder)
        
        return triggered
    
    def process_condition_b(self, source_folder: Path, target_folder: Path) -> bool:
        """
        條件 b: 處理名稱為 exam 的資料夾
        按照結構複製到目標資料夾
        
        Returns:
            是否觸發此條件
        """
        triggered = False
        
        # 尋找 exam 資料夾
        for item in source_folder.rglob("*"):
            if item.is_dir() and item.name.lower() == "exam":
                triggered = True
                self.logger.info(f"條件 b 觸發: 找到 exam 資料夾: {item.relative_to(source_folder)}")
                self.copy_folder_structure(item, source_folder, target_folder)
        
        return triggered
    
    def process_condition_c(self, source_folder: Path, target_folder: Path) -> bool:
        """
        條件 c: 處理包含試題關鍵字的文件或資料夾
        按照結構複製到目標資料夾
        
        Returns:
            是否觸發此條件
        """
        triggered = False
        
        def check_path(path: Path) -> bool:
            """檢查路徑名稱是否包含試題關鍵字"""
            path_name = path.name.lower()
            return any(keyword in path_name for keyword in self.exam_keywords)
        
        # 尋找符合條件的文件和資料夾
        for item in source_folder.rglob("*"):
            # 檢查資料夾名稱
            if item.is_dir() and check_path(item):
                triggered = True
                self.logger.info(f"條件 c 觸發: 找到試題資料夾: {item.relative_to(source_folder)}")
                self.copy_folder_structure(item, source_folder, target_folder)
            
            # 檢查檔案名稱和副檔名
            elif item.is_file():
                if (check_path(item) and 
                    item.suffix.lower() in self.exam_extensions):
                    triggered = True
                    self.logger.info(f"條件 c 觸發: 找到試題文件: {item.relative_to(source_folder)}")
                    self.copy_folder_structure(item, source_folder, target_folder)
        
        return triggered
    
    def print_summary(self):
        """印出處理結果摘要"""
        print("\n" + "="*60)
        print("試題提取處理結果摘要")
        print("="*60)
        print(f"來源目錄: {self.source_dir}")
        print(f"目標目錄: {self.target_dir}")
        print(f"掃描的資料夾總數: {self.stats['folders_scanned']}")
        print(f"找到試題內容的資料夾: {self.stats['exam_folders_found']}")
        print(f"沒有試題內容的資料夾: {self.stats['folders_without_exams']}")
        print(f"複製的檔案總數: {self.stats['files_copied']}")
        print(f"複製的資料夾總數: {self.stats['folders_copied']}")
        print(f"處理錯誤次數: {self.stats['errors']}")
        print("="*60)
        
        if self.stats['folders_scanned'] > 0:
            success_rate = (self.stats['exam_folders_found'] / self.stats['folders_scanned']) * 100
            print(f"試題資料夾發現率: {success_rate:.1f}%")
        
        # 記錄到日誌
        self.logger.info("處理完成")
        self.logger.info(f"統計結果: {self.stats}")
    
    def run(self):
        """主執行流程"""
        try:
            print("="*60)
            print("試題驗證流程工具（選擇性複製版）")
            print("="*60)
            
            # 步驟 1: 選擇來源目錄
            self.select_source_directory()
            
            # 步驟 2: 設定目標目錄
            print(f"\n目標目錄設定:")
            print(f'預設目標目錄名稱: exam_01_{self.source_dir.name}')
            print("請選擇目標目錄設定:")
            print("0 - 使用預設目標目錄名稱")
            print("1 - 輸入自定義目標目錄名稱")
            
            target_name = None
            while True:
                choice = input("請輸入選項 (0/1): ").strip()
                
                if choice == '0':
                    target_name = None  # 使用預設名稱
                    print(f'使用預設目標目錄名稱: exam_01_{self.source_dir.name}')
                    break
                elif choice == '1':
                    target_name = input("請輸入目標目錄名稱: ").strip()
                    if target_name:
                        print(f"目標目錄名稱已設定為: {target_name}")
                        break
                    else:
                        print("目標目錄名稱不能為空，請重新輸入")
                else:
                    print("無效選項，請輸入 0 或 1")
            
            self.set_target_directory(target_name)
            
            # 步驟 3: 執行試題提取
            self.extract_exams()
            
            # 步驟 4: 顯示結果
            self.print_summary()
            
        except KeyboardInterrupt:
            print("\n用戶中斷操作")
            self.logger.info("用戶中斷操作")
        except Exception as e:
            print(f"\n執行過程中發生錯誤: {e}")
            self.logger.error(f"執行錯誤: {e}")
        finally:
            print("\n程式結束")


def main():
    """主程式進入點"""
    extractor = ExamExtractor()
    extractor.run()


if __name__ == "__main__":
    main()