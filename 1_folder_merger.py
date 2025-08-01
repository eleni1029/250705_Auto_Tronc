#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ZIP 檔案解壓縮工具
按檔案名稱排序，逐一解壓縮 01_ori_zipfiles/ 目錄下的 .zip 檔案到 02_merged_projects/
支援覆蓋模式解壓縮
"""

import os
import shutil
import logging
import zipfile
from pathlib import Path
from datetime import datetime
from typing import List, Optional


class ZipExtractor:
    """ZIP 檔案解壓縮器類別"""
    
    def __init__(self, source_dir: str = "01_ori_zipfiles", target_dir: str = "02_merged_projects"):
        """
        初始化解壓縮器
        
        Args:
            source_dir: 來源目錄路徑，包含 ZIP 檔案（預設: "01_ori_zipfiles"）
            target_dir: 目標目錄路徑，解壓縮目的地（預設: "02_merged_projects"）
        """
        self.source_dir = Path(source_dir)
        self.target_dir = Path(target_dir)
        self.stats = {
            'zip_files_processed': 0,
            'files_extracted': 0,
            'folders_created': 0,
            'errors': 0
        }
        
        # 設定日誌
        self._setup_logging()
    
    def _setup_logging(self):
        """設定日誌系統"""
        # 確保 log 資料夾存在
        log_dir = Path("log")
        log_dir.mkdir(exist_ok=True)
        
        log_filename = log_dir / f"merge_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        # 建立日誌格式
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename, encoding='utf-8'),
                logging.StreamHandler()  # 同時輸出到控制台
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"=== 開始 ZIP 解壓縮作業 ===")
        self.logger.info(f"來源目錄: {self.source_dir.absolute()}")
        self.logger.info(f"目標目錄: {self.target_dir.absolute()}")
        self.logger.info(f"日誌檔案: {log_filename}")
    
    def _validate_directories(self) -> bool:
        """驗證目錄狀態並獲取 ZIP 檔案列表"""
        if not self.source_dir.exists():
            self.logger.error(f"來源目錄不存在: {self.source_dir}")
            return False
        
        if not self.source_dir.is_dir():
            self.logger.error(f"來源路徑不是目錄: {self.source_dir}")
            return False
        
        # 獲取 ZIP 檔案列表並按名稱排序
        zip_files = sorted([f for f in self.source_dir.iterdir() if f.suffix.lower() == '.zip'])
        if not zip_files:
            self.logger.warning(f"來源目錄下沒有 ZIP 檔案: {self.source_dir}")
            return False
        
        self.logger.info(f"找到 {len(zip_files)} 個 ZIP 檔案待解壓縮")
        for zip_file in zip_files:
            self.logger.info(f"  - {zip_file.name}")
        
        return True
    
    def _create_target_directory(self):
        """建立目標目錄"""
        try:
            self.target_dir.mkdir(parents=True, exist_ok=True)
            self.logger.info(f"目標目錄已準備: {self.target_dir}")
        except Exception as e:
            self.logger.error(f"無法建立目標目錄: {e}")
            raise
    
    def _extract_zip_file(self, zip_path: Path) -> bool:
        """
        解壓縮 ZIP 檔案到目標目錄
        
        Args:
            zip_path: ZIP 檔案路徑
            
        Returns:
            bool: 解壓縮是否成功
        """
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                # 獲取 ZIP 檔案中的所有檔案列表
                file_list = zip_ref.namelist()
                self.logger.info(f"開始解壓縮: {zip_path.name} (包含 {len(file_list)} 個檔案)")
                
                # 解壓縮所有檔案，支援覆蓋
                for file_info in zip_ref.infolist():
                    try:
                        # 解壓縮檔案
                        zip_ref.extract(file_info, self.target_dir)
                        
                        if file_info.is_dir():
                            self.stats['folders_created'] += 1
                        else:
                            self.stats['files_extracted'] += 1
                            
                    except Exception as e:
                        self.stats['errors'] += 1
                        self.logger.error(f"解壓縮檔案失敗 {file_info.filename}: {e}")
                        
            self.logger.info(f"完成解壓縮: {zip_path.name}")
            return True
            
        except Exception as e:
            self.stats['errors'] += 1
            self.logger.error(f"無法解壓縮 ZIP 檔案 {zip_path}: {e}")
            return False
    
    def _get_sorted_zip_files(self) -> List[Path]:
        """
        獲取排序後的 ZIP 檔案列表
        
        Returns:
            List[Path]: 按檔案名稱排序的 ZIP 檔案列表
        """
        zip_files = [f for f in self.source_dir.iterdir() if f.suffix.lower() == '.zip']
        return sorted(zip_files, key=lambda x: x.name)
    
    def extract_zip_files(self) -> bool:
        """
        執行 ZIP 檔案解壓縮
        
        Returns:
            bool: 解壓縮是否成功完成
        """
        try:
            # 驗證目錄
            if not self._validate_directories():
                return False
            
            # 建立目標目錄
            self._create_target_directory()
            
            # 獲取排序後的 ZIP 檔案列表
            zip_files = self._get_sorted_zip_files()
            
            # 逐一解壓縮 ZIP 檔案
            for zip_file in zip_files:
                self.logger.info(f"開始處理 ZIP 檔案: {zip_file.name}")
                self.stats['zip_files_processed'] += 1
                
                # 解壓縮當前 ZIP 檔案
                success = self._extract_zip_file(zip_file)
                
                if success:
                    self.logger.info(f"成功處理 ZIP 檔案: {zip_file.name}")
                else:
                    self.logger.error(f"處理 ZIP 檔案失敗: {zip_file.name}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"解壓縮過程發生錯誤: {e}")
            return False
    
    def print_summary(self):
        """輸出解壓縮摘要"""
        self.logger.info("=== ZIP 解壓縮作業完成 ===")
        self.logger.info(f"處理 ZIP 檔案數: {self.stats['zip_files_processed']}")
        self.logger.info(f"解壓縮檔案數: {self.stats['files_extracted']}")
        self.logger.info(f"建立資料夾數: {self.stats['folders_created']}")
        self.logger.info(f"錯誤數量: {self.stats['errors']}")
        self.logger.info(f"總處理項目數: {self.stats['files_extracted'] + self.stats['folders_created']}")


def main():
    """主函數"""
    # 設定來源和目標目錄
    SOURCE_DIR = "01_ori_zipfiles"  # 可以修改為您的實際路徑
    TARGET_DIR = "02_merged_projects"  # 可以修改為您希望的目標路徑
    
    # 建立 ZIP 解壓縮器實例
    extractor = ZipExtractor(SOURCE_DIR, TARGET_DIR)
    
    # 執行 ZIP 解壓縮
    success = extractor.extract_zip_files()
    
    # 輸出摘要
    extractor.print_summary()
    
    if success:
        print(f"\n✅ ZIP 解壓縮完成！結果已保存到: {TARGET_DIR}")
    else:
        print(f"\n❌ ZIP 解壓縮過程中發生錯誤，請檢查日誌檔案")
    
    return success


if __name__ == "__main__":
    main()