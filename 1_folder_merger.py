#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
項目資料夾合併工具
合併 1_projects/ 目錄下的所有子資料夾到一個新的目標資料夾
支援深度遞迴合併，同名檔案採用覆蓋策略
"""

import os
import shutil
import logging
from pathlib import Path
from datetime import datetime
from typing import List, Optional


class ProjectMerger:
    """項目合併器類別"""
    
    def __init__(self, source_dir: str = "1_projects", target_dir: str = "2_merged_projects"):
        """
        初始化合併器
        
        Args:
            source_dir: 來源目錄路徑（預設: "1_projects"）
            target_dir: 目標目錄路徑（預設: "2_merged_projects"）
        """
        self.source_dir = Path(source_dir)
        self.target_dir = Path(target_dir)
        self.stats = {
            'folders_processed': 0,
            'files_copied': 0,
            'files_overwritten': 0,
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
        self.logger.info(f"=== 開始合併作業 ===")
        self.logger.info(f"來源目錄: {self.source_dir.absolute()}")
        self.logger.info(f"目標目錄: {self.target_dir.absolute()}")
        self.logger.info(f"日誌檔案: {log_filename}")
    
    def _validate_directories(self) -> bool:
        """驗證目錄狀態"""
        if not self.source_dir.exists():
            self.logger.error(f"來源目錄不存在: {self.source_dir}")
            return False
        
        if not self.source_dir.is_dir():
            self.logger.error(f"來源路徑不是目錄: {self.source_dir}")
            return False
        
        # 獲取子資料夾列表
        subdirs = [d for d in self.source_dir.iterdir() if d.is_dir()]
        if not subdirs:
            self.logger.warning(f"來源目錄下沒有子資料夾: {self.source_dir}")
            return False
        
        self.logger.info(f"找到 {len(subdirs)} 個子資料夾待合併")
        for subdir in subdirs:
            self.logger.info(f"  - {subdir.name}")
        
        return True
    
    def _create_target_directory(self):
        """建立目標目錄"""
        try:
            self.target_dir.mkdir(parents=True, exist_ok=True)
            self.logger.info(f"目標目錄已準備: {self.target_dir}")
        except Exception as e:
            self.logger.error(f"無法建立目標目錄: {e}")
            raise
    
    def _copy_file_with_overwrite(self, src_file: Path, dst_file: Path):
        """
        複製檔案，支援覆蓋
        
        Args:
            src_file: 來源檔案路徑
            dst_file: 目標檔案路徑
        """
        try:
            # 確保目標目錄存在
            dst_file.parent.mkdir(parents=True, exist_ok=True)
            
            # 檢查是否為覆蓋操作
            is_overwrite = dst_file.exists()
            
            # 複製檔案
            shutil.copy2(src_file, dst_file)
            
            if is_overwrite:
                self.stats['files_overwritten'] += 1
                self.logger.info(f"覆蓋檔案: {dst_file.relative_to(self.target_dir)}")
            else:
                self.stats['files_copied'] += 1
                self.logger.info(f"複製檔案: {dst_file.relative_to(self.target_dir)}")
                
        except Exception as e:
            self.stats['errors'] += 1
            self.logger.error(f"複製檔案失敗 {src_file} -> {dst_file}: {e}")
    
    def _merge_directory(self, src_dir: Path, dst_dir: Path):
        """
        遞迴合併目錄
        
        Args:
            src_dir: 來源目錄
            dst_dir: 目標目錄
        """
        try:
            for item in src_dir.iterdir():
                src_path = src_dir / item.name
                dst_path = dst_dir / item.name
                
                if src_path.is_file():
                    # 處理檔案
                    self._copy_file_with_overwrite(src_path, dst_path)
                
                elif src_path.is_dir():
                    # 處理子目錄（遞迴合併）
                    self.logger.info(f"處理子目錄: {src_path.relative_to(self.source_dir)}")
                    self._merge_directory(src_path, dst_path)
                    
        except Exception as e:
            self.stats['errors'] += 1
            self.logger.error(f"合併目錄失敗 {src_dir}: {e}")
    
    def merge_projects(self) -> bool:
        """
        執行項目合併
        
        Returns:
            bool: 合併是否成功完成
        """
        try:
            # 驗證目錄
            if not self._validate_directories():
                return False
            
            # 建立目標目錄
            self._create_target_directory()
            
            # 開始合併每個子資料夾
            subdirs = [d for d in self.source_dir.iterdir() if d.is_dir()]
            
            for subdir in subdirs:
                self.logger.info(f"開始處理資料夾: {subdir.name}")
                self.stats['folders_processed'] += 1
                
                # 合併當前子目錄到目標目錄
                self._merge_directory(subdir, self.target_dir)
                
                self.logger.info(f"完成處理資料夾: {subdir.name}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"合併過程發生錯誤: {e}")
            return False
    
    def print_summary(self):
        """輸出合併摘要"""
        self.logger.info("=== 合併作業完成 ===")
        self.logger.info(f"處理資料夾數: {self.stats['folders_processed']}")
        self.logger.info(f"複製檔案數: {self.stats['files_copied']}")
        self.logger.info(f"覆蓋檔案數: {self.stats['files_overwritten']}")
        self.logger.info(f"錯誤數量: {self.stats['errors']}")
        self.logger.info(f"總檔案處理數: {self.stats['files_copied'] + self.stats['files_overwritten']}")


def main():
    """主函數"""
    # 設定來源和目標目錄
    SOURCE_DIR = "1_projects"  # 可以修改為您的實際路徑
    TARGET_DIR = "2_merged_projects"  # 可以修改為您希望的目標路徑
    
    # 建立合併器實例
    merger = ProjectMerger(SOURCE_DIR, TARGET_DIR)
    
    # 執行合併
    success = merger.merge_projects()
    
    # 輸出摘要
    merger.print_summary()
    
    if success:
        print(f"\n✅ 合併完成！結果已保存到: {TARGET_DIR}")
    else:
        print(f"\n❌ 合併過程中發生錯誤，請檢查日誌檔案")
    
    return success


if __name__ == "__main__":
    main()