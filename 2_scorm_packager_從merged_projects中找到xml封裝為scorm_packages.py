#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SCORM 自動打包工具
自動搜尋包含 imsmanifest 的 XML 檔案並將其所在資料夾打包為 SCORM 壓縮包
支援多層級處理和用戶互動選擇
包含 manifest 檔案標準化重命名功能
"""

import os
import zipfile
import logging
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple, Optional
from collections import defaultdict
import shutil


class ScormPackager:
    """SCORM 打包器類別"""
    
    def __init__(self, source_dir: str, output_dir: str = "scorm_packages"):
        """
        初始化打包器
        
        Args:
            source_dir: 來源目錄路徑
            output_dir: 輸出目錄路徑（預設: "scorm_packages"）
        """
        self.source_dir = Path(source_dir)
        self.output_dir = Path(output_dir)
        self.stats = {
            'directories_scanned': 0,
            'manifests_found': 0,
            'packages_created': 0,
            'files_packaged': 0,
            'conflicts_resolved': 0,
            'errors': 0,
            'skipped_files': 0,
            'manifests_renamed': 0,
            'manifests_backed_up': 0
        }
        
        # 記錄找到的 manifest 檔案
        self.manifest_locations: List[Tuple[Path, Path]] = []  # (xml_file_path, containing_directory)
        self.package_results: List[Dict] = []
        
        # 設定日誌
        self._setup_logging()
    
    def _setup_logging(self):
        """設定日誌系統"""
        log_filename = self.output_dir / f"scorm_package_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

        # 確保 log 資料夾存在（這裡 output_dir 是 scorm_packages）
        self.output_dir.mkdir(parents=True, exist_ok=True)

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("=== 開始 SCORM 打包作業 ===")
        self.logger.info(f"來源目錄: {self.source_dir.absolute()}")
        self.logger.info(f"輸出目錄: {self.output_dir.absolute()}")
        self.logger.info(f"日誌檔案: {log_filename}")
    
    def _is_manifest_file(self, filename: str) -> bool:
        """
        檢查檔案是否為 manifest 檔案
        
        Args:
            filename: 檔案名稱
            
        Returns:
            bool: 是否為 manifest 檔案
        """
        return (filename.lower().endswith('.xml') and 
                'imsmanifest' in filename.lower())
    
    def _validate_source_directory(self) -> bool:
        """驗證來源目錄狀態"""
        if not self.source_dir.exists():
            self.logger.error(f"來源目錄不存在: {self.source_dir}")
            return False
        
        if not self.source_dir.is_dir():
            self.logger.error(f"來源路徑不是目錄: {self.source_dir}")
            return False
        
        self.logger.info(f"來源目錄驗證成功: {self.source_dir}")
        return True
    
    def _create_output_directory(self):
        """建立輸出目錄"""
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            self.logger.info(f"輸出目錄已準備: {self.output_dir}")
        except Exception as e:
            self.logger.error(f"無法建立輸出目錄: {e}")
            raise
    
    def _scan_for_manifests(self) -> Dict[Path, List[Path]]:
        """
        遞迴掃描尋找所有 manifest 檔案
        
        Returns:
            Dict[Path, List[Path]]: {包含目錄: [manifest檔案列表]}
        """
        manifest_by_directory = defaultdict(list)
        
        def _recursive_scan(current_dir: Path):
            try:
                self.stats['directories_scanned'] += 1
                manifest_files_in_dir = []
                
                # 檢查當前目錄下的檔案
                for item in current_dir.iterdir():
                    if item.is_file() and self._is_manifest_file(item.name):
                        manifest_files_in_dir.append(item)
                        self.stats['manifests_found'] += 1
                        self.logger.info(f"找到 manifest 檔案: {item.relative_to(self.source_dir)}")
                
                # 如果當前目錄有 manifest 檔案，記錄之
                if manifest_files_in_dir:
                    manifest_by_directory[current_dir] = manifest_files_in_dir
                
                # 遞迴處理子目錄
                for item in current_dir.iterdir():
                    if item.is_dir():
                        _recursive_scan(item)
                        
            except PermissionError:
                self.stats['errors'] += 1
                self.logger.warning(f"無權限存取目錄: {current_dir}")
            except Exception as e:
                self.stats['errors'] += 1
                self.logger.error(f"掃描目錄時發生錯誤 {current_dir}: {e}")
        
        _recursive_scan(self.source_dir)
        return dict(manifest_by_directory)
    
    def _resolve_conflicts(self, manifest_by_directory: Dict[Path, List[Path]]) -> Dict[Path, Path]:
        """
        解決同層級多個 manifest 檔案的衝突
        
        Args:
            manifest_by_directory: {包含目錄: [manifest檔案列表]}
            
        Returns:
            Dict[Path, Path]: {包含目錄: 選定的manifest檔案}
        """
        resolved_manifests = {}
        
        for directory, manifest_files in manifest_by_directory.items():
            if len(manifest_files) == 1:
                # 只有一個檔案，直接使用
                resolved_manifests[directory] = manifest_files[0]
                self.logger.info(f"目錄 {directory.relative_to(self.source_dir)} 使用檔案: {manifest_files[0].name}")
            else:
                # 多個檔案，需要用戶選擇
                self.stats['conflicts_resolved'] += 1
                print(f"\n⚠️  發現衝突：目錄 '{directory.relative_to(self.source_dir)}' 包含多個 manifest 檔案")
                print(f"完整路徑: {directory.absolute()}")
                print("請選擇要使用的檔案：")
                
                for i, manifest_file in enumerate(manifest_files, 1):
                    print(f"  {i}. {manifest_file.name}")
                
                while True:
                    try:
                        choice = input(f"請輸入選擇 (1-{len(manifest_files)}) 或 's' 跳過此目錄: ").strip().lower()
                        
                        if choice == 's':
                            self.logger.info(f"用戶選擇跳過目錄: {directory.relative_to(self.source_dir)}")
                            break
                        
                        choice_num = int(choice)
                        if 1 <= choice_num <= len(manifest_files):
                            selected_file = manifest_files[choice_num - 1]
                            resolved_manifests[directory] = selected_file
                            self.logger.info(f"用戶選擇檔案: {selected_file.name} (目錄: {directory.relative_to(self.source_dir)})")
                            break
                        else:
                            print(f"請輸入 1 到 {len(manifest_files)} 之間的數字，或 's' 跳過")
                            
                    except ValueError:
                        print(f"請輸入有效的數字 (1-{len(manifest_files)}) 或 's' 跳過")
                    except KeyboardInterrupt:
                        print("\n操作已取消")
                        self.logger.info("用戶取消操作")
                        return {}
        
        return resolved_manifests
    
    def _generate_backup_filename(self, directory: Path, original_name: str) -> str:
        """
        生成備份檔案名稱
        
        Args:
            directory: 目標目錄
            original_name: 原始檔案名稱
            
        Returns:
            str: 唯一的備份檔案名稱
        """
        base_name = f"backup_{original_name}"
        counter = 1
        
        while (directory / base_name).exists():
            name_without_ext = original_name.rsplit('.', 1)[0]
            extension = original_name.rsplit('.', 1)[1] if '.' in original_name else ''
            if extension:
                base_name = f"backup_{name_without_ext}_{counter}.{extension}"
            else:
                base_name = f"backup_{name_without_ext}_{counter}"
            counter += 1
        
        return base_name
    
    def _standardize_manifest_name(self, directory: Path, selected_manifest: Path) -> Tuple[bool, str]:
        """
        標準化 manifest 檔案名稱為 imsmanifest.xml
        
        Args:
            directory: 包含 manifest 檔案的目錄
            selected_manifest: 用戶選擇的 manifest 檔案
            
        Returns:
            Tuple[bool, str]: (是否成功, 錯誤訊息)
        """
        standard_name = "imsmanifest.xml"
        standard_path = directory / standard_name
        
        try:
            # 如果選擇的檔案已經是標準名稱，不需要處理
            if selected_manifest.name.lower() == standard_name.lower():
                self.logger.info(f"檔案 {selected_manifest.name} 已經是標準名稱，無需重命名")
                return True, ""
            
            # 檢查是否存在同名的標準檔案
            if standard_path.exists():
                # 生成備份檔案名稱
                backup_name = self._generate_backup_filename(directory, standard_name)
                backup_path = directory / backup_name
                
                # 備份現有的 imsmanifest.xml
                shutil.move(str(standard_path), str(backup_path))
                self.stats['manifests_backed_up'] += 1
                self.logger.info(f"已備份現有標準檔案: {standard_name} → {backup_name}")
            
            # 將選擇的檔案重命名為標準名稱
            shutil.move(str(selected_manifest), str(standard_path))
            self.stats['manifests_renamed'] += 1
            self.logger.info(f"已重命名 manifest 檔案: {selected_manifest.name} → {standard_name}")
            
            return True, ""
            
        except PermissionError as e:
            error_msg = f"無權限執行重命名操作: {e}"
            self.logger.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"重命名操作失敗: {e}"
            self.logger.error(error_msg)
            return False, error_msg
    
    def _generate_package_name(self, directory: Path, existing_names: Dict[str, int]) -> str:
        """
        生成壓縮包名稱
        
        Args:
            directory: 要打包的目錄
            existing_names: 已存在的名稱計數器
            
        Returns:
            str: 唯一的壓縮包名稱
        """
        # 獲取相對於 source_dir 的路徑組件
        relative_path = directory.relative_to(self.source_dir)
        path_parts = relative_path.parts
        
        # 建構基本名稱
        base_name = "_".join(path_parts) + "_scorm.zip"
        
        # 檢查是否需要加序號
        if base_name in existing_names:
            existing_names[base_name] += 1
            name_without_ext = base_name.rsplit('.', 1)[0]
            extension = base_name.rsplit('.', 1)[1]
            unique_name = f"{name_without_ext}_{existing_names[base_name]}.{extension}"
        else:
            existing_names[base_name] = 0
            unique_name = base_name
        
        return unique_name
    
    def _create_zip_package(self, source_directory: Path, package_name: str, 
                           selected_manifest: Path) -> bool:
        """
        建立 ZIP 壓縮包
        
        Args:
            source_directory: 要打包的來源目錄
            package_name: 壓縮包名稱
            selected_manifest: 選定的 manifest 檔案
            
        Returns:
            bool: 是否成功建立
        """
        # 首先執行 manifest 檔案標準化
        self.logger.info(f"開始標準化 manifest 檔案: {selected_manifest.name}")
        success, error_msg = self._standardize_manifest_name(source_directory, selected_manifest)
        
        if not success:
            self.stats['errors'] += 1
            print(f"❌ 標準化 manifest 檔案失敗: {error_msg}")
            self.logger.error(f"標準化 manifest 檔案失敗，中斷打包操作: {error_msg}")
            return False
        
        package_path = self.output_dir / package_name
        files_added = 0
        
        try:
            with zipfile.ZipFile(package_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # 遞迴添加目錄中的所有檔案
                for root, dirs, files in os.walk(source_directory):
                    root_path = Path(root)
                    
                    for file in files:
                        file_path = root_path / file
                        
                        try:
                            # 計算在壓縮包中的相對路徑
                            arcname = file_path.relative_to(source_directory)
                            zipf.write(file_path, arcname)
                            files_added += 1
                            
                        except PermissionError:
                            self.stats['skipped_files'] += 1
                            self.logger.warning(f"無權限存取檔案，已跳過: {file_path}")
                        except Exception as e:
                            self.stats['skipped_files'] += 1
                            self.logger.warning(f"無法添加檔案到壓縮包，已跳過 {file_path}: {e}")
            
            self.stats['packages_created'] += 1
            self.stats['files_packaged'] += files_added
            
            # 記錄打包結果
            self.package_results.append({
                'package_name': package_name,
                'source_directory': str(source_directory.relative_to(self.source_dir)),
                'selected_manifest': "imsmanifest.xml",  # 標準化後的名稱
                'original_manifest': selected_manifest.name,  # 原始名稱
                'files_count': files_added,
                'package_size': package_path.stat().st_size,
                'package_path': str(package_path.absolute())
            })
            
            self.logger.info(f"成功建立壓縮包: {package_name} (包含 {files_added} 個檔案)")
            return True
            
        except Exception as e:
            self.stats['errors'] += 1
            self.logger.error(f"建立壓縮包失敗 {package_name}: {e}")
            return False
    
    def _write_summary_report(self):
        """寫入詳細的打包報告"""
        report_file = self.output_dir / "packaging_report.log"
        
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                f.write("=== SCORM 打包詳細報告 ===\n")
                f.write(f"打包時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"來源目錄: {self.source_dir.absolute()}\n\n")
                
                f.write("=== 打包統計 ===\n")
                f.write(f"掃描目錄數: {self.stats['directories_scanned']}\n")
                f.write(f"找到 manifest 檔案數: {self.stats['manifests_found']}\n")
                f.write(f"建立壓縮包數: {self.stats['packages_created']}\n")
                f.write(f"打包檔案總數: {self.stats['files_packaged']}\n")
                f.write(f"解決衝突數: {self.stats['conflicts_resolved']}\n")
                f.write(f"重命名 manifest 檔案數: {self.stats['manifests_renamed']}\n")
                f.write(f"備份 manifest 檔案數: {self.stats['manifests_backed_up']}\n")
                f.write(f"跳過檔案數: {self.stats['skipped_files']}\n")
                f.write(f"錯誤數: {self.stats['errors']}\n\n")
                
                f.write("=== 打包詳情 ===\n")
                for result in self.package_results:
                    f.write(f"壓縮包: {result['package_name']}\n")
                    f.write(f"來源目錄: {result['source_directory']}\n")
                    f.write(f"原始 manifest: {result['original_manifest']}\n")
                    f.write(f"標準化後 manifest: {result['selected_manifest']}\n")
                    f.write(f"檔案數量: {result['files_count']}\n")
                    f.write(f"檔案大小: {result['package_size']:,} bytes\n")
                    f.write(f"完整路徑: {result['package_path']}\n")
                    f.write("-" * 50 + "\n")
                
            self.logger.info(f"詳細報告已寫入: {report_file}")
            
        except Exception as e:
            self.logger.error(f"寫入報告失敗: {e}")
    
    def package_scorm_contents(self) -> bool:
        """
        執行 SCORM 打包作業
        
        Returns:
            bool: 打包是否成功完成
        """
        try:
            print("🔍 開始掃描 manifest 檔案...")
            
            # 驗證來源目錄
            if not self._validate_source_directory():
                return False
            
            # 建立輸出目錄
            self._create_output_directory()
            
            # 掃描 manifest 檔案
            manifest_by_directory = self._scan_for_manifests()
            
            if not manifest_by_directory:
                print("❌ 未找到任何包含 'imsmanifest' 的 XML 檔案")
                self.logger.warning("未找到任何 manifest 檔案")
                return False
            
            print(f"✅ 找到 {len(manifest_by_directory)} 個包含 manifest 檔案的目錄")
            
            # 解決衝突
            print("\n🔧 解決檔案衝突...")
            resolved_manifests = self._resolve_conflicts(manifest_by_directory)
            
            if not resolved_manifests:
                print("❌ 沒有可用的 manifest 檔案進行打包")
                return False
            
            # 開始打包
            print(f"\n📦 開始打包 {len(resolved_manifests)} 個目錄...")
            existing_names = {}
            
            for directory, manifest_file in resolved_manifests.items():
                # 生成壓縮包名稱
                package_name = self._generate_package_name(directory, existing_names)
                
                print(f"正在打包: {directory.relative_to(self.source_dir)} → {package_name}")
                
                # 建立壓縮包（包含 manifest 標準化）
                success = self._create_zip_package(directory, package_name, manifest_file)
                
                if success:
                    print(f"✅ 完成: {package_name}")
                else:
                    print(f"❌ 失敗: {package_name}")
                    # 如果是因為 manifest 標準化失敗，中斷整個操作
                    if "標準化 manifest 檔案失敗" in str(self.logger.handlers[0].stream.getvalue() if hasattr(self.logger.handlers[0], 'stream') else ''):
                        print("⚠️  因為 manifest 檔案標準化失敗，中斷打包操作")
                        return False
            
            return True
            
        except KeyboardInterrupt:
            print("\n\n⚠️  操作已被用戶取消")
            self.logger.info("操作被用戶取消")
            return False
        except Exception as e:
            self.logger.error(f"打包過程發生錯誤: {e}")
            print(f"❌ 打包過程發生錯誤: {e}")
            return False
    
    def print_summary(self):
        """輸出打包摘要"""
        print("\n" + "="*50)
        print("📊 SCORM 打包作業完成")
        print("="*50)
        print(f"掃描目錄數: {self.stats['directories_scanned']}")
        print(f"找到 manifest 檔案數: {self.stats['manifests_found']}")
        print(f"建立壓縮包數: {self.stats['packages_created']}")
        print(f"打包檔案總數: {self.stats['files_packaged']}")
        if self.stats['conflicts_resolved'] > 0:
            print(f"解決衝突數: {self.stats['conflicts_resolved']}")
        if self.stats['manifests_renamed'] > 0:
            print(f"重命名 manifest 檔案數: {self.stats['manifests_renamed']}")
        if self.stats['manifests_backed_up'] > 0:
            print(f"備份 manifest 檔案數: {self.stats['manifests_backed_up']}")
        if self.stats['skipped_files'] > 0:
            print(f"跳過檔案數: {self.stats['skipped_files']}")
        if self.stats['errors'] > 0:
            print(f"錯誤數: {self.stats['errors']}")
        
        if self.package_results:
            print(f"\n📦 壓縮包已存放在: {self.output_dir.absolute()}")
            for result in self.package_results:
                size_mb = result['package_size'] / (1024 * 1024)
                print(f"  • {result['package_name']} ({size_mb:.2f} MB, {result['files_count']} 檔案)")
        
        # 寫入詳細報告
        self._write_summary_report()


def main():
    """主函數"""
    print("🚀 SCORM 自動打包工具")
    print("="*30)
    
    # 取得用戶輸入
    while True:
        source_folder = input("請輸入要掃描的資料夾名稱 (預設: merged_projects): ").strip()
        if not source_folder:
            source_folder = "merged_projects"
        
        source_path = Path(source_folder)
        if source_path.exists():
            break
        else:
            print(f"❌ 資料夾 '{source_folder}' 不存在，請重新輸入")
    
    # 建立打包器並執行
    packager = ScormPackager(source_folder)
    success = packager.package_scorm_contents()
    
    # 輸出摘要
    packager.print_summary()
    
    if success and packager.stats['packages_created'] > 0:
        print(f"\n🎉 打包完成！請查看 '{packager.output_dir}' 資料夾")
    elif success:
        print(f"\n⚠️  作業完成，但沒有建立任何壓縮包")
    else:
        print(f"\n❌ 打包過程中發生錯誤，請檢查日誌檔案")


if __name__ == "__main__":
    main()