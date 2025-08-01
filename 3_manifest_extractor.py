#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IMS Manifest 組織結構解析器 (增強版) - Part 1
解析 imsmanifest.xml 檔案中的 organizations 結構
支援 HTML 檔案過濾和影音檔案提取功能
新增路徑對應記錄功能
"""

import os
import json
import logging
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Any
from collections import defaultdict
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup


class ManifestParser:
    """Manifest 解析器類別"""
    
    def __init__(self, source_dir: str, output_dir: str = "04_manifest_structures"):
        """
        初始化解析器
        
        Args:
            source_dir: 來源目錄路徑
            output_dir: 輸出目錄路徑（預設: "04_manifest_structures"）
        """
        self.source_dir = Path(source_dir)
        self.output_dir = Path(output_dir)
        self.skip_non_html = False  # 是否略過非 HTML 檔案
        
        # 支援的影音檔案格式
        self.media_extensions = {
            # 影片格式
            'mpg', 'mpeg', 'mp4', 'mkv', 'avi', 'mov', 'wmv', 'flv', 'webm', 'm3u8',
            # 音訊格式
            'mp3', 'wav', 'aac', 'ogg', 'flac', 'wma', 'midi', 'mid'
        }
        
        self.stats = {
            'directories_scanned': 0,
            'manifests_found': 0,
            'manifests_parsed': 0,
            'json_files_created': 0,
            'conflicts_resolved': 0,
            'parse_errors': 0,
            'resource_missing': 0,
            'non_html_skipped': 0,
            'html_files_analyzed': 0,
            'html_files_missing': 0,
            'media_files_found': 0
        }
        
        # 記錄解析結果
        self.parse_results: List[Dict] = []
        self.error_logs: List[Dict] = []
        
        # 新增：路徑對應記錄
        self.path_mappings: List[Dict] = []
        
        # 設定日誌
        self._setup_logging()
    
    def _setup_logging(self):
        """設定日誌系統"""
        # 確保輸出資料夾存在
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # 確保 log 資料夾存在
        log_dir = Path("log")
        log_dir.mkdir(parents=True, exist_ok=True)

        # log 儲存在 log 資料夾
        log_filename = log_dir / f"manifest_parse_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )

        self.logger = logging.getLogger(__name__)
        self.logger.info("=== 開始 Manifest 解析作業 ===")
        self.logger.info(f"來源路徑: {self.source_dir}")
        self.logger.info(f"輸出資料夾: {self.output_dir}")
        self.logger.info(f"日誌檔案: {log_filename}")
    
    def _get_user_preferences(self):
        """獲取用戶偏好設定"""
        # 直接設定為不略過非 HTML 檔案
        self.skip_non_html = False
        print("⚙️  設定: 將處理所有檔案類型")
        print("-" * 20)
        
        self.logger.info(f"預設設定 - 略過非HTML檔案: {self.skip_non_html}")
    
    def _is_html_file(self, href: str) -> bool:
        """檢查檔案是否為 HTML 檔案"""
        if not href:
            return False
        
        # 移除查詢參數和錨點
        clean_href = href.split('?')[0].split('#')[0]
        return clean_href.lower().endswith(('.html', '.htm'))
    
    def _is_media_file(self, src: str) -> bool:
        """檢查檔案是否為影音檔案"""
        if not src:
            return False
        
        # 移除查詢參數和錨點
        clean_src = src.split('?')[0].split('#')[0]
        extension = Path(clean_src).suffix.lower().lstrip('.')
        return extension in self.media_extensions
    
    def _is_manifest_file(self, filename: str) -> bool:
        """檢查檔案是否為 manifest 檔案"""
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
                    self.logger.warning(f"無權限存取目錄: {current_dir}")
                except Exception as e:
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
                        print(f"請輸入選擇 (1-{len(manifest_files)}) 或 's' 跳過此目錄: ", end="", flush=True)
                        choice = input().strip().lower()
                        
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
    
    def _parse_xml_manifest(self, manifest_file: Path) -> Optional[ET.Element]:
        """
        解析 XML manifest 檔案
        
        Args:
            manifest_file: manifest 檔案路徑
            
        Returns:
            Optional[ET.Element]: 解析後的根元素，失敗時返回 None
        """
        try:
            tree = ET.parse(manifest_file)
            root = tree.getroot()
            self.logger.info(f"成功解析 XML: {manifest_file.name}")
            return root
        except ET.ParseError as e:
            self.stats['parse_errors'] += 1
            error_info = {
                'file': str(manifest_file.relative_to(self.source_dir)),
                'error_type': 'XML Parse Error',
                'error_message': str(e)
            }
            self.error_logs.append(error_info)
            self.logger.error(f"XML 解析錯誤 {manifest_file.name}: {e}")
            return None
        except Exception as e:
            self.stats['parse_errors'] += 1
            error_info = {
                'file': str(manifest_file.relative_to(self.source_dir)),
                'error_type': 'File Access Error',
                'error_message': str(e)
            }
            self.error_logs.append(error_info)
            self.logger.error(f"檔案存取錯誤 {manifest_file.name}: {e}")
            return None
        
    def _extract_media_from_html(self, html_file_path: Path, base_directory: Path) -> List[str]:
            """
            從 HTML 檔案中提取影音檔案路徑
            
            Args:
                html_file_path: HTML 檔案路徑
                base_directory: 基礎目錄（manifest 所在目錄）
                
            Returns:
                List[str]: 影音檔案路徑列表（相對於 manifest 目錄）
            """
            media_files = []
            
            try:
                with open(html_file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                try:
                    with open(html_file_path, 'r', encoding='latin-1') as f:
                        content = f.read()
                except Exception as e:
                    self.logger.warning(f"無法讀取 HTML 檔案 {html_file_path}: {e}")
                    return media_files
            except Exception as e:
                self.logger.warning(f"無法存取 HTML 檔案 {html_file_path}: {e}")
                return media_files
            
            try:
                soup = BeautifulSoup(content, 'html.parser')
                
                # 查找所有可能包含影音檔案的標籤
                media_tags = [
                    # video 標籤
                    ('video', ['src']),
                    # audio 標籤
                    ('audio', ['src']),
                    # source 標籤 (在 video/audio 內)
                    ('source', ['src']),
                    # embed 標籤
                    ('embed', ['src']),
                    # object 標籤
                    ('object', ['data']),
                    # iframe 標籤 (可能包含影音)
                    ('iframe', ['src']),
                    # a 標籤 (鏈接到影音檔案)
                    ('a', ['href'])
                ]
                
                html_dir = html_file_path.parent
                
                for tag_name, attributes in media_tags:
                    tags = soup.find_all(tag_name)
                    for tag in tags:
                        for attr in attributes:
                            src = tag.get(attr)
                            if src and self._is_media_file(src):
                                # 處理相對路徑
                                if not src.startswith(('http://', 'https://', '//')):
                                    # 相對路徑處理
                                    if src.startswith('/'):
                                        # 絕對路徑（相對於根目錄）
                                        media_path = base_directory / src.lstrip('/')
                                    else:
                                        # 相對路徑（相對於 HTML 檔案目錄）
                                        media_path = html_dir / src
                                    
                                    # 轉換為相對於 manifest 目錄的路徑
                                    try:
                                        relative_media_path = media_path.resolve().relative_to(base_directory.resolve())
                                        media_files.append(str(relative_media_path).replace('\\', '/'))
                                    except ValueError:
                                        # 檔案在 manifest 目錄外，記錄但不包含
                                        self.logger.warning(f"影音檔案在 manifest 目錄外: {src}")
                
                self.stats['media_files_found'] += len(media_files)
                if media_files:
                    self.logger.info(f"從 {html_file_path.name} 中找到 {len(media_files)} 個影音檔案")
                
            except Exception as e:
                self.logger.warning(f"解析 HTML 檔案時發生錯誤 {html_file_path}: {e}")
            
            return media_files
    
    def _is_url(self, href: str) -> bool:
        """檢查 href 是否為網路鏈接"""
        if not href:
            return False
        return href.startswith(('http://', 'https://', 'ftp://', '//'))

    def _extract_resources(self, root: ET.Element) -> Dict[str, str]:
        """
        提取 resources 中的 identifier 和 href 對應關係
        只取 resource 本身的 href，不限制檔案格式
        
        Args:
            root: XML 根元素
            
        Returns:
            Dict[str, str]: {identifier: href}
        """
        resources_map = {}
        
        # 尋找 resources 元素，考慮命名空間
        for resources in root.iter():
            if resources.tag.endswith('resources'):
                for resource in resources:
                    if resource.tag.endswith('resource'):
                        identifier = resource.get('identifier')
                        href = resource.get('href')  # 只取 resource 本身的 href
                        
                        if identifier and href:
                            resources_map[identifier] = href
        
        self.logger.info(f"提取到 {len(resources_map)} 個資源映射")
        return resources_map

    def _parse_item(self, item_elem: ET.Element, resources_map: Dict[str, str], 
                    base_directory: Path, item_path: str = "") -> Dict[str, Any]:
        """
        遞迴解析 item 元素，包含影音檔案提取
        
        Args:
            item_elem: item XML 元素
            resources_map: 資源映射字典
            base_directory: manifest 所在的基礎目錄
            item_path: 項目的完整路徑（用於錯誤記錄）
            
        Returns:
            Dict[str, Any]: 解析後的 item 資料
        """
        item_data = {}
        
        # 獲取 title
        title_elem = None
        for child in item_elem:
            if child.tag.endswith('title'):
                title_elem = child
                break
        
        if title_elem is not None:
            item_data['title'] = title_elem.text or ""
        else:
            item_data['title'] = item_elem.get('identifier', 'Untitled')
        
        # 建構當前項目的完整路徑
        current_path = f"{item_path} > {item_data['title']}" if item_path else item_data['title']
        
        # 獲取 identifierref 並查找對應的 href
        identifierref = item_elem.get('identifierref')
        if identifierref and identifierref in resources_map:
            href = resources_map[identifierref]
            
            # 檢查是否為網路鏈接
            if self._is_url(href):
                # 是網路鏈接，直接寫入，不報錯
                item_data['href'] = href
                self.logger.info(f"找到網路鏈接: {href} (項目路徑: {current_path})")
            else:
                # 是本地檔案路徑，檢查檔案是否存在
                item_data['href'] = href
                file_path = base_directory / href
                
                if file_path.exists():
                    # 如果是 HTML 檔案，分析其中的影音檔案
                    if self._is_html_file(href):
                        self.stats['html_files_analyzed'] += 1
                        media_list = self._extract_media_from_html(file_path, base_directory)
                        if media_list:
                            item_data['media_files'] = media_list
                else:
                    # 檔案不存在，寫入 JSON 並記錄 log
                    item_data['href'] = f"{href}  # 檔案不存在"
                    self.stats['html_files_missing'] += 1
                    self.logger.warning(f"檔案不存在: {file_path} (項目路徑: {current_path})")
                    error_info = {
                        'error_type': 'File Missing',
                        'file_path': str(href),
                        'item_title': item_data['title'],
                        'item_full_path': current_path,
                        'manifest_directory': str(base_directory.relative_to(self.source_dir))
                    }
                    self.error_logs.append(error_info)
                
        elif identifierref:
            # 找不到對應的資源，寫入 JSON 並記錄 log
            item_data['href'] = f"# 找不到資源: {identifierref}"
            self.stats['resource_missing'] += 1
            self.logger.warning(f"找不到資源 '{identifierref}' 對應的 href (項目路徑: {current_path})")
            error_info = {
                'error_type': 'Resource Missing',
                'identifierref': identifierref,
                'item_title': item_data['title'],
                'item_full_path': current_path,
                'manifest_directory': str(base_directory.relative_to(self.source_dir))
            }
            self.error_logs.append(error_info)
        
        # 遞迴處理子 items，傳遞完整路徑
        sub_items = []
        for child in item_elem:
            if child.tag.endswith('item'):
                sub_item = self._parse_item(child, resources_map, base_directory, current_path)
                sub_items.append(sub_item)
        
        if sub_items:
            item_data['items'] = sub_items
        
        return item_data
    
    def _parse_organizations(self, root: ET.Element, resources_map: Dict[str, str], 
                            base_directory: Path) -> List[Dict[str, Any]]:
        """
        解析 organizations 結構
        
        Args:
            root: XML 根元素
            resources_map: 資源映射字典
            base_directory: manifest 所在的基礎目錄
            
        Returns:
            List[Dict[str, Any]]: 解析後的組織結構列表
        """
        organizations_data = []
        
        # 尋找 organizations 元素
        for organizations in root.iter():
            if organizations.tag.endswith('organizations'):
                for organization in organizations:
                    if organization.tag.endswith('organization'):
                        org_data = {
                            'identifier': organization.get('identifier', ''),
                            'title': '',
                            'items': []
                        }
                        
                        # 獲取組織標題
                        for child in organization:
                            if child.tag.endswith('title'):
                                org_data['title'] = child.text or ""
                                break
                        
                        # 組織的路徑作為起始路徑
                        org_path = org_data['title'] or org_data['identifier']
                        
                        # 解析所有 item，傳遞組織路徑
                        for child in organization:
                            if child.tag.endswith('item'):
                                item_data = self._parse_item(child, resources_map, base_directory, org_path)
                                org_data['items'].append(item_data)
                        
                        organizations_data.append(org_data)
        
        return organizations_data
    
    def _generate_json_filename(self, directory: Path) -> str:
        """
        生成 JSON 檔案名稱
        
        Args:
            directory: 目錄路徑
            
        Returns:
            str: JSON 檔案名稱
        """
        # 獲取相對於 source_dir 的路徑組件
        relative_path = directory.relative_to(self.source_dir)
        path_parts = relative_path.parts
        
        # 建構檔名
        filename = "_".join(path_parts) + "_imsmanifest.json"
        return filename
    
    def _save_json_file(self, data: List[Dict[str, Any]], filename: str, 
                        source_directory: Path, manifest_file: Path) -> bool:
        """
        保存 JSON 檔案
        
        Args:
            data: 要保存的資料
            filename: 檔案名稱
            source_directory: 來源目錄
            manifest_file: manifest 檔案
            
        Returns:
            bool: 是否成功保存
        """
        try:
            json_path = self.output_dir / filename
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            self.stats['json_files_created'] += 1
            
            # 記錄結果
            self.parse_results.append({
                'json_filename': filename,
                'source_directory': str(source_directory.relative_to(self.source_dir)),
                'manifest_file': manifest_file.name,
                'organizations_count': len(data),
                'json_path': str(json_path.absolute())
            })
            
            # 新增：記錄路徑對應關係
            self.path_mappings.append({
                'json_filename': filename,
                'xml_relative_path': str(manifest_file.relative_to(self.source_dir)),
                'xml_absolute_path': str(manifest_file.absolute()),
                'source_directory_relative': str(source_directory.relative_to(self.source_dir)),
                'generated_time': datetime.now().isoformat()
            })
            
            self.logger.info(f"成功保存 JSON: {filename} (包含 {len(data)} 個組織)")
            return True
            
        except Exception as e:
            error_info = {
                'file': filename,
                'error_type': 'JSON Save Error',
                'error_message': str(e)
            }
            self.error_logs.append(error_info)
            self.logger.error(f"保存 JSON 檔案失敗 {filename}: {e}")
            return False
    
    def _save_path_mappings(self):
        """
        保存路徑對應記錄到 JSON 檔案
        """
        try:
            mapping_file = self.output_dir / "path_mappings.json"
            
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(self.path_mappings, f, ensure_ascii=False, indent=2)
            
            self.logger.info(f"路徑對應記錄已保存: {mapping_file}")
            print(f"📍 路徑對應記錄已保存: {mapping_file}")
            
        except Exception as e:
            self.logger.error(f"保存路徑對應記錄失敗: {e}")
            print(f"❌ 保存路徑對應記錄失敗: {e}")
        
    def _write_summary_report(self):
        """寫入詳細的解析報告"""
        # 確保 log 資料夾存在
        log_dir = Path("log")
        log_dir.mkdir(parents=True, exist_ok=True)
        
        report_file = log_dir / f"parsing_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                f.write("=== Manifest 解析詳細報告 ===\n")
                f.write(f"解析時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"來源目錄: {self.source_dir.absolute()}\n\n")
                
                f.write("=== 解析統計 ===\n")
                f.write(f"掃描目錄數: {self.stats['directories_scanned']}\n")
                f.write(f"找到 manifest 檔案數: {self.stats['manifests_found']}\n")
                f.write(f"成功解析檔案數: {self.stats['manifests_parsed']}\n")
                f.write(f"建立 JSON 檔案數: {self.stats['json_files_created']}\n")
                f.write(f"解決衝突數: {self.stats['conflicts_resolved']}\n")
                f.write(f"解析錯誤數: {self.stats['parse_errors']}\n")
                f.write(f"缺失資源數: {self.stats['resource_missing']}\n")
                f.write(f"略過非HTML檔案數: {self.stats['non_html_skipped']}\n")
                f.write(f"分析HTML檔案數: {self.stats['html_files_analyzed']}\n")
                f.write(f"缺失HTML檔案數: {self.stats['html_files_missing']}\n")
                f.write(f"找到影音檔案數: {self.stats['media_files_found']}\n\n")
                
                f.write("=== 成功解析的檔案 ===\n")
                for result in self.parse_results:
                    f.write(f"JSON 檔案: {result['json_filename']}\n")
                    f.write(f"來源目錄: {result['source_directory']}\n")
                    f.write(f"Manifest 檔案: {result['manifest_file']}\n")
                    f.write(f"組織數量: {result['organizations_count']}\n")
                    f.write(f"完整路徑: {result['json_path']}\n")
                    f.write("-" * 50 + "\n")
                
                if self.error_logs:
                    f.write("\n=== 錯誤記錄 ===\n")
                    for error in self.error_logs:
                        f.write(f"錯誤類型: {error['error_type']}\n")
                        f.write(f"Manifest 目錄: {error.get('manifest_directory', 'N/A')}\n")
                        f.write(f"項目完整路徑: {error.get('item_full_path', 'N/A')}\n")
                        if 'file_path' in error:
                            f.write(f"缺失檔案路徑: {error['file_path']}\n")
                        if 'identifierref' in error:
                            f.write(f"缺失資源: {error['identifierref']}\n")
                        if 'item_title' in error:
                            f.write(f"項目標題: {error['item_title']}\n")
                        if 'error_message' in error:
                            f.write(f"錯誤訊息: {error['error_message']}\n")
                        f.write("-" * 30 + "\n")
                
            self.logger.info(f"詳細報告已寫入: {report_file}")
            
        except Exception as e:
            self.logger.error(f"寫入報告失敗: {e}")
    
    def parse_manifests(self) -> bool:
        """
        執行 manifest 解析作業
        
        Returns:
            bool: 解析是否成功完成
        """
        try:
            print("🔍 開始掃描 manifest 檔案...")
            
            # 獲取用戶偏好設定
            self._get_user_preferences()
            print()
            
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
                print("❌ 沒有可用的 manifest 檔案進行解析")
                return False
            
            # 開始解析
            print(f"\n📋 開始解析 {len(resolved_manifests)} 個 manifest 檔案...")
            
            for directory, manifest_file in resolved_manifests.items():
                print(f"正在解析: {directory.relative_to(self.source_dir)}/{manifest_file.name}")
                
                # 解析 XML
                root = self._parse_xml_manifest(manifest_file)
                if root is None:
                    print(f"❌ 解析失敗: {manifest_file.name}")
                    continue
                
                # 提取資源映射
                resources_map = self._extract_resources(root)
                
                # 解析組織結構
                organizations_data = self._parse_organizations(root, resources_map, directory)
                
                if not organizations_data:
                    self.logger.warning(f"檔案 {manifest_file.name} 中未找到 organizations")
                    print(f"⚠️  未找到組織結構: {manifest_file.name}")
                    continue
                
                self.stats['manifests_parsed'] += 1
                
                # 生成 JSON 檔名並保存
                json_filename = self._generate_json_filename(directory)
                success = self._save_json_file(organizations_data, json_filename, directory, manifest_file)
                
                if success:
                    print(f"✅ 完成: {json_filename}")
                else:
                    print(f"❌ 保存失敗: {json_filename}")
            
            # 保存路徑對應記錄
            if self.path_mappings:
                self._save_path_mappings()
            
            return True
            
        except KeyboardInterrupt:
            print("\n\n⚠️  操作已被用戶取消")
            self.logger.info("操作被用戶取消")
            return False
        except Exception as e:
            self.logger.error(f"解析過程發生錯誤: {e}")
            print(f"❌ 解析過程發生錯誤: {e}")
            return False
    
    def print_summary(self):
        """輸出解析摘要"""
        print("\n" + "="*50)
        print("📊 Manifest 解析作業完成")
        print("="*50)
        print(f"掃描目錄數: {self.stats['directories_scanned']}")
        print(f"找到 manifest 檔案數: {self.stats['manifests_found']}")
        print(f"成功解析檔案數: {self.stats['manifests_parsed']}")
        print(f"建立 JSON 檔案數: {self.stats['json_files_created']}")
        if self.stats['conflicts_resolved'] > 0:
            print(f"解決衝突數: {self.stats['conflicts_resolved']}")
        if self.stats['parse_errors'] > 0:
            print(f"解析錯誤數: {self.stats['parse_errors']}")
        if self.stats['resource_missing'] > 0:
            print(f"缺失資源數: {self.stats['resource_missing']}")
        if self.stats['non_html_skipped'] > 0:
            print(f"略過非HTML檔案數: {self.stats['non_html_skipped']}")
        if self.stats['html_files_analyzed'] > 0:
            print(f"分析HTML檔案數: {self.stats['html_files_analyzed']}")
        if self.stats['html_files_missing'] > 0:
            print(f"缺失HTML檔案數: {self.stats['html_files_missing']}")
        if self.stats['media_files_found'] > 0:
            print(f"找到影音檔案數: {self.stats['media_files_found']}")
        
        if self.parse_results:
            print(f"\n📄 JSON 檔案已存放在: {self.output_dir.absolute()}")
            for result in self.parse_results:
                print(f"  • {result['json_filename']} ({result['organizations_count']} 個組織)")
        
        # 新增：顯示路徑對應記錄資訊
        if self.path_mappings:
            print(f"\n📍 路徑對應記錄: path_mappings.json (包含 {len(self.path_mappings)} 筆記錄)")
        
        # 寫入詳細報告
        self._write_summary_report()


def main():
    """主函數"""
    print("🚀 IMS Manifest 組織結構解析器 (增強版)")
    print("="*40)
    print("支援功能：")
    print("• 組織結構解析")
    print("• HTML 檔案過濾")  
    print("• 影音檔案自動提取")
    print("• 路徑對應記錄生成")
    print("• 完整的錯誤處理和日誌")
    print()
    
    # 檢查依賴
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        print("❌ 缺少必要的依賴套件：beautifulsoup4")
        print("請執行：pip install beautifulsoup4")
        return False
    
    # 取得用戶輸入
    while True:
        print("請輸入要掃描的資料夾名稱 (輸入 '0' 使用預設: 02_merged_projects): ", end="", flush=True)
        source_folder = input().strip()
        if not source_folder:
            print("⚠️ 請輸入有效值，或輸入 '0' 使用預設值")
            continue
        if source_folder == '0':
            source_folder = "02_merged_projects"
        
        source_path = Path(source_folder)
        if source_path.exists():
            break
        else:
            print(f"❌ 資料夾 '{source_folder}' 不存在，請重新輸入")
    
    # 建立解析器並執行
    parser = ManifestParser(source_folder)
    success = parser.parse_manifests()
    
    # 輸出摘要
    parser.print_summary()
    
    if success and parser.stats['json_files_created'] > 0:
        print(f"\n🎉 解析完成！請查看 '{parser.output_dir}' 資料夾")
        if parser.stats['media_files_found'] > 0:
            print(f"💾 成功提取 {parser.stats['media_files_found']} 個影音檔案路徑")
        if parser.path_mappings:
            print(f"📍 生成 {len(parser.path_mappings)} 筆路徑對應記錄")
    elif success:
        print(f"\n⚠️  作業完成，但沒有建立任何 JSON 檔案")
    else:
        print(f"\n❌ 解析過程中發生錯誤，請檢查日誌檔案")
    
    return success


if __name__ == "__main__":
    main()