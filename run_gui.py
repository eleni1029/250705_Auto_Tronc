#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auto Tronc GUI 啟動器
"""

import sys
import os

def main():
    """主函數"""
    print("🚀 啟動 Auto Tronc GUI...")
    
    # 檢查必要文件
    required_files = [
        "auto_tronc_gui.py",
        "config.py",
        "1_folder_merger.py",
        "2_scorm_packager.py",
        "3_manifest_extractor.py", 
        "4_cloud_mapping.py",
        "6_system_todolist_maker.py",
        "7_start_tronc.py"
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print("❌ 缺少必要文件:")
        for file in missing_files:
            print(f"   - {file}")
        print("\n請確保所有文件都在當前目錄中。")
        return False
    
    print("✅ 所有必要文件檢查完成")
    
    # 啟動GUI
    try:
        from auto_tronc_gui import main as gui_main
        gui_main()
    except ImportError as e:
        print(f"❌ 導入GUI模組失敗: {e}")
        return False
    except Exception as e:
        print(f"❌ 啟動GUI失敗: {e}")
        return False
    
    return True

if __name__ == "__main__":
    main()