# Auto Tronc - 自動創課系統

![Auto Tronc GUI](https://img.shields.io/badge/GUI-Tkinter-blue) ![Python](https://img.shields.io/badge/Python-3.8+-green) ![License](https://img.shields.io/badge/License-MIT-yellow)

## 🎯 專案概述

Auto Tronc 是一個功能完整的自動化課程創建系統，專為 TronClass 平台設計。本專案提供了直觀的圖形用戶界面 (GUI)，整合了完整的課程製作工作流程，從資料整理到最終課程上線，實現一鍵式自動化處理。

## ✨ 核心功能

### 🖥️ 圖形化用戶界面
- **現代化 GUI 設計**：基於 Tkinter 開發的專業圖形界面
- **工作流程可視化**：7 個主要步驟的按鈕式操作
- **實時日誌監控**：執行過程的即時反饋和狀態追蹤
- **Excel 文件編輯器**：內建配置文件編輯功能
- **互動式終端**：支援腳本互動式執行

### 📋 完整工作流程

#### 1. 資料夾合併 (`1_folder_merger.py`)
- 將 Google Drive 下載的分散項目文件合併到統一目錄
- 支援深度遞迴合併和智能重複文件處理
- 自動統計處理結果

#### 2. SCORM 打包 (`2_scorm_packager.py`)
- 自動識別包含 `imsmanifest.xml` 的項目
- 標準化 SCORM 包生成
- 支援批量處理和用戶選擇性打包

#### 3. 課程結構提取 (`3_manifest_extractor.py`)
- 從 XML manifest 文件中提取課程結構資訊
- 生成標準化的 Excel 格式課程架構
- 支援 HTML 內容過濾和組織

#### 4. 資源庫映射 (`4_cloud_mapping.py`)
- 根據映射規則生成資源庫路徑
- 自動補充缺失的資源連結
- Excel 格式的資源管理

#### 5. 執行文件生成 (`5_0_to_be_executed_excel_generator.sh`)
- Shell 腳本自動化執行
- 生成待處理的批次文件

#### 6. 系統待辦清單 (`6_system_todolist_maker.py`)
- 分析 Excel 文件生成系統化待辦事項
- 支援多工作表處理和用戶選擇

#### 7. 自動創課執行 (`7_start_tronc.py`)
- **課程創建**：自動建立新課程
- **章節管理**：批量創建課程章節  
- **單元組織**：自動生成學習單元
- **學習活動**：支援多種活動類型（影片、音訊、文件、連結）
- **資源上傳**：自動化材料上傳和管理

## 🚀 快速開始

### 系統需求
- Python 3.8 或更高版本
- 支援的作業系統：Windows、macOS、Linux

### 安裝步驟

1. **克隆專案**
   ```bash
   git clone <repository-url>
   cd 250708_Auto_Tronc
   ```

2. **安裝依賴**
   ```bash
   pip install -r requirements.txt
   ```

3. **配置系統**
   編輯 `config.py` 文件，設定：
   - TronClass 平台 URL
   - 登入憑證
   - 預設課程和章節 ID

4. **啟動 GUI**
   ```bash
   python run_gui.py
   ```

### 快速配置指南

在 `config.py` 中設定以下參數：

```python
BASE_URL = 'https://your-tronclass-domain.com'
USERNAME = 'your-email@domain.com'
PASSWORD = 'your-password'
COURSE_ID = 16401  # 目標課程 ID
MODULE_ID = 28739  # 預設章節 ID
```

## 📁 專案結構

```
250708_Auto_Tronc/
├── auto_tronc_gui.py          # 主要 GUI 應用程序
├── run_gui.py                 # GUI 啟動器
├── config.py                  # 系統配置文件
├── requirements.txt           # Python 依賴清單
├── final_terminal.py          # 終端模擬器
├── tronc_login.py            # 登入管理模組
├── 
├── 工作流程腳本/
│   ├── 1_folder_merger.py     # 資料夾合併工具
│   ├── 2_scorm_packager.py    # SCORM 打包器
│   ├── 3_manifest_extractor.py # 結構提取器
│   ├── 4_cloud_mapping.py     # 資源映射工具
│   ├── 5_0_to_be_executed_excel_generator.sh # 執行文件生成
│   ├── 6_system_todolist_maker.py # 待辦清單生成器
│   └── 7_start_tronc.py       # 自動創課主程序
├── 
├── 創建模組/
│   ├── create_01_course.py    # 課程創建
│   ├── create_02_module.py    # 章節創建  
│   ├── create_03_syllabus.py  # 大綱創建
│   ├── create_04_activity.py  # 活動創建
│   └── create_05_material.py  # 資源上傳
├── 
├── 輔助工具/
│   ├── excel_analyzer.py      # Excel 分析器
│   ├── detailed_excel_checker.py # 詳細檢查器
│   └── sub_*.py              # 子模組工具
└── 
└── 資料目錄/
    ├── projects/             # 原始項目資料
    ├── merged_projects/      # 合併後的項目
    ├── scorm_packages/       # SCORM 包輸出
    ├── manifest_structures/  # 提取的結構文件
    ├── to_be_executed/      # 待執行文件
    └── log/                 # 系統日誌
```

## 🎮 GUI 操作指南

### 主界面功能

1. **工作流程面板** (左側)
   - 7 個步驟按鈕，可獨立執行
   - 每個步驟都有詳細說明
   - 彩色編碼便於識別

2. **文件編輯面板** (右上)
   - 快速打開 Excel 配置文件
   - 內建配置編輯器
   - 支援文件瀏覽和選擇

3. **執行日誌面板** (右中)
   - 實時顯示執行過程
   - 支援日誌保存和清空
   - 彩色編碼的狀態訊息

4. **狀態監控面板** (右下)
   - 當前執行狀態顯示
   - 進度條指示
   - 快捷鍵支援 (F5 刷新, Ctrl+Q 退出)

### 互動式執行

某些腳本支援互動式執行，提供：
- 終端模擬環境
- 實時輸入輸出
- 預設值處理 (輸入 '0' 使用預設值)
- 執行狀態監控

## 🔧 支援的活動類型

系統支援以下學習活動類型：

| 活動類型 | 描述 | API 映射 |
|---------|------|----------|
| 線上連結 | 外部網站連結 | `web_link` |
| 影音連結 | 影音平台連結 | `web_link` |
| 影音教材_影音連結 | 線上影音教材 | `online_video` |
| 參考檔案_圖片 | 圖片資源 | `material` |
| 參考檔案_PDF | PDF 文件 | `material` |
| 影音教材_影片 | 本地影片文件 | `video` |
| 影音教材_音訊 | 音訊文件 | `audio` |

## 📊 Excel 文件格式

### 課程結構 Excel 格式
```
| 課程名稱 | 章節名稱 | 單元名稱 | 活動名稱 | 活動類型 | 資源路徑 | 描述 |
|---------|---------|---------|---------|---------|---------|------|
```

### 資源映射 Excel 格式  
```
| 原始路徑 | 目標路徑 | 資源類型 | 狀態 | 備註 |
|---------|---------|---------|------|------|
```

## 🔐 安全性考量

- 所有敏感資訊存放在 `config.py` 中
- 支援 Cookie 自動管理和更新
- API 請求間隔控制防止被阻擋
- 錯誤日誌記錄和追蹤

## 🐛 疑難排解

### 常見問題

1. **GUI 無法啟動**
   - 檢查 Python 版本 (需要 3.8+)
   - 確認所有依賴已安裝: `pip install -r requirements.txt`

2. **腳本執行失敗**
   - 檢查 `config.py` 配置是否正確
   - 確認網路連接和 TronClass 平台可訪問性

3. **Excel 文件讀取錯誤**
   - 確認文件格式符合規範
   - 檢查文件路徑和權限

4. **SCORM 打包失敗**
   - 確認 `merged_projects` 目錄存在
   - 檢查 `imsmanifest.xml` 文件完整性

### 日誌檢查

系統會在 `log/` 目錄下生成詳細日誌：
- `packaging_report_*.log` - 打包過程日誌
- `scorm_package_*.log` - SCORM 生成日誌
- `*_error_*.json` - 錯誤詳細記錄

## 🚀 進階使用

### 自定義工作流程
可以通過修改 `auto_tronc_gui.py` 中的 `workflow_steps` 來自定義工作流程。

### 批量處理
使用命令列直接執行各個腳本進行批量處理：
```bash
python 1_folder_merger.py
python 2_scorm_packager.py
# ... 依序執行
```

### API 擴展
可以在 `create_*.py` 模組中添加新的創建功能，系統會自動整合。

## 📝 開發資訊

- **開發語言**: Python 3.8+
- **GUI 框架**: Tkinter
- **HTTP 請求**: requests
- **數據處理**: pandas, openpyxl
- **網頁解析**: BeautifulSoup4
- **瀏覽器自動化**: Selenium

## 📄 授權條款

本專案採用 MIT 授權條款。詳見 LICENSE 文件。

## 🤝 貢獻指南

歡迎提交 Issue 和 Pull Request 來改進本專案。

---

**Auto Tronc** - 讓課程創建變得簡單高效！ 🎓✨