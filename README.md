Note  2025.08.01
- 注意：基於目前 1.77 的 TC 權限架構，用戶在創建課程完成後，需要在課程中上傳該課程的 Scorm 包，非創建人的其他用戶才有辦法訪問屬於 Scorm 包的連結。

--
# Auto Tronc - 自動創課系統

![Auto Tronc GUI](https://img.shields.io/badge/GUI-Tkinter-blue) ![Python](https://img.shields.io/badge/Python-3.8+-green) ![License](https://img.shields.io/badge/License-MIT-yellow)

## 🎯 專案概述

Auto Tronc 是一個功能完整的自動化教學系統，專為 TronClass 平台設計。本專案提供了直觀的圖形用戶界面 (GUI)，整合了兩大核心模塊：**課程創建**與**題庫創建**，從資料整理到最終內容上線，實現一鍵式自動化處理。

## 🏗️ 系統架構

本系統分為兩大核心模塊：

### 🎓 模塊一：課程創建系統
專注於 SCORM 課程包的處理與自動上傳，完整支援從原始資料到線上課程的全流程自動化。

### 📝 模塊二：題庫創建系統  
專注於測驗題目的批量處理與上傳，支援多種題型格式轉換與資源管理。

## ✨ 核心功能

### 🖥️ 圖形化用戶界面
- **現代化 GUI 設計**：基於 Tkinter 開發的專業圖形界面
- **工作流程可視化**：7 個主要步驟的按鈕式操作
- **實時日誌監控**：執行過程的即時反饋和狀態追蹤
- **Excel 文件編輯器**：內建配置文件編輯功能
- **互動式終端**：支援腳本互動式執行

### 🎓 模塊一：課程創建工作流程

#### 1. ZIP 檔案解壓縮 (`1_folder_merger.py`)
- 按檔案名稱排序，逐一解壓縮 ZIP 檔案到統一目錄
- 支援覆蓋模式解壓縮和完整統計
- 適用於 Google Drive 壓縮包處理

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
- **學習活動**：支援多種活動類型（影片、文件、連結）**註：音訊功能尚未驗證支持**
- **資源上傳**：自動化材料上傳和管理

### 📝 模塊二：題庫創建工作流程

#### 1. 題目資料提取 (`exam_1_extractor.py`)
- 從各種格式文件中提取題目內容
- 支援 Word 文檔和其他常見格式
- 自動分析題目結構和類型

#### 2. 檔案重命名整理 (`exam_2_rename.py`)
- 標準化題庫檔案命名規則
- 批量整理題目資源文件
- 建立有序的檔案結構

#### 3. XML 題目生成 (`exam_3_xmlpouring.py`)
- 將題目內容轉換為標準 XML 格式
- 支援多種題型（選擇題、填空題、問答題等）
- 自動驗證 XML 格式正確性

#### 4. Word 檔案切分 (`exam_4_wordsplitter.py`)
- 智能切分大型 Word 題庫文件
- 按章節或題目數量自動分割
- 保持原有格式和圖片資源

#### 5. Word 資源批量上傳 (`exam_5_wordpouring.py`)
- 批量上傳 Word 格式題庫到平台
- 自動創建題庫資源庫條目
- 即時更新上傳狀態和資源 ID

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

3. **配置環境變數**
   ```bash
   # 複製環境變數範本
   cp .env.example .env
   
   # 編輯 .env 文件，填入您的實際資訊
   nano .env
   ```

4. **配置 .env 文件**
   在 `.env` 文件中設定以下參數：
   ```bash
   # 登入資訊
   USERNAME=your-email@domain.com
   PASSWORD=your-password
   
   # 平台設定
   BASE_URL=https://your-tronclass-domain.com
   ```

5. **啟動 GUI**
   ```bash
   python auto_tronc_gui.py
   ```

### 🔐 環境變數配置

本系統使用 `.env` 文件來安全地存放敏感資訊，避免將帳號密碼上傳到 GitHub。

#### 設定步驟：

1. **複製範本文件**
   ```bash
   cp .env.example .env
   ```

2. **編輯 .env 文件**
   ```bash
   # TronClass 系統配置
   USERNAME=your-email@domain.com     # 您的登入帳號
   PASSWORD=your-password             # 您的登入密碼
   BASE_URL=https://your-domain.com   # TronClass 平台網址
   ```

3. **重要安全提醒**
   - ✅ `.env` 文件已被 `.gitignore` 忽略，不會上傳到版本控制
   - ✅ 請勿將 `.env` 文件分享給他人
   - ✅ 每個使用者都需要建立自己的 `.env` 文件

#### GUI 配置編輯器

系統提供內建的配置編輯器，可透過 GUI 界面修改所有設定：

1. 啟動 GUI：`python auto_tronc_gui.py`
2. 點擊「⚙️ 編輯配置」按鈕
3. 在視覺化界面中修改設定
4. 點擊「💾 保存配置」自動更新 `.env` 和 `config.py`

## 📁 專案結構

```
250708_Auto_Tronc/
├── auto_tronc_gui.py          # 主要 GUI 應用程序
├── config.py                  # 系統配置文件
├── requirements.txt           # Python 依賴清單
├── final_terminal.py          # 終端模擬器
├── tronc_login.py            # 登入管理模組
├── excel_analyzer.py          # Excel 分析器
├── 
├── 🎓 模塊一：課程創建系統/
│   ├── 1_folder_merger.py     # ZIP 檔案解壓縮
│   ├── 2_scorm_packager.py    # SCORM 打包器
│   ├── 3_manifest_extractor.py # 課程結構提取器
│   ├── 4_cloud_mapping.py     # 資源庫映射工具
│   ├── 5_0_to_be_executed_excel_generator.sh # 執行文件生成
│   ├── 6_system_todolist_maker.py # 系統待辦清單生成器
│   ├── 7_start_tronc.py       # 自動創課主程序
│   └── create_*.py           # 創建功能模組
│       ├── create_01_course.py    # 課程創建
│       ├── create_02_module.py    # 章節創建
│       ├── create_03_syllabus.py  # 大綱創建
│       ├── create_04_activity.py  # 活動創建
│       └── create_05_material.py  # 資源上傳
├── 
├── 📝 模塊二：題庫創建系統/
│   ├── exam_1_extractor.py    # 題目資料提取
│   ├── exam_2_rename.py       # 檔案重命名整理
│   ├── exam_3_xmlpouring.py   # XML 題目生成
│   ├── exam_4_wordsplitter.py # Word 檔案切分
│   ├── exam_5_wordpouring.py  # Word 資源批量上傳
│   └── sub_*.py              # 題庫處理子模組
│       ├── sub_library_creator.py     # 資源庫創建
│       ├── sub_word_splitter.py       # Word 分割工具
│       ├── sub_todolist_resource.py   # 題庫待辦清單
│       └── sub_excel_*.py            # Excel 處理工具
├── 
└── 📂 資料目錄/
    ├── 01_ori_zipfiles/      # 原始 ZIP 檔案
    ├── 02_merged_projects/   # 合併後的課程項目
    ├── 03_scorm_packages/    # SCORM 包輸出
    ├── 04_manifest_structures/ # 提取的課程結構文件
    ├── 05_to_be_executed/    # 課程創建待執行文件
    ├── 06_todolist/          # 課程創建待辦清單文件
    ├── exam_01_02_merged_projects/ # 題庫原始資料
    ├── exam_02_xml_todolist/ # XML 題庫待辦清單
    ├── exam_03_wordsplitter/ # Word 切分結果
    ├── exam_04_docx_todolist/ # Word 題庫待辦清單
    └── log/                 # 系統運行日誌
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
| 影音教材_音訊 | 音訊文件 | `audio` | **尚未驗證支持** |

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

- **環境變數保護**：敏感資訊（用戶名、密碼、URL）存放在 `.env` 文件中
- **版本控制安全**：`.env` 文件已被 `.gitignore` 忽略，不會上傳到 GitHub
- **Cookie 管理**：支援 Session Cookie 自動管理和更新
- **請求控制**：API 請求間隔控制防止被阻擋
- **錯誤追蹤**：完整的錯誤日誌記錄和追蹤系統

## 🐛 疑難排解

### 常見問題

1. **GUI 無法啟動**
   - 檢查 Python 版本 (需要 3.8+)
   - 確認所有依賴已安裝: `pip install -r requirements.txt`
   - 確認已安裝 `python-dotenv`: `pip install python-dotenv`

2. **環境變數配置問題**
   - 確認 `.env` 文件存在並包含正確資訊
   - 檢查 `.env` 文件格式 (無空格、無引號)
   - 使用 GUI 配置編輯器檢查設定

3. **腳本執行失敗**
   - 檢查 `.env` 文件中的登入資訊是否正確
   - 確認網路連接和 TronClass 平台可訪問性
   - 檢查 `BASE_URL` 是否正確

4. **登入認證失敗**
   - 確認用戶名和密碼正確
   - 檢查平台網址是否可訪問
   - 嘗試手動登入平台驗證帳號狀態

5. **Excel 文件讀取錯誤**
   - 確認文件格式符合規範
   - 檢查文件路徑和權限

6. **SCORM 打包失敗**
   - 確認 `2_merged_projects` 目錄存在
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

#### 課程創建模塊批量處理：
```bash
# 課程創建完整流程
python 1_folder_merger.py      # ZIP 檔案解壓縮
python 2_scorm_packager.py     # SCORM 打包
python 3_manifest_extractor.py # 課程結構提取
python 4_cloud_mapping.py      # 資源庫映射
bash 5_0_to_be_executed_excel_generator.sh # 執行文件生成
python 6_system_todolist_maker.py # 待辦清單生成
python 7_start_tronc.py        # 自動創課執行
```

#### 題庫創建模塊批量處理：
```bash
# 題庫創建完整流程
python exam_1_extractor.py     # 題目資料提取
python exam_2_rename.py        # 檔案重命名整理
python exam_3_xmlpouring.py    # XML 題目生成
python exam_4_wordsplitter.py  # Word 檔案切分
python exam_5_wordpouring.py   # Word 資源批量上傳
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