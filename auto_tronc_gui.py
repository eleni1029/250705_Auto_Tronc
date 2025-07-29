#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auto Tronc GUI - 自動創課系統圖形界面
整合所有工作流程步驟的專業GUI應用程序
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess
import threading
import os
import sys
from datetime import datetime
from final_terminal import FinalTerminal

class AutoTroncGUI:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_variables()
        self.create_widgets()
        self.setup_layout()
        
    def setup_window(self):
        """設置主窗口"""
        self.root.title("Auto Tronc - 自動創課系統")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # 設置窗口圖標（如果有的話）
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
            
    def setup_variables(self):
        """設置變量"""
        self.current_step = tk.StringVar(value="待開始")
        self.progress_var = tk.DoubleVar()
        self.log_text = tk.StringVar(value="歡迎使用 Auto Tronc 自動創課系統")
        
        # 工作流程步驟 (使用更柔和的顏色)
        self.workflow_steps = [
            {
                "id": "1",
                "name": "資料夾合併",
                "description": "將GoogleDrive下載解壓文件從projects合併到merged_projects",
                "script": "1_folder_merger.py",
                "button_text": "1. 合併專案資料夾",
                "color": "#5cb85c"  # 柔和綠色
            },
            {
                "id": "2", 
                "name": "SCORM打包",
                "description": "從merged_projects中找到XML封裝為SCORM packages",
                "script": "2_scorm_packager.py",
                "button_text": "2. 建立SCORM包",
                "color": "#5bc0de"  # 柔和藍色
            },
            {
                "id": "3",
                "name": "結構提取",
                "description": "從merged_projects中找到XML抽取待處理結構文件",
                "script": "3_manifest_extractor.py", 
                "button_text": "3. 提取課程結構",
                "color": "#f0ad4e"  # 柔和橙色
            },
            {
                "id": "4",
                "name": "資源庫映射",
                "description": "根據mapping文件生成待補充資源庫路徑的Excel",
                "script": "4_cloud_mapping.py",
                "button_text": "4. 生成資源映射",
                "color": "#d9534f"  # 柔和紅色
            },
            {
                "id": "5",
                "name": "執行文件生成",
                "description": "生成待執行文件",
                "script": "5_0_to_be_executed_excel_generator.sh",
                "button_text": "5. 生成執行文件",
                "color": "#777777"  # 柔和灰色
            },
            {
                "id": "6",
                "name": "系統列表製作",
                "description": "生成系統批次執行文件",
                "script": "6_system_todolist_maker.py",
                "button_text": "6. 製作待辦清單",
                "color": "#5bc0de"  # 重複使用藍色調
            },
            {
                "id": "7",
                "name": "開始自動創課",
                "description": "執行自動創課流程",
                "script": "7_start_tronc.py", 
                "button_text": "7. 開始自動創課",
                "color": "#d9534f"  # 重複使用紅色調
            }
        ]
        
    def create_widgets(self):
        """創建所有界面組件"""
        # 創建主框架
        self.create_main_frame()
        self.create_header_frame()
        self.create_workflow_frame()
        self.create_excel_frame()
        self.create_log_frame()
        self.create_status_frame()
        
    def create_main_frame(self):
        """創建主框架"""
        # 使用PanedWindow創建可調整大小的分割面板
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左側面板（工作流程）
        self.left_frame = ttk.Frame(self.main_paned, width=450)
        self.left_frame.pack_propagate(False)  # 防止內容影響框架大小
        self.main_paned.add(self.left_frame, weight=1)
        
        # 右側面板（日志和控制）
        self.right_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.right_frame, weight=2)
        
        # 設置初始分割位置
        self.root.after(100, lambda: self.main_paned.sashpos(0, 450))
        
    def create_header_frame(self):
        """創建標題框架"""
        header_frame = ttk.Frame(self.left_frame)
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 主標題
        title_label = ttk.Label(header_frame, text="Auto Tronc 自動創課系統", 
                               font=("Arial", 18, "bold"))
        title_label.pack(anchor=tk.W)
        
        # 副標題
        subtitle_label = ttk.Label(header_frame, text="一鍵式自動化課程創建工作流程", 
                                  font=("Arial", 12))
        subtitle_label.pack(anchor=tk.W)
        
        # 分隔線
        separator = ttk.Separator(header_frame, orient=tk.HORIZONTAL)
        separator.pack(fill=tk.X, pady=(10, 0))
        
    def create_workflow_frame(self):
        """創建工作流程框架"""
        workflow_frame = ttk.LabelFrame(self.left_frame, text="工作流程", padding=10)
        workflow_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 創建滾動區域
        canvas = tk.Canvas(workflow_frame)
        scrollbar = ttk.Scrollbar(workflow_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 為每個步驟創建按鈕和說明
        for i, step in enumerate(self.workflow_steps):
            step_frame = ttk.Frame(scrollable_frame)
            step_frame.pack(fill=tk.X, padx=5, pady=5)
            
            # 步驟按鈕
            btn = tk.Button(step_frame, 
                           text=step["button_text"],
                           bg=step["color"],
                           fg="#2c3e50",  # 深灰色文字，提高對比度
                           font=("Arial", 12, "bold"),
                           relief=tk.RAISED,
                           bd=2,
                           command=lambda s=step: self.run_step(s))
            btn.pack(fill=tk.X, pady=(0, 5))
            
            # 步驟說明
            desc_label = ttk.Label(step_frame, text=step["description"], 
                                  font=("Arial", 10), foreground="gray")
            desc_label.pack(anchor=tk.W)
            
            # 分隔線（最後一個步驟除外）
            if i < len(self.workflow_steps) - 1:
                sep = ttk.Separator(step_frame, orient=tk.HORIZONTAL)
                sep.pack(fill=tk.X, pady=10)
        
        # 工作流程說明
        info_frame = ttk.Frame(scrollable_frame)
        info_frame.pack(fill=tk.X, padx=5, pady=20)
        
        info_label = ttk.Label(info_frame, 
                             text="💡 請依需求單獨執行各個步驟，每個步驟完成後請檢查結果", 
                             font=("Arial", 12), 
                             foreground="darkblue",
                             wraplength=400)
        info_label.pack()
        
        # 打包滾動區域
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 綁定滾輪事件
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind('<Enter>', _bind_to_mousewheel)
        canvas.bind('<Leave>', _unbind_from_mousewheel)
        
    def create_excel_frame(self):
        """創建Excel編輯框架"""
        excel_frame = ttk.LabelFrame(self.right_frame, text="文件編輯", padding=10)
        excel_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Excel文件路徑顯示
        path_frame = ttk.Frame(excel_frame)
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(path_frame, text="當前文件:").pack(anchor=tk.W)
        self.excel_path_var = tk.StringVar(value="4_資源庫路徑_補充.xlsx")
        path_label = ttk.Label(path_frame, textvariable=self.excel_path_var, 
                              font=("Arial", 10), foreground="blue")
        path_label.pack(anchor=tk.W)
        
        # 按鈕框架
        btn_frame = ttk.Frame(excel_frame)
        btn_frame.pack(fill=tk.X)
        
        # 打開Excel按鈕
        open_btn = ttk.Button(btn_frame, text="📊 打開Excel編輯", 
                             command=self.open_excel)
        open_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 瀏覽其他文件按鈕
        browse_btn = ttk.Button(btn_frame, text="📁 瀏覽...", 
                               command=self.browse_excel)
        browse_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 配置編輯按鈕
        config_btn = ttk.Button(btn_frame, text="⚙️ 編輯配置", 
                               command=self.open_config_editor)
        config_btn.pack(side=tk.LEFT)
        
        # 說明
        info_label = ttk.Label(excel_frame, 
                              text="可編輯Excel文件和系統配置。\n編輯完成後請保存文件。",
                              font=("Arial", 10),
                              foreground="gray")
        info_label.pack(pady=(10, 0))
        
    def create_log_frame(self):
        """創建日志框架"""
        log_frame = ttk.LabelFrame(self.right_frame, text="執行日誌", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 日志文本區域
        self.log_text_widget = scrolledtext.ScrolledText(log_frame, 
                                                        height=15, 
                                                        font=("Arial", 10))
        self.log_text_widget.pack(fill=tk.BOTH, expand=True)
        
        # 日志控制按鈕
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.pack(fill=tk.X, pady=(10, 0))
        
        clear_btn = ttk.Button(log_control_frame, text="清空日誌", 
                              command=self.clear_log)
        clear_btn.pack(side=tk.LEFT)
        
        save_btn = ttk.Button(log_control_frame, text="保存日誌", 
                             command=self.save_log)
        save_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 添加初始歡迎信息
        self.log_message("歡迎使用 Auto Tronc 自動創課系統！")
        self.log_message("請選擇要執行的工作流程步驟。")
        
    def create_status_frame(self):
        """創建狀態框架"""
        status_frame = ttk.Frame(self.right_frame)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 當前狀態
        ttk.Label(status_frame, text="當前狀態:").pack(anchor=tk.W)
        status_label = ttk.Label(status_frame, textvariable=self.current_step, 
                               font=("Arial", 12, "bold"), foreground="green")
        status_label.pack(anchor=tk.W)
        
        # 進度條
        ttk.Label(status_frame, text="執行進度:").pack(anchor=tk.W, pady=(10, 0))
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, 
                                          maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))
        
    def log_message(self, message):
        """記錄日誌信息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text_widget.insert(tk.END, log_entry)
        self.log_text_widget.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        """清空日誌"""
        self.log_text_widget.delete(1.0, tk.END)
        self.log_message("日誌已清空")
        
    def save_log(self):
        """保存日誌到文件"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"auto_tronc_log_{timestamp}.txt"
            
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.log_text_widget.get(1.0, tk.END))
            
            self.log_message(f"日誌已保存到: {filename}")
            messagebox.showinfo("保存成功", f"日誌已保存到: {filename}")
        except Exception as e:
            error_msg = f"保存日誌失敗: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("保存失敗", error_msg)
            
    def open_excel(self):
        """打開Excel文件進行編輯"""
        excel_path = self.excel_path_var.get()
        
        if not os.path.exists(excel_path):
            messagebox.showerror("文件不存在", f"文件不存在: {excel_path}")
            return
            
        try:
            if sys.platform.startswith('darwin'):  # macOS
                subprocess.run(['open', excel_path])
            elif sys.platform.startswith('win'):   # Windows
                os.startfile(excel_path)
            else:  # Linux
                subprocess.run(['xdg-open', excel_path])
                
            self.log_message(f"已打開Excel文件: {excel_path}")
        except Exception as e:
            error_msg = f"打開Excel失敗: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("打開失敗", error_msg)
            
    def browse_excel(self):
        """瀏覽並選擇Excel文件"""
        filename = filedialog.askopenfilename(
            title="選擇Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ],
            initialdir=os.getcwd()
        )
        
        if filename:
            self.excel_path_var.set(filename)
            self.log_message(f"已選擇Excel文件: {os.path.basename(filename)}")
            
    def open_config_editor(self):
        """打開配置編輯器"""
        ConfigEditor(self.root, self)
    
    def _needs_interaction(self, script_path):
        """檢查腳本是否需要用戶交互"""
        interactive_scripts = [
            'test_interactive.py',
            '2_scorm_packager.py',
            '3_manifest_extractor.py', 
            '6_system_todolist_maker.py',
            '7_start_tronc.py'
        ]
        return os.path.basename(script_path) in interactive_scripts
    
    def _execute_interactive_script(self, step):
        """執行需要交互的腳本"""
        self.log_message(f"🤖 執行交互式腳本: {step['name']}")
        
        try:
            # 使用新的 PTY 終端執行所有需要交互的腳本
            self._run_script_with_pty_terminal(step)
                
        except Exception as e:
            self.log_message(f"❌ 交互式執行錯誤: {str(e)}")
            self.current_step.set("交互式執行錯誤")
    
    def _run_scorm_packager_interactive(self, step):
        """交互式執行SCORM打包器"""
        self.progress_var.set(20)
        
        # 使用自定義對話框代替 simpledialog
        source_folder = self._get_folder_input("SCORM打包設定", "請輸入要掃描的資料夾名稱:", "merged_projects")
        
        if source_folder is None:
            self.current_step.set("取消: 用戶取消操作")
            return
            
        if not source_folder.strip():
            source_folder = "merged_projects"
            
        if not os.path.exists(source_folder):
            messagebox.showerror("錯誤", f"資料夾 '{source_folder}' 不存在", parent=self.root)
            self.current_step.set("錯誤: 資料夾不存在")
            return
            
        self.log_message(f"📂 來源資料夾: {source_folder}")
        self.progress_var.set(40)
        
        # 直接使用命令行執行，提供預設輸入
        inputs = [source_folder]
        self._execute_script_with_input(step, inputs)
    
    def _run_manifest_extractor_interactive(self, step):
        """交互式執行結構提取器"""
        self.progress_var.set(20)
        
        # 使用自定義對話框獲取來源資料夾
        source_folder = self._get_folder_input("結構提取設定", "請輸入要掃描的資料夾名稱:", "merged_projects")
        
        if source_folder is None:
            self.current_step.set("取消: 用戶取消操作")
            return
            
        if not source_folder.strip():
            source_folder = "merged_projects"
            
        if not os.path.exists(source_folder):
            messagebox.showerror("錯誤", f"資料夾 '{source_folder}' 不存在", parent=self.root)
            self.current_step.set("錯誤: 資料夾不存在")
            return
            
        # 詢問是否略過非HTML文件
        skip_non_html = messagebox.askyesno(
            "結構提取設定",
            "是否略過非HTML檔案？",
            parent=self.root
        )
        
        self.log_message(f"📂 來源資料夾: {source_folder}")
        self.log_message(f"⚙️ 略過非HTML: {'是' if skip_non_html else '否'}")
        self.progress_var.set(60)
        
        # 模擬命令行執行
        inputs = [source_folder, 'y' if skip_non_html else 'n']
        self._execute_script_with_input(step, inputs)
    
    def _run_todolist_maker_interactive(self, step):
        """交互式執行待辦清單製作器"""
        self.progress_var.set(20)
        
        # 確保主窗口可見並更新
        self.root.update()
        
        # 檢查analyzed文件
        pattern = os.path.join('to_be_executed', '*analyzed*.xlsx')
        import glob
        files = glob.glob(pattern)
        
        if not files:
            messagebox.showerror("錯誤", "在to_be_executed目錄中找不到analyzed檔案", parent=self.root)
            self.current_step.set("錯誤: 找不到analyzed檔案")
            return
            
        # 如果有多個檔案，讓用戶選擇
        if len(files) > 1:
            file_names = [os.path.basename(f) for f in files]
            choice = self._show_choice_dialog(
                "選擇分析檔案",
                "請選擇要處理的analyzed檔案:",
                file_names
            )
            if choice is None:
                self.current_step.set("取消: 用戶取消選擇")
                return
            file_index = choice
        else:
            file_index = 0
            
        self.log_message(f"📄 選擇檔案: {os.path.basename(files[file_index])}")
        self.progress_var.set(60)
        
        # 模擬選擇（預設選擇第一個檔案和全部sheet）
        inputs = ['', 'all']  # Enter選擇第一個檔案，all選擇全部sheet
        self._execute_script_with_input(step, inputs)
    
    def _run_start_tronc_interactive(self, step):
        """交互式執行自動創課"""
        self.progress_var.set(20)
        
        # 確保主窗口可見並更新
        self.root.update()
        
        # 檢查extracted文件
        pattern = os.path.join('to_be_executed', '*extracted*.xlsx')
        import glob
        files = glob.glob(pattern)
        
        if not files:
            messagebox.showerror("錯誤", "在to_be_executed目錄中找不到extracted檔案", parent=self.root)
            self.current_step.set("錯誤: 找不到extracted檔案")
            return
            
        self.progress_var.set(40)
        
        # 詢問操作類型
        operations = [
            "建立文件內所有元素",
            "建立所有課程", 
            "建立所有章節",
            "建立所有單元",
            "建立所有學習活動",
            "建立特定類型學習活動",
            "建立所有資源"
        ]
        
        operation_choice = self._show_choice_dialog(
            "選擇操作",
            "請選擇要進行的操作:",
            operations
        )
        
        if operation_choice is None:
            self.current_step.set("取消: 用戶取消選擇")
            return
            
        self.log_message(f"🎯 選擇操作: {operations[operation_choice]}")
        
        # 詢問確認
        confirm = messagebox.askyesno(
            "確認執行",
            f"確認執行 '{operations[operation_choice]}' 嗎？",
            parent=self.root
        )
        
        if not confirm:
            self.current_step.set("取消: 用戶取消確認")
            return
            
        self.progress_var.set(60)
        
        # 模擬輸入
        inputs = ['', str(operation_choice + 1), 'y', '']  # 第一個檔案，選擇操作，確認，預設錯誤處理
        self._execute_script_with_input(step, inputs)
    
    def _run_script_with_pty_terminal(self, step):
        """使用 PTY 終端執行腳本"""
        try:
            # 創建 PTY 終端窗口
            script_path = step.get('file_path', step.get('script', ''))
            terminal = FinalTerminal(self.root, step['name'], script_path)
            self.root.wait_window(terminal.window)
            
            # 更新狀態
            if terminal.execution_successful:
                self.log_message(f"✅ {step['name']} 執行完成")
                self.current_step.set(f"{step['name']} - 完成")
                self.progress_var.set(100)
            else:
                self.log_message(f"❌ {step['name']} 執行失敗或被取消")
                self.current_step.set(f"{step['name']} - 失敗/取消")
                
        except Exception as e:
            self.log_message(f"❌ 終端執行錯誤: {str(e)}")
            self.current_step.set("終端執行錯誤")
    
    def _get_folder_input(self, title, message, default_value):
        """獲取資料夾輸入"""
        try:
            # 確保主窗口處於前台
            self.root.lift()
            self.root.focus_force()
            self.root.update()
            
            # 創建自定義輸入對話框
            dialog = InputDialog(self.root, title, message, default_value)
            self.root.wait_window(dialog.window)
            return dialog.result
        except Exception as e:
            self.log_message(f"❌ 輸入對話框錯誤: {str(e)}")
            # 如果自定義對話框失敗，使用預設值
            return default_value
    
    def _show_choice_dialog(self, title, message, choices):
        """顯示選擇對話框"""
        try:
            # 確保主窗口更新
            self.root.update_idletasks()
            dialog = ChoiceDialog(self.root, title, message, choices)
            self.root.wait_window(dialog.window)
            return dialog.result
        except Exception as e:
            self.log_message(f"❌ 對話框顯示錯誤: {str(e)}")
            return None
    
    def _execute_script_with_input(self, step, inputs):
        """執行腳本並提供輸入"""
        try:
            script_path = step['script']
            cmd = [sys.executable, script_path]
            
            self.log_message(f"執行命令: {' '.join(cmd)}")
            self.log_message(f"📥 提供輸入: {inputs}")
            
            # 執行腳本並提供輸入
            process = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                cwd=os.getcwd()
            )
            
            # 提供輸入
            input_text = '\n'.join(inputs) + '\n'
            stdout, _ = process.communicate(input=input_text)
            
            self.progress_var.set(100)
            
            # 顯示輸出
            for line in stdout.split('\n'):
                if line.strip():
                    self.log_message(line.strip())
            
            if process.returncode == 0:
                self.log_message(f"✅ {step['name']} 執行成功")
                self.current_step.set(f"完成: {step['name']}")
            else:
                self.log_message(f"❌ {step['name']} 執行失敗 (返回碼: {process.returncode})")
                self.current_step.set(f"失敗: {step['name']}")
                
        except Exception as e:
            self.log_message(f"❌ 執行錯誤: {str(e)}")
            self.current_step.set("執行錯誤")
            
    def run_step(self, step):
        """執行單個工作流程步驟"""
        self.log_message(f"開始執行: {step['name']}")
        self.current_step.set(f"執行中: {step['name']}")
        
        # 在後台線程中執行腳本
        thread = threading.Thread(target=self._execute_script, args=(step,))
        thread.daemon = True
        thread.start()
        
    def _execute_script(self, step):
        """在後台執行腳本"""
        try:
            script_path = step['script']
            
            if not os.path.exists(script_path):
                self.log_message(f"❌ 腳本文件不存在: {script_path}")
                self.current_step.set("錯誤: 腳本不存在")
                return
                
            self.progress_var.set(20)
            
            # 根據文件擴展名選擇執行方式
            if script_path.endswith('.py'):
                # 檢查是否需要交互式處理
                if self._needs_interaction(script_path):
                    self._execute_interactive_script(step)
                    return
                cmd = [sys.executable, script_path]
            elif script_path.endswith('.sh'):
                # sh腳本使用交互式終端執行以獲取完整日誌
                self._execute_interactive_script(step)
                return
            else:
                cmd = [script_path]
                
            self.log_message(f"執行命令: {' '.join(cmd)}")
            self.progress_var.set(40)
            
            # 執行腳本並捕獲輸出
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                cwd=os.getcwd(),
                bufsize=1,
                universal_newlines=True
            )
            
            self.progress_var.set(60)
            
            # 實時顯示輸出
            for line in iter(process.stdout.readline, ''):
                if line:
                    self.log_message(line.strip())
                    
            process.wait()
            self.progress_var.set(100)
            
            if process.returncode == 0:
                self.log_message(f"✅ {step['name']} 執行成功")
                self.current_step.set(f"完成: {step['name']}")
            else:
                self.log_message(f"❌ {step['name']} 執行失敗 (返回碼: {process.returncode})")
                self.current_step.set(f"失敗: {step['name']}")
                
        except Exception as e:
            error_msg = f"❌ 執行 {step['name']} 時發生錯誤: {str(e)}"
            self.log_message(error_msg)
            self.current_step.set("執行錯誤")
            
        finally:
            # 重置進度條
            self.root.after(2000, lambda: self.progress_var.set(0))
            
            
    def setup_layout(self):
        """設置佈局"""
        # 設置窗口關閉事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 設置鍵盤快捷鍵
        self.root.bind('<Control-q>', lambda e: self.on_closing())
        self.root.bind('<F5>', lambda e: self.refresh_status())
        
    def refresh_status(self):
        """刷新狀態信息"""
        self.log_message("🔄 刷新狀態...")
        
        # 檢查文件狀態
        excel_path = self.excel_path_var.get()
        if os.path.exists(excel_path):
            mtime = os.path.getmtime(excel_path)
            mtime_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            self.log_message(f"📊 Excel文件狀態: {excel_path} (修改時間: {mtime_str})")
        else:
            self.log_message(f"❌ Excel文件不存在: {excel_path}")
            
        # 檢查腳本文件
        missing_scripts = []
        for step in self.workflow_steps:
            if not os.path.exists(step['script']):
                missing_scripts.append(step['script'])
                
        if missing_scripts:
            self.log_message(f"⚠️ 缺少腳本文件: {', '.join(missing_scripts)}")
        else:
            self.log_message("✅ 所有腳本文件都存在")
            
    def on_closing(self):
        """窗口關閉事件處理"""
        if messagebox.askokcancel("退出", "確定要退出 Auto Tronc GUI 嗎？"):
            self.root.destroy()

class ConfigEditor:
    """配置編輯器窗口"""
    
    def __init__(self, parent, main_app):
        self.parent = parent
        self.main_app = main_app
        self.window = tk.Toplevel(parent)
        self.config_data = {}
        self.setup_window()
        self.load_config()
        self.create_widgets()
        
    def setup_window(self):
        """設置窗口"""
        self.window.title("系統配置編輯器")
        self.window.geometry("600x500")
        self.window.resizable(True, True)
        
        # 設置窗口在父窗口中央
        self.window.transient(self.parent)
        self.window.grab_set()
        
    def load_config(self):
        """載入配置文件"""
        try:
            config_path = "config.py"
            if not os.path.exists(config_path):
                messagebox.showerror("錯誤", "config.py 文件不存在！")
                self.window.destroy()
                return
                
            # 讀取配置文件內容
            with open(config_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            # 簡單解析配置值
            self.config_data = {}
            for line in content.split('\n'):
                line = line.strip()
                if '=' in line and not line.startswith('#') and not line.startswith('def'):
                    try:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip()
                        
                        # 處理字符串值（去除引號）
                        if value.startswith(("'", '"')) and value.endswith(("'", '"')):
                            value = value[1:-1]
                        
                        self.config_data[key] = value
                    except:
                        continue
                        
        except Exception as e:
            messagebox.showerror("錯誤", f"載入配置失敗: {str(e)}")
            self.window.destroy()
            
    def create_widgets(self):
        """創建界面組件"""
        # 創建主框架
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 標題
        title_label = ttk.Label(main_frame, text="Auto Tronc 系統配置", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 創建滾動區域
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 配置項
        self.config_vars = {}
        
        # 基本設置
        basic_frame = ttk.LabelFrame(scrollable_frame, text="基本設置", padding=10)
        basic_frame.pack(fill=tk.X, pady=5)
        
        # 基礎URL
        self.create_config_field(basic_frame, "BASE_URL", "基礎網址", 
                                "TronClass 系統的基礎網址")
        
        # 登入設置
        login_frame = ttk.LabelFrame(scrollable_frame, text="登入設置", padding=10)
        login_frame.pack(fill=tk.X, pady=5)
        
        self.create_config_field(login_frame, "USERNAME", "用戶名", 
                                "登入用戶名（email格式）")
        self.create_config_field(login_frame, "PASSWORD", "密碼", 
                                "登入密碼", show="*")
        
        # 預設值設置
        defaults_frame = ttk.LabelFrame(scrollable_frame, text="預設值設置", padding=10)
        defaults_frame.pack(fill=tk.X, pady=5)
        
        self.create_config_field(defaults_frame, "COURSE_ID", "預設課程ID", 
                                "預設的課程ID")
        self.create_config_field(defaults_frame, "MODULE_ID", "預設章節ID", 
                                "預設的章節ID")
        self.create_config_field(defaults_frame, "SLEEP_SECONDS", "請求間隔", 
                                "每次API請求間隔秒數")
        
        # Cookie設置
        cookie_frame = ttk.LabelFrame(scrollable_frame, text="Cookie設置", padding=10)
        cookie_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(cookie_frame, text="Session Cookie:").pack(anchor=tk.W)
        ttk.Label(cookie_frame, text="（通常由自動登入獲取，無需手動編輯）", 
                 font=("Arial", 12), foreground="gray").pack(anchor=tk.W)
        
        self.cookie_text = tk.Text(cookie_frame, height=3, wrap=tk.WORD)
        self.cookie_text.pack(fill=tk.X, pady=5)
        
        if "COOKIE" in self.config_data:
            self.cookie_text.insert(1.0, self.config_data["COOKIE"])
        
        # 按鈕框架
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(fill=tk.X, pady=20)
        
        # 保存按鈕
        save_btn = ttk.Button(button_frame, text="💾 保存配置", 
                             command=self.save_config)
        save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 重置按鈕
        reset_btn = ttk.Button(button_frame, text="🔄 重置", 
                              command=self.reset_config)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 取消按鈕
        cancel_btn = ttk.Button(button_frame, text="❌ 取消", 
                               command=self.window.destroy)
        cancel_btn.pack(side=tk.RIGHT)
        
        # 測試連接按鈕
        test_btn = ttk.Button(button_frame, text="🔗 測試連接", 
                             command=self.test_connection)
        test_btn.pack(side=tk.RIGHT, padx=(0, 10))
        
        # 打包滾動區域
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def create_config_field(self, parent, key, label, description, show=None):
        """創建配置字段"""
        field_frame = ttk.Frame(parent)
        field_frame.pack(fill=tk.X, pady=5)
        
        # 標籤
        label_frame = ttk.Frame(field_frame)
        label_frame.pack(fill=tk.X)
        
        ttk.Label(label_frame, text=f"{label}:", 
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        ttk.Label(label_frame, text=description, 
                 font=("Arial", 12), foreground="gray").pack(side=tk.RIGHT)
        
        # 輸入框
        var = tk.StringVar()
        if key in self.config_data:
            var.set(self.config_data[key])
            
        entry = ttk.Entry(field_frame, textvariable=var, show=show)
        entry.pack(fill=tk.X, pady=(2, 0))
        
        self.config_vars[key] = var
        
    def save_config(self):
        """保存配置"""
        try:
            # 更新配置數據
            for key, var in self.config_vars.items():
                self.config_data[key] = var.get()
                
            # 更新Cookie
            self.config_data["COOKIE"] = self.cookie_text.get(1.0, tk.END).strip()
            
            # 重新生成配置文件
            self.write_config_file()
            
            messagebox.showinfo("成功", "配置已保存！")
            self.main_app.log_message("✅ 系統配置已更新")
            self.window.destroy()
            
        except Exception as e:
            error_msg = f"保存配置失敗: {str(e)}"
            messagebox.showerror("錯誤", error_msg)
            self.main_app.log_message(f"❌ {error_msg}")
            
    def write_config_file(self):
        """寫入配置文件"""
        # 更新 .env 文件
        env_content = f'''# TronClass 系統配置
# 請根據您的實際情況填寫以下資訊

# 登入資訊
USERNAME={self.config_data.get("USERNAME", "")}
PASSWORD={self.config_data.get("PASSWORD", "")}

# 平台設定
BASE_URL={self.config_data.get("BASE_URL", "https://staging.tronclass.com")}'''
        
        with open(".env", 'w', encoding='utf-8') as f:
            f.write(env_content)
        
        # 更新 config.py 文件
        config_content = f'''# config.py
# 全局設定檔案，請在此填寫共用參數

import os
from dotenv import load_dotenv

# 載入環境變數
load_dotenv()

# 從環境變數獲取敏感資訊
USERNAME = os.getenv('USERNAME', '')  # 從 .env 文件讀取
PASSWORD = os.getenv('PASSWORD', '')  # 從 .env 文件讀取  
BASE_URL = os.getenv('BASE_URL', 'https://staging.tronclass.com')  # 從 .env 文件讀取

# 其他系統設定
COOKIE = '{self.config_data.get("COOKIE", "")}'  # 自動登入獲取
SLEEP_SECONDS = {self.config_data.get("SLEEP_SECONDS", "0.1")}  # 每次請求間隔，避免被擋
LOGIN_URL = f'{{BASE_URL}}/login'  # 登入網址
COURSE_ID = {self.config_data.get("COURSE_ID", "16401")}  # 預設的課程 ID
MODULE_ID = {self.config_data.get("MODULE_ID", "28739")}  # 預設的章節 ID

# 活動類型映射
ACTIVITY_TYPE_MAPPING = {{
    '線上連結': 'web_link',
    '影音連結': 'web_link',
    '影音教材_影音連結': 'online_video',
    '參考檔案_圖片': 'material',
    '參考檔案_PDF': 'material',
    '影音教材_影片': 'video',
    '影音教材_音訊': 'audio'
}}

# 支援的活動類型
SUPPORTED_ACTIVITY_TYPES = [
    '線上連結',
    '影音連結',
    '影音教材_影音連結',
    '參考檔案_圖片',
    '參考檔案_PDF',
    '影音教材_影片',
    '影音教材_音訊'
]

def get_api_urls():
    """獲取 API URLs"""
    return {{
        'COURSE_CREATE_URL': f'{{BASE_URL}}/api/course',
        'MODULE_CREATE_URL': f'{{BASE_URL}}/api/course/{{COURSE_ID}}/module',
        'SYLLABUS_CREATE_URL': f'{{BASE_URL}}/api/syllabus',
        'ACTIVITY_CREATE_URL': f'{{BASE_URL}}/api/courses/{{COURSE_ID}}/activities',
        'MATERIAL_UPLOAD_URL': f'{{BASE_URL}}/api/uploads',
        'MATERIAL_CREATE_URL': f'{{BASE_URL}}/api/material'
    }}'''
        
        with open("config.py", 'w', encoding='utf-8') as f:
            f.write(config_content)
            
    def reset_config(self):
        """重置配置"""
        if messagebox.askyesno("確認", "確定要重置所有配置嗎？"):
            self.load_config()
            # 重新設置所有字段
            for key, var in self.config_vars.items():
                if key in self.config_data:
                    var.set(self.config_data[key])
                else:
                    var.set("")
                    
            if "COOKIE" in self.config_data:
                self.cookie_text.delete(1.0, tk.END)
                self.cookie_text.insert(1.0, self.config_data["COOKIE"])
                
    def test_connection(self):
        """測試連接"""
        try:
            base_url = self.config_vars["BASE_URL"].get()
            if not base_url:
                messagebox.showerror("錯誤", "請先設置基礎網址")
                return
                
            import requests
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            
            # 測試基本連接
            response = requests.get(f"{base_url}/login", timeout=10, verify=False)
            
            if response.status_code == 200:
                messagebox.showinfo("成功", f"連接測試成功！\n網址: {base_url}")
                self.main_app.log_message(f"✅ 連接測試成功: {base_url}")
            else:
                messagebox.showwarning("警告", f"連接測試返回狀態碼: {response.status_code}")
                self.main_app.log_message(f"⚠️ 連接測試警告: {response.status_code}")
                
        except Exception as e:
            error_msg = f"連接測試失敗: {str(e)}"
            messagebox.showerror("錯誤", error_msg)
            self.main_app.log_message(f"❌ {error_msg}")

class InteractiveTerminalOld:
    """互動式終端組件 - 基於subprocess"""
    
    def __init__(self, parent, script_name, script_path):
        self.parent = parent
        self.script_name = script_name
        self.script_path = script_path
        self.execution_successful = False
        self.process = None
        self.update_job = None
        self.waiting_for_input = False
        self.output_queue = None
        self.read_thread = None
        
        # 創建窗口
        self.window = tk.Toplevel(parent)
        self.window.title(f"互動式終端 - {script_name}")
        self.window.geometry("900x700")
        self.window.resizable(True, True)
        
        # 設置窗口屬性
        self.window.transient(parent)
        self.window.grab_set()
        
        self.setup_ui()
        self.start_script()
        
        # 居中顯示
        self.center_window()
        
    def setup_ui(self):
        """設置用戶界面"""
        # 主框架
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 標題
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(title_frame, text=f"🖥️ {self.script_name}", 
                 font=("Arial", 14, "bold")).pack(side=tk.LEFT)
        
        # 狀態標籤
        self.status_var = tk.StringVar(value="正在啟動...")
        self.status_label = ttk.Label(title_frame, textvariable=self.status_var, 
                                     font=("Arial", 12), foreground="blue")
        self.status_label.pack(side=tk.RIGHT)
        
        # 終端輸出區域
        terminal_frame = ttk.Frame(main_frame)
        terminal_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 創建文本區域和滾動條
        self.output_text = tk.Text(terminal_frame, 
                                  font=("Arial", 12),
                                  bg="#1e1e1e", 
                                  fg="#ffffff",
                                  insertbackground="#ffffff",
                                  selectbackground="#404040",
                                  wrap=tk.WORD,
                                  state=tk.DISABLED)
        
        scrollbar = ttk.Scrollbar(terminal_frame, orient=tk.VERTICAL, command=self.output_text.yview)
        self.output_text.configure(yscrollcommand=scrollbar.set)
        
        self.output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 輸入區域
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(input_frame, text="輸入:", font=("Arial", 12)).pack(side=tk.LEFT, padx=(0, 5))
        
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_var, 
                                    font=("Arial", 12))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.input_entry.bind('<Return>', self.send_input)
        # 默認啟用輸入框並設置焦點
        self.input_entry.focus_set()
        
        self.send_button = ttk.Button(input_frame, text="發送", command=self.send_input)
        self.send_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        # 添加提示標籤
        hint_frame = ttk.Frame(main_frame)
        hint_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(hint_frame, text="💡 提示：輸入 '0' 使用預設值，輸入框始終可用", 
                 font=("Arial", 10), foreground="gray").pack(side=tk.LEFT)
        
        # 按鈕區域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="終止腳本", command=self.terminate_script).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="清除輸出", command=self.clear_output).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="關閉", command=self.close_terminal).pack(side=tk.RIGHT)
        
        # 設置文本標籤
        self.output_text.tag_configure("error", foreground="#ff6b6b")
        self.output_text.tag_configure("warning", foreground="#feca57")
        self.output_text.tag_configure("success", foreground="#48ca61")
        self.output_text.tag_configure("input", foreground="#74b9ff")
        self.output_text.tag_configure("prompt", foreground="#fdcb6e")
        
    def center_window(self):
        """居中顯示窗口"""
        self.window.update_idletasks()
        x = (self.parent.winfo_rootx() + self.parent.winfo_width() // 2) - (self.window.winfo_width() // 2)
        y = (self.parent.winfo_rooty() + self.parent.winfo_height() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        
    def start_script(self):
        """啟動腳本 - 簡化但有效的版本"""
        try:
            import os
            import subprocess
            import threading
            import queue
            
            # 設置環境變量
            env = os.environ.copy()
            env['PYTHONUNBUFFERED'] = '1'
            env['PYTHONIOENCODING'] = 'utf-8'
            
            # 根據腳本類型選擇執行命令
            if self.script_path.endswith('.py'):
                cmd = [sys.executable, '-u', self.script_path]
            elif self.script_path.endswith('.sh'):
                cmd = ['bash', self.script_path]
            else:
                cmd = [self.script_path]
            
            # 啟動進程
            self.process = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,  # 合併stderr到stdout
                text=True,
                bufsize=0,
                universal_newlines=True,
                env=env
            )
            
            self.output_queue = queue.Queue()
            self.waiting_for_input = False  # 初始化時不等待輸入
            self.status_var.set("運行中...")
            self.append_output(f"🚀 啟動腳本: {os.path.basename(self.script_path)}", "success")
            
            # 啟動讀取線程
            self.read_thread = threading.Thread(target=self.read_output, daemon=True)
            self.read_thread.start()
            
            # 開始處理輸出
            self.process_output()
            
        except Exception as e:
            self.append_output(f"❌ 啟動腳本失敗: {str(e)}", "error")
            self.status_var.set("啟動失敗")
            
    def read_output(self):
        """讀取進程輸出的線程函數"""
        import time
        
        buffer = ""
        last_char_time = time.time()
        
        try:
            while self.process and self.process.poll() is None:
                try:
                    # 逐字符讀取
                    char = self.process.stdout.read(1)
                    if char:
                        buffer += char
                        last_char_time = time.time()
                        
                        # 如果遇到換行符，立即輸出
                        if char == '\n':
                            if buffer.strip():
                                self.output_queue.put(('output', buffer.rstrip()))
                            buffer = ""
                    else:
                        # 檢查是否有未輸出的提示符（無換行符的情況）
                        current_time = time.time()
                        if buffer and (current_time - last_char_time) > 0.8:
                            # 檢查是否包含提示關鍵詞
                            if any(keyword in buffer for keyword in ['請輸入', '選擇', ': ', '?', '確認']):
                                self.output_queue.put(('output', buffer.rstrip()))
                                self.output_queue.put(('prompt', True))  # 標記為提示
                                buffer = ""
                                last_char_time = current_time
                        
                        time.sleep(0.01)
                        
                except Exception as e:
                    break
            
            # 處理剩餘輸出
            if buffer.strip():
                self.output_queue.put(('output', buffer.rstrip()))
                
            # 標記進程結束
            if self.process:
                self.output_queue.put(('status', f'exit_{self.process.returncode}'))
                
        except Exception as e:
            self.output_queue.put(('error', f"讀取輸出錯誤: {str(e)}"))
            
    def check_for_input_prompt(self, text):
        """檢查文本是否包含輸入提示"""
        try:
            # 檢查是否有常見的輸入提示
            input_prompts = [
                '請輸入', '選擇', '確認', '輸入', 'input', 'enter',
                ': ', '? ', '：', '？', '(y/n)', '[Enter]', '預設:', '(default)', '(Y/n)', '(y/N)'
            ]
            
            text_lower = text.lower()
            prompt_detected = any(prompt.lower() in text_lower for prompt in input_prompts)
            
            if prompt_detected and not self.waiting_for_input:
                self.waiting_for_input = True
                # 使用延遲設置焦點，確保可靠性
                self.window.after(10, lambda: self.input_entry.focus_set())
                self.status_var.set("等待輸入...")
                return True
                
        except Exception as e:
            pass  # 忽略檢查錯誤
        return False
    
            
    def process_output(self):
        """處理輸出隊列 - 更頻繁檢查版本"""
        try:
            messages_processed = 0
            while messages_processed < 20:  # 限制每次處理的消息數量
                try:
                    msg_type, content = self.output_queue.get_nowait()
                    messages_processed += 1
                    
                    if msg_type == 'output':
                        self.display_output(content)
                    elif msg_type == 'status':
                        if content == 'success':
                            self.status_var.set("執行完成")
                            self.append_output("✅ 腳本執行完成", "success")
                            self.execution_successful = True
                        elif content == 'error':
                            self.status_var.set("執行失敗")
                            self.append_output("❌ 腳本執行失敗", "error")
                        elif content.startswith('exit_'):
                            exit_code = content.split('_')[1]
                            if exit_code == '0':
                                self.status_var.set("執行完成")
                                self.append_output("✅ 腳本執行完成", "success")
                                self.execution_successful = True
                            else:
                                self.status_var.set("執行失敗")
                                self.append_output(f"❌ 腳本執行失敗 (退出碼: {exit_code})", "error")
                    elif msg_type == 'prompt':
                        self.waiting_for_input = True
                        # 確保輸入框始終有焦點
                        self.window.after(10, lambda: self.input_entry.focus_set())
                        self.status_var.set("等待輸入...")
                    elif msg_type == 'error':
                        self.append_output(content, "error")
                        
                except:
                    break
                    
        except Exception:
            pass
        
        # 檢查窗口是否還存在
        try:
            # 更頻繁地處理輸出隊列
            if self.process and self.process.poll() is None:
                self.window.after(50, self.process_output)  # 每50ms檢查一次
            else:
                # 進程結束，繼續處理一段時間以確保處理完所有輸出
                self.window.after(100, self.process_output)
        except:
            # 窗口已關閉
            pass
            
    def display_output(self, text):
        """顯示輸出文本"""
        # 不要分割任何輸出，保持完整性
        # 移除分割邏輯以避免截斷
        
        # 檢測不同類型的消息並應用樣式
        if "錯誤" in text or "ERROR" in text or "❌" in text:
            tag = "error"
        elif "警告" in text or "WARNING" in text or "⚠️" in text:
            tag = "warning"
        elif "成功" in text or "SUCCESS" in text or "✅" in text:
            tag = "success"
        elif "請選擇" in text or "請輸入" in text or "選擇" in text:
            tag = "prompt"
        else:
            tag = None
            
        self.append_output(text, tag)
        
    def append_output(self, text, tag=None):
        """添加文本到輸出區域"""
        self.output_text.config(state=tk.NORMAL)
        self.output_text.insert(tk.END, text + "\n", tag)
        self.output_text.config(state=tk.DISABLED)
        self.output_text.see(tk.END)
        
    def send_input(self, _=None):
        """發送用戶輸入 - 簡化版本（輸入框始終啟用）"""
        # 檢查進程是否存在且運行中
        if not self.process or self.process.poll() is not None:
            self.append_output("⚠️ 沒有運行中的腳本", "warning")
            return
            
        user_input = self.input_var.get()
        
        # 新邏輯：空輸入不允許，需要明確輸入 0 來使用預設值
        if not user_input.strip():
            self.append_output("⚠️ 請輸入有效值，或輸入 '0' 使用預設值", "warning")
            return
        
        # 處理預設值邏輯
        if user_input.strip() == "0":
            self.append_output(">>> 0 [使用預設值]", "input")
            user_input = ""  # 發送空字符串給腳本使用預設值
        else:
            self.append_output(f">>> {user_input}", "input")
            
        # 發送到腳本
        try:
            input_line = user_input + "\n"
            self.process.stdin.write(input_line)
            self.process.stdin.flush()
            if hasattr(self, 'waiting_for_input'):
                self.waiting_for_input = False
            self.status_var.set("運行中...")
        except Exception as e:
            self.append_output(f"❌ 發送輸入失敗: {str(e)}", "error")
                
        self.input_var.set("")
        
    def clear_output(self):
        """清除輸出"""
        self.output_text.config(state=tk.NORMAL)
        self.output_text.delete(1.0, tk.END)
        self.output_text.config(state=tk.DISABLED)
        
    
    def terminate_script(self):
        """終止腳本 - subprocess版本"""
        if self.process and self.process.poll() is None:
            try:
                self.process.terminate()
                self.append_output("⚠️ 腳本已被用戶終止", "warning")
                self.status_var.set("已終止")
                if self.update_job:
                    self.window.after_cancel(self.update_job)
            except Exception as e:
                self.append_output(f"❌ 終止腳本失敗: {str(e)}", "error")
                
    def close_terminal(self):
        """關閉終端 - subprocess版本"""
        if self.process and self.process.poll() is None:
            result = messagebox.askyesno("確認", "腳本仍在運行，確定要關閉嗎？", parent=self.window)
            if result:
                self.terminate_script()
            else:
                return
        
        # 清理資源
        try:
            if self.update_job:
                self.window.after_cancel(self.update_job)
        except:
            pass
            
        self.window.destroy()

class InputDialog:
    """輸入對話框"""
    
    def __init__(self, parent, title, message, default_value=""):
        self.result = None
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("400x200")
        self.window.resizable(False, False)
        
        # 設置窗口居中
        self.window.transient(parent)
        self.window.grab_set()
        
        # 創建界面
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 標題
        title_label = ttk.Label(main_frame, text=title, font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 訊息
        msg_label = ttk.Label(main_frame, text=message, wraplength=350)
        msg_label.pack(pady=(0, 10))
        
        # 輸入框
        self.entry_var = tk.StringVar(value=default_value)
        self.entry = ttk.Entry(main_frame, textvariable=self.entry_var, width=40)
        self.entry.pack(pady=(0, 20))
        self.entry.focus()
        self.entry.select_range(0, tk.END)
        
        # 按鈕區域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="確定", command=self.ok_clicked).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="取消", command=self.cancel_clicked).pack(side=tk.RIGHT)
        
        # 綁定事件
        self.entry.bind('<Return>', lambda e: self.ok_clicked())
        self.window.bind('<Escape>', lambda e: self.cancel_clicked())
        
        # 居中顯示
        self.window.update_idletasks()
        x = (parent.winfo_rootx() + parent.winfo_width() // 2) - (self.window.winfo_width() // 2)
        y = (parent.winfo_rooty() + parent.winfo_height() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        
    def ok_clicked(self):
        self.result = self.entry_var.get()
        self.window.destroy()
        
    def cancel_clicked(self):
        self.result = None
        self.window.destroy()

class ChoiceDialog:
    """選擇對話框"""
    
    def __init__(self, parent, title, message, choices):
        self.result = None
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("400x300")
        self.window.resizable(False, False)
        
        # 設置窗口居中
        self.window.transient(parent)
        self.window.grab_set()
        
        # 創建界面
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 標題
        title_label = ttk.Label(main_frame, text=title, font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 訊息
        msg_label = ttk.Label(main_frame, text=message)
        msg_label.pack(pady=(0, 10))
        
        # 選項列表
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        # 添加選項
        for choice in choices:
            self.listbox.insert(tk.END, choice)
            
        # 預設選擇第一個
        if choices:
            self.listbox.selection_set(0)
            
        # 按鈕
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ok_btn = ttk.Button(button_frame, text="確定", command=self.ok_clicked)
        ok_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        cancel_btn = ttk.Button(button_frame, text="取消", command=self.cancel_clicked)
        cancel_btn.pack(side=tk.RIGHT)
        
        # 綁定雙擊事件
        self.listbox.bind('<Double-1>', lambda e: self.ok_clicked())
        
        # 綁定按鍵事件
        self.window.bind('<Return>', lambda e: self.ok_clicked())
        self.window.bind('<Escape>', lambda e: self.cancel_clicked())
        
        # 等待用戶操作
        self.window.wait_window()
        
    def ok_clicked(self):
        """確定按鈕點擊"""
        selection = self.listbox.curselection()
        if selection:
            self.result = selection[0]
        self.window.destroy()
        
    def cancel_clicked(self):
        """取消按鈕點擊"""
        self.result = None
        self.window.destroy()

def main():
    """主函數"""
    # 檢查是否在正確的目錄中
    if not os.path.exists("1_folder_merger.py"):
        messagebox.showerror(
            "錯誤", 
            "請在包含 Auto Tronc 腳本文件的目錄中運行此程序！\n"
            "需要的文件包括: 1_folder_merger.py, 2_scorm_packager.py 等"
        )
        return
        
    # 創建主窗口
    root = tk.Tk()
    app = AutoTroncGUI(root)
    
    # 顯示歡迎信息
    welcome_msg = """
🚀 Auto Tronc 自動創課系統已啟動！

功能特色:
✅ 一鍵執行完整工作流程
✅ 分步驟執行各個功能
✅ 實時查看執行日誌
✅ 直接編輯Excel配置文件
✅ 進度追蹤和狀態監控

使用說明:
1. 可以點擊單個步驟按鈕執行特定功能
2. 或點擊"執行完整工作流程"自動執行所有步驟
3. 需要編輯Excel時，點擊"打開Excel編輯"
4. 所有執行過程都會在右側日誌中顯示

快捷鍵:
- F5: 刷新狀態
- Ctrl+Q: 退出程序
"""
    
    app.log_message(welcome_msg)
    
    # 運行主循環
    root.mainloop()

if __name__ == "__main__":
    main()