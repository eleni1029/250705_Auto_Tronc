#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
功能完整的交互式終端執行器
支持用戶輸入交互和實時輸出顯示
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
import subprocess
import threading
import os
import sys
import queue
import time
from datetime import datetime


class FinalTerminal:
    def __init__(self, parent, task_name, script_path):
        """
        初始化交互式終端執行器
        
        Args:
            parent: 父窗口
            task_name: 任務名稱
            script_path: 腳本路徑
        """
        self.parent = parent
        self.task_name = task_name
        self.script_path = script_path
        self.execution_successful = False
        self.process = None
        self.output_queue = queue.Queue()
        self.input_queue = queue.Queue()
        self.is_running = False
        self.waiting_for_input = False
        
        self.create_window()
        self.start_execution()
    
    def create_window(self):
        """創建終端窗口"""
        self.window = tk.Toplevel(self.parent)
        self.window.title(f"交互式終端 - {self.task_name}")
        self.window.geometry("900x700")
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # 創建主框架
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 標題框架
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = ttk.Label(title_frame, 
                               text=f"🖥️  {self.task_name}", 
                               font=("Arial", 14, "bold"))
        title_label.pack(side=tk.LEFT)
        
        # 狀態標籤
        self.status_var = tk.StringVar(value="正在啟動...")
        self.status_label = ttk.Label(title_frame, 
                                     textvariable=self.status_var, 
                                     font=("Arial", 10), 
                                     foreground="blue")
        self.status_label.pack(side=tk.RIGHT)
        
        # 腳本路徑
        path_label = ttk.Label(main_frame, 
                              text=f"腳本：{self.script_path}", 
                              font=("Arial", 9),
                              foreground="gray")
        path_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 終端輸出區域
        output_frame = ttk.LabelFrame(main_frame, text="終端輸出", padding=5)
        output_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.output_text = scrolledtext.ScrolledText(output_frame, 
                                                    height=25, 
                                                    font=("Consolas", 10),
                                                    bg="#1e1e1e",
                                                    fg="#ffffff",
                                                    insertbackground="#ffffff",
                                                    selectbackground="#404040",
                                                    wrap=tk.WORD)
        self.output_text.pack(fill=tk.BOTH, expand=True)
        
        # 配置文本標籤樣式
        self.output_text.tag_configure("error", foreground="#ff6b6b")
        self.output_text.tag_configure("warning", foreground="#feca57")
        self.output_text.tag_configure("success", foreground="#48ca61")
        self.output_text.tag_configure("input", foreground="#74b9ff")
        self.output_text.tag_configure("prompt", foreground="#fdcb6e")
        self.output_text.tag_configure("timestamp", foreground="#636e72")
        
        # 輸入區域
        input_frame = ttk.LabelFrame(main_frame, text="用戶輸入", padding=5)
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 輸入提示
        self.input_prompt_var = tk.StringVar(value="等待腳本輸出...")
        prompt_label = ttk.Label(input_frame, 
                                textvariable=self.input_prompt_var, 
                                font=("Consolas", 9),
                                foreground="darkblue")
        prompt_label.pack(anchor=tk.W, pady=(0, 5))
        
        # 輸入控件框架
        input_control_frame = ttk.Frame(input_frame)
        input_control_frame.pack(fill=tk.X)
        
        ttk.Label(input_control_frame, text="輸入:", font=("Consolas", 10)).pack(side=tk.LEFT, padx=(0, 5))
        
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(input_control_frame, 
                                    textvariable=self.input_var, 
                                    font=("Consolas", 10),
                                    state='disabled')
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.input_entry.bind('<Return>', self.send_input)
        self.input_entry.bind('<KP_Enter>', self.send_input)
        
        self.send_button = ttk.Button(input_control_frame, 
                                     text="發送", 
                                     command=self.send_input,
                                     state='disabled')
        self.send_button.pack(side=tk.RIGHT)
        
        # 提示信息
        hint_label = ttk.Label(input_frame, 
                              text="💡 提示：Enter使用預設值，q退出 | F2/Ctrl+I手動啟用輸入 | ⌨️按鈕手動啟用", 
                              font=("Arial", 8), 
                              foreground="gray")
        hint_label.pack(anchor=tk.W, pady=(5, 0))
        
        # 進度條
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, 
                                          variable=self.progress_var, 
                                          mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # 控制按鈕框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        # 終止按鈕
        self.terminate_button = ttk.Button(button_frame, 
                                          text="🛑 終止腳本", 
                                          command=self.terminate_script)
        self.terminate_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # 清除輸出按鈕
        clear_button = ttk.Button(button_frame, 
                                 text="🧹 清除輸出", 
                                 command=self.clear_output)
        clear_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # 手動啟用輸入按鈕
        self.manual_input_button = ttk.Button(button_frame, 
                                             text="⌨️ 啟用輸入", 
                                             command=self.manual_enable_input)
        self.manual_input_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # 關閉按鈕（初始時禁用）
        self.close_button = ttk.Button(button_frame, 
                                      text="❌ 關閉", 
                                      command=self.close_window,
                                      state='disabled')
        self.close_button.pack(side=tk.RIGHT)
        
        # 綁定窗口關閉事件
        self.window.protocol("WM_DELETE_WINDOW", self.on_window_close)
        
        # 綁定鍵盤快捷鍵
        self.window.bind('<F2>', lambda e: self.manual_enable_input())  # F2快速啟用輸入
        self.window.bind('<Control-i>', lambda e: self.manual_enable_input())  # Ctrl+I啟用輸入
    
    def append_output(self, text, tag=None):
        """追加輸出文本"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 插入時間戳
        self.output_text.insert(tk.END, f"[{timestamp}] ", "timestamp")
        
        # 插入正文
        if tag:
            self.output_text.insert(tk.END, f"{text}\n", tag)
        else:
            self.output_text.insert(tk.END, f"{text}\n")
        
        self.output_text.see(tk.END)
        self.window.update_idletasks()
    
    def update_input_prompt(self, prompt_text):
        """更新輸入提示"""
        self.input_prompt_var.set(prompt_text)
        
    def enable_input(self):
        """啟用輸入控件"""
        self.input_entry.config(state='normal')
        self.send_button.config(state='normal')
        self.input_entry.focus_set()
        self.waiting_for_input = True
        
    def disable_input(self):
        """禁用輸入控件"""
        self.input_entry.config(state='disabled')
        self.send_button.config(state='disabled')
        self.waiting_for_input = False
    
    def manual_enable_input(self):
        """手動啟用輸入功能 - 智能解決方案"""
        if not self.is_running:
            self.append_output("⚠️ 腳本未運行，無法啟用輸入", "warning")
            return
            
        # 強制啟用輸入
        self.enable_input()
        self.update_input_prompt("⌨️ 手動啟用輸入 - 請輸入內容或按Enter使用預設值")
        self.append_output("🔓 輸入功能已手動啟用，可以進行輸入", "success")
        
        # 如果有最後一行輸出包含提示，重新顯示
        last_lines = self.output_text.get("end-10c linestart", "end-1c").strip().split('\n')
        if last_lines:
            last_line = last_lines[-1].strip()
            if last_line and not last_line.startswith('['):  # 不是時間戳行
                self.update_input_prompt(f"根據輸出推測: {last_line}")
                
        # 更新按鈕狀態
        self.manual_input_button.config(text="✅ 輸入已啟用", state='disabled')
        
        # 自動重新啟用按鈕（以防需要再次手動啟用）
        self.window.after(5000, lambda: self.manual_input_button.config(text="⌨️ 啟用輸入", state='normal'))
    
    def start_execution(self):
        """開始執行腳本"""
        self.is_running = True
        self.progress_bar.start()
        self.append_output(f"開始執行腳本：{self.script_path}", "success")
        self.status_var.set("運行中...")
        
        # 在後台線程中執行
        thread = threading.Thread(target=self._execute_script)
        thread.daemon = True
        thread.start()
        
        # 開始處理輸出
        self.process_output()
    
    def _execute_script(self):
        """在後台執行腳本"""
        try:
            # 判斷腳本類型並構建命令
            if self.script_path.endswith('.py'):
                cmd = [sys.executable, '-u', self.script_path]
            elif self.script_path.endswith('.sh'):
                cmd = ['bash', self.script_path]
            else:
                cmd = [self.script_path]
            
            self.output_queue.put(('output', f"執行命令：{' '.join(cmd)}"))
            
            # 設置環境變量
            env = os.environ.copy()
            env['PYTHONUNBUFFERED'] = '1'
            env['PYTHONIOENCODING'] = 'utf-8'
            
            # 啟動進程
            self.process = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=0,
                universal_newlines=True,
                env=env,
                cwd=os.getcwd()
            )
            
            # 啟動輸出讀取線程
            output_thread = threading.Thread(target=self._read_output, daemon=True)
            output_thread.start()
            
            # 處理用戶輸入
            while self.is_running and self.process.poll() is None:
                try:
                    user_input = self.input_queue.get(timeout=0.1)
                    if user_input is not None:
                        # 發送用戶輸入到進程
                        self.process.stdin.write(user_input + '\n')
                        self.process.stdin.flush()
                        self.output_queue.put(('user_input', f">>> {user_input}"))
                except queue.Empty:
                    continue
                except Exception as e:
                    self.output_queue.put(('error', f"發送輸入錯誤: {str(e)}"))
                    break
            
            # 等待進程結束
            if self.process:
                self.process.wait()
                exit_code = self.process.returncode
                
                if exit_code == 0:
                    self.execution_successful = True
                    self.output_queue.put(('status', 'success'))
                else:
                    self.output_queue.put(('status', f'failed_{exit_code}'))
                    
        except Exception as e:
            self.output_queue.put(('error', f"執行過程中發生錯誤：{str(e)}"))
            self.output_queue.put(('status', 'error'))
        finally:
            self.is_running = False
    
    def _read_output(self):
        """讀取進程輸出的線程函數"""
        buffer = ""
        last_char_time = time.time()
        
        try:
            while self.is_running and self.process and self.process.poll() is None:
                try:
                    # 逐字符讀取以便檢測提示符
                    char = self.process.stdout.read(1)
                    if char:
                        buffer += char
                        last_char_time = time.time()
                        
                        # 如果遇到換行符，立即輸出
                        if char == '\n':
                            if buffer.strip():
                                self.output_queue.put(('output', buffer.rstrip()))
                                # 檢查是否是輸入提示
                                if self._is_input_prompt(buffer.strip()):
                                    self.output_queue.put(('prompt', buffer.strip()))
                            buffer = ""
                    else:
                        # 檢查是否有未輸出的提示符（無換行符的情況）
                        current_time = time.time()
                        if buffer and (current_time - last_char_time) > 0.3:  # 減少等待時間
                            # 更積極的提示檢測
                            buffer_text = buffer.strip()
                            is_prompt = (
                                self._is_input_prompt(buffer_text) or 
                                buffer_text.endswith(':') or 
                                buffer_text.endswith('\\') or
                                '預設' in buffer_text or
                                len(buffer_text) > 10 and not buffer_text.endswith('...')  # 可能是完整提示
                            )
                            
                            if is_prompt:
                                self.output_queue.put(('output', buffer.rstrip()))
                                self.output_queue.put(('prompt', buffer.strip()))
                                buffer = ""
                                last_char_time = current_time
                        
                        time.sleep(0.01)
                        
                except Exception as e:
                    break
            
            # 處理剩餘輸出
            if buffer.strip():
                self.output_queue.put(('output', buffer.rstrip()))
                
        except Exception as e:
            self.output_queue.put(('error', f"讀取輸出錯誤: {str(e)}"))
    
    def _is_input_prompt(self, text):
        """檢查文本是否包含輸入提示"""
        input_indicators = [
            '請輸入', '選擇', '確認', '輸入', 'input', 'enter',
            ': ', '? ', '：', '？', '(y/n)', '[Enter]', '預設:', 
            '(default)', '(Y/n)', '(y/N)', '>>>', 'choice',
            '或按Enter', '請選擇', 'Select', 'Choose', '\\', '\\ '
        ]
        
        text_lower = text.lower()
        
        # 檢查是否包含輸入指示符
        has_indicator = any(indicator.lower() in text_lower for indicator in input_indicators)
        
        # 特別檢查以 \ 結尾的行（shell續行提示符）
        if text.strip().endswith('\\'):
            return True
            
        # 檢查是否包含預設值提示格式，如 "(預設: xxx):"
        if '預設:' in text and ':' in text:
            return True
            
        return has_indicator
    
    def process_output(self):
        """處理輸出隊列"""
        try:
            processed_count = 0
            while processed_count < 10:  # 限制每次處理的消息數量
                try:
                    msg_type, content = self.output_queue.get_nowait()
                    processed_count += 1
                    
                    if msg_type == 'output':
                        # 檢測消息類型並應用樣式
                        if any(keyword in content for keyword in ['錯誤', 'ERROR', '❌', 'Error']):
                            self.append_output(content, "error")
                        elif any(keyword in content for keyword in ['警告', 'WARNING', '⚠️', 'Warning']):
                            self.append_output(content, "warning")
                        elif any(keyword in content for keyword in ['成功', 'SUCCESS', '✅', 'Success']):
                            self.append_output(content, "success")
                        elif self._is_input_prompt(content):
                            self.append_output(content, "prompt")
                        else:
                            self.append_output(content)
                            
                    elif msg_type == 'user_input':
                        self.append_output(content, "input")
                        self.disable_input()
                        
                    elif msg_type == 'prompt':
                        self.update_input_prompt(f"等待輸入: {content}")
                        self.enable_input()
                        
                    elif msg_type == 'status':
                        if content == 'success':
                            self.status_var.set("執行完成")
                            self.append_output("✅ 腳本執行完成", "success")
                            self._execution_finished()
                        elif content.startswith('failed_'):
                            exit_code = content.split('_')[1]
                            self.status_var.set("執行失敗")
                            self.append_output(f"❌ 腳本執行失敗 (退出碼: {exit_code})", "error")
                            self._execution_finished()
                        elif content == 'error':
                            self.status_var.set("執行錯誤")
                            self.append_output("❌ 腳本執行過程中發生錯誤", "error")
                            self._execution_finished()
                            
                    elif msg_type == 'error':
                        self.append_output(content, "error")
                        
                except queue.Empty:
                    break
                    
        except Exception:
            pass
        
        # 繼續處理隊列
        if self.is_running or not self.output_queue.empty():
            self.window.after(100, self.process_output)
    
    def send_input(self, event=None):
        """發送用戶輸入"""
        if not self.waiting_for_input:
            return
            
        user_input = self.input_var.get().strip()
        
        # 檢查是否是退出命令
        if user_input.lower() in ['q', 'quit', 'exit']:
            self.terminate_script()
            return
        
        # 將輸入放入隊列
        self.input_queue.put(user_input if user_input else "")
        
        # 清空輸入框
        self.input_var.set("")
    
    def clear_output(self):
        """清除輸出"""
        self.output_text.delete(1.0, tk.END)
        self.append_output("輸出已清除", "warning")
    
    def _execution_finished(self):
        """執行完成後的清理工作"""
        self.is_running = False
        self.progress_bar.stop()
        self.disable_input()
        self.terminate_button.config(state='disabled')
        self.close_button.config(state='normal')
        self.update_input_prompt("腳本已結束")
        
        # 顯示完成提示
        if self.execution_successful:
            self.append_output("🎉 腳本執行完成！您可以關閉此窗口。", "success")
        else:
            self.append_output("⚠️  腳本執行結束。請檢查上方的輸出日誌。", "warning")
    
    def terminate_script(self):
        """終止腳本"""
        if self.process and self.process.poll() is None:
            try:
                self.process.terminate()
                self.append_output("🛑 用戶終止執行", "warning")
                self.status_var.set("已終止")
                self.is_running = False
            except:
                pass
        
        self._execution_finished()
    
    def close_window(self):
        """關閉窗口"""
        if self.is_running and self.process and self.process.poll() is None:
            result = messagebox.askyesno(
                "確認關閉", 
                "腳本仍在執行中，確定要關閉窗口嗎？\n這將終止腳本執行。",
                parent=self.window
            )
            if result:
                self.terminate_script()
            else:
                return
        
        self.window.destroy()
    
    def on_window_close(self):
        """處理窗口關閉事件"""
        self.close_window()


# 測試程序
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # 隱藏主窗口
    
    # 測試終端
    terminal = FinalTerminal(root, "測試任務", "python3")
    root.wait_window(terminal.window)
    
    print(f"執行結果：{'成功' if terminal.execution_successful else '失敗'}")
    root.destroy()