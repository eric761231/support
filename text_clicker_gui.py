# -*- coding: utf-8 -*-
"""
體力監控工具 - 極簡 GUI
"""
import sys
import io
import os
import threading
import queue
import time
import tkinter as tk
import ctypes
import ctypes.wintypes
from tkinter import ttk, messagebox, scrolledtext
from text_clicker import HPMonitor

# 設定 UTF-8 編碼輸出 (Windows 環境下使用 reconfigure 更安全)
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        # 為了兼容舊版 Python (3.7 以下)，保留回退方案
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

class HPMonitorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("天堂簡易補血器")
        self.root.geometry("600x650")
        
        self.monitor = HPMonitor()
        self.is_running = False
        self.stop_flag = False
        self.log_queue = queue.Queue()
        self.action_queue = queue.Queue()
        self.last_heal_time = 0
        self.key_spam_flag = False  # 控制連點測試的旗標
        
        self.setup_ui()
        self.check_log_queue()
        
        # 延遲 1 秒後執行自動抓取與測試 (給予使用者和系統緩衝時間)
        self.root.after(1000, self.startup_sequence)
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(expand=True, fill="both")
        
        # 標題
        ttk.Label(main_frame, text="天堂自動補血監控", font=("Arial", 18, "bold")).pack(pady=(0, 20))
        
        # 參數設定區
        settings_frame = ttk.LabelFrame(main_frame, text="監控設定", padding="15")
        settings_frame.pack(fill="x", pady=5)
        
        # 視窗標題
        row1 = ttk.Frame(settings_frame)
        row1.pack(fill="x", pady=5)
        ttk.Label(row1, text="視窗名稱:", width=10).pack(side=tk.LEFT)
        self.window_title_var = tk.StringVar(value="Lineage")
        ttk.Entry(row1, textvariable=self.window_title_var).pack(side=tk.LEFT, expand=True, fill="x", padx=5)
        
        # HP 區域 (ROI)
        row2 = ttk.Frame(settings_frame)
        row2.pack(fill="x", pady=5)
        ttk.Label(row2, text="偵測區域:", width=10).pack(side=tk.LEFT)
        self.hp_roi_var = tk.StringVar(value="")
        ttk.Entry(row2, textvariable=self.hp_roi_var).pack(side=tk.LEFT, expand=True, fill="x", padx=5)
        ttk.Button(row2, text="🤖 自動抓取位置", command=self.auto_calibrate).pack(side=tk.LEFT)
        
        # 參數設定
        row3 = ttk.Frame(settings_frame)
        row3.pack(fill="x", pady=10)
        
        ttk.Label(row3, text="補血點(%):").pack(side=tk.LEFT)
        self.threshold_var = tk.StringVar(value="70")
        ttk.Spinbox(row3, from_=1, to=100, increment=5, textvariable=self.threshold_var, width=5).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row3, text="按鍵:").pack(side=tk.LEFT, padx=(15, 0))
        self.key_var = tk.StringVar(value="f5")
        ttk.Entry(row3, textvariable=self.key_var, width=5).pack(side=tk.LEFT, padx=5)

        ttk.Label(row3, text="按壓(ms):").pack(side=tk.LEFT, padx=(15, 0))
        self.duration_var = tk.StringVar(value="50")
        ttk.Entry(row3, textvariable=self.duration_var, width=5).pack(side=tk.LEFT, padx=5)

        ttk.Label(row3, text="間隔(ms):").pack(side=tk.LEFT, padx=(15, 0))
        self.interval_var = tk.StringVar(value="150")
        ttk.Entry(row3, textvariable=self.interval_var, width=5).pack(side=tk.LEFT, padx=5)

        # 檢測按鈕區
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="🔍 測試體力識別", command=self.test_hp).pack(side=tk.LEFT, padx=5)
        self.spam_btn = ttk.Button(btn_frame, text="⌨️ 啟動連點測試", command=self.toggle_key_spam)
        self.spam_btn.pack(side=tk.LEFT, padx=5)
        
        # 視覺化血條區
        hp_display_frame = ttk.LabelFrame(main_frame, text="即時體力狀態", padding="10")
        hp_display_frame.pack(fill="x", pady=5)
        
        self.hp_label = ttk.Label(hp_display_frame, text="0 / 0 (0%)", font=("Arial", 14, "bold"))
        self.hp_label.pack()
        
        self.hp_progress = ttk.Progressbar(hp_display_frame, length=500, mode='determinate')
        self.hp_progress.pack(pady=5, fill="x")
        
        # 控制按鈕
        ctrl_frame = ttk.Frame(main_frame)
        ctrl_frame.pack(pady=10)
        self.start_btn = ttk.Button(ctrl_frame, text="▶ 開始監控", command=self.start, width=15)
        self.start_btn.pack(side=tk.LEFT, padx=10)
        self.stop_btn = ttk.Button(ctrl_frame, text="⏹ 停止", command=self.stop, state=tk.DISABLED, width=15)
        self.stop_btn.pack(side=tk.LEFT, padx=10)
        
        # 日誌
        log_label = ttk.Label(main_frame, text="執行日誌:")
        log_label.pack(anchor="w")
        self.log_text = scrolledtext.ScrolledText(main_frame, height=12, wrap=tk.WORD)
        self.log_text.pack(expand=True, fill="both", pady=5)

    def log(self, msg):
        self.log_queue.put(msg)

    def check_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
                self.log_text.see(tk.END)
        except queue.Empty: pass
        self.root.after(100, self.check_log_queue)

    def get_roi(self):
        roi_str = self.hp_roi_var.get().strip()
        if not roi_str: return None
        try:
            return tuple(map(int, roi_str.replace(',', ' ').split()))
        except:
            messagebox.showerror("錯誤", "區域格式錯誤! 例: 10, 50, 150, 30")
            return None

    def auto_calibrate(self, silent=False):
        """調用引擎自動偵測血條位置"""
        self.monitor._log_callback = self.log
        window = self.window_title_var.get().strip()
        template = os.path.join("hpbar", "hp.jpg")
        
        self.log(f"正在全自動搜尋 '{window}' 內的體力位置...")
        roi = self.monitor.find_hp_roi(window, template)
        
        if roi:
            roi_str = f"{roi[0]}, {roi[1]}, {roi[2]}, {roi[3]}"
            self.hp_roi_var.set(roi_str)
            self.log(f"抓取成功！偵測區域已更新為: {roi_str}")
            if not silent:
                messagebox.showinfo("成功", f"自動抓取成功！\n位置: {roi_str}")
            return True
        else:
            if not silent:
                messagebox.showwarning("失敗", "無法在視窗內找到體力顯示。\n請確認遊戲視窗未被遮擋，或手動設定。")
            return False

    def test_hp(self, silent=False):
        self.monitor._log_callback = self.log
        window = self.window_title_var.get().strip()
        roi = self.get_roi()
        self.log(f"Testing OCR on {window}...")
        
        curr, m_hp, ratio = self.monitor.detect_hp(window, roi)
        if m_hp > 0:
            self.log(f"Result: {curr} / {m_hp} ({ratio:.1%})")
            self.update_hp_display(curr, m_hp, ratio)
            if not silent:
                messagebox.showinfo("Success", f"HP Detected: {curr} / {m_hp} ({ratio:.1%})")
        else:
            self.log("Failed: Could not read HP values")
            if not silent:
                messagebox.showwarning("Failed", "Could not read HP from screen. Check ROI settings.")

    def _perform_macro_keypress(self, vk_code):
        """Keyboard-only input simulation using ScanCodes and timing."""
        scan_code = ctypes.windll.user32.MapVirtualKeyW(vk_code, 0)
        window = self.window_title_var.get().strip()
        game_hwnd = self.monitor._find_window_by_keyword(window)
        
        if game_hwnd:
            curr_hwnd = ctypes.windll.user32.GetForegroundWindow()
            if curr_hwnd != game_hwnd:
                ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)
                ctypes.windll.user32.SetForegroundWindow(game_hwnd)
                time.sleep(0.05)
            
            try:
                duration = float(self.duration_var.get()) / 1000.0
            except:
                duration = 0.05
            
            try:
                ctypes.windll.user32.keybd_event(vk_code, scan_code, 0, 0)
                time.sleep(duration)
                ctypes.windll.user32.keybd_event(vk_code, scan_code, 2, 0)
                return True
            except Exception as e:
                self.log(f"Key error: {e}")
                return False
        return False

    def toggle_key_spam(self):
        """切換連點測試的開/關狀態"""
        if not self.key_spam_flag:
            self.key_spam_flag = True
            self.spam_btn.config(text="⏹ 停止連點測試")
            key = self.key_var.get().strip().lower()
            vk_keys = {'f5': 0x74, 'f6': 0x75, 'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B}
            vk_code = vk_keys.get(key, 0x74)
            self.log(f"[連點測試] 啟動按鍵連點模式 ({key.upper()})")
            threading.Thread(target=self._spam_loop, args=(vk_code,), daemon=True).start()
        else:
            self.key_spam_flag = False
            self.spam_btn.config(text="⌨️ 啟動連點測試")
            self.log("[連點測試] 已停止。")

    def _spam_loop(self, vk_code):
        """連點測試的無限迴圈"""
        while self.key_spam_flag:
            self._perform_macro_keypress(vk_code)
            # 每秒執行一次連點序列
            for _ in range(10):
                if not self.key_spam_flag: break
                time.sleep(0.1)

    def test_keypress(self):
        """手動測試按鍵連點 (單次巨集)"""
        key = self.key_var.get().strip().lower()
        vk_keys = {'f5': 0x74, 'f6': 0x75, 'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B}
        vk_code = vk_keys.get(key, 0x74)
        self.log(f"[測試按鍵] 試圖發送 {key.upper()} 到遊戲...")
        if self._perform_macro_keypress(vk_code):
            self.log(f"[測試按鍵] ✅ 巨集發送完成")
        else:
            self.log("[測試按鍵] ❌ 發送失敗")


    def startup_sequence(self):
        """啟動時自動執行的動作腳本"""
        self.log("=== 開始執行開機自動校正 ===")
        if self.auto_calibrate(silent=True):
            self.log("=== 位置抓取完成，1 秒後進行體力識別測試 ===")
            self.root.update() # 強制更新 GUI 介面顯示
            
            # 使用 after 延遲 1 秒，確保視窗與座標已經徹底準備完畢，不引發時序衝突
            self.root.after(1000, self._delayed_test_hp)
        else:
            self.log("=== 開機自動校正失敗 ===")

    def _delayed_test_hp(self):
        self.test_hp(silent=True)
        self.log("=== 開機自動校正完畢 ===")

    def start(self):
        self.is_running = True
        self.stop_flag = False
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # 清空佇列，確保沒有上次的殘留指令
        while not self.action_queue.empty():
            self.action_queue.get_nowait()
        
        # 啟動雙執行緒
        threading.Thread(target=self.run_loop, daemon=True).start()
        threading.Thread(target=self.action_loop, daemon=True).start()

    def stop(self):
        self.stop_flag = True
        self.log("正在停止監控...")

    def action_loop(self):
        """動作執行緒：接收補血指令並執行按鍵觸發"""
        self.log("動作執行緒已啟動 (純按鍵模式)")
        while not self.stop_flag:
            try:
                action = self.action_queue.get(timeout=0.1)
                if action[0] == 'heal':
                    _, vk_code, presses = action
                    try:
                        interval = float(self.interval_var.get()) / 1000.0
                    except:
                        interval = 0.15
                        
                    self.log(f"[觸發補血] 正在執行 {presses} 次按鍵...")
                    for i in range(presses):
                        self._perform_macro_keypress(vk_code)
                        if i < presses - 1:
                            time.sleep(interval)
            except queue.Empty:
                continue
            except Exception as e:
                self.log(f"[動作異常] {e}")

            except queue.Empty:
                continue
            except Exception as e:
                self.log(f"[動作異常] {e}")


    def run_loop(self):
        """監控執行緒：專門負責擷取畫面與辨識數字"""
        self.monitor._log_callback = self.log
        window = self.window_title_var.get().strip()
        threshold = float(self.threshold_var.get()) / 100.0
        key = self.key_var.get().strip().lower()
        presses = int(self.presses_var.get())
        
        # 鍵盤按鍵碼對應表 (VK Codes)
        vk_keys = {'f5': 0x74, 'f6': 0x75, 'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B}
        vk_code = vk_keys.get(key, 0x74)
        
        self.log(f"視覺監控開始! 門檻: {threshold:.0%}")
        
        try:
            while not self.stop_flag:
                loop_start_time = time.time()
                try:
                    roi = self.get_roi()
                    curr, m_hp, ratio = self.monitor.detect_hp(window, roi)
                    
                    if m_hp > 0:
                        # 更新 GUI (Thread-safe)
                        self.root.after(0, self.update_hp_display, curr, m_hp, ratio)
                        
                        # 如果低於門檻，並且冷卻時間已過 (1.0 秒)
                        if ratio < threshold:
                            current_time = time.time()
                            if current_time - self.last_heal_time > 1.0:
                                self.log(f"[觸發警告] 體力過低 ({ratio:.1%})，發送補血指令到動作佇列。")
                                # 將動作送進佇列給另一個執行緒執行
                                self.action_queue.put(('heal', vk_code, presses))
                                self.last_heal_time = current_time
                            else:
                                # self.log("[冷卻中] 角色可能正在喝水，忽略此次判定。")
                                pass
                                
                except Exception as loop_e:
                    self.log(f"[監控警告] 偵測發生異常: {loop_e}")
                    
                # 扣除掉 OCR 運算時間，確保穩定的每秒 2 次掃描
                elapsed = time.time() - loop_start_time
                sleep_time = max(0.1, 0.5 - elapsed)
                time.sleep(sleep_time)
                
        except Exception as e:
            self.log(f"[嚴重錯誤] 監控中斷: {e}")
            
        self.is_running = False
        self.root.after(0, lambda: [self.start_btn.config(state=tk.NORMAL), self.stop_btn.config(state=tk.DISABLED)])
        self.log("監控已停止。")

    def update_hp_display(self, curr, m_hp, ratio):
        """更新 GUI 上的血條與文字"""
        self.hp_label.config(text=f"{curr} / {m_hp} ({ratio:.1%})")
        self.hp_progress['value'] = ratio * 100
        
        # 根據比例變換顏色 (自定義樣式)
        style = ttk.Style()
        if ratio < 0.3:
            color = 'red'
        elif ratio < 0.7:
            color = 'orange'
        else:
            color = 'green'
        
        style.configure("TProgressbar", background=color)

def is_admin():
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if __name__ == "__main__":
    if not is_admin():
        # 如果不是管理員，則主動請求 UAC 權限重新啟動自己
        import ctypes
        import sys
        
        script = os.path.abspath(sys.argv[0])
        params = ' '.join([script] + sys.argv[1:])
        
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        except Exception as e:
            print(f"無法取得系統管理員權限: {e}")
        
        sys.exit()
        
    # 如果已經是系統管理員，啟動時自動縮小 CMD 視窗
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 6) # 6 = SW_MINIMIZE (最小化)
    except:
        pass
        
    root = tk.Tk()
    app = HPMonitorGUI(root)
    root.mainloop()
