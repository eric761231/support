import os

file_path = r'c:\c\code_workspace\support\text_clicker_gui.py'

with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# 1. Update row3 in setup_ui to include Duration and Interval
target1 = '''        ttk.Label(row3, text="補血點(%):").pack(side=tk.LEFT)
        self.threshold_var = tk.StringVar(value="70")
        ttk.Spinbox(row3, from_=1, to=100, increment=5, textvariable=self.threshold_var, width=5).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row3, text="按鍵:").pack(side=tk.LEFT, padx=(15, 0))
        self.key_var = tk.StringVar(value="f5")
        ttk.Entry(row3, textvariable=self.key_var, width=5).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row3, text="次數:").pack(side=tk.LEFT, padx=(15, 0))
        self.presses_var = tk.StringVar(value="2")
        ttk.Spinbox(row3, from_=1, to=10, increment=1, textvariable=self.presses_var, width=5).pack(side=tk.LEFT, padx=5)'''

# Note: The target might have garbled text in my view, so I'll use a more robust regex or anchor
# Actually, I'll search for the structure.

import re

# Replace the row3 section
pattern_row3 = re.compile(r'row3 = ttk\.Frame\(settings_frame\).*?presses_var.*?width=5\)\.pack\(side=tk\.LEFT, padx=5\)', re.DOTALL)
replacement_row3 = '''row3 = ttk.Frame(settings_frame)
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
        ttk.Entry(row3, textvariable=self.interval_var, width=5).pack(side=tk.LEFT, padx=5)'''

content = pattern_row3.sub(replacement_row3, content)

# 2. Update _perform_macro_keypress
pattern_macro = re.compile(r'def _perform_macro_keypress\(self, vk_code\):.*?return False\s+return False', re.DOTALL)
# The "return False return False" is due to the previous messy merges/edits in the file.
# I'll replace the whole function.

replacement_macro = '''def _perform_macro_keypress(self, vk_code):
        """核心按鍵觸發 (純鍵盤模擬，移除滑鼠點擊)"""
        # 取得掃描碼 (Scan Code) 使模擬更貼近真實硬體
        scan_code = ctypes.windll.user32.MapVirtualKeyW(vk_code, 0)
        
        # 取得時間參數
        try:
            duration = float(self.duration_var.get()) / 1000.0
        except:
            duration = 0.05
            
        # 找到遊戲視窗
        window = self.window_title_var.get().strip()
        game_hwnd = self.monitor._find_window_by_keyword(window)
        
        if game_hwnd:
            # 確保視窗在最前端
            curr_hwnd = ctypes.windll.user32.GetForegroundWindow()
            if curr_hwnd != game_hwnd:
                ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)
                ctypes.windll.user32.SetForegroundWindow(game_hwnd)
                time.sleep(0.05)
            
            try:
                # 執行按鍵按下
                ctypes.windll.user32.keybd_event(vk_code, scan_code, 0, 0)
                time.sleep(duration)
                # 執行按鍵放開 (KEYEVENTF_KEYUP = 2)
                ctypes.windll.user32.keybd_event(vk_code, scan_code, 2, 0)
                
                return True
            except Exception as e:
                self.log(f"[按鍵失敗] {e}")
                return False
        return False'''

content = pattern_macro.sub(replacement_macro, content)

# 3. Update action_loop heal logic to use multiple presses with interval
pattern_action = re.compile(r'if action\[0\] == \'heal\':.*?self\._perform_macro_keypress\(vk_code\)', re.DOTALL)
replacement_action = '''if action[0] == 'heal':
                    _, vk_code, presses = action
                    try:
                        interval = float(self.interval_var.get()) / 1000.0
                    except:
                        interval = 0.15
                        
                    self.log(f"[觸發補血] 正在執行 {presses} 次按鍵...")
                    for i in range(presses):
                        self._perform_macro_keypress(vk_code)
                        if i < presses - 1:
                            time.sleep(interval)'''

content = pattern_action.sub(replacement_action, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Success")
