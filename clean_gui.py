import os
import re

file_path = r'c:\c\code_workspace\support\text_clicker_gui.py'

with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# 1. Fix the botched test_hp and macro function
# We need to find where test_hp starts and where the next valid function starts.
# Looking at the previous output, test_hp was interrupted by the macro insertion.

new_test_hp = '''    def test_hp(self, silent=False):
        self.monitor._log_callback = self.log
        window = self.window_title_var.get().strip()
        roi = self.get_roi()
        self.log(f"測試模式: 正在擷取 {window} 畫面...")
        
        curr, m_hp, ratio = self.monitor.detect_hp(window, roi)
        if m_hp > 0:
            self.log(f"測試結果: {curr} / {m_hp} ({ratio:.1%})")
            # 更新介面
            self.update_hp_display(curr, m_hp, ratio)
            if not silent:
                messagebox.showinfo("測試結果", f"體力偵測成功！\n數值: {curr} / {m_hp} ({ratio:.1%})")
        else:
            self.log("測試失敗: 無法讀取體力數值")
            if not silent:
                messagebox.showwarning("測試失敗", "無法從目前畫面讀取體力數值，請檢查位置設定。")

    def _perform_macro_keypress(self, vk_code):
        """核心按鍵觸發 (純鍵盤模擬，移除滑鼠點擊)"""
        # 取得掃描碼 (Scan Code)
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
        return False
'''

# Find the start of test_hp and everything until the next valid function start (toggle_key_spam)
# toggle_key_spam starts with "    def toggle_key_spam(self):"

pattern_mess = re.compile(r'    def test_hp\(self, silent=False\):.*?    def toggle_key_spam\(self\):', re.DOTALL)
content = pattern_mess.sub(new_test_hp + '\n    def toggle_key_spam(self):', content)

# 4. Final check on action_loop (ensure it's clean)
pattern_action_fix = re.compile(r'if action\[0\] == \'heal\':.*?for i in range\(presses\):.*?time\.sleep\(interval\)', re.DOTALL)
# This one should already be correct from the previous pass, but let's be sure.

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Cleaned up and Optimized")
