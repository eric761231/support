import os
import re

file_path = r'c:\c\code_workspace\support\text_clicker_gui.py'

with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Scan for the problematic f-string around line 179
# We'll replace the block from "def test_hp" to "_perform_macro_keypress" properly.

start_idx = -1
end_idx = -1

for i, line in enumerate(lines):
    if 'def test_hp(self, silent=False):' in line:
        start_idx = i
    if 'def toggle_key_spam(self):' in line:
        end_idx = i
        break

if start_idx != -1 and end_idx != -1:
    new_block = [
        '    def test_hp(self, silent=False):\n',
        '        self.monitor._log_callback = self.log\n',
        '        window = self.window_title_var.get().strip()\n',
        '        roi = self.get_roi()\n',
        '        self.log(f"Testing OCR on {window}...")\n',
        '        \n',
        '        curr, m_hp, ratio = self.monitor.detect_hp(window, roi)\n',
        '        if m_hp > 0:\n',
        '            self.log(f"Result: {curr} / {m_hp} ({ratio:.1%})")\n',
        '            self.update_hp_display(curr, m_hp, ratio)\n',
        '            if not silent:\n',
        '                messagebox.showinfo("Success", f"HP Detected: {curr} / {m_hp} ({ratio:.1%})")\n',
        '        else:\n',
        '            self.log("Failed: Could not read HP values")\n',
        '            if not silent:\n',
        '                messagebox.showwarning("Failed", "Could not read HP from screen. Check ROI settings.")\n',
        '\n',
        '    def _perform_macro_keypress(self, vk_code):\n',
        '        """Keyboard-only input simulation using ScanCodes and timing."""\n',
        '        scan_code = ctypes.windll.user32.MapVirtualKeyW(vk_code, 0)\n',
        '        window = self.window_title_var.get().strip()\n',
        '        game_hwnd = self.monitor._find_window_by_keyword(window)\n',
        '        \n',
        '        if game_hwnd:\n',
        '            curr_hwnd = ctypes.windll.user32.GetForegroundWindow()\n',
        '            if curr_hwnd != game_hwnd:\n',
        '                ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)\n',
        '                ctypes.windll.user32.SetForegroundWindow(game_hwnd)\n',
        '                time.sleep(0.05)\n',
        '            \n',
        '            try:\n',
        '                duration = float(self.duration_var.get()) / 1000.0\n',
        '            except:\n',
        '                duration = 0.05\n',
        '            \n',
        '            try:\n',
        '                ctypes.windll.user32.keybd_event(vk_code, scan_code, 0, 0)\n',
        '                time.sleep(duration)\n',
        '                ctypes.windll.user32.keybd_event(vk_code, scan_code, 2, 0)\n',
        '                return True\n',
        '            except Exception as e:\n',
        '                self.log(f"Key error: {e}")\n',
        '                return False\n',
        '        return False\n',
        '\n'
    ]
    lines[start_idx:end_idx] = new_block

with open(file_path, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Fixed Syntax Errors")
