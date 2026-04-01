import os
import re

file_path = r'c:\c\code_workspace\support\text_clicker.py'

with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# Add ctypes to imports if not present
if 'import ctypes' not in content:
    content = content.replace('import sys', 'import sys\nimport ctypes')

# Update check_and_trigger
pattern_trigger = re.compile(r'def check_and_trigger\(self,.*?\):.*?return False', re.DOTALL)
replacement_trigger = '''def check_and_trigger(self, window_title: Optional[str] = None, region: Optional[Tuple[int, int, int, int]] = None, 
                           threshold: float = 0.7, key: str = 'f5', presses: int = 2, duration: float = 0.05, interval: float = 0.15) -> bool:
        """核心監控邏輯 - 使用更可靠的鍵盤模擬"""
        curr, m_hp, ratio = self.detect_hp(window_title, region)
        if m_hp > 0:
            self._log_callback(f"[體力狀態] {curr} / {m_hp} ({ratio:.1%})")
            if ratio < threshold:
                self._log_callback(f"[觸發警告] 體力低於 {threshold:.0%}，正在執行 {presses} 次 {key.upper()}")
                
                # 簡單鍵盤映射 (常用 F5-F12)
                vk_keys = {
                    'f5': 0x74, 'f6': 0x75, 'f7': 0x76, 'f8': 0x77, 
                    'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B
                }
                vk_code = vk_keys.get(key.lower(), 0x74)
                scan_code = ctypes.windll.user32.MapVirtualKeyW(vk_code, 0)
                
                for i in range(presses):
                    # Key Down
                    ctypes.windll.user32.keybd_event(vk_code, scan_code, 0, 0)
                    time.sleep(duration)
                    # Key Up
                    ctypes.windll.user32.keybd_event(vk_code, scan_code, 2, 0)
                    
                    if i < presses - 1:
                        time.sleep(interval)
                return True
        return False'''

content = pattern_trigger.sub(replacement_trigger, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Success")
