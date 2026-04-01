# -*- coding: utf-8 -*-
"""
體力監控工具 - 專注於體力偵測與自動按鍵觸發 (F5 x2)
"""
import sys
import ctypes
import io
import pyautogui
import pytesseract
from PIL import Image
import cv2
import numpy as np
import time
from typing import Tuple, Optional
import win32gui
import os
import re

# UTF-8 編碼已在 GUI 層級處理，此處不再重複設定

class HPMonitor:
    """體力監控工具類"""
    
    def __init__(self):
        """初始化 HPMonitor"""
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.1
        self._setup_tesseract()
        self._log_callback = print
    
    def _setup_tesseract(self):
        """自動檢測並設定 Tesseract OCR 路徑"""
        common_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            r'C:\Tesseract-OCR\tesseract.exe',
        ]
        
        # 檢查是否在 PATH 中
        import shutil
        tesseract_path = shutil.which('tesseract')
        if tesseract_path and os.path.exists(tesseract_path):
            return
        
        # 嘗試從常見路徑中找到
        for path in common_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                return

    def _find_window_by_keyword(self, keyword: str):
        """根據關鍵字搜尋視窗"""
        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd).strip()
                if not title:
                    return True
                    
                # 排除程式本身的視窗、終端機與常見系統視窗
                blacklist = ['天堂簡易補血器', 'cmd.exe', 'powershell', 'python', 'text_clicker']
                for b in blacklist:
                    if b.lower() in title.lower():
                        return True # 跳過此視窗
                        
                if keyword.lower() in title.lower():
                    ctx.append(hwnd)
            return True
        
        windows = []
        try:
            win32gui.EnumWindows(enum_handler, windows)
            return windows[-1] if windows else 0 # 常常遊戲視窗在 Z-order 較後面
        except:
            return 0

    def capture_window_region(self, window_title: Optional[str] = None, region: Optional[Tuple[int, int, int, int]] = None) -> Image.Image:
        """擷取指定視窗或區域的截圖"""
        if window_title:
            hwnd = win32gui.FindWindow(None, window_title) or self._find_window_by_keyword(window_title)
            if hwnd == 0:
                raise ValueError(f"找不到視窗: {window_title}")
            
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            if region:
                rx, ry, rw, rh = region
                return pyautogui.screenshot(region=(left + rx, top + ry, rw, rh))
            else:
                return pyautogui.screenshot(region=(left, top, right - left, bottom - top))
        else:
            return pyautogui.screenshot(region=region) if region else pyautogui.screenshot()

    def detect_hp(self, window_title: Optional[str] = None, region: Optional[Tuple[int, int, int, int]] = None) -> Tuple[int, int, float]:
        """偵測體力數值並計算百分比"""
        try:
            screenshot = self.capture_window_region(window_title, region)
            processed_img = self._preprocess_hp_image(screenshot)
            
            # 存偵錯圖片 (原始擷取 + 處理後)
            debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'hpbar')
            screenshot.save(os.path.join(debug_dir, 'debug_capture.png'))
            processed_img.save(os.path.join(debug_dir, 'debug_processed.png'))
            
            # OCR 識別
            # --psm 11: 稀疏文字模式 (恢復使用 PSM 11，因 PSM 7 會漏掉數字)
            # -c tessedit_char_whitelist: 只偵測數字、斜線與 HP: 符號
            config = '--psm 11 --oem 3 -c tessedit_char_whitelist=0123456789/HP: '
            hp_text = pytesseract.image_to_string(processed_img, config=config).strip()
            
            # [重要] 針對 Lineage 特殊字體的後處理轉換
            trans_table = str.maketrans({
                'I': '1', 'l': '1', '|': '1',  # 1 的誤讀
                'G': '6', 'b': '6',             # 6 的誤讀
                'T': '7', 't': '7',             # 7 的誤讀
                'S': '5', 's': '5',             # 5 的誤讀
                'B': '8',                       # 8 的誤讀
                'O': '0', 'o': '0',             # 0 的誤讀
                'A': '4',                       # 4 的誤讀
            })
            hp_text = hp_text.translate(trans_table)
            
            self._log_callback(f"[體力偵測] OCR 修復後結果: '{hp_text}'")
            # 確保提取出來的數字符合使用者的「最多 3 位數」鐵律
            def validate(c, m):
                if m > 999:
                    self._log_callback(f"[體力偵測] 忽略異常大數值: {c} / {m}")
                    return 0, 0, 1.0
                if c > 999 or c > m:
                    # 嘗試用總血量來修正目前血量
                    if hasattr(self, '_last_max_hp') and self._last_max_hp > 0:
                        c_str = str(c)
                        m_len = len(str(self._last_max_hp))
                        # 取前 m_len 位數作為修正後的目前血量
                        if len(c_str) > m_len:
                            corrected = int(c_str[:m_len])
                            if corrected <= self._last_max_hp:
                                self._log_callback(f"[體力修正] {c} -> {corrected} / {self._last_max_hp}")
                                return corrected, self._last_max_hp, (corrected / self._last_max_hp)
                    self._log_callback(f"[體力偵測] 忽略不可能的比例: {c} / {m}")
                    return 0, 0, 1.0
                return c, m, (c / m if m > 0 else 1.0)

            # 使用更寬鬆的 regex (優先處理完美的 '108 / 167')
            match = re.search(r'(\d+)\s*/\s*(\d+)', hp_text)
            if match:
                curr, m_hp = int(match.group(1)), int(match.group(2))
                if m_hp > 0:
                    self._last_max_hp = m_hp
                    return validate(curr, m_hp)
            
            # 記憶修復法：如果 OCR 把斜線看錯變成 '1007167' 這種黏在一起的字串
            if hasattr(self, '_last_max_hp') and self._last_max_hp > 0:
                m_hp_str = str(self._last_max_hp)
                last_idx = hp_text.rfind(m_hp_str)
                if last_idx > 0:
                    left_part = hp_text[:last_idx]
                    left_nums = re.findall(r'\d+', left_part)
                    if left_nums:
                        left_str = "".join(left_nums)
                        if len(left_str) >= 2 and left_str[-1] in ('7', '1'):
                            left_str = left_str[:-1]
                        
                        if left_str.isdigit():
                            curr = int(left_str)
                            return validate(curr, self._last_max_hp)
            
            # 最底層的暴力拆解 (當完全沒有記憶可以用的時候)
            nums = re.findall(r'\d+', hp_text)
            if len(nums) == 1:
                s_num = str(nums[0])
                if len(s_num) % 2 == 0 and len(s_num) >= 2:
                    half = len(s_num) // 2
                    curr, m_hp = int(s_num[:half]), int(s_num[half:])
                    if m_hp > 0:
                        self._last_max_hp = m_hp
                        return validate(curr, m_hp)
                        
            elif len(nums) >= 2:
                curr, m_hp = int(nums[-2]), int(nums[-1])
                if m_hp > 0:
                    self._last_max_hp = m_hp
                    return validate(curr, m_hp)
                
            return 0, 0, 1.0
        except Exception as e:
            self._log_callback(f"[體力偵測] 錯誤: {e}")
            return 0, 0, 1.0


    def find_hp_roi(self, window_title: str, template_path: str = None) -> Optional[Tuple[int, int, int, int]]:
        """使用 OCR 搜尋畫面中的 'HP:' 文字，自動定位體力數字區域"""
        try:
            self._log_callback("[自動偵測] 正在擷取視窗畫面...")
            screenshot = self.capture_window_region(window_title)
            img_array = np.array(screenshot)
            img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            
            h, w = gray.shape
            # 放大 2 倍以提高 OCR 準確度
            gray_scaled = cv2.resize(gray, (w * 2, h * 2), interpolation=cv2.INTER_CUBIC)
            
            self._log_callback("[自動偵測] 正在掃描畫面中的體力 (HP:) 文字...")
            
            # 使用 Tesseract 取得每個文字的位置資訊
            pil_img = Image.fromarray(gray_scaled)
            data = pytesseract.image_to_data(pil_img, config='--psm 11 --oem 3', output_type=pytesseract.Output.DICT)
            
            # 搜尋 "HP:" 文字
            for i, text in enumerate(data['text']):
                text_clean = text.strip().upper()
                if 'HP:' in text_clean or 'HP;' in text_clean or 'HP' == text_clean:
                    conf = float(data['conf'][i]) if data['conf'][i] != '-1' else 0
                    if conf < 10:
                        continue
                    
                    # 取得 HP 文字的位置 (座標除以 2 還原回原始尺寸)
                    x = data['left'][i] // 2
                    y = data['top'][i] // 2
                    text_w = data['width'][i] // 2
                    text_h = data['height'][i] // 2
                    
                    # 從 HP: 文字開始，往右延伸以包含數字
                    roi_x = x
                    roi_y = max(0, y - 5)
                    roi_w = min(text_w + 200, w - roi_x)
                    roi_h = text_h + 10
                    
                    # 產生偵錯截圖：在原圖上畫黃色矩形與圓點
                    debug_img = img_cv.copy()
                    # 黃色矩形框出偵測區域
                    cv2.rectangle(debug_img, (roi_x, roi_y), (roi_x + roi_w, roi_y + roi_h), (0, 255, 255), 2)
                    # 黃色大圓點標記 HP 文字起始位置
                    cv2.circle(debug_img, (x + text_w // 2, y + text_h // 2), 8, (0, 255, 255), -1)
                    # 加上文字標記
                    cv2.putText(debug_img, 'HP Detected', (roi_x, roi_y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)
                    
                    debug_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'hpbar', 'debug_detect.png')
                    cv2.imwrite(debug_path, debug_img)
                    
                    self._log_callback(f"[自動偵測] 成功！找到體力位置 (信心度: {conf:.0f}%, 位置: {roi_x}, {roi_y})")
                    self._log_callback(f"[自動偵測] 偵錯截圖已存至: {debug_path}")
                    return (roi_x, roi_y, roi_w, roi_h)
            
            # 若找不到，也存一張截圖供偵錯
            debug_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'hpbar', 'debug_detect.png')
            cv2.imwrite(debug_path, img_cv)
            self._log_callback(f"[自動偵測] 失敗: 畫面中找不到體力 (HP:) 文字")
            self._log_callback(f"[自動偵測] 原始截圖已存至: {debug_path}")
            return None
        except Exception as e:
            self._log_callback(f"[自動偵測] 發生錯誤: {e}")
            import traceback
            self._log_callback(traceback.format_exc())
            return None

    def _preprocess_hp_image(self, image: Image.Image) -> Image.Image:
        """針對遊戲體力數字優化的圖像處理"""
        img_array = np.array(image)
        img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        
        # 放大 3 倍以提高 OCR 準確度 (Lineage 字體較細)
        h, w = gray.shape
        gray = cv2.resize(gray, (w * 3, h * 3), interpolation=cv2.INTER_CUBIC)
        
        # 使用大津二值化 (Otsu's Binarization) 自動尋找最佳門檻
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        
        # 移除 Dilation (膨脹)，避免細體字體數字糊掉
        # 僅保留清晰的二值化影像即可
        
        return Image.fromarray(thresh)

    def check_and_trigger(self, window_title: Optional[str] = None, region: Optional[Tuple[int, int, int, int]] = None, 
                           threshold: float = 0.7, key: str = 'f5', presses: int = 2, interval: float = 0.1) -> bool:
        """核心監控邏輯"""
        curr, m_hp, ratio = self.detect_hp(window_title, region)
        if m_hp > 0:
            self._log_callback(f"[體力狀態] {curr} / {m_hp} ({ratio:.1%})")
            if ratio < threshold:
                self._log_callback(f"[觸發警告] 體力低於 {threshold:.0%}，正在按 {key.upper()} {presses} 下")
                pyautogui.press(key, presses=presses, interval=interval)
                return True
        return False
