import cv2
import numpy as np
import pytesseract
from PIL import Image
import os

def brute_force_ocr():
    image_path = r"c:\c\code_workspace\support\hpbar\hp.jpg"
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Try different scales
    for scale in [1, 2, 3]:
        resized = cv2.resize(gray, (0, 0), fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
        
        # Try different thresholding
        # 1. Simple Threshold
        for thresh_val in [150, 180, 200]:
            _, thresh = cv2.threshold(resized, thresh_val, 255, cv2.THRESH_BINARY)
            # Tesseract likes black text on white
            inv = cv2.bitwise_not(thresh)
            
            for psm in [6, 7, 11]:
                config = f'--psm {psm} --oem 3'
                text = pytesseract.image_to_string(inv, config=config).strip()
                if text:
                    print(f"Scale={scale}, Thresh={thresh_val}, PSM={psm} => '{text}'")

        # 2. Otsu
        _, otsu = cv2.threshold(resized, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        inv_otsu = cv2.bitwise_not(otsu)
        for psm in [6, 7, 11]:
            config = f'--psm {psm} --oem 3'
            text = pytesseract.image_to_string(inv_otsu, config=config).strip()
            if text:
                print(f"Scale={scale}, Otsu, PSM={psm} => '{text}'")

if __name__ == "__main__":
    brute_force_ocr()
