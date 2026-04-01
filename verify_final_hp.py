from text_clicker import HPMonitor
from PIL import Image
import os
import pytesseract
import re

def test_real_screenshot():
    monitor = HPMonitor()
    image_path = r"c:\c\code_workspace\support\hpbar\hp.jpg"
    
    if not os.path.exists(image_path):
        print(f"Error: {image_path} not found.")
        return

    print(f"Testing OCR on: {image_path}")
    img = Image.open(image_path)
    monitor._log_callback = print
    
    # Use the actual engine method
    processed = monitor._preprocess_hp_image(img)
    processed.save("processed_hp_test.png")
    
    # Test OCR with PSM 11
    config = '--psm 11 --oem 3'
    hp_text = pytesseract.image_to_string(processed, config=config).strip()
    print(f"OCR Output: '{hp_text}'")
    
    match = re.search(r'(\d+)\s*/\s*(\d+)', hp_text)
    if match:
        curr, m_hp = int(match.group(1)), int(match.group(2))
        print(f"Parsed Result: Current={curr}, Max={m_hp}, Ratio={curr/m_hp:.1%}")
    else:
        # Fallback to search for any two numbers
        nums = re.findall(r'\d+', hp_text)
        if len(nums) >= 2:
            curr, m_hp = int(nums[0]), int(nums[1])
            print(f"Parsed Result (Fallback): Current={curr}, Max={m_hp}, Ratio={curr/m_hp:.1%}")
        else:
            print("Regex failed to find numbers.")

if __name__ == "__main__":
    test_real_screenshot()
