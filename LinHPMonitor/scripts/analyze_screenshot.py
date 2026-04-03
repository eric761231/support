import cv2
import numpy as np
import os
from pathlib import Path

def find_window_bbox(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150)
    # dilate to close gaps
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5,5))
    dil = cv2.dilate(edges, kernel, iterations=2)
    contours, _ = cv2.findContours(dil, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    h,w = img.shape[:2]
    best = None
    best_area = 0
    for cnt in contours:
        x,y,ww,hh = cv2.boundingRect(cnt)
        area = ww*hh
        # ignore tiny contours
        if area < 10000: continue
        # prefer large rectangles not equal to full image
        if area > best_area:
            best_area = area
            best = (x,y,ww,hh)
    if best is None:
        return (0,0,w,h)
    return best

def find_bar_by_color(img, is_mp=False):
    # search near bottom half
    h,w = img.shape[:2]
    roi = img[int(h*0.55):int(h*0.9), int(w*0.15):int(w*0.85)]
    hsv = cv2.cvtColor(roi, cv2.COLOR_BGR2HSV)
    if is_mp:
        lower = np.array([95,40,40])
        upper = np.array([140,255,255])
    else:
        # red range: combine two
        lower1 = np.array([0,80,40]); upper1 = np.array([10,255,255])
        lower2 = np.array([160,80,40]); upper2 = np.array([180,255,255])
        mask1 = cv2.inRange(hsv, lower1, upper1)
        mask2 = cv2.inRange(hsv, lower2, upper2)
        mask = cv2.bitwise_or(mask1, mask2)
        # find contours
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            return None
        best = max(contours, key=cv2.contourArea)
        x,y,ww,hh = cv2.boundingRect(best)
        # convert to full image coords
        return (x + int(w*0.15), y + int(h*0.55), ww, hh)
    mask = cv2.inRange(hsv, lower, upper)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None
    best = max(contours, key=cv2.contourArea)
    x,y,ww,hh = cv2.boundingRect(best)
    return (x + int(w*0.15), y + int(h*0.55), ww, hh)


def compute_normalized_from_window(win, bar):
    wx,wy,ww,wh = win
    bx,by,bw,bh = bar
    # center of bar relative to window
    center_x = (bx - wx) + bw/2
    center_y = (by - wy) + bh/2
    return (center_x / ww, center_y / wh, bw / ww, bh / wh)

if __name__ == '__main__':
    base = Path(__file__).resolve().parents[1]
    samples = base / 'smaples'
    p = samples / 'purple.jpg'
    if not p.exists():
        p = samples / 'loacl.jpg'
    if not p.exists():
        print('No sample image found in', samples)
        exit(1)
    img = cv2.imread(str(p))
    win = find_window_bbox(img)
    print('window bbox:', win)
    hp = find_bar_by_color(img, is_mp=False)
    mp = find_bar_by_color(img, is_mp=True)
    print('hp bbox:', hp)
    print('mp bbox:', mp)
    if hp:
        hp_norm = compute_normalized_from_window(win, hp)
        print('hp norm (cx,cy,w,h ratios):', [round(x,4) for x in hp_norm])
    if mp:
        mp_norm = compute_normalized_from_window(win, mp)
        print('mp norm (cx,cy,w,h ratios):', [round(x,4) for x in mp_norm])
    # save debug image with boxes
    dbg = img.copy()
    x,y,ww,hh = win
    cv2.rectangle(dbg,(x,y),(x+ww,y+hh),(0,255,0),2)
    if hp:
        x,y,ww,hh = hp
        cv2.rectangle(dbg,(x,y),(x+ww,y+hh),(0,0,255),2)
    if mp:
        x,y,ww,hh = mp
        cv2.rectangle(dbg,(x,y),(x+ww,y+hh),(255,0,0),2)
    outp = samples / 'analysis_debug.jpg'
    cv2.imwrite(str(outp), dbg)
    print('wrote debug image to', outp)
