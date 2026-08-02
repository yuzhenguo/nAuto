import os
import cv2
import numpy as np
import re

# ── EasyOCR 초기화 ──────────────────────────────────────
easyocr_available = False
try:
    import easyocr
    easyocr_available = True
    print("[EasyOCR] Library loaded successfully.")
except ImportError:
    print("[EasyOCR] Library not installed. (pip install easyocr is required)")

# ── ddddocr 초기화 ──────────────────────────────────────
ddddocr_available = False
try:
    import ddddocr
    ddddocr_available = True
    print("[ddddocr] Library loaded successfully.")
except ImportError:
    print("[ddddocr] Library not installed. (pip install ddddocr is required)")


def test_easyocr(img_path, zipcode):
    if not easyocr_available:
        print("\n[EasyOCR] Test skipped: easyocr library not available.")
        return
        
    print("\n=== 1. EasyOCR Test Start ===")
    
    # 이미지 로드 (Unicode 경로 안전하게 로드)
    img_arr = np.fromfile(img_path, dtype=np.uint8)
    img = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
    if img is None:
        print(f"Error: Unable to load image: {img_path}")
        return
        
    # 1회성 reader 초기화
    reader = easyocr.Reader(['ko', 'en'], gpu=False)
    # 1. Preprocessing: 2x upscale and sharpening
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    scale = 2.0
    enlarged = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
    sharpened = cv2.filter2D(enlarged, -1, kernel)
    
    print("  Running EasyOCR...")
    results = reader.readtext(sharpened, detail=1, paragraph=False)
    
    target_clean = zipcode.strip()
    found_box = None
    found_text = None
    best_conf = 0
    
    print("  [All Recognitions]")
    for (bbox, text, conf) in results:
        # Filter text output for safe printing
        safe_text = "".join(c if ord(c) < 128 else f"\\u{ord(c):04x}" for c in text)
        text_clean = re.sub(r'\s+', '', safe_text)
        print(f"    - '{safe_text}' (Cleaned: '{text_clean}'), Conf: {conf:.2f}")
        
        # Check matching
        raw_text_clean = re.sub(r'\s+', '', text)
        if raw_text_clean == target_clean and conf > best_conf:
            found_box = bbox
            found_text = text
            best_conf = conf
        elif target_clean in raw_text_clean and conf > best_conf:
            found_box = bbox
            found_text = text
            best_conf = conf
            
    if found_box is not None:
        xs = [pt[0] for pt in found_box]
        ys = [pt[1] for pt in found_box]
        cx = int((min(xs) + max(xs)) / 2 / scale)
        cy = int((min(ys) + max(ys)) / 2 / scale)
        safe_found = "".join(c if ord(c) < 128 else f"\\u{ord(c):04x}" for c in found_text)
        print(f"  [SUCCESS] Found '{safe_found}'! (Conf: {best_conf:.2f})")
        print(f"  Target Click Center Coordinates: ({cx}, {cy})")
        
        debug_img = img.copy()
        pts = np.array([[int(pt[0]/scale), int(pt[1]/scale)] for pt in found_box], dtype=np.int32)
        cv2.polylines(debug_img, [pts], True, (0, 255, 0), 3)
        cv2.circle(debug_img, (cx, cy), 10, (0, 0, 255), -1)
        
        out_path = "test_easyocr_result.png"
        _, enc_dbg = cv2.imencode('.png', debug_img)
        enc_dbg.tofile(out_path)
        print(f"  Debug image saved to: {out_path}")
    else:
        print(f"  [FAILED] EasyOCR could not find zipcode '{zipcode}'")


def test_ddddocr_slice(img_path, zipcode):
    if not ddddocr_available:
        print("\n[ddddocr] Test skipped: ddddocr library not available.")
        return
        
    print("\n=== 2. ddddocr Slice Scan Test Start ===")
    
    img_arr = np.fromfile(img_path, dtype=np.uint8)
    img = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
    if img is None:
        print(f"Error: Unable to load image: {img_path}")
        return
        
    ocr = ddddocr.DdddOcr(show_ad=False)
    h, w = img.shape[:2]
    
    found = False
    
    # Slice vertically from y=450 to y=1800 with 20px step
    # Restrict width to x=[0, 260] for the zipcode column on the left side
    for y in range(450, 1800, 20):
        chunk = img[y:y+100, 0:260]
        _, chunk_png = cv2.imencode('.png', chunk)
        try:
            text = ocr.classification(chunk_png.tobytes())
        except Exception:
            continue
            
        clean_text = "".join(filter(str.isdigit, text))
        
        if zipcode in clean_text:
            tap_y = y + 50
            safe_text = "".join(c if ord(c) < 128 else f"\\u{ord(c):04x}" for c in text)
            print(f"  [SUCCESS] Found zipcode '{zipcode}' around y={y}! (OCR text: '{safe_text}')")
            print(f"  Target Click Y Coordinate: {tap_y}")
            
            debug_img = img.copy()
            cv2.rectangle(debug_img, (0, y), (260, y+100), (0, 255, 0), 2)
            cv2.line(debug_img, (0, tap_y), (w, tap_y), (0, 0, 255), 2)
            
            out_path = "test_ddddocr_result.png"
            _, enc_dbg = cv2.imencode('.png', debug_img)
            enc_dbg.tofile(out_path)
            print(f"  Debug image saved to: {out_path}")
            
            found = True
            break
            
    if not found:
        print(f"  [FAILED] ddddocr slice scan could not find zipcode '{zipcode}'")


if __name__ == "__main__":
    target_img = "debug_zipcode_screenshot.png"
    target_zip = "52157"
    
    if not os.path.exists(target_img):
        print(f"Error: '{target_img}' file not found.")
    else:
        print(f"Checking zipcode '{target_zip}' in image '{target_img}'")
        test_easyocr(target_img, target_zip)
        test_ddddocr_slice(target_img, target_zip)
