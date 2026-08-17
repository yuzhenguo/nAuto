import os
import sys
import time
import subprocess
import cv2
import numpy as np

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def test_match_without_scroll():
    # 1. 기기 확인
    res = subprocess.run(["adb", "devices"], capture_output=True, text=True)
    devices = []
    for line in res.stdout.strip().splitlines()[1:]:
        parts = line.split()
        if len(parts) >= 2 and parts[1] == "device":
            devices.append(parts[0])
            
    if not devices:
        print("❌ 연결된 ADB 기기가 없습니다.")
        os.system("pause")
        return
        
    device_id = devices[0]
    print(f"📱 기기 선택됨: {device_id}")
    print("📸 현재 핸드폰 화면 캡처 중...")
    
    # 2. 캡처
    cmd = ["adb", "-s", device_id, "exec-out", "screencap", "-p"]
    res = subprocess.run(cmd, capture_output=True, timeout=10)
    if not res.stdout or len(res.stdout) < 200:
        print("❌ 스크린샷 캡처 실패")
        os.system("pause")
        return

    img_arr = np.frombuffer(res.stdout, np.uint8)
    screen = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
    if screen is None:
        print("❌ 이미지 디코딩 실패")
        os.system("pause")
        return

    # 3. 템플릿 매칭 테스트 (다른결재 관련 이미지들)
    templates = [
        "다른결재.png",
        "다른결재2.png",
        "다른결제.png",
        "다른결제수단.png"
    ]
    
    best_match = None
    best_score = 0.0
    best_name = ""
    best_loc = None
    best_tw, best_th = 0, 0

    for t_name in templates:
        t_path = os.path.join(_BASE_DIR, t_name)
        if not os.path.exists(t_path):
            continue
            
        tmpl = cv2.imdecode(np.fromfile(t_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        if tmpl is None:
            tmpl = cv2.imread(t_path, cv2.IMREAD_COLOR)
        if tmpl is None:
            continue
            
        th, tw = tmpl.shape[:2]
        
        # 멀티 스케일 매칭
        for scale in [1.0, 0.9, 1.1, 0.85, 0.95, 1.05, 1.15]:
            new_w = int(tw * scale)
            new_h = int(th * scale)
            if new_w < 10 or new_h < 10 or new_w > screen.shape[1] or new_h > screen.shape[0]:
                continue
            
            resized_tmpl = cv2.resize(tmpl, (new_w, new_h), interpolation=cv2.INTER_AREA if scale < 1.0 else cv2.INTER_CUBIC)
            res_map = cv2.matchTemplate(screen, resized_tmpl, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res_map)
            
            if max_val > best_score:
                best_score = max_val
                best_name = t_name
                best_loc = max_loc
                best_tw, best_th = new_w, new_h

    print("\n=======================================================")
    if best_score > 0 and best_loc is not None:
        cx = best_loc[0] + best_tw // 2
        cy = best_loc[1] + best_th // 2
        print(f"🎯 [매칭 결과] 이미지: {best_name}")
        print(f"   - 점수 (신뢰도): {best_score:.4f}")
        print(f"   - 중심 좌표: ({cx}, {cy})")
        
        # 인식된 영역 빨간색 상자로 표시 후 저장
        matched_img = screen.copy()
        cv2.rectangle(matched_img, best_loc, (best_loc[0] + best_tw, best_loc[1] + best_th), (0, 0, 255), 3)
        cv2.circle(matched_img, (cx, cy), 10, (0, 255, 0), -1)
        
        save_path = os.path.join(_BASE_DIR, "matched_result.png")
        cv2.imwrite(save_path, matched_img)
        print(f"📸 인식 구역 표시 이미지 저장 완료: {save_path}")
    else:
        print("❌ 매칭 가능한 다른결재 이미지 파일을 찾지 못했거나 매칭에 실패했습니다.")
    print("=======================================================\n")

if __name__ == "__main__":
    test_match_without_scroll()
    os.system("pause")
