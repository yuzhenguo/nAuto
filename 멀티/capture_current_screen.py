import os
import sys
import time
import subprocess
import cv2
import numpy as np

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def capture_and_crop():
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
    print(f"📱 스크린샷 캡처 중... 기기: {device_id}")
    
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

    save_path = os.path.join(_BASE_DIR, "captured_screen.png")
    cv2.imwrite(save_path, screen)
    print(f"✅ 현재 전체 화면이 저장되었습니다: {save_path}")
    print("💡 이제 캡처된 화면을 기준으로 매칭 테스트를 진행합니다.")

if __name__ == "__main__":
    capture_and_crop()
    os.system("pause")
