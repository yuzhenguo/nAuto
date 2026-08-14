"""
appium_helper.py
Appium 공통 유틸리티 모듈 (네이버 앱 전용)
- 드라이버 생성
- 요소 대기/클릭/입력 래퍼
"""

import time
from appium import webdriver
from appium.options.android import UiAutomator2Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException
)

# ─── 앱 설정 ───────────────────────────────────────────────────────────────────
NAVER_PACKAGE  = "com.nhn.android.search"
NAVER_ACTIVITY = "com.nhn.android.search.ui.pages.SearchHomePage"  # adb 확인된 실제 런처 Activity

# 로드 확인용 (쿠팡 appium_helper가 로드되면 이 메시지가 안 나옴)
print(f"[✅ naver_appium 로드됨] PACKAGE={NAVER_PACKAGE}, FILE={__file__}")


def _kill_port_process(port: int):
    """지정한 포트를 사용하는 프로세스 종료 (Windows 전용)"""
    import subprocess
    try:
        result = subprocess.run(
            ["cmd", "/c", f"netstat -ano | findstr :{port}"],
            capture_output=True, text=True, timeout=3
        )
        lines = result.stdout.strip().splitlines()
        for line in lines:
            parts = line.split()
            if len(parts) >= 5:
                local_addr = parts[1]
                addr_parts = local_addr.rsplit(':', 1)
                if len(addr_parts) == 2 and addr_parts[1] == str(port):
                    pid = parts[-1]
                    subprocess.run(["taskkill", "/F", "/PID", pid],
                                   capture_output=True, timeout=3)
    except Exception:
        pass


def wake_screen_if_off(device_id: str, log_callback=None) -> bool:
    """
    화면이 꺼져 있으면 adb로 깨운 뒤 True 반환.
    이미 켜져 있으면 아무것도 안 하고 False 반환.
    """
    import subprocess
    try:
        # 화면 상태 확인: mHoldingWakeLockSuspendBlocker=true 이면 켜져 있음
        res = subprocess.run(
            ["adb", "-s", device_id, "shell", "dumpsys", "power"],
            capture_output=True, text=True, timeout=5
        )
        output = res.stdout
        screen_on = (
            "mHoldingInteractiveSuspendBlocker=true" in output
            or "Display Power: state=ON" in output
            or "mWakefulness=Awake" in output
        )
        if not screen_on:
            _log(log_callback, f"  📺 [{device_id}] 화면 꺼짐 감지 → WAKEUP 키 전송")
            # KEYCODE_WAKEUP (224)으로 화면 켜기
            subprocess.run(
                ["adb", "-s", device_id, "shell", "input", "keyevent", "224"],
                capture_output=True, timeout=5
            )
            time.sleep(1)  # 화면 켜지는 대기 시간
            _log(log_callback, f"  📺 [{device_id}] 화면 잠금 해제를 위한 스와이프 전송 (500, 1600 -> 500, 600)")
            # 화면 아래에서 위로 쓸어넘겨서 드래그 잠금 해제
            subprocess.run(
                ["adb", "-s", device_id, "shell", "input", "swipe", "500", "1600", "500", "600", "300"],
                capture_output=True, timeout=5
            )
            time.sleep(2)  # 화면 전환/안정화 대기
            return True
    except Exception as e:
        _log(log_callback, f"  ⚠ [{device_id}] 화면 상태 확인 중 오류 (무시): {e}")
    return False


def get_screenshot_safe(driver, device_id: str, log_callback=None,
                        max_retries: int = 3, black_threshold: int = 10) -> bytes:
    """
    화면이 꺼져 있어 캡처가 검게 나오는 경우를 자동 복구하는 스크린샷 함수.

    1. 화면 상태 확인 → 꺼져 있으면 wake 후 재캡처
    2. 캡처된 이미지가 거의 완전히 검으면(평균 밝기 < black_threshold) 재시도

    Returns:
        PNG bytes (bytes)
    """
    import numpy as np

    for attempt in range(1, max_retries + 1):
        # 스크린샷 전 화면 상태 보장
        woke = wake_screen_if_off(device_id, log_callback)
        if woke:
            _log(log_callback, f"  📺 [{device_id}] 화면을 깨운 뒤 스크린샷 재시도 ({attempt}/{max_retries})")

        png_bytes = driver.get_screenshot_as_png()

        # 검은 화면 여부 판별 (numpy로 평균 밝기 계산)
        try:
            arr = np.frombuffer(png_bytes, np.uint8)
            import cv2
            img = cv2.imdecode(arr, cv2.IMREAD_GRAYSCALE)
            if img is not None:
                mean_brightness = float(img.mean())
                if mean_brightness < black_threshold:
                    _log(log_callback,
                         f"  ⚠ [{device_id}] 캡처 이미지가 검음 (평균 밝기={mean_brightness:.1f}) "
                         f"→ 화면 깨우기 후 재시도 ({attempt}/{max_retries})")
                    # 강제 화면 켜기
                    import subprocess
                    subprocess.run(
                        ["adb", "-s", device_id, "shell", "input", "keyevent", "224"],
                        capture_output=True, timeout=5
                    )
                    time.sleep(1)
                    # 여기도 스와이프 잠금해제 추가
                    subprocess.run(
                        ["adb", "-s", device_id, "shell", "input", "swipe", "500", "1600", "500", "600", "300"],
                        capture_output=True, timeout=5
                    )
                    time.sleep(2)
                    continue  # 다음 시도
        except Exception:
            pass  # 밝기 판별 실패 시 그냥 사용

        return png_bytes  # 정상 캡처

    # 모든 재시도 실패 시 마지막 캡처 반환
    _log(log_callback, f"  ⚠ [{device_id}] 스크린샷 {max_retries}회 모두 검은 화면 → 마지막 캡처 사용")
    return driver.get_screenshot_as_png()


def set_default_browser_to_chrome(device_id: str, log_callback=None):
    """지정한 기기의 기본 브라우저를 크롬(com.android.chrome)으로 설정"""
    import subprocess
    try:
        # 먼저 Chrome 패키지가 설치되어 있는지 확인
        check_cmd = ["adb", "-s", device_id, "shell", "pm", "path", "com.android.chrome"]
        check_res = subprocess.run(check_cmd, capture_output=True, text=True, timeout=5)
        if "package:" not in check_res.stdout:
            _log(log_callback, f"  ℹ️ [{device_id}] Chrome 브라우저가 설치되어 있지 않습니다. 기본 브라우저 설정을 변경하지 않습니다.")
            return

        cmd = ["adb", "-s", device_id, "shell", "role", "set-role-holder", "android.app.role.BROWSER", "com.android.chrome"]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
        if res.returncode == 0:
            _log(log_callback, f"  ℹ️ [{device_id}] 기본 브라우저를 Chrome으로 변경 완료 (삼성 인터넷 SurfaceView 이슈 우회)")
        else:
            stderr_str = res.stderr.strip() if res.stderr else ""
            if "not found" in stderr_str or "inaccessible" in stderr_str or res.returncode == 127:
                _log(log_callback, f"  ℹ️ [{device_id}] 기기에서 'role' 명령이 지원되지 않아 기본 브라우저 설정 변경을 건너뜁니다 (Android 버전 등)")
            else:
                _log(log_callback, f"  ⚠️ [{device_id}] 기본 브라우저 Chrome 설정 실패 (ExitCode {res.returncode}): {stderr_str}")
    except Exception as e:
        _log(log_callback, f"  ⚠️ [{device_id}] 기본 브라우저 설정 중 예외 발생: {e}")


def reset_device_uiautomation(device_id: str, log_callback=None):
    """기기의 UiAutomation 상태 및 Appium 서버 패키지를 완전히 리셋 (자가 치유)"""
    import subprocess
    _log(log_callback, f"  🔄 [{device_id}] 기기 UiAutomation 상태 및 Appium 패키지 리셋 시도...")
    try:
        # 1. accessibility_enabled 토글로 AccessibilityManager 초기화
        subprocess.run(["adb", "-s", device_id, "shell", "settings", "put", "secure", "accessibility_enabled", "0"], capture_output=True, timeout=5)
        time.sleep(1.5)
        subprocess.run(["adb", "-s", device_id, "shell", "settings", "put", "secure", "accessibility_enabled", "1"], capture_output=True, timeout=5)
        time.sleep(1.5)
        
        # 2. uiautomator2 server 패키지 삭제 (재설치 유도)
        subprocess.run(["adb", "-s", device_id, "shell", "pm", "uninstall", "io.appium.uiautomator2.server"], capture_output=True, timeout=5)
        subprocess.run(["adb", "-s", device_id, "shell", "pm", "uninstall", "io.appium.uiautomator2.server.test"], capture_output=True, timeout=5)
        
        # 3. uiautomator2 server 프로세스 강제 종료
        subprocess.run(["adb", "-s", device_id, "shell", "am", "force-stop", "io.appium.uiautomator2.server"], capture_output=True, timeout=5)
        subprocess.run(["adb", "-s", device_id, "shell", "am", "force-stop", "io.appium.uiautomator2.server.test"], capture_output=True, timeout=5)
        subprocess.run(["adb", "-s", device_id, "shell", "pkill", "-f", "uiautomator"], capture_output=True, timeout=5)
    except Exception as e:
        _log(log_callback, f"  ⚠ [{device_id}] 리셋 중 예외 발생 (무시): {e}")


def reboot_device_and_wait(device_id: str, log_callback=None) -> bool:
    """기기(핸드폰) 강제 재부팅 후 부팅이 완료될 때까지 대기"""
    import subprocess
    _log(log_callback, f"  🚨 [{device_id}] 'UiAutomation not connected' 에러 감지 → 설비(핸드폰) 강제 재부팅 진행")
    try:
        subprocess.run(["adb", "-s", device_id, "reboot"], capture_output=True, timeout=10)
        _log(log_callback, f"  ⏳ [{device_id}] 재부팅 명령 전송 완료. 재부팅 및 ADB 연결 대기 (최대 150초)...")
        
        # 기기가 꺼지고 완전히 켜질 때까지 루프 돌며 체크
        max_wait = 150
        check_interval = 5
        start_time = time.time()
        
        # 기기가 일단 offline으로 꺼질 때까지 잠시 대기
        time.sleep(10)
        
        while time.time() - start_time < max_wait:
            # 1. 기기가 adb devices 목록에 'device' 상태로 나타나는지 확인
            res_devices = subprocess.run(["adb", "devices"], capture_output=True, text=True, timeout=5)
            lines = res_devices.stdout.strip().splitlines()
            online = False
            for line in lines[1:]:
                parts = line.split()
                if len(parts) >= 2 and parts[0] == device_id and parts[1] == "device":
                    online = True
                    break
            
            if online:
                # 2. boot_completed 프로퍼티가 '1'인지 확인
                res_boot = subprocess.run(
                    ["adb", "-s", device_id, "shell", "getprop", "sys.boot_completed"],
                    capture_output=True, text=True, timeout=5
                )
                if res_boot.stdout.strip() == "1":
                    _log(log_callback, f"  ✅ [{device_id}] 기기 재부팅 완료 (부팅 완료 감지됨)!")
                    time.sleep(5)  # 부팅 완료 후 시스템 안정화 대기
                    return True
            
            elapsed = int(time.time() - start_time)
            _log(log_callback, f"  ⏳ [{device_id}] 부팅 대기 중... ({elapsed}초 경과)")
            time.sleep(check_interval)
            
        _log(log_callback, f"  ⚠️ [{device_id}] 150초 대기 시간 초과. 부팅 확인 없이 연결 시도 진행")
    except Exception as e:
        _log(log_callback, f"  ⚠ [{device_id}] 재부팅 및 대기 중 예외 발생: {e}")
    return False


def create_driver(device_id: str, appium_port: int, log_callback=None) -> webdriver.Remote:
    """
    지정 기기와 Appium 포트로 드라이버 생성

    Args:
        device_id: adb 기기 ID
        appium_port: 해당 기기의 Appium 서버 포트
        log_callback: GUI 로그 함수 (선택)

    Returns:
        Appium WebDriver 인스턴스
    """
    import subprocess

    # 기본 브라우저 설정 (C:\appium_auto 로직 참고)
    set_default_browser_to_chrome(device_id, log_callback)

    opts = UiAutomator2Options()
    opts.platform_name = "Android"
    opts.device_name = device_id
    opts.udid = device_id
    opts.app_package = NAVER_PACKAGE
    opts.app_activity = NAVER_ACTIVITY
    opts.no_reset = True
    opts.dont_stop_app_on_reset = True
    opts.auto_launch = True         # 앱 자동 시작 활성화 (C:\appium_auto 로직 참고)
    opts.new_command_timeout = 1200  # OCR(EasyOCR) 처리 시간 고려하여 20분으로 증가

    # 동시 기기 실행 시 포트 충돌 방지
    # systemPort: appium_port 기반 오프셋 (7723~8500 범위 → 12000~13000 범위)
    BASE_APPIUM_PORT = 7723
    system_port = 12000 + (appium_port - BASE_APPIUM_PORT)
    if system_port < 12000 or system_port > 13500:
        # 범위 밖이면 랜덤 오프셋 사용
        import random
        system_port = random.randint(12100, 13400)
    opts.set_capability("appium:systemPort", system_port)

    chrome_driver_port = 9515 + (appium_port - BASE_APPIUM_PORT)
    if chrome_driver_port < 9515 or chrome_driver_port > 11000:
        import random
        chrome_driver_port = random.randint(10100, 10900)
    opts.set_capability("appium:chromedriverPort", chrome_driver_port)

    # uiautomator2 타임아웃 및 안정성 향상 설정 추가
    opts.set_capability("appium:uiautomator2ServerLaunchTimeout", 90000)
    opts.set_capability("appium:uiautomator2ServerInstallTimeout", 90000)
    opts.set_capability("appium:adbExecTimeout", 90000)
    opts.set_capability("appium:skipServerInstall", False)          # 서버 재설치 강제
    opts.set_capability("appium:forceAppLaunch", True)
    opts.set_capability("appium:disableWindowAnimation", True)      # 윈도우 애니메이션 비활성화 (성능 향상)
    opts.set_capability("appium:ensureWebviewsHavePages", True)

    server_url = f"http://127.0.0.1:{appium_port}"

    # 초기 드라이버 생성 전 uiautomator2 서버 및 uiautomator 강제 종료
    _log(log_callback, f"  🧹 [{device_id}] 기기 내 기존 Appium uiautomator2 server 강제 종료...")
    try:
        subprocess.run(
            ["adb", "-s", device_id, "shell", "am", "force-stop", "io.appium.uiautomator2.server"],
            capture_output=True, timeout=5
        )
        subprocess.run(
            ["adb", "-s", device_id, "shell", "am", "force-stop", "io.appium.uiautomator2.server.test"],
            capture_output=True, timeout=5
        )
        subprocess.run(
            ["adb", "-s", device_id, "shell", "pkill", "-f", "uiautomator"],
            capture_output=True, timeout=5
        )
    except Exception:
        pass

    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            driver = webdriver.Remote(server_url, options=opts)
            _log(log_callback,
                 f"드라이버 연결 완료: {device_id} (포트 {appium_port}, systemPort {system_port})")
            return driver
        except Exception as e:
            err_msg = str(e)
            _log(log_callback, f"⚠️ [드라이버 연결 시도 {attempt}/{max_retries} 실패]: {err_msg[:200]}")
            if attempt < max_retries:
                _log(log_callback, "🔄 포트 및 ADB 포워딩 정리 후 재시도합니다 (대기 3초)...")
                try:
                    _kill_port_process(system_port)
                    subprocess.run(
                        ["adb", "-s", device_id, "forward", "--remove", f"tcp:{system_port}"],
                        capture_output=True, timeout=3
                    )
                    
                    # 'UiAutomation not connected' 오류 발생 시 설비 강제 재부팅, 그 외에는 기존 리셋 수행
                    if "UiAutomation not connected" in err_msg:
                        reboot_device_and_wait(device_id, log_callback)
                    else:
                        reset_device_uiautomation(device_id, log_callback)
                except Exception as clean_err:
                    _log(log_callback, f"  [정리 중 오류 발생]: {clean_err}")
                time.sleep(3)
            else:
                raise e


def force_stop_and_restart_app(driver, device_id: str, log_callback=None):
    """
    네이버 앱을 강제 종료한 뒤 재실행
    """
    import subprocess

    _log(log_callback, f"🔄 네이버 앱 강제 종료 시작 (device={device_id})")

    try:
        subprocess.run(
            ["adb", "-s", device_id, "shell", "input", "keyevent", "3"],
            capture_output=True, timeout=5
        )
        _log(log_callback, "  ▶ HOME 키 전송 완료")
        time.sleep(1)
    except Exception as e:
        _log(log_callback, f"  ⚠ HOME 키 전송 실패: {e}")

    try:
        result = subprocess.run(
            ["adb", "-s", device_id, "shell", "am", "force-stop", NAVER_PACKAGE],
            capture_output=True, text=True, timeout=8
        )
        if result.returncode == 0:
            _log(log_callback, "  ▶ adb force-stop 완료")
        else:
            _log(log_callback, f"  ⚠ adb force-stop returncode={result.returncode}")
    except Exception as e:
        _log(log_callback, f"  ⚠ adb force-stop 예외: {e}")
    time.sleep(1)

    try:
        driver.terminate_app(NAVER_PACKAGE)
        _log(log_callback, "  ▶ terminate_app 완료")
    except Exception as e:
        _log(log_callback, f"  ⚠ terminate_app 오류 (무시): {e}")

    _log(log_callback, "  ⏳ 앱 완전 종료 대기 (2초)...")
    time.sleep(2)

    _log(log_callback, "  🚀 네이버 앱 재실행 중 (adb am start)...")
    try:
        # 1순위: adb shell am start (start_activity는 신버전 Appium에서 제거됨)
        result = subprocess.run(
            ["adb", "-s", device_id, "shell", "am", "start",
             "-n", f"{NAVER_PACKAGE}/{NAVER_ACTIVITY}"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            _log(log_callback, "  ✅ adb am start 완료")
        else:
            _log(log_callback, f"  ⚠ adb am start 실패 ({result.stderr.strip()[:80]}), activate_app 폴백")
            driver.activate_app(NAVER_PACKAGE)
        time.sleep(7)
        _log(log_callback, "  ✅ 네이버 앱 재실행 완료")
    except Exception as act_err:
        _log(log_callback, f"  ⚠ 앱 재실행 실패, activate_app 폴백: {act_err}")
        try:
            driver.activate_app(NAVER_PACKAGE)
            time.sleep(7)
        except Exception as fb_err:
            _log(log_callback, f"  ⚠ activate_app 폴백도 실패: {fb_err}")


def wait_for_element(driver, xpath: str, timeout: int = 10, log_callback=None):
    """
    XPath 요소가 나타날 때까지 대기 후 반환
    타임아웃 시 None 반환 (예외 없음)
    """
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        return el
    except TimeoutException:
        _log(log_callback, f"[대기 타임아웃] {xpath[:80]}")
        return None


def wait_and_click(driver, xpath: str, timeout: int = 10, log_callback=None) -> bool:
    """
    XPath 요소 대기 후 클릭. 성공 시 True, 실패 시 False
    """
    el = wait_for_element(driver, xpath, timeout, log_callback)
    if el:
        try:
            el.click()
            return True
        except WebDriverException as e:
            _log(log_callback, f"[클릭 실패] {xpath[:60]}: {e}")
            return False
    return False


def safe_click(driver, xpath: str, log_callback=None) -> bool:
    """즉시 클릭 시도 (대기 없음). 요소 없으면 False"""
    try:
        el = driver.find_element(By.XPATH, xpath)
        el.click()
        return True
    except (NoSuchElementException, WebDriverException):
        return False


def wait_and_input(driver, xpath: str, text: str,
                   timeout: int = 10, log_callback=None) -> bool:
    """
    XPath 요소 대기 후 텍스트 입력 (mobile: type 및 send_keys 이중 지원)
    """
    el = wait_for_element(driver, xpath, timeout, log_callback)
    if el:
        try:
            el.click()
            time.sleep(0.3)
            
            # 1차 시도: mobile: type
            try:
                driver.execute_script("mobile: type", {"text": text})
                time.sleep(0.3)
            except Exception:
                pass
                
            # 2차 시도: 입력이 확인되지 않으면 send_keys 수행
            try:
                curr_val = el.text or ""
            except Exception:
                curr_val = ""
                
            if not curr_val or text not in curr_val:
                el.clear()
                time.sleep(0.3)
                el.send_keys(text)
                time.sleep(0.3)
            return True
        except Exception as e:
            _log(log_callback, f"[입력 실패] {xpath[:60]}: {e}")
            return False
    return False


def element_exists(driver, xpath: str, timeout: int = 3) -> bool:
    """요소 존재 여부 확인 (짧은 타임아웃)"""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        return True
    except TimeoutException:
        return False


def go_to_main_page(driver, log_callback=None):
    """네이버 앱 메인 페이지로 이동 (앱 재활성화)"""
    try:
        driver.activate_app(NAVER_PACKAGE)
        time.sleep(2)
        _log(log_callback, "메인 페이지 이동 완료")
    except WebDriverException as e:
        _log(log_callback, f"[오류] 메인 페이지 이동 실패: {e}")


def tap_by_coords(driver, x: int, y: int, log_callback=None) -> bool:
    """지정 좌표 탭 (W3C Actions 지원 및 driver.tap 폴백)"""
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.actions.action_builder import ActionBuilder
        from selenium.webdriver.common.actions.pointer_input import PointerInput
        from selenium.webdriver.common.actions import interaction

        touch = PointerInput(interaction.POINTER_TOUCH, "touch")
        actions = ActionChains(driver)
        actions.w3c_actions = ActionBuilder(driver, mouse=touch)
        actions.w3c_actions.pointer_action.move_to_location(x, y)
        actions.w3c_actions.pointer_action.click()
        actions.perform()
        return True
    except Exception as e:
        _log(log_callback, f"[W3C Actions 좌표 탭 실패] ({x},{y}): {e} → driver.tap 폴백 시도")
        try:
            driver.tap([(x, y)])
            return True
        except Exception as fb_err:
            _log(log_callback, f"[폴백 좌표 탭 실패] ({x},{y}): {fb_err}")
            return False


def _log(callback, message: str):
    """로그 콜백 호출 (없으면 print)"""
    if callback:
        callback(message)
    else:
        print(f"[AppiumHelper] {message}")
