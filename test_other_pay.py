import sys
import os
import time
import socket
import subprocess
from appium import webdriver
from appium.options.android import UiAutomator2Options

# 모듈 경로 추가
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_NAVER_DIR = os.path.join(_BASE_DIR, "naver_address_auto")
if _NAVER_DIR not in sys.path:
    sys.path.insert(0, _NAVER_DIR)

# 네이버 자동 주문 워커 임포트
from naver_order_worker import NaverOrderWorker
import naver_appium as ah_mod


def print_log(did, msg):
    print(f"[TEST] {msg}")


FIXED_APPIUM_PORT = 7723  # 메인 프로그램에서 실제 사용 중인 포트로 수정하세요


def check_appium_port(port, timeout=2):
    """
    지정된 포트에 Appium 서버가 응답하는지 확인한다.
    """
    print(f"📡 포트 {port} 확인 중...")

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(timeout)
        if s.connect_ex(("127.0.0.1", port)) != 0:
            print(f"❌ 포트 {port} 가 열려있지 않습니다.")
            return False

    try:
        import urllib.request
        with urllib.request.urlopen(
            f"http://127.0.0.1:{port}/status", timeout=timeout
        ) as resp:
            if resp.status == 200:
                print(f"✅ 포트 {port} 에서 Appium 서버 응답 확인!")
                return True
    except Exception as e:
        print(f"❌ 포트 {port} 는 열려있지만 Appium 응답이 없습니다: {e}")
        return False

    return False


def main():
    device_id = "ce10171a4c4faf0504"
    print(f"📱 대상 기기 고정됨: {device_id}")

    driver = None

    worker = NaverOrderWorker(
        device_id=device_id,
        appium_port=7723,
        order_manager=None,
        log_callback=print_log,
        status_callback=lambda did, msg: print(f"[Status] {msg}"),
    )

    opts = UiAutomator2Options()
    opts.platform_name = "Android"
    opts.device_name = device_id
    opts.udid = device_id
    opts.no_reset = True
    opts.dont_stop_app_on_reset = True
    opts.set_capability("appium:ensureWebviewsHavePages", True)
    opts.set_capability("appium:newCommandTimeout", 300)

    print(f"⏳ 메인 프로그램의 Appium 서버에 연결 시도 중... (고정 포트: {FIXED_APPIUM_PORT})")

    if check_appium_port(FIXED_APPIUM_PORT):
        try:
            driver = webdriver.Remote(
                f"http://127.0.0.1:{FIXED_APPIUM_PORT}", options=opts
            )
            worker.appium_port = FIXED_APPIUM_PORT
            worker.driver = driver
            print(f"✅ 포트 {FIXED_APPIUM_PORT} 에서 Appium 세션 연결 성공!")
        except Exception as e:
            print(f"❌ 세션 연결 중 오류 발생: {e}")
            driver = None

    if not driver:
        print("❌ Appium 연결 실패: 현재 실행 중인 메인 프로그램의 Appium 서버를 찾을 수 없습니다.")
        print(f"💡 FIXED_APPIUM_PORT 값({FIXED_APPIUM_PORT})이 메인 프로그램의 실제 포트와 일치하는지 확인하세요.")
        print("💡 메인 프로그램에서 해당 기기가 '연결됨' 상태인지 확인하세요.")
        os.system("pause")
        return

    recognize_dir = os.path.join(_BASE_DIR, "인식")
    os.makedirs(recognize_dir, exist_ok=True)

    print("\n=======================================================")
    print("🎯 '다른결재' 버튼 이미지 매칭 및 스크롤 테스트 시작")
    print("   핸드폰 화면을 그대로 두고 결과를 확인하세요!")
    print(f"   📸 인식되는 모든 잘라낸 이미지는 [{recognize_dir}] 에 저장됩니다.")
    print("=======================================================\n")

    success = worker._click_other_pay_button(max_scroll_attempts=15)

    print("\n=======================================================")
    if success:
        print("✅ 테스트 완료: 다른결재 인식 성공 및 검증 완료!")
    else:
        print("❌ 테스트 완료: 다른결재 버튼을 찾지 못했거나 클릭 검증에 실패했습니다.")
    print("=======================================================\n")

    try:
        driver.quit()
    except Exception:
        pass

    os.system("pause")


if __name__ == "__main__":
    main()