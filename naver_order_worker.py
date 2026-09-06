"""
naver_order_worker.py

네이버 쇼핑 앱 자동 주문 워커 (핸드폰 1대 기준)
개발리스트.md의 19단계 흐름을 구현합니다.

단계:
1~2   결재목록.xlsx 에서 완료여부 공백 데이터 루프
3     메인 페이지 진입, 7초 대기
4     네이버 플러스 스토어 탭 클릭, 5초 대기
5     스토어홈 팝업 처리 (하루/7일 보지 않기), 2초 대기
6     마이쇼핑 클릭, 5초 대기 → 팝업 재처리
7     마이쇼핑 화면 검색 버튼 클릭, 2초 대기
8     검색입력.png 이미지 인식 → 검색어 입력
9     검색아이콘.png 이미지 인식 클릭, 5초 대기
10    상품 리스트에서 판매자명 + 상품명 매칭 클릭, 5초 대기
11    구매하기 버튼 클릭, 5초 대기
12    체크박스.png 이미지 인식 클릭, 2초 대기
13    바로구매.png / 바로구매2/3/4 이미지 인식 클릭, 8초 대기
14    변경 버튼 클릭, 3초 대기
15    스크롤 다운
16    배송지 목록에서 수취인/전화번호 매칭 클릭, 5초 대기
17    스크롤 + 전액사용.png 이미지 인식 클릭, 3초 대기
18    결제하기 버튼 클릭, 5초 대기
19    비밀번호 숫자 이미지(p0~p9.png) 인식으로 각 자리 클릭
"""

import os
import sys
import time
import random
import threading
from typing import Callable, Optional

from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException, NoSuchElementException

# 기존 naver_address_auto의 헬퍼 재사용
_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "naver_address_auto")
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

try:
    import naver_appium as ah
except ImportError:
    import appium_helper as ah  # type: ignore

from order_manager import OrderManager, OrderRow

# ─── 이미지 파일 경로 (개발문서 폴더) ───────────────────────────────────────
_IMG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "개발문서")
_NUM_DIR = os.path.join(_IMG_DIR, "숫자")

IMG_SEARCH_INPUT  = os.path.join(_IMG_DIR, "검색입력.png")   # 검색 입력창 (단계 8)
IMG_SEARCH_INPUT2 = os.path.join(_IMG_DIR, "검색입력2.png")  # 검색 입력창 예비용
IMG_SEARCH_ICON   = os.path.join(_IMG_DIR, "검색아이콘.png") # 검색 아이콘 (단계 9)
IMG_CHECKBOX      = os.path.join(_IMG_DIR, "체크박스.png")   # 체크박스 (단계 12)
IMG_CHECKBOX2     = os.path.join(_IMG_DIR, "체크박스2.png")
IMG_CHECKBOX4     = os.path.join(_IMG_DIR, "체크박스4.png")
IMG_OPTION_SELECT = os.path.join(_IMG_DIR, "옵션 선택.png")  # 옵션 선택 텍스트 (체크박스 위)
IMG_DELIVERY_INFO = os.path.join(_IMG_DIR, "배송정보.png")  # 배송정보 텍스트 (체크박스 아래)
IMG_BUY_NOW       = os.path.join(_IMG_DIR, "바로구매.png")   # 바로구매 버튼 (단계 13)
IMG_BUY_NOW2      = os.path.join(_IMG_DIR, "바로구매2.png")
IMG_BUY_NOW3      = os.path.join(_IMG_DIR, "바로구매3.png")
IMG_BUY_NOW4      = os.path.join(_IMG_DIR, "바로구매4.png")
IMG_DELIVERY_MEMO = os.path.join(_IMG_DIR, "배송메모.png")   # 배송메모 드롭다운 (단계 16.5)
IMG_MEMO_NO_SELECT = os.path.join(_IMG_DIR, "선택안함.png")  # 배송메모 '선택안함' 옵션
IMG_ORDER_PAY     = os.path.join(_IMG_DIR, "주문결재.png")   # 주문결재 확인용 (단계 13 폴백)
IMG_FULL_USE      = os.path.join(_IMG_DIR, "전액사용.png")   # 전액사용 버튼 (단계 17)

# 추가된 결제방식 이미지
IMG_OTHER_PAY     = os.path.join(_IMG_DIR, "다른결재.png")
IMG_OTHER_PAY2    = os.path.join(_IMG_DIR, "다른결재수단2.png")
IMG_OTHER_PAY4    = os.path.join(_IMG_DIR, "다른결재4.png")
IMG_PAY_METHOD    = os.path.join(_IMG_DIR, "결재수단.png")
IMG_BOGI          = os.path.join(_IMG_DIR, "보기.png")
IMG_NORMAL_PAY    = os.path.join(_IMG_DIR, "일반결재.png")
IMG_NORMAL_PAY3   = os.path.join(_IMG_DIR, "일반결재3.png")
IMG_NORMAL_PAY_CHECK = os.path.join(_IMG_DIR, "일반결재체크.png")
IMG_BANK_TRANSFER = os.path.join(_IMG_DIR, "무통장입금.png")
IMG_BANK_TRANSFER_CHECK = os.path.join(_IMG_DIR, "무통장체크.png")
IMG_SELECT_BANK   = os.path.join(_IMG_DIR, "은행을.png")
IMG_SELECT_BANK2  = os.path.join(_IMG_DIR, "은행을2.png")
IMG_SELECT_BANK3  = os.path.join(_IMG_DIR, "은행을3.png")
IMG_SELECT_BANK4  = os.path.join(_IMG_DIR, "은행을4.png")
IMG_BANK_SELECT   = os.path.join(_IMG_DIR, "은행선택.png")
IMG_SHINHAN_BANK  = os.path.join(_IMG_DIR, "신한은행.png")
IMG_SHINHAN_BANK2 = os.path.join(_IMG_DIR, "신한은행2.png")
IMG_HANA_BANK     = os.path.join(_IMG_DIR, "하나은행.png")
IMG_HANA_BANK2    = os.path.join(_IMG_DIR, "하나은행2.png")
IMG_NONGHYUP_BANK = os.path.join(_IMG_DIR, "농협.png")
IMG_NONGHYUP_BANK2= os.path.join(_IMG_DIR, "농협2.png")
IMG_WOORI_BANK    = os.path.join(_IMG_DIR, "우리은행.png")
IMG_WOORI_BANK2   = os.path.join(_IMG_DIR, "우리은행2.png")
IMG_KB_BANK       = os.path.join(_IMG_DIR, "국민은행.png")
IMG_KB_BANK2      = os.path.join(_IMG_DIR, "국민은행2.png")
IMG_IBK_BANK      = os.path.join(_IMG_DIR, "기업은행.png")
IMG_IBK_BANK2     = os.path.join(_IMG_DIR, "기업은행2.png")
IMG_BUSAN_BANK    = os.path.join(_IMG_DIR, "부산은행.png")
IMG_BUSAN_BANK2   = os.path.join(_IMG_DIR, "부산은행2.png")
IMG_NOT_APPLY     = os.path.join(_IMG_DIR, "미신청.png")
IMG_NOT_APPLY2    = os.path.join(_IMG_DIR, "미신청2.png")
IMG_NOT_APPLY3    = os.path.join(_IMG_DIR, "미신청3.png")
IMG_NOT_APPLY4    = os.path.join(_IMG_DIR, "미신청4.png")
IMG_DO_PAY        = os.path.join(_IMG_DIR, "결재하기.png")
IMG_DO_ORDER      = os.path.join(_IMG_DIR, "주문하기.png")
IMG_BUY_BTN       = os.path.join(_IMG_DIR, "구매하기.png")
IMG_BUY_BTN2      = os.path.join(_IMG_DIR, "구매하기2.png")
IMG_BUY_BTN3      = os.path.join(_IMG_DIR, "구매하기3.png")
IMG_BUY_BTN4      = os.path.join(_IMG_DIR, "구매하기4.png")
IMG_MONEY_PAY     = os.path.join(_IMG_DIR, "머니.png")
IMG_PAY_MONEY_KR  = os.path.join(_IMG_DIR, "pay머니.png")
IMG_PAYL_MONEY    = os.path.join(_IMG_DIR, "payl머니.png")
IMG_PAY_BENEFIT   = os.path.join(_IMG_DIR, "결제혜택.png")  # 결제혜택 팝업 감지용
IMG_CLOSE_POPUP   = os.path.join(_IMG_DIR, "닫기.png")      # 팝업 닫기 버튼

# 현대카드 결제 이미지 (단계 22)
_HYUNDAI_NUM_DIR = os.path.join(_IMG_DIR, "현대숫자")
IMG_HYUNDAI_NUMS = {
    str(d): os.path.join(_HYUNDAI_NUM_DIR, f"{d}.png") for d in range(10)
}
HYUNDAI_PIN6 = "115080"  # 현대카드 1차 PIN (6자리, 개발리스트 22-8)
IMG_HYUNDAI_CARDS = [
    (os.path.join(_IMG_DIR, "카드를.png"), "카드를"),
    (os.path.join(_IMG_DIR, "카드를1.png"), "카드를1"),
    (os.path.join(_IMG_DIR, "카드를2.png"), "카드를2"),
    (os.path.join(_IMG_DIR, "카드를3.png"), "카드를3"),
    (os.path.join(_IMG_DIR, "카드를4.png"), "카드를4"),
]
IMG_HYUNDAI_BRAND = [
    (os.path.join(_IMG_DIR, "현대1.png"), "현대1"),
    (os.path.join(_IMG_DIR, "현대2.png"), "현대2"),
    (os.path.join(_IMG_DIR, "현대3.png"), "현대3"),
    (os.path.join(_IMG_DIR, "현대4.png"), "현대4"),
]
# 국민카드 (결제방식=국민카드): kb국민1/2 선택 → 결재하기 후 종료
IMG_KB_BRAND = [
    (os.path.join(_IMG_DIR, "kb국민1.png"), "kb국민1"),
    (os.path.join(_IMG_DIR, "kb국민2.png"), "kb국민2"),
]
IMG_HYUNDAI_DO_PAY = [
    (os.path.join(_IMG_DIR, "결재하기.png"), "결재하기"),
    (os.path.join(_IMG_DIR, "결재하기1.png"), "결재하기1"),
    (os.path.join(_IMG_DIR, "결재하기2.png"), "결재하기2"),
    (os.path.join(_IMG_DIR, "결재하기3.png"), "결재하기3"),
    (os.path.join(_IMG_DIR, "결재하기4.png"), "결재하기4"),
]
IMG_HYUNDAI_PIN_BTN = [
    (os.path.join(_IMG_DIR, "현대핀1.png"), "현대핀1"),
    (os.path.join(_IMG_DIR, "현대핀2.png"), "현대핀2"),
    (os.path.join(_IMG_DIR, "현대핀3.png"), "현대핀3"),
    (os.path.join(_IMG_DIR, "현대핀4.png"), "현대핀4"),
    (os.path.join(_IMG_DIR, "현대핀5.png"), "현대핀5"),
]
IMG_HYUNDAI_PIN_INPUT = [
    (os.path.join(_IMG_DIR, "핀입력1.png"), "핀입력1"),
    (os.path.join(_IMG_DIR, "핀입력2.png"), "핀입력2"),
    (os.path.join(_IMG_DIR, "핀입력3.png"), "핀입력3"),
    (os.path.join(_IMG_DIR, "핀입력4.png"), "핀입력4"),
    (os.path.join(_IMG_DIR, "핀입력5.png"), "핀입력5"),
]
IMG_HYUNDAI_CONFIRM = [
    (os.path.join(_IMG_DIR, "현대확인1.png"), "현대확인1"),
    (os.path.join(_IMG_DIR, "현대확인2.png"), "현대확인2"),
    (os.path.join(_IMG_DIR, "현대확인3.png"), "현대확인3"),
    (os.path.join(_IMG_DIR, "현대확인4.png"), "현대확인4"),
]
IMG_HYUNDAI_PAY_NOW = [
    (os.path.join(_IMG_DIR, "현대결제하기1.png"), "현대결제하기1"),
    (os.path.join(_IMG_DIR, "현대결제하기2.png"), "현대결제하기2"),
    (os.path.join(_IMG_DIR, "현대결제하기3.png"), "현대결제하기3"),
    (os.path.join(_IMG_DIR, "현대결제하기4.png"), "현대결제하기4"),
]
IMG_HYUNDAI_CARD_PW = [
    (os.path.join(_IMG_DIR, "현대카드비번1.png"), "현대카드비번1"),
    (os.path.join(_IMG_DIR, "현대카드비번2.png"), "현대카드비번2"),
    (os.path.join(_IMG_DIR, "현대카드비번3.png"), "현대카드비번3"),
    (os.path.join(_IMG_DIR, "현대카드비번4.png"), "현대카드비번4"),
]
# 현대결제하기 클릭 후 안전/추가인증 팝업 감지
IMG_HYUNDAI_SAFE_DETECT = [
    (os.path.join(_IMG_DIR, "안전한.png"), "안전한"),
    (os.path.join(_IMG_DIR, "안전한2.png"), "안전한2"),
    (os.path.join(_IMG_DIR, "안전한3.png"), "안전한3"),
    (os.path.join(_IMG_DIR, "안전결재.png"), "안전결재"),
    (os.path.join(_IMG_DIR, "안전결재2.png"), "안전결재2"),
    (os.path.join(_IMG_DIR, "추가인증.png"), "추가인증"),
]
# 팝업 본체(화이트 모달) — bbox 하단 = 확인 버튼
IMG_HYUNDAI_SAFE_POPUP_BODY = [
    (os.path.join(_IMG_DIR, "안전한3.png"), "안전한3"),
    (os.path.join(_IMG_DIR, "안전결재.png"), "안전결재"),
    (os.path.join(_IMG_DIR, "안전결재2.png"), "안전결재2"),
]
# 전체화면 참고 (안전결재3)
IMG_HYUNDAI_SAFE_POPUP_FULL = os.path.join(_IMG_DIR, "안전결재3.png")
# 팝업 닫기: 안전확인1~3 중 하나 클릭
IMG_HYUNDAI_SAFE_CONFIRM = [
    (os.path.join(_IMG_DIR, "안전확인1.png"), "안전확인1"),
    (os.path.join(_IMG_DIR, "안전확인2.png"), "안전확인2"),
    (os.path.join(_IMG_DIR, "안전확인3.png"), "안전확인3"),
]
SAFE_AUTH_TEXT_XPATHS = [
    '//*[contains(@text,"안전한 결제를 위해")]',
    '//*[contains(@text,"추가 인증을 진행합니다")]',
    '//*[contains(@text,"추가 인증")]',
]
# 현대비번.png = 키패드 영역 템플릿 (인식 후 ROI 커팅 → 숫자 입력)
IMG_HYUNDAI_PW_KEYPAD = os.path.join(_IMG_DIR, "현대비번.png")
IMG_HYUNDAI_PW_CONFIRM_FULL = os.path.join(_IMG_DIR, "현대비번 확인.png")
ORDER_COMPLETE_XPATH = '//android.widget.TextView[@text="주문완료 되었습니다"]'

# 비밀번호 숫자 이미지 (단계 19): p0.png ~ p9.png
IMG_NUMS = {
    str(d): os.path.join(_NUM_DIR, f"p{d}.png") for d in range(10)
}

# 핫스팟 활성 상태 이미지 (재부팅 후 핫스팟 켜기 감지)
IMG_ACTIVE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "활성.PNG")
if not os.path.exists(IMG_ACTIVE):
    IMG_ACTIVE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "naver_address_auto", "활성.PNG")


CREATE_NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0

def _run_cmd(cmd, **kwargs):
    """Windows GUI/GDI 프로세스 핸들 누수 방지 안전 subprocess 래퍼"""
    if sys.platform == "win32" and "creationflags" not in kwargs:
        kwargs["creationflags"] = CREATE_NO_WINDOW
    import subprocess
    return subprocess.run(cmd, **kwargs)


def _check_image_exists_on_device(device_id: str, template_path: str, threshold: float = 0.65) -> bool:
    """ADB screencap으로 화면을 캡처하여 템플릿 이미지가 존재하는지 확인 (main_order.py의 check_image_exists_on_device와 동일)"""
    if not os.path.exists(template_path):
        return False
    try:
        import cv2
        import numpy as np
        res = _run_cmd(
            ["adb", "-s", device_id, "exec-out", "screencap", "-p"],
            capture_output=True, timeout=8
        )
        if not res.stdout or len(res.stdout) < 100:
            return False
        img_arr = np.frombuffer(res.stdout, np.uint8)
        screen = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
        template = cv2.imread(template_path, cv2.IMREAD_COLOR)
        if screen is None or template is None:
            return False
        result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, _ = cv2.minMaxLoc(result)
        return max_val >= threshold
    except Exception:
        return False

# ─── XPath 상수 ─────────────────────────────────────────────────────────────

# 메인화면 - 네이버 플러스 스토어 탭 (단계 4)
STORE_TAB_XPATH = (
    '//android.view.ViewGroup[@content-desc="네이버 플러스 스토어, 버튼,"]'
    '/android.view.ViewGroup[@resource-id="com.nhn.android.search:id/tabIconLayout"]'
    '/android.widget.ImageView[@resource-id="com.nhn.android.search:id/tabIcon"]'
)

# 팝업 처리 (단계 5, 6)
HIDE_BTN_1DAY    = '//android.widget.Button[@text="하루 동안 보지 않기"]'
HIDE_BTN_7DAY    = '//android.widget.Button[@text="7일 동안 보지 않기"]'

# 3.1 & 7.1 웰컴 모달 / 팝업 닫기 버튼 목록
WELCOME_MODAL_XPATHS = [
    '//android.view.View[@resource-id="joinBeginWelcomeModal"]/android.view.View/android.view.View[2]',
    '//android.widget.Button[@resource-id="btnWelcomeClose"]',
    '//android.view.ViewGroup[@content-desc="홈 버튼"]/android.view.ViewGroup/android.widget.ImageView[@resource-id="com.nhn.android.search.InAppBrowser:id/toolbarIconView"]',
]

# 마이쇼핑 (단계 6)
MY_SHOPPING_XPATH = '//android.view.View[@content-desc="마이쇼핑"]'

# 마이쇼핑 화면 검색 버튼 (단계 7) - 마이쇼핑.xml 참고
SEARCH_BTN_IN_MY_XPATH = '//android.widget.Button[@text="검색"]'

# 구매하기 버튼 (단계 11)
BUY_BTN_XPATH = '//android.widget.Button[@text="구매하기"]'

# 변경 버튼 (단계 14)
CHANGE_BTN_XPATH = '//android.widget.Button[@text="변경"]'

# 결제하기 버튼 (단계 18)
PAY_BTN_XPATH = '//android.widget.Button[contains(@text,"결제하기")]'

# 배송지 목록 (단계 16) - 배송지목록.xml 참고
DELIVERY_LIST_WEBVIEW_XPATHS = [
    '//android.webkit.WebView[@text="주문/결제"]/android.view.View/android.view.View/android.view.View',
    '//android.webkit.WebView/android.view.View/android.view.View',
    '//android.webkit.WebView',
]

# 타임아웃
TASK_TIMEOUT_SEC = 900  # 주문 1건 최대 15분 (현대카드 PIN 대기 포함)


class NaverOrderWorker:
    """
    네이버 자동 주문 워커

    기존 NaverWorker(배송지 등록)와 다르게 구매/결제 흐름을 담당합니다.
    기기 1대 기준으로 단독 실행합니다.
    """

    def __init__(self,
                 device_id: str,
                 appium_port: int,
                 order_manager: OrderManager,
                 log_callback: Optional[Callable] = None,
                 status_callback: Optional[Callable] = None,
                 machine_num: int = 1,
                 test_mode: bool = False,
                 manual_mode: bool = False):
        self.device_id      = device_id
        self.appium_port    = appium_port
        self.order_manager  = order_manager
        self._log_cb        = log_callback
        self._status_cb     = status_callback
        self.machine_num    = machine_num
        self.test_mode      = test_mode
        # 수동시작: 배송지 선택까지 진행 + 엑셀 Y 기록 후 종료
        self.manual_mode    = manual_mode
        self.driver         = None
        self._stop_event    = threading.Event()

    def _skip_final_order_click(self) -> bool:
        """테스트/수동시작 모드에서는 주문하기·결제하기 최종 클릭을 생략"""
        return bool(self.test_mode or self.manual_mode)

    # ─── 공개 메서드 ─────────────────────────────────────────────────────────

    def run(self) -> bool:
        """워커 메인 실행 (별도 스레드에서 호출)"""
        if self.manual_mode:
            self._log("🖐 수동시작 워커 시작 (배송지 선택까지 → Y 기록 후 종료)")
        else:
            self._log("🚀 자동 주문 워커 시작")

        # ── 재부팅 후 핫스팟 활성 여부 사전 확인 (활성.PNG 이미지 인식) ──
        if os.path.exists(IMG_ACTIVE):
            self._log(f"🔍 [활성.PNG] 핫스팟 활성 상태 사전 확인 중...")
            if _check_image_exists_on_device(self.device_id, IMG_ACTIVE, threshold=0.65):
                self._log("✅ [활성.PNG] 인식 성공 → 핫스팟 활성 확인, 다음 작업으로 진행")
            else:
                self._log("ℹ️ [활성.PNG] 미감지 → 핫스팟 미활성 또는 이미지 없음, 계속 진행")

        max_restarts = 10

        for attempt in range(1, max_restarts + 1):
            if self._stop_event.is_set():
                break

            self._set_status(f"연결 중... ({attempt}/{max_restarts})")

            try:
                self.driver = ah.create_driver(self.device_id, self.appium_port, self._log)
                self._log("🔄 네이버 앱 재시작")
                ah.force_stop_and_restart_app(self.driver, self.device_id, self._log)
            except Exception as e:
                self._log(f"❌ 드라이버 연결 실패: {e}")
                self._set_status("연결 실패")
                if attempt < max_restarts:
                    time.sleep(10)
                    continue
                return False

            success = False
            try:
                # [단계 3~6] 메인 → 스토어 → 마이쇼핑 진입
                self._go_main_and_enter_store()

                # [단계 2~19] 결재목록 루프 주문 처리
                self._order_loop()
                success = True

            except Exception as e:
                self._log(f"❌ 예기치 않은 오류: {e}")
                self._set_status("오류 발생")
            finally:
                if self.driver:
                    try:
                        self.driver.quit()
                    except Exception:
                        pass
                self.driver = None

            if success or self._stop_event.is_set():
                break

            self._log(f"🔄 오류 회복 재시작... ({attempt}/{max_restarts})")
            time.sleep(5)

        self._log("🏁 워커 종료")
        self._set_status("완료")
        return True

    def stop(self):
        self._stop_event.set()
        self._log("⏹ 중지 요청됨")

    # ─── 단계 3~6: 메인 → 스토어 → 마이쇼핑 ─────────────────────────────────

    # ─── 단계 3~6: 메인 → 스토어 → 마이쇼핑 (3.2 계정전환 포함) ─────────────────

    def _go_main_and_enter_store(self, login_id: str = "") -> bool:
        """[단계 3] 메인 페이지 진입, 7초 대기 및 3.1 웰컴 모달 처리 -> [단계 3.2] 계정 전환 -> [단계 4~6] 스토어/마이쇼핑 진입"""
        self._set_status("메인 페이지 이동 중")
        ah.force_stop_and_restart_app(self.driver, self.device_id, self._log)
        time.sleep(3)

        # [단계 3.1] 웰컴 모달 / 팝업 발견 시 클릭
        self._check_and_close_welcome_modals(step_label="3.1")

        # [단계 3.2] 로그인 아이디가 지정된 경우 네이버 계정 전환 수행
        if login_id:
            if not self._switch_account(login_id):
                self._log(f"❌ [단계 3.2] 계정 전환 실패 (아이디: {login_id}) -> 4번 이하 단계 수행 취소 및 다음 레코드 이동")
                return False
            # 계정 전환 성공 후 → 네이버 앱 강제 재시작으로 UiAutomator2 안정화
            self._log("🔄 [단계 3.2 완료] 계정 전환 후 네이버 앱 재시작으로 UiAutomator2 안정화")
            ah.force_stop_and_restart_app(self.driver, self.device_id, self._log)
            # UiAutomator2 instrumentation 완전 안정화 대기
            self._wait_for_uiautomator_ready(max_wait=15)

        # [단계 4] 네이버 플러스 스토어 탭 버튼 클릭
        if ah.element_exists(self.driver, STORE_TAB_XPATH, timeout=5):
            self._log("📌 네이버 플러스 스토어 탭 감지 → 클릭")
            ah.wait_and_click(self.driver, STORE_TAB_XPATH, timeout=7, log_callback=self._log)
            time.sleep(3)
        else:
            self._log("⏭ 스토어 탭 없음 (이미 스토어 화면)")
            time.sleep(2)

        # [단계 5] 팝업 처리
        self._dismiss_popups()

        # [단계 6] 마이쇼핑 클릭 -> 5초 대기
        self._set_status("마이쇼핑 클릭")
        if ah.element_exists(self.driver, MY_SHOPPING_XPATH, timeout=8):
            ah.wait_and_click(self.driver, MY_SHOPPING_XPATH, timeout=7, log_callback=self._log)
            self._log("✅ 마이쇼핑 클릭 완료 (5초 대기)")
            time.sleep(5)
        else:
            self._log("⚠ 마이쇼핑 버튼 미발견")
            time.sleep(2)

        # 마이쇼핑 진입 후 팝업 및 웰컴 모달 재처리
        self._dismiss_popups()
        self._check_and_close_welcome_modals(step_label="6.1")
        return True

    def _switch_account(self, login_id: str) -> bool:
        """
        [단계 3.2] 네이버 계정 자동 전환
        1) //android.widget.ImageView[@content-desc="메뉴"] 클릭 (3초 대기)
        2) //android.widget.ImageView[@content-desc="설정"] 클릭 (3초 대기)
        3) //android.view.ViewGroup[contains(@content-desc, "로그인 아이디 관리")] 클릭 (3초 대기)
        4) 아이디선택.xml 구조 참고하여 content-desc에 login_id가 포함된 요소를 탐색 및 클릭 (3초 대기)
           - 미발견 시: 작업을 중단하고 로그 기록 후 False 반환
        5) (//android.view.View[contains(@content-desc, "{login_id}") and contains(@content-desc, "로그인 중")]) 요소를 통해 로그인 검증
           - 미발견/불일치 시: 작업 중단 및 로그 기록 후 False 반환
        """
        if not login_id:
            self._log("ℹ [단계 3.2] 로그인 아이디 미지정 -> 계정 전환 건너뜀")
            return True

        self._set_status(f"계정 전환 시도: {login_id}")
        self._log(f"🔑 [단계 3.2] 네이버 계정 전환 시작 (타겟 아이디: {login_id})")

        # 1. 메뉴 버튼 발견 하면 클릭 (3초 대기)
        menu_xpaths = [
            '//android.widget.ImageView[@content-desc="메뉴"]',
            '//*[@content-desc="메뉴"]',
        ]
        menu_clicked = False
        for xpath in menu_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=4):
                self._log("  📌 메뉴 버튼 발견 -> 클릭")
                ah.wait_and_click(self.driver, xpath, timeout=4, log_callback=self._log)
                time.sleep(2)
                menu_clicked = True
                break

        if not menu_clicked:
            self._log("  ❌ [단계 3.2] 메뉴 버튼 미발견 -> 계정 전환 실패")
            return False

        # 2. 설정 버튼 발견하면 클릭 (3초 대기)
        setting_xpaths = [
            '//android.widget.ImageView[@content-desc="설정"]',
            '//*[@content-desc="설정"]',
        ]
        setting_clicked = False
        for xpath in setting_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=4):
                self._log("  📌 설정 버튼 발견 -> 클릭")
                ah.wait_and_click(self.driver, xpath, timeout=4, log_callback=self._log)
                time.sleep(2)
                setting_clicked = True
                break

        if not setting_clicked:
            self._log("  ❌ [단계 3.2] 설정 버튼 미발견 -> 계정 전환 실패")
            return False

        # 3. 로그인 아이디 관리 링크 발견하면 클릭 (3초 대기)
        mgmt_xpaths = [
            '//android.view.ViewGroup[contains(@content-desc, "로그인 아이디 관리")]',
            '//android.widget.LinearLayout[@resource-id="com.nhn.android.search.Setup:id/pref_title_cell"]/android.widget.LinearLayout[1]',
            '//android.widget.LinearLayout[@resource-id="com.nhn.android.search.Setup:id/pref_title_cell"]',
            '//*[contains(@content-desc, "로그인 아이디 관리")]',
        ]
        mgmt_clicked = False
        for xpath in mgmt_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=4):
                self._log("  📌 로그인 아이디 관리 링크 발견 -> 클릭")
                ah.wait_and_click(self.driver, xpath, timeout=4, log_callback=self._log)
                time.sleep(2)
                mgmt_clicked = True
                break

        if not mgmt_clicked:
            self._log("  ❌ [단계 3.2] 로그인 아이디 관리 링크 미발견 -> 계정 전환 실패")
            return False

        # 3.5. 오타 보정 로직: resource-id 비존재 시 content-desc 기반으로도 수집
        actual_login_id = login_id
        try:
            import difflib as _difflib
            id_els = self.driver.find_elements(
                By.XPATH,
                '//*[@resource-id="com.nhn.android.search:id/idText"]'
                ' | //android.view.View[contains(@content-desc, "간편로그인")]'
                ' | //android.view.View[contains(@content-desc, "로그인 중")]'
            )
            available_ids = []
            for el in id_els:
                try:
                    txt = el.get_attribute("text") or el.get_attribute("content-desc") or ""
                    if txt:
                        for suffix in [", 로그인 중, 간편로그인", ", 로그인 중",
                                       " , 간편로그인", ", 간편로그인", ", 더보기"]:
                            txt = txt.replace(suffix, "")
                        txt = txt.strip()
                        if len(txt) >= 3 and " " not in txt:
                            available_ids.append(txt)
                except Exception:
                    pass

            if available_ids:
                if login_id in available_ids:
                    actual_login_id = login_id
                else:
                    matches = _difflib.get_close_matches(login_id, available_ids, n=1, cutoff=0.6)
                    if matches:
                        actual_login_id = matches[0]
                        self._log(f"  ℹ 아이디 오타 보정: '{login_id}' -> '{actual_login_id}' 로 매칭됨")
                    else:
                        for aid in available_ids:
                            if aid.startswith(login_id[:4]) or login_id.startswith(aid[:4]):
                                actual_login_id = aid
                                self._log(f"  ℹ 아이디 접두사 보정: '{login_id}' -> '{actual_login_id}' 로 매칭됨")
                                break
        except Exception as e:
            self._log(f"  ⚠ 아이디 목록 추출 중 예외: {e}")

        # 4. '로그인 중' 상태 확인
        # 방법1: XPath contains 체크 (login_id / actual_login_id 양쪽)
        # 방법2: 화면의 모든 '로그인 중' 요소를 수집해서 difflib 유사도 체크 (오타 대응)
        already_logged_in = False
        for cid in [login_id, actual_login_id]:
            for xpath in [
                f'//android.view.View[contains(@content-desc, "{cid}") and contains(@content-desc, "로그인 중")]',
                f'//*[contains(@content-desc, "{cid}") and contains(@content-desc, "로그인 중")]',
            ]:
                if ah.element_exists(self.driver, xpath, timeout=2):
                    already_logged_in = True
                    self._log(f"  ✅ [단계 3.2] 계정 [{actual_login_id}] 이미 '로그인 중' 상태임 (XPath 매칭)")
                    break
            if already_logged_in:
                break

        # XPath 미감지 시 → '로그인 중' 요소 전수 조사 (difflib 유사도)
        if not already_logged_in:
            try:
                import difflib as _difflib
                login_els = self.driver.find_elements(
                    By.XPATH, '//*[contains(@content-desc, "로그인 중")]'
                )
                for el in login_els:
                    desc = el.get_attribute("content-desc") or ""
                    raw_id = desc
                    for suffix in [", 로그인 중, 간편로그인", ", 로그인 중", " , 간편로그인", ", 간편로그인"]:
                        raw_id = raw_id.replace(suffix, "")
                    raw_id = raw_id.strip()
                    if not raw_id:
                        continue
                    for cid in [login_id, actual_login_id]:
                        sim = _difflib.SequenceMatcher(None, cid, raw_id).ratio()
                        if sim >= 0.75 or raw_id.startswith(cid[:5]) or cid.startswith(raw_id[:5]):
                            actual_login_id = raw_id
                            already_logged_in = True
                            self._log(
                                f"  ✅ [단계 3.2] 계정 [{raw_id}] 이미 '로그인 중' 상태임"
                                f" (유사도 {sim:.2f}, 입력 아이디: {cid})"
                            )
                            break
                    if already_logged_in:
                        break
            except Exception as e:
                self._log(f"  ⚠ '로그인 중' 전수 조사 예외: {e}")

        # verify_xpaths는 클릭 후 검증에서 재사용
        verify_xpaths = [
            f'//android.view.View[contains(@content-desc, "{actual_login_id}") and contains(@content-desc, "로그인 중")]',
            f'//*[contains(@content-desc, "{actual_login_id}") and contains(@content-desc, "로그인 중")]',
        ]

        if not already_logged_in:
            target_account_xpaths = [
                # 1순위: clickable=true인 부모 컨테이너 직접 탐색
                f'//android.view.View[contains(@content-desc, "{actual_login_id}") and contains(@content-desc, "간편로그인")]/ancestor-or-self::android.view.View[@clickable="true"]',
                f'//android.view.View[contains(@content-desc, "{actual_login_id}") and not(contains(@content-desc, "더보기"))]/ancestor-or-self::android.view.View[@clickable="true"]',
                # 2순위: content-desc 정확히 매칭
                f'//android.view.View[@content-desc="{actual_login_id} , 간편로그인"]',
                f'//android.view.View[contains(@content-desc, "{actual_login_id}") and contains(@content-desc, "간편로그인")]',
                # 3순위: 더보기 제외한 일반 탐색
                f'//android.view.View[contains(@content-desc, "{actual_login_id}") and not(contains(@content-desc, "더보기"))]',
                f'//*[contains(@content-desc, "{actual_login_id}") and not(contains(@content-desc, "더보기"))]',
                f'//*[contains(@text, "{actual_login_id}") and not(contains(@text, "더보기"))]',
            ]
            target_el = None
            for xpath in target_account_xpaths:
                if ah.element_exists(self.driver, xpath, timeout=4):
                    try:
                        target_el = self.driver.find_element(By.XPATH, xpath)
                        self._log(f"  📌 아이디선택 화면에서 '{actual_login_id}' 발견 -> 강력 클릭 진행")
                        break
                    except Exception:
                        continue

            if not target_el:
                self._log(f"  ❌ [단계 3.2] 로그인 아이디 [{actual_login_id}] 미발견 -> 작업 중지 및 로그 기록")
                return False

            # 클릭 전 좌표 미리 계산 (클릭 후 StaleElementReferenceException 방지)
            import subprocess
            tap_x, tap_y = None, None
            try:
                rect = target_el.rect
                tap_x = int(rect['x'] + min(200, rect['width'] * 0.4))
                tap_y = int(rect['y'] + rect['height'] // 2)
            except Exception as e:
                self._log(f"  ⚠ 클릭 전 좌표 획득 예외: {e}")

            click_triggered = False
            for attempt in range(1, 3):
                self._log(f"  👉 [단계 3.2] 로그인 아이디 [{actual_login_id}] 클릭 시도 ({attempt}/2)")
                
                # 클릭: clickable 여부와 무관하게 항상 ADB 좌표 탭 사용 (clickable=false 요소 대응)
                if tap_x and tap_y:
                    try:
                        self._log(f"  👉 ADB input tap: ({tap_x}, {tap_y})")
                        subprocess.run(
                            ["adb", "-s", self.device_id, "shell", "input", "tap", str(tap_x), str(tap_y)],
                            capture_output=True, timeout=5
                        )
                        click_triggered = True
                        self._log("  ✅ ADB tap 전송 완료")
                    except Exception as e:
                        self._log(f"  ⚠ ADB input tap 예외: {e}")

                # ADB tap 실패 시에만 direct click 폴백
                if not click_triggered:
                    try:
                        target_el.click()
                        self._log("  ✅ direct click 전송 완료 (폴백)")
                        click_triggered = True
                    except Exception as e:
                        err_str = str(e)
                        if any(kw in err_str for kw in ["StaleElement", "StaleObject", "not linked", "does not exist"]):
                            self._log("  ✅ direct click 성공으로 화면 전환됨 (StaleElement 감지)")
                            click_triggered = True
                            break
                        else:
                            self._log(f"  ⚠ direct click 예외: {e}")

                time.sleep(3)

                # 클릭 후 확인 팝업(button1) 존재 시 클릭 (계정 전환 확인 다이얼로그)
                CONFIRM_BTN_XPATH = '//android.widget.Button[@resource-id="android:id/button1"]'
                if ah.element_exists(self.driver, CONFIRM_BTN_XPATH, timeout=2):
                    self._log("  📌 계정 전환 확인 팝업(button1) 감지 → 클릭")
                    ah.wait_and_click(self.driver, CONFIRM_BTN_XPATH, timeout=2, log_callback=self._log)
                    time.sleep(2)

                # 클릭 후 '로그인 중' 상태 최종 검증
                verified = False
                for xpath in verify_xpaths:
                    if ah.element_exists(self.driver, xpath, timeout=4):
                        verified = True
                        break

                if verified:
                    self._log(f"  ✅ [단계 3.2] 로그인 아이디 [{actual_login_id}] '로그인 중' 전환 확인 성공!")
                    break
                else:
                    self._log(f"  ⚠ [{actual_login_id}] '로그인 중' 전환 미확인 -> 재시도")
                    if attempt < 2:
                        time.sleep(1)

            # 5. 최종 검증: 이미 로그인중이었거나 클릭 후 로그인중 상태가 검증되어야 함
            if not already_logged_in and not verified:
                self._log(f"  ❌ [단계 3.2] 계정 [{actual_login_id}] '로그인 중' 상태 검증 실패 -> 작업 중단 및 다음 레드로 이동")
                return False

        # 6. 존재하면 //android.widget.ScrollView/android.view.View[1]/android.widget.Button 클릭 (2초 대기)
        back_btn_xpaths = [
            '//android.widget.ScrollView/android.view.View[1]/android.widget.Button',
            '//android.view.View[@content-desc="뒤로"]',
        ]
        for xpath in back_btn_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=3):
                self._log("  📌 아이디선택 화면 상단 뒤로/버튼 발견 -> 클릭 (2초 대기)")
                ah.wait_and_click(self.driver, xpath, timeout=3, log_callback=self._log)
                time.sleep(2)
                break

        # 7. //android.widget.ImageView[@content-desc="이전"] 발견하면 클릭 (2초 대기)
        prev_btn_xpaths = [
            '//android.widget.ImageView[@content-desc="이전"]',
            '//*[@content-desc="이전"]',
        ]
        for xpath in prev_btn_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=3):
                self._log("  📌 설정 화면 '이전' 버튼 발견 -> 클릭 (2초 대기)")
                ah.wait_and_click(self.driver, xpath, timeout=3, log_callback=self._log)
                time.sleep(2)
                break

        return True

    def _wait_for_uiautomator_ready(self, max_wait: int = 15):
        """
        UiAutomator2 instrumentation 서버가 정상 동작할 때까지 대기.
        계정 전환 후 설정 앱 왕복 과정에서 instrumentation이 불안정해지는 현상 대응.
        간단한 get_window_size 호출로 서버 응답을 확인하고 실패 시 재시도.
        """
        self._log(f"⏳ UiAutomator2 안정화 대기 (최대 {max_wait}초)...")
        interval = 2
        for i in range(max_wait // interval):
            try:
                self.driver.get_window_size()
                self._log(f"  ✅ UiAutomator2 정상 응답 확인 ({i * interval}초 경과)")
                return
            except Exception as e:
                err_str = str(e)
                if "instrumentation" in err_str or "not connected" in err_str or "not running" in err_str:
                    self._log(f"  ⚠ UiAutomator2 아직 불안정 ({i * interval}초 경과), {interval}초 후 재시도...")
                    time.sleep(interval)
                else:
                    # 다른 종류의 에러는 그냥 넘어감
                    time.sleep(interval)
        # 최대 대기 후에도 실패하면 경고만 남기고 계속 진행 (재시작 로직이 외부에 있음)
        self._log(f"  ⚠ UiAutomator2 안정화 대기 {max_wait}초 초과, 계속 진행합니다.")

    def _check_and_close_welcome_modals(self, step_label: str = "3.1/7.1"):
        """3.1 및 7.1 웰컴 모달 / 팝업 버튼 발견 시 클릭"""
        for xpath in WELCOME_MODAL_XPATHS:
            try:
                if ah.element_exists(self.driver, xpath, timeout=2):
                    self._log(f"📌 [{step_label}] 웰컴 모달/팝업 버튼 발견 → 클릭 시도: {xpath[:50]}")
                    ah.wait_and_click(self.driver, xpath, timeout=3, log_callback=self._log)
                    time.sleep(1.0)
            except Exception as e:
                self._log(f"  ⚠ [{step_label}] 모달 닫기 예외: {e}")

    def _dismiss_popups(self):
        """하루/7일 동안 보지 않기 팝업 처리"""
        for xpath in [HIDE_BTN_1DAY, HIDE_BTN_7DAY]:
            if ah.element_exists(self.driver, xpath, timeout=5):
                label = "하루" if "하루" in xpath else "7일"
                self._log(f"📌 '{label} 동안 보지 않기' 팝업 감지 → 클릭")
                ah.wait_and_click(self.driver, xpath, timeout=5, log_callback=self._log)
                time.sleep(1.5)
        # 한 번 더 하루 보지 않기 확인 (문서 요구사항)
        if ah.element_exists(self.driver, HIDE_BTN_1DAY, timeout=3):
            ah.wait_and_click(self.driver, HIDE_BTN_1DAY, timeout=5, log_callback=self._log)
            time.sleep(1.5)

    # ─── 단계 7: 마이쇼핑 검색 버튼 클릭 ────────────────────────────────────

    def _click_search_in_my_shopping(self) -> bool:
        """
        [단계 7] 마이쇼핑.xml 참고 - 검색 버튼 클릭
        bounds="[417,2052][630,2220]"
        + [단계 7.1] 웰컴 모달/팝업 처리
        """
        self._set_status("검색 버튼 클릭")
        clicked = False
        if ah.element_exists(self.driver, SEARCH_BTN_IN_MY_XPATH, timeout=5):
            ah.wait_and_click(self.driver, SEARCH_BTN_IN_MY_XPATH, timeout=5, log_callback=self._log)
            self._log("✅ 마이쇼핑 검색 버튼 클릭 완료")
            time.sleep(1.5)
            clicked = True
        else:
            self._log("⚠ 검색 버튼 미발견 → 좌표 탭 시도")
            try:
                size = self.driver.get_window_size()
                w, h = size['width'], size['height']
                # 마이쇼핑.xml 기준 비율: x=523/1080, y=2136/2400
                tap_x = int(w * 0.484)
                tap_y = int(h * 0.890)
                ah.tap_by_coords(self.driver, tap_x, tap_y, self._log)
                self._log(f"✅ 검색 버튼 좌표 탭 완료 ({tap_x}, {tap_y})")
                time.sleep(1.5)
                clicked = True
            except Exception as e:
                self._log(f"❌ 검색 버튼 클릭 실패: {e}")
                clicked = False

        # [단계 7.1] 웰컴 모달/팝업 닫기 버튼 발견 시 클릭
        self._check_and_close_welcome_modals(step_label="7.1")
        return clicked


    # ─── 단계 8~9: 검색어 입력 + 검색 실행 ──────────────────────────────────

    def _input_search_keyword(self, keyword: str) -> bool:
        """
        [단계 8] 검색입력 이미지 / OCR 인식 → 검색 입력창 클릭 → 검색어 입력
        """
        self._set_status(f"검색어 입력: {keyword}")
        self._log(f"🔍 검색어 입력: '{keyword}'")

        tap_coords = None

        # (사용자 요청으로 검색입력.png 등 이미지 매칭 방식은 제외됨)

        # 2순위: OCR 탐색 (상품명, 브랜드, 검색 등)
        if not tap_coords:
            self._log("🔍 [OCR] '상품명' 또는 '브랜드' 등 텍스트 검색 시도")
            try:
                import cv2, numpy as np
                from paddleocr import PaddleOCR
                res = _run_cmd(
                    ["adb", "-s", self.device_id, "exec-out", "screencap", "-p"],
                    capture_output=True, timeout=8
                )
                if res.stdout and len(res.stdout) > 100:
                    img_arr = np.frombuffer(res.stdout, np.uint8)
                    screen = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
                    if screen is not None:
                        ocr_engine = PaddleOCR(use_angle_cls=True, lang="korean")
                        results = ocr_engine.ocr(screen, cls=True)
                        if results:
                            for page in results:
                                if not page: continue
                                for item in page:
                                    box = item[0]
                                    text_info = item[1]
                                    text = text_info[0].replace(" ", "")
                                    conf = float(text_info[1])
                                    if conf > 0.5 and any(kw in text for kw in ["상품", "브랜드", "쇼핑몰", "검색어", "입력"]):
                                        cx = int((box[0][0] + box[2][0]) / 2)
                                        cy = int((box[0][1] + box[2][1]) / 2)
                                        tap_coords = (cx, cy)
                                        self._log(f"  ✅ [OCR] '{text}' 발견 → 좌표: {tap_coords}")
                                        break
                                if tap_coords: break
            except Exception as e:
                self._log(f"  ⚠ [OCR 검색] 오류: {e}")

        # 3순위: EditText XPath
        if not tap_coords:
            search_xpaths = [
                '//android.widget.EditText[@hint="검색어를 입력해주세요"]',
                '//android.widget.EditText[@hint="검색어 입력"]',
                '//android.widget.EditText[@hint="상품, 브랜드, 쇼핑몰 검색"]',
                '//android.widget.EditText',
            ]
            for xpath in search_xpaths:
                if ah.element_exists(self.driver, xpath, timeout=2):
                    try:
                        el = self.driver.find_element(By.XPATH, xpath)
                        rect = el.rect
                        tap_coords = (rect['x'] + rect['width'] // 2,
                                      rect['y'] + rect['height'] // 2)
                        self._log(f"  ✅ EditText 발견 → 좌표: {tap_coords}")
                        break
                    except Exception:
                        continue

        if tap_coords:
            # 입력창 클릭
            ah.tap_by_coords(self.driver, tap_coords[0], tap_coords[1], self._log)
            time.sleep(1.0)
        else:
            self._log("  ⚠ 검색 입력창 좌표 획득 실패. (기본 화면 중앙 상단 클릭 폴백 시도)")
            try:
                sz = self.driver.get_window_size()
                ah.tap_by_coords(self.driver, int(sz['width'] * 0.5), int(sz['height'] * 0.15), self._log)
                time.sleep(1.0)
            except Exception:
                pass

        # 텍스트 입력
        try:
            focused_xpath = '//android.widget.EditText[@focused="true"]'
            el = None
            if ah.element_exists(self.driver, focused_xpath, timeout=2):
                el = self.driver.find_element(By.XPATH, focused_xpath)
            elif ah.element_exists(self.driver, '//android.widget.EditText', timeout=2):
                el = self.driver.find_element(By.XPATH, '//android.widget.EditText')

            if el:
                try:
                    el.clear()
                    time.sleep(0.3)
                except Exception:
                    pass
                try:
                    self.driver.execute_script("mobile: type", {"text": keyword})
                    time.sleep(0.5)
                except Exception:
                    el.send_keys(keyword)
                    time.sleep(0.5)
            else:
                raise Exception("EditText 요소를 찾지 못함")

            self._log(f"  ✅ 검색어 입력 완료: '{keyword}'")
            return True
            
        except Exception as e:
            self._log(f"  ❌ 검색어 입력 실패: {e}")
            return False

    def _click_search_button(self) -> bool:
        """
        [단계 9] 검색 실행 (키보드 엔터 우선), 5초 대기
        """
        self._set_status("검색 실행")

        # 1순위: 키보드 엔터 (검색)
        try:
            import subprocess
            subprocess.run(
                ["adb", "-s", self.device_id, "shell", "input", "keyevent", "66"],
                capture_output=True, timeout=5
            )
            self._log("  ✅ 엔터 키 전송 완료 (검색)")
            time.sleep(4)
            return True
        except Exception as e:
            self._log(f"  ⚠ 엔터 키 전송 실패: {e}")

        # 2순위: 이미지 매칭 (검색아이콘.png 폴백)
        if os.path.exists(IMG_SEARCH_ICON):
            coords = self._find_image_coords(IMG_SEARCH_ICON, threshold=0.7)
            if coords:
                ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                self._log("  ✅ 검색아이콘 이미지 인식 클릭 완료 (폴백)")
                time.sleep(3)
                return True

        # 3순위: 좌표 탭
        try:
            size = self.driver.get_window_size()
            w, h = size['width'], size['height']
            tap_x = int(w * 0.94)
            tap_y = int(h * 0.08)
            ah.tap_by_coords(self.driver, tap_x, tap_y, self._log)
            self._log(f"  ✅ 검색 버튼 좌표 탭 완료 ({tap_x}, {tap_y})")
            time.sleep(3)
            return True
        except Exception as e:
            self._log(f"  ❌ 검색 버튼 클릭 최종 실패: {e}")
            return False

    # ─── 단계 10: 상품 매칭 클릭 ────────────────────────────────────────────

    def _click_product(self, seller_name: str, product_name: str) -> bool:
        """
        [단계 10] 상품 리스트에서 판매자명 + 상품명 매칭 클릭
        - 판매자명 TextView/View + 상품명 View/TextView 조합 매칭
        - 수량/숫자 단위(예: 2박스 vs 5박스) 불일치 시 매칭 배제
        - 스크롤 탐색 중에는 완전 매칭(score >= 2)만 즉시 클릭하며, 부분 매칭은 스크롤 완료 후 폴백
        """
        self._set_status(f"상품 선택: {product_name[:15]}")
        self._log(f"🔍 상품 매칭 시도: 판매자='{seller_name}', 상품='{product_name}'")

        # 상품명 키워드 정제 (2글자 이상의 의미있는 키워드 목록)
        import re
        clean_prod = re.sub(r'[\+\-\*\/\(\)\[\]\{\}\?\!\,]', ' ', product_name)
        keywords = [k.strip() for k in clean_prod.split() if len(k.strip()) >= 2]
        main_keyword = keywords[0] if keywords else product_name[:5]
        safe_keyword = main_keyword.replace('"', '').replace("'", "")
        clean_full_name = product_name.replace('"', '').replace("'", "")

        # 타겟 상품명에서 숫자+단위 패턴 추출 (예: '2박스', '14포', '1박스', '5박스', '10개' 등)
        target_num_units = re.findall(r'\d+\s*(?:박스|포|개|입|g|ml|kg|L|세트|EA|ea|가지|회)?', product_name, re.IGNORECASE)
        target_num_units = [nu.replace(" ", "").lower() for nu in target_num_units if nu.strip()]

        def check_num_conflict(cand_txt: str) -> bool:
            """타겟에 포함된 숫자/수량 단위가 후보 텍스트와 다르게 들어있는지 검사"""
            cand_txt_clean = cand_txt.replace(" ", "").lower()
            for t_nu in target_num_units:
                # 숫자 파싱
                m_num = re.search(r'\d+', t_nu)
                if not m_num:
                    continue
                num_str = m_num.group(0)
                # 만약 t_nu가 '2박스'라면, cand_txt에 '5박스', '1박스', '3박스' 등이 있으면 충돌
                unit_m = re.search(r'(박스|포|개|입|g|ml|kg|L|세트|EA|ea)', t_nu, re.IGNORECASE)
                if unit_m:
                    unit_str = unit_m.group(0)
                    # cand_txt에서 같은 단위를 사용하는 숫자패턴 찾기
                    cand_matches = re.findall(r'(\d+)\s*' + re.escape(unit_str), cand_txt_clean)
                    if cand_matches and num_str not in cand_matches:
                        return True  # 충돌 발견!
                else:
                    # 단위 없이 단순 숫자 (예: 2 vs 5)
                    # target에는 '2'가 있는데 cand에는 '5'만 있고 '2'가 없다면 충돌 우려
                    pass
            return False

        fallback_candidates = []
        scroll_max = 20

        for scroll_cnt in range(scroll_max + 1):
            try:
                # 방법 0: 직접 XPath 텍스트 매칭 시도 (우선)
                direct_xpaths = [
                    f'//android.view.View[contains(@text, "{safe_keyword}")]',
                    f'//android.widget.TextView[contains(@text, "{safe_keyword}")]',
                    f'//*[contains(@content-desc, "{safe_keyword}")]'
                ]
                for dxpath in direct_xpaths:
                    try:
                        els = self.driver.find_elements(By.XPATH, dxpath)
                        for el in els:
                            text = el.text or ''
                            desc = el.get_attribute('content-desc') or ''
                            full_txt = f"{text} {desc}"

                            if check_num_conflict(full_txt):
                                continue

                            matched_kws = [kw for kw in keywords if kw in full_txt]
                            is_full_match = (len(matched_kws) == len(keywords)) or (clean_full_name.replace(" ", "") in full_txt.replace(" ", ""))

                            if is_full_match:
                                self._log(f"  📌 상품명 직접 매칭 발견: {dxpath}")
                                if self._safe_click_element(el):
                                    time.sleep(3)
                                    return True
                    except Exception:
                        pass

                # 방법 A: 화면의 모든 View, TextView 검색하여 판매자 + 상품명 둘 다 포함된 상품 카드 찾기
                all_views = self.driver.find_elements(
                    By.XPATH,
                    '//android.view.View | //android.widget.TextView'
                )

                candidate_elements = []

                for el in all_views:
                    try:
                        text = el.text or ''
                        desc = el.get_attribute('content-desc') or ''
                        full_txt = f"{text} {desc}"

                        if not full_txt.strip():
                            continue

                        # 수량/숫자 충돌 검사
                        if check_num_conflict(full_txt):
                            continue

                        matched_kws = [kw for kw in keywords if kw in full_txt]
                        is_full_match = (len(matched_kws) == len(keywords)) or (clean_full_name.replace(" ", "") in full_txt.replace(" ", ""))

                        # 1. 판매자명 + 상품명 전체 일치
                        if (seller_name and seller_name in full_txt) and is_full_match:
                            candidate_elements.append((3, len(matched_kws), el, full_txt))
                        # 2. 상품명 전체 일치
                        elif is_full_match:
                            candidate_elements.append((2, len(matched_kws), el, full_txt))
                        # 3. 상품명 키워드 대부분(길이-1) 일치 (폴백 후보로 저장은 하되 스크롤 완료 후 선택)
                        elif len(keywords) > 2 and len(matched_kws) >= len(keywords) - 1:
                            fallback_candidates.append((1, len(matched_kws), el, full_txt))
                    except Exception:
                        continue

                # 점수 높은 순 정렬 (우선순위 -> 키워드 일치 개수)
                candidate_elements.sort(key=lambda x: (x[0], x[1]), reverse=True)

                if candidate_elements:
                    best_score, best_kw_cnt, best_el, best_txt = candidate_elements[0]
                    # 완전 매칭(score >= 2)인 경우에만 스크롤 진행 도중 즉시 클릭!
                    if best_score >= 2:
                        self._log(f"  📌 상품 완전 매칭 성공 (점수={best_score}, 키워드={best_kw_cnt}/{len(keywords)}): '{best_txt[:40]}...'")
                        if self._safe_click_element(best_el):
                            time.sleep(3)
                            return True
                        else:
                            self._log("  ⚠ 좌표 클릭 실패, 계속 탐색...")

            except Exception as e:
                self._log(f"  ⚠ 상품 탐색 중 오류: {e}")

            if scroll_cnt < scroll_max:
                self._log(f"  ⬇ 스크롤 다운 ({scroll_cnt + 1}/{scroll_max})")
                # 상품 리스트: 지문검증/재시도 없이 빠른 ADB 스와이프 (간격 단축)
                self._scroll_down_fast(distance_ratio=0.28)
                time.sleep(0.35)

        # 20회 스크롤 완료 후에도 완전 매칭이 없었던 경우 폴백 후보 사용
        if fallback_candidates:
            fallback_candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
            best_score, best_kw_cnt, best_el, best_txt = fallback_candidates[0]
            self._log(f"  📌 [폴백] 상품 부분 매칭 선택 (점수={best_score}, 키워드={best_kw_cnt}/{len(keywords)}): '{best_txt[:40]}...'")
            if self._safe_click_element(best_el):
                time.sleep(3)
                return True

        self._log(f"  ❌ 상품 매칭 최종 실패: {seller_name} / {product_name}")
        return False

    def _safe_click_element(self, el) -> bool:
        """
        요소의 bounds를 파싱하여 중심 좌표를 ADB tap으로 클릭
        언제나 좌표 탭만 사용 (el.click() 다중 실행 금지)
        """
        import re as _re
        import subprocess
        try:
            # 1. bounds attribute 직접 파싱 (가장 정확)
            bounds_str = None
            try:
                bounds_str = el.get_attribute("bounds")
            except Exception:
                pass

            if bounds_str:
                m = _re.match(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', bounds_str)
                if m:
                    x1, y1, x2, y2 = map(int, m.groups())
                    x = (x1 + x2) // 2
                    y = (y1 + y2) // 2
                else:
                    bounds_str = None

            # 2. bounds 파싱 실패 시 rect 사용
            if not bounds_str:
                rect = el.rect
                x = rect['x'] + rect['width'] // 2
                y = rect['y'] + rect['height'] // 2

            w_w, w_h = 1080, 2400
            try:
                size = self.driver.get_window_size()
                w_w, w_h = size['width'], size['height']
            except Exception:
                pass

            # 화면 영역 내부인지 검사 후, 중앙이 아닐 경우 스크롤로 중앙 정렬
            if 150 <= y <= w_h - 200 and 0 < x < w_w:
                # 화면 중앙(target_y)과 현재 y 차이를 스크롤로 보정
                target_y = w_h // 2
                offset = y - target_y
                if abs(offset) > int(w_h * 0.15):  # 화면 높이 15% 이상 차이날 때만 보정
                    import subprocess as _sp
                    self._log(f"  📐 상품을 화면 중앙으로 이동 (y={y} → 목표={target_y}, offset={offset})")
                    # swipe: 아래로 치우친 경우 위로 스크롤, 위로 치우친 경우 아래로 스크롤
                    swipe_start_y = int(w_h * 0.5)
                    swipe_end_y   = swipe_start_y - offset  # offset 만큼 반대로 스와이프
                    swipe_end_y   = max(50, min(w_h - 50, swipe_end_y))
                    swipe_x       = w_w // 2
                    _sp.run(
                        ["adb", "-s", self.device_id, "shell", "input", "swipe",
                         str(swipe_x), str(swipe_start_y),
                         str(swipe_x), str(swipe_end_y), "400"],
                        capture_output=True, timeout=5
                    )
                    time.sleep(0.8)
                    # 스크롤 후 좌표 갱신
                    try:
                        bounds_str3 = el.get_attribute("bounds")
                        import re as _re2
                        m3 = _re2.match(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', bounds_str3 or "")
                        if m3:
                            x1, y1, x2, y2 = map(int, m3.groups())
                            x, y = (x1 + x2) // 2, (y1 + y2) // 2
                        else:
                            rect = el.rect
                            x = rect['x'] + rect['width'] // 2
                            y = rect['y'] + rect['height'] // 2
                    except Exception:
                        pass

                self._log(f"  👉 ADB 좌표 탭: ({x}, {y})")
                if not self._is_visible_coord(x, y):
                    self._log(f"  ⏭ 화면 밖 좌표 클릭 생략: ({x}, {y})")
                    return False
                subprocess.run(
                    ["adb", "-s", self.device_id, "shell", "input", "tap",
                     str(x), str(y)],
                    capture_output=True, timeout=5
                )
                return True
            else:
                self._log(f"  ⚠ 요소 좌표가 화면 표시 범위를 벗어남 (y={y}, 화면높이={w_h}) -> 미세 스크롤")
                self._scroll_down()
                time.sleep(1)
                # 스크롤 후 재시도
                try:
                    bounds_str2 = el.get_attribute("bounds")
                    m2 = _re.match(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', bounds_str2 or "")
                    if m2:
                        x1, y1, x2, y2 = map(int, m2.groups())
                        x, y = (x1 + x2) // 2, (y1 + y2) // 2
                    else:
                        rect = el.rect
                        x = rect['x'] + rect['width'] // 2
                        y = rect['y'] + rect['height'] // 2
                except Exception:
                    pass
                if self._is_visible_coord(x, y) and 150 <= y <= w_h - 200:
                    subprocess.run(
                        ["adb", "-s", self.device_id, "shell", "input", "tap",
                         str(x), str(y)],
                        capture_output=True, timeout=5
                    )
                    return True
                self._log(f"  ⏭ 스크롤 후에도 화면 밖 ({x}, {y}) → 클릭 안 함")
                return False

        except Exception as e:
            self._log(f"  ⚠ 안전 클릭 실패: {e}")

        # 폴백: el.click()
        try:
            el.click()
            return True
        except Exception:
            return False

    def _click_element_or_parent(self, el) -> bool:
        """기존 하위 호환 래퍼"""
        return self._safe_click_element(el)


    # ─── 단계 11: 구매하기 버튼 ──────────────────────────────────────────────

    def _click_buy_button(self) -> bool:
        """[단계 11] 구매하기 버튼 클릭 (구매하기, 구매하기2, 구매하기3, 구매하기4 이미지 중 하나 인식 시 즉시 클릭), 3초 대기"""
        self._set_status("구매하기 클릭")

        buy_img_candidates = [
            (IMG_BUY_BTN,  "구매하기"),
            (IMG_BUY_BTN2, "구매하기2"),
            (IMG_BUY_BTN3, "구매하기3"),
            (IMG_BUY_BTN4, "구매하기4"),
        ]

        # 1순위: 4가지 이미지 후보 중 하나라도 매칭되면 즉시 클릭
        for img_path, img_name in buy_img_candidates:
            if os.path.exists(img_path):
                coords = self._find_image_coords(img_path, threshold=0.70)
                if coords:
                    self._log(f"  🎯 [{img_name}] 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    self._log(f"✅ [{img_name}] 이미지 인식 클릭 완료")
                    time.sleep(3)
                    return True

        # 2순위: XPath 매칭 폴백
        if ah.element_exists(self.driver, BUY_BTN_XPATH, timeout=3):
            ah.wait_and_click(self.driver, BUY_BTN_XPATH, timeout=5, log_callback=self._log)
            self._log("✅ 구매하기 XPath 버튼 클릭 완료")
            time.sleep(3)
            return True

        self._log("⚠ 구매하기 버튼 미발견 → 계속 진행")
        return True  # 없어도 계속 진행

    # ─── 단계 12: 체크박스 이미지 인식 클릭 ──────────────────────────────────

    def _score_template_in_region(self, template_path, screen_gray, screen_w, screen_h,
                                  min_x, max_x, min_y, max_y):
        """한 장의 그레이 스크린샷에서 템플릿 최고점/좌표. 영역 밖이면 무효.
        returns (cx, cy, score) 또는 None"""
        try:
            import cv2
            import numpy as np
        except ImportError:
            return None
        if not os.path.exists(template_path):
            return None
        template_bgr = cv2.imdecode(
            np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR
        )
        if template_bgr is None:
            return None
        template_gray = cv2.cvtColor(template_bgr, cv2.COLOR_BGR2GRAY)
        t_h, t_w = template_gray.shape
        roi = screen_gray.copy()
        if min_y > 0:
            roi[:min_y, :] = 0
        if max_y < screen_h:
            roi[max_y:, :] = 0
        if min_x > 0:
            roi[:, :min_x] = 0
        if max_x < screen_w:
            roi[:, max_x:] = 0

        best_score, best_loc, best_tw, best_th = -1.0, None, t_w, t_h
        for scale in np.linspace(0.55, 1.65, 12):
            new_w, new_h = int(t_w * scale), int(t_h * scale)
            if new_w >= screen_w or new_h >= screen_h or new_w < 8 or new_h < 8:
                continue
            resized = cv2.resize(template_gray, (new_w, new_h), interpolation=cv2.INTER_AREA)
            try:
                r = cv2.matchTemplate(roi, resized, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(r)
                if max_val > best_score:
                    best_score, best_loc, best_tw, best_th = max_val, max_loc, new_w, new_h
            except Exception:
                continue
        if best_loc is None:
            return None
        cx = best_loc[0] + best_tw // 2
        cy = best_loc[1] + best_th // 2
        if not (min_x <= cx <= max_x and min_y <= cy <= max_y):
            return None
        return cx, cy, float(best_score)

    def _click_checkbox(self, product_name: str = "") -> bool:
        """[단계 12] 체크박스2/4/원본 중 인식률이 가장 높은 것을,
        옵션선택과 배송정보 사이(왼쪽 열)에서만 찾아 클릭."""
        self._set_status("체크박스/옵션 선택")
        self._log("🔍 체크박스 및 옵션 항목 탐색 시도 중...")

        w_h, w_w = 2400, 1080
        try:
            size = self.driver.get_window_size()
            w_h, w_w = size['height'], size['width']
        except Exception:
            pass

        opt_y = del_y = None
        if os.path.exists(IMG_OPTION_SELECT):
            opt_coords = self._find_image_coords(IMG_OPTION_SELECT, threshold=0.70)
            if opt_coords:
                opt_y = opt_coords[1]
        if os.path.exists(IMG_DELIVERY_INFO):
            del_coords = self._find_image_coords(IMG_DELIVERY_INFO, threshold=0.70)
            if del_coords:
                del_y = del_coords[1]

        # 옵션선택 라벨·'옵션 필수선택' 헤더를 건너뛴 뒤 ~ 배송정보 직전
        # = 화살표가 가리키는 옵션 행 체크박스 간격
        header_skip = max(85, int(w_h * 0.036))
        if opt_y and del_y and opt_y < del_y:
            min_y_check = opt_y + header_skip
            max_y_check = del_y - 45
            if min_y_check >= max_y_check:
                min_y_check = opt_y + 50
                max_y_check = del_y - 20
            self._log(
                f"  📌 체크박스 탐색 구간 (옵션선택~배송정보 사이): "
                f"y={min_y_check}~{max_y_check}, 옵션y={opt_y}, 배송y={del_y}"
            )
        elif opt_y:
            min_y_check = opt_y + header_skip
            max_y_check = int(w_h * 0.82)
            self._log(f"  📌 체크박스 탐색 구간: 옵션선택 아래 y={min_y_check}~{max_y_check}")
        else:
            min_y_check = int(w_h * 0.55)
            max_y_check = int(w_h * 0.82)
            self._log("  ⚠ 옵션선택 미검출 → 화면 하단 시트로 제한")

        min_x_check = 0
        max_x_check = int(w_w * 0.22)

        checkbox_imgs = [
            p for p in (IMG_CHECKBOX2, IMG_CHECKBOX4, IMG_CHECKBOX) if os.path.exists(p)
        ]
        if not checkbox_imgs:
            self._log("  ⚠ 체크박스 템플릿 파일 없음")
            return True

        min_score = 0.78

        def _pick_best():
            try:
                import cv2
                import numpy as np
                from PIL import Image
                import io
            except ImportError:
                self._log("  [이미지 매칭] cv2/numpy/PIL 미설치")
                return None
            png = self._get_screenshot()
            pil = Image.open(io.BytesIO(png))
            screen_bgr = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
            screen_gray = cv2.cvtColor(screen_bgr, cv2.COLOR_BGR2GRAY)
            sh, sw = screen_gray.shape
            ranked = []
            for path in checkbox_imgs:
                hit = self._score_template_in_region(
                    path, screen_gray, sw, sh,
                    min_x_check, max_x_check, min_y_check, max_y_check,
                )
                name = os.path.basename(path)
                if hit is None:
                    self._log(f"  ℹ {name}: 구간 내 매칭 없음")
                    continue
                cx, cy, score = hit
                if not self._is_visible_coord(cx, cy):
                    self._log(
                        f"  ⏭ {name}: 점수 {score:.4f} 좌표 ({cx}, {cy}) 는 화면 밖 → 제외"
                    )
                    continue
                ranked.append((score, cx, cy, path))
                self._log(f"  ℹ {name}: 점수 {score:.4f} 좌표 ({cx}, {cy})")
            if not ranked:
                return None
            ranked.sort(key=lambda t: t[0], reverse=True)
            best = ranked[0]
            self._log(
                f"  🎯 최고 인식: {os.path.basename(best[3])} "
                f"점수 {best[0]:.4f} @ ({best[1]}, {best[2]})"
            )
            if best[0] < min_score:
                self._log(f"  ⚠ 최고점도 {best[0]:.4f} < {min_score} → 오탐 가능, 채택 안 함")
                return None
            return best

        best = _pick_best()
        if best:
            score, cx, cy, path = best
            img_name = os.path.basename(path)
            if not self._is_visible_coord(cx, cy):
                self._log(f"  ⏭ {img_name} 좌표 ({cx}, {cy}) 화면 밖 → 클릭 안 함")
            else:
                self._log(f"  👉 {img_name} 체크박스 ADB soft tap: ({cx}, {cy})")
                if not self._soft_tap(cx, cy, duration_ms=180):
                    pass
                else:
                    time.sleep(1.2)
                    again = _pick_best()
                    if again is None:
                        self._log(f"✅ {img_name} 클릭 후 구간 내 미체크 소멸 → 선택 완료")
                        return True
                    if again[3] == path and abs(again[2] - cy) <= 40 and again[0] >= min_score:
                        self._log("  ⚠ 같은 위치 미체크 잔존 → 한 번 더 탭")
                        self._soft_tap(cx, cy, duration_ms=180)
                        time.sleep(0.8)
                    else:
                        self._log(f"✅ {img_name} 클릭 완료 (원래 위치 미체크 아님)")
                    return True

        if product_name:
            import re
            hangul = re.sub(r'[^가-힣0-9]', ' ', product_name)
            kws = [k for k in hangul.split() if len(k) >= 2][:4]
            extra = []
            for k in list(kws):
                if len(k) >= 4:
                    extra.append(k[:4])
            for kw in kws + extra:
                try:
                    els = self.driver.find_elements(
                        By.XPATH, f'//*[contains(@text, "{kw}")]'
                    )
                    for el in els:
                        text = (el.get_attribute("text") or "")
                        if any(s in text for s in ("옵션", "배송", "바로구매", "장바구니")):
                            continue
                        rect = el.rect
                        cy = rect['y'] + rect['height'] // 2
                        if min_y_check <= cy <= max_y_check:
                            tap_x = int(w_w * 0.11)
                            if not self._is_visible_coord(tap_x, cy):
                                self._log(f"  ⏭ 옵션 행 좌표 ({tap_x}, {cy}) 화면 밖 → 스킵")
                                continue
                            self._log(f"  👉 옵션 행 '{text[:40]}' 왼쪽 체크박스 탭: ({tap_x}, {cy})")
                            self._soft_tap(tap_x, cy, duration_ms=180)
                            time.sleep(0.8)
                            return True
                except Exception:
                    continue

        if opt_y and del_y and opt_y < del_y:
            tap_x = int(w_w * 0.11)
            tap_y = (min_y_check + max_y_check) // 2
            if not self._is_visible_coord(tap_x, tap_y):
                self._log(f"  ⏭ 폴백 좌표 ({tap_x}, {tap_y}) 화면 밖 → 클릭 안 함")
            else:
                self._log(f"  ⚠ 이미지 미채택 → 옵션~배송 사이 왼쪽 탭 ({tap_x}, {tap_y})")
                self._soft_tap(tap_x, tap_y, duration_ms=180)
                time.sleep(0.8)
                return True

        self._log("  ⚠ 체크박스 미발견 → 계속 진행")
        return True

    # ─── 단계 13: 바로구매 이미지 인식 클릭 ──────────────────────────────────

    def _click_bottom_cta(self, min_y: int, max_y: int, min_x: int = 0) -> bool:
        """옵션 시트 하단 CTA(바로구매 / 바로 구매 / 구매하기)를 XPath로 클릭"""
        xpaths = [
            '//android.widget.Button[@text="바로구매"]',
            '//android.widget.Button[@text="바로 구매"]',
            '//android.widget.Button[contains(@text,"바로구매")]',
            '//android.widget.Button[contains(@text,"바로 구매")]',
            '//*[@content-desc="바로구매"]',
            '//*[@content-desc="바로 구매"]',
            '//*[contains(@content-desc,"바로구매")]',
            '//android.view.View[@text="바로구매"]',
            '//android.widget.TextView[@text="바로구매"]',
            '//*[@text="바로구매"]',
            '//*[@text="바로 구매"]',
            '//android.widget.Button[@text="구매하기"]',
        ]
        seen = set()
        for xpath in xpaths:
            try:
                if not ah.element_exists(self.driver, xpath, timeout=1):
                    continue
                for el in self.driver.find_elements(By.XPATH, xpath):
                    try:
                        rect = el.rect
                    except Exception:
                        continue
                    cx = rect['x'] + rect['width'] // 2
                    cy = rect['y'] + rect['height'] // 2
                    key = (cx, cy)
                    if key in seen:
                        continue
                    seen.add(key)
                    if cy < min_y or cy > max_y:
                        continue
                    if cx < min_x:
                        continue
                    txt = (el.get_attribute("text") or el.get_attribute("content-desc") or "").strip()
                    if "구매조건" in txt:
                        continue
                    if self._safe_click_element(el):
                        self._log(f"  ✅ 하단 구매 CTA 클릭 (text={txt!r}, x={cx}, y={cy})")
                        time.sleep(5)
                        return True
            except Exception:
                continue
        return False

    def _is_order_pay_screen(self) -> bool:
        """주문/결제 화면 진입 여부 확인 (바로구매 성공 판정용)"""
        markers = [
            '//android.webkit.WebView[@text="주문/결제"]',
            '//android.widget.TextView[@text="주문/결제"]',
            '//android.widget.Button[@text="변경"]',
            '//*[contains(@text,"결제하기")]',
            '//*[contains(@text,"배송지명")]',
            '//*[contains(@text,"배송메모")]',
        ]
        for xpath in markers:
            try:
                if ah.element_exists(self.driver, xpath, timeout=1):
                    return True
            except Exception:
                continue
        if os.path.exists(IMG_ORDER_PAY):
            try:
                if self._find_image_coords(IMG_ORDER_PAY, threshold=0.65):
                    return True
            except Exception:
                pass
        return False

    def _click_buy_now(self) -> bool:
        """[단계 13] 바로구매 클릭 (XPath 우선 → 하단 이미지만 → 검증)"""
        self._set_status("바로구매 클릭")
        time.sleep(1.0)

        # 이미 주문/결제 화면이면 성공
        if self._is_order_pay_screen():
            self._log("  ✅ 이미 주문/결제 화면 → 바로구매 생략")
            return True

        w_h, w_w = 2400, 1080
        try:
            size = self.driver.get_window_size()
            w_h, w_w = size['height'], size['width']
        except Exception:
            pass

        # CTA는 화면 하단. 중단(y≈1540) 오매칭 방지를 위해 72% 이상으로 제한
        min_y_buynow = int(w_h * 0.72)
        max_y_buynow = int(w_h * 0.98)
        min_x_right = int(w_w * 0.30)

        buy_now_imgs = [
            (p, n) for p, n in (
                (IMG_BUY_NOW, "바로구매"),
                (IMG_BUY_NOW2, "바로구매2"),
                (IMG_BUY_NOW3, "바로구매3"),
                (IMG_BUY_NOW4, "바로구매4"),
            ) if os.path.exists(p)
        ]

        def _confirm_after_click(label: str) -> bool:
            time.sleep(4.0)
            if self._is_order_pay_screen():
                self._log(f"  ✅ {label} 후 주문/결제 화면 확인")
                return True
            self._log(f"  ⚠ {label} 후 주문/결제 화면 미확인 → 오클릭 가능")
            return False

        for attempt in range(1, 4):
            # 1) XPath 우선 (하단 CTA만)
            if self._click_bottom_cta(min_y_buynow, max_y_buynow, min_x=0):
                if _confirm_after_click("XPath CTA"):
                    return True

            # 2) 이미지: 하단만, threshold 완화하되 Y는 엄격
            for thr in (0.65, 0.58, 0.52):
                for img_path, img_name in buy_now_imgs:
                    coords = self._find_image_coords(
                        img_path, threshold=thr,
                        min_x=min_x_right, min_y=min_y_buynow, max_y=max_y_buynow,
                    )
                    if not coords:
                        continue
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    self._log(f"✅ {img_name} 이미지 인식 클릭 (threshold={thr}, y={coords[1]})")
                    if _confirm_after_click(img_name):
                        return True

            # 3) 옵션시트 확인이 '구매하기'인 경우 (하단만)
            for img_path, img_name in (
                (IMG_BUY_BTN, "구매하기"),
                (IMG_BUY_BTN2, "구매하기2"),
                (IMG_BUY_BTN3, "구매하기3"),
                (IMG_BUY_BTN4, "구매하기4"),
            ):
                if not os.path.exists(img_path):
                    continue
                coords = self._find_image_coords(
                    img_path, threshold=0.70,
                    min_x=min_x_right, min_y=min_y_buynow, max_y=max_y_buynow,
                )
                if not coords:
                    continue
                ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                self._log(f"✅ 옵션시트 '{img_name}' 이미지 클릭 (바로구매 대체, y={coords[1]})")
                if _confirm_after_click(img_name):
                    return True

            if attempt < 3:
                self._log(f"  ⚠ 바로구매 미확인 ({attempt}회차) → 옵션 시트 재오픈 후 재시도")
                self._click_buy_button()
                time.sleep(1.2)

        self._log("  ❌ 바로구매 버튼 미발견/미확인")
        try:
            btns = self.driver.find_elements(By.XPATH, '//android.widget.Button')
            names = []
            for b in btns[:12]:
                t = (b.get_attribute("text") or b.get_attribute("content-desc") or "").strip()
                if t:
                    names.append(t)
            if names:
                self._log(f"  ℹ 현재 화면 Button: {names}")
        except Exception:
            pass

        if self._is_order_pay_screen():
            self._log("  ✅ 주문/결제 화면 확인됨 → 바로구매 성공으로 간주")
            return True
        return False

    # ─── 단계 14: 변경 버튼 클릭 ─────────────────────────────────────────────

    def _click_change_button(self) -> bool:
        """[단계 14] 변경 버튼 클릭 (XPath → 좌표 탭), 미발견 시 False"""
        self._set_status("변경 버튼 클릭")
        time.sleep(1.0)

        w_h = 2400
        try:
            w_h = self.driver.get_window_size()['height']
        except Exception:
            pass
        # 배송지 영역 '변경'은 상단~중상단. 하단 CTA와 구분
        max_y_change = int(w_h * 0.55)

        xpaths = [
            CHANGE_BTN_XPATH,
            '//android.widget.Button[contains(@text,"변경")]',
            '//*[@text="변경" and @clickable="true"]',
            '//android.view.View[@text="변경"]',
        ]

        for xpath in xpaths:
            try:
                if not ah.element_exists(self.driver, xpath, timeout=3):
                    continue
                for el in self.driver.find_elements(By.XPATH, xpath):
                    try:
                        rect = el.rect
                    except Exception:
                        continue
                    cx = rect['x'] + rect['width'] // 2
                    cy = rect['y'] + rect['height'] // 2
                    if cy > max_y_change:
                        continue
                    txt = (el.get_attribute("text") or "").strip()
                    self._log(f"  📌 변경 버튼 발견 (text={txt!r}, x={cx}, y={cy})")
                    if ah.tap_by_coords(self.driver, cx, cy, self._log):
                        self._log("✅ 변경 버튼 좌표 클릭 완료")
                        time.sleep(2.5)
                        return True
                    if self._safe_click_element(el):
                        self._log("✅ 변경 버튼 클릭 완료")
                        time.sleep(2.5)
                        return True
            except Exception:
                continue

        self._log("❌ 변경 버튼 미발견")
        try:
            btns = self.driver.find_elements(By.XPATH, '//android.widget.Button')
            names = [
                (b.get_attribute("text") or b.get_attribute("content-desc") or "").strip()
                for b in btns[:15]
            ]
            names = [n for n in names if n]
            if names:
                self._log(f"  ℹ 현재 화면 Button: {names}")
        except Exception:
            pass
        return False

    # ─── 단계 16: 배송지 선택 ────────────────────────────────────────────────

    def _recipient_name_matches(self, text: str, recipient_name: str) -> bool:
        """배송지/수취인 텍스트에 목표 수취인명이 포함되는지 확인.
        예: '박경아', '배송지명박경아', '배송지명박경아(박경아)'
        """
        if not recipient_name or not text:
            return False
        name = recipient_name.strip()
        if not name:
            return False
        if name in text:
            return True
        # '배송지명' 접두사 제거 후 비교 (배송지명박경아(박경아) 등)
        if text.startswith("배송지명"):
            rest = text[len("배송지명"):].strip()
            if name in rest or rest.startswith(name):
                return True
            # '이름(이름)' 형태에서 괄호 앞 이름만 비교
            if "(" in rest:
                head = rest.split("(", 1)[0].strip()
                if head == name or name in head:
                    return True
        return False

    def _parse_element_bounds(self, el):
        """요소 bounds → (x1,y1,x2,y2) 또는 None."""
        import re as _re
        try:
            bs = el.get_attribute("bounds") or ""
            mm = _re.match(r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]", bs)
            if mm:
                return tuple(map(int, mm.groups()))
            rect = el.rect
            return (rect["x"], rect["y"],
                    rect["x"] + rect["width"], rect["y"] + rect["height"])
        except Exception:
            return None

    def _iter_address_select_buttons(self):
        """배송지 목록의 '선택'/'선택됨' 후보 (WebView에서 class가 비는 경우 대비)."""
        seen = set()
        xpaths = [
            '//android.widget.Button[@text="선택" or @text="선택됨"]',
            '//*[@text="선택" or @text="선택됨"]',
        ]
        for xp in xpaths:
            try:
                els = self.driver.find_elements(By.XPATH, xp)
            except Exception:
                continue
            for el in els:
                try:
                    t = (el.get_attribute("text") or "").strip()
                    if t not in ("선택", "선택됨"):
                        continue
                    bb = self._parse_element_bounds(el)
                    key = bb if bb else id(el)
                    if key in seen:
                        continue
                    seen.add(key)
                    yield el
                except Exception:
                    continue

    def _find_card_select_button(self, name_el, recipient_name: str):
        """
        수취인 이름 View를 덮는(같은 카드) '선택' 버튼 반환.

        네이버 배송지 XML: 각 카드의 선택 버튼 bounds가 이름·전화·주소를 통째로 덮음.
        → 이름 중심점이 들어있는 선택 버튼 중 면적이 가장 작은 것 = 해당 카드.
        (리스트 전체가 아닌 카드 단위; 위쪽 다른 수취인 오선택 방지)
        """
        name_bb = self._parse_element_bounds(name_el)
        if not name_bb:
            return None
        name_cx = (name_bb[0] + name_bb[2]) // 2
        name_cy = (name_bb[1] + name_bb[3]) // 2

        all_btns = list(self._iter_address_select_buttons())
        if not all_btns:
            self._log("  ℹ 화면에서 text='선택'/'선택됨' 버튼 0개")
            return None

        containing = []  # (area, dist, btn, bb)
        named = []       # 이름 포함 확인된 후보
        for btn in all_btns:
            try:
                bb = self._parse_element_bounds(btn)
                if not bb:
                    continue
                # 이름 중심이 버튼(카드) 안에 있어야 동일 카드
                if not (bb[0] - 8 <= name_cx <= bb[2] + 8 and bb[1] - 24 <= name_cy <= bb[3] + 24):
                    continue
                area = max(1, (bb[2] - bb[0]) * (bb[3] - bb[1]))
                # 화면 전체/리스트 전체를 덮는 비정상적으로 큰 버튼 제외
                if area > 1080 * 900:
                    continue
                btn_cy = (bb[1] + bb[3]) // 2
                dist = abs(btn_cy - name_cy)
                item = (area, dist, btn, bb)
                containing.append(item)
                # 조상에 목표 이름이 있으면 가점 후보
                node = btn
                has_name = False
                for _ in range(4):
                    try:
                        node = node.find_element(By.XPATH, "..")
                        area_txt = " ".join(
                            (e.get_attribute("text") or "")
                            for e in node.find_elements(By.XPATH, ".//*")
                        )
                        sel_cnt = sum(
                            1 for e in node.find_elements(By.XPATH, ".//*")
                            if (e.get_attribute("text") or "").strip() in ("선택", "선택됨")
                        )
                        if sel_cnt > 1:
                            break
                        if self._recipient_name_matches(area_txt, recipient_name):
                            has_name = True
                            break
                    except Exception:
                        break
                if has_name:
                    named.append(item)
            except Exception:
                continue

        pool = named if named else containing
        if not pool:
            self._log(
                f"  ℹ 선택 버튼 {len(all_btns)}개 중 이름 Y({name_cy})를 덮는 카드 없음"
            )
            # 최후: 세로만 겹치는 가장 가까운 버튼 (이름 X가 카드 밖인 경우)
            nearest = None
            for btn in all_btns:
                bb = self._parse_element_bounds(btn)
                if not bb:
                    continue
                if not (bb[1] - 24 <= name_cy <= bb[3] + 24):
                    continue
                area = max(1, (bb[2] - bb[0]) * (bb[3] - bb[1]))
                if area > 1080 * 900:
                    continue
                dist = abs((bb[1] + bb[3]) // 2 - name_cy)
                cand = (area, dist, btn, bb)
                if nearest is None or cand[:2] < nearest[:2]:
                    nearest = cand
            if nearest is None:
                return None
            pool = [nearest]

        pool.sort(key=lambda x: (x[0], x[1]))
        best = pool[0][2]
        self._log(
            f"  🎯 이름 Y={name_cy} 덮는 선택 카드 채택 "
            f"(후보 {len(pool)}/{len(all_btns)}, bounds={pool[0][3]})"
        )
        return best

    def _tap_recipient_card_select(self, name_el, recipient_name: str):
        """
        이름 View → 동일 카드 '선택' 버튼 중앙 탭.
        반환: True=탭 완료, False=스크롤 후 재탐색 필요, None=버튼 못 찾음
        """
        btn = self._find_card_select_button(name_el, recipient_name)
        if btn is None:
            self._log(f"  ⚠ 이름 '{recipient_name}' 동일 카드의 선택 버튼 미발견")
            return None

        bb = self._parse_element_bounds(btn)
        if not bb:
            self._log("  ⚠ 선택 버튼 bounds 파싱 실패")
            return None

        btn_h = bb[3] - bb[1]
        if btn_h < 50:
            self._log(f"  ⚠ 선택 버튼 높이 {btn_h}px 잘림 → 스크롤 후 재탐색")
            if not self._scroll_address_list("down"):
                self._scroll_down(distance_ratio=0.15)
            time.sleep(0.5)
            return False

        tap_x = (bb[0] + bb[2]) // 2
        tap_y = (bb[1] + bb[3]) // 2
        btn_txt = (btn.get_attribute("text") or "").strip()

        try:
            scr_h = self.driver.get_window_size()["height"]
        except Exception:
            scr_h = 2400
        safe_top = int(scr_h * 0.10)
        safe_bottom = int(scr_h * 0.90)

        if tap_y < safe_top:
            self._log(f"  ⚠ 선택 버튼 Y={tap_y} 상단 밖 → 스크롤 다운")
            if not self._scroll_address_list("down"):
                self._scroll_down(distance_ratio=0.15)
            time.sleep(0.5)
            return False
        if tap_y > safe_bottom:
            self._log(f"  ⚠ 선택 버튼 Y={tap_y} 하단 밖 → 스크롤 업")
            if not self._scroll_address_list("up"):
                self._scroll_up(distance_ratio=0.18)
            time.sleep(0.5)
            return False

        self._log(
            f"  🎯 '{recipient_name}' 카드 '{btn_txt}' 탭 → ({tap_x}, {tap_y}) bounds={bb}"
        )
        self._soft_tap(tap_x, tap_y)
        time.sleep(2.5)
        return True

    def _scroll_address_list(self, direction: str = "down") -> bool:
        """배송지 목록 팝업 ListView 안에서만 스크롤. direction: down|up"""
        try:
            lists = self.driver.find_elements(By.XPATH, "//android.widget.ListView")
            best = None
            best_area = 0
            for lv in lists:
                has_sel = False
                try:
                    for e in lv.find_elements(By.XPATH, ".//*"):
                        if (e.get_attribute("text") or "").strip() in ("선택", "선택됨"):
                            has_sel = True
                            break
                except Exception:
                    pass
                if not has_sel:
                    continue
                bb = self._parse_element_bounds(lv)
                if not bb:
                    continue
                area = (bb[2] - bb[0]) * (bb[3] - bb[1])
                if area > best_area:
                    best_area = area
                    best = bb
            if not best:
                return False
            x1, y1, x2, y2 = best
            cx = (x1 + x2) // 2
            if direction == "up":
                y_from = int(y1 + (y2 - y1) * 0.35)
                y_to = int(y1 + (y2 - y1) * 0.75)
                arrow = "⬆"
            else:
                y_from = int(y1 + (y2 - y1) * 0.75)
                y_to = int(y1 + (y2 - y1) * 0.35)
                arrow = "⬇"
            self._log(f"  {arrow} 배송지 ListView 스크롤 {direction}: ({cx},{y_from})→({cx},{y_to})")
            _run_cmd(
                ["adb", "-s", self.device_id, "shell", "input", "swipe",
                 str(cx), str(y_from), str(cx), str(y_to), "350"],
                capture_output=True, timeout=5,
            )
            return True
        except Exception:
            return False

    def _scroll_address_list_up(self) -> bool:
        return self._scroll_address_list("up")

    def _check_current_delivery_address(self, recipient_name: str, phone: str) -> bool:
        """
        현재 주문/결제 화면에 목표 배송지가 이미 선택되어 있는지 확인합니다.
        (변경 버튼을 클릭하기 전에 호출)
        """
        phone_digits = ''.join(filter(str.isdigit, phone)) if phone else ""
        last4 = phone_digits[-4:] if (phone_digits and len(phone_digits) >= 4) else ""
        
        try:
            if phone_digits and len(phone_digits) >= 8:
                formatted_phone = f"{phone_digits[:3]}-{phone_digits[3:7]}-{phone_digits[7:]}"
                # '연락처' 접두사가 포함된 전화번호 텍스트를 우선 탐색
                phone_xpaths = [
                    f'//*[contains(@text, "연락처{formatted_phone}")]',
                    f'//*[contains(@text, "연락처") and contains(@text, "{last4}")]',
                    f'//*[contains(@text, "연락처{phone}")]'
                ]
                
                for px in phone_xpaths:
                    try:
                        views = self.driver.find_elements(By.XPATH, px)
                        for view in views:
                            text = view.get_attribute("text") or ""
                            if "연락처" in text:
                                # 전화번호를 찾았으므로 조상 컨테이너(최대 4단계)에서 수취인명 확인
                                node = view
                                for _ in range(4):
                                    try:
                                        node = node.find_element(By.XPATH, "..")
                                        els = node.find_elements(By.XPATH, ".//*")
                                        area_text = " ".join([e.get_attribute("text") or "" for e in els])
                                        if self._recipient_name_matches(area_text, recipient_name):
                                            return True
                                    except Exception:
                                        break
                    except Exception:
                        continue
        except Exception:
            pass
            
        # 폴백: '배송지명{수취인}' 패턴이면 이름만으로도 현재 선택으로 인정
        # (연락처가 다른 View에 있거나 마스킹되어 last4 검증이 실패하는 경우 대응)
        try:
            views = self.driver.find_elements(By.XPATH, '//*[contains(@text, "배송지명")]')
            for view in views:
                text = view.get_attribute("text") or ""
                if not self._recipient_name_matches(text, recipient_name):
                    continue
                self._log(f"  ✅ 현재 배송지명 매칭: '{text[:60]}' ← '{recipient_name}'")
                if not last4:
                    return True
                try:
                    parent = view.find_element(By.XPATH, "..")
                    area_views = parent.find_elements(By.XPATH, ".//*")
                    area_text = " ".join([av.get_attribute("text") or "" for av in area_views])
                    if last4 in area_text or "***" in area_text:
                        return True
                    # 부모에 전화가 없어도 배송지명에 수취인이 명확히 있으면 성공
                    # (결제창 '배송지명박경아(박경아)' 형태)
                    self._log(f"  ℹ 배송지명에 '{recipient_name}' 확인됨 (전화 미확인 → 이름만으로 인정)")
                    return True
                except Exception:
                    return True
        except Exception:
            pass
            
        return False

    def _visible_bounds(self):
        """사용자가 실제로 보는 화면 좌표 범위 (상태바·하단 내비 제외). ADB tap 기준."""
        w, h = self._get_window_size()
        try:
            import re as _re
            res = _run_cmd(
                ["adb", "-s", self.device_id, "shell", "wm", "size"],
                capture_output=True, timeout=3, text=True,
            )
            out = (res.stdout or "") + (res.stderr or "")
            m = _re.search(r"(\d+)\s*x\s*(\d+)", out)
            if m:
                w, h = int(m.group(1)), int(m.group(2))
        except Exception:
            pass
        left = 8
        right = max(left + 1, w - 8)
        top = max(48, int(h * 0.04))       # 상태바 아래
        bottom = min(h - 12, int(h * 0.96))  # 제스처/내비 위
        return left, top, right, bottom

    def _is_visible_coord(self, x, y) -> bool:
        """보이는 화면 안 좌표인지. 밖이면 클릭 금지."""
        try:
            x, y = int(x), int(y)
        except Exception:
            return False
        left, top, right, bottom = self._visible_bounds()
        return left <= x <= right and top <= y <= bottom

    def _soft_tap(self, x: int, y: int, duration_ms: int = 120) -> bool:
        """
        제자리 짧은 swipe(꾹 눌렀다 떼기)로 부드러운 탭을 수행합니다.
        네이버 주문/결제 WebView에서는 순간적인 'input tap'이 무시되는 경우가 있어,
        약 120ms 동안 눌렀다 떼는 방식이 훨씬 안정적으로 클릭을 인식합니다.
        사용자 가시 범위 밖 좌표는 클릭하지 않습니다.
        """
        if not self._is_visible_coord(x, y):
            left, top, right, bottom = self._visible_bounds()
            self._log(
                f"  ⏭ 화면 밖 좌표 클릭 생략: ({x}, {y}) "
                f"가시범위 x={left}~{right} y={top}~{bottom}"
            )
            return False
        _run_cmd(
            ["adb", "-s", self.device_id, "shell", "input", "swipe",
             str(int(x)), str(int(y)), str(int(x)), str(int(y)), str(duration_ms)],
            capture_output=True, timeout=5
        )
        return True

    def _find_recipient_on_screen(self, recipient_name: str, phone_digits: str) -> bool:
        """
        현재 화면(스크롤 없이)에서 수취인을 탐색하여 클릭. 찾으면 True 반환.
        """
        last4 = phone_digits[-4:] if (phone_digits and len(phone_digits) >= 4) else ""
        formatted_phone = f"{phone_digits[:3]}-{phone_digits[3:7]}-{phone_digits[7:]}" if len(phone_digits) >= 8 else ""

        # ── 방법 1: 수취인 이름 View → 동일 카드 형제 '선택' 버튼만 탭 ──
        # XML: View[text=이름] 과 Button[text=선택] 이 같은 부모 아래 형제
        # (전역 point-in-bounds / 전화 동일번호 카드 오선택 방지)
        try:
            search_xpath = f'//*[contains(@text, "{recipient_name}")]'
            views = self.driver.find_elements(By.XPATH, search_xpath)
            for view in views:
                try:
                    text = (view.get_attribute("text") or "").strip()
                    if not self._recipient_name_matches(text, recipient_name):
                        continue
                    # 결제창 뒤쪽 '배송지명…' 스킵
                    if "배송지명" in text:
                        popup_open = False
                        try:
                            popup_open = bool(self.driver.find_elements(
                                By.XPATH,
                                '//android.widget.Button[@text="선택" or @text="선택됨"]'
                            ))
                        except Exception:
                            pass
                        if not popup_open:
                            self._log(f"  ✅ 결제창 인라인 배송지명에 '{recipient_name}' 확인 → 이미 선택됨")
                            return True
                        self._log("  ℹ '배송지명' 인라인 텍스트는 팝업 뒤쪽 → 스킵")
                        continue

                    # 잘린/비정상 이름 View 스킵 (상단 일부만 보임)
                    nbb = self._parse_element_bounds(view)
                    if nbb and (nbb[3] - nbb[1]) < 20:
                        self._log(f"  ⏭ 잘린 이름 View 스킵: '{text[:40]}' h={nbb[3]-nbb[1]}")
                        continue

                    self._log(f"  🔎 이름 패턴 View 발견: '{text[:60]}'")
                    # True=탭완료, False=스크롤후재시도, None=이 View 스킵
                    tapped = self._tap_recipient_card_select(view, recipient_name)
                    if tapped is True:
                        return True
                    if tapped is False:
                        return False
                    continue
                except Exception:
                    continue
        except Exception:
            pass

        # ── 방법 2: 선택 버튼 조상에 목표 이름이 있는 카드만 탭 ──
        try:
            sel_views = list(self._iter_address_select_buttons())
            self._log(f"  ℹ 선택 버튼 후보 {len(sel_views)}개 (이름 매칭 폴백)")
            for sv in sel_views:
                try:
                    bb = self._parse_element_bounds(sv)
                    if not bb or (bb[3] - bb[1]) < 50:
                        continue
                    # 조상(최대 4단)에 목표 수취인 이름이 있고, 선택 버튼이 1개뿐인 카드만
                    matched = False
                    node = sv
                    for _ in range(4):
                        try:
                            node = node.find_element(By.XPATH, "..")
                            texts = []
                            sel_cnt = 0
                            for e in node.find_elements(By.XPATH, ".//*"):
                                t = (e.get_attribute("text") or "").strip()
                                if not t:
                                    continue
                                texts.append(t)
                                if t in ("선택", "선택됨"):
                                    sel_cnt += 1
                            if sel_cnt > 1:
                                break
                            if self._recipient_name_matches(" ".join(texts), recipient_name):
                                matched = True
                                break
                        except Exception:
                            break
                    if not matched:
                        continue

                    cx, cy = (bb[0] + bb[2]) // 2, (bb[1] + bb[3]) // 2
                    try:
                        scr_h = self.driver.get_window_size()["height"]
                    except Exception:
                        scr_h = 2400
                    if cy < int(scr_h * 0.10):
                        if not self._scroll_address_list("down"):
                            self._scroll_down(distance_ratio=0.15)
                        time.sleep(0.5)
                        return False
                    if cy > int(scr_h * 0.90):
                        if not self._scroll_address_list("up"):
                            self._scroll_up(distance_ratio=0.18)
                        time.sleep(0.5)
                        return False

                    sv_text = (sv.get_attribute("text") or "").strip()
                    self._log(
                        f"  🎯 폴백 매칭! '{recipient_name}' 카드 '{sv_text}' "
                        f"→ ({cx}, {cy}) bounds={bb}"
                    )
                    self._soft_tap(cx, cy)
                    time.sleep(2.5)
                    return True
                except Exception:
                    continue
        except Exception:
            pass

        return False

    def _select_delivery_address(self, recipient_name: str, phone: str) -> bool:
        """
        [단계 16] 배송지 선택:
        1) 화면 안정화 후 스크롤 없이 현재 화면에서 수취인 탐색
        2) 없으면 스크롤하며 재탐색 (최대 8회)

        [지원 화면 형태 - XML 참고]
        - 주문/결제 페이지 인라인 배송지: text="배송지명{이름}" 패턴
        - 별도 배송지 목록 팝업: 수취인명이 바로 text 속성에 있는 ListView
        """
        self._set_status(f"배송지 선택: {recipient_name}")
        self._log(f"🔍 배송지 선택: 수취인={recipient_name!r}, 전화={phone!r}")

        phone_digits = ''.join(filter(str.isdigit, phone)) if phone else ""
        scroll_max = 8  # 스크롤 횟수 증가 (5 → 8)

        # ── 화면 안정화 대기 (변경 버튼 클릭 후 팝업/페이지 로딩) ──
        time.sleep(1.5)

        def _try_and_verify() -> bool:
            # 안전 영역 스크롤 후 재시도 포함: 최대 3번 탐색 시도
            for _attempt in range(3):
                found = self._find_recipient_on_screen(recipient_name, phone_digits)
                if found:
                    break
                # False 반환 = 안전 영역 밖이라 스크롤하고 리턴한 경우 → 재탐색
                # (스크롤은 _find_recipient_on_screen 내부에서 이미 수행됨)
            else:
                return False

            self._log("  ⏳ 클릭 후 결제창 복귀·수취인 확인 대기 (3초)...")
            time.sleep(3.0)

            # 성공 조건: 반드시 결제 화면 '배송지명'이 목표 수취인과 일치해야 함.
            # (팝업만 닫히거나 '결제하기' 텍스트만 보이면 SONG TAO 등 엉뚱한 배송지로
            #  남아 있어도 Y 기록되던 오탐 방지)
            def _recipient_confirmed() -> bool:
                try:
                    return bool(self._check_current_delivery_address(recipient_name, phone))
                except Exception:
                    return False

            popup_still_open = False
            try:
                popup_still_open = bool(self.driver.find_elements(
                    By.XPATH, '//android.widget.Button[@text="선택"]'))
            except Exception:
                pass

            if popup_still_open:
                self._log("  ⚠ 배송지 목록 팝업이 아직 열려있음 → 선택 미완료로 판단")
            elif _recipient_confirmed():
                self._log("  ✅ 결제 화면 배송지가 목표 수취인으로 변경 확인됨")
                self._log("  ✅ 배송지 선택 성공 및 결제창 복귀 확인됨")
                return True
            else:
                # 결제창으로 돌아간 것처럼 보여도 수취인이 다르면 실패
                try:
                    views = self.driver.find_elements(By.XPATH, '//*[contains(@text, "배송지명")]')
                    shown = [(v.get_attribute("text") or "")[:60] for v in views[:5]]
                    if shown:
                        self._log(f"  ⚠ 목표 수취인 '{recipient_name}' 미확인. 현재 배송지명: {shown}")
                except Exception:
                    pass
                back_on_pay = False
                try:
                    img_dest = os.path.join(_IMG_DIR, "배송지.png")
                    if os.path.exists(img_dest) and self._find_image_coords(img_dest, threshold=0.75):
                        back_on_pay = True
                    elif os.path.exists(IMG_DELIVERY_INFO) and self._find_image_coords(IMG_DELIVERY_INFO, threshold=0.75):
                        back_on_pay = True
                    elif self.driver.find_elements(
                            By.XPATH, '//*[contains(@text, "결제하기") or contains(@text, "주문하기")]'):
                        back_on_pay = True
                except Exception:
                    pass
                if back_on_pay:
                    self._log(
                        f"  ❌ 결제창 복귀는 됐으나 수취인 '{recipient_name}' 미매칭 → 오선택으로 실패 처리"
                    )
                    return False

                self._log("  ⚠ 클릭을 시도했으나 결제창 복귀 확인 실패 (요소가 뒤쪽에 가려진 것으로 의심)")
                self._log("  👉 살짝 스크롤 업해서 요소를 앞단으로 노출 후 재탐색/클릭 시도합니다.")
                for _ in range(2):
                    self._scroll_up(distance_ratio=0.12)
                    time.sleep(0.8)

                self._log("  📋 스크롤 업 후 수취인 재탐색...")
                if _recipient_confirmed():
                    self._log("  ✅ 스크롤 후 결제창 배송지명 매칭 → 선택 완료")
                    return True
                if self._find_recipient_on_screen(recipient_name, phone_digits):
                    time.sleep(3.0)
                    if _recipient_confirmed():
                        self._log("  ✅ 재탐색 후 배송지명 매칭 확인")
                        return True
                    self._log("  ⚠ 재클릭 후에도 목표 수취인 미확인 → 실패 처리 (스크롤 후 재시도)")
                    return False
            return False

        # ── 1단계: 스크롤 없이 현재 화면에서 먼저 탐색 ──
        self._log("  📋 현재 화면에서 수취인 탐색 중 (스크롤 없음)...")
        if _try_and_verify():
            return True

        # ── 디버그: 현재 화면의 배송지명 View 목록 출력 ──
        try:
            debug_views = self.driver.find_elements(By.XPATH, '//*[contains(@text, "배송지명")]')
            if debug_views:
                self._log(f"  🔎 현재 화면 '배송지명' View 목록:")
                for dv in debug_views[:5]:
                    self._log(f"       - text='{(dv.get_attribute('text') or '')[:80]}'")
            else:
                self._log("  🔎 현재 화면에 '배송지명' View 없음 (배송지 목록 팝업 가능)")
                # 배송지 목록 팝업 화면의 ListView 존재 여부 확인
                listview = None
                try:
                    listview = self.driver.find_element(
                        By.XPATH, '//android.widget.ListView'
                    )
                    self._log("  ✅ android.widget.ListView 발견 → 배송지 목록 팝업 화면")
                except Exception:
                    self._log("  ⚠ ListView 없음")
        except Exception:
            pass

        # ── 2단계: 없으면 스크롤하며 재탐색 ──
        for scroll_cnt in range(1, scroll_max + 1):
            self._log(f"  ⬇ 배송지 목록 스크롤 ({scroll_cnt}/{scroll_max})")
            if not self._scroll_address_list("down"):
                self._scroll_down()
            time.sleep(1.0)

            self._log(f"  📋 스크롤 후 수취인 재탐색...")
            if _try_and_verify():
                return True

        self._log(f"  ❌ 배송지 '{recipient_name}' 매칭 실패")
        return False

    # ─── 단계 16.5: 배송메모 '선택안함' 처리 ─────────────────────────────────

    def _handle_delivery_memo(self) -> None:
        """
        [단계 16.5] 배송지 선택 직후 배송메모 처리:
        '배송메모 선택' 팝업이 떠 있는 경우에만 '선택 안 함'(선택안함.png)을
        이미지 인식으로 찾아 1회 클릭합니다.
        (배송메모 필드를 직접 클릭해서 팝업을 여는 동작은 하지 않음)
        """
        self._set_status("배송메모 처리")
        self._log("🔍 [배송메모] '선택 안 함' 이미지 인식 시도...")
        time.sleep(1.0)

        if not os.path.exists(IMG_MEMO_NO_SELECT):
            self._log("  ℹ [배송메모] 선택안함.png 템플릿 없음 → 건너뛰고 다음 작업 진행")
            return

        # '선택 안 함' 이미지 인식 (최대 3회 재시도)
        for ns_try in range(1, 4):
            ns_coords = self._find_image_coords(IMG_MEMO_NO_SELECT, threshold=0.60)
            if ns_coords:
                self._log(f"  🎯 [선택 안 함] 이미지 발견! 좌표 ({ns_coords[0]}, {ns_coords[1]}) -> 1회 탭")
                ah.tap_by_coords(self.driver, ns_coords[0], ns_coords[1], self._log)
                time.sleep(1.2)
                self._log("✅ [배송메모] '선택 안 함' 선택 완료 → 다음 작업 진행")
                return
            if ns_try < 3:
                time.sleep(1.0)

        self._log("  ℹ [배송메모] '선택 안 함' 미감지 → 건너뛰고 다음 작업 진행")

    # ─── 단계 17: 전액사용 클릭 ──────────────────────────────────────────────

    def _click_full_use(self) -> bool:
        """
        [단계 17] 전액사용.png 인식될 때까지 아래로 부드럽게 미세 스크롤하며 대기 -> 인식 후 클릭 -> 3초 대기
        """
        self._set_status("전액사용 탐색 중")
        self._log("🔍 [단계 17] 전액사용 버튼 탐색 시작 (부드러운 미세 스크롤 탐색)")

        w_h = 2400
        try:
            w_h = self.driver.get_window_size()['height']
        except Exception:
            pass

        min_y_full_use = int(w_h * 0.25)  # 상단 헤더/툴바 오탐지 방지 (Y >= 25% 영역)
        max_y_full_use = int(w_h * 0.65)  # 하단 고정 결제버튼(플로팅 바) 가림 방지 (Y <= 65% 영역)
        max_scroll_attempts = 15

        for attempt in range(1, max_scroll_attempts + 1):
            # 1. 전액사용.png 이미지 매칭
            if os.path.exists(IMG_FULL_USE):
                coords = self._find_image_coords(IMG_FULL_USE, threshold=0.80, min_y=min_y_full_use, max_y=max_y_full_use)
                if coords:
                    cx, cy = coords
                    self._log(f"  🎯 전액사용.png 이미지 발견! 좌표 ({cx}, {cy}) -> 탭 클릭")
                    ah.tap_by_coords(self.driver, cx, cy, self._log)
                    self._log("✅ [단계 17] 전액사용 이미지 인식 클릭 완료 (3초 대기)")
                    time.sleep(2)
                    return True

            # 2. XPath 텍스트 매칭 폴백 ("전액사용", "전액 사용", "전액")
            full_use_xpaths = [
                '//android.widget.Button[@text="전액사용"]',
                '//android.widget.Button[contains(@text,"전액")]',
                '//android.view.View[@text="전액사용"]',
                '//android.view.View[contains(@text,"전액사용")]',
                '//*[@text="전액사용"]',
                '//*[contains(@text,"전액사용")]',
                '//*[contains(@text,"전액 사용")]',
            ]
            for xpath in full_use_xpaths:
                try:
                    if ah.element_exists(self.driver, xpath, timeout=1):
                        els = self.driver.find_elements(By.XPATH, xpath)
                        for el in els:
                            rect = el.rect
                            cy = rect['y'] + rect['height'] // 2
                            if min_y_full_use <= cy <= max_y_full_use:
                                self._log(f"  🎯 전액사용 XPath 발견: {xpath} (y={cy}) -> 클릭 시도")
                                if self._safe_click_element(el):
                                    self._log("✅ [단계 17] 전액사용 XPath 클릭 완료 (3초 대기)")
                                    time.sleep(2)
                                    return True
                except Exception:
                    continue

            # 3. 미발견 시 부드럽게 미세 스크롤 다운
            self._log(f"  ⬇ [단계 17] 전액사용 미발견 또는 Y범위 밖 -> 미세 스크롤 다운 ({attempt}/{max_scroll_attempts})")
            self._scroll_down(distance_ratio=0.18)
            time.sleep(0.8)

        self._log("  ❌ 전액사용 버튼 탐색 실패 (최대 스크롤 초과)")
        return False


    # ─── 단계 18: 결제하기 버튼 ──────────────────────────────────────────────

    def _click_pay_button(self) -> bool:
        """[단계 18] 결제하기 버튼 클릭, 5초 대기"""
        self._set_status("결제하기 클릭")

        if self._skip_final_order_click():
            mode = "수동시작" if self.manual_mode else "테스트 모드"
            self._log(f"🖐 [{mode}] 결제하기 버튼 클릭 생략 (성공 처리)")
            return True

        if ah.element_exists(self.driver, PAY_BTN_XPATH, timeout=5):
            ah.wait_and_click(self.driver, PAY_BTN_XPATH, timeout=5, log_callback=self._log)
            self._log("✅ 결제하기 버튼 클릭 완료")
            time.sleep(3)
            return True

        # 폴백: 좌표
        try:
            size = self.driver.get_window_size()
            w, h = size['width'], size['height']
            tap_x = int(w * 0.5)
            tap_y = int(h * 0.92)
            ah.tap_by_coords(self.driver, tap_x, tap_y, self._log)
            self._log(f"  ✅ 결제하기 좌표 탭 ({tap_x}, {tap_y})")
            time.sleep(3)
            return True
        except Exception as e:
            self._log(f"  ❌ 결제하기 클릭 실패: {e}")
            return False

    # ─── 단계 19: 비밀번호 입력 ──────────────────────────────────────────────

    def _find_digit_coords(self, img_path: str, min_y: int,
                           screenshot_png: Optional[bytes] = None,
                           max_y: Optional[int] = None,
                           min_x: Optional[int] = None,
                           max_x: Optional[int] = None,
                           prefer_xy: Optional[tuple] = None,
                           prefer_radius: int = 220,
                           min_score: float = 0.50) -> Optional[tuple]:
        """
        숫자 키패드 이미지 인식 (경량화 버전).
        - 스케일 20단계 (0.6~1.8), UI 과부하 방지
        - screenshot_png를 외부에서 주입하면 재캡처 생략 (재시도 성능 개선)
        - prefer_xy가 있으면 예상 좌표 근처 매칭을 우선
        - min_x/max_x/min_y/max_y 로 ROI 제한 (현대비번 커팅 영역)
        """
        try:
            import cv2
            import numpy as np
            from PIL import Image
            import io
        except ImportError:
            return None

        try:
            if screenshot_png is None:
                screenshot_png = self._get_screenshot()
            screenshot_pil = Image.open(io.BytesIO(screenshot_png))
            screen_bgr = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)
            screen_gray = cv2.cvtColor(screen_bgr, cv2.COLOR_BGR2GRAY)
            screen_h, screen_w = screen_gray.shape

            # 키패드 영역 외 차단
            masked = screen_gray.copy()
            if min_y > 0:
                masked[:min_y, :] = 0
            if max_y is not None and 0 < max_y < screen_h:
                masked[max_y:, :] = 0
            if min_x is not None and min_x > 0:
                masked[:, :min_x] = 0
            if max_x is not None and 0 < max_x < screen_w:
                masked[:, max_x:] = 0

            template_bgr = cv2.imdecode(
                np.fromfile(img_path, dtype=np.uint8), cv2.IMREAD_COLOR
            )
            if template_bgr is None:
                return None
            template_gray = cv2.cvtColor(template_bgr, cv2.COLOR_BGR2GRAY)
            t_h, t_w = template_gray.shape

            best_score, best_loc, best_tw, best_th = -1, None, t_w, t_h
            best_near_score, best_near_loc, best_near_tw, best_near_th = -1, None, t_w, t_h
            # 키패드 버튼은 크기 편차가 작으므로 0.6~1.8 범위, 20단계만 탐색
            for scale in np.linspace(0.6, 1.8, 20):
                nw = int(t_w * scale)
                nh = int(t_h * scale)
                if nw >= screen_w or nh >= screen_h or nw < 8 or nh < 4:
                    continue
                resized = cv2.resize(template_gray, (nw, nh), interpolation=cv2.INTER_AREA)
                result = cv2.matchTemplate(masked, resized, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)
                if max_val > best_score:
                    best_score, best_loc, best_tw, best_th = max_val, max_loc, nw, nh
                if prefer_xy is not None:
                    # 예상 좌표 주변 결과맵에서 최고점 탐색
                    px, py = prefer_xy
                    # result 좌표 = 템플릿 좌상단 → 중심 기준 역산
                    rx = int(px - nw // 2)
                    ry = int(py - nh // 2)
                    rh, rw = result.shape
                    pad = max(40, prefer_radius // 2)
                    x1 = max(0, rx - pad)
                    y1 = max(0, ry - pad)
                    x2 = min(rw, rx + pad + 1)
                    y2 = min(rh, ry + pad + 1)
                    if x2 > x1 and y2 > y1:
                        region = result[y1:y2, x1:x2]
                        _, near_val, _, near_loc = cv2.minMaxLoc(region)
                        abs_loc = (near_loc[0] + x1, near_loc[1] + y1)
                        if near_val > best_near_score:
                            best_near_score = near_val
                            best_near_loc = abs_loc
                            best_near_tw, best_near_th = nw, nh

            # 예상 좌표 근처 매칭이 충분하면 그쪽 우선
            if prefer_xy is not None and best_near_loc is not None and best_near_score >= max(0.45, min_score - 0.08):
                cx = best_near_loc[0] + best_near_tw // 2
                cy = best_near_loc[1] + best_near_th // 2
                self._log(
                    f"    🎯 [숫자인식] 점수: {best_near_score:.4f} → 좌표 ({cx}, {cy})"
                    f" (예상근처 {prefer_xy})"
                )
                return cx, cy

            if best_score >= min_score and best_loc is not None:
                cx = best_loc[0] + best_tw // 2
                cy = best_loc[1] + best_th // 2
                if prefer_xy is not None:
                    dist = ((cx - prefer_xy[0]) ** 2 + (cy - prefer_xy[1]) ** 2) ** 0.5
                    if dist > prefer_radius * 1.6:
                        self._log(
                            f"    ↩ [숫자인식] 후보 거부 ({cx},{cy}) "
                            f"점수 {best_score:.4f} — 예상 {prefer_xy}과 거리 {dist:.0f}"
                        )
                        return None
                self._log(f"    🎯 [숫자인식] 점수: {best_score:.4f} → 좌표 ({cx}, {cy})")
                return cx, cy

            self._log(f"    ↩ [숫자인식] 미발견 (점수 {best_score:.4f} < {min_score:.2f})")
            return None
        except Exception as e:
            self._log(f"    [숫자인식 오류] {e}")
            return None

    def _input_password(self, password: str) -> bool:
        """
        [단계 19] 비밀번호 입력
        숫자 폴더의 p0~p9.png 이미지로 각 자리 숫자 버튼의 위치를 인식하여 순서대로 클릭.
        - 입력 전 3초 대기 (키패드 완전 표시 보장)
        - 각 숫자: 최대 3회 재시도, 첫 시도 실패 시 새 스크린샷으로 재시도
        - 경량화된 20스케일 매칭 (UI 과부하 방지)
        """
        self._set_status("비밀번호 입력")

        if self._skip_final_order_click():
            mode = "수동시작" if self.manual_mode else "테스트 모드"
            self._log(f"🖐 [{mode}] 비밀번호 입력 생략 (강제 성공)")
            return True

        pwd_digits = ''.join(filter(str.isdigit, password))

        if not pwd_digits:
            self._log("  ⚠ 비밀번호 없음 → 건너뜀")
            return True

        self._log(f"  🔐 비밀번호 입력: {'*' * len(pwd_digits)}자리")

        w_h = 2400
        try:
            w_h = self.driver.get_window_size()['height']
        except Exception:
            pass

        min_y_keypad = int(w_h * 0.45)  # 비밀번호 키패드는 화면 하단에만 존재

        # ── 비밀번호 키패드 완전 표시 대기 (3초) ──
        self._log("  ⏳ 키패드 완전 표시 대기 (3초)...")
        time.sleep(3.0)

        MAX_DIGIT_RETRY = 3   # 숫자 1개당 최대 재시도 횟수 (UI 과부하 방지)
        RETRY_INTERVAL  = 0.8 # 재시도 간격 (초)

        all_success = True
        for idx, digit in enumerate(pwd_digits):
            img_path = IMG_NUMS.get(digit)
            if not img_path or not os.path.exists(img_path):
                self._log(f"  ⚠ 숫자 이미지 없음: p{digit}.png → 실패")
                all_success = False
                break

            self._log(f"  🔢 {idx+1}번째 자리 '{digit}' 클릭 시도 (최대 {MAX_DIGIT_RETRY}회)")
            digit_ok = False
            current_screenshot = None  # 첫 시도: None → _find_digit_coords 내부에서 캡처
            for retry in range(1, MAX_DIGIT_RETRY + 1):
                coords = self._find_digit_coords(img_path, min_y_keypad,
                                                  screenshot_png=current_screenshot)
                if coords:
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    self._log(f"    ✅ '{digit}' 클릭 완료 ({coords}) [시도 {retry}회]")
                    digit_ok = True
                    time.sleep(0.6)
                    break
                else:
                    self._log(f"    ↩ '{digit}' 인식 실패 ({retry}/{MAX_DIGIT_RETRY}) → {RETRY_INTERVAL}초 후 재시도")
                    time.sleep(RETRY_INTERVAL)
                    current_screenshot = None  # 재시도 시 새 스크린샷 캡처

            if not digit_ok:
                self._log(f"    ❌ '{digit}' 최종 인식 실패 (하단 키패드 영역, {MAX_DIGIT_RETRY}회 모두 실패)")
                all_success = False
                break

        if all_success:
            self._log("  ✅ 비밀번호 입력 완료")
            return True
        else:
            self._log("  ❌ 비밀번호 입력 실패 (일부 숫자 인식 불가)")
            return False

    def _capture_and_log_bank_transfer(self, row: OrderRow) -> bool:
        """무통장 결제 완료 후 스크린샷 캡쳐 및 주문번호/계좌번호 로그 저장.
        주문번호가 확인되면 True, 미확인 시 False 반환."""
        self._log("📸 무통장입금 완료 화면 캡쳐 및 정보 추출 대기 중...")

        import re
        order_num_xpath = '//android.widget.Button[contains(@text, "복사하기")]'
        bank_xpath = '//android.widget.Button[contains(@text, "은행")]'

        loaded = False
        for _ in range(10):
            if ah.element_exists(self.driver, order_num_xpath, timeout=1) or ah.element_exists(self.driver, bank_xpath, timeout=1):
                loaded = True
                break
            time.sleep(1)

        if not loaded:
            self._log("⚠ 무통장 결제 완료 화면 요소 확인 지연. 스크린샷 캡쳐를 진행합니다.")

        time.sleep(2)

        today_str = time.strftime("%Y%m%d")
        save_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "캡쳐", "무통장", today_str)
        os.makedirs(save_dir, exist_ok=True)

        timestamp = time.strftime("%Y%m%d_%H%M%S")
        safe_keyword = re.sub(r'[\\/*?"<>|]', "", row.search_keyword)
        screenshot_filename = f"무통장_{row.row_index}행_{timestamp}_{safe_keyword}.png"
        screenshot_path = os.path.join(save_dir, screenshot_filename)

        try:
            screenshot_data = self._get_screenshot()
            with open(screenshot_path, "wb") as f:
                f.write(screenshot_data)
            self._log(f"✅ 무통장 스크린샷 저장 완료: {screenshot_path}")
        except Exception as e:
            self._log(f"❌ 스크린샷 캡쳐 실패: {e}")

        order_number_text = ""
        bank_info_text = "알수없음"

        # ── 주문번호 추출 (복사하기 버튼 옆 텍스트) ──
        try:
            els = self.driver.find_elements(By.XPATH, order_num_xpath)
            if els:
                text = els[0].text or ""
                extracted = text.replace("복사하기", "").strip()
                if extracted:
                    order_number_text = extracted
        except Exception as e:
            self._log(f"⚠ 주문번호 추출 실패: {e}")

        # ── 주문번호 미추출 시 전체 화면 텍스트에서 패턴 탐색 ──
        if not order_number_text:
            try:
                all_els = self.driver.find_elements(By.XPATH, '//*[@text]')
                for el in all_els:
                    t = (el.get_attribute("text") or "").strip()
                    # 주문번호 패턴: 숫자로만 구성된 10자리 이상
                    if re.fullmatch(r'\d{10,}', t):
                        order_number_text = t
                        self._log(f"  🔍 주문번호 패턴 추출: {order_number_text}")
                        break
            except Exception as e:
                self._log(f"⚠ 주문번호 패턴 탐색 실패: {e}")

        # ── 계좌정보 추출 ──
        try:
            els = self.driver.find_elements(By.XPATH, bank_xpath)
            for el in els:
                text = el.text or ""
                if any(char.isdigit() for char in text) and len(text) > 5:
                    bank_info_text = text.strip()
                    break
        except Exception as e:
            self._log(f"⚠ 계좌번호 추출 실패: {e}")

        # ── 로그 파일 기록 ──
        order_display = order_number_text if order_number_text else "알수없음"
        log_filename = "무통장로그.txt"
        log_path = os.path.join(save_dir, log_filename)
        log_content = f"[{timestamp}] [순번: {row.row_index}행] 키워드: {row.search_keyword} | 주문번호: {order_display} | 계좌정보: {bank_info_text}\n"

        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(log_content)
            self._log(f"✅ 무통장 정보 로그 기록 완료 (주문번호: {order_display}, 계좌정보: {bank_info_text})")
        except Exception as e:
            self._log(f"❌ 무통장 로그 기록 실패: {e}")

        # ── 주문번호 확인 여부로 성공/실패 반환 ──
        if order_number_text:
            self._log(f"✅ [무통장 성공 확인] 주문번호: {order_number_text}")
            return True
        else:
            self._log("❌ [무통장 실패] 주문번호 미확인 → 성공으로 처리하지 않음")
            return False


    def _click_image_with_scroll(self, img_path: str, name: str, threshold: float = 0.82, max_scroll_attempts: int = 15,
                                 min_x: Optional[int] = None, max_x: Optional[int] = None,
                                 min_y: Optional[int] = None, max_y: Optional[int] = None) -> bool:
        """지정된 이미지를 미세 스크롤하며 찾고, 화면 중앙 영역에 정렬하여 클릭"""
        self._set_status(f"{name} 탐색 중")
        self._log(f"🔍 {name} 버튼 탐색 시작 (미세 스크롤 탐색, 최대 {max_scroll_attempts}회 시도)")

        w_h = 2400
        try:
            w_h = self.driver.get_window_size()['height']
        except Exception:
            pass

        mid_top    = int(w_h * 0.35)
        mid_bottom = int(w_h * 0.65)

        for attempt in range(1, max_scroll_attempts + 1):
            if os.path.exists(img_path):
                coords = self._find_image_coords(img_path, threshold=threshold, min_x=min_x, max_x=max_x, min_y=min_y, max_y=max_y)
                if coords:
                    if coords[1] < mid_top:
                        self._log(f"  📌 {name} 상단 치우침(y={coords[1]}) -> 미세 스크롤 업")
                        self._scroll_up(distance_ratio=0.18)
                        time.sleep(1.0)
                        adj = self._find_image_coords(img_path, threshold=threshold, min_x=min_x, max_x=max_x, min_y=min_y, max_y=max_y)
                        if adj: coords = adj
                    elif coords[1] > mid_bottom:
                        self._log(f"  📌 {name} 하단 치우침(y={coords[1]}) -> 미세 스크롤 다운")
                        self._scroll_down(distance_ratio=0.18)
                        time.sleep(1.0)
                        adj = self._find_image_coords(img_path, threshold=threshold, min_x=min_x, max_x=max_x, min_y=min_y, max_y=max_y)
                        if adj: coords = adj

                    self._log(f"  🎯 {name} 이미지 발견! 화면 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    time.sleep(2)
                    return True

            self._log(f"  ⬇ {name} 미발견 -> 미세 스크롤 다운 ({attempt}/{max_scroll_attempts})")
            self._scroll_down(distance_ratio=0.20)
            time.sleep(0.8)
        self._log(f"  ❌ {name} 버튼 탐색 실패")
        return False
        
    def _click_any_image_with_scroll(self, images: list, threshold: float = 0.82, max_scroll_attempts: int = 15,
                                     min_x: Optional[int] = None, max_x: Optional[int] = None,
                                     min_y: Optional[int] = None, max_y: Optional[int] = None) -> bool:
        """여러 이미지 중 하나라도 발견되면 미세 스크롤 조정 후 클릭"""
        names_str = " / ".join(n for _, n in images)
        self._set_status(f"{names_str} 탐색 중")
        self._log(f"🔍 [{names_str}] 중 하나 탐색 시작 (미세 스크롤 탐색, 최대 {max_scroll_attempts}회 시도)")

        w_h = 2400
        try:
            w_h = self.driver.get_window_size()['height']
        except Exception:
            pass

        mid_top    = int(w_h * 0.35)
        mid_bottom = int(w_h * 0.65)

        for attempt in range(1, max_scroll_attempts + 1):
            for img_path, name in images:
                if os.path.exists(img_path):
                    coords = self._find_image_coords(img_path, threshold=threshold, min_x=min_x, max_x=max_x, min_y=min_y, max_y=max_y)
                    if coords:
                        if coords[1] < mid_top:
                            self._log(f"  📌 {name} 상단 치우침(y={coords[1]}) -> 미세 스크롤 업")
                            self._scroll_up(distance_ratio=0.18)
                            time.sleep(1.0)
                            adj = self._find_image_coords(img_path, threshold=threshold, min_x=min_x, max_x=max_x, min_y=min_y, max_y=max_y)
                            if adj: coords = adj
                        elif coords[1] > mid_bottom:
                            self._log(f"  📌 {name} 하단 치우침(y={coords[1]}) -> 미세 스크롤 다운")
                            self._scroll_down(distance_ratio=0.18)
                            time.sleep(1.0)
                            adj = self._find_image_coords(img_path, threshold=threshold, min_x=min_x, max_x=max_x, min_y=min_y, max_y=max_y)
                            if adj: coords = adj

                        self._log(f"  🎯 {name} 이미지 발견! 화면 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                        ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                        time.sleep(2)
                        return True

            self._log(f"  ⬇ [{names_str}] 미발견 -> 미세 스크롤 다운 ({attempt}/{max_scroll_attempts})")
            self._scroll_down(distance_ratio=0.20)
            time.sleep(0.8)
        self._log(f"  ❌ [{names_str}] 버튼 모두 탐색 실패")
        return False

    def _click_image_basic(self, img_path: str, name: str, threshold: float = 0.82) -> bool:
        """지정된 이미지를 스크롤 없이 한 번만 찾아서 클릭 (또는 짧게 대기하며 재시도)"""
        self._set_status(f"{name} 탐색 중")
        for _ in range(3):
            if os.path.exists(img_path):
                coords = self._find_image_coords(img_path, threshold=threshold)
                if coords:
                    self._log(f"  🎯 {name} 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    time.sleep(2)
                    return True
            time.sleep(1)
        self._log(f"  ❌ {name} 버튼 탐색 실패 (스크롤 없음)")
        return False

    def _click_any_image_basic(self, images: list, threshold: float = 0.70,
                               attempts: int = 5, wait_after: float = 2.0,
                               min_y: Optional[int] = None,
                               max_y: Optional[int] = None,
                               min_x: Optional[int] = None,
                               max_x: Optional[int] = None) -> bool:
        """여러 이미지 중 하나라도 발견되면 스크롤 없이 클릭."""
        names_str = " / ".join(n for _, n in images)
        self._set_status(f"{names_str} 탐색 중")
        valid = [(p, n) for p, n in images if os.path.exists(p)]
        if not valid:
            self._log(f"  ❌ [{names_str}] 템플릿 파일 없음")
            return False
        for attempt in range(1, attempts + 1):
            for img_path, name in valid:
                coords = self._find_image_coords(
                    img_path, threshold=threshold,
                    min_x=min_x, max_x=max_x, min_y=min_y, max_y=max_y,
                )
                if coords:
                    self._log(f"  🎯 {name} 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    time.sleep(wait_after)
                    return True
            if attempt < attempts:
                time.sleep(1.0)
        self._log(f"  ❌ [{names_str}] 버튼 탐색 실패 (스크롤 없음, {attempts}회)")
        return False

    def _click_hyundai_safe_confirm(self, attempts: int = 6, threshold: float = 0.60,
                                    force_tap: bool = False) -> bool:
        """
        안전결재3 화면 기준: 중앙 화이트 팝업 하단 '확인' 클릭.
        기대 좌표 ≈ (544, 1253). 상단(y=316) 오인 차단.
        """
        w, h = 1080, 2400
        try:
            size = self.driver.get_window_size()
            w, h = size["width"], size["height"]
        except Exception:
            pass

        # 팝업 확인 버튼 ROI (안전결재3 기준 중하단)
        min_y = int(h * 0.42)
        max_y = int(h * 0.72)
        min_x = int(w * 0.20)
        max_x = int(w * 0.80)
        expect_x, expect_y = w // 2, int(h * 0.52)  # ≈544,1250 @1080x2400

        # 1) 팝업 본체(안전결재/안전한3) bbox → 하단 중앙 탭 (확인 위치)
        for img_path, name in IMG_HYUNDAI_SAFE_POPUP_BODY:
            if not os.path.exists(img_path):
                continue
            box = self._find_image_bbox(
                img_path, threshold=0.50, save_crop=True, crop_label=f"팝업_{name}"
            )
            if not box:
                continue
            # 확인은 모달 하단 ~85~92% 지점
            tap_x = (box["x1"] + box["x2"]) // 2
            tap_y = int(box["y1"] + (box["y2"] - box["y1"]) * 0.88)
            if not (min_y - int(h * 0.05) <= tap_y <= max_y + int(h * 0.08)):
                # bbox가 전체화면(안전결재 축소본)일 수 있음 → 기대좌표 사용
                tap_x, tap_y = expect_x, expect_y
            self._log(
                f"  🎯 [안전확인] 팝업본체 '{name}' bbox → 확인 탭 ({tap_x},{tap_y})"
            )
            ah.tap_by_coords(self.driver, tap_x, tap_y, self._log)
            time.sleep(2.0)
            return True

        # 2) 안전확인1~3 이미지 (중하단 ROI만)
        self._log(
            f"  🔍 [안전확인] 중하단 ROI 탐색 "
            f"x={min_x}~{max_x} y={min_y}~{max_y} (기대≈{expect_x},{expect_y})"
        )
        if self._click_any_image_basic(
            IMG_HYUNDAI_SAFE_CONFIRM,
            threshold=threshold,
            attempts=attempts,
            wait_after=2.0,
            min_x=min_x, max_x=max_x, min_y=min_y, max_y=max_y,
        ):
            return True

        # 3) XPath: 팝업 문구 근처 / 중하단 '확인'
        for xp in SAFE_AUTH_TEXT_XPATHS:
            try:
                if ah.element_exists(self.driver, xp, timeout=1):
                    # 같은 계층/근처의 확인 버튼
                    parent_xp = xp + '/ancestor::*[1]//*[contains(@text,"확인")]'
                    for c_xp in [parent_xp, '//android.widget.Button[contains(@text,"확인")]',
                                 '//*[contains(@text,"확인")]']:
                        try:
                            els = self.driver.find_elements(By.XPATH, c_xp)
                            for el in els:
                                loc = el.location
                                sz = el.size
                                cx = int(loc["x"] + sz["width"] / 2)
                                cy = int(loc["y"] + sz["height"] / 2)
                                if not (min_x <= cx <= max_x and min_y <= cy <= max_y):
                                    self._log(f"  ↩ [안전확인] XPath 거부 ({cx},{cy})")
                                    continue
                                if self._safe_click_element(el):
                                    self._log(f"  ✅ [안전확인] XPath 클릭 ({cx},{cy})")
                                    time.sleep(2.0)
                                    return True
                        except Exception:
                            continue
            except Exception:
                continue

        # 4) 팝업이 확인된 경우에만 안전결재3 기준 강제 탭
        if force_tap:
            self._log(f"  ⚠ [안전확인] 인식 실패 → 안전결재3 기준 강제 탭 ({expect_x},{expect_y})")
            ah.tap_by_coords(self.driver, expect_x, expect_y, self._log)
            time.sleep(2.0)
            return True
        return False

    def _hyundai_safe_detect_visible(self, threshold: float = 0.60) -> Optional[str]:
        """안전한/안전결재/추가인증 또는 텍스트로 팝업 감지."""
        for img_path, name in IMG_HYUNDAI_SAFE_DETECT:
            if not os.path.exists(img_path):
                continue
            coords = self._find_image_coords(img_path, threshold=threshold)
            if coords:
                return name
        # 전체화면 참고 템플릿
        if os.path.exists(IMG_HYUNDAI_SAFE_POPUP_FULL):
            box = self._find_image_bbox(
                IMG_HYUNDAI_SAFE_POPUP_FULL, threshold=0.40, save_crop=False, crop_label="안전결재3"
            )
            if box and box.get("score", 0) >= 0.40:
                return "안전결재3"
        # 텍스트
        for xp in SAFE_AUTH_TEXT_XPATHS:
            try:
                if ah.element_exists(self.driver, xp, timeout=0.8):
                    return "텍스트팝업"
            except Exception:
                continue
        return None

    def _handle_hyundai_safe_auth_popup(self) -> bool:
        """
        [22-10 이후] 현대결제하기 → 3초 대기 →
        안전결재3 팝업 확인 1회 클릭 후 바로 현대카드비번 단계로 진행.
        (잔존 감지/재시도 생략 — 안전한2 상단 오인으로 루프 방지)
        """
        self._log("  ⏳ [안전인증] 현대결제하기 후 3초 대기...")
        time.sleep(3.0)

        detect_names = " / ".join(n for _, n in IMG_HYUNDAI_SAFE_DETECT) + " / 안전결재3"
        confirm_names = " / ".join(n for _, n in IMG_HYUNDAI_SAFE_CONFIRM)

        detected = self._hyundai_safe_detect_visible(threshold=0.55)
        if detected:
            self._log(f"  🔒 [안전인증] 팝업 감지: '{detected}' → 확인 클릭")
        else:
            self._log(f"  🔍 [안전인증] 감지 미매칭 ({detect_names}) → 안전확인 1회 시도")

        self._log(f"  🔍 [안전인증] 팝업닫기: {confirm_names}")
        clicked = self._click_hyundai_safe_confirm(
            attempts=5, threshold=0.58, force_tap=bool(detected)
        )
        if clicked:
            self._log("  ✅ [안전인증] 확인 클릭 완료 → 현대카드비번 단계로 진행 (잔존검사 패스)")
        else:
            self._log("  ℹ [안전인증] 확인 미클릭/팝업없음 → 현대카드비번 단계로 진행")
        time.sleep(1.0)
        return True

    def _click_bank_select_with_scroll(self, max_scroll_attempts: int = 8) -> bool:
        """
        [무통장입금 클릭 후 판단]
        '은행을.png' 탐색 (최대 8회 스크롤)
        '은행을.png' 이미지 인식 시 화면 중앙 조절 후 탭 클릭.
        """
        self._set_status("은행 선택 탐색 및 판단 중")
        self._log(f"🔍 [무통장입금 클릭 후 판단] 은행 선택 버튼 탐색 시작 ('은행을.png', 최대 {max_scroll_attempts}회 시도)")

        bank_images = [
            (IMG_SELECT_BANK,  "은행을"),
            (IMG_SELECT_BANK2, "은행을2"),
            (IMG_SELECT_BANK3, "은행을3"),
            (IMG_SELECT_BANK4, "은행을4"),
        ]
        bank_xpaths = [
            '//android.widget.Button[contains(@text,"은행")]',
            '//android.view.View[contains(@text,"은행")]',
            '//*[contains(@text,"은행선택")]',
            '//*[contains(@text,"은행을")]',
            '//*[contains(@content-desc,"은행")]',
        ]

        w_h = 2400
        try:
            w_h = self.driver.get_window_size()['height']
        except Exception:
            pass

        mid_top = int(w_h * 0.35)
        mid_bottom = int(w_h * 0.65)

        for attempt in range(1, max_scroll_attempts + 1):
            # 1. 이미지 매칭 (threshold 0.70)
            for img_path, name in bank_images:
                if os.path.exists(img_path):
                    coords = self._find_image_coords(img_path, threshold=0.70)
                    if coords:
                        if coords[1] < mid_top:
                            self._log(f"  📌 {name} 상단 치우침 -> 미세 스크롤 업")
                            self._scroll_up(distance_ratio=0.18)
                            time.sleep(1.0)
                            adj = self._find_image_coords(img_path, threshold=0.70)
                            if adj: coords = adj
                        elif coords[1] > mid_bottom:
                            self._log(f"  📌 {name} 하단 치우침 -> 미세 스크롤 다운")
                            self._scroll_down(distance_ratio=0.18)
                            time.sleep(1.0)
                            adj = self._find_image_coords(img_path, threshold=0.70)
                            if adj: coords = adj

                        self._log(f"  🎯 {name} 이미지 발견! 화면 중앙 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭 및 존재 확인 성공")
                        ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                        time.sleep(2)
                        return True

            # 2. XPath 매칭 폴백
            for xpath in bank_xpaths:
                try:
                    if ah.element_exists(self.driver, xpath, timeout=1):
                        self._log(f"  🎯 은행 선택 XPath 발견: {xpath} -> 클릭 및 존재 확인 성공")
                        el = self.driver.find_element(By.XPATH, xpath)
                        if self._safe_click_element(el):
                            time.sleep(2)
                            return True
                except Exception:
                    continue

            self._log(f"  ⬇ '은행을' 미발견 -> 미세 스크롤 다운 ({attempt}/{max_scroll_attempts})")
            self._scroll_down(distance_ratio=0.20)
            time.sleep(0.8)

        self._log(f"  ⚠ 무통장입금 클릭 후 '은행을' 미발견 (최대 {max_scroll_attempts}회 시도 초과). 무시하고 계속 진행합니다.")
        return True

    def _click_other_pay_button(self, max_scroll_attempts: int = 20) -> bool:
        """
        다른결재 버튼을 3중 인식 방식으로 탐색 후 화면 중앙에 안착시켜 클릭합니다.
        0순위: 최우선 XPath (btn_payment_method_accordion 등)
        1순위: 이미지 매칭 (다른결재.png, 다른결재수단2.png 등)
        2순위: OCR (pytesseract)
        3순위: XPath 텍스트 탐색
        """
        self._set_status("다른 결제수단 탐색 중")
        self._log("🔍 [다른결재 버튼] 이미지/OCR/XPath 3중 탐색 시작")

        other_pay_keywords = [
            "다른결재수단", "다른 결재수단", "다른결제수단", "다른 결제수단",
            "다른결재", "다른 결재", "결제수단보기",
            "다른결재4", "보기"
        ]

        img_candidates = []
        for img_path, name in [
            (IMG_OTHER_PAY,  "다른결재"),
            (IMG_OTHER_PAY2, "다른결재수단2"),
            (IMG_OTHER_PAY4, "다른결재4"),
            (IMG_BOGI,       "보기"),
        ]:
            if os.path.exists(img_path):
                img_candidates.append((img_path, name))

        w_h = 2400
        w_w = 1080
        try:
            sz = self.driver.get_window_size()
            w_h, w_w = sz['height'], sz['width']
        except Exception:
            pass

        # 상/하단 20% 이외의 영역(20%~80%)에 위치하면 즉시 클릭
        mid_top    = int(w_h * 0.20)
        mid_bottom = int(w_h * 0.80)

        def _tap_coords_and_return(cx, cy, label):
            if cy < mid_top:
                self._log(f"  📌 [{label}] 상단 20% 치우침 (y={cy} < {mid_top}) -> 미세 안전스크롤 업")
                self._scroll_up(distance_ratio=0.08)
                time.sleep(1.0)
                return False
            elif cy > mid_bottom:
                self._log(f"  📌 [{label}] 하단 20% 치우침 (y={cy} > {mid_bottom}) -> 미세 안전스크롤 다운")
                self._scroll_down_safe(distance_ratio=0.08)
                time.sleep(1.0)
                return False

            self._log(f"  🎯 [{label}] 앞단 안착! 좌표 ({cx}, {cy}) -> 캡처 후 1회 탭")
            
            # 사용자 요청: 클릭 직전 화면 캡처
            try:
                ss_data = self._get_screenshot()
                if ss_data:
                    ss_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_screenshots")
                    os.makedirs(ss_dir, exist_ok=True)
                    ss_path = os.path.join(ss_dir, f"다른결재_탭전_{self.device_id}_{int(time.time())}.png")
                    with open(ss_path, "wb") as f:
                        f.write(ss_data)
                    self._log(f"  📸 화면 캡처 완료: {os.path.basename(ss_path)}")
            except Exception as e:
                self._log(f"  ⚠ 화면 캡처 실패: {e}")

            ah.tap_by_coords(self.driver, cx, cy, self._log)
            time.sleep(2.0)

            self._log(f"  ✅ [{label}] 다른결재 버튼 클릭 완료 -> 즉시 다음 결제 단계 진행")
            return True

        for attempt in range(1, max_scroll_attempts + 1):
            # ── 이미지 매칭 (threshold 0.55: 로그상 0.50~0.61대 후보 허용) ──
            found_and_handled = False
            for img_path, name in img_candidates:
                coords = self._find_image_coords(img_path, threshold=0.55)
                if coords:
                    res = _tap_coords_and_return(coords[0], coords[1], f"이미지/{name}")
                    if res is True:
                        return True
                    elif res is False:
                        found_and_handled = True
                        break

            if found_and_handled:
                continue

            # 결제화면: 짧은·느린 안전스크롤 (앱 창밖 오버스크롤 방지)
            self._log(f"  ⬇ [다른결재] 이미지 미발견 -> 안전스크롤 ({attempt}/{max_scroll_attempts})")
            self._scroll_down_safe(distance_ratio=0.14)
            time.sleep(0.6)

        self._log("  ❌ [다른결재 버튼] 이미지 매칭 탐색 실패")
        return False

    # ─── 결제화면 PaddleOCR 텍스트 추출 & 로그 저장 ────────────────────────────

    def _ocr_payment_screen(self, label: str = "결제화면"):
        """
        현재 화면을 ADB 스크린샷으로 캡처하여 PaddleOCR로 한국어 텍스트를 추출하고,
        기기별 OCR 전용 로그 파일(ocr_기기{machine_num}_{device_id}.log)에 저장합니다.

        Args:
            label: 로그에 기록될 행위 구분 레이블 (예: '결제화면진입', '다른결재탐색' 등)
        """
        try:
            import datetime
            import numpy as np
            import cv2
            from paddleocr import PaddleOCR

            # ── 스크린샷 캡처 ──
            res = _run_cmd(
                ["adb", "-s", self.device_id, "exec-out", "screencap", "-p"],
                capture_output=True, timeout=10
            )
            if not res.stdout or len(res.stdout) < 200:
                self._log("  ⚠ [OCR] 스크린샷 캡처 실패")
                return

            img_arr = np.frombuffer(res.stdout, np.uint8)
            screen_bgr = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
            if screen_bgr is None:
                self._log("  ⚠ [OCR] 이미지 디코딩 실패")
                return

            # ── PaddleOCR 실행 (한국어, GPU 없이) ──
            try:
                ocr_engine = PaddleOCR(use_angle_cls=True, lang="korean")
                results = ocr_engine.ocr(screen_bgr, cls=True)
            except Exception as pe:
                self._log(f"  ⚠ [OCR] PaddleOCR 오류: {pe}")
                return

            # ── 텍스트 추출 ──
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            lines_extracted = []
            if results:
                for page in results:
                    if not page:
                        continue
                    for item in page:
                        # item: [[x1,y1],[x2,y2],[x3,y3],[x4,y4]], (text, confidence)
                        try:
                            box = item[0]  # 4개 꼭지점 좌표
                            text_info = item[1]  # (텍스트, 신뢰도)
                            text = text_info[0].strip() if text_info and text_info[0] else ""
                            conf = float(text_info[1]) if text_info and len(text_info) > 1 else 0.0
                            if text and conf >= 0.4:
                                cx = int((box[0][0] + box[2][0]) / 2)
                                cy = int((box[0][1] + box[2][1]) / 2)
                                lines_extracted.append(f"  [{cx:4d},{cy:4d}] (신뢰도:{conf:.2f}) {text}")
                        except Exception:
                            continue

            # ── OCR 전용 로그 파일 저장 ──
            today_ocr = datetime.datetime.now().strftime("%Y%m%d")
            log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs", today_ocr)
            os.makedirs(log_dir, exist_ok=True)
            ocr_log_path = os.path.join(
                log_dir,
                f"ocr_기기{self.machine_num}_{self.device_id}.log"
            )
            separator = "=" * 60
            with open(ocr_log_path, "a", encoding="utf-8") as f:
                f.write(f"\n{separator}\n")
                f.write(f"[{now}] [{self.device_id}] 📸 {label}\n")
                f.write(f"{separator}\n")
                if lines_extracted:
                    for line in lines_extracted:
                        f.write(line + "\n")
                else:
                    f.write("  (텍스트 없음)\n")

            self._log(f"  📋 [OCR] '{label}' 화면 텍스트 {len(lines_extracted)}줄 추출 완료 → {ocr_log_path}")

        except ImportError:
            self._log("  ℹ [OCR] PaddleOCR 미설치 → OCR 건너뜀 (pip install paddleocr 필요)")
        except Exception as e:
            self._log(f"  ⚠ [OCR] 전체 오류: {e}")

    def _process_bank_transfer(self) -> bool:
        self._log("💰 [무통장 결제] 프로세스 시작")
        self._ocr_payment_screen(label="무통장결제_화면진입")
            
        if not self._click_other_pay_button(max_scroll_attempts=20):
            self._log("⚠ '다른결재 관련 버튼' 미발견 -> 스크롤을 위로 올린 후 탐색 시작")
            for _ in range(5):
                self._scroll_up(distance_ratio=0.5)
                time.sleep(0.4)

            self._log("🔍 '미신청2' 또는 '일반결재' 등을 찾아 스크롤 탐색합니다.")
            found_action = False
            for attempt in range(8):
                # 1. 미신청2 가장 먼저 확인 (발견 시 결제수단 설정 스킵하고 바로 return True)
                if os.path.exists(IMG_NOT_APPLY2) and self._find_image_coords(IMG_NOT_APPLY2, threshold=0.75):
                    self._log("✅ '미신청2' 이미지 확인! 결제수단 셋팅 생략 후 즉시 다음 단계(주문하기)로 넘어갑니다.")
                    return True
                
                # 2. 체크 상태 확인
                if os.path.exists(IMG_NORMAL_PAY_CHECK) and self._find_image_coords(IMG_NORMAL_PAY_CHECK, threshold=0.70):
                    self._log("✅ '일반결재체크' 확인됨. 계속 진행합니다.")
                    found_action = True
                    break
                elif os.path.exists(IMG_BANK_TRANSFER_CHECK) and self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70):
                    self._log("✅ '무통장체크' 확인됨. 계속 진행합니다.")
                    found_action = True
                    break
                
                # 3. 일반결재 클릭 시도
                normal_pay_images = [
                    (IMG_NORMAL_PAY, "일반결재"),
                    (IMG_NORMAL_PAY3, "일반결재3"),
                ]
                clicked_normal = False
                for img_path, name in normal_pay_images:
                    if os.path.exists(img_path):
                        coords = self._find_image_coords(img_path, threshold=0.75)
                        if coords:
                            ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                            time.sleep(1.5)
                            self._log(f"✅ '{name}' 발견 및 클릭 성공! 계속 진행합니다.")
                            clicked_normal = True
                            break
                
                if clicked_normal:
                    found_action = True
                    break
                    
                # 미발견 시 미세 스크롤 다운
                self._scroll_down(distance_ratio=0.20)
                time.sleep(0.8)
                
            if not found_action:
                self._log("❌ '다른결재', '미신청2', '일반결재' 및 체크 상태 모두 미발견 -> 무통장 결제 실패")
                return False
        
        # 2. 일반결재 탐색 및 클릭 (이미 체크되어 있으면 스킵)
        is_already_checked = False
        if os.path.exists(IMG_NORMAL_PAY_CHECK) and self._find_image_coords(IMG_NORMAL_PAY_CHECK, threshold=0.70):
            self._log("✅ '일반결재체크' 상태 감지됨. 일반결재 클릭 건너뜀.")
            is_already_checked = True
        elif os.path.exists(IMG_BANK_TRANSFER_CHECK) and self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70):
            self._log("✅ '무통장체크' 상태 감지됨. 일반결재 클릭 건너뜀.")
            is_already_checked = True
            
        if not is_already_checked:
            normal_pay_images = [
                (IMG_NORMAL_PAY, "일반결재"),
                (IMG_NORMAL_PAY3, "일반결재3"),
            ]
            if not self._click_any_image_with_scroll(normal_pay_images, threshold=0.75, max_scroll_attempts=8):
                if os.path.exists(IMG_NORMAL_PAY_CHECK) and self._find_image_coords(IMG_NORMAL_PAY_CHECK, threshold=0.70):
                    self._log("✅ '일반결재체크' 발견! 성공으로 간주하고 진행합니다.")
                elif os.path.exists(IMG_BANK_TRANSFER_CHECK) and self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70):
                    self._log("✅ '무통장체크' 발견! 성공으로 간주하고 진행합니다.")
                else:
                    self._log("❌ '일반결재/일반결재3' 버튼 미발견 -> 무통장 결제 실패")
                    return False
                    
        # 3. 무통장입금 탐색 (이미 체크되어 있으면 한 번 클릭 후 은행 선택으로 진행)
        is_bank_transfer_checked = False
        if os.path.exists(IMG_BANK_TRANSFER_CHECK) and self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70):
            self._log("✅ '무통장체크' 상태 감지됨! 이미지를 한 번 클릭 후 은행 선택으로 진행합니다.")
            coords = self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70)
            if coords:
                ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
            time.sleep(1.5)
            is_bank_transfer_checked = True
            
        if not is_bank_transfer_checked:
            if not self._click_image_with_scroll(IMG_BANK_TRANSFER, "무통장입금", max_scroll_attempts=8):
                if os.path.exists(IMG_BANK_TRANSFER_CHECK) and self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70):
                    self._log("✅ '무통장체크' 상태 감지됨! 이미지를 한 번 클릭 후 은행 선택으로 진행합니다.")
                    coords = self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70)
                    if coords:
                        ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    time.sleep(1.5)
                else:
                    self._log("⚠ '무통장입금' 버튼 미발견. 중단하지 않고 계속 진행합니다.")
            else:
                time.sleep(1.5)

        # 무통장입금 클릭(또는 스킵) 후 '은행을' / '은행선택' 탐색
        if not self._click_bank_select_with_scroll(max_scroll_attempts=8):
            self._log("⚠ 무통장입금 클릭 후 '은행을' / '은행선택' 미발견. 무시하고 계속 진행합니다.")

        # 모든 은행 템플릿 후보 (1번/2번 포함)
        all_bank_candidates = [
            (IMG_HANA_BANK, "하나은행"), (IMG_HANA_BANK2, "하나은행2"),
            (IMG_NONGHYUP_BANK, "농협"), (IMG_NONGHYUP_BANK2, "농협2"),
            (IMG_WOORI_BANK, "우리은행"), (IMG_WOORI_BANK2, "우리은행2"),
            (IMG_KB_BANK, "국민은행"), (IMG_KB_BANK2, "국민은행2"),
            (IMG_IBK_BANK, "기업은행"), (IMG_IBK_BANK2, "기업은행2"),
            (IMG_SHINHAN_BANK, "신한은행"), (IMG_SHINHAN_BANK2, "신한은행2"),
            (IMG_BUSAN_BANK, "부산은행"), (IMG_BUSAN_BANK2, "부산은행2"),
        ]
        valid_bank_candidates = [(p, name) for p, name in all_bank_candidates if os.path.exists(p)]
        random.shuffle(valid_bank_candidates)

        bank_selected = False
        self._log("🎲 [은행 선택] 노출된 은행(1번/2번 이미지) 즉시 탐색 및 선택 시작...")

        for attempt in range(1, 8):
            for bank_img, bank_name in valid_bank_candidates:
                coords = self._find_image_coords(bank_img, threshold=0.70)
                if coords:
                    self._log(f"✅ 은행 [{bank_name}] 발견! 좌표 ({coords[0]}, {coords[1]}) -> 클릭하여 선택 완료")
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    time.sleep(2.0)
                    bank_selected = True
                    break

            if bank_selected:
                break

            self._log(f"  ⬇ [은행 탐색] 현재 화면 미노출 -> 미세 스크롤 다운 ({attempt}/7)")
            self._scroll_down(distance_ratio=0.25)
            time.sleep(1.0)

        if not bank_selected:
            self._log("⚠ 랜덤 은행 선택 실패! 하지만 중단하지 않고 다음 단계(미신청 및 결제/주문하기)로 계속 진행합니다.")
            # 실패 시 중단하지 않고 계속 진행 (사용자 요청)
            
        # ── 미신청 탐색 전: 팝업 점검 및 위에서부터 탐색 ──
        self._dismiss_payment_benefit_popup()

        # 위에서부터 살짝 밑으로 내리며 탐색하기 위해 스크롤을 살짝 위로 올림
        for _ in range(2):
            self._scroll_up(distance_ratio=0.4)
            time.sleep(0.3)

        not_apply_images = [
            (IMG_NOT_APPLY, "미신청"),
            (IMG_NOT_APPLY2, "미신청2"),
            (IMG_NOT_APPLY3, "미신청3"),
            (IMG_NOT_APPLY4, "미신청4"),
        ]

        # 미신청 라디오 버튼은 화면 우측 25% 이내(x >= 75%) 영역에 위치하므로 min_x 제약을 지정
        w_w = 1080
        try:
            w_w = self.driver.get_window_size()['width']
        except Exception:
            pass
        min_x_right = int(w_w * 0.75)

        if not self._click_any_image_with_scroll(not_apply_images, threshold=0.65, max_scroll_attempts=10, min_x=min_x_right):
            self._log("⚠ '미신청' 이미지 미발견 -> XPath 텍스트 탐색 시도")
            for _ in range(3):
                self._scroll_up(distance_ratio=0.5)
                time.sleep(0.3)
            self._dismiss_payment_benefit_popup()
            
            not_apply_xpaths = [
                '//*[contains(@text, "미신청")]',
                '//*[contains(@content-desc, "미신청")]',
                '//android.widget.RadioButton[contains(@text, "신청안함")]',
                '//*[contains(@text, "신청 안 함")]',
            ]
            found_xpath = False
            for xp in not_apply_xpaths:
                try:
                    if ah.element_exists(self.driver, xp, timeout=1):
                        el = self.driver.find_element(By.XPATH, xp)
                        self._safe_click_element(el)
                        self._log(f"✅ '미신청' XPath 발견 및 클릭 성공: {xp}")
                        found_xpath = True
                        time.sleep(1.0)
                        break
                except Exception:
                    continue

            if not found_xpath:
                self._log("❌ '미신청' 및 '미신청2' 버튼 이미지/XPath 모두 미발견 -> 무통장 결제 실패")
                return False

        # ── 주문하기 탐색 전: 팝업 점검 및 화면 제일 밑으로 스크롤 ──
        self._dismiss_payment_benefit_popup()

        self._log("🔍 [무통장] 화면 제일 밑으로 스크롤 후 '주문하기.png' 이미지 인식 클릭 시도")
        for _ in range(4):
            self._scroll_down(distance_ratio=0.6)
            time.sleep(0.4)

        found_order = False
        if os.path.exists(IMG_DO_ORDER):
            for attempt in range(5):
                coords = self._find_image_coords(IMG_DO_ORDER, threshold=0.70)
                if coords:
                    if self._skip_final_order_click():
                        mode = "수동시작" if self.manual_mode else "테스트 모드"
                        self._log(f"✅ '주문하기' 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 🖐 [{mode}] 클릭 생략")
                    else:
                        self._log(f"✅ '주문하기' 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                        ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    found_order = True
                    time.sleep(2.0)
                    break
                else:
                    self._log(f"  📌 '주문하기.png' 미발견 -> 미세 스크롤 다운 ({attempt+1}/5)")
                    self._scroll_down(distance_ratio=0.20)
                    time.sleep(0.8)

        if not found_order:
            self._log("⚠ '주문하기' 이미지 미발견 -> 스크롤을 위로 올리며 '결재하기' / '주문결재' 이미지 탐색 시도")
            for _ in range(5):
                self._scroll_up(distance_ratio=0.4)
                time.sleep(0.6)
                for pay_img, pay_name in [(IMG_DO_PAY, "결재하기"), (IMG_ORDER_PAY, "주문결재")]:
                    if os.path.exists(pay_img):
                        coords = self._find_image_coords(pay_img, threshold=0.70)
                        if coords:
                            if self._skip_final_order_click():
                                mode = "수동시작" if self.manual_mode else "테스트 모드"
                                self._log(f"✅ '{pay_name}' 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 🖐 [{mode}] 클릭 생략")
                            else:
                                self._log(f"✅ '{pay_name}' 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                                ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                            found_order = True
                            time.sleep(2.0)
                            break
                if found_order:
                    break

        if not found_order:
            self._log("❌ '주문하기' 및 '결재하기' 버튼 이미지 모두 미발견 -> 무통장 결제 실패")
            return False

        return True

    def _dismiss_payment_benefit_popup(self) -> None:
        """
        화면에 '결제혜택.png' 팝업이 존재할 때만 '닫기.png' 또는 '닫기' 버튼을 눌러 팝업을 닫습니다. (1회만 실행)
        """
        if getattr(self, 'has_dismissed_payment_benefit', False):
            return

        self.has_dismissed_payment_benefit = True
        self._log("🔍 [팝업 닫기 점검] 결제혜택.png 감지 시도 (최초 1회)...")
        try:
            # 반드시 '결제혜택.png'가 화면에 인식되어야 팝업 닫기 작동
            if os.path.exists(IMG_PAY_BENEFIT):
                coords = self._find_image_coords(IMG_PAY_BENEFIT, threshold=0.65)
                if coords:
                    self._log(f"  🎯 [결제혜택 팝업] 감지! 좌표 ({coords[0]}, {coords[1]}) → 닫기 버튼 탐색")
                    
                    # 1순위: 닫기.png 이미지 매칭
                    if os.path.exists(IMG_CLOSE_POPUP):
                        close_coords = self._find_image_coords(IMG_CLOSE_POPUP, threshold=0.65)
                        if close_coords:
                            self._log(f"  🎯 [닫기.png] 발견! 좌표 ({close_coords[0]}, {close_coords[1]}) → 클릭 닫기")
                            import naver_appium as _ah
                            _ah.tap_by_coords(self.driver, close_coords[0], close_coords[1], self._log)
                            time.sleep(1.0)
                            self._log("  ✅ [결제혜택 팝업] '닫기.png' 이미지로 팝업 닫기 완료")
                            return

                    # 2순위: XPath 닫기 버튼
                    close_xpaths = [
                        '//*[@content-desc="닫기"]',
                        '//*[@text="닫기"]',
                        '//android.widget.Button[@text="닫기"]',
                        '//*[@content-desc="close"]',
                        '//*[@content-desc="Close"]',
                    ]
                    for xpath in close_xpaths:
                        try:
                            if ah.element_exists(self.driver, xpath, timeout=1):
                                el = self.driver.find_element(By.XPATH, xpath)
                                self._safe_click_element(el)
                                self._log(f"  ✅ [결제혜택 팝업] XPath 닫기 클릭 완료: {xpath}")
                                time.sleep(1.0)
                                return
                        except Exception:
                            continue

            self._log("  ℹ [팝업 닫기 점검] 결제혜택 팝업 미감지 → 계속 진행")
        except Exception as e:
            self._log(f"  ⚠ [팝업 닫기 점검] 처리 중 예외 발생: {e} → 계속 진행")

    def _process_money_payment(self, password: str) -> bool:
        self._log("💸 [머니 결제] 프로세스 시작")
        if not self._click_image_with_scroll(IMG_MONEY_PAY, "머니"):
            return False
        if not self._click_pay_button():
            self._log("❌ 결제하기 클릭 실패")
            return False
        if password:
            if not self._input_password(password):
                self._log("❌ 머니 결제 비밀번호 입력 실패")
                return False
        else:
            self._log("  ℹ 비밀번호 없음 → 건너뜀")
        return True

    def _is_hyundai_card_payment(self, method: str) -> bool:
        m = (method or "").replace(" ", "")
        if self._is_kb_card_payment(method):
            return False
        return any(k in m for k in ("현대카드", "현대하드", "현대"))

    def _is_kb_card_payment(self, method: str) -> bool:
        """결제방식: 국민카드 / kb국민 / 국만카드."""
        m = (method or "").replace(" ", "")
        m_lower = m.lower()
        if any(k in m for k in ("국민카드", "국만카드", "KB국민", "kb국민")):
            return True
        if m_lower in ("kb", "kb국민카드") or m == "국민":
            return True
        return False

    def _ensure_normal_pay_checked(self) -> bool:
        """[22-2] 일반결재가 체크되어 있어야 함. 아니면 클릭."""
        if os.path.exists(IMG_NORMAL_PAY_CHECK) and self._find_image_coords(IMG_NORMAL_PAY_CHECK, threshold=0.70):
            self._log("✅ [22-2] '일반결재체크' 상태 확인")
            return True
        normal_pay_images = [
            (IMG_NORMAL_PAY, "일반결재"),
            (IMG_NORMAL_PAY3, "일반결재3"),
        ]
        if self._click_any_image_with_scroll(normal_pay_images, threshold=0.75, max_scroll_attempts=8):
            self._log("✅ [22-2] 일반결재 클릭 완료")
            return True
        if os.path.exists(IMG_NORMAL_PAY_CHECK) and self._find_image_coords(IMG_NORMAL_PAY_CHECK, threshold=0.70):
            self._log("✅ [22-2] '일반결재체크' 발견")
            return True
        self._log("❌ [22-2] 일반결재 미확인")
        return False

    def _card_placeholder_visible(self) -> bool:
        """드롭다운이 아직 '카드를 선택해주세요' 상태인지."""
        xpaths = [
            '//*[contains(@text,"카드를 선택해주세요")]',
            '//*[contains(@text,"카드를 선택")]',
        ]
        for xp in xpaths:
            try:
                if ah.element_exists(self.driver, xp, timeout=0.8):
                    return True
            except Exception:
                continue
        return False

    def _hyundai_card_selected(self) -> bool:
        """카드 필드에 현대가 선택되었는지 (플레이스홀더가 사라진 상태)."""
        if self._card_placeholder_visible():
            return False
        xpaths = [
            '//*[contains(@text,"현대카드")]',
            '//*[contains(@text,"현대")]',
        ]
        for xp in xpaths:
            try:
                els = self.driver.find_elements(By.XPATH, xp)
                for el in els:
                    t = (el.get_attribute("text") or "").strip()
                    if "현대" in t and "선택해주세요" not in t:
                        bb = self._parse_element_bounds(el)
                        if not bb:
                            continue
                        cy = (bb[1] + bb[3]) // 2
                        # 상단 헤더(오탐) 제외
                        if cy < 280:
                            continue
                        self._log(f"  ✅ 현대카드 선택 확인: '{t[:30]}' y={cy}")
                        return True
            except Exception:
                continue
        return False

    def _open_card_select_dropdown(self, brand: str = "hyundai") -> bool:
        """[22-3] '카드를 선택해주세요' 드롭다운을 연다. brand=hyundai|kb"""
        self._set_status("카드 선택 드롭다운")
        self._log(f"🔍 [22-3] 카드 선택 드롭다운 열기 (brand={brand})")
        if brand == "kb":
            if self._kb_card_selected():
                self._log("  ℹ 이미 국민카드(KB)가 선택되어 있음")
                return True
        elif self._hyundai_card_selected():
            self._log("  ℹ 이미 현대카드가 선택되어 있음")
            return True

        w_h = 2400
        try:
            w_h = self.driver.get_window_size()["height"]
        except Exception:
            pass
        min_y, max_y = int(w_h * 0.22), int(w_h * 0.88)

        xpaths = [
            '//*[contains(@text,"카드를 선택해주세요")]',
            '//android.widget.Button[contains(@text,"카드를 선택")]',
            '//android.view.View[contains(@text,"카드를 선택")]',
        ]

        for attempt in range(1, 7):
            clicked = False
            for xp in xpaths:
                try:
                    els = self.driver.find_elements(By.XPATH, xp)
                except Exception:
                    continue
                for el in els:
                    bb = self._parse_element_bounds(el)
                    if not bb:
                        continue
                    cy = (bb[1] + bb[3]) // 2
                    if not (min_y <= cy <= max_y):
                        continue
                    cx = (bb[0] + bb[2]) // 2
                    self._log(f"  🎯 [22-3] XPath 드롭다운 탭 ({cx}, {cy})")
                    self._soft_tap(cx, cy)
                    clicked = True
                    break
                if clicked:
                    break

            if not clicked:
                if self._click_any_image_with_scroll(
                    IMG_HYUNDAI_CARDS, threshold=0.75, max_scroll_attempts=3,
                    min_y=min_y, max_y=max_y,
                ):
                    clicked = True

            if clicked:
                time.sleep(1.8)
                if brand == "kb":
                    if self._find_kb_list_target():
                        self._log("  ✅ [22-3] 카드 목록에서 KB국민 항목 감지")
                        return True
                    self._log("  ℹ [22-3] 드롭다운 클릭 완료 → 목록에서 KB국민 탐색")
                else:
                    if self._find_hyundai_list_target():
                        self._log("  ✅ [22-3] 카드 목록에서 현대 항목 감지")
                        return True
                    self._log("  ℹ [22-3] 드롭다운 클릭 완료 → 목록에서 현대 탐색")
                return True
            else:
                self._log(f"  ⬇ [22-3] 드롭다운 미발견 → 스크롤 ({attempt}/6)")
                self._scroll_down_safe(distance_ratio=0.10)
                time.sleep(0.6)

        self._log("❌ [22-3] 카드 선택 드롭다운을 열지 못함")
        return False

    def _kb_card_selected(self) -> bool:
        """카드 필드에 KB국민이 선택되었는지."""
        if self._card_placeholder_visible():
            return False
        xpaths = [
            '//*[contains(@text,"KB국민")]',
            '//*[contains(@text,"kb국민")]',
            '//*[contains(@text,"국민카드")]',
        ]
        for xp in xpaths:
            try:
                els = self.driver.find_elements(By.XPATH, xp)
                for el in els:
                    t = (el.get_attribute("text") or "").strip()
                    if not t:
                        continue
                    if "선택해주세요" in t:
                        continue
                    if any(k in t for k in ("KB국민", "kb국민", "국민카드", "국민")):
                        bb = self._parse_element_bounds(el)
                        if not bb:
                            continue
                        cy = (bb[1] + bb[3]) // 2
                        if cy < 280:
                            continue
                        self._log(f"  ✅ 국민카드 선택 확인: '{t[:30]}' y={cy}")
                        return True
            except Exception:
                continue
        return False

    def _find_kb_list_target(self):
        """카드 목록에서 클릭할 KB국민 항목 (x,y) 또는 None."""
        w_h, w_w = 2400, 1080
        try:
            sz = self.driver.get_window_size()
            w_h, w_w = sz["height"], sz["width"]
        except Exception:
            pass
        min_y, max_y = int(w_h * 0.22), int(w_h * 0.88)

        xpaths = [
            '//*[contains(@text,"KB국민")]',
            '//*[contains(@text,"kb국민")]',
            '//*[contains(@text,"국민카드")]',
            '//android.widget.TextView[contains(@text,"국민")]',
            '//android.view.View[contains(@text,"국민")]',
        ]
        best = None
        for xp in xpaths:
            try:
                els = self.driver.find_elements(By.XPATH, xp)
            except Exception:
                continue
            for el in els:
                try:
                    t = (el.get_attribute("text") or "").strip()
                    if "선택해주세요" in t:
                        continue
                    if not any(k in t for k in ("KB국민", "kb국민", "국민카드", "국민")):
                        continue
                    if "현대" in t:
                        continue
                    bb = self._parse_element_bounds(el)
                    if not bb:
                        continue
                    cy = (bb[1] + bb[3]) // 2
                    cx = (bb[0] + bb[2]) // 2
                    if not (min_y <= cy <= max_y):
                        continue
                    # KB국민 우선
                    score = 2 if ("KB" in t or "kb" in t.lower()) else 1
                    cand = (score, cx, cy, t)
                    if best is None or cand[0] > best[0]:
                        best = cand
                except Exception:
                    continue
        if best:
            return best[1], best[2]

        for img_path, name in IMG_KB_BRAND:
            if not os.path.exists(img_path):
                continue
            coords = self._find_image_coords(
                img_path, threshold=0.70, min_y=min_y, max_y=max_y
            )
            if coords:
                self._log(f"  🎯 KB목록 이미지 '{name}' @ {coords}")
                return coords
        return None

    def _pick_kb_card(self) -> bool:
        """목록에서 kb국민1/2 클릭."""
        self._set_status("국민카드 선택")
        self._log("🔍 [국민카드] kb국민1/2 선택 시도")
        if self._kb_card_selected():
            self._log("  ✅ 이미 국민카드 선택됨")
            return True

        w_h = 2400
        try:
            w_h = self.driver.get_window_size()["height"]
        except Exception:
            pass
        min_y, max_y = int(w_h * 0.22), int(w_h * 0.88)

        for attempt in range(1, 8):
            target = self._find_kb_list_target()
            if target:
                self._soft_tap(target[0], target[1])
                time.sleep(1.5)
                if self._kb_card_selected() or not self._card_placeholder_visible():
                    self._log("  ✅ [국민카드] 선택 확인")
                    return True
                self._log("  ⚠ 탭 후에도 플레이스홀더 잔존 → 재시도")

            # 이미지 직접 클릭
            if self._click_any_image_basic(
                IMG_KB_BRAND, threshold=0.68, attempts=2, wait_after=1.5,
                min_y=min_y, max_y=max_y,
            ):
                time.sleep(1.0)
                if self._kb_card_selected() or not self._card_placeholder_visible():
                    self._log("  ✅ [국민카드] 이미지 클릭 후 선택 확인")
                    return True

            # 드롭다운 재오픈
            if self._card_placeholder_visible() or attempt % 2 == 0:
                self._open_card_select_dropdown(brand="kb")
            else:
                self._scroll_down_safe(distance_ratio=0.10)
            time.sleep(0.6)

        self._log("❌ [국민카드] kb국민 선택 실패")
        return False

    def _find_hyundai_list_target(self):
        """카드 목록에서 클릭할 현대 항목 (x,y) 또는 없으면 None.
        상단 헤더(y<22%) 오탐을 제외한다.
        """
        w_h, w_w = 2400, 1080
        try:
            sz = self.driver.get_window_size()
            w_h, w_w = sz["height"], sz["width"]
        except Exception:
            pass
        min_y, max_y = int(w_h * 0.22), int(w_h * 0.88)

        xpaths = [
            '//*[contains(@text,"현대카드")]',
            '//android.widget.TextView[@text="현대"]',
            '//android.view.View[@text="현대"]',
            '//android.widget.TextView[contains(@text,"현대")]',
            '//android.view.View[contains(@text,"현대")]',
        ]
        best = None
        for xp in xpaths:
            try:
                els = self.driver.find_elements(By.XPATH, xp)
            except Exception:
                continue
            for el in els:
                try:
                    t = (el.get_attribute("text") or "").strip()
                    if "현대" not in t:
                        continue
                    if "선택해주세요" in t:
                        continue
                    bb = self._parse_element_bounds(el)
                    if not bb:
                        continue
                    cy = (bb[1] + bb[3]) // 2
                    cx = (bb[0] + bb[2]) // 2
                    if not (min_y <= cy <= max_y):
                        continue
                    score = 2 if "현대카드" in t else 1
                    cand = (score, cx, cy, t)
                    if best is None or cand[0] > best[0]:
                        best = cand
                except Exception:
                    continue
        if best:
            # (cx, cy, label, src) — _pick_hyundai_card 호환
            return best[1], best[2], best[3], "xpath"

        for img_path, name in IMG_HYUNDAI_BRAND:
            if not os.path.exists(img_path):
                continue
            coords = self._find_image_coords(
                img_path, threshold=0.70, min_y=min_y, max_y=max_y
            )
            if coords:
                self._log(f"  🎯 현대목록 이미지 '{name}' @ {coords}")
                return coords[0], coords[1], name, "image"
        return None

    def _pick_hyundai_card(self) -> bool:
        """[22-4] 열린 카드 목록에서 현대만 클릭하고, 선택 여부를 검증한다."""
        self._set_status("현대카드 선택")
        self._log("🔍 [22-4] 카드 목록에서 현대 선택")
        if self._hyundai_card_selected():
            self._log("  ✅ 이미 현대카드 선택됨")
            return True

        for attempt in range(1, 9):
            target = self._find_hyundai_list_target()
            if target:
                cx, cy, label, src = target
                self._log(f"  🎯 [22-4] 현대 탭 ({cx}, {cy}) src={src} '{str(label)[:24]}'")
                self._soft_tap(cx, cy)
                time.sleep(1.8)
                if self._hyundai_card_selected():
                    self._log("  ✅ [22-4] 현대카드 선택 확인")
                    return True
                if not self._card_placeholder_visible():
                    # 목록이 닫혔고 플레이스홀더가 없음 → 선택 성공으로 간주
                    self._log("  ✅ [22-4] 드롭다운 닫힘(플레이스홀더 소멸) → 선택 완료")
                    return True
                self._log("  ⚠ [22-4] 탭 후에도 '카드를 선택해주세요' 유지 → 재시도")
                if attempt in (3, 6) and self._card_placeholder_visible():
                    self._log("  🔄 [22-4] 드롭다운 다시 열기")
                    self._open_card_select_dropdown(brand="hyundai")
            else:
                self._log(f"  ⬇ [22-4] 목록에서 현대 미발견 → 스크롤 ({attempt}/8)")
                self._scroll_down_safe(distance_ratio=0.10)
                time.sleep(0.6)
                if attempt == 3 and self._card_placeholder_visible():
                    self._log("  🔄 [22-4] 목록 미검출 → 드롭다운 다시 열기")
                    self._open_card_select_dropdown(brand="hyundai")

        self._log("❌ [22-4] 현대카드 선택 실패")
        return False

    def _find_image_bbox(self, template_path: str, threshold: float = 0.55,
                         save_crop: bool = True, crop_label: str = "") -> Optional[dict]:
        """
        템플릿 매칭 후 bbox 반환.
        반환: {cx, cy, x1, y1, x2, y2, score, tw, th} 또는 None
        """
        try:
            import cv2
            import numpy as np
            from PIL import Image
            import io
        except ImportError:
            return None

        if not os.path.exists(template_path):
            self._log(f"  [bbox매칭] 템플릿 없음: {template_path}")
            return None

        try:
            screenshot_png = self._get_screenshot()
            screenshot_pil = Image.open(io.BytesIO(screenshot_png))
            screen_bgr = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)
            screen_gray = cv2.cvtColor(screen_bgr, cv2.COLOR_BGR2GRAY)
            screen_h, screen_w = screen_gray.shape

            template_bgr = cv2.imdecode(
                np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR
            )
            if template_bgr is None:
                return None
            template_gray = cv2.cvtColor(template_bgr, cv2.COLOR_BGR2GRAY)
            t_h, t_w = template_gray.shape

            best_score, best_loc, best_tw, best_th = -1, None, t_w, t_h
            for scale in np.linspace(0.45, 1.80, 18):
                nw = int(t_w * scale)
                nh = int(t_h * scale)
                if nw >= screen_w or nh >= screen_h or nw < 20 or nh < 20:
                    continue
                resized = cv2.resize(template_gray, (nw, nh), interpolation=cv2.INTER_AREA)
                result = cv2.matchTemplate(screen_gray, resized, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)
                if max_val > best_score:
                    best_score, best_loc, best_tw, best_th = max_val, max_loc, nw, nh

            if best_score < threshold or best_loc is None:
                self._log(f"  ❌ [bbox매칭] 실패 (점수 {best_score:.4f} < {threshold})")
                return None

            x1, y1 = int(best_loc[0]), int(best_loc[1])
            x2, y2 = min(screen_w, x1 + best_tw), min(screen_h, y1 + best_th)
            cx, cy = x1 + best_tw // 2, y1 + best_th // 2
            self._log(
                f"  🎯 [bbox매칭] 발견 score={best_score:.4f} "
                f"center=({cx},{cy}) box=({x1},{y1})-({x2},{y2})"
            )

            if save_crop:
                try:
                    root_dir = os.path.dirname(os.path.abspath(__file__))
                    recognize_dir = os.path.join(root_dir, "인식")
                    os.makedirs(recognize_dir, exist_ok=True)
                    label = crop_label or os.path.splitext(os.path.basename(template_path))[0]
                    timestamp = time.strftime("%Y%m%d_%H%M%S")
                    crop_path = os.path.join(
                        recognize_dir,
                        f"crop_{label}_({cx},{cy})_{best_score:.3f}_{timestamp}.png",
                    )
                    cropped = screen_bgr[y1:y2, x1:x2]
                    if cropped.size > 0:
                        cv2.imwrite(crop_path, cropped)
                        self._log(f"  📸 [키패드 커팅 저장] {crop_path}")
                except Exception:
                    pass

            return {
                "cx": cx, "cy": cy,
                "x1": x1, "y1": y1, "x2": x2, "y2": y2,
                "score": float(best_score),
                "tw": best_tw, "th": best_th,
                "screen_w": screen_w, "screen_h": screen_h,
            }
        except Exception as e:
            self._log(f"  [bbox매칭 오류] {e}")
            return None

    def _locate_hyundai_pw_keypad(self, attempts: int = 5) -> Optional[dict]:
        """
        현대비번.png 인식 → 키패드 ROI 커팅.
        반환 ROI: {min_x, max_x, min_y, max_y, cx, cy, ...}
        """
        if not os.path.exists(IMG_HYUNDAI_PW_KEYPAD):
            self._log("  ⚠ 현대비번.png 파일 없음 → ROI 없이 진행")
            return None

        self._log("  🔍 [현대비번] 키패드 영역 인식/커팅 시도...")
        for attempt in range(1, attempts + 1):
            box = self._find_image_bbox(
                IMG_HYUNDAI_PW_KEYPAD,
                threshold=0.50,
                save_crop=True,
                crop_label="현대비번",
            )
            if box:
                # 여유 마진 (키 가장자리 포함)
                pad_x = max(20, int(box["tw"] * 0.04))
                pad_y = max(20, int(box["th"] * 0.03))
                sw, sh = box["screen_w"], box["screen_h"]
                roi = {
                    "min_x": max(0, box["x1"] - pad_x),
                    "max_x": min(sw, box["x2"] + pad_x),
                    "min_y": max(0, box["y1"] - pad_y),
                    "max_y": min(sh, box["y2"] + pad_y),
                    "cx": box["cx"],
                    "cy": box["cy"],
                    "score": box["score"],
                }
                self._log(
                    f"  ✅ [현대비번] 키패드 ROI 확정 "
                    f"x={roi['min_x']}~{roi['max_x']} y={roi['min_y']}~{roi['max_y']} "
                    f"(시도 {attempt}/{attempts})"
                )
                return roi
            self._log(f"  ↩ [현대비번] 미발견 ({attempt}/{attempts})")
            time.sleep(1.0)
        self._log("  ⚠ [현대비번] 키패드 ROI 실패 → 기본 영역으로 숫자 입력")
        return None

    def _input_hyundai_digits_with_fallback(self, password: str, expected_len: int,
                                            use_keypad_crop: bool = True) -> bool:
        """현대 PIN 입력: 현대비번 ROI 커팅 → 이미지 키패드 → send_keys 폴백."""
        pwd_digits = ''.join(filter(str.isdigit, password or ""))
        if expected_len and len(pwd_digits) != expected_len:
            self._log(f"  ❌ PIN 자릿수 불일치: {len(pwd_digits)}자리 (기대 {expected_len}자리)")
            return False
        if not pwd_digits:
            self._log("  ❌ PIN 값 없음")
            return False

        self._log(f"  🔐 PIN 입력 시도: {'*' * len(pwd_digits)}자리")
        time.sleep(1.0)

        roi = None
        if use_keypad_crop:
            roi = self._locate_hyundai_pw_keypad(attempts=4)

        # 방법1: 이미지 키패드 (현대비번 ROI 내부 우선)
        img_ok = self._input_hyundai_digits(pwd_digits, expected_len=expected_len, roi=roi)
        if img_ok:
            return True

        self._log("  ⚠ 이미지 키패드 실패 → send_keys 폴백 시도")

        # 방법2: EditText send_keys
        for xp in ['//android.widget.EditText', '//*[@class="android.widget.EditText"]']:
            try:
                if ah.element_exists(self.driver, xp, timeout=2):
                    el = self.driver.find_element(By.XPATH, xp)
                    el.clear()
                    el.send_keys(pwd_digits)
                    self._log(f"  ✅ send_keys 입력 완료: {xp}")
                    time.sleep(0.5)
                    return True
            except Exception as e:
                self._log(f"  ⚠ send_keys 실패 ({xp}): {e}")

        # 방법3: ADB input text
        try:
            _run_cmd(
                ["adb", "-s", self.device_id, "shell", "input", "text", pwd_digits],
                timeout=8,
            )
            self._log(f"  ✅ ADB input text 완료")
            time.sleep(0.5)
            return True
        except Exception as e:
            self._log(f"  ❌ ADB input text 실패: {e}")

        return False

    def _input_hyundai_digits(self, password: str, expected_len: int,
                              roi: Optional[dict] = None) -> bool:
        """현대숫자 폴더 0.png~9.png 로 PIN 입력 (현대비번 ROI + 키패드 그리드 보정)."""
        pwd_digits = ''.join(filter(str.isdigit, password or ""))
        if expected_len and len(pwd_digits) != expected_len:
            self._log(f"  ❌ 현대 PIN 자릿수 불일치: {len(pwd_digits)}자리 (기대 {expected_len}자리)")
            return False
        if not pwd_digits:
            self._log("  ❌ 현대 PIN 값 없음")
            return False

        self._log(f"  🔐 현대 PIN 입력: {'*' * len(pwd_digits)}자리")
        w_w, w_h = 1080, 2400
        try:
            size = self.driver.get_window_size()
            w_w, w_h = size['width'], size['height']
        except Exception:
            pass

        # 일반 전화 키패드 상대 좌표 (열, 행) — '1' 기준
        # 1 2 3
        # 4 5 6
        # 7 8 9
        #   0
        KEYPAD_RC = {
            '1': (0, 0), '2': (1, 0), '3': (2, 0),
            '4': (0, 1), '5': (1, 1), '6': (2, 1),
            '7': (0, 2), '8': (1, 2), '9': (2, 2),
            '0': (1, 3),
        }

        if roi:
            min_x_keypad = int(roi.get("min_x", 0))
            max_x_keypad = int(roi.get("max_x", w_w))
            min_y_keypad = int(roi.get("min_y", int(w_h * 0.28)))
            max_y_keypad = int(roi.get("max_y", int(w_h * 0.95)))
            roi_w = max(1, max_x_keypad - min_x_keypad)
            roi_h = max(1, max_y_keypad - min_y_keypad)
            # ROI 기준 키 간격 (현대비번 커팅 영역 내부)
            dx = max(80, int(roi_w / 3))
            dy = max(80, int(roi_h / 4.5))
            self._log(
                f"  📐 현대비번 ROI 적용: x={min_x_keypad}~{max_x_keypad} "
                f"y={min_y_keypad}~{max_y_keypad} dx={dx} dy={dy}"
            )
        else:
            min_x_keypad = 0
            max_x_keypad = w_w
            min_y_keypad = int(w_h * 0.28)
            max_y_keypad = int(w_h * 0.95)
            dx = int(w_w * 0.28)
            dy = int(w_h * 0.075)

        anchor_digit = None
        anchor_xy = None  # '1' 위치 (가상 원점)

        self._log(f"  ⏳ 현대 키패드 표시 대기 (1.5초)... (탐색 y={min_y_keypad}~{max_y_keypad})")
        time.sleep(1.5)

        MAX_DIGIT_RETRY = 3
        RETRY_INTERVAL = 0.8
        for idx, digit in enumerate(pwd_digits):
            img_path = IMG_HYUNDAI_NUMS.get(digit)
            if not img_path or not os.path.exists(img_path):
                self._log(f"  ⚠ 현대숫자 이미지 없음: {digit}.png → 실패")
                return False
            self._log(f"  🔢 {idx + 1}번째 자리 '{digit}' 클릭 시도")

            prefer_xy = None
            if anchor_xy is not None and digit in KEYPAD_RC:
                col, row = KEYPAD_RC[digit]
                prefer_xy = (anchor_xy[0] + col * dx, anchor_xy[1] + row * dy)
                self._log(f"    ℹ 키패드 예상좌표: {prefer_xy} (기준 '{anchor_digit}'={anchor_xy})")
            elif roi and digit in KEYPAD_RC and anchor_xy is None:
                # ROI만 있을 때: 좌상단을 '1' 가정한 초기 예상좌표
                col, row = KEYPAD_RC[digit]
                prefer_xy = (
                    min_x_keypad + int(dx * (col + 0.5)),
                    min_y_keypad + int(dy * (row + 0.35)),
                )
                self._log(f"    ℹ ROI 초기 예상좌표: {prefer_xy}")

            digit_ok = False
            for retry in range(1, MAX_DIGIT_RETRY + 1):
                coords = self._find_digit_coords(
                    img_path,
                    min_y_keypad,
                    max_y=max_y_keypad,
                    min_x=min_x_keypad,
                    max_x=max_x_keypad,
                    prefer_xy=prefer_xy,
                    prefer_radius=max(160, int(min(dx, dy) * 0.85)),
                    min_score=0.55,
                )
                if coords:
                    cx, cy = coords
                    # ROI 밖이면 거부
                    if not (min_x_keypad <= cx <= max_x_keypad and min_y_keypad <= cy <= max_y_keypad):
                        self._log(f"    ↩ ROI 밖 좌표 거부 ({cx},{cy})")
                        time.sleep(RETRY_INTERVAL)
                        continue
                    # 첫 숫자로 키패드 원점('1' 위치) 추정
                    if anchor_xy is None and digit in KEYPAD_RC:
                        col, row = KEYPAD_RC[digit]
                        anchor_xy = (cx - col * dx, cy - row * dy)
                        anchor_digit = digit
                        # ROI와 교차하는 범위로 더 좁힘
                        min_y_keypad = max(min_y_keypad, int(anchor_xy[1] - dy * 0.6))
                        max_y_keypad = min(max_y_keypad, int(anchor_xy[1] + dy * 3.8))
                        min_x_keypad = max(min_x_keypad, int(anchor_xy[0] - dx * 0.6))
                        max_x_keypad = min(max_x_keypad, int(anchor_xy[0] + dx * 2.6))
                        self._log(
                            f"    📌 키패드 원점 추정: '1'≈{anchor_xy} "
                            f"(from '{digit}'@{coords}), "
                            f"범위 x={min_x_keypad}~{max_x_keypad} y={min_y_keypad}~{max_y_keypad}"
                        )
                    elif anchor_xy is not None and digit in KEYPAD_RC:
                        col, row = KEYPAD_RC[digit]
                        exp_x = anchor_xy[0] + col * dx
                        exp_y = anchor_xy[1] + row * dy
                        if digit == '5' and (cx < anchor_xy[0] or cy < anchor_xy[1] + int(dy * 0.35)):
                            self._log(
                                f"    ↩ '5' 위치 거부 ({cx},{cy}) — "
                                f"'1'({anchor_xy})보다 오른쪽 아래여야 함 (예상≈{exp_x},{exp_y})"
                            )
                            time.sleep(RETRY_INTERVAL)
                            continue
                        if col > 0:
                            dx = max(int(dx * 0.7), abs(cx - anchor_xy[0]) // col)
                        if row > 0:
                            dy = max(int(dy * 0.7), abs(cy - anchor_xy[1]) // row)

                    ah.tap_by_coords(self.driver, cx, cy, self._log)
                    self._log(f"    ✅ '{digit}' 클릭 완료 ({coords}) [시도 {retry}회]")
                    digit_ok = True
                    time.sleep(0.5)
                    break
                self._log(f"    ↩ '{digit}' 인식 실패 ({retry}/{MAX_DIGIT_RETRY})")
                time.sleep(RETRY_INTERVAL)
            if not digit_ok:
                # 앵커/ROI 예상 좌표 강제 탭
                if prefer_xy is not None:
                    px = max(min_x_keypad, min(max_x_keypad, prefer_xy[0]))
                    py = max(min_y_keypad, min(max_y_keypad, prefer_xy[1]))
                    self._log(f"    ⚠ 인식 실패 → 예상좌표 강제 탭 ({px},{py})")
                    ah.tap_by_coords(self.driver, px, py, self._log)
                    time.sleep(0.5)
                    continue
                self._log(f"    ❌ '{digit}' 최종 인식 실패")
                return False
        self._log("  ✅ 현대 PIN 입력 완료")
        return True

    def _click_hyundai_pw_confirm(self) -> bool:
        """현대확인 이미지 클릭 (현대비번 확인 화면 감지 후)."""
        # 전체화면 템플릿은 클릭용이 아니라 화면 존재 확인용
        if os.path.exists(IMG_HYUNDAI_PW_CONFIRM_FULL):
            box = self._find_image_bbox(
                IMG_HYUNDAI_PW_CONFIRM_FULL,
                threshold=0.45,
                save_crop=False,
                crop_label="현대비번확인화면",
            )
            if box:
                self._log("  ✅ [현대비번 확인] 화면 감지됨 → 확인 버튼 탐색")

        if self._click_any_image_basic(IMG_HYUNDAI_CONFIRM, threshold=0.65, attempts=6, wait_after=5.0):
            return True
        # XPath 폴백
        for xp in [
            '//*[contains(@text,"확인")]',
            '//android.widget.Button[contains(@text,"확인")]',
        ]:
            try:
                if ah.element_exists(self.driver, xp, timeout=2):
                    el = self.driver.find_element(By.XPATH, xp)
                    if self._safe_click_element(el):
                        self._log(f"  ✅ 확인 XPath 클릭: {xp}")
                        time.sleep(2.0)
                        return True
            except Exception:
                continue
        return False

    def _verify_hyundai_order_complete(self) -> bool:
        """[22-14] 주문완료 되었습니다 존재 여부."""
        xpaths = [
            ORDER_COMPLETE_XPATH,
            '//android.widget.TextView[contains(@text,"주문완료 되었습니다")]',
            '//*[contains(@text,"주문완료 되었습니다")]',
        ]
        for xpath in xpaths:
            try:
                if ah.element_exists(self.driver, xpath, timeout=4):
                    self._log(f"✅ [22-14] 주문완료 확인: {xpath}")
                    return True
            except Exception:
                continue
        self._log("❌ [22-14] '주문완료 되었습니다' 미발견 → 실패")
        return False

    def _process_kb_card_payment(self) -> bool:
        """
        국민카드 결제: 다른결재 → 일반결재 → 카드선택(kb국민1/2) → 결재하기
        이후 PIN/비번 등 후속 작업 없이 종료 (성공 처리).
        """
        self._log("💳 [국민카드 결제] 프로세스 시작 (선택→결재하기 후 종료)")
        if self._skip_final_order_click():
            self._log("🖐 테스트/수동시작 모드 → 국민카드 결제 최종 단계 생략")
            return True

        # 1) 다른결재
        if not self._click_other_pay_button(max_scroll_attempts=20):
            self._log("❌ [국민카드] 다른결재 버튼 미발견")
            return False

        # 2) 일반결재 체크
        if not self._ensure_normal_pay_checked():
            return False

        # 3) 카드 드롭다운
        if not self._open_card_select_dropdown(brand="kb"):
            return False

        # 4) kb국민1/2 선택
        if not self._pick_kb_card():
            return False

        if self._card_placeholder_visible():
            self._log("❌ [국민카드] 카드가 아직 '카드를 선택해주세요' → 결재하기 클릭 안 함")
            return False

        # 5) 결재하기
        if not self._click_any_image_with_scroll(IMG_HYUNDAI_DO_PAY, threshold=0.70, max_scroll_attempts=8):
            self._log("  ⚠ [국민카드] 결재하기 이미지 미발견 → XPath 폴백")
            if not self._click_pay_button():
                self._log("❌ [국민카드] 결재하기 클릭 실패")
                return False

        self._log("✅ [국민카드] 결재하기 클릭 완료 → 후속 작업 없이 종료")
        return True

    def _process_hyundai_card_payment(self, second_password: str) -> bool:
        """[단계 22] 현대카드 결제."""
        self._log("💳 [현대카드 결제] 프로세스 시작")
        if self._skip_final_order_click():
            self._log("🖐 테스트/수동시작 모드 → 현대카드 결제 최종 단계 생략")
            return True

        # 22-1 다른결재 / 다른결재4 / 다른결재수단2
        if not self._click_other_pay_button(max_scroll_attempts=20):
            self._log("❌ [22-1] 다른결재 버튼 미발견")
            return False

        # 22-2 일반결재 체크
        if not self._ensure_normal_pay_checked():
            return False

        # 22-3 카드를 선택해주세요 드롭다운 열기
        if not self._open_card_select_dropdown():
            return False

        # 22-4 목록에서 현대 클릭 + 선택 검증 (상단 y=193 오탐 금지)
        if not self._pick_hyundai_card():
            return False

        if self._card_placeholder_visible():
            self._log("❌ [22-4] 카드가 아직 '카드를 선택해주세요' → 결제하기 클릭 안 함")
            return False

        # 22-5 결재하기.png ~ 결재하기4.png
        if not self._click_any_image_with_scroll(IMG_HYUNDAI_DO_PAY, threshold=0.70, max_scroll_attempts=8):
            self._log("  ⚠ [22-5] 결재하기 이미지 미발견 → XPath 폴백")
            if not self._click_pay_button():
                self._log("❌ [22-5] 결재하기 클릭 실패")
                return False

        # 22-6 현대핀1~5.png 클릭 (PIN번호 결제 버튼), 이후 최대 8초 대기
        pin_btn_clicked = self._click_any_image_basic(IMG_HYUNDAI_PIN_BTN, threshold=0.70, attempts=6, wait_after=2.0)
        if not pin_btn_clicked:
            # XPath 폴백: "PIN번호 결제" 텍스트
            pin_xpaths = [
                '//*[contains(@text,"PIN번호 결제")]',
                '//*[contains(@text,"PIN번호")]',
                '//*[contains(@text,"핀번호")]',
                '//*[contains(@content-desc,"PIN")]',
            ]
            for xp in pin_xpaths:
                try:
                    if ah.element_exists(self.driver, xp, timeout=2):
                        el = self.driver.find_element(By.XPATH, xp)
                        if self._safe_click_element(el):
                            self._log(f"  ✅ [22-6] XPath 폴백으로 핀 버튼 클릭: {xp}")
                            pin_btn_clicked = True
                            break
                except Exception:
                    continue
        if not pin_btn_clicked:
            self._log("❌ [22-6] 현대핀 버튼 미발견")
            return False
        self._log("  ⏳ [22-6] 핀 입력창 표시 대기 (8초)...")
        time.sleep(8.0)

        # 22-7 핀입력1~5.png 인식 후 탭 → 키보드/키패드 활성화
        pin_names = " / ".join(n for _, n in IMG_HYUNDAI_PIN_INPUT)
        self._log(f"  🔍 [22-7] 핀입력 탐색: {pin_names}")
        pin_field_tapped = self._click_any_image_basic(
            IMG_HYUNDAI_PIN_INPUT, threshold=0.60, attempts=4, wait_after=1.5
        )

        if not pin_field_tapped:
            # 방법2: EditText XPath
            for xp in ['//android.widget.EditText', '//*[@class="android.widget.EditText"]']:
                try:
                    if ah.element_exists(self.driver, xp, timeout=1):
                        el = self.driver.find_element(By.XPATH, xp)
                        el.click()
                        self._log(f"  ✅ [22-7] EditText 탭: {xp}")
                        pin_field_tapped = True
                        time.sleep(1.5)
                        break
                except Exception:
                    continue

        if not pin_field_tapped:
            # 방법3: 화면 중앙 상단(○ 예상 위치) 강제 탭
            w, h = self._get_window_size()
            tap_x, tap_y = w // 2, int(h * 0.33)
            self._log(f"  ⚠ [22-7] 이미지/XPath 미발견 → 화면 중앙({tap_x},{tap_y}) 강제 탭")
            ah.tap_by_coords(self.driver, tap_x, tap_y, self._log)
            time.sleep(1.5)

        # 22-8 현대숫자 6자리 PIN – 현대비번.png ROI 커팅 후 입력
        if not self._input_hyundai_digits_with_fallback(HYUNDAI_PIN6, expected_len=6, use_keypad_crop=True):
            self._log("❌ [22-8] 현대카드 6자리 PIN 입력 실패")
            return False

        # 22-9 현대확인 / 현대비번 확인
        if not self._click_hyundai_pw_confirm():
            self._log("❌ [22-9] 현대확인 이미지 미발견")
            return False

        # 22-10 현대결제하기1~4.png
        if not self._click_any_image_basic(IMG_HYUNDAI_PAY_NOW, threshold=0.70, attempts=6, wait_after=2.0):
            if not self._click_any_image_with_scroll(IMG_HYUNDAI_PAY_NOW, threshold=0.70, max_scroll_attempts=6):
                self._log("❌ [22-10] 현대결제하기 이미지 미발견")
                return False

        # 22-10.5 안전한/추가인증 팝업 → 안전확인1~2 클릭 후 카드비번 진행
        if not self._handle_hyundai_safe_auth_popup():
            self._log("❌ [22-10.5] 안전인증 확인 클릭 실패")
            return False

        # 22-11 현대카드비번1~4.png 클릭 → 비번 입력창
        if not self._click_any_image_basic(IMG_HYUNDAI_CARD_PW, threshold=0.70, attempts=6, wait_after=2.0):
            self._log("❌ [22-11] 현대카드비번 이미지 미발견")
            return False
        self._log("  ⏳ [22-11] 카드비번 키패드 표시 대기 (2초)...")
        time.sleep(2.0)

        # 22-12 현대비번.png 인식/커팅 → 2차비밀번호 4자리 입력
        pin4 = ''.join(filter(str.isdigit, second_password or ""))
        self._log(f"  🔐 [22-12] 2차비밀번호 입력 (현대비번 ROI 커팅 강화)")
        if not self._input_hyundai_digits_with_fallback(pin4, expected_len=4, use_keypad_crop=True):
            self._log("❌ [22-12] 2차비밀번호 4자리 입력 실패")
            return False

        # 22-13 현대비번 확인 / 현대확인, 5초 + 7초 대기
        if not self._click_hyundai_pw_confirm():
            self._log("❌ [22-13] 현대확인 이미지 미발견")
            return False
        self._log("  ⏳ [22-13] 추가 7초 대기...")
        time.sleep(7.0)

        # 22-14 주문완료 확인
        return self._verify_hyundai_order_complete()

    # ─── 주문 루프 ────────────────────────────────────────────────────────────

    def _order_loop(self):
        """[단계 2~19] 결재목록 루프 주문 처리"""
        self._log(f"📋 주문 루프 시작 (폰ID 필터: {self.device_id})")
        try:
            extra = self.order_manager.describe_pending_filter(self.device_id)
            self._log(f"  ℹ {extra}")
        except Exception:
            pass

        while not self._stop_event.is_set():
            # OrderManager 버전에 따라 device_id 인자 지원 여부 호환
            try:
                row = self.order_manager.get_next_pending(device_id=self.device_id)
            except TypeError:
                # 구버전: device_id 미지원 → 전체 pending에서 기기ID 필터
                try:
                    rows = self.order_manager.get_pending_rows(device_id=self.device_id)
                except TypeError:
                    rows = self.order_manager.get_pending_rows()
                    rows = [r for r in rows if getattr(r, "device_id", "") == self.device_id]
                row = rows[0] if rows else None
            if not row:
                try:
                    extra = self.order_manager.describe_pending_filter(self.device_id)
                    self._log(f"⚠ 해당 기기 미처리 행 없음 ({extra})")
                except Exception:
                    pass
                self._log("✅ 모든 주문 처리 완료 (해당 기기 대상)")
                break

            self._log(
                f"📌 처리 중: row={row.row_index}, keyword={row.search_keyword!r}, "
                f"폰ID={row.device_id!r}, 결재방식={row.payment_method!r}"
            )
            self._set_status(f"주문 중: {row.search_keyword}")
            self.has_dismissed_payment_benefit = False

            try:
                # 23. 실패 시 재작업하지 않음 (1회만 시도)
                success = self._process_order_with_timeout(row)
            except Exception as fatal_err:
                self.order_manager.mark_failed(row.row_index)
                self._log(f"❌ 치명적 오류: {fatal_err}")
                raise

            if success:
                if row.payment_method == "무통장":
                    # 무통장: 주문번호 확인되어야 최종 성공 처리
                    try:
                        if self._skip_final_order_click():
                            mode = "수동시작" if self.manual_mode else "테스트 모드"
                            self._log(f"🖐 [{mode}] 주문번호 캡처/확인 생략 → 엑셀 Y 기록")
                            order_confirmed = True
                        else:
                            order_confirmed = self._capture_and_log_bank_transfer(row)
                    except Exception as e:
                        self._log(f"❌ 무통장 스크린샷/로그 처리 중 예외 발생: {e}")
                        order_confirmed = False

                    if order_confirmed:
                        self.order_manager.mark_success(row.row_index)
                        if self.manual_mode:
                            self._log(f"✅ [수동시작] 배송지 선택 완료 → Y 기록: {row.search_keyword}")
                        else:
                            self._log(f"✅ 무통장 주문 성공 (주문번호 확인됨): {row.search_keyword} → Y 기록")
                    else:
                        self.order_manager.mark_failed(row.row_index)
                        self._log(f"❌ 무통장 주문번호 미확인 → F 기록: {row.search_keyword}")
                        self._log("⏹ F 기록 → 다음 작업 없이 워커 종료")
                        break
                else:
                    self.order_manager.mark_success(row.row_index)
                    if self.manual_mode:
                        self._log(f"✅ [수동시작] 배송지 선택 완료 → Y 기록: {row.search_keyword}")
                    else:
                        self._log(f"✅ 주문 성공: {row.search_keyword} → Y 기록")

                # 수동시작: Y 기록 후 해당 기기 작업 종료 (다음 행 계속하지 않음)
                if self.manual_mode:
                    self._log("🖐 [수동시작] 엑셀 Y 기록 완료 → 프로그램(워커) 종료")
                    break

                self._log("⏳ [주문 성공] 완료 후 30초 대기 중...")
                import gc
                gc.collect()
                time.sleep(30)
            else:
                self.order_manager.mark_failed(row.row_index)
                self._log(f"❌ 주문 실패: {row.search_keyword} → F 기록")
                self._log("⏹ F 기록 → 다음 작업 없이 워커 종료")
                import gc
                gc.collect()
                break

        self._log("📋 주문 루프 종료")



    def _process_order_with_timeout(self, row: OrderRow) -> bool:
        """주문 1건 처리 (타임아웃 적용)"""
        result = [False]
        exception = [None]

        def task():
            try:
                result[0] = self._process_one_order(row)
            except Exception as e:
                exception[0] = e

        t = threading.Thread(target=task, daemon=True)
        t.start()
        t.join(timeout=TASK_TIMEOUT_SEC)

        if t.is_alive():
            self._log(f"⏰ 타임아웃 ({TASK_TIMEOUT_SEC}초) - 다음 행으로")
            return False

        if exception[0]:
            FATAL_PATTERNS = [
                "A session is either terminated or not started",
                "UiAutomation not connected",
                "instrumentation process is not running",
            ]
            err_str = str(exception[0])
            if any(p in err_str for p in FATAL_PATTERNS):
                raise exception[0]
            self._log(f"❌ 처리 중 예외: {exception[0]}")
            return False

        return result[0]

    def _process_one_order(self, row: OrderRow) -> bool:
        """
        주문 1건 전체 흐름 (단계 3~19)
        각 주문마다 처음부터 메인 페이지로 이동하여 처리
        """
        # [단계 3~6] 매 주문마다 메인→계정전환(3.2)→스토어→마이쇼핑 재진입
        if not self._go_main_and_enter_store(login_id=row.login_id):
            self._log("❌ 계정 전환 또는 메인 페이지/스토어/마이쇼핑 진입 실패 -> 다음 레드로 이동")
            return False

        # [단계 7] 마이쇼핑 검색 버튼 클릭
        if not self._click_search_in_my_shopping():
            self._log("❌ 검색 버튼 클릭 실패")
            return False

        # [단계 8] 검색어 입력
        if not self._input_search_keyword(row.search_keyword):
            self._log("❌ 검색어 입력 실패")
            return False

        # [단계 9] 검색 실행
        if not self._click_search_button():
            self._log("❌ 검색 실행 실패")
            return False

        # [단계 10] 상품 매칭 클릭
        if not self._click_product(row.seller_name, row.product_name):
            self._log("❌ 상품 매칭 실패")
            return False

        # [단계 11] 구매하기 버튼
        self._click_buy_button()
        time.sleep(1.0)

        # [단계 12] 체크박스/옵션 항목 클릭
        self._click_checkbox(row.product_name)
        time.sleep(0.8)

        # [단계 13] 바로구매 클릭
        if not self._click_buy_now():
            self._log("❌ 바로구매 클릭 실패")
            return False

        # [단계 14 & 16] 배송지 확인 및 선택
        if self._check_current_delivery_address(row.recipient_name, row.phone):
            self._log(f"✅ 목표 배송지 '{row.recipient_name}'가 이미 선택되어 있습니다 (변경 불필요)")
        else:
            self._log("🔄 목표 배송지가 선택되어 있지 않아 변경을 시도합니다.")
            if not self._click_change_button():
                self._log("❌ 변경 버튼 클릭 실패 (주문/결제 화면 미진입 가능)")
                return False

            # [단계 15] 스크롤 다운 → 제거 (배송지 목록이 바로 표시되므로 불필요)

            if not self._select_delivery_address(row.recipient_name, row.phone):
                self._log("❌ 배송지 선택 실패")
                return False

        # 수동시작: 배송지 선택(결제창 복귀)까지 완료하면 결제 단계 생략 → Y 기록
        if self.manual_mode:
            self._log("🖐 [수동시작] 배송지 선택 완료 → 결제 단계 생략, Y 기록 후 종료")
            return True

        # [단계 16.5] 배송메모 처리 (배송메모.png 인식 시 '선택안함' 1회 클릭)
        self._handle_delivery_memo()

        # [단계 17] 전액사용 클릭 등 결제 방식 분기
        if self._is_kb_card_payment(row.payment_method):
            if not self._process_kb_card_payment():
                self._log("❌ 국민카드 결제 진행 실패")
                return False
        elif self._is_hyundai_card_payment(row.payment_method):
            second_pw = getattr(row, "second_password", "") or ""
            if not self._process_hyundai_card_payment(second_pw):
                self._log("❌ 현대카드 결제 진행 실패")
                return False
        elif row.payment_method == "무통장":
            if not self._process_bank_transfer():
                self._log("❌ 무통장 결제 진행 실패")
                return False
        elif row.payment_method == "머니":
            if not self._process_money_payment(row.password):
                self._log("❌ 머니 결제 진행 실패")
                return False
        else:
            # 포인트 또는 기본 결제
            if not self._click_full_use():
                self._log("❌ 전액사용 버튼 클릭 실패")
                return False
            # [단계 18] 결제하기 버튼
            if not self._click_pay_button():
                self._log("❌ 결제하기 클릭 실패")
                return False
            # [단계 19] 비밀번호 입력
            if row.password:
                if not self._input_password(row.password):
                    self._log("❌ 비밀번호 입력 실패")
                    return False
            else:
                self._log("  ℹ 비밀번호 없음 → 건너뜀")

        self._log(f"✅ 주문 완료: {row.search_keyword}")
        return True

    # ─── 이미지 인식 (naver_worker.py 동일 로직 재구현) ─────────────────────

    def _find_image_coords(self, template_path: str,
                           threshold: float = 0.75,
                           min_x: Optional[int] = None,
                           max_x: Optional[int] = None,
                           min_y: Optional[int] = None,
                           max_y: Optional[int] = None) -> Optional[tuple]:
        """멀티스케일 OpenCV 템플릿 매칭으로 이미지 위치 탐색"""
        try:
            import cv2
            import numpy as np
            from PIL import Image
            import io
        except ImportError:
            self._log("  [이미지 매칭] cv2/numpy/PIL 라이브러리 미설치")
            return None

        try:
            # 화면 캡처
            screenshot_png = self._get_screenshot()
            screenshot_pil = Image.open(io.BytesIO(screenshot_png))
            screen_bgr = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)
            screen_gray = cv2.cvtColor(screen_bgr, cv2.COLOR_BGR2GRAY)
            screen_h, screen_w = screen_gray.shape

            # min_x, max_x, min_y, max_y 제약이 있다면 탐색 영역 밖을 검정색으로 지워서 오탐 방지
            if min_y is not None and min_y > 0:
                screen_gray[:min_y, :] = 0
            if max_y is not None and max_y < screen_h:
                screen_gray[max_y:, :] = 0
            if min_x is not None and min_x > 0:
                screen_gray[:, :min_x] = 0
            if max_x is not None and max_x < screen_w:
                screen_gray[:, max_x:] = 0

            if not os.path.exists(template_path):
                self._log(f"  [이미지 매칭] 템플릿 없음: {template_path}")
                return None

            template_bgr = cv2.imdecode(
                np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR
            )
            if template_bgr is None:
                self._log(f"  [이미지 매칭] 템플릿 로드 실패: {template_path}")
                return None

            template_gray = cv2.cvtColor(template_bgr, cv2.COLOR_BGR2GRAY)
            t_h, t_w = template_gray.shape

            best_score = -1
            best_loc   = None
            best_tw    = t_w
            best_th    = t_h

            # 표준 Grayscale 매칭 (CLAHE 왜곡 없이 템플릿 원본 정밀 비교)
            # 다양한 모바일 해상도(DPI) 대응을 위한 멀티스케일 (0.55 ~ 1.65x, 12단계)
            scales = np.linspace(0.55, 1.65, 12)
            for scale in scales:
                new_w = int(t_w * scale)
                new_h = int(t_h * scale)
                if new_w >= screen_w or new_h >= screen_h:
                    continue
                if new_w < 10 or new_h < 5:
                    continue

                resized_templ = cv2.resize(template_gray, (new_w, new_h), interpolation=cv2.INTER_AREA)

                # 표준 TM_CCOEFF_NORMED 적용 (가장 정밀하고 오탐이 없는 방식)
                try:
                    r = cv2.matchTemplate(screen_gray, resized_templ, cv2.TM_CCOEFF_NORMED)
                    _, max_val, _, max_loc = cv2.minMaxLoc(r)
                    if max_val > best_score:
                        best_score = max_val
                        best_loc   = max_loc
                        best_tw    = new_w
                        best_th    = new_h
                except Exception:
                    pass

            if best_score >= threshold and best_loc is not None:
                cx = best_loc[0] + best_tw // 2
                cy = best_loc[1] + best_th // 2
                if min_x is not None and cx < min_x:
                    self._log(f"  ❌ [이미지 매칭] 매칭 좌표 x({cx}) < min_x({min_x}) → 무효 처리")
                    return None
                if max_x is not None and cx > max_x:
                    self._log(f"  ❌ [이미지 매칭] 매칭 좌표 x({cx}) > max_x({max_x}) → 무효 처리")
                    return None
                if min_y is not None and cy < min_y:
                    self._log(f"  ❌ [이미지 매칭] 매칭 좌표 y({cy}) < min_y({min_y}) → 무효 처리")
                    return None
                if max_y is not None and cy > max_y:
                    self._log(f"  ❌ [이미지 매칭] 매칭 좌표 y({cy}) > max_y({max_y}) → 무효 처리")
                    return None
                self._log(f"  🎯 [이미지 매칭] 발견! 중심좌표: ({cx}, {cy}), 점수: {best_score:.4f}")

                # 인식된 부분만 크롭하여 저장 ('인식' 폴더: 프로젝트 최상위 루트)
                try:
                    root_dir = os.path.dirname(os.path.abspath(__file__))
                    recognize_dir = os.path.join(root_dir, "인식")
                    os.makedirs(recognize_dir, exist_ok=True)
                    file_name = os.path.basename(template_path)
                    name_no_ext = os.path.splitext(file_name)[0]
                    timestamp = time.strftime("%Y%m%d_%H%M%S")
                    
                    crop_name = f"crop_{name_no_ext}_({cx},{cy})_{best_score:.3f}_{timestamp}.png"
                    crop_path = os.path.join(recognize_dir, crop_name)
                    x1, y1 = max(0, best_loc[0]), max(0, best_loc[1])
                    x2, y2 = min(screen_bgr.shape[1], x1 + best_tw), min(screen_bgr.shape[0], y1 + best_th)
                    cropped = screen_bgr[y1:y2, x1:x2]
                    if cropped.size > 0:
                        cv2.imwrite(crop_path, cropped)
                        self._log(f"  📸 [인식 캡처 저장 완료] {crop_path}")
                except Exception as save_err:
                    pass

                return cx, cy
            else:
                self._log(f"  ❌ [이미지 매칭] 실패 (점수 {best_score:.4f} < {threshold})")
                return None
        except Exception as e:
            self._log(f"  [이미지 매칭] 오류: {e}")
            return None

    def _get_screenshot(self) -> bytes:
        """화면 꺼짐 자동 복구 스크린샷"""
        return ah.get_screenshot_safe(self.driver, self.device_id, log_callback=self._log)

    # ─── 스크롤 유틸 ──────────────────────────────────────────────────────────

    def _capture_screen_fingerprint(self):
        """ADB screencap 기반 화면 지문 (64x128 그레이스케일 축소본). 실패 시 None"""
        try:
            import cv2
            import numpy as np
            res = _run_cmd(
                ["adb", "-s", self.device_id, "exec-out", "screencap", "-p"],
                capture_output=True, timeout=8
            )
            if not res.stdout or len(res.stdout) < 100:
                return None
            arr = np.frombuffer(res.stdout, np.uint8)
            img = cv2.imdecode(arr, cv2.IMREAD_GRAYSCALE)
            if img is None:
                return None
            return cv2.resize(img, (64, 128), interpolation=cv2.INTER_AREA)
        except Exception:
            return None

    @staticmethod
    def _fingerprints_differ(before, after) -> bool:
        """두 화면 지문이 유의미하게 다른지 판정 (평균 픽셀 차 > 1.5)"""
        try:
            import cv2
            diff = cv2.absdiff(before, after)
            return float(diff.mean()) > 1.5
        except Exception:
            return True  # 비교 불가 시 '변화 있음'으로 간주 (기존 동작 유지)

    def _scroll_gesture(self, direction: str, distance_ratio: float) -> bool:
        """UiAutomator2 네이티브 'mobile: scrollGesture' 수행.

        시스템 제스처 라이브러리를 사용하므로 터치 슬롭이 보장되어
        스크롤 도중 요소 클릭/롱클릭이 절대 발생하지 않습니다.
        화면 중앙 밴드(세로 25%~75%)에서만 제스처를 수행하여
        상단 헤더/하단 고정 결제바를 건드리지 않습니다.

        Returns:
            제스처 수행 성공 여부 (드라이버 미지원/오류 시 False)
        """
        try:
            w, h = self._get_window_size()
            top = int(h * 0.25)
            area_h = int(h * 0.50)
            percent = max(0.15, min(1.0, (h * distance_ratio) / area_h))
            self.driver.execute_script('mobile: scrollGesture', {
                'left': int(w * 0.10),
                'top': top,
                'width': int(w * 0.80),
                'height': area_h,
                'direction': direction,
                'percent': percent,
                'speed': 1200,  # px/s. 낮은 속도 = 관성(fling) 없는 부드러운 드래그
            })
            return True
        except Exception:
            return False

    def _adb_swipe(self, sx: int, sy: int, ex: int, ey: int, duration_ms: int = 500):
        """ADB input swipe (긴 duration = 관성 없는 드래그 스크롤)"""
        try:
            _run_cmd(
                ["adb", "-s", self.device_id, "shell", "input", "swipe",
                 str(sx), str(sy), str(ex), str(ey), str(duration_ms)],
                capture_output=True, timeout=5
            )
        except Exception:
            pass

    def _get_window_size(self):
        w, h = 1080, 2400
        try:
            size = self.driver.get_window_size()
            w, h = size['width'], size['height']
        except Exception:
            pass
        return w, h

    def _scroll_down_fast(self, distance_ratio: float = 0.28):
        """상품 리스트용 빠른 스크롤 (지문 검증/재시도 생략)."""
        w, h = self._get_window_size()
        start_y = int(h * 0.72)
        end_y = max(100, int(start_y - (h * distance_ratio)))
        self._adb_swipe(w // 2, start_y, w // 2, end_y, duration_ms=280)
        time.sleep(0.2)

    def _scroll_down_safe(self, distance_ratio: float = 0.14):
        """
        결제/주문 화면용 안전 스크롤.
        - 웹뷰 콘텐츠 밴드(약 32%~62%)만 드래그 → 하단 네비/앱 밖으로 오버스크롤 방지
        - 중간 거리 + 느린 duration → fling/관성으로 창 밖 이탈 방지
        - 지문 재시도(시작점 변경) 없음 → 과도한 연속 스와이프 방지
        """
        w, h = self._get_window_size()
        # 콘텐츠 중앙 부근에서 조금 더 넓게 올림 (네이버 앱 WebView 내부)
        start_y = int(h * 0.62)
        end_y = max(int(h * 0.32), int(start_y - (h * max(0.08, min(0.18, distance_ratio)))))
        self._log(
            f"  ↕ [안전스크롤] down ({start_y}→{end_y}, ratio≈{distance_ratio:.2f})"
        )
        # 네이티브 제스처 우선 (느린 속도)
        try:
            area_top = int(h * 0.32)
            area_h = int(h * 0.36)
            percent = max(0.18, min(0.48, (h * distance_ratio) / max(1, area_h)))
            self.driver.execute_script('mobile: scrollGesture', {
                'left': int(w * 0.12),
                'top': area_top,
                'width': int(w * 0.76),
                'height': area_h,
                'direction': 'down',
                'percent': percent,
                'speed': 750,
            })
        except Exception:
            self._adb_swipe(w // 2, start_y, w // 2, end_y, duration_ms=950)
        time.sleep(0.85)

    def _scroll_down(self, distance_ratio: float = 0.20):
        """아래로 미세 스크롤 (요소 클릭이 발생하지 않는 방식).

        1차: UiAutomator2 네이티브 scrollGesture (클릭 이벤트 미발생 보장)
        폴백/재시도: ADB 장거리 저속 드래그 (이동 거리가 터치 슬롭을 크게
        초과하므로 클릭으로 인식되지 않음)

        스크롤 후 화면 지문을 비교하여 실제로 화면이 움직였는지 검증하고,
        변화가 없으면 시작점(Y)을 바꿔 재시도합니다.
        (결제화면의 가로 스크롤 카드영역/드롭다운 오버레이 등이 세로 스와이프를
        가로채 스크롤이 무시되는 현상 대응)
        """
        w, h = self._get_window_size()
        before = self._capture_screen_fingerprint()

        # 1차: 네이티브 scrollGesture, 미지원 시 ADB 드래그
        if not self._scroll_gesture("down", distance_ratio):
            start_y = int(h * 0.72)
            end_y = max(100, int(start_y - (h * distance_ratio)))
            self._adb_swipe(w // 2, start_y, w // 2, end_y, duration_ms=700)

        if before is None:
            return
        time.sleep(0.5)
        after = self._capture_screen_fingerprint()
        if after is None or self._fingerprints_differ(before, after):
            return

        # 화면 무변화 → 시작점을 바꿔 ADB 드래그 스와이프로 재시도
        for retry_idx, start_ratio in enumerate((0.60, 0.50), start=1):
            self._log(f"  ⚠ [스크롤 다운] 화면 변화 없음 → 시작점 변경 재시도 ({retry_idx}/2, y={int(start_ratio*100)}%)")
            start_y = int(h * start_ratio)
            end_y = max(100, int(start_y - (h * distance_ratio)))
            self._adb_swipe(w // 2, start_y, w // 2, end_y, duration_ms=700)
            time.sleep(0.7)
            after = self._capture_screen_fingerprint()
            if after is None or self._fingerprints_differ(before, after):
                self._log("  ✅ [스크롤 다운] 재시도 후 화면 이동 확인")
                return

        self._log("  ⚠ [스크롤 다운] 재시도에도 화면이 움직이지 않음 (페이지 끝 또는 스크롤 불가 상태)")

    def _scroll_up(self, distance_ratio: float = 0.4):
        """위로 스크롤 (화면을 아래로 내림). 요소 클릭 미발생 방식 + 무변화 시 재시도"""
        w, h = self._get_window_size()
        before = self._capture_screen_fingerprint()

        # 1차: 네이티브 scrollGesture, 미지원 시 ADB 드래그
        if not self._scroll_gesture("up", distance_ratio):
            start_y = int(h * 0.3)
            end_y = min(h - 100, int(start_y + (h * distance_ratio)))
            self._adb_swipe(w // 2, start_y, w // 2, end_y, duration_ms=700)

        if before is None:
            return
        time.sleep(0.5)
        after = self._capture_screen_fingerprint()
        if after is None or self._fingerprints_differ(before, after):
            return

        for retry_idx, start_ratio in enumerate((0.42, 0.55), start=1):
            self._log(f"  ⚠ [스크롤 업] 화면 변화 없음 → 시작점 변경 재시도 ({retry_idx}/2, y={int(start_ratio*100)}%)")
            start_y = int(h * start_ratio)
            end_y = min(h - 100, int(start_y + (h * distance_ratio)))
            self._adb_swipe(w // 2, start_y, w // 2, end_y, duration_ms=700)
            time.sleep(0.7)
            after = self._capture_screen_fingerprint()
            if after is None or self._fingerprints_differ(before, after):
                self._log("  ✅ [스크롤 업] 재시도 후 화면 이동 확인")
                return

        self._log("  ⚠ [스크롤 업] 재시도에도 화면이 움직이지 않음 (페이지 처음 또는 스크롤 불가 상태)")


    # ─── 로그/상태 ────────────────────────────────────────────────────────────

    def _log(self, message: str):
        full_msg = f"[{self.device_id}] {message}"
        if self._log_cb:
            self._log_cb(self.device_id, message)
        else:
            print(full_msg)
        # 로그 파일 기록
        try:
            import datetime
            today_folder = datetime.datetime.now().strftime("%Y%m%d")
            log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs", today_folder)
            os.makedirs(log_dir, exist_ok=True)
            log_path = os.path.join(log_dir, f"naver_order_기기{self.machine_num}_{self.device_id}.log")
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{now}] {full_msg}\n")
        except Exception:
            pass

    def _set_status(self, status: str):
        if self._status_cb:
            self._status_cb(self.device_id, status)
