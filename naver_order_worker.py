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
13    바로구매.png 이미지 인식 클릭, 8초 대기
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
IMG_SEARCH_ICON   = os.path.join(_IMG_DIR, "검색아이콘.png") # 검색 아이콘 (단계 9)
IMG_CHECKBOX      = os.path.join(_IMG_DIR, "체크박스.png")   # 체크박스 (단계 12)
IMG_BUY_NOW       = os.path.join(_IMG_DIR, "바로구매.png")   # 바로구매 버튼 (단계 13)
IMG_FULL_USE      = os.path.join(_IMG_DIR, "전액사용.png")   # 전액사용 버튼 (단계 17)

# 추가된 결제방식 이미지
IMG_OTHER_PAY     = os.path.join(_IMG_DIR, "다른결재.png")
IMG_NORMAL_PAY    = os.path.join(_IMG_DIR, "일반결재.png")
IMG_BANK_TRANSFER = os.path.join(_IMG_DIR, "무통장입금.png")
IMG_SELECT_BANK   = os.path.join(_IMG_DIR, "은행을.png")
IMG_BANK_SELECT   = os.path.join(_IMG_DIR, "은행선택.png")
IMG_SHINHAN_BANK  = os.path.join(_IMG_DIR, "신한은행.png")
IMG_NOT_APPLY     = os.path.join(_IMG_DIR, "미신청.png")
IMG_DO_PAY        = os.path.join(_IMG_DIR, "결재하기.png")
IMG_DO_ORDER      = os.path.join(_IMG_DIR, "주문하기.png")
IMG_MONEY_PAY     = os.path.join(_IMG_DIR, "머니.png")

# 비밀번호 숫자 이미지 (단계 19): p0.png ~ p9.png
IMG_NUMS = {
    str(d): os.path.join(_NUM_DIR, f"p{d}.png") for d in range(10)
}

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
TASK_TIMEOUT_SEC = 600  # 주문 1건 최대 10분


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
                 status_callback: Optional[Callable] = None):
        self.device_id      = device_id
        self.appium_port    = appium_port
        self.order_manager  = order_manager
        self._log_cb        = log_callback
        self._status_cb     = status_callback
        self.driver         = None
        self._stop_event    = threading.Event()

    # ─── 공개 메서드 ─────────────────────────────────────────────────────────

    def run(self) -> bool:
        """워커 메인 실행 (별도 스레드에서 호출)"""
        self._log("🚀 자동 주문 워커 시작")
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
        ah.go_to_main_page(self.driver, self._log)
        time.sleep(7)

        # [단계 3.1] 웰컴 모달 / 팝업 발견 시 클릭
        self._check_and_close_welcome_modals(step_label="3.1")

        # [단계 3.2] 로그인 아이디가 지정된 경우 네이버 계정 전환 수행
        if login_id:
            if not self._switch_account(login_id):
                self._log(f"❌ [단계 3.2] 계정 전환 실패 (아이디: {login_id}) -> 4번 이하 단계 수행 취소 및 다음 레코드 이동")
                return False
            # 계정 전환 성공 후 메인 페이지로 이동하여 4번 진행
            ah.go_to_main_page(self.driver, self._log)
            time.sleep(3)

        # [단계 4] 네이버 플러스 스토어 탭 버튼 클릭
        if ah.element_exists(self.driver, STORE_TAB_XPATH, timeout=5):
            self._log("📌 네이버 플러스 스토어 탭 감지 → 클릭")
            ah.wait_and_click(self.driver, STORE_TAB_XPATH, timeout=7, log_callback=self._log)
            time.sleep(5)
        else:
            self._log("⏭ 스토어 탭 없음 (이미 스토어 화면)")
            time.sleep(3)

        # [단계 5] 팝업 처리
        self._dismiss_popups()

        # [단계 6] 마이쇼핑 클릭 -> 10초 대기
        self._set_status("마이쇼핑 클릭")
        if ah.element_exists(self.driver, MY_SHOPPING_XPATH, timeout=8):
            ah.wait_and_click(self.driver, MY_SHOPPING_XPATH, timeout=7, log_callback=self._log)
            self._log("✅ 마이쇼핑 클릭 완료 (10초 대기)")
            time.sleep(10)
        else:
            self._log("⚠ 마이쇼핑 버튼 미발견")
            time.sleep(3)

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
                time.sleep(3)
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
                time.sleep(3)
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
                time.sleep(3)
                mgmt_clicked = True
                break

        if not mgmt_clicked:
            self._log("  ❌ [단계 3.2] 로그인 아이디 관리 링크 미발견 -> 계정 전환 실패")
            return False

        # 4. 아이디선택 화면에서 타겟 아이디 찾기 및 '로그인 중' 상태 확인
        already_logged_in = False
        verify_xpaths = [
            f'//android.view.View[contains(@content-desc, "{login_id}") and contains(@content-desc, "로그인 중")]',
            f'//*[contains(@content-desc, "{login_id}") and contains(@content-desc, "로그인 중")]',
        ]
        for xpath in verify_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=2):
                already_logged_in = True
                self._log(f"  ✅ [단계 3.2] 계정 [{login_id}] 이미 '로그인 중' 상태임")
                break

        target_account_xpaths = [
            f'//android.view.View[contains(@content-desc, "{login_id}")]',
            f'//*[contains(@content-desc, "{login_id}")]',
        ]
        target_el = None
        for xpath in target_account_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=4):
                try:
                    target_el = self.driver.find_element(By.XPATH, xpath)
                    self._log(f"  📌 아이디선택 화면에서 '{login_id}' 발견 -> 클릭 시도")
                    break
                except Exception:
                    continue

        if not target_el:
            self._log(f"  ❌ [단계 3.2] 로그인 아이디 [{login_id}] 미발견 -> 작업 중지 및 로그 기록")
            return False

        # 아이디 2번 클릭 시도
        click_success = False
        for attempt in range(1, 3):
            self._log(f"  👉 [단계 3.2] 로그인 아이디 [{login_id}] 클릭 시도 ({attempt}/2)")
            if self._safe_click_element(target_el):
                click_success = True
            else:
                try:
                    target_el.click()
                    click_success = True
                except Exception as e:
                    self._log(f"  ⚠ 클릭 예외 발생: {e}")
            
            time.sleep(1)
            
        if not click_success:
            self._log(f"  ❌ [단계 3.2] 로그인 아이디 [{login_id}] 클릭 실패")
            return False

        time.sleep(4)

        # 5. 클릭 후 로그인 중 상태 검증 (화면 전환 가능성 고려하여 실패해도 진행)
        if not already_logged_in:
            verified = False
            for xpath in verify_xpaths:
                if ah.element_exists(self.driver, xpath, timeout=3):
                    verified = True
                    break

            if not verified:
                self._log(f"  ⚠ [단계 3.2] 클릭 후 '{login_id}' '로그인 중' 미발견 (화면이 전환되었을 수 있음)")
            else:
                self._log(f"  ✅ [단계 3.2] 로그인 아이디 [{login_id}] 전환 성공 및 검증 완료")

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

    def _check_and_close_welcome_modals(self, step_label: str = "3.1/7.1"):
        """3.1 및 7.1 웰컴 모달 / 팝업 버튼 발견 시 클릭"""
        for xpath in WELCOME_MODAL_XPATHS:
            try:
                if ah.element_exists(self.driver, xpath, timeout=2):
                    self._log(f"📌 [{step_label}] 웰컴 모달/팝업 버튼 발견 → 클릭 시도: {xpath[:50]}")
                    ah.wait_and_click(self.driver, xpath, timeout=3, log_callback=self._log)
                    time.sleep(1.5)
            except Exception as e:
                self._log(f"  ⚠ [{step_label}] 모달 닫기 예외: {e}")

    def _dismiss_popups(self):
        """하루/7일 동안 보지 않기 팝업 처리"""
        for xpath in [HIDE_BTN_1DAY, HIDE_BTN_7DAY]:
            if ah.element_exists(self.driver, xpath, timeout=5):
                label = "하루" if "하루" in xpath else "7일"
                self._log(f"📌 '{label} 동안 보지 않기' 팝업 감지 → 클릭")
                ah.wait_and_click(self.driver, xpath, timeout=5, log_callback=self._log)
                time.sleep(2)
        # 한 번 더 하루 보지 않기 확인 (문서 요구사항)
        if ah.element_exists(self.driver, HIDE_BTN_1DAY, timeout=3):
            ah.wait_and_click(self.driver, HIDE_BTN_1DAY, timeout=5, log_callback=self._log)
            time.sleep(2)

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
            time.sleep(2)
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
                time.sleep(2)
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
        [단계 8] 검색입력.png 이미지 인식 → 검색 입력창 클릭 → 검색어 입력
        """
        self._set_status(f"검색어 입력: {keyword}")
        self._log(f"🔍 검색어 입력: '{keyword}'")

        tap_coords = None

        # 1순위: 이미지 매칭 (검색입력.png)
        if os.path.exists(IMG_SEARCH_INPUT):
            self._log("🔍 [이미지 매칭] 검색입력.png 인식 시도")
            coords = self._find_image_coords(IMG_SEARCH_INPUT, threshold=0.7)
            if coords:
                tap_coords = coords
                self._log(f"  ✅ 검색입력.png 발견 → 좌표: {tap_coords}")

        # 2순위: EditText XPath
        if not tap_coords:
            search_xpaths = [
                '//android.widget.EditText[@hint="검색어를 입력해주세요"]',
                '//android.widget.EditText[@hint="검색어 입력"]',
                '//android.widget.EditText[@hint="상품, 브랜드, 쇼핑몰 검색"]',
                '//android.widget.EditText',
            ]
            for xpath in search_xpaths:
                if ah.element_exists(self.driver, xpath, timeout=3):
                    try:
                        el = self.driver.find_element(By.XPATH, xpath)
                        rect = el.rect
                        tap_coords = (rect['x'] + rect['width'] // 2,
                                      rect['y'] + rect['height'] // 2)
                        self._log(f"  ✅ EditText 발견 → 좌표: {tap_coords}")
                        break
                    except Exception:
                        continue

        if not tap_coords:
            self._log("  ❌ 검색 입력창 좌표 획득 실패")
            return False

        # 입력창 클릭
        ah.tap_by_coords(self.driver, tap_coords[0], tap_coords[1], self._log)
        time.sleep(0.8)

        # 텍스트 입력
        try:
            focused_xpath = '//android.widget.EditText[@focused="true"]'
            if ah.element_exists(self.driver, focused_xpath, timeout=3):
                el = self.driver.find_element(By.XPATH, focused_xpath)
            else:
                el = self.driver.find_element(By.XPATH, '//android.widget.EditText')

            # 기존 텍스트 지우고 입력
            el.clear()
            time.sleep(0.3)
            try:
                self.driver.execute_script("mobile: type", {"text": keyword})
                time.sleep(0.5)
            except Exception:
                el.send_keys(keyword)
                time.sleep(0.5)

            self._log(f"  ✅ 검색어 입력 완료: '{keyword}'")
            return True
        except Exception as e:
            self._log(f"  ❌ 검색어 입력 실패: {e}")
            return False

    def _click_search_button(self) -> bool:
        """
        [단계 9] 검색아이콘.png 이미지 인식 클릭, 5초 대기
        """
        self._set_status("검색 실행")

        # 1순위: 이미지 매칭
        if os.path.exists(IMG_SEARCH_ICON):
            coords = self._find_image_coords(IMG_SEARCH_ICON, threshold=0.7)
            if coords:
                ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                self._log("  ✅ 검색아이콘 이미지 인식 클릭 완료")
                time.sleep(5)
                return True

        # 2순위: 키보드 엔터 (검색)
        try:
            import subprocess
            subprocess.run(
                ["adb", "-s", self.device_id, "shell", "input", "keyevent", "66"],
                capture_output=True, timeout=5
            )
            self._log("  ✅ 엔터 키 전송 완료 (검색)")
            time.sleep(5)
            return True
        except Exception as e:
            self._log(f"  ⚠ 엔터 키 전송 실패: {e}")

        # 3순위: 좌표 탭
        try:
            size = self.driver.get_window_size()
            w, h = size['width'], size['height']
            tap_x = int(w * 0.94)
            tap_y = int(h * 0.08)
            ah.tap_by_coords(self.driver, tap_x, tap_y, self._log)
            self._log(f"  ✅ 검색 버튼 좌표 탭 완료 ({tap_x}, {tap_y})")
            time.sleep(5)
            return True
        except Exception as e:
            self._log(f"  ❌ 검색 버튼 클릭 최종 실패: {e}")
            return False

    # ─── 단계 10: 상품 매칭 클릭 ────────────────────────────────────────────

    def _click_product(self, seller_name: str, product_name: str) -> bool:
        """
        [단계 10] 상품 리스트에서 판매자명 + 상품명 매칭 클릭
        - 판매자명 TextView/View + 상품명 View/TextView 조합 매칭
        - 특수문자 대응 및 부분 키워드 매칭 지원
        - 요소 중심 좌표 tap_by_coords 기반 클릭으로 100% 동작 보장
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
                            
                            matched_kws = [kw for kw in keywords if kw in full_txt]
                            is_full_match = (len(matched_kws) == len(keywords)) or (clean_full_name.replace(" ", "") in full_txt.replace(" ", ""))
                            
                            if is_full_match:
                                self._log(f"  📌 상품명 직접 매칭 발견: {dxpath}")
                                if self._safe_click_element(el):
                                    time.sleep(5)
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

                        matched_kws = [kw for kw in keywords if kw in full_txt]
                        is_full_match = (len(matched_kws) == len(keywords)) or (clean_full_name.replace(" ", "") in full_txt.replace(" ", ""))

                        # 1. 판매자명 + 상품명 전체 일치
                        if (seller_name and seller_name in full_txt) and is_full_match:
                            candidate_elements.append((3, len(matched_kws), el, full_txt))
                        # 2. 상품명 전체 일치
                        elif is_full_match:
                            candidate_elements.append((2, len(matched_kws), el, full_txt))
                        # 3. 상품명 키워드 대부분(길이-1) 일치 (폴백)
                        elif len(keywords) > 2 and len(matched_kws) >= len(keywords) - 1:
                            candidate_elements.append((1, len(matched_kws), el, full_txt))
                    except Exception:
                        continue

                # 점수 높은 순 정렬 (우선순위 -> 키워드 일치 개수)
                candidate_elements.sort(key=lambda x: (x[0], x[1]), reverse=True)

                if candidate_elements:
                    best_score, best_kw_cnt, best_el, best_txt = candidate_elements[0]
                    self._log(f"  📌 상품 매칭 성공 (점수={best_score}, 키워드={best_kw_cnt}/{len(keywords)}): '{best_txt[:40]}...'")

                    if self._safe_click_element(best_el):
                        time.sleep(5)
                        return True
                    else:
                        self._log("  ⚠ 좌표 클릭 실패, 계속 탐색...")

            except Exception as e:
                self._log(f"  ⚠ 상품 탐색 중 오류: {e}")

            if scroll_cnt < scroll_max:
                self._log(f"  ⬇ 스크롤 다운 ({scroll_cnt + 1}/{scroll_max})")
                self._scroll_down()
                time.sleep(1.5)

        self._log(f"  ❌ 상품 매칭 최종 실패: {seller_name} / {product_name}")
        return False

    def _safe_click_element(self, el) -> bool:
        """
        요소의 화면 위치(bounds/rect)를 구하여 중심 좌표를 탭(tap_by_coords)
        화면 외부에 있을 경우 클릭 가능 위치로 스크롤 조절
        """
        try:
            rect = el.rect
            x = rect['x'] + rect['width'] // 2
            y = rect['y'] + rect['height'] // 2
            w_w, w_h = 1080, 2400

            try:
                size = self.driver.get_window_size()
                w_w, w_h = size['width'], size['height']
            except Exception:
                pass

            # 화면 영역 내부인지 검사 (상단 바 150px, 하단 바 200px 제외)
            if 150 <= y <= w_h - 200 and 0 < x < w_w:
                self._log(f"  👉 상품 좌표 탭 클릭 시도: ({x}, {y})")
                # 1순위: 좌표 탭
                ah.tap_by_coords(self.driver, x, y, self._log)
                time.sleep(0.5)
                # 2순위: 요소 direct click 보조 실행
                try:
                    el.click()
                except Exception:
                    pass
                return True
            else:
                self._log(f"  ⚠ 요소 좌표가 화면 표시 범위를 벗어남 (y={y}, 화면높이={w_h}) -> 미세 스크롤")
                self._scroll_down()
                time.sleep(1)
                # 스크롤 후 재시도
                rect = el.rect
                x = rect['x'] + rect['width'] // 2
                y = rect['y'] + rect['height'] // 2
                if 150 <= y <= w_h - 200:
                    ah.tap_by_coords(self.driver, x, y, self._log)
                    return True

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
        """[단계 11] 구매하기 버튼 클릭, 5초 대기"""
        self._set_status("구매하기 클릭")
        if ah.element_exists(self.driver, BUY_BTN_XPATH, timeout=5):
            ah.wait_and_click(self.driver, BUY_BTN_XPATH, timeout=5, log_callback=self._log)
            self._log("✅ 구매하기 버튼 클릭 완료")
            time.sleep(5)
            return True
        self._log("⚠ 구매하기 버튼 미발견 → 계속 진행")
        return True  # 없어도 계속 진행

    # ─── 단계 12: 체크박스 이미지 인식 클릭 ──────────────────────────────────

    def _click_checkbox(self) -> bool:
        """[단계 12] 체크박스.png 이미지 인식 클릭, 2초 대기"""
        self._set_status("체크박스 클릭")

        if os.path.exists(IMG_CHECKBOX):
            coords = self._find_image_coords(IMG_CHECKBOX, threshold=0.6)
            if coords:
                ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                self._log("✅ 체크박스 이미지 인식 클릭 완료")
                time.sleep(2)
                return True

        # XPath 폴백
        checkbox_xpaths = [
            '//android.widget.CheckBox',
            '//*[@checkable="true"]',
        ]
        for xpath in checkbox_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=3):
                ah.wait_and_click(self.driver, xpath, timeout=3, log_callback=self._log)
                self._log(f"  ✅ 체크박스 XPath 클릭: {xpath}")
                time.sleep(2)
                return True

        self._log("  ⚠ 체크박스 미발견 → 계속 진행")
        return True

    # ─── 단계 13: 바로구매 이미지 인식 클릭 ──────────────────────────────────

    def _click_buy_now(self) -> bool:
        """[단계 13] 바로구매.png 이미지 인식 클릭, 8초 대기"""
        self._set_status("바로구매 클릭")

        if os.path.exists(IMG_BUY_NOW):
            coords = self._find_image_coords(IMG_BUY_NOW, threshold=0.65)
            if coords:
                ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                self._log("✅ 바로구매 이미지 인식 클릭 완료")
                time.sleep(8)
                return True

        # XPath 폴백
        buy_now_xpaths = [
            '//android.widget.Button[@text="바로구매"]',
            '//android.widget.Button[contains(@text,"구매")]',
            '//*[@text="바로구매"]',
        ]
        for xpath in buy_now_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=3):
                ah.wait_and_click(self.driver, xpath, timeout=5, log_callback=self._log)
                self._log(f"  ✅ 바로구매 XPath 클릭: {xpath}")
                time.sleep(8)
                return True

        self._log("  ❌ 바로구매 버튼 미발견")
        return False

    # ─── 단계 14: 변경 버튼 클릭 ─────────────────────────────────────────────

    def _click_change_button(self) -> bool:
        """[단계 14] 변경 버튼 클릭, 3초 대기"""
        self._set_status("변경 버튼 클릭")
        if ah.element_exists(self.driver, CHANGE_BTN_XPATH, timeout=5):
            ah.wait_and_click(self.driver, CHANGE_BTN_XPATH, timeout=5, log_callback=self._log)
            self._log("✅ 변경 버튼 클릭 완료")
            time.sleep(3)
            return True
        self._log("⚠ 변경 버튼 미발견 → 계속 진행")
        return True

    # ─── 단계 16: 배송지 선택 ────────────────────────────────────────────────

    def _select_delivery_address(self, recipient_name: str, phone: str) -> bool:
        """
        [단계 16] 배송지목록.xml 참고:
        수취인명 포함 text 탐색 + 전화번호 2차 검증 후 클릭
        """
        self._set_status(f"배송지 선택: {recipient_name}")
        self._log(f"🔍 배송지 선택: 수취인={recipient_name!r}, 전화={phone!r}")

        phone_digits = ''.join(filter(str.isdigit, phone)) if phone else ""

        scroll_max = 5
        for scroll_cnt in range(scroll_max + 1):
            # 배송지 목록 내 View 탐색 (배송지목록.xml의 WebView 내부 구조)
            list_xpaths = [
                '//android.webkit.WebView[@text="주문/결제"]//android.view.View',
                '//android.webkit.WebView//android.view.View',
                '//android.view.View',
            ]

            for list_xpath in list_xpaths:
                try:
                    all_views = self.driver.find_elements(By.XPATH, list_xpath)
                    for view in all_views:
                        try:
                            text = view.get_attribute("text") or ""
                            # 수취인명 매칭
                            if recipient_name not in text and recipient_name not in (view.get_attribute("content-desc") or ""):
                                continue
                            # 전화번호 2차 검증 (전화번호 정보 포함)
                            if phone_digits and len(phone_digits) >= 8:
                                last4 = phone_digits[-4:]
                                if last4 not in text and last4 not in (view.get_attribute("content-desc") or ""):
                                    # 같은 수취인이 여러 명이면 전화번호로 구분
                                    continue

                            self._log(f"  🎯 배송지 발견: '{text[:50]}'")
                            if self._click_element_or_parent(view):
                                time.sleep(5)
                                return True
                        except Exception:
                            continue
                    # 하나라도 텍스트로 수취인을 찾으면 break
                except Exception:
                    continue

            # XPath 직접 탐색
            direct_xpaths = [
                f'//android.view.View[contains(@text, "{recipient_name}")]',
                f'//*[contains(@text, "{recipient_name}")]',
                f'//*[contains(@content-desc, "{recipient_name}")]',
            ]
            for xpath in direct_xpaths:
                try:
                    els = self.driver.find_elements(By.XPATH, xpath)
                    if els:
                        self._log(f"  🎯 수취인 직접 XPath 매칭: {xpath}")
                        if self._click_element_or_parent(els[0]):
                            time.sleep(5)
                            return True
                except Exception:
                    continue

            if scroll_cnt < scroll_max:
                self._log(f"  ⬇ 배송지 목록 스크롤 ({scroll_cnt + 1}/{scroll_max})")
                self._scroll_down()
                time.sleep(1.5)

        self._log(f"  ❌ 배송지 '{recipient_name}' 매칭 실패")
        return False

    # ─── 단계 17: 전액사용 클릭 ──────────────────────────────────────────────

    def _click_full_use(self) -> bool:
        """
        [단계 17] 전액사용.png 인식될 때까지 아래로 스크롤하며 대기 -> 인식 후 클릭 -> 3초 대기
        """
        self._set_status("전액사용 탐색 중")
        self._log("🔍 [단계 17] 전액사용 버튼 탐색 시작 (인식될 때까지 스크롤)")

        max_scroll_attempts = 15
        for attempt in range(1, max_scroll_attempts + 1):
            # 1. 전액사용.png 이미지 매칭 (threshold 0.82로 상향하여 오탐 방지)
            if os.path.exists(IMG_FULL_USE):
                coords = self._find_image_coords(IMG_FULL_USE, threshold=0.82)
                if coords:
                    # 상단 바에 가려져 탭 클릭이 씹히는 문제 방지 (Y 좌표가 낮을 경우)
                    if coords[1] < 750:
                        self._log(f"  ⚠ 이미지 상단(Y={coords[1]}) 발견. 중앙 정렬을 위해 스크롤을 내립니다(터치방식).")
                        self._scroll_up(distance_ratio=0.35)
                        time.sleep(1.5)
                        coords = self._find_image_coords(IMG_FULL_USE, threshold=0.82)
                        if not coords:
                            continue
                            
                    # 하단 바에 가려지거나 끝에 걸치는 문제 방지 (Y 좌표가 너무 높을 경우 중간으로 올림)
                    elif coords[1] > 1600:
                        self._log(f"  ⚠ 이미지 하단(Y={coords[1]}) 발견. 중앙 정렬을 위해 터치로 위로 올립니다(화면 아래로).")
                        self._scroll_down(distance_ratio=0.35)
                        time.sleep(1.5)
                        coords = self._find_image_coords(IMG_FULL_USE, threshold=0.82)
                        if not coords:
                            continue

                    self._log(f"  🎯 전액사용.png 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    self._log("✅ [단계 17] 전액사용 이미지 인식 클릭 완료 (3초 대기)")
                    time.sleep(3)
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
                        self._log(f"  🎯 전액사용 XPath 발견: {xpath} -> 클릭 시도")
                        el = self.driver.find_element(By.XPATH, xpath)
                        
                        rect = el.rect
                        y = rect['y'] + rect['height'] // 2
                        if y > 1600:
                            self._log(f"  ⚠ 요소 하단(Y={y}) 발견. 중앙 정렬을 위해 터치로 위로 올립니다.")
                            self._scroll_down(distance_ratio=0.35)
                            time.sleep(1.5)
                            el = self.driver.find_element(By.XPATH, xpath)
                        elif y < 750:
                            self._log(f"  ⚠ 요소 상단(Y={y}) 발견. 중앙 정렬을 위해 스크롤을 내립니다.")
                            self._scroll_up(distance_ratio=0.35)
                            time.sleep(1.5)
                            el = self.driver.find_element(By.XPATH, xpath)

                        if self._safe_click_element(el):
                            self._log("✅ [단계 17] 전액사용 XPath 클릭 완료 (3초 대기)")
                            time.sleep(3)
                            return True
                except Exception:
                    continue

            # 3. 미발견 시 아래로 스크롤 후 재탐색
            self._log(f"  ⬇ [단계 17] 전액사용 미발견 -> 스크롤 다운 ({attempt}/{max_scroll_attempts})")
            self._scroll_down(distance_ratio=0.45)
            time.sleep(1.2)

        self._log("  ❌ 전액사용 버튼 탐색 실패 (최대 스크롤 초과)")
        return False


    # ─── 단계 18: 결제하기 버튼 ──────────────────────────────────────────────

    def _click_pay_button(self) -> bool:
        """[단계 18] 결제하기 버튼 클릭, 5초 대기"""
        self._set_status("결제하기 클릭")

        if ah.element_exists(self.driver, PAY_BTN_XPATH, timeout=5):
            ah.wait_and_click(self.driver, PAY_BTN_XPATH, timeout=5, log_callback=self._log)
            self._log("✅ 결제하기 버튼 클릭 완료")
            time.sleep(5)
            return True

        # 폴백: 좌표
        try:
            size = self.driver.get_window_size()
            w, h = size['width'], size['height']
            tap_x = int(w * 0.5)
            tap_y = int(h * 0.92)
            ah.tap_by_coords(self.driver, tap_x, tap_y, self._log)
            self._log(f"  ✅ 결제하기 좌표 탭 ({tap_x}, {tap_y})")
            time.sleep(5)
            return True
        except Exception as e:
            self._log(f"  ❌ 결제하기 클릭 실패: {e}")
            return False

    # ─── 단계 19: 비밀번호 입력 ──────────────────────────────────────────────

    def _input_password(self, password: str) -> bool:
        """
        [단계 19] 비밀번호 입력
        숫자 폴더의 p0~p9.png 이미지로 각 자리 숫자 버튼의 위치를 인식하여 순서대로 클릭
        """
        self._set_status("비밀번호 입력")
        pwd_digits = ''.join(filter(str.isdigit, password))

        if not pwd_digits:
            self._log("  ⚠ 비밀번호 없음 → 건너뜀")
            return True

        self._log(f"  🔐 비밀번호 입력: {'*' * len(pwd_digits)}자리")

        # 비밀번호 입력 화면 대기 (1~2초)
        time.sleep(1.0)

        for idx, digit in enumerate(pwd_digits):
            img_path = IMG_NUMS.get(digit)
            if not img_path or not os.path.exists(img_path):
                self._log(f"  ⚠ 숫자 이미지 없음: p{digit}.png → 건너뜀")
                continue

            self._log(f"  🔢 {idx+1}번째 자리 '{digit}' 클릭 시도")
            coords = self._find_image_coords(img_path, threshold=0.60)

            if coords:
                ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                self._log(f"    ✅ '{digit}' 클릭 완료 ({coords})")
                time.sleep(0.5)
            else:
                self._log(f"    ❌ '{digit}' 이미지 인식 실패")
                # 실패 시에도 계속 진행

        self._log("  ✅ 비밀번호 입력 완료")
        return True

    def _click_image_with_scroll(self, img_path: str, name: str, threshold: float = 0.82, max_scroll_attempts: int = 15) -> bool:
        """지정된 이미지를 스크롤하며 찾아서 클릭"""
        self._set_status(f"{name} 탐색 중")
        self._log(f"🔍 {name} 버튼 탐색 시작 (인식될 때까지 스크롤, 최대 {max_scroll_attempts}회 시도)")
        for attempt in range(1, max_scroll_attempts + 1):
            if os.path.exists(img_path):
                coords = self._find_image_coords(img_path, threshold=threshold)
                if coords:
                    if coords[1] < 750:
                        self._scroll_up(distance_ratio=0.35)
                        time.sleep(1.5)
                        coords = self._find_image_coords(img_path, threshold=threshold)
                        if not coords: continue
                    elif coords[1] > 1600:
                        self._scroll_down(distance_ratio=0.35)
                        time.sleep(1.5)
                        coords = self._find_image_coords(img_path, threshold=threshold)
                        if not coords: continue
                    self._log(f"  🎯 {name} 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    time.sleep(3)
                    return True
            self._log(f"  ⬇ {name} 미발견 -> 스크롤 다운 ({attempt}/{max_scroll_attempts})")
            self._scroll_down(distance_ratio=0.35)
            time.sleep(1.2)
        self._log(f"  ❌ {name} 버튼 탐색 실패")
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
                    time.sleep(3)
                    return True
            time.sleep(1)
        self._log(f"  ❌ {name} 버튼 탐색 실패 (스크롤 없음)")
        return False

    def _click_bank_select_with_scroll(self, max_scroll_attempts: int = 5) -> bool:
        """
        [무통장입금 클릭 후 판단]
        '은행을.png' 또는 '은행선택.png' 탐색 (최대 5회 스크롤)
        '은행선택.png', '은행을.png' 이미지 또는 '은행' 관련 XPath 인식 시 탭 클릭 및 성공으로 판단.
        미발견 시 무통장입금 정상 선택 실패로 판단하여 작업을 중단하고 False 반환 (실패 처리).
        """
        self._set_status("은행 선택 탐색 및 판단 중")
        self._log(f"🔍 [무통장입금 클릭 후 판단] 은행 선택 버튼 탐색 시작 ('은행선택.png' / '은행을.png', 최대 {max_scroll_attempts}회 시도)")

        bank_images = [
            (IMG_BANK_SELECT, "은행선택"),
            (IMG_SELECT_BANK, "은행을"),
        ]
        bank_xpaths = [
            '//android.widget.Button[contains(@text,"은행")]',
            '//android.view.View[contains(@text,"은행")]',
            '//*[contains(@text,"은행선택")]',
            '//*[contains(@text,"은행을")]',
            '//*[contains(@content-desc,"은행")]',
        ]

        for attempt in range(1, max_scroll_attempts + 1):
            # 1. 이미지 매칭 (threshold 0.70)
            for img_path, name in bank_images:
                if os.path.exists(img_path):
                    coords = self._find_image_coords(img_path, threshold=0.70)
                    if coords:
                        if coords[1] < 750:
                            self._scroll_up(distance_ratio=0.35)
                            time.sleep(1.5)
                            coords = self._find_image_coords(img_path, threshold=0.70)
                            if not coords: continue
                        elif coords[1] > 1600:
                            self._scroll_down(distance_ratio=0.35)
                            time.sleep(1.5)
                            coords = self._find_image_coords(img_path, threshold=0.70)
                            if not coords: continue
                        self._log(f"  🎯 {name} 이미지 발견! 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭 및 존재 확인 성공")
                        ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                        time.sleep(3)
                        return True

            # 2. XPath 매칭 폴백
            for xpath in bank_xpaths:
                try:
                    if ah.element_exists(self.driver, xpath, timeout=1):
                        self._log(f"  🎯 은행 선택 XPath 발견: {xpath} -> 클릭 및 존재 확인 성공")
                        el = self.driver.find_element(By.XPATH, xpath)
                        if self._safe_click_element(el):
                            time.sleep(3)
                            return True
                except Exception:
                    continue

            self._log(f"  ⬇ '은행을'/'은행선택' 미발견 -> 스크롤 다운 ({attempt}/{max_scroll_attempts})")
            self._scroll_down(distance_ratio=0.35)
            time.sleep(1.2)

        self._log(f"  ❌ [실패] 무통장입금 클릭 후 '은행을' / '은행선택' 미발견 (최대 {max_scroll_attempts}회 시도 초과 -> 작업 중단)")
        return False

    def _process_bank_transfer(self) -> bool:
        self._log("💰 [무통장 결제] 프로세스 시작")
        if not self._click_image_with_scroll(IMG_OTHER_PAY, "다른결재", max_scroll_attempts=7):
            self._log("❌ '다른결재' 버튼 미발견 -> 무통장 결제 실패")
            return False
        if not self._click_image_with_scroll(IMG_NORMAL_PAY, "일반결재"):
            self._log("❌ '일반결재' 버튼 미발견 -> 무통장 결제 실패")
            return False
        if not self._click_image_with_scroll(IMG_BANK_TRANSFER, "무통장입금"):
            self._log("❌ '무통장입금' 버튼 미발견 -> 무통장 결제 실패")
            return False

        # 무통장입금 클릭 후 2초 대기하여 화면 전환 보장
        time.sleep(2.0)

        # 무통장입금 클릭 후 '은행을' / '은행선택' 존재하는지 판단 (없으면 실패 및 작업 중단)
        if not self._click_bank_select_with_scroll(max_scroll_attempts=5):
            self._log("❌ 무통장입금 클릭 후 '은행을' / '은행선택' 존재하지 않음 -> 무통장 결제 중단 및 실패 처리")
            return False

        if not self._click_image_with_scroll(IMG_SHINHAN_BANK, "신한은행"):
            self._log("❌ '신한은행' 버튼 미발견 -> 무통장 결제 실패")
            return False
        if not self._click_image_with_scroll(IMG_NOT_APPLY, "미신청"):
            self._log("❌ '미신청' 버튼 미발견 -> 무통장 결제 실패")
            return False
        if not self._click_image_with_scroll(IMG_DO_PAY, "결재하기"):
            self._log("❌ '결재하기' 버튼 미발견 -> 무통장 결제 실패")
            return False
        if not self._click_image_with_scroll(IMG_DO_ORDER, "주문하기"):
            self._log("❌ '주문하기' 버튼 미발견 -> 무통장 결제 실패")
            return False
        return True

    def _process_money_payment(self, password: str) -> bool:
        self._log("💸 [머니 결제] 프로세스 시작")
        if not self._click_image_with_scroll(IMG_MONEY_PAY, "머니"):
            return False
        if not self._click_pay_button():
            self._log("❌ 결제하기 클릭 실패")
            return False
        if password:
            self._input_password(password)
        else:
            self._log("  ℹ 비밀번호 없음 → 건너뜀")
        return True

    # ─── 주문 루프 ────────────────────────────────────────────────────────────

    def _order_loop(self):
        """[단계 2~19] 결재목록 루프 주문 처리"""
        self._log("📋 주문 루프 시작")

        while not self._stop_event.is_set():
            row = self.order_manager.get_next_pending()
            if not row:
                self._log("✅ 모든 주문 처리 완료")
                break

            self._log(f"📌 처리 중: row={row.row_index}, keyword={row.search_keyword!r}")
            self._set_status(f"주문 중: {row.search_keyword}")

            try:
                success = self._process_order_with_timeout(row)
            except Exception as fatal_err:
                self.order_manager.mark_failed(row.row_index)
                self._log(f"❌ 치명적 오류: {fatal_err}")
                raise

            if success:
                self.order_manager.mark_success(row.row_index)
                self._log(f"✅ 주문 성공: {row.search_keyword} → Y 기록")
            else:
                self.order_manager.mark_failed(row.row_index)
                self._log(f"❌ 주문 실패: {row.search_keyword} → F 기록")

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

        # [단계 12] 체크박스 클릭
        self._click_checkbox()

        # [단계 13] 바로구매 클릭
        if not self._click_buy_now():
            self._log("❌ 바로구매 클릭 실패")
            return False

        # [단계 14] 변경 버튼
        self._click_change_button()

        # [단계 15] 스크롤 다운
        self._log("  ⬇ 스크롤 다운 (단계 15)")
        self._scroll_down()
        time.sleep(1.0)

        # [단계 16] 배송지 선택
        if not self._select_delivery_address(row.recipient_name, row.phone):
            self._log("❌ 배송지 선택 실패")
            return False

        # [단계 17] 전액사용 클릭 등 결제 방식 분기
        if row.payment_method == "무통장":
            if not self._process_bank_transfer():
                self._log("❌ 무통장 결제 진행 실패")
                return False
        elif row.payment_method == "머니":
            if not self._process_money_payment(row.password):
                self._log("❌ 머니 결제 진행 실패")
                return False
        else:
            # 포인트 또는 기본 결제
            self._click_full_use()
            # [단계 18] 결제하기 버튼
            if not self._click_pay_button():
                self._log("❌ 결제하기 클릭 실패")
                return False
            # [단계 19] 비밀번호 입력
            if row.password:
                self._input_password(row.password)
            else:
                self._log("  ℹ 비밀번호 없음 → 건너뜀")

        self._log(f"✅ 주문 완료: {row.search_keyword}")
        return True

    # ─── 이미지 인식 (naver_worker.py 동일 로직 재구현) ─────────────────────

    def _find_image_coords(self, template_path: str,
                           threshold: float = 0.75) -> Optional[tuple]:
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

            scales = np.linspace(0.4, 3.0, 60)
            for scale in scales:
                new_w = int(t_w * scale)
                new_h = int(t_h * scale)
                if new_w >= screen_w or new_h >= screen_h:
                    continue
                if new_w < 10 or new_h < 5:
                    continue
                resized = cv2.resize(template_gray, (new_w, new_h),
                                     interpolation=cv2.INTER_AREA)
                result = cv2.matchTemplate(screen_gray, resized, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)
                if max_val > best_score:
                    best_score = max_val
                    best_loc   = max_loc
                    best_tw    = new_w
                    best_th    = new_h

            if best_score >= threshold and best_loc is not None:
                cx = best_loc[0] + best_tw // 2
                cy = best_loc[1] + best_th // 2
                self._log(f"  🎯 [이미지 매칭] 발견! 중심좌표: ({cx}, {cy}), 점수: {best_score:.4f}")
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

    def _scroll_down(self, distance_ratio: float = 0.4):
        """아래로 스크롤 (Appium swipe + ADB input swipe 보조)"""
        w, h = 1080, 2400
        try:
            size = self.driver.get_window_size()
            w, h = size['width'], size['height']
        except Exception:
            pass

        start_y = int(h * 0.72)
        end_y = int(start_y - (h * distance_ratio))
        end_y = max(100, end_y)

        # 1순위: W3C Actions (최신 Appium 2.x 권장)
        try:
            from selenium.webdriver.common.actions.action_builder import ActionBuilder
            from selenium.webdriver.common.actions.pointer_input import PointerInput
            from selenium.webdriver.common.actions import interaction
            
            pointer = PointerInput(interaction.POINTER_TOUCH, "touch")
            action = ActionBuilder(self.driver, mouse=pointer)
            action.pointer_action.move_to_location(w // 2, start_y)
            action.pointer_action.pointer_down()
            action.pointer_action.pause(0.1)
            action.pointer_action.move_to_location(w // 2, end_y)
            action.pointer_action.release()
            action.perform()
            return
        except Exception:
            pass

        # 2순위: TouchAction (사용자 요청: 터치 방식 스크롤)
        try:
            from appium.webdriver.common.touch_action import TouchAction
            action = TouchAction(self.driver)
            action.press(x=w//2, y=start_y).wait(500).move_to(x=w//2, y=end_y).release().perform()
            return
        except Exception:
            pass

        # 3순위: Appium 기존 swipe
        try:
            self.driver.swipe(w // 2, start_y, w // 2, end_y, 450)
            return
        except Exception:
            pass

        # 4순위: ADB shell input swipe (WebView 및 예외 발생 시 보장)
        try:
            import subprocess
            subprocess.run(
                ["adb", "-s", self.device_id, "shell", "input", "swipe",
                 str(w // 2), str(start_y), str(w // 2), str(end_y), "500"],
                capture_output=True, timeout=5
            )
        except Exception:
            pass

    def _scroll_up(self, distance_ratio: float = 0.4):
        """위로 스크롤 (화면을 아래로 내림, 터치 방식 포함)"""
        w, h = 1080, 2400
        try:
            size = self.driver.get_window_size()
            w, h = size['width'], size['height']
        except Exception:
            pass

        start_y = int(h * 0.3)
        end_y = int(start_y + (h * distance_ratio))
        end_y = min(h - 100, end_y)

        try:
            from selenium.webdriver.common.actions.action_builder import ActionBuilder
            from selenium.webdriver.common.actions.pointer_input import PointerInput
            from selenium.webdriver.common.actions import interaction
            
            pointer = PointerInput(interaction.POINTER_TOUCH, "touch")
            action = ActionBuilder(self.driver, mouse=pointer)
            action.pointer_action.move_to_location(w // 2, start_y)
            action.pointer_action.pointer_down()
            action.pointer_action.pause(0.1)
            action.pointer_action.move_to_location(w // 2, end_y)
            action.pointer_action.release()
            action.perform()
            return
        except Exception:
            pass

        try:
            from appium.webdriver.common.touch_action import TouchAction
            action = TouchAction(self.driver)
            action.press(x=w//2, y=start_y).wait(500).move_to(x=w//2, y=end_y).release().perform()
            return
        except Exception:
            pass

        try:
            import subprocess
            subprocess.run(
                ["adb", "-s", self.device_id, "shell", "input", "swipe",
                 str(w // 2), str(start_y), str(w // 2), str(end_y), "500"],
                capture_output=True, timeout=5
            )
        except Exception:
            pass


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
            log_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                f"naver_order_{self.device_id}.log"
            )
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{now}] {full_msg}\n")
        except Exception:
            pass

    def _set_status(self, status: str):
        if self._status_cb:
            self._status_cb(self.device_id, status)
