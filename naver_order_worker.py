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
IMG_SEARCH_ICON   = os.path.join(_IMG_DIR, "검색아이콘.png") # 검색 아이콘 (단계 9)
IMG_CHECKBOX      = os.path.join(_IMG_DIR, "체크박스.png")   # 체크박스 (단계 12)
IMG_OPTION_SELECT = os.path.join(_IMG_DIR, "옵션 선택.png")  # 옵션 선택 텍스트 (체크박스 위)
IMG_DELIVERY_INFO = os.path.join(_IMG_DIR, "배송정보.png")  # 배송정보 텍스트 (체크박스 아래)
IMG_BUY_NOW       = os.path.join(_IMG_DIR, "바로구매.png")   # 바로구매 버튼 (단계 13)
IMG_ORDER_PAY     = os.path.join(_IMG_DIR, "주문결재.png")   # 주문결재 확인용 (단계 13 폴백)
IMG_FULL_USE      = os.path.join(_IMG_DIR, "전액사용.png")   # 전액사용 버튼 (단계 17)

# 추가된 결제방식 이미지
IMG_OTHER_PAY     = os.path.join(_IMG_DIR, "다른결재.png")
IMG_NORMAL_PAY    = os.path.join(_IMG_DIR, "일반결재.png")
IMG_NORMAL_PAY_CHECK = os.path.join(_IMG_DIR, "일반결재체크.png")
IMG_BANK_TRANSFER = os.path.join(_IMG_DIR, "무통장입금.png")
IMG_BANK_TRANSFER_CHECK = os.path.join(_IMG_DIR, "무통장체크.png")
IMG_SELECT_BANK   = os.path.join(_IMG_DIR, "은행을.png")
IMG_BANK_SELECT   = os.path.join(_IMG_DIR, "은행선택.png")
IMG_SHINHAN_BANK  = os.path.join(_IMG_DIR, "신한은행.png")
IMG_HANA_BANK     = os.path.join(_IMG_DIR, "하나은행.png")
IMG_NONGHYUP_BANK = os.path.join(_IMG_DIR, "농협.png")
IMG_WOORI_BANK    = os.path.join(_IMG_DIR, "우리은행.png")
IMG_KB_BANK       = os.path.join(_IMG_DIR, "국민은행.png")
IMG_IBK_BANK      = os.path.join(_IMG_DIR, "기업은행.png")
IMG_NOT_APPLY     = os.path.join(_IMG_DIR, "미신청.png")
IMG_NOT_APPLY2    = os.path.join(_IMG_DIR, "미신청2.png")
IMG_DO_PAY        = os.path.join(_IMG_DIR, "결재하기.png")
IMG_DO_ORDER      = os.path.join(_IMG_DIR, "주문하기.png")
IMG_MONEY_PAY     = os.path.join(_IMG_DIR, "머니.png")
IMG_PAY_BENEFIT   = os.path.join(_IMG_DIR, "결제혜택.png")  # 결제혜택 팝업 감지용
IMG_CLOSE_POPUP   = os.path.join(_IMG_DIR, "닫기.png")      # 팝업 닫기 버튼

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

        # [단계 6] 마이쇼핑 클릭 -> 10초 대기
        self._set_status("마이쇼핑 클릭")
        if ah.element_exists(self.driver, MY_SHOPPING_XPATH, timeout=8):
            ah.wait_and_click(self.driver, MY_SHOPPING_XPATH, timeout=7, log_callback=self._log)
            self._log("✅ 마이쇼핑 클릭 완료 (10초 대기)")
            time.sleep(7)
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
                time.sleep(3)
                return True

        # 2순위: 키보드 엔터 (검색)
        try:
            import subprocess
            subprocess.run(
                ["adb", "-s", self.device_id, "shell", "input", "keyevent", "66"],
                capture_output=True, timeout=5
            )
            self._log("  ✅ 엔터 키 전송 완료 (검색)")
            time.sleep(3)
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
                self._scroll_down()
                time.sleep(1.0)

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

            # 화면 영역 내부인지 검사
            if 150 <= y <= w_h - 200 and 0 < x < w_w:
                self._log(f"  👉 ADB 좌표 탭: ({x}, {y})")
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
                if 150 <= y <= w_h - 200:
                    subprocess.run(
                        ["adb", "-s", self.device_id, "shell", "input", "tap",
                         str(x), str(y)],
                        capture_output=True, timeout=5
                    )
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
            time.sleep(3)
            return True
        self._log("⚠ 구매하기 버튼 미발견 → 계속 진행")
        return True  # 없어도 계속 진행

    # ─── 단계 12: 체크박스 이미지 인식 클릭 ──────────────────────────────────

    def _click_checkbox(self, product_name: str = "") -> bool:
        """[단계 12] 체크박스.png 이미지 인식 및 옵션 항목 클릭, 2초 대기 (threshold 0.65, min_y 0.40*h, max_y 0.90*h)"""
        self._set_status("체크박스/옵션 선택")
        self._log("🔍 체크박스 및 옵션 항목 탐색 시도 중...")

        w_h = 2400
        w_w = 1080
        try:
            size = self.driver.get_window_size()
            w_h = size['height']
            w_w = size['width']
        except Exception:
            pass

        min_y_check = int(w_h * 0.40)  # 화면 하단 40% 이하 영역
        max_y_check = int(w_h * 0.90)  # y=2160 이하, 최하단 네비게이션 바(y>0.93) 제외
        max_x_check = int(w_w * 0.40)  # 체크박스는 화면 좌측에 위치하므로 X축 제한

        # 0순위: 옵션 선택(위) ~ 배송정보(아래) 사이로 Y축 제한 탐색 (사용자 요청 사항)
        if os.path.exists(IMG_OPTION_SELECT) and os.path.exists(IMG_DELIVERY_INFO):
            opt_coords = self._find_image_coords(IMG_OPTION_SELECT, threshold=0.70)
            del_coords = self._find_image_coords(IMG_DELIVERY_INFO, threshold=0.70)
            
            if opt_coords and del_coords and opt_coords[1] < del_coords[1]:
                min_y_check = opt_coords[1]
                max_y_check = del_coords[1]
                self._log(f"  📌 체크박스 탐색 Y영역 동적 설정 (옵션~배송정보): {min_y_check} ~ {max_y_check}")

        # 1순위: 템플릿 이미지 매칭 (threshold 0.65 로 설정하여 0.7518 점수 수용)
        if os.path.exists(IMG_CHECKBOX):
            coords = self._find_image_coords(IMG_CHECKBOX, threshold=0.65, min_x=0, max_x=max_x_check, min_y=min_y_check, max_y=max_y_check)
            if coords:
                ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                self._log("✅ 체크박스 이미지 인식 클릭 완료")
                time.sleep(1.5)
                return True

        # 2순위: 상품명 키워드 기반 옵션 텍스트 XPath 탐색 (예: "유기농 올리브", "2박스", "레드비트샷" 등)
        if product_name:
            import re
            clean_prod = re.sub(r'[\+\-\*\/\(\)\[\]\{\}\?\!\,]', ' ', product_name)
            keywords = [k.strip() for k in clean_prod.split() if len(k.strip()) >= 2]
            for kw in keywords[:3]:
                safe_kw = kw.replace('"', '').replace("'", "")
                option_xpaths = [
                    f'//android.view.View[contains(@text, "{safe_kw}")]',
                    f'//android.widget.TextView[contains(@text, "{safe_kw}")]',
                    f'//*[contains(@text, "{safe_kw}")]',
                ]
                for xpath in option_xpaths:
                    if ah.element_exists(self.driver, xpath, timeout=1):
                        try:
                            els = self.driver.find_elements(By.XPATH, xpath)
                            for el in els:
                                rect = el.rect
                                cx = rect['x'] + rect['width'] // 2
                                cy = rect['y'] + rect['height'] // 2
                                if min_y_check <= cy <= max_y_check and cx <= max_x_check:
                                    self._safe_click_element(el)
                                    self._log(f"  ✅ 옵션 상품 텍스트 XPath 클릭 ({kw}, x={cx}, y={cy}): {xpath}")
                                    time.sleep(1.5)
                                    return True
                        except Exception:
                            pass

        # 3순위: 기본 CheckBox / checkable / 옵션 키워드 XPath 폴백
        checkbox_xpaths = [
            '//android.widget.CheckBox',
            '//*[@checkable="true"]',
            '//*[contains(@text, "박스")]',
            '//*[contains(@text, "포")]',
            '//*[contains(@text, "개")]',
            '//*[contains(@text, "동의")]',
            '//*[contains(@text, "구매조건")]',
        ]
        for xpath in checkbox_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=1):
                try:
                    els = self.driver.find_elements(By.XPATH, xpath)
                    for el in els:
                        rect = el.rect
                        cx = rect['x'] + rect['width'] // 2
                        cy = rect['y'] + rect['height'] // 2
                        if min_y_check <= cy <= max_y_check and cx <= max_x_check:
                            self._safe_click_element(el)
                            self._log(f"  ✅ 체크박스 XPath 클릭 (x={cx}, y={cy}): {xpath}")
                            time.sleep(1.5)
                            return True
                except Exception:
                    pass

        self._log("  ⚠ 체크박스 미발견 → 계속 진행")
        return True

    # ─── 단계 13: 바로구매 이미지 인식 클릭 ──────────────────────────────────

    def _click_buy_now(self) -> bool:
        """[단계 13] 바로구매.png 이미지 인식 클릭, 8초 대기 (threshold 0.70, min_y 0.50*h, max_y 0.93*h)"""
        self._set_status("바로구매 클릭")

        w_h = 2400
        try:
            w_h = self.driver.get_window_size()['height']
        except Exception:
            pass

        min_y_buynow = int(w_h * 0.50)  # 바로구매 버튼은 화면 하단 50% 이하 영역만 유효
        max_y_buynow = int(w_h * 0.93)  # 시스템 바 제외

        for attempt in range(1, 3):
            if os.path.exists(IMG_BUY_NOW):
                coords = self._find_image_coords(IMG_BUY_NOW, threshold=0.65, min_y=min_y_buynow, max_y=max_y_buynow)
                if coords:
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    self._log("✅ 바로구매 이미지 인식 클릭 완료")
                    time.sleep(5)
                    return True

            # XPath 폴백 ("구매조건" 동의 문구 오매칭 방지를 위해 정확한 바로구매 문구만 탐색)
            buy_now_xpaths = [
                '//android.widget.Button[@text="바로구매"]',
                '//android.widget.Button[contains(@text,"바로구매")]',
                '//android.view.View[contains(@text,"바로구매")]',
                '//*[@content-desc="바로구매"]',
                '//*[@text="바로구매"]',
            ]
            for xpath in buy_now_xpaths:
                if ah.element_exists(self.driver, xpath, timeout=2):
                    try:
                        els = self.driver.find_elements(By.XPATH, xpath)
                        for el in els:
                            rect = el.rect
                            cy = rect['y'] + rect['height'] // 2
                            if min_y_buynow <= cy <= max_y_buynow:
                                self._safe_click_element(el)
                                self._log(f"  ✅ 바로구매 XPath 클릭 (y={cy}): {xpath}")
                                time.sleep(5)
                                return True
                    except Exception:
                        pass

            # 바로구매 버튼을 찾지 못한 경우 (하단 바가 닫혔을 가능성 대응)
            if attempt == 1:
                self._log("  ⚠ 바로구매 버튼 미발견 (1회차) → '구매하기' 버튼 재클릭으로 옵션 드로어 복구 시도")
                if ah.element_exists(self.driver, BUY_BTN_XPATH, timeout=3):
                    ah.wait_and_click(self.driver, BUY_BTN_XPATH, timeout=3, log_callback=self._log)
                    self._log("  ✅ '구매하기' 버튼 재클릭 완료 → 3초 대기 후 바로구매 재시도")
                    time.sleep(3)
                else:
                    self._log("  ⚠ '구매하기' 버튼도 화면에 없음")
                    break

        self._log("  ❌ 바로구매 버튼 미발견")

        # 미발견 시 주문결재.png 가 존재하는지 확인 (이미 다음 페이지로 넘어갔을 경우 대비)
        if os.path.exists(IMG_ORDER_PAY):
            self._log("  🔍 주문결재.png 존재 여부 확인 중...")
            if self._find_image_coords(IMG_ORDER_PAY, threshold=0.65):
                self._log("  ✅ 주문결재.png 확인됨! 바로구매 버튼 클릭 성공으로 간주하고 계속 진행")
                return True
                
        return False

    # ─── 단계 14: 변경 버튼 클릭 ─────────────────────────────────────────────

    def _click_change_button(self) -> bool:
        """[단계 14] 변경 버튼 클릭, 3초 대기"""
        self._set_status("변경 버튼 클릭")
        if ah.element_exists(self.driver, CHANGE_BTN_XPATH, timeout=5):
            ah.wait_and_click(self.driver, CHANGE_BTN_XPATH, timeout=5, log_callback=self._log)
            self._log("✅ 변경 버튼 클릭 완료")
            time.sleep(2)
            return True
        self._log("⚠ 변경 버튼 미발견 → 계속 진행")
        return True

    # ─── 단계 16: 배송지 선택 ────────────────────────────────────────────────

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
                                        if recipient_name in area_text:
                                            return True
                                    except Exception:
                                        break
                    except Exception:
                        continue
        except Exception:
            pass
            
        # 기존 폴백 로직 (배송지명으로 탐색)
        try:
            views = self.driver.find_elements(By.XPATH, '//*[contains(@text, "배송지명")]')
            for view in views:
                text = view.get_attribute("text") or ""
                if recipient_name in text:
                    if not last4:
                        return True
                    try:
                        parent = view.find_element(By.XPATH, "..")
                        area_views = parent.find_elements(By.XPATH, ".//*")
                        area_text = " ".join([av.get_attribute("text") or "" for av in area_views])
                        if last4 in area_text:
                            return True
                    except Exception:
                        pass
        except Exception:
            pass
            
        return False

    def _find_recipient_on_screen(self, recipient_name: str, phone_digits: str) -> bool:
        """
        현재 화면(스크롤 없이)에서 수취인을 탐색하여 클릭. 찾으면 True 반환.
        """
        last4 = phone_digits[-4:] if (phone_digits and len(phone_digits) >= 4) else ""
        formatted_phone = f"{phone_digits[:3]}-{phone_digits[3:7]}-{phone_digits[7:]}" if len(phone_digits) >= 8 else ""

        # ── 방법 1: 배송지 목록 팝업 ('선택' 또는 '선택됨' 버튼 우선 탐색) ──
        # 팝업 리스트에서는 "선택", "선택됨" 텍스트를 가진 버튼을 찾는 것이 좌표 오차를 없애는 가장 확실한 방법입니다.
        try:
            sel_views = self.driver.find_elements(By.XPATH, '//*[contains(@text, "선택")]')
            for sv in sel_views:
                try:
                    sv_text = (sv.get_attribute("text") or "").strip()
                    if sv_text not in ["선택", "선택됨"]:
                        continue

                    node = sv
                    found_name = False
                    # 조상 노드를 5단계까지 올라가며 동일 블록 안에 수취인명과 전화번호가 있는지 확인
                    for _ in range(5):
                        try:
                            node = node.find_element(By.XPATH, "..")
                            els = node.find_elements(By.XPATH, ".//*")
                            area_text = " ".join((e.get_attribute("text") or "") for e in els)

                            if recipient_name in area_text:
                                # 전화번호 검증 (팝업은 마스킹 안됨, 하지만 혹시 모르니 확인)
                                if (formatted_phone in area_text) or (last4 in area_text) or ("***" in area_text):
                                    found_name = True
                                    break
                                elif not last4:
                                    found_name = True
                                    break
                        except Exception:
                            break
                    
                    if found_name:
                        bounds_str = sv.get_attribute("bounds") or ""
                        self._log(f"  🎯 팝업 매칭! '{recipient_name}' 그룹의 '{sv_text}' 버튼 클릭 시도")
                        
                        import re as _re, subprocess
                        m = _re.match(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', bounds_str)
                        if m:
                            x1, y1, x2, y2 = map(int, m.groups())
                            cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
                            self._log(f"  👉 '{sv_text}' 직접 ADB 탭: ({cx}, {cy})")
                            subprocess.run(
                                ["adb", "-s", self.device_id, "shell", "input", "tap", str(cx), str(cy)],
                                capture_output=True, timeout=5
                            )
                        else:
                            self._log("  ⚠ 좌표 파싱 실패 -> 기본 클릭")
                            self._safe_click_element(sv)
                        
                        time.sleep(3)
                        return True
                except Exception:
                    continue
        except Exception:
            pass

        # ── 방법 2: 수취인 이름 기반 탐색 (팝업 폴백, '배송지명' 인라인 텍스트 제외) ──
        # 팝업에서 '선택' 버튼을 못 찾았을 경우, 이름 텍스트 뷰 자체를 클릭합니다.
        try:
            search_xpath = f'//*[contains(@text, "{recipient_name}")]'
            views = self.driver.find_elements(By.XPATH, search_xpath)
            for view in views:
                try:
                    text = view.get_attribute("text") or ""
                    # 사용자 요청: '배송지명'이 포함된 인라인 배경 텍스트는 오작동 유발하므로 무조건 제외
                    if "배송지명" in text:
                        continue
                    
                    if recipient_name not in text:
                        continue

                    self._log(f"  🔎 이름 패턴 View 발견: '{text[:60]}'")

                    if last4:
                        phone_ok = False
                        try:
                            parent = view.find_element(By.XPATH, "..")
                            sibling_texts = []
                            try:
                                grand_parent = parent.find_element(By.XPATH, "..")
                                area_views = grand_parent.find_elements(By.XPATH, ".//*")
                                for av in area_views:
                                    try:
                                        at = av.get_attribute("text") or ""
                                        if at: sibling_texts.append(at)
                                    except Exception:
                                        pass
                            except Exception:
                                pass

                            area_text = " ".join(sibling_texts)
                            if "***" in area_text:
                                phone_ok = True
                                self._log("  ℹ 전화번호 마스킹 감지 → 이름만으로 매칭")
                            elif last4 in area_text:
                                phone_ok = True
                            elif "연락처" in area_text:
                                for st in sibling_texts:
                                    if "연락처" in st and last4 in st:
                                        phone_ok = True
                                        break
                        except Exception:
                            phone_ok = True
                            self._log("  ⚠ 전화번호 범위 탐색 실패 → 이름만으로 매칭 시도")

                        if not phone_ok:
                            self._log(f"  ⚠ 전화번호 뒷4자리 '{last4}' 미매칭 → 스킵")
                            continue

                    self._log(f"  🎯 이름 패턴 매칭 성공: '{text[:60]}'")
                    
                    # ── 사용자 요청: 연락처(전화번호) View를 찾아 클릭 ──
                    target_view = view
                    if last4:
                        try:
                            parent = view.find_element(By.XPATH, "..")
                            grand_parent = parent.find_element(By.XPATH, "..")
                            search_xpath = f'.//*[contains(@text, "{formatted_phone}") or contains(@text, "{last4}")]'
                            phone_views = grand_parent.find_elements(By.XPATH, search_xpath)
                            for pv in phone_views:
                                ptext = pv.get_attribute("text") or ""
                                if "연락처" in ptext or formatted_phone in ptext or last4 in ptext:
                                    target_view = pv
                                    self._log(f"  👉 연락처 View로 클릭 대상 변경: {ptext[:30]}")
                                    break
                        except Exception:
                            pass

                    import re as _re, subprocess
                    try:
                        bounds_str = target_view.get_attribute("bounds") or ""
                        m = _re.match(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', bounds_str)
                        if m:
                            x1, y1, x2, y2 = map(int, m.groups())
                            tap_x = (x1 + x2) // 2
                            tap_y = (y1 + y2) // 2
                        else:
                            rect = target_view.rect
                            tap_x = rect['x'] + rect['width'] // 2
                            tap_y = rect['y'] + rect['height'] // 2
                        self._log(f"  👉 직접 ADB 탭: ({tap_x}, {tap_y})")
                        _run_cmd(
                            ["adb", "-s", self.device_id, "shell", "input", "tap",
                             str(tap_x), str(tap_y)],
                            capture_output=True, timeout=5
                        )
                        time.sleep(3)
                        return True
                    except Exception as e:
                        self._log(f"  ⚠ 좌표 클릭 실패: {e}")
                        if self._safe_click_element(target_view):
                            time.sleep(3)
                            return True
                except Exception:
                    continue
        except Exception:
            pass

        # ── 방법 B: 배송지 목록 팝업 형태 (명확한 '선택' 버튼 좌표 클릭) ──
        # 전략: 전화번호 View를 찾은 후, 동일 블록 내 수취인명과 "선택"(또는 "선택됨") 버튼을 확인하고,
        # 버튼의 중앙 좌표(bounds 파싱)를 직접 계산하여 터치합니다. (엉뚱한 신규배송지 클릭 방지)
        try:
            if phone_digits and len(phone_digits) >= 8:
                formatted_phone = f"{phone_digits[:3]}-{phone_digits[3:7]}-{phone_digits[7:]}"
                phone_xpaths = [
                    f'//*[contains(@text, "{formatted_phone}")]',
                    f'//*[contains(@text, "{last4}")]'
                ]
                for px in phone_xpaths:
                    try:
                        p_views = self.driver.find_elements(By.XPATH, px)
                        for pv in p_views:
                            try:
                                pv_text = pv.get_attribute("text") or ""
                                if "연락처" in pv_text:
                                    continue
                                
                                node = pv
                                target_button = None
                                found_name = False
                                
                                # 조상 노드를 5단계까지 올라가며 수취인명과 선택버튼 탐색
                                for _ in range(5):
                                    try:
                                        node = node.find_element(By.XPATH, "..")
                                        els = node.find_elements(By.XPATH, ".//*")
                                        
                                        # 이름 검증
                                        if not found_name:
                                            area_text = " ".join((e.get_attribute("text") or "") for e in els)
                                            if recipient_name in area_text:
                                                found_name = True
                                        
                                        # 버튼 탐색
                                        if not target_button:
                                            for e in els:
                                                e_class = e.get_attribute("class") or ""
                                                e_text = e.get_attribute("text") or ""
                                                if "Button" in e_class and e_text in ["선택", "선택됨"]:
                                                    target_button = e
                                                    break
                                                    
                                        if found_name and target_button:
                                            break
                                    except Exception:
                                        break
                                
                                if found_name and target_button:
                                    btn_text = target_button.get_attribute("text")
                                    bounds_str = target_button.get_attribute("bounds")
                                    self._log(f"  🎯 팝업 매칭! '{pv_text}' 그룹의 '{btn_text}' 버튼 클릭 시도 (bounds: {bounds_str})")
                                    
                                    # 명시적으로 bounds 중앙을 파싱해 클릭
                                    import re
                                    match = re.match(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', bounds_str)
                                    if match:
                                        x1, y1, x2, y2 = map(int, match.groups())
                                        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
                                        self._log(f"  👉 '{btn_text}' 버튼 정중앙 탭: X={cx}, Y={cy}")
                                        self.driver.tap([(cx, cy)])
                                    else:
                                        self._log("  ⚠ 좌표 파싱 실패 -> 기본 요소 클릭 사용")
                                        self._click_element_or_parent(target_button)
                                        
                                    time.sleep(3)
                                    return True
                            except Exception:
                                continue
                    except Exception:
                        pass
        except Exception:
            pass

        # ── 방법 B-2: 이름으로 먼저 찾는 폴백 로직 (전화번호가 마스킹된 경우 대비) ──
        try:
            name_xpaths = [
                f'//*[contains(@text, "{recipient_name}")]',
                f'//*[contains(@content-desc, "{recipient_name}")]',
            ]
            for nxp in name_xpaths:
                try:
                    name_views = self.driver.find_elements(By.XPATH, nxp)
                    for nv in name_views:
                        try:
                            nv_text = nv.get_attribute("text") or ""
                            nv_desc = nv.get_attribute("content-desc") or ""

                            # '배송지명' prefix가 붙은 경우는 방법 A에서 이미 처리했으므로 패스
                            if "배송지명" in nv_text:
                                continue
                            # 수취인명 매칭
                            if recipient_name not in (nv_text + nv_desc):
                                continue

                            # 전화번호 검증 (같은 View에 있는지 먼저 확인)
                            if last4:
                                combined = nv_text + " " + nv_desc
                                if last4 in combined:
                                    # 같은 View에 전화번호도 있으면 바로 클릭
                                    self._log(f"  🎯 배송지 발견 (이름+전화 동일 View): '{nv_text[:50]}'")
                                    if self._click_element_or_parent(nv):
                                        time.sleep(3)
                                        return True
                                else:
                                    # 같은 View에 없으면 부모 영역 탐색
                                    area_text = ""
                                    try:
                                        parent = nv.find_element(By.XPATH, "..")
                                        gparent = parent.find_element(By.XPATH, "..")
                                        area_els = gparent.find_elements(By.XPATH, ".//*")
                                        area_text = " ".join(
                                            (ae.get_attribute("text") or "") for ae in area_els
                                        )
                                    except Exception:
                                        pass
                                    # 마스킹 감지: '***' 포함 시 전화번호 검증 스킵
                                    if "***" in area_text:
                                        self._log(f"  ℹ 전화번호 마스킹 감지 → 이름만으로 클릭: '{nv_text[:40]}'")
                                    elif area_text and last4 not in area_text:
                                        self._log(f"  ⚠ 전화 뒷4자리 '{last4}' 부모 영역에도 없음 → 스킵")
                                        continue

                            self._log(f"  🎯 배송지 발견: '{nv_text[:50]}'")
                            if self._click_element_or_parent(nv):
                                time.sleep(3)
                                return True
                        except Exception:
                            continue
                except Exception:
                    continue
        except Exception:
            pass

        # ── 방법 C: RadioButton[text="선택"] 을 이름+전화가 매칭된 블록에서 직접 클릭 ──
        # XML: <android.widget.RadioButton text="선택" resource-id="delivery_option_..."/>
        try:
            radio_buttons = self.driver.find_elements(
                By.XPATH, '//android.widget.RadioButton[@text="선택"]'
            )
            for rb in radio_buttons:
                try:
                    # RadioButton의 부모 컨테이너에서 이름 탐색
                    container = rb.find_element(By.XPATH, "..")
                    container_els = container.find_elements(By.XPATH, ".//*")
                    container_text = " ".join(
                        (ce.get_attribute("text") or "") for ce in container_els
                    )
                    if recipient_name not in container_text:
                        continue
                    # 마스킹 감지: '***' 포함 시 전화번호 검증 스킵
                    if last4 and "***" not in container_text and last4 not in container_text:
                        self._log(f"  ⚠ RadioButton 블록 전화 '{last4}' 미매칭 → 스킵")
                        continue
                    self._log(f"  🎯 RadioButton '선택' 클릭 - 수취인 '{recipient_name}' 매칭")
                    rb.click()
                    time.sleep(3)
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

        # ── 1단계: 스크롤 없이 현재 화면에서 먼저 탐색 ──
        self._log("  📋 현재 화면에서 수취인 탐색 중 (스크롤 없음)...")
        if self._find_recipient_on_screen(recipient_name, phone_digits):
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
            self._scroll_down()
            time.sleep(1.0)

            self._log(f"  📋 스크롤 후 수취인 재탐색...")
            if self._find_recipient_on_screen(recipient_name, phone_digits):
                return True

        self._log(f"  ❌ 배송지 '{recipient_name}' 매칭 실패")
        return False

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
                           screenshot_png: Optional[bytes] = None) -> Optional[tuple]:
        """
        숫자 키패드 이미지 인식 (경량화 버전).
        - 스케일 20단계 (0.6~1.8), UI 과부하 방지
        - screenshot_png를 외부에서 주입하면 재캡처 생략 (재시도 성능 개선)
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

            template_bgr = cv2.imdecode(
                np.fromfile(img_path, dtype=np.uint8), cv2.IMREAD_COLOR
            )
            if template_bgr is None:
                return None
            template_gray = cv2.cvtColor(template_bgr, cv2.COLOR_BGR2GRAY)
            t_h, t_w = template_gray.shape

            best_score, best_loc, best_tw, best_th = -1, None, t_w, t_h
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

            if best_score >= 0.50 and best_loc is not None:
                cx = best_loc[0] + best_tw // 2
                cy = best_loc[1] + best_th // 2
                self._log(f"    🎯 [숫자인식] 점수: {best_score:.4f} → 좌표 ({cx}, {cy})")
                return cx, cy

            self._log(f"    ↩ [숫자인식] 미발견 (점수 {best_score:.4f} < 0.50)")
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

        save_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "캡쳐", "무통장")
        os.makedirs(save_dir, exist_ok=True)

        timestamp = time.strftime("%Y%m%d_%H%M%S")
        safe_keyword = re.sub(r'[\\/*?"<>|]', "", row.search_keyword)
        screenshot_filename = f"무통장_{timestamp}_{safe_keyword}.png"
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
        log_content = f"[{timestamp}] 키워드: {row.search_keyword} | 주문번호: {order_display} | 계좌정보: {bank_info_text}\n"

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


    def _click_image_with_scroll(self, img_path: str, name: str, threshold: float = 0.82, max_scroll_attempts: int = 15) -> bool:
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
                coords = self._find_image_coords(img_path, threshold=threshold)
                if coords:
                    if coords[1] < mid_top:
                        self._log(f"  📌 {name} 상단 치우침(y={coords[1]}) -> 미세 스크롤 업")
                        self._scroll_up(distance_ratio=0.18)
                        time.sleep(1.0)
                        adj = self._find_image_coords(img_path, threshold=threshold)
                        if adj: coords = adj
                    elif coords[1] > mid_bottom:
                        self._log(f"  📌 {name} 하단 치우침(y={coords[1]}) -> 미세 스크롤 다운")
                        self._scroll_down(distance_ratio=0.18)
                        time.sleep(1.0)
                        adj = self._find_image_coords(img_path, threshold=threshold)
                        if adj: coords = adj

                    self._log(f"  🎯 {name} 이미지 발견! 화면 중앙 좌표 ({coords[0]}, {coords[1]}) -> 탭 클릭")
                    ah.tap_by_coords(self.driver, coords[0], coords[1], self._log)
                    time.sleep(2)
                    return True

            self._log(f"  ⬇ {name} 미발견 -> 미세 스크롤 다운 ({attempt}/{max_scroll_attempts})")
            self._scroll_down(distance_ratio=0.20)
            time.sleep(0.8)
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
                    time.sleep(2)
                    return True
            time.sleep(1)
        self._log(f"  ❌ {name} 버튼 탐색 실패 (스크롤 없음)")
        return False

    def _click_bank_select_with_scroll(self, max_scroll_attempts: int = 5) -> bool:
        """
        [무통장입금 클릭 후 판단]
        '은행을.png' 탐색 (최대 5회 스크롤)
        '은행을.png' 이미지 인식 시 화면 중앙 조절 후 탭 클릭.
        """
        self._set_status("은행 선택 탐색 및 판단 중")
        self._log(f"🔍 [무통장입금 클릭 후 판단] 은행 선택 버튼 탐색 시작 ('은행을.png', 최대 {max_scroll_attempts}회 시도)")

        bank_images = [
            (IMG_SELECT_BANK, "은행을"),
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

        self._log(f"  ❌ [실패] 무통장입금 클릭 후 '은행을' 미발견 (최대 {max_scroll_attempts}회 시도 초과 -> 작업 중단)")
        return False

    def _process_bank_transfer(self) -> bool:
        self._log("💰 [무통장 결제] 프로세스 시작")
        
        # 1. 다른결재 무조건 탐색 및 클릭 시도
        if not self._click_image_with_scroll(IMG_OTHER_PAY, "다른결재", max_scroll_attempts=7):
            self._log("⚠ '다른결재' 버튼 미발견")
            if os.path.exists(IMG_NORMAL_PAY_CHECK) and self._find_image_coords(IMG_NORMAL_PAY_CHECK, threshold=0.70):
                self._log("✅ '일반결재체크' 확인됨. 계속 진행합니다.")
            elif os.path.exists(IMG_BANK_TRANSFER_CHECK) and self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70):
                self._log("✅ '무통장체크' 확인됨. 계속 진행합니다.")
            else:
                self._log("❌ '다른결재' 및 체크 상태 미발견 -> 무통장 결제 실패")
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
            if not self._click_image_with_scroll(IMG_NORMAL_PAY, "일반결재"):
                if os.path.exists(IMG_NORMAL_PAY_CHECK) and self._find_image_coords(IMG_NORMAL_PAY_CHECK, threshold=0.70):
                    self._log("✅ '일반결재체크' 발견! 성공으로 간주하고 진행합니다.")
                elif os.path.exists(IMG_BANK_TRANSFER_CHECK) and self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70):
                    self._log("✅ '무통장체크' 발견! 성공으로 간주하고 진행합니다.")
                else:
                    self._log("❌ '일반결재' 버튼 미발견 -> 무통장 결제 실패")
                    return False
                    
        # 3. 무통장입금 탐색 (이미 체크되어 있으면 스킵)
        is_bank_transfer_checked = False
        if os.path.exists(IMG_BANK_TRANSFER_CHECK) and self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70):
            self._log("✅ '무통장체크' 상태 감지됨! 무통장입금 클릭 건너뜀.")
            is_bank_transfer_checked = True
            
        if not is_bank_transfer_checked:
            if not self._click_image_with_scroll(IMG_BANK_TRANSFER, "무통장입금"):
                if os.path.exists(IMG_BANK_TRANSFER_CHECK) and self._find_image_coords(IMG_BANK_TRANSFER_CHECK, threshold=0.70):
                    self._log("✅ '무통장체크' 상태 감지됨! 성공으로 간주하고 진행합니다.")
                else:
                    self._log("❌ '무통장입금' 버튼 미발견 -> 무통장 결제 실패")
                    return False
            time.sleep(1.5)

        # 무통장입금 클릭(또는 스킵) 후 '은행을' / '은행선택' 존재하는지 판단 (없으면 실패 및 작업 중단)
        if not self._click_bank_select_with_scroll(max_scroll_attempts=5):
            self._log("❌ 무통장입금 클릭 후 '은행을' / '은행선택' 존재하지 않음 -> 무통장 결제 중단 및 실패 처리")
            return False

        # 5개 은행 중 랜덤으로 1개 지정하여 선택 (하나은행, 농협, 우리은행, 국민은행, 기업은행)
        random_bank_options = [
            (IMG_HANA_BANK, "하나은행"),
            (IMG_NONGHYUP_BANK, "농협"),
            (IMG_WOORI_BANK, "우리은행"),
            (IMG_KB_BANK, "국민은행"),
            (IMG_IBK_BANK, "기업은행"),
        ]
        random.shuffle(random_bank_options)

        bank_selected = False
        for bank_img, bank_name in random_bank_options:
            self._log(f"🎲 랜덤 은행 선택 시도: {bank_name}")
            if self._click_image_with_scroll(bank_img, bank_name, threshold=0.70, max_scroll_attempts=5):
                self._log(f"✅ 랜덤 은행 [{bank_name}] 선택 완료")
                bank_selected = True
                break
            else:
                self._log(f"🔄 {bank_name} 미발견 -> 다음 은행 탐색을 위해 스크롤을 맨 위로 올립니다.")
                for _ in range(4):
                    self._scroll_up(distance_ratio=0.5)
                    time.sleep(0.5)

        if not bank_selected:
            self._log("❌ 랜덤 은행 선택 실패 -> 무통장 결제 실패")
            return False
            
        # ── 미신청 탐색 전: 팝업/닫기 버튼 점검 ──
        self._dismiss_payment_benefit_popup()

        if not self._click_image_with_scroll(IMG_NOT_APPLY, "미신청", max_scroll_attempts=4):
            self._log("⚠ '미신청' 미발견 -> '미신청2' 탐색 시도")
            self._scroll_up(distance_ratio=0.6)
            time.sleep(1)
            # 미신청2 전에도 팝업 점검
            self._dismiss_payment_benefit_popup()
            if not self._click_image_with_scroll(IMG_NOT_APPLY2, "미신청2", max_scroll_attempts=4):
                self._log("❌ '미신청' 및 '미신청2' 버튼 미발견 -> 무통장 결제 실패")
                return False

        # ── 결재하기 탐색 전: 팝업/닫기 버튼 점검 ──
        self._dismiss_payment_benefit_popup()

        if not self._click_image_with_scroll(IMG_DO_PAY, "결재하기"):
            self._log("❌ '결재하기' 버튼 미발견 -> 무통장 결제 실패")
            return False

        # ── 주문하기 탐색 전: 팝업/닫기 버튼 점검 ──
        self._dismiss_payment_benefit_popup()

        if not self._click_image_with_scroll(IMG_DO_ORDER, "주문하기"):
            self._log("❌ '주문하기' 버튼 미발견 -> 무통장 결제 실패")
            return False
        return True

    def _dismiss_payment_benefit_popup(self) -> None:
        """
        화면에 '닫기.png' 또는 '결제혜택.png' 팝업이 존재하는지 확인하고, 존재하면 '닫기.png' (또는 닫기 버튼)를 눌러 팝업을 닫습니다.
        팝업이 없거나 닫기 버튼을 찾지 못해도 오류 없이 통과합니다.
        """
        self._log("🔍 [팝업 닫기 점검] 닫기.png / 결제혜택.png 감지 시도...")
        try:
            # 1순위: 화면에 닫기.png 가 존재하는지 직접 탐색하여 클릭
            if os.path.exists(IMG_CLOSE_POPUP):
                close_coords = self._find_image_coords(IMG_CLOSE_POPUP, threshold=0.65)
                if close_coords:
                    self._log(f"  🎯 [닫기.png] 발견! 좌표 ({close_coords[0]}, {close_coords[1]}) → 클릭 닫기")
                    import naver_appium as _ah
                    _ah.tap_by_coords(self.driver, close_coords[0], close_coords[1], self._log)
                    time.sleep(1.0)
                    self._log("  ✅ [닫기.png] 팝업 닫기 완료")
                    return

            # 2순위: 결제혜택.png 가 존재하는지 탐색
            if os.path.exists(IMG_PAY_BENEFIT):
                coords = self._find_image_coords(IMG_PAY_BENEFIT, threshold=0.65)
                if coords:
                    self._log(f"  🎯 [결제혜택 팝업] 감지! 좌표 ({coords[0]}, {coords[1]}) → 닫기 버튼 탐색")
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

            self._log("  ℹ [팝업 닫기 점검] 팝업 미감지 또는 닫기 처리 불필요 → 계속 진행")
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

    # ─── 주문 루프 ────────────────────────────────────────────────────────────

    def _order_loop(self):
        """[단계 2~19] 결재목록 루프 주문 처리"""
        self._log(f"📋 주문 루프 시작 (폰ID 필터: {self.device_id})")

        while not self._stop_event.is_set():
            row = self.order_manager.get_next_pending(device_id=self.device_id)
            if not row:
                self._log("✅ 모든 주문 처리 완료 (해당 기기 대상)")
                break

            self._log(f"📌 처리 중: row={row.row_index}, keyword={row.search_keyword!r}, 폰ID={row.device_id!r}")
            self._set_status(f"주문 중: {row.search_keyword}")

            try:
                success = False
                max_retries = 3
                for attempt in range(max_retries + 1):
                    if self._stop_event.is_set():
                        break
                    
                    if attempt > 0:
                        self._log(f"  🔄 [{attempt}차 실패 재시도] {row.search_keyword} (row={row.row_index}) {attempt}회 재시도 진행...")
                        try:
                            ah.force_stop_and_restart_app(self.driver, self.device_id, self._log)
                            time.sleep(2)
                        except Exception:
                            pass

                    success = self._process_order_with_timeout(row)
                    if success:
                        if attempt > 0:
                            self._log(f"  ✅ [재시도 성공!] {row.search_keyword} (row={row.row_index}) {attempt}회 재시도 성공 -> Y 기록 진행")
                        break
            except Exception as fatal_err:
                self.order_manager.mark_failed(row.row_index)
                self._log(f"❌ 치명적 오류: {fatal_err}")
                raise

            if success:
                if row.payment_method == "무통장":
                    # 무통장: 주문번호 확인되어야 최종 성공 처리
                    try:
                        order_confirmed = self._capture_and_log_bank_transfer(row)
                    except Exception as e:
                        self._log(f"❌ 무통장 스크린샷/로그 처리 중 예외 발생: {e}")
                        order_confirmed = False

                    if order_confirmed:
                        self.order_manager.mark_success(row.row_index)
                        self._log(f"✅ 무통장 주문 성공 (주문번호 확인됨): {row.search_keyword} → Y 기록")
                    else:
                        self.order_manager.mark_failed(row.row_index)
                        self._log(f"❌ 무통장 주문번호 미확인 → F 기록: {row.search_keyword}")
                        self._log("⏳ [무통장 실패] 30초 대기 중...")
                        time.sleep(30)
                        continue
                else:
                    self.order_manager.mark_success(row.row_index)
                    self._log(f"✅ 주문 성공: {row.search_keyword} → Y 기록")

                self._log("⏳ [주문 성공] 완료 후 30초 대기 중...")
                import gc
                gc.collect()
                time.sleep(30)
            else:
                self.order_manager.mark_failed(row.row_index)
                self._log(f"❌ 주문 실패 (총 {max_retries + 1}회 시도 모두 실패): {row.search_keyword} → F 기록")
                import gc
                gc.collect()

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

        # [단계 12] 체크박스/옵션 항목 클릭
        self._click_checkbox(row.product_name)

        # [단계 13] 바로구매 클릭
        if not self._click_buy_now():
            self._log("❌ 바로구매 클릭 실패")
            return False

        # [단계 14 & 16] 배송지 확인 및 선택
        if self._check_current_delivery_address(row.recipient_name, row.phone):
            self._log(f"✅ 목표 배송지 '{row.recipient_name}'가 이미 선택되어 있습니다 (변경 불필요)")
        else:
            self._log("🔄 목표 배송지가 선택되어 있지 않아 변경을 시도합니다.")
            self._click_change_button()
            
            # [단계 15] 스크롤 다운 → 제거 (배송지 목록이 바로 표시되므로 불필요)
            
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

            scales = np.linspace(0.5, 2.0, 20)  # 경량화: 60→20단계, UI 과부하 방지
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

    def _scroll_down(self, distance_ratio: float = 0.20):
        """아래로 미세 스크롤 (Appium swipe + ADB input swipe 보조)"""
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
            _run_cmd(
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
            _run_cmd(
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
            log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
            os.makedirs(log_dir, exist_ok=True)
            log_path = os.path.join(log_dir, f"naver_order_{self.device_id}.log")
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{now}] {full_msg}\n")
        except Exception:
            pass

    def _set_status(self, status: str):
        if self._status_cb:
            self._status_cb(self.device_id, status)
