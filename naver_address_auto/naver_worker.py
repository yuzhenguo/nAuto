"""

naver_worker.py

네이버 앱 자동 주소 등록 워커 (핸드폰 1대 기준)

개발문서.txt의 전체 자동화 흐름 구현

"""



import os

import time

import threading

import subprocess

from typing import Callable, Optional



from selenium.webdriver.common.by import By

from selenium.common.exceptions import WebDriverException, NoSuchElementException



import appium_helper as ah  # noqa: kept for compatibility

try:

    # 네이버 전용 helper를 우선 사용 (쿠팡 appium_helper와 이름 충돌 방지)

    import naver_appium as ah  # type: ignore

except ImportError:

    import appium_helper as ah  # type: ignore  # fallback



from address_manager import AddressManager, AddressRow



# ─── 타임아웃 설정 ────────────────────────────────────────────────────────────

TASK_TIMEOUT_SEC   = 1200   # 행당 최대 작업 시간: 20분

DELETE_LOOP_MAX    = 30     # 삭제 루프 최대 반복 횟수 (무한루프 방지)

SCROLL_MAX         = 20     # 주소 목록 스크롤 최대 횟수



# ─── XPath 상수 ───────────────────────────────────────────────────────────────

# 메인화면 - 네이버 플러스 스토어 탭

STORE_TAB_XPATH = (

    '//android.view.ViewGroup[@content-desc="네이버 플러스 스토어, 버튼,"]'

    '/android.view.ViewGroup[@resource-id="com.nhn.android.search:id/tabIconLayout"]'

    '/android.widget.ImageView[@resource-id="com.nhn.android.search:id/tabIcon"]'

)

# 스토어홈 - 하루 동안 보지 않기

HIDE_BTN_XPATH     = '//android.widget.Button[@text="하루 동안 보지 않기"]'

# 마이쇼핑

MY_SHOPPING_XPATH  = '//android.view.View[@content-desc="마이쇼핑"]'

# 설정

SETTING_XPATH      = '//android.view.View[@content-desc="설정"]'

# 배송지 관리

DELIVERY_MGMT_XPATH = '//android.widget.Button[@text="배송지 관리"]'

# 배송지 목록 ListView (네이버+ 스토어 WebView 내부)

LIST_VIEW_XPATH    = (

    '//android.webkit.WebView[@text="네이버+ 스토어"]'

    '/android.view.View/android.widget.ListView'

)

# 삭제 버튼

DELETE_BTN_XPATH   = '//android.widget.Button[@text="삭제"]'

# 삭제 확인 팝업 OK 버튼

DELETE_CONFIRM_XPATH = '//android.widget.Button[@resource-id="android:id/button1"]'

# 웰컴 모달 / 팝업 닫기 버튼 목록
WELCOME_MODAL_XPATHS = [
    '//android.view.View[@resource-id="joinBeginWelcomeModal"]/android.view.View/android.view.View[2]',
    '//android.widget.Button[@resource-id="btnWelcomeClose"]',
    '//android.view.ViewGroup[@content-desc="홈 버튼"]/android.view.ViewGroup/android.widget.ImageView[@resource-id="com.nhn.android.search.InAppBrowser:id/toolbarIconView"]',
]

# 배송지 신규입력

NEW_ADDRESS_BTN_XPATH = '//android.widget.Button[@text="배송지 신규입력"]'

# 수취인 입력

RECEIVER_XPATH     = '//android.widget.EditText[@resource-id="receiver_name"]'

# 주소검색 버튼

ADDR_SEARCH_BTN_XPATH = '//android.widget.Button[@text="주소검색"]'

# 주소검색 화면 로딩 확인용 (주소 검색 팝업 헤더 View)

ADDR_SEARCH_SCREEN_XPATH = '//android.view.View[@text="주소 검색"]'

# 주소검색 입력창 - hint 기반 (resource-id 없음, WebView 내부)

ADDR_SEARCH_INPUT_XPATHS = [

    '//android.widget.EditText[@hint="지번/도로명 주소"]',

    '//android.widget.EditText[@hint="지번/도로명을 입력해주세요."]',

    '//android.widget.EditText[@hint="지번/도로명"]',

    '//android.view.View[@text="주소 검색"]/following-sibling::*//android.widget.EditText',

]

# 검색 버튼 - 주소검색 화면 내 (text="검색")

SEARCH_BTN_XPATH   = '//android.widget.Button[@text="검색"]'

# 주소 목록 ListView

ADDR_LIST_XPATH    = '//android.widget.ListView'

# 전화번호 중간 4자리 입력 EditText[1]

PHONE_MID_XPATH    = (

    '//android.webkit.WebView[@text="네이버+ 스토어"]'

    '/android.view.View/android.view.View[2]/android.view.View'

    '/android.view.View[2]/android.view.View[2]/android.widget.EditText[1]'

)

# 전화번호 마지막 4자리 입력 EditText[2]

PHONE_LAST_XPATH   = (

    '//android.webkit.WebView[@text="네이버+ 스토어"]'

    '/android.view.View/android.view.View[2]/android.view.View'

    '/android.view.View[2]/android.view.View[2]/android.widget.EditText[2]'

)

# 등록 버튼

REGISTER_BTN_XPATH = '//android.widget.Button[@text="등록"]'





class NaverWorker:

    """

    네이버 앱 배송지 자동 등록 워커

    - 기존 배송지 삭제 (9열 Y인 경우, 첫 행 최초 1회만 판단)

    - 엑셀 기반 새 배송지 등록 루프

    """



    def __init__(self,

                 device_id: str,

                 appium_port: int,

                 address_manager: AddressManager,

                 log_callback: Optional[Callable] = None,

                 status_callback: Optional[Callable] = None):

        self.device_id       = device_id

        self.appium_port     = appium_port

        self.address_manager = address_manager

        self._log_cb         = log_callback

        self._status_cb      = status_callback

        self.driver          = None

        self._stop_event     = threading.Event()

        self.has_checked_initial_delete = False  # 삭제 여부 최초 1회만 판단

        self.current_naver_id = None

        self.is_initialized = False



    # ─── 공개 메서드 ──────────────────────────────────────────────────────────



    def run(self) -> bool:

        """워커 메인 실행 (별도 스레드에서 호출)"""

        self._log("🚀 워커 시작")

        max_restarts = 20



        for attempt in range(1, max_restarts + 1):

            if self._stop_event.is_set():

                break



            self._set_status(f"연결 중... (시도 {attempt}/{max_restarts})")

            try:

                self.driver = ah.create_driver(

                    self.device_id, self.appium_port, self._log

                )

                self._log("🔄 네이버 앱 종료 후 재시작 진행")

                ah.force_stop_and_restart_app(self.driver, self.device_id, self._log)

            except Exception as e:

                self._log(f"❌ 드라이버 연결 실패: {e}")

                self._set_status("연결 실패")

                if attempt < max_restarts:

                    time.sleep(3)

                    continue

                else:

                    return False



            success = False

            try:

                # 배송지 등록 루프 (내부에서 계정 전환 및 화면 진입 처리)

                self._register_address_loop()

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



            if attempt < max_restarts:

                self._log(f"🔄 오류 회복을 위해 5초 후 워커를 재시작합니다... ({attempt}/{max_restarts})")

                time.sleep(5)

            else:

                self._log("❌ 최대 재시작 횟수 초과. 워커를 완전히 종료합니다.")



        self._log("🏁 워커 종료")

        self._set_status("완료")

        return True



    def stop(self):

        """워커 중지 요청"""

        self._stop_event.set()

        self._log("⏹ 중지 요청됨")



    # ─── WebView 로딩 대기 헬퍼 ────────────────────────────────────────────────



    def _wait_webview_ready(self, label: str = "", timeout: int = 20,

                             webview_text: str = "네이버+ 스토어") -> bool:

        """

        WebView 요소가 화면에 나타날 때까지 대기.

        UiAutomator2 NATIVE 컨텍스트에서 JS(execute_script) 가 미지원이므로

        JavaScript 없이 element 존재 확인 방식으로 대체.

        """

        tag = f"[WebView 로딩{'/' + label if label else ''}]"

        self._log(f"  ⏳ {tag} 대기 시작 (최대 {timeout}초)")



        wv_xpaths = [

            f'//android.webkit.WebView[@text="{webview_text}"]',

            '//android.webkit.WebView',

            '//com.naver.xwhale.WebView',

        ]



        for elapsed in range(timeout):

            for wv_xpath in wv_xpaths:

                try:

                    els = self.driver.find_elements(By.XPATH, wv_xpath)

                    if els:

                        self._log(f"  ✅ {tag} WebView 감지 ({elapsed}초) → 1초 추가 안정화 대기")

                        time.sleep(1)

                        return True

                except Exception:

                    pass

            self._log(f"  📊 {tag} {elapsed}s | WebView 미감지")

            time.sleep(1)



        self._log(f"  ⚠ {tag} {timeout}초 대기 후에도 WebView 미감지 → 계속 진행")

        return False



    # ─── 내비게이션 ───────────────────────────────────────────────────────────



    def _go_main_and_enter_store(self, login_id: str = "") -> bool:

        """

        [단계 1~6] 앱 메인 → 네이버 플러스 스토어 탭 클릭

        → 팝업 처리 → 마이쇼핑 클릭 → 팝업 처리

        """

        self._set_status("메인 페이지 이동 중")

        ah.go_to_main_page(self.driver, self._log)

        time.sleep(7)

        # [단계 3.1] 웰컴 모달 / 팝업 발견 시 클릭
        self._check_and_close_welcome_modals(step_label="3.1")

        # [단계 3.2] 로그인 아이디가 지정된 경우 네이버 계정 전환 수행
        if login_id:
            if not self._switch_account(login_id):
                self._log(f"❌ [단계 3.2] 계정 전환 실패 (아이디: {login_id})")
                return False
            # 계정 전환 성공 후 메인 페이지로 이동
            ah.go_to_main_page(self.driver, self._log)
            time.sleep(3)


        # [단계 3] 네이버 플러스 스토어 탭 버튼이 존재하면 클릭

        if ah.element_exists(self.driver, STORE_TAB_XPATH, timeout=5):

            self._log("📌 네이버 플러스 스토어 탭 감지 → 클릭")

            ah.wait_and_click(self.driver, STORE_TAB_XPATH, timeout=7, log_callback=self._log)

            time.sleep(5)

        else:

            self._log("⏭ 네이버 플러스 스토어 탭 없음 (이미 스토어 화면이거나 탭 없음)")

            time.sleep(3)



        # [단계 4] 스토어홈 - 팝업 처리 ('7일간 보지 않기' / '하루 동안 보지 않기' / 7일간.png) (최대 2회)
        self._dismiss_hide_popup(max_count=2)

        # [단계 5] 마이쇼핑 클릭 (7일간 팝업 가림 방지)
        self._set_status("마이쇼핑 클릭")
        my_shopping_xpaths = [
            MY_SHOPPING_XPATH,
            '//*[contains(@content-desc, "마이쇼핑")]',
            '//*[contains(@text, "마이쇼핑")]',
        ]
        my_shopping_clicked = False
        for attempt in range(1, 3):
            self._dismiss_hide_popup(max_count=1)
            for xpath in my_shopping_xpaths:
                if ah.element_exists(self.driver, xpath, timeout=4):
                    self._log("📌 마이쇼핑 버튼 발견 → 클릭")
                    ah.wait_and_click(self.driver, xpath, timeout=5, log_callback=self._log)
                    self._log("✅ 마이쇼핑 클릭 완료 (5초 대기)")
                    time.sleep(5)
                    my_shopping_clicked = True
                    break
            if my_shopping_clicked:
                break
            time.sleep(1)

        if not my_shopping_clicked:
            self._log("⚠ 마이쇼핑 버튼 미발견. 계속 진행합니다.")
            time.sleep(2)

        self._dismiss_hide_popup(max_count=1)
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

        # 1.5. 설정 버튼 찾기 전 "7일간 보지 않기" 팝업이 뜨면 클릭 (7일간.png 참고)
        dismiss_7days_xpaths = [
            '//android.widget.Button[@text="7일간 보지 않기"]',
            '//android.widget.Button[contains(@text, "7일")]',
            '//*[contains(@text, "7일간 보지 않기")]',
            '//*[contains(@text, "7일간 보이지 않기")]',
            '//*[contains(@text, "7일 동안 보지 않기")]',
            '//*[contains(@text, "7일간")]',
            '//*[contains(@content-desc, "7일간")]',
            '//*[contains(@content-desc, "7일 동안")]',
        ]
        for xpath in dismiss_7days_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=2):
                self._log("  📌 [단계 3.2] '7일간 보이지 않기' 팝업 발견 -> 클릭")
                ah.wait_and_click(self.driver, xpath, timeout=3, log_callback=self._log)
                time.sleep(2)
                break

        # 7일간.png 이미지 템플릿 매칭 보조 (XPath 미감지 시 대비)
        tpl_7days = os.path.join(os.path.dirname(__file__), "7일간.png")
        if os.path.exists(tpl_7days):
            try:
                import cv2
                import numpy as np
                import subprocess

                res = subprocess.run(["adb", "-s", self.device_id, "exec-out", "screencap", "-p"], capture_output=True, timeout=5)
                if res.stdout and len(res.stdout) > 100:
                    img_arr = np.frombuffer(res.stdout, np.uint8)
                    screen_img = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
                    template_img = cv2.imread(tpl_7days, cv2.IMREAD_COLOR)
                    if screen_img is not None and template_img is not None:
                        result = cv2.matchTemplate(screen_img, template_img, cv2.TM_CCOEFF_NORMED)
                        _, max_val, _, max_loc = cv2.minMaxLoc(result)
                        if max_val >= 0.65:
                            h, w = template_img.shape[:2]
                            cx = max_loc[0] + w // 2
                            cy = max_loc[1] + h // 2
                            self._log(f"  📌 [이미지 인식] '7일간.png' 발견 (일치율: {max_val:.2f}) -> 좌표 탭: ({cx}, {cy})")
                            subprocess.run(["adb", "-s", self.device_id, "shell", "input", "tap", str(cx), str(cy)], capture_output=True, timeout=3)
                            time.sleep(2)
            except Exception as e:
                self._log(f"  ⚠ 7일간.png 이미지 인식 예외: {e}")

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

        # 3.5. 오타 보정 로직: 화면에 보이는 아이디들과 비교하여 가장 비슷한 아이디(actual_login_id)로 보정
        # resource-id 비존재 시 content-desc (간편로그인/로그인 중) 기반으로 수집
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

        # XPath 미감지 시 → 화면의 '로그인 중' 요소 전수 조사 (difflib 유사도)
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
            target_el = None
            for scroll_idx in range(5):
                # 아이디 오타/접두사 보정 재검사
                try:
                    id_els = self.driver.find_elements(By.XPATH, '//*[@resource-id="com.nhn.android.search:id/idText"] | //*[@content-desc]')
                    available_ids = []
                    for el in id_els:
                        txt = el.text or el.get_attribute("content-desc") or ""
                        if txt:
                            txt = txt.replace(", 더보기", "").replace(" , 간편로그인", "").replace(", 간편로그인", "").replace(", 로그인 중", "").strip()
                            if len(txt) >= 3:
                                available_ids.append(txt)

                    if available_ids:
                        if login_id not in available_ids:
                            import difflib
                            matches = difflib.get_close_matches(login_id, available_ids, n=1, cutoff=0.55)
                            if matches:
                                actual_login_id = matches[0]
                                self._log(f"  ℹ 아이디 유사 보정 (스크롤 {scroll_idx+1}): '{login_id}' -> '{actual_login_id}'")
                            else:
                                for aid in available_ids:
                                    if aid.startswith(login_id[:3]) or login_id.startswith(aid[:3]):
                                        actual_login_id = aid
                                        self._log(f"  ℹ 아이디 접두사 보정 (스크롤 {scroll_idx+1}): '{login_id}' -> '{actual_login_id}'")
                                        break
                except Exception:
                    pass

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

                for xpath in target_account_xpaths:
                    if ah.element_exists(self.driver, xpath, timeout=2):
                        try:
                            target_el = self.driver.find_element(By.XPATH, xpath)
                            self._log(f"  📌 아이디선택 화면에서 '{actual_login_id}' 발견 -> 강력 클릭 진행 (스크롤 {scroll_idx+1}회차)")
                            break
                        except Exception:
                            continue

                if target_el:
                    break

                if scroll_idx < 4:
                    self._log(f"  📜 타겟 아이디 [{actual_login_id}] 탐색 중... 목록 하단 스크롤 ({scroll_idx+1}/4)")
                    try:
                        ah.swipe_up(self.driver)
                    except Exception:
                        subprocess.run(
                            ["adb", "-s", self.device_id, "shell", "input", "swipe", "500", "1500", "500", "800", "300"],
                            capture_output=True, timeout=5
                        )
                    time.sleep(1.5)

            if not target_el:
                self._log(f"  ❌ [단계 3.2] 로그인 아이디 [{actual_login_id}] 5회 스크롤 후에도 미발견 -> 계정 전환 실패")
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
                import subprocess
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

                time.sleep(1.5)

                # 아이디 클릭 후 팝업(android:id/message) 감지 시 확인 버튼(android:id/button1) 클릭
                msg_xpaths = [
                    '//android.widget.TextView[@resource-id="android:id/message"]',
                    '//*[@resource-id="android:id/message"]',
                ]
                btn1_xpaths = [
                    '//android.widget.Button[@resource-id="android:id/button1"]',
                    '//*[@resource-id="android:id/button1"]',
                ]
                for msg_xpath in msg_xpaths:
                    if ah.element_exists(self.driver, msg_xpath, timeout=2):
                        self._log("  📌 [단계 3.2] 알림 메시지 팝업 감지 (android:id/message) -> 확인 버튼(button1) 클릭")
                        for btn_xpath in btn1_xpaths:
                            if ah.element_exists(self.driver, btn_xpath, timeout=3):
                                ah.wait_and_click(self.driver, btn_xpath, timeout=3, log_callback=self._log)
                                self._log("  ✅ 확인 버튼(button1) 클릭 완료")
                                time.sleep(2)
                                break
                        break

                time.sleep(1.5)

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


    def _dismiss_hide_popup(self, max_count: int = 2):
        """'하루 동안 보지 않기' / '7일간 보지 않기' / '동일한 주소 존재' 팝업 반복 처리"""
        popup_xpaths = [
            '//android.widget.Button[@text="7일간 보지 않기"]',
            '//android.widget.Button[@text="하루 동안 보지 않기"]',
            '//android.widget.Button[contains(@text, "7일")]',
            '//*[contains(@text, "7일간 보지 않기")]',
            '//*[contains(@text, "7일간 보이지 않기")]',
            '//*[contains(@text, "7일 동안 보지 않기")]',
            '//*[contains(@text, "하루 동안 보지 않기")]',
        ]

        duplicate_pop_xpaths = [
            '//*[contains(@text, "동일한 주소가 존재")]',
            '//*[contains(@text, "입력/수정할 수 없습니다")]',
            '//*[contains(@text, "shopping.naver.com")]',
        ]

        for i in range(max_count):
            dismissed = False

            # 1. 중복 배송지 알림 (페이지 내용: 회원의 배송지 목록에 동일한 주소가 존재...) 감지 및 확인 클릭
            for dup_xpath in duplicate_pop_xpaths:
                if ah.element_exists(self.driver, dup_xpath, timeout=1):
                    self._log("📌 중복 배송지 알림 팝업 감지 (동일한 주소 존재) → '확인' 버튼 클릭")
                    confirm_btn_xpaths = [
                        '//android.widget.Button[@text="확인"]',
                        '//android.widget.Button[@resource-id="android:id/button1"]',
                        '//*[@text="확인"]',
                    ]
                    for c_xpath in confirm_btn_xpaths:
                        if ah.element_exists(self.driver, c_xpath, timeout=2):
                            ah.wait_and_click(self.driver, c_xpath, timeout=3, log_callback=self._log)
                            time.sleep(1.5)
                            dismissed = True
                            break

            # 2. 일반 팝업 감지
            for xpath in popup_xpaths:
                if ah.element_exists(self.driver, xpath, timeout=2):
                    self._log(f"📌 팝업 감지 ({xpath[:30]}) → 클릭 (회차 {i+1})")
                    ah.wait_and_click(self.driver, xpath, timeout=3, log_callback=self._log)
                    time.sleep(2)
                    dismissed = True
                    break

            # 7일간.png 이미지 템플릿 매칭 검사
            tpl_7days = os.path.join(os.path.dirname(__file__), "7일간.png")
            if not dismissed and os.path.exists(tpl_7days):
                try:
                    import cv2
                    import numpy as np
                    import subprocess

                    res = subprocess.run(["adb", "-s", self.device_id, "exec-out", "screencap", "-p"], capture_output=True, timeout=5)
                    if res.stdout and len(res.stdout) > 100:
                        img_arr = np.frombuffer(res.stdout, np.uint8)
                        screen_img = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
                        template_img = cv2.imread(tpl_7days, cv2.IMREAD_COLOR)
                        if screen_img is not None and template_img is not None:
                            result = cv2.matchTemplate(screen_img, template_img, cv2.TM_CCOEFF_NORMED)
                            _, max_val, _, max_loc = cv2.minMaxLoc(result)
                            if max_val >= 0.65:
                                h, w = template_img.shape[:2]
                                cx = max_loc[0] + w // 2
                                cy = max_loc[1] + h // 2
                                self._log(f"📌 [이미지 인식] '7일간.png' 발견 (일치율: {max_val:.2f}) -> 좌표 탭: ({cx}, {cy})")
                                subprocess.run(["adb", "-s", self.device_id, "shell", "input", "tap", str(cx), str(cy)], capture_output=True, timeout=3)
                                time.sleep(2)
                                dismissed = True
                except Exception as e:
                    self._log(f"  ⚠ 7일간.png 이미지 인식 예외: {e}")

            if not dismissed:
                break

    def _navigate_to_delivery_mgmt(self) -> bool:
        """[단계 7~8] 설정 → 배송지 관리 화면 진입"""
        # [단계 7 전] "7일간 보지 않기" / "하루 동안 보지 않기" / 7일간.png 사전 팝업 감지 및 클릭
        self._dismiss_hide_popup(max_count=3)

        # [단계 7] 설정 클릭 (7일간 팝업 재감지 및 다중 후보 XPath 적용)
        self._set_status("설정 클릭")
        setting_clicked = False

        setting_xpaths = [
            '//android.view.View[@content-desc="설정"]',
            '//*[@content-desc="설정"]',
            '//android.widget.ImageView[@content-desc="설정"]',
            '//android.view.ViewGroup[@content-desc="설정"]',
            '//android.view.View[contains(@content-desc, "설정")]',
            '//*[contains(@content-desc, "설정")]',
            '//android.widget.Button[@text="설정"]',
            '//android.view.View[@text="설정"]',
        ]

        for attempt in range(1, 4):
            # 매 시도마다 7일간 보지 않기 / 7일간.png 팝업 체크
            self._dismiss_hide_popup(max_count=1)

            for xpath in setting_xpaths:
                if ah.element_exists(self.driver, xpath, timeout=3):
                    self._log(f"📌 설정 버튼 발견 ({xpath[:35]}) → 클릭 (시도 {attempt}/3)")
                    if ah.wait_and_click(self.driver, xpath, timeout=4, log_callback=self._log):
                        setting_clicked = True
                        time.sleep(2.5)
                        break

            if setting_clicked:
                break

            self._log(f"  ⚠ 설정 버튼 미발견 ({attempt}/3) -> 7일간 팝업 재감지 후 재시도...")
            time.sleep(1.5)

        if not setting_clicked:
            self._log("  🔄 설정 버튼 미발견 -> 메인 복구 후 스토어/마이쇼핑 재진입 시도...")
            ah.go_to_main_page(self.driver, self._log)
            time.sleep(3)
            self._dismiss_hide_popup(max_count=2)

            if ah.element_exists(self.driver, STORE_TAB_XPATH, timeout=5):
                ah.wait_and_click(self.driver, STORE_TAB_XPATH, timeout=5, log_callback=self._log)
                time.sleep(4)

            self._dismiss_hide_popup(max_count=2)
            if ah.element_exists(self.driver, MY_SHOPPING_XPATH, timeout=5):
                ah.wait_and_click(self.driver, MY_SHOPPING_XPATH, timeout=5, log_callback=self._log)
                time.sleep(4)

            self._dismiss_hide_popup(max_count=2)
            for xpath in setting_xpaths:
                if ah.element_exists(self.driver, xpath, timeout=4):
                    if ah.wait_and_click(self.driver, xpath, timeout=4, log_callback=self._log):
                        setting_clicked = True
                        time.sleep(2.5)
                        break

        if not setting_clicked:
            self._log("❌ 설정 버튼을 최종적으로 찾지 못했습니다.")
            return False

        # [단계 8] 배송지 관리 클릭
        self._set_status("배송지 관리 클릭")
        delivery_mgmt_xpaths = [
            DELIVERY_MGMT_XPATH,
            '//*[contains(@text, "배송지 관리")]',
            '//*[contains(@content-desc, "배송지 관리")]',
        ]
        mgmt_clicked = False
        for xpath in delivery_mgmt_xpaths:
            if ah.element_exists(self.driver, xpath, timeout=5):
                if ah.wait_and_click(self.driver, xpath, timeout=5, log_callback=self._log):
                    mgmt_clicked = True
                    break

        if not mgmt_clicked:
            self._log("❌ 배송지 관리 버튼을 찾지 못했습니다.")
            return False

        time.sleep(3)
        self._log("✅ 배송지 관리 화면 진입")
        return True



    # ─── 삭제 로직 ────────────────────────────────────────────────────────────



    def _delete_existing_addresses(self):

        """

        [단계 9.1~9.3] 기본배송지 제외 기존 배송지 모두 삭제
        삭제.png 이미지 인식으로 삭제 버튼 탐색, 스크롤 반복 (최대 15회)

        """

        self._set_status("기존 배송지 삭제 중")

        self._log("🗑 기존 배송지 삭제 시작")

        # 배송지 관리 화면 로드 대기 (최대 10초)
        ah.wait_for_element(self.driver, NEW_ADDRESS_BTN_XPATH, timeout=10, log_callback=self._log)

        delete_img_path = os.path.join(os.path.dirname(__file__), "삭제.png")
        img_based = os.path.exists(delete_img_path)
        if img_based:
            self._log(f"  📂 삭제.png 이미지 파일 발견: {delete_img_path}")
        else:
            self._log(f"  ⚠ 삭제.png 파일 없음 → XPath 기반 삭제로 전환 ({delete_img_path})")

        MAX_LOOPS = 15
        deleted_count = 0

        for loop_count in range(MAX_LOOPS):

            if self._stop_event.is_set():
                break

            time.sleep(1)

            tap_x, tap_y = None, None

            # ─── 1순위: 삭제.png 이미지 매칭으로 삭제 버튼 좌표 탐색 ───
            if img_based:
                coords = self._find_image_coords(delete_img_path, threshold=0.70)
                if coords:
                    tap_x, tap_y = coords
                    self._log(f"  🎯 [이미지 매칭] 삭제.png 발견! 좌표: ({tap_x}, {tap_y})")

            # ─── 2순위: 이미지 미감지 시 XPath로 폴백 ───
            if tap_x is None:
                target_btn = self._find_non_default_delete_button()
                if target_btn:
                    try:
                        rect = target_btn.rect
                        tap_x = int(rect['x'] + rect['width'] // 2)
                        tap_y = int(rect['y'] + rect['height'] // 2)
                        self._log(f"  📌 [XPath 폴백] 삭제 버튼 발견 좌표: ({tap_x}, {tap_y})")
                    except Exception:
                        try:
                            target_btn.click()
                            self._log(f"  📌 [XPath 폴백] 삭제 버튼 direct click (loop {loop_count + 1})")
                            tap_x = 0  # click 완료 표시용
                            tap_y = 0
                        except Exception as e:
                            self._log(f"  ⚠ 삭제 버튼 클릭 실패: {e}")

            # ─── 삭제 버튼 미발견: 스크롤 후 재탐색, 없으면 종료 ───
            if tap_x is None:
                self._log(f"  📜 삭제 버튼 미발견 → 스크롤 후 재탐색 (loop {loop_count + 1})")
                self._scroll_down()
                time.sleep(1.5)

                # 스크롤 후 재탐색
                if img_based:
                    coords = self._find_image_coords(delete_img_path, threshold=0.70)
                    if coords:
                        tap_x, tap_y = coords
                        self._log(f"  🎯 [스크롤 후 이미지 매칭] 삭제.png 발견! 좌표: ({tap_x}, {tap_y})")

                if tap_x is None:
                    target_btn = self._find_non_default_delete_button()
                    if target_btn:
                        try:
                            rect = target_btn.rect
                            tap_x = int(rect['x'] + rect['width'] // 2)
                            tap_y = int(rect['y'] + rect['height'] // 2)
                        except Exception:
                            pass

                if tap_x is None:
                    self._log("✅ 삭제 완료 (더 이상 삭제할 주소 없음)")
                    break

            # ─── ADB tap 으로 삭제 버튼 클릭 ───
            if tap_x is not None and not (tap_x == 0 and tap_y == 0):
                try:
                    import subprocess
                    subprocess.run(
                        ["adb", "-s", self.device_id, "shell", "input", "tap",
                         str(tap_x), str(tap_y)],
                        capture_output=True, timeout=5
                    )
                    self._log(f"  ✅ 삭제 버튼 탭 완료 (loop {loop_count + 1}, 좌표: {tap_x},{tap_y})")
                except Exception as e:
                    self._log(f"  ⚠ ADB tap 실패: {e}")

            time.sleep(1)

            # [9.2] 삭제 확인 팝업에서 OK 버튼 클릭
            if ah.element_exists(self.driver, DELETE_CONFIRM_XPATH, timeout=4):
                ah.wait_and_click(self.driver, DELETE_CONFIRM_XPATH, timeout=4, log_callback=self._log)
                self._log("  삭제 확인 OK 클릭 완료")
            else:
                if ah.element_exists(self.driver, '//android.widget.Button[@text="확인"]', timeout=2):
                    ah.wait_and_click(self.driver, '//android.widget.Button[@text="확인"]',
                                      timeout=2, log_callback=self._log)

            deleted_count += 1

            # 삭제 완료 후 WebView 재로딩 대기
            self._wait_webview_ready(label="삭제 후", timeout=10)

        self._log(f"🗑 삭제 루프 종료 (총 {deleted_count}건 삭제)")

        # 삭제 루프 전체 완료 후 WebView 완전 로딩 대기
        self._wait_webview_ready(label="삭제 루프 완료", timeout=15)




    def _find_non_default_delete_button(self):

        """

        ListView 하위 요소 중 '기본배송지' 텍스트가 없는 항목의 삭제 버튼 반환

        """

        try:

            # 1. full XPATH로 직접 아이템들 찾기 (상대 경로 버그 회피 및 가로폭 필터링)

            items = []

            for items_xpath in [

                '//android.webkit.WebView[@text="네이버+ 스토어"]/android.view.View/android.widget.ListView/android.view.View',

                '//android.webkit.WebView/android.view.View/android.widget.ListView/android.view.View',

                '//android.widget.ListView/android.view.View'

            ]:

                try:

                    raw_items = self.driver.find_elements(By.XPATH, items_xpath)

                    if raw_items:

                        # 가로 폭이 넓은 항목(실제 배송지 카드)만 필터링 (버튼 등 내부 요소 제외)

                        filtered_items = []

                        for item in raw_items:

                            try:

                                rect = item.rect

                                if rect['width'] > 800:

                                    filtered_items.append(item)

                            except Exception:

                                filtered_items.append(item)

                        

                        if filtered_items:

                            items = filtered_items

                            self._log(f"  [디버그] 배송지 항목 수: {len(items)} (xpath: {items_xpath})")

                            break

                except WebDriverException:

                    continue



            if not items:

                self._log("  [디버그] 배송지 항목(ListView 하위 View)을 찾을 수 없습니다.")

                raise NoSuchElementException("ListView items not found")



            for idx, item in enumerate(items):

                try:

                    # 기본배송지 텍스트 포함 여부 확인

                    default_tags = item.find_elements(

                        By.XPATH, './/*[@text="기본배송지"]'

                    )

                    if default_tags:

                        self._log(f"  [디버그] 항목 {idx}: 기본배송지 → 건너뜀")

                        continue



                    # 이 항목의 삭제 버튼 반환

                    del_btns = item.find_elements(By.XPATH, './/android.widget.Button[@text="삭제"]')

                    if del_btns:

                        self._log(f"  [디버그] 항목 {idx}: 삭제 버튼 발견")

                        return del_btns[0]

                except (NoSuchElementException, WebDriverException):

                    continue



        except (NoSuchElementException, WebDriverException) as e:

            self._log(f"  [디버그] ListView 조회 중 오류/폴백 진행: {e}")

            # 폴백: 화면 전체에서 삭제 버튼 목록을 찾아 첫 번째 삭제 버튼 반환

            # 기본배송지는 삭제 버튼이 없으므로, 화면에 보이는 첫 번째 삭제 버튼은 무조건 삭제 대상임!

            try:

                all_del = self.driver.find_elements(By.XPATH, DELETE_BTN_XPATH)

                if all_del:

                    self._log(f"  [디버그] 폴백: 전체 화면에서 삭제 버튼 {len(all_del)}개 발견, 첫 번째 반환")

                    return all_del[0]

            except WebDriverException:

                pass



        return None



    # ─── 등록 루프 ────────────────────────────────────────────────────────────



    def _register_address_loop(self):

        """

        [단계 9~22] 엑셀에서 미처리 행을 순차적으로 읽어 배송지 등록

        """

        self._log("📋 주소 등록 루프 시작")



        while not self._stop_event.is_set():

            row = self.address_manager.get_next_pending_row(self.device_id)

            if not row:

                self._log("✅ 모든 주소 처리 완료")

                break
                
            # 폰 ID(기기 ID) 명시적 확인 (자신의 기기에 할당된 작업만 수행)
            if row.device_id != self.device_id:
                self._log(f"⏭ [건너뜀] 기기 ID 불일치 (내 기기: {self.device_id}, 할당: {row.device_id})")
                self.address_manager.mark_failed(row.row_index)
                continue



            self._log(f"📌 처리 중: row={row.row_index}, name={row.name}")

            self._set_status(f"등록 중: {row.name}")



            try:
                success = False
                max_retries = 3
                for attempt in range(max_retries + 1):
                    if self._stop_event.is_set():
                        break

                    if attempt > 0:
                        self._log(f"  🔄 [{attempt}차 실패 재시도] {row.name} (row={row.row_index}) {attempt}회 재시도 진행...")
                        try:
                            ah.go_to_main_page(self.driver, self._log)
                            time.sleep(2)
                            self._dismiss_hide_popup(max_count=2)
                        except Exception:
                            pass
                        self.is_initialized = False

                    success = self._process_row_with_timeout(row)
                    if success:
                        if attempt > 0:
                            self._log(f"  ✅ [재시도 성공!] {row.name} (row={row.row_index}) {attempt}회 재시도 성공 -> Y 기록 진행")
                        break

            except Exception as fatal_err:
                # 치명적 세션/서버 에러 → 현재 행을 F 처리 후 루프 밖으로 전파
                self.address_manager.mark_failed(row.row_index)
                self._log(f"❌ 실패/타임아웃: {row.name} → F 기록")
                self._log(f"🔴 [등록 루프] 치명적 오류로 루프 중단 → run() 재연결 유도")
                raise  # run()의 except Exception as e 로 전파

            if success:
                self.address_manager.mark_success(row.row_index)
                self._log(f"✅ 성공: {row.name} → Y 기록")
            else:
                self.address_manager.mark_failed(row.row_index)
                self._log(f"❌ 실패/타임아웃 (총 {max_retries + 1}회 시도 모두 실패): {row.name} → F 기록")



        self._log("📋 등록 루프 종료")



    def _process_row_with_timeout(self, row: AddressRow) -> bool:

        """행 처리를 타임아웃으로 실행"""

        result = [False]

        exception = [None]



        def task():

            try:

                result[0] = self._register_one_address(row)

            except Exception as e:

                exception[0] = e

                result[0] = False



        t = threading.Thread(target=task, daemon=True)

        t.start()

        t.join(timeout=TASK_TIMEOUT_SEC)



        if t.is_alive():

            self._log(f"⏰ 타임아웃 ({TASK_TIMEOUT_SEC}초) - 다음 행으로")

            return False



        if exception[0]:

            self._log(f"❌ 처리 중 예외: {exception[0]}")

            # ── 치명적 세션/서버 에러 감지 → 상위로 전파 (재연결 유도) ──
            err_str = str(exception[0])
            FATAL_PATTERNS = [
                "instrumentation process is not running",
                "A session is either terminated or not started",
                "NoSuchDriverError",
                "UnknownError",
                "UiAutomation not connected",
                "cannot be proxied to UiAutomator2",
            ]
            if any(p in err_str for p in FATAL_PATTERNS):
                self._log("🔴 [치명적 오류] UiAutomator2/세션 크래시 감지 → 재연결을 위해 예외 전파")
                raise exception[0]  # run() 루프가 재연결 처리

            return False



        return result[0]



    def _register_one_address(self, row: AddressRow) -> bool:

        """

        배송지 1건 등록 전체 흐름 (단계 9~22)

        """

        # 계정 전환 및 화면 진입
        if getattr(self, 'current_naver_id', None) != row.naver_id or not getattr(self, 'is_initialized', False):
            self._log(f"🔄 계정 전환/초기화 (현재: {getattr(self, 'current_naver_id', '없음')} -> 대상: {row.naver_id})")
            if not self._go_main_and_enter_store(login_id=row.naver_id):
                self._log("❌ 메인 이동 및 계정 전환 실패")
                return False
            if not self._navigate_to_delivery_mgmt():
                self._log("❌ 배송지 관리 화면 진입 실패")
                return False
            self.current_naver_id = row.naver_id
            self.is_initialized = True

        # [단계 9] 9열(주소초기화) 값을 최초 1회만 확인하여 기존 배송지 삭제 여부 결정

        if not self.has_checked_initial_delete:

            self.has_checked_initial_delete = True

            if row.delete_existing:

                self._log("🗑 [첫 번째 행] 9열 값 Y → 기존 배송지 삭제 진행 (9.1~9.3)")

                self._delete_existing_addresses()

            else:

                self._log("⏭ [첫 번째 행] 9열 값 Y 아님 → 기존 배송지 삭제 건너뜀")

        else:

            self._log("⏭ [이후 행] 삭제 판단은 이미 1회 수행됨 → 건너뜀")



        # [단계 10] 배송지 신규입력 클릭, 3초 대기

        self._set_status("배송지 신규입력 클릭")

        if not self._ensure_delivery_mgmt_screen():

            self._log("❌ 배송지 관리 화면으로 복귀 실패")

            return False



        if not ah.wait_and_click(self.driver, NEW_ADDRESS_BTN_XPATH, timeout=10, log_callback=self._log):

            self._log("❌ 배송지 신규입력 버튼 클릭 실패")

            return False

        time.sleep(3)



        # [단계 11] 배송지입력.xml 구조 - 수취인 입력 (1열)

        self._set_status(f"수취인 입력: {row.name}")

        if not ah.wait_and_input(self.driver, RECEIVER_XPATH, row.name, timeout=10, log_callback=self._log):

            self._log("❌ 수취인 입력 실패")

            return False



        # [단계 12] 주소검색 버튼 클릭

        time.sleep(1)

        if not ah.wait_and_click(self.driver, ADDR_SEARCH_BTN_XPATH, timeout=5, log_callback=self._log):

            self._log("❌ 주소검색 버튼 클릭 실패")

            return False

        # 팝업 렌더링 대기 (2→5초)

        time.sleep(5)



        # [단계 13] 주소검색 화면 로딩 대기 후 주소 검색어 입력 (2열)

        self._set_status(f"주소 검색: {row.address_search}")

        if not self._input_address_search(row.address_search):

            self._log("❌ 주소 검색창 입력 실패")

            return False



        # [단계 14] 검색 버튼 클릭 (좌표 폴백 포함)

        if not self._click_search_button():

            self._log("❌ 검색 버튼 클릭 실패")

            return False

        self._log("  ⏳ 검색 버튼 클릭 후 3.5초 대기 (결과 로딩)")

        time.sleep(3.5)



        # [단계 15] 우편번호 매칭 주소 선택 (3열)

        self._set_status(f"우편번호 매칭: {row.zipcode}")

        if not self._select_address_by_zipcode(row.zipcode):

            self._log(f"❌ 우편번호 {row.zipcode} 매칭 실패")

            return False



        # [단계 16/17] 상세주소 입력 화면 - 5열 값 타이핑, 2초 대기

        self._set_status("상세주소 입력")

        

        # 선택 클릭 완료 후 2초 대기

        time.sleep(2.0)



        # 바로 상세주소 입력 진행

        self._log("  👉 상세주소 입력 진행")

        if row.detail_address:

            try:

                # 포커스가 기본적으로 잡혀 있으므로 좌표 클릭 없이 바로 mobile: type으로 입력

                self.driver.execute_script("mobile: type", {"text": row.detail_address})

                time.sleep(0.5)

                self._log(f"  ✅ 상세주소 입력 완료 (입력값: '{row.detail_address}')")

            except Exception as e_det:

                self._log(f"  ❌ 상세주소 입력 실패: {e_det}")



        # 상세 주소 입력 다하고 커서 없애기 (헤더 영역 탭 및 키보드 숨김)

        try:

            self._log("  👉 상세주소 입력 후 포커스 해제 (헤더 탭)")

            self.driver.tap([(500, 150)])

            time.sleep(0.5)

            self.driver.hide_keyboard()

            self._log("  ⌨ 키보드 숨기기 실행 완료")

            time.sleep(1.0)

        except Exception:

            pass



        # 스크롤 젤 밑으로 내리기

        self._log("  ⬇ 스크롤 젤 밑으로 이동 (2회)")

        self._scroll_down()

        time.sleep(0.8)

        self._scroll_down()

        time.sleep(0.8)



        # 선택완료 버튼 클릭

        self._log("  👉 선택완료 버튼 클릭 시도")

        if not self._click_confirm_button():

            self._log("  ❌ 선택완료 버튼 매칭 실패 → 주소 등록 작업 실패 처리")

            return False

        time.sleep(2)



        # [단계 19] 전화번호 중간 4자리 입력

        self._set_status("전화번호 입력")

        

        # //android.view.View[@text="통신사 번호"] 클릭후 2초 대기

        carrier_xpath = '//android.view.View[@text="통신사 번호"]'

        if ah.element_exists(self.driver, carrier_xpath, timeout=3):

            self._log("  📞 [전화번호] '통신사 번호' 발견 → 클릭")

            ah.wait_and_click(self.driver, carrier_xpath, timeout=5, log_callback=self._log)

            time.sleep(2)

            

            # 그후 //android.widget.CheckedTextView[@resource-id="android:id/text1" and @text="010"] 존재 하면 클릭 후 2초 대기

            prefix_010_xpaths = [

                '//android.widget.CheckedTextView[@resource-id="android:id/text1" and @text="010"]',

                '//android.widget.CheckedTextView[@text="010"]',

                '//android.widget.TextView[@text="010"]',

                '//*[@text="010"]',

            ]

            clicked_010 = False

            for p_xpath in prefix_010_xpaths:

                if ah.element_exists(self.driver, p_xpath, timeout=2):

                    self._log(f"  📞 [전화번호] '010' 옵션 발견 ({p_xpath[:50]}) → 클릭")

                    if ah.wait_and_click(self.driver, p_xpath, timeout=5, log_callback=self._log):

                        clicked_010 = True

                        time.sleep(2)

                        break

            if not clicked_010:

                self._log("  ⚠ [전화번호] '010' 옵션을 찾을 수 없어 010 선택 단계를 패스합니다.")

                

        phone_mid = row.get_phone_middle()

        phone_last = row.get_phone_last()



        if phone_mid:

            mid_xpaths = [

                '//android.widget.EditText[@hint="휴대전화번호 앞자리"]',

                '//android.widget.EditText[@hint="휴대전화번호 중간자리"]',

                '//android.widget.EditText[@hint="휴대전화번호 앞"]',

                PHONE_MID_XPATH,

            ]

            mid_success = False

            for xpath in mid_xpaths:

                if ah.element_exists(self.driver, xpath, timeout=3):

                    try:

                        # 1. 엘리먼트 탭

                        el = self.driver.find_element(By.XPATH, xpath)

                        self._log(f"  📞 전화번호 중간 4자리 입력 필드 발견: {xpath}")

                        self._click_element_center_coordinates(el)

                        time.sleep(0.5)

                        

                        # 2. mobile: type

                        try:

                            self.driver.execute_script("mobile: type", {"text": phone_mid})

                            time.sleep(0.5)

                        except Exception:

                            pass

                            

                        # 3. 검증 (stale 방지 위해 새로 찾기)

                        el = self.driver.find_element(By.XPATH, xpath)

                        try:

                            curr_val = el.text or ""

                        except Exception:

                            curr_val = ""

                        

                        if phone_mid not in curr_val:

                            self._log("  ⌨ [전화번호 중간] mobile: type 미입력 확인 -> send_keys 시도 (stale 방지)")

                            el = self.driver.find_element(By.XPATH, xpath)

                            el.clear()

                            time.sleep(0.3)

                            el = self.driver.find_element(By.XPATH, xpath)

                            el.send_keys(phone_mid)

                            time.sleep(0.5)

                        mid_success = True

                        break

                    except Exception as e_mid:

                        self._log(f"  ⚠ 중간 번호 입력 시도 실패 ({xpath}): {e_mid}")

            if not mid_success:

                self._log("  ❌ 전화번호 중간 4자리 입력 최종 실패")



        # [단계 20] 전화번호 마지막 4자리 입력, 2초 대기

        time.sleep(2)

        if phone_last:

            last_xpaths = [

                '//android.widget.EditText[@hint="휴대전화번호 뒷자리"]',

                '//android.widget.EditText[@hint="휴대전화번호 뒤"]',

                PHONE_LAST_XPATH,

            ]

            last_success = False

            for xpath in last_xpaths:

                if ah.element_exists(self.driver, xpath, timeout=3):

                    try:

                        # 1. 엘리먼트 탭

                        el = self.driver.find_element(By.XPATH, xpath)

                        self._log(f"  📞 전화번호 마지막 4자리 입력 필드 발견: {xpath}")

                        # 좌표탭 하지 않고 바로 값 입력

                        el.clear()

                        time.sleep(0.3)

                        el = self.driver.find_element(By.XPATH, xpath)

                        el.send_keys(phone_last)

                        time.sleep(0.5)

                        

                        # 키보드 숨기기

                        try:

                            self.driver.hide_keyboard()

                            self._log("  ⌨ 키보드 숨기기 실행 완료")

                            time.sleep(0.5)

                        except Exception:

                            pass

                        

                        last_success = True

                        break

                    except Exception as e_last:

                        self._log(f"  ⚠ 마지막 번호 입력 시도 실패 ({xpath}): {e_last}")

            if not last_success:

                self._log("  ❌ 전화번호 마지막 4자리 입력 최종 실패")



        # [단계 21] 등록 버튼 클릭
        time.sleep(2)
        self._set_status("등록 버튼 클릭")
        if not ah.wait_and_click(self.driver, REGISTER_BTN_XPATH, timeout=10, log_callback=self._log):
            self._log("❌ 등록 버튼 클릭 실패")
            return False

        time.sleep(2.5)

        # 중복 배송지 팝업 ("회원의 배송지 목록에 동일한 주소가 존재해서...") 감지 시 확인 버튼 처리
        duplicate_pop_xpaths = [
            '//*[contains(@text, "동일한 주소가 존재")]',
            '//*[contains(@text, "입력/수정할 수 없습니다")]',
            '//*[contains(@text, "shopping.naver.com")]',
        ]
        confirm_btn_xpaths = [
            '//android.widget.Button[@text="확인"]',
            '//android.widget.Button[@resource-id="android:id/button1"]',
            '//*[@text="확인"]',
        ]
        has_duplicate = False
        for dup_xpath in duplicate_pop_xpaths:
            if ah.element_exists(self.driver, dup_xpath, timeout=2):
                has_duplicate = True
                self._log("  ⚠ [중복 배송지 팝업 감지] '동일한 주소가 존재해서 입력/수정할 수 없습니다.' 팝업 발견 -> '확인' 버튼 클릭")
                for c_xpath in confirm_btn_xpaths:
                    if ah.element_exists(self.driver, c_xpath, timeout=3):
                        ah.wait_and_click(self.driver, c_xpath, timeout=3, log_callback=self._log)
                        self._log("  ✅ 중복 팝업 '확인' 클릭 완료")
                        time.sleep(2)
                        break
                break

        if has_duplicate:
            self._log(f"  ℹ {row.name} 이미 동일한 배송지가 존재함 -> 팝업 닫고 정상 처리 완료로 간주 진행")
            try:
                ah.go_back(self.driver, self._log)
                time.sleep(2)
            except Exception:
                pass
            return True

        self._log(f"✅ {row.name} 배송지 등록 완료")
        time.sleep(2)
        return True



    # ─── 주소 선택 ────────────────────────────────────────────────────────────



    def _select_address_by_zipcode(self, zipcode: str) -> bool:

        """

        [단계 15] 주소 목록에서 우편번호 일치 항목 선택

        우편번호 공백이면 첫 번째 항목 즉시 클릭

        """

        zipcode_clean = ''.join(filter(str.isdigit, str(zipcode))) if zipcode else ""



        # 우편번호가 공백이면 첫 번째 항목 즉시 클릭, 3초 대기

        if not zipcode_clean:

            self._log("ℹ 우편번호 공백 → 첫 번째 주소 즉시 클릭")

            result = self._click_first_list_item()

            time.sleep(3)

            return result



        # 0순위: OCR 인식 시도 (최대 5회 재시도, 각 시도 시 다시 화면 캡처 및 저장)

        self._log(f"🔍 [우편번호 OCR 검색] 시작 - 대상 우편번호: '{zipcode_clean}'")

        ocr_clicked = False

        for attempt in range(5):

            self._log(f"  👉 [우편번호 OCR] {attempt + 1}/5회차 시도 중...")

            ocr_clicked = self._click_zipcode_by_ocr(zipcode_clean)

            if ocr_clicked:

                self._log(f"🎯 [OCR 인식 성공] 우편번호 {zipcode_clean} 클릭 완료")

                time.sleep(3)

                return True

            time.sleep(2.0)  # 목록 로딩 대기용 2초 휴식



        # 1순위: XML/XPath 기반 매칭 및 좌표/요소 클릭

        # ListView 대기 (ADDR_LIST_XPATH = '//android.widget.ListView')

        if not ah.element_exists(self.driver, ADDR_LIST_XPATH, timeout=10):

            self._log("⚠ 주소 목록 ListView 없음 → 첫 번째 항목 클릭")

            result = self._click_first_list_item()

            time.sleep(3)

            return result



        for scroll_count in range(SCROLL_MAX):

            try:

                # full XPATH로 직접 아이템들 찾기 (상대 경로 버그 회피 및 가로폭 필터링)

                items = []

                for items_xpath in [

                    '//android.webkit.WebView[@text="네이버+ 스토어"]/android.view.View/android.view.View/android.widget.ListView/android.view.View',

                    '//android.webkit.WebView/android.view.View/android.view.View/android.widget.ListView/android.view.View',

                    '//android.widget.ListView/android.view.View'

                ]:

                    try:

                        raw_items = self.driver.find_elements(By.XPATH, items_xpath)

                        if raw_items:

                            # 가로 폭이 넓은 항목(실제 배송지 카드)만 필터링 (버튼 등 내부 요소 제외)

                            filtered_items = []

                            for item in raw_items:

                                try:

                                    rect = item.rect

                                    if rect['width'] > 800:

                                        filtered_items.append(item)

                                except Exception:

                                    filtered_items.append(item)

                            

                            if filtered_items:

                                items = filtered_items

                                break

                    except WebDriverException:

                        continue

            except Exception as e:

                self._log(f"  [디버그] 검색 결과 아이템 획득 에러: {e}")

                items = []



            for idx, item in enumerate(items, start=1):

                try:

                    # 이 항목의 모든 TextView 텍스트 수집

                    texts = item.find_elements(By.XPATH, ".//android.widget.TextView")

                    item_texts = [t.text for t in texts if t.text]



                    for t_el in texts:

                        text = t_el.text

                        text_digits = ''.join(filter(str.isdigit, text))

                        if zipcode_clean and text_digits == zipcode_clean:

                            self._log(f"  우편번호 {zipcode} 매칭 → 항목 {idx} 선택")

                            

                            # 1. 텍스트 요소 좌표 클릭 시도 (가장 안전함)

                            if self._click_element_center_coordinates(t_el):

                                self._log("  ✅ 우편번호 텍스트 좌표 탭 완료")

                                time.sleep(3)

                                return True

                                

                            # 2. RadioButton click 시도

                            radio_btns = item.find_elements(By.XPATH, ".//android.widget.RadioButton")

                            if radio_btns:

                                try:

                                    radio_btns[0].click()

                                    self._log("  ✅ RadioButton click() 완료")

                                    time.sleep(3)

                                    return True

                                except Exception:

                                    if self._click_element_center_coordinates(radio_btns[0]):

                                        self._log("  ✅ RadioButton 좌표 탭 완료")

                                        time.sleep(3)

                                        return True



                            # 3. item 자체 click 시도

                            try:

                                item.click()

                                self._log("  ✅ 항목 click() 완료")

                                time.sleep(3)

                                return True

                            except Exception:

                                if self._click_element_center_coordinates(item):

                                    self._log("  ✅ 항목 좌표 탭 완료")

                                    time.sleep(3)

                                    return True

                except WebDriverException:

                    continue



            # 스크롤 다운 후 재탐색

            if scroll_count < SCROLL_MAX - 1:

                try:

                    self._scroll_down()

                    time.sleep(0.8)

                except Exception:

                    break



        self._log(f"⚠ 우편번호 {zipcode} 미발견 → 첫 번째 항목 클릭")

        result = self._click_first_list_item()

        time.sleep(3)

        return result



    def _click_first_list_item(self) -> bool:

        """주소 목록 첫 번째 항목 클릭"""

        # 1. '선택버튼.png' 매칭을 통한 첫 번째 항목 클릭 시도
        try:
            import numpy as np
            import cv2
            import os
            import io
            from PIL import Image

            btn_template_path = os.path.join(os.path.dirname(__file__), "선택버튼.png")
            if os.path.exists(btn_template_path):
                self._log("🔍 [첫 번째 항목 선택버튼 매칭] 선택버튼.png 이미지 매칭 시도")
                # 화면 캡처
                screenshot_png = self._get_screenshot()
                screenshot_arr = np.frombuffer(screenshot_png, np.uint8)
                img = cv2.imdecode(screenshot_arr, cv2.IMREAD_COLOR)

                btn_template_bgr = cv2.imdecode(np.fromfile(btn_template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
                if btn_template_bgr is not None:
                    btn_template_gray = cv2.cvtColor(btn_template_bgr, cv2.COLOR_BGR2GRAY)
                    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

                    s_h, s_w = img_gray.shape
                    t_h, t_w = btn_template_gray.shape

                    candidates = [] # 후보 리스트: (score, x, y, scale)
                    scales = np.linspace(0.4, 3.0, 80)
                    for scale in scales:
                        new_w = int(t_w * scale)
                        new_h = int(t_h * scale)
                        if new_w >= s_w or new_h >= s_h:
                            continue
                        if new_w < 10 or new_h < 5:
                            continue

                        resized = cv2.resize(btn_template_gray, (new_w, new_h), interpolation=cv2.INTER_AREA)
                        res = cv2.matchTemplate(img_gray, resized, cv2.TM_CCOEFF_NORMED)

                        # 0.55 임계값 이상의 매칭점을 수집
                        loc = np.where(res >= 0.55)
                        for pt in zip(*loc[::-1]):
                            btn_cx = pt[0] + new_w // 2
                            btn_cy = pt[1] + new_h // 2

                            # 선택 버튼은 무조건 화면 우측 영역(오른쪽 절반 이상)에 존재
                            # 또한 너무 위쪽(y < 400)이나 너무 아래쪽(y > s_h - 200)에 있는 노이즈는 제외
                            if btn_cx < s_w * 0.5 or btn_cy < 400 or btn_cy > s_h - 200:
                                continue

                            score = res[pt[1], pt[0]]
                            candidates.append((score, btn_cx, btn_cy, scale))

                    # 근접한 매칭 후보들을 Y좌표 50px 범위로 그룹핑하여 중복 제거
                    groups = []
                    for cand in candidates:
                        score, cx, cy, scale = cand
                        added = False
                        for g in groups:
                            if abs(cy - g[0][2]) < 50:
                                g.append(cand)
                                added = True
                                break
                        if not added:
                            groups.append([cand])

                    best_group_candidates = []
                    for g in groups:
                        g.sort(key=lambda x: -x[0]) # 각 그룹 내에서 가장 매칭 점수가 높은 대표 선정
                        best_group_candidates.append(g[0])

                    # y좌표 기준으로 정렬 (가장 위쪽에 있는 '선택' 버튼이 첫 번째 항목임)
                    best_group_candidates.sort(key=lambda x: x[2])

                    if best_group_candidates:
                        best_score, abs_x, abs_y, best_scale = best_group_candidates[0]
                        self._log(f"🎯 [첫 번째 항목 선택버튼 매칭 성공] 최적 좌표: ({abs_x}, {abs_y}), 점수: {best_score:.4f}, 스케일: {best_scale:.2f}x")
                        
                        # 디버그용 매칭 이미지 저장
                        try:
                            debug_img = img.copy()
                            t_w_resized = int(t_w * best_scale)
                            t_h_resized = int(t_h * best_scale)
                            top_left = (abs_x - t_w_resized // 2, abs_y - t_h_resized // 2)
                            bottom_right = (abs_x + t_w_resized // 2, abs_y + t_h_resized // 2)
                            cv2.rectangle(debug_img, top_left, bottom_right, (0, 255, 0), 2)
                            cv2.circle(debug_img, (abs_x, abs_y), 5, (0, 0, 255), -1)
                            debug_path = os.path.join(os.path.dirname(btn_template_path), f"debug_select_match_first_{self.device_id}.png")
                            _, enc_dbg = cv2.imencode('.png', debug_img)
                            enc_dbg.tofile(debug_path)
                            self._log(f"  [디버그] 첫 번째 항목 선택버튼 매칭 이미지 저장 완료: {debug_path}")
                        except Exception as dbg_err:
                            self._log(f"  [디버그] 매칭 이미지 저장 중 에러: {dbg_err}")

                        # 클릭 실행
                        if self._tap_coordinates(abs_x, abs_y):
                            self._log("  ✅ 첫 번째 주소 항목 선택버튼 매칭 클릭 완료")
                            return True
                    else:
                        self._log("  ⚠ [첫 번째 항목 선택버튼 매칭] 화면에서 매칭 조건에 맞는 선택 버튼을 찾지 못함")
            else:
                self._log(f"  ⚠ '선택버튼.png' 파일이 존재하지 않습니다: {btn_template_path}")
        except Exception as e_matching:
            self._log(f"  ⚠ [첫 번째 항목 선택버튼 매칭] 이미지 매칭 처리 중 에러: {e_matching}")

        # 2. 폴백: 기존 XPath 기반 첫 번째 항목 클릭 로직
        self._log("🔄 [첫 번째 항목 클릭 폴백] 기존 XPath/요소 기반 클릭 진행")
        try:

            list_view = self.driver.find_element(By.XPATH, ADDR_LIST_XPATH)

            items = list_view.find_elements(By.XPATH, "./android.view.View")

            if not items:

                items = list_view.find_elements(By.XPATH, ".//android.view.View")

            if items:

                radio_btns = items[0].find_elements(By.XPATH, ".//android.widget.RadioButton")

                if radio_btns:

                    radio_btns[0].click()

                else:

                    items[0].click()

                self._log("  ✅ 첫 번째 주소 항목 클릭 완료")

                return True

        except Exception as e:

            self._log(f"  ⚠ 첫 번째 항목 클릭 실패: {e}")

        return False




    # ─── 주소검색 입력 / 검색버튼 ────────────────────────────────────────────



    def _input_address_search(self, address_text: str) -> bool:

        """

        [단계 13] 주소검색 팝업 입력창에 텍스트 입력 (도로명.png 이미지 인식 후 클릭 후 NATIVE 키보드 입력)

        """

        self._log(f"  📍 [주소검색 입력] 입력할 주소: '{address_text}'")



        tap_coords = None



        # 1순위: 이미지 매칭 (도로명.png)

        import os

        template_path = os.path.join(os.path.dirname(__file__), "도로명.png")

        if os.path.exists(template_path):

            self._log("🔍 [이미지 매칭] 도로명.png 인식 시도")

            coords = self._find_image_coords(template_path, threshold=0.7)

            if coords:

                tap_coords = coords

                self._log(f"  ✅ [이미지 매칭] 도로명.png 발견 → 좌표: {tap_coords}")

        else:

            self._log("  ⚠ [이미지 매칭] 도로명.png 파일이 존재하지 않아 다음 순위 시도")







        # 3순위: 이미지 및 OCR 모두 실패 시 폴백 (Native XPath 엘리먼트의 좌표)

        if not tap_coords:

            self._log("  ⚠ [이미지/OCR 실패] Native XPath 기반 좌표 클릭 시도")

            native_xpaths = ADDR_SEARCH_INPUT_XPATHS + [

                '//android.widget.EditText'

            ]

            for xpath in native_xpaths:

                try:

                    if ah.element_exists(self.driver, xpath, timeout=2):

                        el = self.driver.find_element(By.XPATH, xpath)

                        rect = el.rect

                        tap_coords = (rect['x'] + rect['width'] // 2, rect['y'] + rect['height'] // 2)

                        self._log(f"  ✅ [Native XPath] 입력창 발견 (XPath: {xpath}) → 좌표: {tap_coords}")

                        break

                except Exception:

                    continue



        if not tap_coords:

            self._log("  ❌ 주소 입력창 좌표 획득 실패")

            return False



        # 2. 좌표 클릭하여 포커스 부여

        self._log(f"  📌 주소 입력창 좌표 탭: {tap_coords}")

        self._tap_coordinates(tap_coords[0], tap_coords[1])

        time.sleep(1.0)



        # 3. 키보드 방식으로 주소 입력 (NATIVE_APP focused EditText에 입력)

        self._log("  ⌨ [키보드 입력] focused EditText 탐색 및 입력 시작")

        try:

            focused_xpath = '//android.widget.EditText[@focused="true"]'

            if ah.element_exists(self.driver, focused_xpath, timeout=3):

                el = self.driver.find_element(By.XPATH, focused_xpath)

                self._log("  ✅ 포커스된 EditText 발견")

            else:

                el = self.driver.find_element(By.XPATH, '//android.widget.EditText')

                self._log("  ⚠ 포커스된 EditText 없음 → 첫 번째 EditText 사용")

                try:

                    el.click()

                    time.sleep(0.5)

                except Exception as click_err:

                    self._log(f"  ⚠ EditText 클릭 실패 (무시): {click_err}")

            

            # 1차 시도: mobile: type (UiAutomator2 전용, 가장 강력하고 키보드 상태 무관하게 동작)

            try:

                self._log("  ⌨ [키보드 입력] 1차 시도: mobile: type 실행")

                self.driver.execute_script("mobile: type", {"text": address_text})

                time.sleep(0.5)

            except Exception as type_err:

                self._log(f"  ⚠ [키보드 입력] 1차 시도 (mobile: type) 실패: {type_err}")



            # 2차 시도: 입력이 제대로 안 되었을 경우 fallback으로 send_keys 시도

            try:

                current_val = el.text or ""

            except Exception:

                current_val = ""



            placeholders = ["지번/도로명 주소", "지번/도로명을 입력해주세요.", "지번/도로명", "지번, 도로명, 건물명으로 검색하세요"]

            is_empty_or_placeholder = (not current_val) or (current_val in placeholders) or (address_text not in current_val)

            

            if is_empty_or_placeholder:

                self._log(f"  ⌨ [키보드 입력] 2차 시도: send_keys 실행 (현재값: '{current_val}')")

                try:

                    el.clear()

                    time.sleep(0.3)

                    el.send_keys(address_text)

                    time.sleep(0.5)

                except Exception as send_err:

                    self._log(f"  ⚠ [키보드 입력] 2차 시도 (send_keys) 실패: {send_err}")

            

            try:

                final_val = el.text or ""

            except Exception:

                final_val = "확인 불가"

            self._log(f"  ✅ [키보드 입력 완료] 최종 확인된 텍스트: '{final_val}'")

            return True

        except Exception as input_err:

            self._log(f"  ❌ [키보드 입력 실패] {input_err}")



        return False



    def _click_search_button(self) -> bool:

        """

        [단계 14] 주소검색 화면의 검색(돋보기) 버튼 클릭

        XML: Button text="검색", bounds=[900,291][1080,409]

        이미지 인식 -> XPath -> 좌표 탭 순으로 시도

        """

        import os

        template_path = os.path.join(os.path.dirname(__file__), "검색.png")



        # 1순위: 이미지 매칭 (검색.png)

        if os.path.exists(template_path):

            self._log("🔍 [이미지 매칭] 검색.png 인식 시도")

            coords = self._find_image_coords(template_path)

            if coords:

                x, y = coords

                self._tap_coordinates(x, y)

                self._log("  ✅ 검색 버튼 이미지 인식 클릭 완료")

                return True



        # 2순위: XPath 클릭 (주소검색 화면 로딩 후)

        if ah.element_exists(self.driver, SEARCH_BTN_XPATH, timeout=5):

            if ah.wait_and_click(self.driver, SEARCH_BTN_XPATH, timeout=5,

                                 log_callback=self._log):

                self._log("  ✅ 검색 버튼 XPath 클릭 완료")

                return True



        # 3순위: 좌표 탭

        # XML 기준 검색버튼 중앙: x=990, y=350 (1080x2340 기준)

        self._log("  ⚠ 검색 버튼 이미지/XPath 실패 → 좌표 탭 시도")

        try:

            size = self.driver.get_window_size()

            w, h = size['width'], size['height']

            # 검색버튼 중앙 비율: x=990/1080=0.917, y=350/2340=0.150

            tap_x = int(w * 0.917)

            tap_y = int(h * 0.150)

            self._log(f"  📌 검색버튼 좌표 탭: ({tap_x}, {tap_y})")

            self.driver.tap([(tap_x, tap_y)])

            self._log("  ✅ 검색버튼 좌표 탭 완료")

            return True

        except Exception as e:

            self._log(f"  ❌ 검색버튼 좌표 탭 실패: {e}")



        return False



    # ─── 상세주소 / 선택완료 ──────────────────────────────────────────────────



    def _input_detail_address(self, detail: str) -> bool:

        """

        [단계 16] 상세주소 입력 - 포커스가 기본적으로 잡혀 있으므로 좌표 클릭 없이 바로 mobile: type으로 입력

        """

        try:

            self.driver.execute_script("mobile: type", {"text": detail})

            time.sleep(0.5)

            self._log(f"  ✅ 상세주소 입력 완료 (입력값: '{detail}')")

            return True

        except Exception as e_det:

            self._log(f"  ❌ 상세주소 입력 실패: {e_det}")

            return False



    def _click_confirm_button(self) -> bool:

        """

        [단계 18] 선택완료 버튼 인식 및 클릭 (이미지 매칭, OCR 방식 또는 XPath)

        """

        # 선택완료 버튼 탐색 시작 하기전에 스크롤을 젤 밑으로 내려주세요

        self._log("  ⬇ 선택완료 버튼 탐색 전 스크롤 다운 진행 (2회)")

        self._scroll_down()

        time.sleep(1)

        self._scroll_down()

        time.sleep(1)



        # 1순위: 이미지 매칭으로 "선택완료.png" 찾아서 클릭
        # ※ threshold를 0.60으로 낮춤: 실제 버튼과 템플릿 색상 차이(연녹/진녹)로
        #   0.69~0.70 구간에서 매칭 실패하는 케이스 방지

        template_path = os.path.join(os.path.dirname(__file__), "선택완료.png")

        if os.path.exists(template_path):

            self._log("🔍 [이미지 매칭] '선택완료.png' 탐색 시작")

            coords = self._find_image_coords(template_path, threshold=0.60)

            if coords:

                self._log(f"🎯 [이미지 매칭] '선택완료.png' 발견! 좌표: {coords}")

                self.driver.tap([coords])

                return True

            else:

                self._log("  [이미지 매칭] '선택완료.png' 탐색 실패")

        else:

            self._log(f"  ⚠ '선택완료.png' 파일이 존재하지 않습니다: {template_path}")



        # 2순위: OCR로 "선택완료" 또는 "완료" 버튼 찾아서 클릭

        try:

            import ddddocr

            import numpy as np

            import cv2

            

            self._log("🔍 [OCR] '선택완료' 버튼 탐색 시작")

            screenshot_png = self._get_screenshot()

            screenshot_arr = np.frombuffer(screenshot_png, np.uint8)

            img = cv2.imdecode(screenshot_arr, cv2.IMREAD_COLOR)

            

            det = ddddocr.DdddOcr(det=True, show_ad=False)

            bboxes = det.detection(screenshot_png)

            

            if bboxes:

                ocr = ddddocr.DdddOcr(show_ad=False)

                for bbox in bboxes:

                    x1, y1, x2, y2 = bbox

                    cropped = img[y1:y2, x1:x2]

                    _, cropped_png = cv2.imencode('.png', cropped)

                    recognized_text = ocr.classification(cropped_png.tobytes())

                    

                    if "선택완료" in recognized_text or "완료" in recognized_text or "확인" in recognized_text:

                        tap_x = (x1 + x2) // 2

                        tap_y = (y1 + y2) // 2

                        self._log(f"🎯 [OCR] 버튼 '{recognized_text}' 매칭 성공! 좌표: ({tap_x}, {tap_y})")

                        self.driver.tap([(tap_x, tap_y)])

                        return True

        except Exception as e:

            self._log(f"  [OCR] 완료 버튼 탐색 중 오류 (XPath 폴백 진행): {e}")



        # 3순위: XPath 클릭

        confirm_xpaths = [

            '//android.widget.Button[@text="선택완료"]',

            '//android.widget.Button[@text="완료"]',

            '//android.widget.Button[@text="확인"]',

            '//*[@content-desc="선택완료"]',

            '//*[@text="선택완료"]',

            '//*[@text="완료"]',

        ]

        for xpath in confirm_xpaths:

            if ah.element_exists(self.driver, xpath, timeout=3):

                # XPath 엘리먼트를 찾아서, 일반 .click()이 아니라 좌표 탭으로 클릭하면 WebView 이슈 회피 가능!

                el = self.driver.find_element(By.XPATH, xpath)

                if self._click_element_center_coordinates(el):

                    self._log(f"  ✅ 선택완료 버튼 좌표 탭 완료 ({xpath[:50]})")

                    return True

                elif ah.wait_and_click(self.driver, xpath, timeout=5, log_callback=self._log):

                    self._log(f"  ✅ 선택완료 버튼 click() 완료 ({xpath[:50]})")

                    return True

        # 4순위: 최후 폴백 - 화면 하단 우측 고정 좌표 탭
        # 선택완료 버튼은 항상 화면 하단 우측(약 75% x, 93% y 위치)에 고정되어 있음
        # 이미지/OCR/XPath 모두 실패해도 버튼이 화면에 보이는 경우 직접 탭
        try:

            size = self.driver.get_window_size()

            w, h = size['width'], size['height']

            # 선택완료 버튼 중앙 비율 (1080x2340 기준: x≈789, y≈2205)
            # x: 우측 절반의 중앙 → 0.75 * w, y: 하단 내비게이션 바 위 → 0.93 * h

            tap_x = int(w * 0.75)

            tap_y = int(h * 0.930)

            self._log(f"  ⚠ [최후 폴백] 선택완료 고정 좌표 탭 시도: ({tap_x}, {tap_y})")

            self.driver.tap([(tap_x, tap_y)])

            time.sleep(1.5)

            self._log("  ✅ [최후 폴백] 선택완료 고정 좌표 탭 완료")

            return True

        except Exception as fallback_err:

            self._log(f"  ❌ [최후 폴백] 선택완료 좌표 탭 실패: {fallback_err}")

        return False



    def _tap_coordinates(self, x: int, y: int) -> bool:

        """지정 좌표 탭 (W3C Actions 지원 및 driver.tap 폴백)"""

        return ah.tap_by_coords(self.driver, x, y, log_callback=self._log)



    def _get_screenshot(self) -> bytes:

        """화면 꺼짐 자동 복구 + 검은 화면 재시도 스크린샷.

        모든 캡처는 반드시 이 메서드를 통해 호출한다."""

        return ah.get_screenshot_safe(

            self.driver, self.device_id, log_callback=self._log

        )



    # ─── 화면 복귀 / 스크롤 ──────────────────────────────────────────────────



    def _ensure_delivery_mgmt_screen(self) -> bool:

        """배송지 관리 화면인지 확인, 아니면 재진입"""

        if ah.element_exists(self.driver, NEW_ADDRESS_BTN_XPATH, timeout=5):

            return True



        # 스크롤 후 재확인

        self._scroll_to_top()

        time.sleep(1)

        if ah.element_exists(self.driver, NEW_ADDRESS_BTN_XPATH, timeout=3):

            return True



        # 재진입 시도

        self._log("⚠ 배송지 관리 화면 이탈 감지 → 재진입 시도")

        try:

            self._go_main_and_enter_store()

            if self._navigate_to_delivery_mgmt():

                return ah.element_exists(self.driver, NEW_ADDRESS_BTN_XPATH, timeout=5)

        except Exception as e:

            self._log(f"⚠ 재진입 중 오류: {e}")

        return False



    def _scroll_down(self):

        """화면 스크롤 다운"""

        try:

            size = self.driver.get_window_size()

            w, h = size['width'], size['height']

            self.driver.swipe(w // 2, int(h * 0.7), w // 2, int(h * 0.3), 500)

        except Exception:

            pass



    def _scroll_to_top(self, max_scroll: int = 3):

        """화면 스크롤 맨 위로"""

        for _ in range(max_scroll):

            try:

                size = self.driver.get_window_size()

                w, h = size['width'], size['height']

                self.driver.swipe(w // 2, int(h * 0.3), w // 2, int(h * 0.75), 500)

                time.sleep(0.5)

            except Exception:

                break



    # ─── 유틸리티 ─────────────────────────────────────────────────────────────



    def _log(self, message: str):

        full_msg = f"[{self.device_id}] {message}"

        if self._log_cb:

            self._log_cb(self.device_id, message)

        else:

            print(full_msg)

        # 로그 파일로도 함께 기록 저장
        try:
            import datetime
            log_dir = os.path.join(os.path.dirname(__file__), "logs")
            os.makedirs(log_dir, exist_ok=True)
            log_file_path = os.path.join(log_dir, f"naver_worker_{self.device_id}.log")
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(log_file_path, "a", encoding="utf-8") as f_log:
                f_log.write(f"[{now_str}] {full_msg}\n")
        except Exception:
            pass



    def _set_status(self, status: str):

        if self._status_cb:

            self._status_cb(self.device_id, status)



    def _click_element_center_coordinates(self, element) -> bool:

        """엘리먼트의 bounds/location을 구해 중앙 좌표를 탭"""

        try:

            rect = element.rect

            x = rect['x'] + rect['width'] // 2

            y = rect['y'] + rect['height'] // 2

            self._log(f"    [좌표탭] 엘리먼트 중앙: ({x}, {y})")

            self._tap_coordinates(x, y)

            return True

        except Exception as e:

            self._log(f"    ⚠ 엘리먼트 좌표 탭 실패: {e}")

            return False



    def _find_image_coords(self, template_path: str, threshold: float = 0.75) -> Optional[tuple]:

        """지정한 템플릿 이미지가 화면에 있는지 멀티스케일 OpenCV 매칭으로 찾고 중심 좌표 반환"""

        try:

            import cv2

            import numpy as np

            from PIL import Image

            import io

            import os

        except ImportError:

            self._log("  [이미지 매칭] cv2/numpy/PIL 라이브러리 미설치")

            return None



        try:

            # 1. 화면 캡처 (화면 꺼짐 자동 복구)

            screenshot_png = self._get_screenshot()

            screenshot_pil = Image.open(io.BytesIO(screenshot_png))

            screen_bgr = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)

            screen_gray = cv2.cvtColor(screen_bgr, cv2.COLOR_BGR2GRAY)



            screen_h, screen_w = screen_gray.shape

            self._log(f"  [이미지 매칭] 스크린샷 크기: {screen_w}x{screen_h}")



            # 원본 캡처 이미지 무조건 저장 (Unicode 경로 대응)

            ss_dir = os.path.join(os.path.dirname(template_path), "캡쳐")
            os.makedirs(ss_dir, exist_ok=True)
            screenshot_path = os.path.join(ss_dir, f"debug_screenshot_{self.device_id}.png")

            try:

                _, enc_ss = cv2.imencode('.png', screen_bgr)

                enc_ss.tofile(screenshot_path)

                self._log(f"  [디버그] 원본 캡처 이미지 저장 완료: {screenshot_path}")

            except Exception as ss_err:

                self._log(f"  [디버그] 원본 캡처 이미지 저장 실패: {ss_err}")



            # 2. 템플릿 로드 (Unicode 경로 대응)

            if not os.path.exists(template_path):

                self._log(f"  [이미지 매칭] 템플릿 파일이 없음: {template_path}")

                return None

            try:

                template_bgr = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)

            except Exception as read_err:

                self._log(f"  [이미지 매칭] 템플릿 로드 중 예외 발생: {read_err}")

                template_bgr = None



            if template_bgr is None:

                self._log(f"  [이미지 매칭] 템플릿 로드 실패 (imdecode 결과 None): {template_path}")

                return None

            template_gray = cv2.cvtColor(template_bgr, cv2.COLOR_BGR2GRAY)



            t_h, t_w = template_gray.shape

            self._log(f"  [이미지 매칭] 템플릿 크기: {t_w}x{t_h}")



            # 3. 멀티스케일 매칭

            best_score = -1

            best_loc = None

            best_scale = 1.0

            best_tw = t_w

            best_th = t_h



            # 스케일 범위: 0.4x ~ 3.0x (기기 해상도 차이 대응)

            scales = np.linspace(0.4, 3.0, 80)



            for scale in scales:

                new_w = int(t_w * scale)

                new_h = int(t_h * scale)



                # 스크린 크기보다 크면 스킵

                if new_w >= screen_w or new_h >= screen_h:

                    continue

                # 너무 작으면 스킵

                if new_w < 10 or new_h < 5:

                    continue



                # 템플릿 리사이즈

                resized = cv2.resize(template_gray, (new_w, new_h),

                                     interpolation=cv2.INTER_AREA)



                # 매칭

                result = cv2.matchTemplate(screen_gray, resized, cv2.TM_CCOEFF_NORMED)

                _, max_val, _, max_loc = cv2.minMaxLoc(result)



                if max_val > best_score:

                    best_score = max_val

                    best_loc = max_loc

                    best_scale = scale

                    best_tw = new_w

                    best_th = new_h



            self._log(f"  [이미지 매칭] 최고 매칭 점수: {best_score:.4f} (스케일: {best_scale:.2f}x)")



            # 4. 결과 반환

            if best_score >= threshold and best_loc is not None:

                cx = best_loc[0] + best_tw // 2

                cy = best_loc[1] + best_th // 2

                self._log(f"  🎯 [이미지 매칭] 발견! 중심좌표: ({cx}, {cy})")



                # 디버그: 매칭 영역 시각화 저장 (Unicode 경로 대응)

                try:

                    debug_img = screen_bgr.copy()

                    top_left = best_loc

                    bottom_right = (top_left[0] + best_tw, top_left[1] + best_th)

                    cv2.rectangle(debug_img, top_left, bottom_right, (0, 255, 0), 3)

                    cv2.circle(debug_img, (cx, cy), 8, (0, 0, 255), -1)

                    

                    debug_path = os.path.join(os.path.dirname(template_path), f"debug_match_{self.device_id}.png")

                    _, enc_dbg = cv2.imencode('.png', debug_img)

                    enc_dbg.tofile(debug_path)

                    self._log(f"  [디버그] 디버그 이미지 저장: {debug_path}")

                except Exception as dbg_err:

                    self._log(f"  [디버그] 디버그 이미지 저장 중 오류: {dbg_err}")



                return cx, cy

            else:

                self._log(f"  ❌ [이미지 매칭] 매칭 실패 (점수 {best_score:.4f} < 임계값 {threshold})")

                # 실패한 시점의 화면 저장 (Unicode 경로 대응)

                failed_path = os.path.join(os.path.dirname(template_path), f"debug_failed_{self.device_id}.png")

                try:

                    _, enc_fail = cv2.imencode('.png', screen_bgr)

                    enc_fail.tofile(failed_path)

                    self._log(f"  [디버그] 매칭 실패 이미지 저장 완료: {failed_path}")

                except Exception as fail_err:

                    self._log(f"  [디버그] 매칭 실패 이미지 저장 중 오류: {fail_err}")

        except Exception as e:

            self._log(f"  [이미지 매칭] 오류 발생: {e}")



        return None



    def _tap_coordinates(self, x: int, y: int) -> bool:

        """지정 좌표 탭 (W3C Actions 지원 및 driver.tap 폴백)"""

        return ah.tap_by_coords(self.driver, x, y, log_callback=self._log)



    def _start_session_keepalive(self, interval: int = 60):

        """OCR 등 장시간 CPU 작업 중 Appium 세션 타임아웃 방지용 백그라운드 핑 시작"""

        import threading

        self._keepalive_stop = threading.Event()

        def _ping():

            while not self._keepalive_stop.wait(interval):

                try:

                    _ = self.driver.current_activity  # 가벼운 명령으로 세션 유지

                    self._log(f"  [세션 유지] keep-alive 핑 전송 완료")

                except Exception:

                    pass  # 이미 세션이 끊어진 경우 무시

        self._keepalive_thread = threading.Thread(target=_ping, daemon=True)

        self._keepalive_thread.start()



    def _stop_session_keepalive(self):

        """백그라운드 핑 중지"""

        if hasattr(self, '_keepalive_stop') and self._keepalive_stop:

            self._keepalive_stop.set()



    def _click_zipcode_by_ocr(self, zipcode: str) -> bool:

        """EasyOCR/ddddocr을 사용하여 화면에서 우편번호를 찾고 해당 우편번호 라인의 선택 버튼을 클릭"""

        try:

            import numpy as np

            import cv2

            import os

            import io

            import re

            from PIL import Image

        except ImportError:

            self._log("  [OCR] cv2/numpy/PIL 라이브러리 미설치 → OCR 건너뜀")

            return False



        # 1. 화면 캡처 및 저장 (Unicode 경로 대응, 화면 꺼짐 자동 복구)

        try:

            screenshot_png = self._get_screenshot()

            screenshot_arr = np.frombuffer(screenshot_png, np.uint8)

            img = cv2.imdecode(screenshot_arr, cv2.IMREAD_COLOR)

            

            # 디버그 캡처 저장

            ss_dir = os.path.join(os.path.dirname(__file__), "캡쳐")
            os.makedirs(ss_dir, exist_ok=True)
            debug_zip_path = os.path.join(ss_dir, f"debug_zipcode_screenshot_{self.device_id}.png")

            _, enc_zip = cv2.imencode('.png', img)

            enc_zip.tofile(debug_zip_path)

            self._log(f"  [디버그] 주소검색 우편번호 화면 캡처 저장 완료: {debug_zip_path}")

        except Exception as e_cap:

            self._log(f"  ❌ 화면 캡처 실패: {e_cap}")

            return False



        tap_x, tap_y = None, None



        # ── 0순위: '선택버튼.png' 이미지 매칭 선행 체크 ──────────────────────
        # 검색 결과가 1건만 나오는 경우 등, 우편번호 OCR 없이 선택 버튼이 바로 보이는 경우
        # ※ 임계값 0.65 이상 + 스케일 0.5x 이상 매칭만 신뢰
        try:
            btn_template_path_pre = os.path.join(os.path.dirname(__file__), "선택버튼.png")
            if os.path.exists(btn_template_path_pre):
                self._log("  🔍 [0순위] 선택버튼.png 이미지 매칭 선행 체크")
                btn_template_bgr_pre = cv2.imdecode(np.fromfile(btn_template_path_pre, dtype=np.uint8), cv2.IMREAD_COLOR)
                if btn_template_bgr_pre is not None:
                    btn_tmpl_gray = cv2.cvtColor(btn_template_bgr_pre, cv2.COLOR_BGR2GRAY)
                    img_gray_pre = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                    s_h_pre, s_w_pre = img_gray_pre.shape
                    t_h_pre, t_w_pre = btn_tmpl_gray.shape

                    pre_candidates = []
                    # 스케일 0.5x 이상만 검색 (0.5x 미만은 버튼이 너무 작아 오매칭 위험)
                    for scale_pre in np.linspace(0.5, 3.0, 55):
                        nw = int(t_w_pre * scale_pre); nh = int(t_h_pre * scale_pre)
                        if nw >= s_w_pre or nh >= s_h_pre or nw < 10 or nh < 5:
                            continue
                        resized_pre = cv2.resize(btn_tmpl_gray, (nw, nh), interpolation=cv2.INTER_AREA)
                        res_pre = cv2.matchTemplate(img_gray_pre, resized_pre, cv2.TM_CCOEFF_NORMED)
                        # 임계값을 0.65로 상향하여 오매칭 방지
                        loc_pre = np.where(res_pre >= 0.65)
                        for pt_pre in zip(*loc_pre[::-1]):
                            cx_pre = pt_pre[0] + nw // 2
                            cy_pre = pt_pre[1] + nh // 2
                            # 화면 우측(>50%) + 상단 노이즈 제외(y>200)
                            if cx_pre < s_w_pre * 0.5 or cy_pre < 200:
                                continue
                            pre_candidates.append((res_pre[pt_pre[1], pt_pre[0]], cx_pre, cy_pre, scale_pre))

                    if pre_candidates:
                        # 점수 내림차순으로 정렬 → 가장 확실한 매칭 선택
                        pre_candidates.sort(key=lambda x: -x[0])
                        best_pre = pre_candidates[0]
                        best_pre_score, pre_x, pre_y, pre_scale = best_pre
                        self._log(f"  🎯 [0순위 선택버튼 매칭 성공] 좌표: ({pre_x}, {pre_y}), 점수: {best_pre_score:.4f}, 스케일: {pre_scale:.2f}x → 바로 클릭")
                        self._tap_coordinates(pre_x, pre_y)
                        return True
                    else:
                        self._log("  ⚠ [0순위] 임계값(0.65) 이상 선택버튼 미발견 → OCR 탐색 진행")
        except Exception as e_pre:
            self._log(f"  ⚠ [0순위] 선택버튼 선행 매칭 에러: {e_pre}")




        # ── 1순위: EasyOCR 사용 ───────────────────────

        easyocr_available = False

        try:

            import easyocr

            easyocr_available = True

        except ImportError:

            self._log("  [EasyOCR] easyocr 라이브러리 미설치 → ddddocr 폴백 사용")



        if easyocr_available:

            try:

                if not hasattr(self, '_easyocr_reader') or self._easyocr_reader is None:

                    self._log("  [EasyOCR] reader 초기화 중 (최초 1회, GPU=False)...")

                    # ── 초기화 전 keep-alive 시작 (초기화에 수 분 소요될 수 있음) ──

                    self._start_session_keepalive(interval=60)

                    try:

                        self._easyocr_reader = easyocr.Reader(['ko', 'en'], gpu=False)

                    finally:

                        self._stop_session_keepalive()

                

                self._log(f"  [EasyOCR] 우편번호 '{zipcode}' 탐색 시작")

                

                # 전처리 (업스케일 및 샤프닝)

                gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

                scale = 2.0

                enlarged = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)

                kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])

                sharpened = cv2.filter2D(enlarged, -1, kernel)

                

                results = self._easyocr_reader.readtext(sharpened, detail=1, paragraph=False)

                

                target_clean = zipcode.strip()

                found_box = None

                found_text = None

                best_conf = 0

                

                for (bbox, text, conf) in results:

                    text_clean = re.sub(r'\s+', '', text)

                    # 완전 일치 또는 포함

                    if text_clean == target_clean and conf > best_conf:

                        found_box = bbox

                        found_text = text

                        best_conf = conf

                    elif target_clean in text_clean and conf > best_conf:

                        found_box = bbox

                        found_text = text

                        best_conf = conf

                

                if found_box is not None:

                    xs = [pt[0] for pt in found_box]

                    ys = [pt[1] for pt in found_box]

                    tap_x = int((min(xs) + max(xs)) / 2 / scale)

                    tap_y = int((min(ys) + max(ys)) / 2 / scale)

                    self._log(f"🎯 [EasyOCR] 우편번호 '{found_text}' 발견! 신뢰도: {best_conf:.2f}  좌표: ({tap_x}, {tap_y})")

                    

                    # 디버그 이미지 저장 (debug_ocr.png)

                    try:

                        debug_img = img.copy()

                        pts = np.array([[int(pt[0]/scale), int(pt[1]/scale)] for pt in found_box], dtype=np.int32)

                        cv2.polylines(debug_img, [pts], True, (0, 255, 0), 3)

                        cv2.circle(debug_img, (tap_x, tap_y), 10, (0, 0, 255), -1)

                        cv2.imwrite(f"debug_ocr_{self.device_id}.png", debug_img)

                    except Exception:

                        pass

            except Exception as e_easy:

                self._log(f"  ⚠ [EasyOCR] 탐색 중 에러: {e_easy} → ddddocr 폴백 시도")



        # ── 2순위: ddddocr 폴백 (세로 슬라이딩 스캔 방식) ──────────

        if tap_y is None:

            try:

                import ddddocr

                self._log("  [ddddocr] 좌측 우편번호 열 세로 슬라이딩 탐색 시작")



                ocr = ddddocr.DdddOcr(show_ad=False)

                h, w = img.shape[:2]

                

                # y=300부터 y=1800까지 15픽셀 간격으로 세로 슬라이드 (범위 확장)

                # 가로 영역은 0~400 픽셀로 확장하여 주소 블록 내 우편번호도 포함

                for y in range(300, min(1800, img.shape[0] - 60), 15):

                    chunk = img[y:y+80, 0:400]

                    _, chunk_png = cv2.imencode('.png', chunk)

                    try:

                        text = ocr.classification(chunk_png.tobytes())

                    except Exception:

                        continue

                        

                    clean_text = "".join(filter(str.isdigit, text))

                    

                    if zipcode in clean_text:

                        tap_x = 130  # 우편번호 열의 중앙 근처

                        tap_y = y + 50

                        self._log(f"🎯 [ddddocr] 우편번호 '{zipcode}' 발견! (인식텍스트: '{text}', Y좌표: {tap_y})")

                        break

                        

                if tap_y is None:

                    self._log(f"  [ddddocr] 우편번호 '{zipcode}' 세로 슬라이드 탐색 실패")



            except Exception as e_dddd:

                self._log(f"  ⚠ [ddddocr] 탐색 중 에러: {e_dddd}")





        # 우편번호 좌표를 최종적으로 찾지 못했으면 실패 리턴

        if tap_y is None:

            self._log(f"  ❌ 화면에서 우편번호 '{zipcode}'를 찾지 못했습니다.")

            return False



        # ── 3. 우편번호 라인의 선택 버튼 찾기 ────────────────

        # y축 범위 설정: 우편번호 Y 중심(tap_y)보다 작고(화면 위쪽), 거리가 150 이내에 있는 버튼을 찾아야 하므로,

        # 넉넉하게 tap_y - 180 에서 tap_y + 40 범위로 세로 스트립을 자릅니다.

        strip_y1 = max(0, tap_y - 180)

        strip_y2 = min(img.shape[0], tap_y + 40)

        

        # 가로 스트립 이미지 추출

        strip_img = img[strip_y1:strip_y2, 0:img.shape[1]]

        

        btn_template_path = os.path.join(os.path.dirname(__file__), "선택버튼.png")

        

        if os.path.exists(btn_template_path):

            self._log(f"🔍 [선택버튼 이미지 매칭] 선택버튼.png 매칭 시도 (스트립 높이: {strip_y1}~{strip_y2})")

            try:

                # 템플릿 로드 (Unicode 경로 대응)

                btn_template_bgr = cv2.imdecode(np.fromfile(btn_template_path, dtype=np.uint8), cv2.IMREAD_COLOR)

                if btn_template_bgr is not None:

                    btn_template_gray = cv2.cvtColor(btn_template_bgr, cv2.COLOR_BGR2GRAY)

                    strip_gray = cv2.cvtColor(strip_img, cv2.COLOR_BGR2GRAY)

                    

                    s_h, s_w = strip_gray.shape

                    t_h, t_w = btn_template_gray.shape

                    

                    candidates = [] # 조건을 만족하는 매칭 후보 저장: (score, abs_x, abs_y, scale, y_diff)

                    

                    # 스케일 범위: 0.4x ~ 3.0x (80단계)

                    scales = np.linspace(0.4, 3.0, 80)

                    for scale in scales:

                        new_w = int(t_w * scale)

                        new_h = int(t_h * scale)

                        

                        if new_w >= s_w or new_h >= s_h:

                            continue

                        if new_w < 10 or new_h < 5:

                            continue

                            

                        resized = cv2.resize(btn_template_gray, (new_w, new_h), interpolation=cv2.INTER_AREA)

                        res = cv2.matchTemplate(strip_gray, resized, cv2.TM_CCOEFF_NORMED)

                        

                        # 0.55 임계값 이상의 모든 매칭점을 수집

                        loc = np.where(res >= 0.55)

                        for pt in zip(*loc[::-1]):

                            btn_cx = pt[0] + new_w // 2

                            btn_cy = pt[1] + new_h // 2

                            

                            abs_x = btn_cx

                            abs_y = strip_y1 + btn_cy

                            

                            # 선택 버튼은 무조건 화면 우측 영역(오른쪽 절반 이상)에 존재

                            if abs_x < img.shape[1] * 0.5:

                                continue

                                

                            y_diff = tap_y - abs_y

                            # 조건 검사: 

                            # 1. button_y < tap_y (y_diff > 0)

                            # 2. tap_y - button_y <= 150 (y_diff <= 150)

                            if 0 < y_diff <= 150:

                                score = res[pt[1], pt[0]]

                                candidates.append((score, abs_x, abs_y, scale, y_diff))

                    

                    # 후보가 있으면 필터링 및 최적 후보 선택

                    # 정렬 기준: 점수(score) 내림차순, y_diff 오름차순(제일 가까운 거리)

                    if candidates:

                        candidates.sort(key=lambda x: (-x[0], x[4]))

                        best_candidate = candidates[0]

                        best_score, abs_x, abs_y, best_scale, y_diff = best_candidate

                        

                        self._log(f"🎯 [선택버튼 매칭 성공] 최적 후보 선택 - 절대좌표: ({abs_x}, {abs_y}), 점수: {best_score:.4f}, 거리차이: {y_diff}px, 스케일: {best_scale:.2f}x")

                        

                        # 디버그용 매칭 결과 이미지 저장 (Unicode 경로 대응)

                        try:

                            debug_img = strip_img.copy()

                            btn_cx_strip = abs_x

                            btn_cy_strip = abs_y - strip_y1

                            t_w_resized = int(t_w * best_scale)

                            t_h_resized = int(t_h * best_scale)

                            

                            top_left = (btn_cx_strip - t_w_resized // 2, btn_cy_strip - t_h_resized // 2)

                            bottom_right = (btn_cx_strip + t_w_resized // 2, btn_cy_strip + t_h_resized // 2)

                            

                            cv2.rectangle(debug_img, top_left, bottom_right, (0, 255, 0), 2)

                            cv2.circle(debug_img, (btn_cx_strip, btn_cy_strip), 5, (0, 0, 255), -1)

                            

                            # 우편번호 Y선 표시

                            cv2.line(debug_img, (0, tap_y - strip_y1), (img.shape[1], tap_y - strip_y1), (255, 0, 0), 1)

                            

                            debug_path = os.path.join(os.path.dirname(btn_template_path), f"debug_select_match_{self.device_id}.png")

                            _, enc_dbg = cv2.imencode('.png', debug_img)

                            enc_dbg.tofile(debug_path)

                            self._log(f"  [디버그] 선택버튼 매칭 디버그 이미지 저장 완료: {debug_path}")

                        except Exception as dbg_err:

                            self._log(f"  [디버그] 선택버튼 매칭 디버그 이미지 저장 중 오류: {dbg_err}")

                            

                        self._tap_coordinates(abs_x, abs_y)

                        return True

                    else:

                        self._log("  ❌ [선택버튼 이미지 매칭] 조건(y < tap_y 이고 차이 <= 150)을 만족하는 매칭 후보 없음")

                else:

                    self._log("  ❌ [선택버튼 이미지 매칭] 템플릿 디코딩 실패")

            except Exception as e_match:

                self._log(f"  ⚠ [선택버튼 이미지 매칭] 에러: {e_match}")

        else:

            self._log(f"  ⚠ '선택버튼.png' 파일이 존재하지 않습니다: {btn_template_path}")

 

        # 폴백 1: OCR로 가로 스트립 영역 내에서 '선택' 글자 찾기

        self._log("  🔍 [폴백 1] 스트립 영역 내에서 OCR로 '선택' 글자 탐색")

        try:

            _, strip_png_bytes = cv2.imencode('.png', strip_img)

            import ddddocr

            strip_det = ddddocr.DdddOcr(det=True, show_ad=False)

            strip_bboxes = strip_det.detection(strip_png_bytes.tobytes())

            if strip_bboxes:

                strip_ocr = ddddocr.DdddOcr(show_ad=False)

                fallback_candidates = []

                for s_bbox in strip_bboxes:

                    sx1, sy1, sx2, sy2 = s_bbox

                    sc_y1 = max(0, sy1 - 2)

                    sc_y2 = min(strip_img.shape[0], sy2 + 2)

                    sc_x1 = max(0, sx1 - 5)

                    sc_x2 = min(strip_img.shape[1], sx2 + 5)

                    s_cropped = strip_img[sc_y1:sc_y2, sc_x1:sc_x2]

                    _, s_cropped_png = cv2.imencode('.png', s_cropped)

                    s_text = strip_ocr.classification(s_cropped_png.tobytes())

                    if "선택" in s_text or "선" in s_text or "택" in s_text:

                        s_cx = (sx1 + sx2) // 2

                        s_cy = (sy1 + sy2) // 2

                        abs_x = s_cx

                        abs_y = strip_y1 + s_cy

                        

                        y_diff = tap_y - abs_y

                        if 0 < y_diff <= 150:

                            fallback_candidates.append((abs_x, abs_y, y_diff))

                

                if fallback_candidates:

                    # y_diff가 가장 작은 순서로 정렬

                    fallback_candidates.sort(key=lambda x: x[2])

                    abs_x, abs_y, y_diff = fallback_candidates[0]

                    self._log(f"  🎯 [폴백 1 성공] OCR로 '선택' 발견 → 절대좌표: ({abs_x}, {abs_y}), 거리: {y_diff}px")

                    self._tap_coordinates(abs_x, abs_y)

                    return True

        except Exception as fallback_ocr_err:

            self._log(f"  ⚠ [폴백 1] OCR 탐색 실패: {fallback_ocr_err}")

 

        # 폴백 2: 우측 영역 강제 클릭

        # 선택 버튼이 우편번호 y축보다 작고 차이가 150 이내에 있는 제일 가까운 곳이므로,

        # 기본 탭 또한 약간 위(tap_y - 90)를 탭하도록 보정합니다.

        fallback_x = int(img.shape[1] * 0.90)

        fallback_y = tap_y - 90

        self._log(f"  🔍 [폴백 2] 기본 오른쪽 좌표 클릭 시도: ({fallback_x}, {fallback_y}) (우편번호 대비 -90px)")

        self._tap_coordinates(fallback_x, fallback_y)

        return True
