# -*- coding: utf-8 -*-
"""
naver_pay_flow.py

[Pay 방식] 네이버페이 설정 → 배송지 관리 경로로 배송지 등록 (pay주소지 등록.md 기준)

 1. 메인화면 네이버페이 버튼 클릭 (content-desc 에 '네이버페이' 포함, 텍스트 변동형)
 2. pay화면 우측 상단 햄버거(전체메뉴) 클릭
 3. pay설정 화면 '설정' 클릭
 4. 페이 설정 화면 '배송지 관리, 링크' 클릭
 5~6. 기본배송지 제외 주소 '삭제' → 빨간색 '삭제하기' 확인 (전부 삭제될 때까지 반복)
 7. '배송지 추가' 클릭
 8. 배송지 등록/수정 화면(배송지 등록/수정.xml) 입력 + '기본 배송지로 설정' 체크
 9. '저장' 클릭
 엑셀 주소록 행을 폰ID 별로 반복 (반복/상태기록은 NaverWorker._register_address_loop 가 담당)

NaverWorker 가 이 Mixin 을 상속하여 사용합니다.
"""

import os
import re
import time
import subprocess
import xml.etree.ElementTree as ET
from typing import Optional

from selenium.webdriver.common.by import By

try:
    import naver_appium as ah  # type: ignore
except ImportError:  # pragma: no cover
    import appium_helper as ah  # type: ignore


# ─── Pay 방식 XPath 상수 ────────────────────────────────────────────────────

# [1] 메인화면 - 네이버페이 버튼 (content-desc 변동: "네이버페이, 새 항목 2개 있음, 버튼" 등)
PAY_MAIN_BTN_XPATHS = [
    '//android.widget.ImageView[contains(@content-desc, "네이버페이")]',
    '//*[@resource-id="com.nhn.android.search:id/payImageView"]',
    '//*[contains(@content-desc, "네이버페이") and contains(@content-desc, "버튼")]',
]

# [2] pay화면 - 햄버거(전체메뉴) 버튼
PAY_MENU_BTN_XPATHS = [
    '//android.view.View[@content-desc="전체메뉴"]',
    '//*[@content-desc="전체메뉴"]',
    '//*[contains(@content-desc, "전체메뉴")]',
    '//android.widget.Button[@text="전체메뉴"]',
]

# [3] pay 전체메뉴 - 설정
PAY_SETTING_XPATHS = [
    '//android.view.View[@content-desc="설정"]',
    '//*[@content-desc="설정"]',
    '//android.widget.Button[@text="설정"]',
]

# [4] 페이 설정 - 배송지 관리, 링크
PAY_DELIVERY_MGMT_XPATHS = [
    '//android.view.View[@content-desc="배송지 관리, 링크"]',
    '//*[contains(@content-desc, "배송지 관리")]',
    '//*[contains(@text, "배송지 관리") and not(contains(@text, "배송지 관리에서"))]',
]

# [5] 배송지 관리 화면 - 삭제 버튼 / 기본배송지 라벨
PAY_DELETE_BTN_XPATH = (
    '//android.widget.Button[@text="삭제"]'
    ' | //android.view.View[@content-desc="삭제"]'
    ' | //android.widget.Button[@content-desc="삭제"]'
)
PAY_DEFAULT_LABEL_XPATH = (
    '//*[@text="기본배송지" or @content-desc="기본배송지"'
    ' or (contains(@text, "기본배송지") and not(contains(@text, "설정")))]'
)

# [5] 삭제 확인 팝업 - 빨간색 '삭제하기'
PAY_DELETE_CONFIRM_XPATHS = [
    '//android.widget.Button[@text="삭제하기"]',
    '//*[@text="삭제하기"]',
    '//*[@content-desc="삭제하기"]',
    '//android.widget.Button[@resource-id="android:id/button1"]',
]

# [7] 배송지 추가
PAY_ADD_ADDRESS_XPATHS = [
    '//android.widget.Button[@text="배송지 추가"]',
    '//*[@text="배송지 추가"]',
    '//*[@content-desc="배송지 추가"]',
    '//*[contains(@text, "배송지 추가")]',
]

# [8] 배송지 등록/수정 화면 (배송지 등록/수정.xml)
PAY_FORM_TITLE_XPATH      = '//*[@text="배송지 등록/수정"]'
PAY_RECEIVER_XPATHS       = [
    '//android.widget.EditText[@hint="이름"]',
    '//*[@text="수령인"]/following::android.widget.EditText[1]',
]
PAY_ADDR_SEARCH_BTN_XPATHS = [
    '//android.view.View[@content-desc="주소검색"]',
    '//*[@content-desc="주소검색"]',
    '//android.widget.TextView[@text="주소검색"]',
    '//android.widget.Button[@text="주소검색"]',
]
PAY_ZIPCODE_FIELD_XPATH   = '//android.widget.EditText[@hint="우편번호"]'
PAY_ADDRESS_FIELD_XPATH   = '//android.widget.EditText[@hint="주소"]'
PAY_DETAIL_ADDR_XPATHS    = [
    '//android.widget.EditText[@resource-id="detailed-address"]',
    '//android.widget.EditText[@hint="상세주소 입력"]',
]

# [8-2-1] 상세주소입력창.xml ('상세주소를 알려주세요' 팝업/화면)
PAY_POPUP_DETAIL_INPUT_XPATHS = [
    '//android.widget.EditText[@resource-id="address-detail"]',
    '//android.widget.EditText[contains(@resource-id, "address-detail")]',
    '//*[contains(@text, "상세주소를 알려주세요")]/ancestor::*//android.widget.EditText',
    '//android.widget.EditText[@hint="주소" and @resource-id="address-detail"]',
]
PAY_POPUP_DETAIL_CONFIRM_XPATHS = [
    '//*[contains(@text, "상세주소")]/ancestor::*//android.widget.Button[@text="확인"]',
    '//android.widget.Button[@text="확인"]',
    '//*[@text="확인" and @clickable="true"]',
    '//android.view.View/android.widget.Button[@text="확인"]',
]

PAY_PHONE_FIRST_XPATH     = '//*[@resource-id="first-number{n}"]'
PAY_PHONE_MID_XPATH       = '//android.widget.EditText[@resource-id="second-number{n}"]'
PAY_PHONE_LAST_XPATH      = '//android.widget.EditText[@resource-id="last-number{n}"]'
PAY_DEFAULT_CHECK_XPATHS  = [
    '//android.widget.CheckBox[@resource-id="default-address"]',
    '//android.widget.CheckBox[@text="기본 배송지로 설정"]',
]
PAY_DEFAULT_TEXT_XPATH    = '//android.widget.TextView[@text="기본 배송지로 설정"]'

# [9] 저장
PAY_SAVE_BTN_XPATHS = [
    '//android.widget.Button[@text="저장"]',
    '//*[@text="저장" and @clickable="true"]',
]


class PayAddressMixin:
    """네이버페이 경로 배송지 등록 기능 (NaverWorker 에 Mixin 으로 결합)"""

    # ─── 공통 유틸 ──────────────────────────────────────────────────────────

    def _is_pay_mode(self) -> bool:
        return str(getattr(self, "registration_mode", "")).strip().lower() in ("pay", "naverpay", "npay")

    def _pay_click_first(self, xpaths, label: str, timeout: float = 3, wait_after: float = 2.5) -> bool:
        """xpaths 중 처음 발견되는 요소 클릭 (click 실패 시 좌표 탭 폴백)"""
        for xp in xpaths:
            if not ah.element_exists(self.driver, xp, timeout=timeout):
                continue
            self._log(f"📌 [Pay] '{label}' 발견 ({xp[:60]}) → 클릭")
            if ah.wait_and_click(self.driver, xp, timeout=3, log_callback=self._log):
                time.sleep(wait_after)
                return True
            try:
                el = self.driver.find_element(By.XPATH, xp)
                if self._click_element_center_coordinates(el):
                    time.sleep(wait_after)
                    return True
            except Exception:
                pass
        return False

    def _pay_adb_tap(self, x: int, y: int):
        try:
            subprocess.run(
                ["adb", "-s", self.device_id, "shell", "input", "tap", str(int(x)), str(int(y))],
                capture_output=True, timeout=5,
            )
        except Exception as e:
            self._log(f"  ⚠ [Pay] ADB tap 실패: {e}")

    def _pay_input_field(self, xpath: str, value: str, label: str) -> bool:
        """EditText 에 값 입력 (기존값 clear 후 입력, 입력값 검증)"""
        if not value:
            self._log(f"  ⚠ [Pay] {label} 값이 비어 있음 → 건너뜀")
            return False
        el = ah.wait_for_element(self.driver, xpath, timeout=4)
        if el is None:
            self._log(f"  ⚠ [Pay] {label} 입력창 미발견: {xpath}")
            return False
        try:
            try:
                el.click()
            except Exception:
                self._click_element_center_coordinates(el)
            time.sleep(0.3)
            try:
                el.clear()
            except Exception:
                pass
            try:
                el.send_keys(value)
            except Exception:
                self.driver.execute_script("mobile: type", {"text": value})
            time.sleep(0.4)
            try:
                el = self.driver.find_element(By.XPATH, xpath)
                cur = (el.text or el.get_attribute("text") or "").strip()
            except Exception:
                cur = ""
            if value not in cur:
                self._log(f"  ⌨ [Pay] {label} 값 미확인(현재: '{cur}') → mobile: type 재입력")
                try:
                    el = self.driver.find_element(By.XPATH, xpath)
                    el.click()
                    el.clear()
                except Exception:
                    pass
                self.driver.execute_script("mobile: type", {"text": value})
                time.sleep(0.4)
            self._log(f"  ✅ [Pay] {label} 입력 완료: '{value}'")
            return True
        except Exception as e:
            self._log(f"  ❌ [Pay] {label} 입력 실패: {str(e).splitlines()[0]}")
            return False

    def _pay_hide_keyboard(self):
        try:
            self.driver.hide_keyboard()
            time.sleep(0.4)
        except Exception:
            pass

    # ─── [1~4] 배송지 관리 화면 진입 ────────────────────────────────────────

    def _pay_go_main(self, login_id: str = "") -> bool:
        """메인 페이지 이동 + 웰컴 모달 처리 + (필요 시) 계정 전환"""
        self._set_status("[Pay] 메인 페이지 이동")
        ah.go_to_main_page(self.driver, self._log)
        time.sleep(4)
        self._check_and_close_welcome_modals(step_label="Pay-0")
        if login_id:
            if not self._switch_account(login_id):
                self._log(f"❌ [Pay] 계정 전환 실패 (아이디: {login_id})")
                return False
            ah.go_to_main_page(self.driver, self._log)
            time.sleep(3)
            self._check_and_close_welcome_modals(step_label="Pay-0")
        return True

    def _navigate_to_pay_delivery_mgmt(self, row=None) -> bool:
        """메인 → 네이버페이 → 전체메뉴 → 설정 → 배송지 관리"""
        self._log("💳 [Pay 방식] 네이버페이 → 전체메뉴 → 설정 → 배송지 관리 경로 진입 시작")

        # [1] 네이버페이 버튼
        self._set_status("[Pay] 네이버페이 클릭")
        clicked = False
        for attempt in range(1, 4):
            if self._pay_click_first(PAY_MAIN_BTN_XPATHS, "네이버페이", timeout=4, wait_after=4):
                clicked = True
                break
            self._log(f"  ⚠ [Pay] 네이버페이 버튼 미발견 ({attempt}/3) → 메인 복귀 후 재시도")
            ah.go_to_main_page(self.driver, self._log)
            time.sleep(3)
            self._check_and_close_welcome_modals(step_label="Pay-1")
        if not clicked:
            self._log("❌ [Pay] 네이버페이 버튼을 찾지 못했습니다.")
            return False
        self._wait_webview_ready(label="Pay 홈", timeout=15, webview_text="네이버페이")
        self._dismiss_hide_popup(max_count=2)

        # [2] 햄버거(전체메뉴)
        self._set_status("[Pay] 전체메뉴 클릭")
        if not self._pay_click_first(PAY_MENU_BTN_XPATHS, "전체메뉴(햄버거)", timeout=5, wait_after=3):
            # 폴백: image.png 기준 우측 상단 햄버거 위치 (1080 기준 x≈997, y≈136)
            try:
                size = self.driver.get_window_size()
                x, y = int(size["width"] * 0.923), int(size["height"] * 0.061)
                self._log(f"  ⚠ [Pay] 전체메뉴 XPath 미발견 → 우측 상단 좌표 탭 ({x}, {y})")
                self._pay_adb_tap(x, y)
                time.sleep(3)
            except Exception:
                pass

        # [3] 설정
        self._set_status("[Pay] 설정 클릭")
        if not self._pay_click_first(PAY_SETTING_XPATHS, "설정", timeout=5, wait_after=3):
            self._log("❌ [Pay] 전체메뉴의 '설정' 버튼을 찾지 못했습니다.")
            return False

        # [4] 배송지 관리, 링크
        self._set_status("[Pay] 배송지 관리 클릭")
        mgmt_ok = False
        for attempt in range(1, 4):
            if self._pay_click_first(PAY_DELIVERY_MGMT_XPATHS, "배송지 관리", timeout=4, wait_after=3):
                mgmt_ok = True
                break
            self._log(f"  ⚠ [Pay] '배송지 관리' 미발견 ({attempt}/3) → 스크롤 후 재탐색")
            self._scroll_down()
            time.sleep(1.2)
        if not mgmt_ok:
            self._log("❌ [Pay] '배송지 관리, 링크' 를 찾지 못했습니다.")
            return False

        # 본인인증(생년월일) 요구 시 처리 (기존 로직 재사용)
        self._check_and_process_birthday_auth(row=row)
        self._dismiss_hide_popup(max_count=1)

        if self._pay_find_add_button(timeout=10):
            self._log("✅ [Pay] 배송지 관리 화면 진입 완료")
            return True
        self._log("⚠ [Pay] '배송지 추가' 버튼 미확인 → 그래도 진행")
        return True

    def _pay_find_add_button(self, timeout: float = 5) -> Optional[str]:
        for i, xp in enumerate(PAY_ADD_ADDRESS_XPATHS):
            if ah.element_exists(self.driver, xp, timeout=timeout if i == 0 else 1):
                return xp
        return None

    def _ensure_pay_delivery_mgmt_screen(self) -> bool:
        """Pay 배송지 관리 화면인지 확인 → 아니면 재진입"""
        if self._pay_find_add_button(timeout=4):
            return True
        self._scroll_to_top()
        time.sleep(1)
        if self._pay_find_add_button(timeout=2):
            return True
        self._log("⚠ [Pay] 배송지 관리 화면 이탈 감지 → 재진입 시도")
        try:
            row = getattr(self, "current_row", None)
            if self._pay_go_main(login_id=""):
                if self._navigate_to_pay_delivery_mgmt(row=row):
                    return bool(self._pay_find_add_button(timeout=5))
        except Exception as e:
            self._log(f"⚠ [Pay] 재진입 중 오류: {e}")
            if self._is_driver_session_error(e) or self._is_connection_refused(e):
                raise
        return False

    # ─── [5~6] 기본배송지 제외 삭제 ─────────────────────────────────────────

    def _pay_parse_delete_targets_from_xml(self) -> Optional[list]:
        """
        driver.page_source XML 트리를 정밀 분석하여
        기본배송지 카드에 속하지 않는 모든 '삭제' 버튼의 (cx, cy) 좌표 목록 반환.
        반환값:
            - [(cx, cy), ...]: 비기본 배송지의 삭제 버튼 좌표 목록 (y 오름차순)
            - []: 화면에 배송지 카드가 있지만 모두 기본배송지임 (삭제 완료 상태)
            - None: XML 파싱 불가로 판별 실패 (Selenium 폴백 필요)
        """
        try:
            source = self.driver.page_source
            if not source:
                return None
            root = ET.fromstring(source)
        except Exception as e:
            self._log(f"  ⚠ [Pay] page_source XML 파싱 예외: {e}")
            return None

        def parse_bounds(b_str):
            m = re.match(r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]", b_str or "")
            if m:
                return int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
            return None

        # 1. '기본배송지' 텍스트를 가진 모든 요소 및 주변 y 범위 수집
        default_y_ranges = []
        for node in root.iter():
            txt = (node.attrib.get("text") or "") + " " + (node.attrib.get("content-desc") or "")
            if "기본배송지" in txt or "기본 배송지" in txt:
                b = parse_bounds(node.attrib.get("bounds"))
                if b:
                    # 기본배송지 라벨 기준 위로 250px (삭제버튼은 라벨 위 45~100px), 아래로 350px (주소/연락처)
                    default_y_ranges.append((b[1] - 250, b[3] + 350))

        # 2. is-collection-item="true" 인 카드 단위 정밀 분석
        non_default_btns = []
        cards_found = False

        for node in root.iter():
            if node.attrib.get("is-collection-item") == "true":
                del_btns = [
                    e for e in node.iter()
                    if (e.attrib.get("text") == "삭제" or e.attrib.get("content-desc") == "삭제")
                ]
                if not del_btns:
                    continue
                cards_found = True

                texts = [
                    (e.attrib.get("text") or "") + " " + (e.attrib.get("content-desc") or "")
                    for e in node.iter()
                ]
                card_text = " ".join(texts)
                card_bounds = parse_bounds(node.attrib.get("bounds"))
                c_y1, c_y2 = (card_bounds[1], card_bounds[3]) if card_bounds else (-1, -1)

                # 카드 내에 '기본배송지' 텍스트가 있으면 해당 카드 전체 삭제 버튼 무조건 제외!
                if "기본배송지" in card_text or "기본 배송지" in card_text:
                    self._log(f"  🛡 [Pay] 기본배송지 카드 감지 (bounds={node.attrib.get('bounds')}) → 삭제 대상에서 제외")
                    if c_y1 != -1:
                        default_y_ranges.append((c_y1 - 10, c_y2 + 10))
                    continue

                for btn in del_btns:
                    b = parse_bounds(btn.attrib.get("bounds"))
                    if not b:
                        continue
                    bx1, by1, bx2, by2 = b
                    cx, cy = (bx1 + bx2) // 2, (by1 + by2) // 2
                    # 화면 가시 영역 내인지 체크 (헤더 240, 하단 바 1850 제외)
                    if 250 <= cy <= 1850:
                        # 추가 2차 방어: 기본배송지 y 범위에 겹치지 않는지 확인
                        if not any(y1 <= cy <= y2 for y1, y2 in default_y_ranges):
                            non_default_btns.append((cy, cx))
                        else:
                            self._log(f"  🛡 [Pay] y={cy} 삭제 버튼이 기본배송지 범위와 겹침 → 제외")

        if cards_found:
            non_default_btns.sort()
            return [(cx, cy) for cy, cx in non_default_btns]

        # 3. is-collection-item 속성이 없는 레이아웃용 폴백
        all_delete_btns = []
        for node in root.iter():
            if node.attrib.get("text") == "삭제" or node.attrib.get("content-desc") == "삭제":
                b = parse_bounds(node.attrib.get("bounds"))
                if not b:
                    continue
                bx1, by1, bx2, by2 = b
                cy = (by1 + by2) // 2
                cx = (bx1 + bx2) // 2
                if not (250 <= cy <= 1850):
                    continue
                if any(y1 <= cy <= y2 for y1, y2 in default_y_ranges):
                    self._log(f"  🛡 [Pay] y={cy} 삭제 버튼은 기본배송지 범위 소속 → 제외")
                    continue
                all_delete_btns.append((cy, cx))

        all_delete_btns.sort()
        return [(cx, cy) for cy, cx in all_delete_btns]

    def _pay_find_non_default_delete_btn(self):
        """기본배송지 카드에 속하지 않는 첫 번째 '삭제' 버튼의 중앙 좌표 (cx, cy) 반환"""
        # [방식 1] XML 기반 정밀 분석 (가장 정확하고 기본배송지 100% 보호)
        targets = self._pay_parse_delete_targets_from_xml()
        if targets is not None:
            if not targets:
                self._log("  ✅ [Pay] 기본배송지 외에 더 이상 삭제할 비기본 배송지가 없습니다.")
                return None
            cx, cy = targets[0]
            self._log(f"  🎯 [Pay] 비기본 배송지 삭제 대상 감지: ({cx}, {cy}), 잔여 {len(targets)}건")
            return cx, cy

        # [방식 2] Selenium find_elements 폴백 (XML 실패 시)
        self._log("  ℹ [Pay] XML 분석 불가 → Selenium 기반 삭제 대상 탐색")
        try:
            btns = self.driver.find_elements(By.XPATH, PAY_DELETE_BTN_XPATH)
        except Exception:
            return None
        if not btns:
            return None

        # 기본배송지 라벨 및 카드 범위 수집
        default_ranges = []
        try:
            for lb in self.driver.find_elements(By.XPATH, PAY_DEFAULT_LABEL_XPATH):
                try:
                    r = lb.rect
                    if r.get("height", 0) > 0:
                        # 라벨 기준 상하 여유 확장 (-250, +350)
                        default_ranges.append((r["y"] - 250, r["y"] + r["height"] + 350))
                    try:
                        cards = lb.find_elements(By.XPATH, "./ancestor::*[@is-collection-item='true']")
                        for c in cards:
                            cr = c.rect
                            if cr.get("height", 0) > 0:
                                default_ranges.append((cr["y"] - 10, cr["y"] + cr["height"] + 10))
                    except Exception:
                        pass
                except Exception:
                    continue
        except Exception:
            pass

        cands = []
        for b in btns:
            try:
                r = b.rect
                if r.get("width", 0) <= 0 or r.get("height", 0) <= 0:
                    continue
                cy = r["y"] + r["height"] // 2
                cx = r["x"] + r["width"] // 2
                if not (250 <= cy <= 1850):
                    continue
                is_default = any(top <= cy <= bottom for top, bottom in default_ranges)
                if is_default:
                    self._log(f"  ⛔ [Pay] y={cy} 삭제 버튼은 기본배송지 소속 → 제외")
                    continue
                cands.append((cy, cx))
            except Exception:
                continue

        if not cands:
            return None
        cands.sort()
        cy, cx = cands[0]
        return cx, cy

    def _pay_find_red_button(self) -> Optional[tuple]:
        """화면에서 빨간색 '삭제하기' 버튼 영역(222.png) 인식 → 중앙 좌표 반환"""
        try:
            import cv2
            import numpy as np
            png = self._get_screenshot()
            if not png:
                return None
            img = cv2.imdecode(np.frombuffer(png, np.uint8), cv2.IMREAD_COLOR)
            if img is None:
                return None
            hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
            m1 = cv2.inRange(hsv, (0, 150, 150), (8, 255, 255))
            m2 = cv2.inRange(hsv, (170, 150, 150), (180, 255, 255))
            mask = cv2.morphologyEx(cv2.bitwise_or(m1, m2), cv2.MORPH_CLOSE, np.ones((9, 9), np.uint8))
            contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            h_img, w_img = img.shape[:2]
            best = None
            for c in contours:
                x, y, w, h = cv2.boundingRect(c)
                # 버튼 크기 필터 (1080 기준 약 330x125)
                if w < w_img * 0.15 or h < h_img * 0.025 or w > w_img * 0.9:
                    continue
                area = w * h
                if best is None or area > best[0]:
                    best = (area, x + w // 2, y + h // 2)
            if best:
                self._log(f"  🎯 [Pay] 빨간색 '삭제하기' 버튼 인식 → ({best[1]}, {best[2]})")
                return best[1], best[2]
        except Exception as e:
            self._log(f"  ⚠ [Pay] 빨간 버튼 인식 예외: {e}")
        return None

    def _pay_confirm_delete(self) -> bool:
        """삭제 확인 팝업의 빨간색 '삭제하기' 클릭"""
        time.sleep(1)
        for xp in PAY_DELETE_CONFIRM_XPATHS:
            if ah.element_exists(self.driver, xp, timeout=1.5):
                try:
                    el = self.driver.find_element(By.XPATH, xp)
                    try:
                        el.click()
                    except Exception:
                        self._click_element_center_coordinates(el)
                    self._log(f"  ✅ [Pay] 삭제 확인 클릭 (XPath: {xp})")
                    return True
                except Exception:
                    continue
        coords = self._pay_find_red_button()
        if coords:
            self._pay_adb_tap(*coords)
            self._log("  ✅ [Pay] 빨간색 '삭제하기' 버튼 좌표 탭 완료")
            return True
        for img_name in ("삭제하기.png", "삭제하기1.png", "삭제하기2.png", "삭제하기3.png"):
            p = os.path.join(os.path.dirname(os.path.abspath(__file__)), img_name)
            if os.path.exists(p):
                c = self._find_image_coords(p, threshold=0.75)
                if c:
                    self._pay_adb_tap(c[0], c[1])
                    self._log(f"  ✅ [Pay] {img_name} 인식 클릭")
                    return True
        self._log("  ⚠ [Pay] 삭제 확인 버튼 미발견")
        return False

    def _pay_delete_existing_addresses(self, max_loops: int = 30):
        """기본배송지 제외 주소를 전부 삭제될 때까지 반복 삭제"""
        self._set_status("[Pay] 기존 배송지 삭제 중")
        self._log("🗑 [Pay] 기본배송지 제외 배송지 삭제 시작")
        deleted = 0
        no_found_count = 0
        last = None
        same_count = 0
        for loop in range(1, max_loops + 1):
            if self._stop_event.is_set():
                break
            target = self._pay_find_non_default_delete_btn()
            if not target:
                no_found_count += 1
                if no_found_count >= 2:
                    break
                self._log("  📜 [Pay] 삭제 대상 미발견 → 스크롤 후 재확인")
                self._scroll_down()
                time.sleep(1.2)
                continue
            no_found_count = 0
            if target == last:
                same_count += 1
                if same_count >= 3:
                    self._log(f"  ⚠ [Pay] 동일 좌표 {target} 3회 연속 → 삭제 루프 종료")
                    break
            else:
                same_count = 0
            last = target

            self._log(f"  👉 [Pay] 삭제 버튼 탭 (loop {loop}, 좌표 {target})")
            self._pay_adb_tap(*target)
            if self._pay_confirm_delete():
                deleted += 1
            time.sleep(2.5)
            self._dismiss_hide_popup(max_count=1)

        self._log(f"🗑 [Pay] 삭제 완료 (총 {deleted}건)")
        self._scroll_to_top()
        time.sleep(1)

    # ─── [7~9] 배송지 1건 등록 ──────────────────────────────────────────────

    def _pay_select_phone_prefix(self, n: int) -> None:
        """연락처{n} 첫번째 번호가 010 이 아니면 010 선택"""
        xp = PAY_PHONE_FIRST_XPATH.format(n=n)
        try:
            el = ah.wait_for_element(self.driver, xp, timeout=2)
            if el is None:
                return
            cur = (el.text or el.get_attribute("text") or el.get_attribute("content-desc") or "")
            if "010" in cur:
                return
            # 하위 텍스트 확인
            try:
                if el.find_elements(By.XPATH, './/*[@text="010"]'):
                    return
            except Exception:
                pass
            self._log(f"  📞 [Pay] 연락처{n} 첫번째 번호 선택 → 010")
            el.click()
            time.sleep(1.2)
            for p in ('//android.widget.CheckedTextView[@text="010"]',
                      '//android.widget.TextView[@text="010"]',
                      '//*[@text="010" and @clickable="true"]',
                      '//*[@text="010"]'):
                if ah.element_exists(self.driver, p, timeout=1.5):
                    ah.wait_and_click(self.driver, p, timeout=2, log_callback=self._log)
                    time.sleep(0.8)
                    return
        except Exception as e:
            self._log(f"  ⚠ [Pay] 010 선택 예외: {str(e).splitlines()[0]}")

    def _pay_input_phone(self, row, n: int) -> bool:
        mid, last = row.get_phone_middle(), row.get_phone_last()
        if not mid or not last:
            self._log(f"  ⚠ [Pay] 전화번호 추출 실패 (원본: '{row.phone}')")
            return False
        mid_xp = PAY_PHONE_MID_XPATH.format(n=n)
        if not ah.element_exists(self.driver, mid_xp, timeout=2):
            self._scroll_down()
            time.sleep(0.8)
        self._pay_select_phone_prefix(n)
        ok1 = self._pay_input_field(mid_xp, mid, f"연락처{n} 중간번호")
        ok2 = self._pay_input_field(PAY_PHONE_LAST_XPATH.format(n=n), last, f"연락처{n} 마지막번호")
        return ok1 and ok2

    def _pay_check_default(self) -> bool:
        """'기본 배송지로 설정' 체크 (단 1회만 클릭하여 토글 해제 방지)"""
        self._pay_hide_keyboard()
        time.sleep(0.5)

        # 1. 이미 체크되어 있는지 먼저 확인
        for xp in PAY_DEFAULT_CHECK_XPATHS:
            try:
                els = self.driver.find_elements(By.XPATH, xp)
                for el in els:
                    if self._is_checkbox_checked(el):
                        self._log("  ✅ [Pay] '기본 배송지로 설정' 이미 체크됨 (중복 클릭 방지)")
                        return True
            except Exception:
                pass

        # 2. 화면 노출 여부 확인 및 미노출 시 스크롤 다운 (1회)
        check_candidates = [
            PAY_DEFAULT_TEXT_XPATH,  # '//android.widget.TextView[@text="기본 배송지로 설정"]'
            '//android.widget.CheckBox[@resource-id="default-address"]',
            '//android.widget.CheckBox[@text="기본 배송지로 설정"]',
            '//*[@text="기본 배송지로 설정"]',
        ]
        has_visible = False
        for xp in check_candidates:
            if ah.element_exists(self.driver, xp, timeout=1):
                has_visible = True
                break

        if not has_visible:
            self._log("  ⬇ [Pay] '기본 배송지로 설정' 미노출 → 스크롤 다운")
            self._scroll_down()
            time.sleep(0.8)

        # 3. 단 1회만 클릭 (텍스트뷰 또는 체크박스 중 1개 클릭 성공 시 즉시 종료)
        self._set_status("[Pay] 기본배송지 체크")
        for xp in check_candidates:
            try:
                el = ah.wait_for_element(self.driver, xp, timeout=1.5)
                if el and el.is_displayed():
                    clicked = False
                    try:
                        el.click()
                        clicked = True
                    except Exception:
                        clicked = self._click_element_center_coordinates(el)

                    if clicked:
                        self._log(f"  ✅ [Pay] '기본 배송지로 설정' 1회 클릭 완료 (중복 클릭 방지, xpath: {xp})")
                        time.sleep(0.8)
                        return True
            except Exception:
                continue

        # 4. 폴백: 기존 공통 기본배송지 체크 로직 호출
        self._log("  ℹ [Pay] 전용 체크 요소 미발견 → 공통 기본배송지 체크 로직 시도")
        return self._check_default_delivery_address()

    def _pay_open_address_search(self) -> bool:
        if self._pay_click_first(PAY_ADDR_SEARCH_BTN_XPATHS, "주소검색", timeout=4, wait_after=3):
            return True
        self._log("❌ [Pay] 주소검색 버튼 클릭 실패")
        return False

    def _pay_handle_detail_address(self, row) -> bool:
        """주소 선택 이후 상세주소 처리
        1. 상세주소입력창.xml ('상세주소를 알려주세요' 팝업/화면):
           - //android.widget.EditText[@resource-id="address-detail"] 에 상세주소 입력
           - //android.widget.Button[@text="확인"] 버튼 클릭하여 배송지 등록/수정 화면으로 복귀
        2. 구형 스토어 팝업:
           - 상세주소 입력 후 '선택완료' 버튼 클릭
        3. 배송지 등록/수정 화면 (수정.xml):
           - //android.widget.EditText[@resource-id="detailed-address"] 값 확인 및 미입력 시 보충 입력
        """
        detail = (row.detail_address or "").strip()
        time.sleep(1.5)

        # ── 1. 상세주소입력창.xml ('상세주소를 알려주세요') 감지 ──
        is_popup_addr_detail = False
        addr_detail_xp = None
        for xp in PAY_POPUP_DETAIL_INPUT_XPATHS:
            if ah.element_exists(self.driver, xp, timeout=2):
                is_popup_addr_detail = True
                addr_detail_xp = xp
                break

        if not is_popup_addr_detail and ah.element_exists(self.driver, '//*[contains(@text, "상세주소를 알려주세요")]', timeout=2):
            is_popup_addr_detail = True

        if is_popup_addr_detail:
            self._log(f"  📌 [Pay] '상세주소를 알려주세요' 화면 감지 (address-detail) → 상세주소 '{detail}' 입력 진행")
            if detail:
                input_ok = False
                if addr_detail_xp:
                    input_ok = self._pay_input_field(addr_detail_xp, detail, "상세주소(address-detail)")
                if not input_ok:
                    try:
                        # address-detail 포커스 후 mobile: type
                        self.driver.execute_script("mobile: type", {"text": detail})
                        time.sleep(0.5)
                        input_ok = True
                        self._log(f"  ✅ [Pay] 상세주소 입력 완료 (mobile: type): '{detail}'")
                    except Exception as e:
                        self._log(f"  ⚠ [Pay] 상세주소 mobile: type 입력 예외: {e}")

            self._pay_hide_keyboard()
            time.sleep(0.5)

            # '확인' 버튼 클릭
            self._set_status("[Pay] 상세주소 확인 클릭")
            confirm_clicked = False
            for c_xp in PAY_POPUP_DETAIL_CONFIRM_XPATHS:
                if ah.element_exists(self.driver, c_xp, timeout=2):
                    self._log(f"  👉 [Pay] 상세주소 '확인' 버튼 발견 ({c_xp}) → 클릭")
                    if ah.wait_and_click(self.driver, c_xp, timeout=3, log_callback=self._log):
                        confirm_clicked = True
                        time.sleep(2.5)
                        break
                    try:
                        el = self.driver.find_element(By.XPATH, c_xp)
                        if self._click_element_center_coordinates(el):
                            confirm_clicked = True
                            time.sleep(2.5)
                            break
                    except Exception:
                        pass

            if not confirm_clicked:
                # 좌표 폴백: 상세주소입력창.xml 기준 bounds="[49,955][1031,1094]" -> (540, 1024)
                try:
                    sz = self.driver.get_window_size()
                    cx = sz.get("width", 1080) // 2
                    cy = int(sz.get("height", 2094) * 0.49)
                    self._log(f"  ⚠ [Pay] '확인' 버튼 XPath 실패 → 좌표 탭 시도 ({cx}, {cy})")
                    self._pay_adb_tap(cx, cy)
                    time.sleep(2.5)
                    confirm_clicked = True
                except Exception as e_tap:
                    self._log(f"  ⚠ [Pay] 좌표 탭 실패: {e_tap}")

        # ── 2. 구형 주소검색 팝업 ('선택완료') 감지 (폴백 호환성) ──
        popup_select_done = (
            ah.element_exists(self.driver, '//*[@text="선택완료" or @content-desc="선택완료"]', timeout=1)
            or ah.element_exists(self.driver, '//*[contains(@text, "상세주소를 입력")]', timeout=1)
        )
        if popup_select_done:
            self._log("  ℹ [Pay] 구형 주소검색 팝업('선택완료') 감지 → 상세주소 입력 후 선택완료 클릭")
            if detail:
                try:
                    self.driver.execute_script("mobile: type", {"text": detail})
                    time.sleep(0.5)
                except Exception as e:
                    self._log(f"  ⚠ [Pay] 팝업 상세주소 입력 실패: {e}")
            self._pay_hide_keyboard()
            if not self._pay_click_first(['//android.widget.Button[@text="선택완료"]', '//*[@text="선택완료"]'],
                                         "선택완료", timeout=2, wait_after=2):
                self._click_confirm_button()
                time.sleep(2)

        # ── 3. 등록/수정 폼 (수정.xml) 복귀 확인 및 detailed-address 필드 점검 ──
        time.sleep(1.5)
        for xp in PAY_DETAIL_ADDR_XPATHS:
            el = ah.wait_for_element(self.driver, xp, timeout=2)
            if el is None:
                continue
            try:
                cur = (el.text or el.get_attribute("text") or "").strip()
            except Exception:
                cur = ""
            if detail and detail not in cur:
                self._log(f"  ℹ [Pay] 등록/수정 화면 detailed-address 확인(현재: '{cur}') → 보충 입력")
                return self._pay_input_field(xp, detail, "상세주소")
            self._log(f"  ✅ [Pay] 등록/수정 화면 상세주소 확인: '{cur}'")
            return True

        if is_popup_addr_detail or popup_select_done:
            self._log("  ✅ [Pay] 상세주소 팝업 단계 처리 완료 (등록/수정 화면으로 복귀)")
            return True

        if not detail:
            return True
        self._log("  ⚠ [Pay] 상세주소 입력창(address-detail / detailed-address) 미발견")
        return False

    def _register_one_address_pay(self, row) -> bool:
        """[Pay 방식] 배송지 1건 등록 (pay주소지 등록.md 5~9 단계)"""

        # 계정 전환 및 배송지 관리 화면 진입
        if getattr(self, "current_naver_id", None) != row.naver_id or not getattr(self, "is_initialized", False):
            self._log(f"🔄 [Pay] 계정 전환/초기화 (현재: {getattr(self, 'current_naver_id', '없음')} → 대상: {row.naver_id})")
            if not self._pay_go_main(login_id=row.naver_id):
                return False
            if not self._navigate_to_pay_delivery_mgmt(row=row):
                self._log("❌ [Pay] 배송지 관리 화면 진입 실패")
                return False
            self.current_naver_id = row.naver_id
            self.is_initialized = True
            self.has_checked_initial_delete = False

        # [5~6] 기본배송지 제외 삭제 (폰ID+아이디 별 최초 1회, 9열 Y 인 경우)
        if not self.has_checked_initial_delete:
            self.has_checked_initial_delete = True
            if row.delete_existing:
                self._log("🗑 [Pay] 9열 값 Y → 기본배송지 제외 배송지 삭제 진행")
                self._pay_delete_existing_addresses()
            else:
                self._log("⏭ [Pay] 9열 값 Y 아님 → 기존 배송지 삭제 건너뜀")

        # [7] 배송지 추가
        self._set_status(f"[Pay] 배송지 추가: {row.name}")
        if not self._ensure_pay_delivery_mgmt_screen():
            self._log("❌ [Pay] 배송지 관리 화면 복귀 실패")
            return False
        if not self._pay_click_first(PAY_ADD_ADDRESS_XPATHS, "배송지 추가", timeout=5, wait_after=3):
            self._log("❌ [Pay] '배송지 추가' 버튼 클릭 실패")
            return False

        if not ah.element_exists(self.driver, PAY_FORM_TITLE_XPATH, timeout=8):
            self._log("  ⚠ [Pay] '배송지 등록/수정' 타이틀 미확인 → 입력창 기준으로 계속 진행")

        # [8-1] 수령인
        self._set_status(f"[Pay] 수령인 입력: {row.name}")
        receiver_ok = False
        for xp in PAY_RECEIVER_XPATHS:
            if ah.element_exists(self.driver, xp, timeout=3):
                receiver_ok = self._pay_input_field(xp, row.name, "수령인")
                if receiver_ok:
                    break
        if not receiver_ok:
            self._log("❌ [Pay] 수령인 입력 실패")
            self._skip_row_retry = True
            return False
        self._pay_hide_keyboard()

        # [8-2] 주소검색 → 검색어 입력 → 검색 → 우편번호 매칭 선택 (기존 로직 재사용)
        self._set_status(f"[Pay] 주소 검색: {row.address_search}")
        if not self._pay_open_address_search():
            return False
        if not self._input_address_search(row.address_search):
            self._log("❌ [Pay] 주소 검색어 입력 실패")
            return False
        if not self._click_search_button():
            self._log("❌ [Pay] 검색 버튼 클릭 실패")
            return False
        time.sleep(2.5)
        self._set_status(f"[Pay] 우편번호 매칭: {row.zipcode}")
        if not self._select_address_by_zipcode(row.zipcode):
            self._log(f"❌ [Pay] 우편번호 {row.zipcode} 매칭 실패")
            return False

        # [8-3] 상세주소
        self._set_status("[Pay] 상세주소 입력")
        if not self._pay_handle_detail_address(row):
            self._log("❌ [Pay] 상세주소 입력 실패")
            return False
        self._pay_hide_keyboard()

        # 우편번호/주소 필드가 채워졌는지 확인 (로그용)
        try:
            z = self.driver.find_element(By.XPATH, PAY_ZIPCODE_FIELD_XPATH)
            self._log(f"  ℹ [Pay] 폼 우편번호 값: '{(z.text or '').strip()}'")
        except Exception:
            pass

        # [8-4] 연락처1 / 연락처2
        self._set_status("[Pay] 연락처 입력")
        if not self._pay_input_phone(row, 1):
            self._log("❌ [Pay] 연락처1 입력 실패")
            return False
        if ah.element_exists(self.driver, PAY_PHONE_MID_XPATH.format(n=2), timeout=2):
            self._pay_input_phone(row, 2)
        self._pay_hide_keyboard()

        # [8-5] 기본 배송지로 설정 체크
        self._pay_check_default()

        # [9] 저장
        self._set_status("[Pay] 저장 클릭")
        time.sleep(1)
        if not self._pay_click_first(PAY_SAVE_BTN_XPATHS, "저장", timeout=3, wait_after=2.5):
            self._scroll_down()
            time.sleep(0.8)
            if not self._pay_click_first(PAY_SAVE_BTN_XPATHS, "저장", timeout=3, wait_after=2.5):
                self._log("❌ [Pay] 저장 버튼 클릭 실패")
                return False

        # 저장 후 알림/오류 팝업 처리 (동일 주소 존재 등)
        self._dismiss_hide_popup(max_count=1)
        for xp in ('//android.widget.Button[@text="확인"]',
                   '//android.widget.Button[@resource-id="android:id/button1"]'):
            if ah.element_exists(self.driver, xp, timeout=1.5):
                self._log("  📌 [Pay] 저장 후 알림 팝업 '확인' 클릭")
                ah.wait_and_click(self.driver, xp, timeout=2, log_callback=self._log)
                time.sleep(1.5)
                break

        # 저장 성공 판정: 배송지 관리 목록 복귀
        if self._pay_find_add_button(timeout=6):
            self._log(f"✅ [Pay] {row.name} 배송지 등록 완료")
            return True

        if ah.element_exists(self.driver, PAY_FORM_TITLE_XPATH, timeout=2):
            self._log("❌ [Pay] 저장 후에도 등록/수정 화면 유지 → 입력값 오류 가능성")
            try:
                self.driver.back()
                time.sleep(2)
            except Exception:
                pass
            return False

        self._log(f"✅ [Pay] {row.name} 저장 완료 (목록 복귀 미확인)")
        return True
