"""
naver_review_worker.py
네이버 쇼핑 앱 자동 리뷰 워커 (핸드폰 1대 기준)

자동주문 워커(NaverOrderWorker)를 상속하여 앱 재시작 / 계정 전환 / 팝업 처리 /
마이쇼핑 진입 로직은 그대로 재사용하고, 리뷰 작성 흐름만 새로 구현한다.

흐름 (개발문서_자동리뷰.md):
1   네이버 앱 종료 → 로그인 아이디 선택(Sheet2) → 마이쇼핑 진입
2   마이쇼핑 '리뷰작성' 클릭
3   리뷰 목록에서 '한달사용리뷰 쓰기' / '리뷰쓰기' 버튼 탐색
    → 버튼이 속한 카드의 상품명과 댓글목록 검색어 매칭 (작업여부 공백 행만)
    → 매칭 댓글 W 예약 후 버튼 클릭
4   별점 5 + 질문별 RadioGroup 마지막 항목(촉촉해요/아주좋아요 등) 선택 (조금씩 스크롤)
    → reviewInput 에 댓글 입력 → 작업여부 Y → 스크롤 2회 → 등록 클릭 (3초)
5   '다음에 동의하기' 있으면 클릭
6   '리뷰가등록되었어요' 포함 요소 있으면 성공, 없으면 작업여부 F (3초)
7   '닫기' 클릭
    → 3~7 반복 (리뷰 버튼이 안 보일 때까지, 최대 100회)
8   해당 폰ID 의 다음 로그인 아이디로 진행
"""

import os
import re
import sys
import time
import datetime
import threading
import xml.etree.ElementTree as ET
from typing import List, Optional

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from selenium.webdriver.common.by import By

import naver_order_worker as _now
from naver_order_worker import NaverOrderWorker, _LOG_WRITE_QUEUE, _run_cmd
from review_manager import ReviewManager, CommentRow, AccountRow, ExcelBusyError


def _ah():
    """naver_order_worker 가 드라이버 생성 시 모듈을 재로딩하므로 항상 최신 참조 사용"""
    return _now.ah


# ─── XPath 상수 ─────────────────────────────────────────────────────────────

# [단계 2] 마이쇼핑 → 리뷰작성
REVIEW_MENU_XPATHS = [
    '//android.view.View[@content-desc="리뷰작성"]/android.view.View/android.widget.Image',
    '//android.view.View[@content-desc="리뷰작성"]',
    '//*[@text="리뷰작성"]',
]

# [단계 3] 리뷰 목록 버튼
REVIEW_BTN_TEXTS = ("한달사용리뷰 쓰기", "리뷰쓰기")
REVIEW_BTN_XPATH = '//android.widget.Button[@text="한달사용리뷰 쓰기" or @text="리뷰쓰기"]'

# [단계 4] 리뷰 작성 화면
#   별점/만족도는 RadioGroup 단위로 선택 (parse_radio_groups 참고)
#   - 질문 그룹: 항상 마지막 RadioButton (촉촉해요 / 아주좋아요 / 아주만족해요 ...)
#   - 별점 그룹: 순서가 5,4,3,2,1 이라 마지막이 1점 → '5' 클릭
REVIEW_INPUT_XPATH = '//android.widget.EditText[@resource-id="reviewInput"]'
SUBMIT_BTN_XPATH = '//android.widget.Button[@text="등록"]'
REVIEW_FORM_XPATH = " | ".join([
    REVIEW_INPUT_XPATH,
    '//android.widget.RadioButton',
    '//android.webkit.WebView[contains(@text,"리뷰쓰기")]',
])

# [단계 5~7]
AGREE_LATER_XPATH = '//android.widget.Button[@text="다음에 동의하기"]'
SUCCESS_TEXT = "리뷰가등록되었어요"   # 공백 제거 후 비교
# 리뷰 목록 화면 표시 (공백 제거 후 비교) — 등록 후 완료 화면 없이 목록으로 바로 돌아오는 경우
LIST_PAGE_MARKERS = ("작성가능한리뷰", "마이쇼핑리뷰작성")
CLOSE_BTN_XPATH = '//android.widget.Button[@text="닫기"]'
# 리뷰 페이지 진입 시 시스템 알림창 (이어서작성.XML)
#   "작성 중이던 리뷰가 있습니다. 이어서 작성하시겠어요?"  [취소] [확인=button1]
ALERT_MESSAGE_XPATH = '//android.widget.TextView[@resource-id="android:id/message"]'
ALERT_OK_XPATH = '//android.widget.Button[@resource-id="android:id/button1"]'
LEAVE_CONFIRM_XPATH = (
    '//android.widget.Button[@resource-id="android:id/button1"]'
    ' | //android.widget.Button[@text="확인" or @text="나가기"]'
)

HIDE_POPUP_XPATHS = [
    '//*[contains(@text,"하루") and contains(@text,"보지 않")]',
    '//*[contains(@text,"하루") and contains(@text,"보기 않")]',
    '//*[contains(@text,"7일") and contains(@text,"보지 않")]',
    '//*[contains(@content-desc,"하루") and contains(@content-desc,"보지 않")]',
    '//*[contains(@content-desc,"7일") and contains(@content-desc,"보지 않")]',
    '//*[contains(@text,"다시 보지 않기")]',
]

MAX_REVIEWS_PER_ACCOUNT = 100   # 아이디당 최대 반복 (단계 3~7)
MAX_CONSECUTIVE_FAILS = 3       # 연속 실패 시 해당 아이디 중단
MAX_LIST_SCROLLS = 8            # 리뷰 목록 아래 방향 탐색 스크롤 횟수
MAX_RATING_ROUNDS = 16          # 별점/만족도 탐색 반복(스크롤) 횟수

_BOUNDS_RE = re.compile(r"\[(-?\d+),(-?\d+)\]\[(-?\d+),(-?\d+)\]")


class StopRequested(Exception):
    pass


# ─── 리뷰 목록 화면 파싱 (page_source XML) ───────────────────────────────────

def _parse_bounds(s: str):
    m = _BOUNDS_RE.match(s or "")
    return tuple(map(int, m.groups())) if m else None


def _norm(text) -> str:
    return re.sub(r"\s+", "", str(text or ""))


class ReviewButton:
    def __init__(self, text, bounds, texts, visible):
        self.text = text
        self.bounds = bounds
        self.texts = texts          # 카드의 상품명 등 텍스트
        self.visible = visible
        x1, y1, x2, y2 = bounds
        self.center = ((x1 + x2) // 2, (y1 + y2) // 2)

    @property
    def product(self) -> str:
        return max(self.texts, key=len) if self.texts else ""

    @property
    def product_key(self) -> str:
        return _norm(self.product)


def _subtree_has_review_button(node) -> bool:
    for el in node.iter():
        if el.tag.endswith("Button") and (el.get("text") or "").strip() in REVIEW_BTN_TEXTS:
            return True
    return False


def _subtree_texts(node) -> List[str]:
    out = []
    for el in node.iter():
        for key in ("text", "content-desc"):
            val = (el.get(key) or "").strip()
            if not val or val in out:
                continue
            if val.startswith("[") or val.startswith("작성기한"):
                continue  # 스토어명 / 작성기한
            out.append(val)
    return out


def _card_texts(button, parent_map) -> List[str]:
    """
    리뷰작성.xml 구조: 상품 카드 View 바로 다음 형제 View 안에 리뷰 버튼이 있다.
    버튼에서 위로 올라가며 가장 가까운 '앞쪽 형제' 중 텍스트가 있는 카드를 찾는다.
    """
    node = button
    for _ in range(4):
        parent = parent_map.get(node)
        if parent is None:
            break
        siblings = list(parent)
        idx = siblings.index(node)
        for sib in reversed(siblings[:idx]):
            if _subtree_has_review_button(sib):
                return []  # 이전 카드의 버튼에 도달 → 이 버튼의 카드 정보 없음
            texts = _subtree_texts(sib)
            if texts:
                return texts
        node = parent
    return []


def parse_review_buttons(page_source: str) -> List[ReviewButton]:
    """page_source 에서 리뷰 버튼과 버튼이 속한 상품 카드 텍스트를 위→아래 순으로 추출"""
    root = ET.fromstring(page_source.encode("utf-8"))
    parent_map = {c: p for p in root.iter() for c in p}

    # 보이는 영역 = 가장 큰 WebView
    view = None
    for el in root.iter():
        if el.tag.endswith("WebView"):
            b = _parse_bounds(el.get("bounds"))
            if b and (view is None or (b[2] - b[0]) * (b[3] - b[1]) > (view[2] - view[0]) * (view[3] - view[1])):
                view = b
    top, bottom = (view[1], view[3]) if view else (0, 10 ** 6)

    buttons = []
    for el in root.iter():
        if not el.tag.endswith("Button"):
            continue
        text = (el.get("text") or "").strip()
        if text not in REVIEW_BTN_TEXTS:
            continue
        b = _parse_bounds(el.get("bounds"))
        if not b:
            continue
        visible = (b[3] - b[1]) >= 40 and b[1] >= top + 4 and b[3] <= bottom - 4
        buttons.append(ReviewButton(text, b, _card_texts(el, parent_map), visible))
    buttons.sort(key=lambda x: x.bounds[1])
    return buttons


class RadioChoice:
    """RadioGroup 1개 = 질문 1개. 클릭할 RadioButton 정보"""

    def __init__(self, question, options, label, bounds, visible):
        self.question = question
        self.options = options
        self.label = label
        self.bounds = bounds
        self.visible = visible
        x1, y1, x2, y2 = bounds
        self.center = ((x1 + x2) // 2, (y1 + y2) // 2)

    @property
    def key(self) -> str:
        return f"{self.question}|{'/'.join(self.options)}"


def parse_radio_groups(page_source: str):
    """
    리뷰 작성 화면의 RadioGroup 목록과 클릭 대상 추출 (촉촉해요.xml 구조).
    WebView 의 RadioButton 은 checkable=false 라 선택 여부를 알 수 없으므로
    질문 텍스트로 그룹을 구분해 호출 측에서 1회만 클릭한다.
    Returns: (RadioChoice 목록, 리뷰 입력란이 화면에 보이는지)
    """
    root = ET.fromstring(page_source.encode("utf-8"))
    nodes = list(root.iter())

    # 보이는 영역: 가장 큰 WebView, 상단 고정 헤더('리뷰쓰기' 제목/닫기 버튼) 아래부터
    view = None
    header_bottom = 0
    for el in nodes:
        b = _parse_bounds(el.get("bounds"))
        if not b:
            continue
        if el.tag.endswith("WebView"):
            if view is None or (b[2] - b[0]) * (b[3] - b[1]) > (view[2] - view[0]) * (view[3] - view[1]):
                view = b
        elif el.tag.endswith("Button") and (el.get("text") or "").strip() == "닫기":
            header_bottom = max(header_bottom, b[3])
    top, bottom = (view[1], view[3]) if view else (0, 10 ** 6)
    top = max(top, header_bottom)

    def on_screen(b):
        cy = (b[1] + b[3]) // 2
        return (b[3] - b[1]) >= 40 and top + 15 < cy < bottom - 15

    groups = []
    inside_group = set()
    question = ""
    input_visible = False
    for el in nodes:
        if el in inside_group:
            continue
        if el.get("resource-id") == "reviewInput":
            b = _parse_bounds(el.get("bounds"))
            input_visible = input_visible or bool(b and on_screen(b))
            continue
        if el.tag.endswith("RadioGroup"):
            inside_group.update(d for d in el.iter() if d is not el)
            radios = [d for d in el.iter() if d.tag.endswith("RadioButton")]
            options = [(r.get("text") or "").strip() for r in radios]
            if not radios:
                continue
            if "5" in options and all(o.isdigit() for o in options if o):
                target = radios[options.index("5")]      # 별점: 5점
            else:
                target = radios[-1]                       # 질문: 마지막 항목
            b = _parse_bounds(target.get("bounds"))
            if b:
                groups.append(RadioChoice(question, options, (target.get("text") or "").strip(),
                                          b, on_screen(b)))
            continue
        txt = (el.get("text") or "").strip()
        if txt and not el.tag.endswith("WebView") and txt not in ("리뷰쓰기", "닫기"):
            question = txt   # 그룹 바로 앞의 텍스트 = 질문 (예: '촉촉함은 어떤가요?')
    return groups, input_visible


def classify_page(page_source: str) -> str:
    """
    현재 화면 판별
      'success' : '리뷰가등록되었어요' 표시
      'form'    : 리뷰 작성 화면 (입력란 / '리뷰쓰기' 웹뷰 제목 / 등록 버튼)
      'list'    : 리뷰 목록 (작성 가능한 리뷰 탭 / 리뷰 작성 버튼)
      'other'   : 그 외
    """
    if SUCCESS_TEXT in _norm(page_source):
        return "success"
    try:
        root = ET.fromstring(page_source.encode("utf-8"))
    except ET.ParseError:
        return "other"
    for el in root.iter():
        text = (el.get("text") or "").strip()
        if el.get("resource-id") == "reviewInput":
            return "form"
        if el.tag.endswith("WebView") and "리뷰쓰기" in text:
            return "form"
        if el.tag.endswith("Button") and text == "등록":
            return "form"
    if any(m in _norm(page_source) for m in LIST_PAGE_MARKERS):
        return "list"
    if parse_review_buttons(page_source):
        return "list"
    return "other"


class _Target:
    def __init__(self, button: ReviewButton, row: CommentRow):
        self.button = button
        self.row = row


# ─── 워커 ───────────────────────────────────────────────────────────────────

class NaverReviewWorker(NaverOrderWorker):
    """네이버 자동 리뷰 워커 (기기 1대)"""

    def __init__(self, device_id: str, appium_port: int, review_manager: ReviewManager,
                 log_callback=None, status_callback=None, machine_num: int = 1):
        super().__init__(
            device_id=device_id,
            appium_port=appium_port,
            order_manager=None,
            log_callback=log_callback,
            status_callback=status_callback,
            machine_num=machine_num,
        )
        self.review_manager = review_manager
        self.current_account: Optional[AccountRow] = None
        self.current_comment: Optional[CommentRow] = None
        self._failed_products = set()
        self._review_no = 0

    # ─── 공개 메서드 ─────────────────────────────────────────────────────────

    def run(self) -> bool:
        """워커 메인 (별도 스레드). 드라이버 오류 시 재연결 후 미완료 아이디부터 이어서 진행."""
        if self._stop_event.is_set():
            return False
        self._log("🚀 자동 리뷰 워커 시작")
        try:
            _ah().wake_screen_if_off(self.device_id, self._log)
        except Exception:
            pass

        max_restarts = 10
        for attempt in range(1, max_restarts + 1):
            if self._stop_event.is_set():
                break
            self._set_status(f"연결 중... ({attempt}/{max_restarts})")
            try:
                self.driver = self._create_driver()
            except Exception as e:
                if self._stop_event.is_set() or "중지 요청" in str(e):
                    break
                self._log(f"❌ 드라이버 연결 실패: {str(e).splitlines()[0][:200]}")
                self._set_status("연결 실패")
                if attempt < max_restarts and self._sleep_interruptible(10):
                    continue
                return False

            finished = False
            try:
                self._account_loop()
                finished = True
            except StopRequested:
                pass
            except ExcelBusyError as e:
                self._log(f"❌ {e}")
                self._set_status("엑셀 저장 실패")
                finished = True
            except Exception as e:
                if not self._stop_event.is_set():
                    self._log(f"❌ 예기치 않은 오류: {str(e).splitlines()[0][:200]}")
                    self._set_status("오류 발생")
            finally:
                self._release_current_comment()
                if self.driver:
                    try:
                        self.driver.quit()
                    except Exception:
                        pass
                self.driver = None

            if finished or self._stop_event.is_set():
                break
            self._log(f"🔄 오류 회복 재시작... ({attempt}/{max_restarts})")
            if not self._sleep_interruptible(5):
                break

        if self._stop_event.is_set():
            self._log("✅ 중지 완료")
            self._set_status("중지됨", "-")
        else:
            self._log("🏁 워커 종료")
            self._set_status("완료", "-")
        return True

    def stop(self):
        self._stop_event.set()
        self._log("⏹ 중지 요청됨")
        # GUI 스레드에서 호출되므로 엑셀 기록은 백그라운드로 (run()의 finally 에서도 한 번 더 정리)
        threading.Thread(target=self._release_current_comment, daemon=True).start()
        self._set_status("중지 중...", "-")
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        self.driver = None

    # ─── 아이디 루프 (단계 1, 8, 9) ──────────────────────────────────────────

    def _account_loop(self):
        while True:
            self._check_stop()
            account = self.review_manager.next_pending_account(self.device_id)
            if not account:
                self._log("✅ 이 기기의 미완료 로그인 아이디 없음 → 작업 종료")
                break

            self.current_account = account
            self.current_payment_method = account.login_id   # 패널 '아이디' 표시용
            result = self._process_account(account)
            self._check_stop()

            self.review_manager.set_account_status(account.row_index, result)
            mark = "완료(Y)" if result == "Y" else "실패(F)"
            self._log(f"👤 [{account.login_id}] 아이디 작업 {mark} 기록")
            self.current_account = None

        self.current_payment_method = ""

    def _process_account(self, account: AccountRow) -> str:
        """로그인 아이디 1개 처리. 반환: 'Y' (리뷰 버튼 소진) / 'F' (로그인·진입 실패, 연속 실패)"""
        self._log(f"👤 로그인 아이디 작업 시작: {account.login_id} (row={account.row_index}, 순번={account.seq})")
        self._failed_products = set()

        # [단계 1] 앱 종료 → 아이디 선택 → 마이쇼핑
        if not self._go_main_and_enter_store(login_id=account.login_id):
            self._log(f"❌ [{account.login_id}] 계정 전환/마이쇼핑 진입 실패")
            return "F"
        self._check_stop()

        # [단계 2] 리뷰작성
        if not self._open_review_list(retry_nav=True):
            return "F"

        done = failed = consecutive = 0
        reopened = False
        for i in range(1, MAX_REVIEWS_PER_ACCOUNT + 1):
            self._check_stop()
            self._set_status(f"리뷰 목록 탐색 ({i}회차)")
            target, saw_buttons = self._find_and_claim_target()

            if target is None:
                if not saw_buttons and not reopened:
                    # 닫기 후 다른 화면에 남아있을 수 있으므로 리뷰작성 화면 1회 재진입 확인
                    reopened = True
                    self._log("ℹ 리뷰 버튼 미발견 → 리뷰작성 화면 재진입 후 재확인")
                    self._go_main_and_enter_store(login_id="")
                    if self._open_review_list(retry_nav=False):
                        continue
                if saw_buttons:
                    self._log("ℹ 리뷰 버튼은 남아있으나 매칭되는 댓글(작업여부 공백) 없음 → 아이디 종료")
                else:
                    self._log("✅ 리뷰쓰기 / 한달사용리뷰 쓰기 버튼 없음 → 아이디 종료")
                break

            reopened = False
            ok = self._write_one_review(target)
            if ok:
                done += 1
                consecutive = 0
            else:
                failed += 1
                consecutive += 1
                self._failed_products.add(target.button.product_key)
                if consecutive >= MAX_CONSECUTIVE_FAILS:
                    self._log(f"❌ 연속 {consecutive}회 실패 → 아이디 작업 중단")
                    self._log(f"📊 [{account.login_id}] 성공 {done} / 실패 {failed}")
                    return "F"
        else:
            self._log(f"⚠ 최대 {MAX_REVIEWS_PER_ACCOUNT}회 도달 → 아이디 종료")

        self._log(f"📊 [{account.login_id}] 성공 {done} / 실패 {failed}")
        return "Y"

    # ─── 단계 2: 리뷰작성 진입 ───────────────────────────────────────────────

    def _open_review_list(self, retry_nav: bool) -> bool:
        ah = _ah()
        for round_no in (1, 2):
            self._set_status("리뷰작성 클릭")
            self._close_popups()
            for i, xp in enumerate(REVIEW_MENU_XPATHS):
                if not ah.element_exists(self.driver, xp, timeout=6 if i == 0 else 2):
                    continue
                if self._click_xpath(xp):
                    self._log("✅ [단계 2] 리뷰작성 클릭 (3초 대기)")
                    time.sleep(3)
                    self._close_popups()
                    return True
            if not retry_nav or round_no == 2:
                break
            self._log("⚠ 리뷰작성 메뉴 미발견 → 마이쇼핑 재진입 후 재시도")
            self._go_main_and_enter_store(login_id="")
        self._log("❌ [단계 2] 리뷰작성 메뉴 진입 실패")
        return False

    # ─── 단계 3: 리뷰 버튼 탐색 + 댓글 매칭(W 예약) ──────────────────────────

    def _scan(self) -> List[ReviewButton]:
        try:
            return parse_review_buttons(self.driver.page_source)
        except ET.ParseError as e:
            self._log(f"  ⚠ 화면 구조 파싱 실패: {e}")
            return []

    @staticmethod
    def _signature(buttons: List[ReviewButton]):
        return tuple((b.text, b.bounds, b.product_key) for b in buttons)

    def _try_claim(self, buttons: List[ReviewButton]) -> Optional[_Target]:
        cands = [b for b in buttons
                 if b.visible and b.texts and b.product_key not in self._failed_products]
        if not cands:
            return None
        idx, row = self.review_manager.claim_comment([b.texts for b in cands])
        if row is None:
            return None
        return _Target(cands[idx], row)

    def _find_and_claim_target(self):
        """
        보이는 리뷰 버튼 중 위에서부터 첫 번째로 댓글 매칭되는 버튼을 찾고 댓글을 W 예약.
        현재 위치 → 맨 위 → 아래로 스크롤하며 탐색.
        Returns: (_Target 또는 None, 리뷰 버튼 존재 여부)
        """
        ah = _ah()
        ah.element_exists(self.driver, REVIEW_BTN_XPATH, timeout=6)  # 목록 로딩 대기
        buttons = self._scan()
        saw = bool(buttons)
        target = self._try_claim(buttons)
        if target:
            return target, True

        # 맨 위로
        sig = self._signature(buttons)
        for _ in range(5):
            self._check_stop()
            self._swipe("up", 0.5)
            buttons = self._scan()
            saw = saw or bool(buttons)
            target = self._try_claim(buttons)
            if target:
                return target, True
            new_sig = self._signature(buttons)
            if new_sig == sig:
                break
            sig = new_sig

        # 아래로
        for _ in range(MAX_LIST_SCROLLS):
            self._check_stop()
            self._swipe("down", 0.35)
            buttons = self._scan()
            saw = saw or bool(buttons)
            target = self._try_claim(buttons)
            if target:
                return target, True
            new_sig = self._signature(buttons)
            if new_sig == sig:
                break
            sig = new_sig
        return None, saw

    # ─── 단계 3~7: 리뷰 1건 작성 ─────────────────────────────────────────────

    def _write_one_review(self, target: _Target) -> bool:
        row, btn = target.row, target.button
        self.current_comment = row
        self._review_no += 1
        self._log(
            f"📝 [리뷰 {self._review_no}] '{btn.text}' 상품: {btn.product[:40]} "
            f"← 검색어 '{row.keyword}' (댓글 row={row.row_index} → W)"
        )
        self._set_status(f"리뷰 작성 중: {row.keyword[:20]}")

        # [단계 3] 리뷰 버튼 클릭
        self._tap(*btn.center, label=btn.text)
        time.sleep(3)
        if not self._wait_review_form(timeout=12):
            self._log("⚠ 리뷰 작성 화면 미진입 → 댓글 예약 해제(공백)")
            self._release_current_comment()
            return False
        self._check_stop()

        # [단계 4-1 ~ 4-4] 별점/만족도
        self._select_ratings()
        self._check_stop()

        # [단계 4-5] 댓글 입력 → Y 기록 → 키보드 숨김 → 스크롤 2회
        if not self._input_review_text(row.comment):
            self._log("❌ [단계 4-5] 리뷰 입력란 입력 실패")
            return self._finish_failed(row)
        self._check_stop()
        self.review_manager.set_comment_status(row.row_index, "Y")
        self._log(f"  ✅ 댓글 입력 완료 → 작업여부 Y (row={row.row_index})")
        self._hide_keyboard()   # 키보드가 화면 하단(등록 버튼)을 가림
        for _ in range(2):
            self._swipe("down", 0.14)

        # [단계 4-6] 등록
        if not self._click_submit(row.comment):
            self._log("❌ [단계 4-6] 등록 버튼 클릭 실패")
            return self._finish_failed(row)
        self.current_comment = None   # 등록 클릭 이후에는 중지해도 예약 해제하지 않음
        time.sleep(3)

        # [단계 5~6] 다음에 동의하기 / 등록 완료 확인
        result = self._wait_review_result(timeout=10)
        if result == "success":
            self._log(f"✅ 리뷰 등록 성공 ('리뷰가등록되었어요' 확인): {btn.product[:40]}")
        elif result == "list":
            self._log(f"✅ 리뷰 등록 성공 (리뷰 목록으로 바로 복귀): {btn.product[:40]}")
        else:
            self._log("❌ [단계 6] 등록 완료 미확인 (리뷰 작성 화면 유지) → 작업여부 F")
            self.review_manager.set_comment_status(row.row_index, "F")
        ok = result is not None
        time.sleep(3)

        # [단계 7] 닫기 — 이미 리뷰 목록이면 닫을 화면이 없음 → 바로 다음 리뷰(단계 3)
        if result == "list":
            self._log("  ℹ 이미 리뷰 목록 화면 → 닫기 생략, 다음 리뷰 진행")
        else:
            self._close_review_page(after_failure=not ok)
        return ok

    def _finish_failed(self, row: CommentRow) -> bool:
        self.review_manager.set_comment_status(row.row_index, "F")
        self.current_comment = None
        self._log(f"  ❌ 작업여부 F 기록 (row={row.row_index})")
        self._close_review_page(after_failure=True)
        return False

    def _handle_alert_ok(self) -> bool:
        """android:id/message 알림창(이어서 작성 등)이 있으면 확인(button1) 클릭"""
        ah = _ah()
        if not ah.element_exists(self.driver, ALERT_MESSAGE_XPATH, timeout=1):
            return False
        try:
            msg = self.driver.find_element(By.XPATH, ALERT_MESSAGE_XPATH).get_attribute("text") or ""
        except Exception:
            msg = ""
        if ah.element_exists(self.driver, ALERT_OK_XPATH, timeout=2) and self._click_xpath(ALERT_OK_XPATH):
            self._log(f"  📌 알림창 '{msg[:40]}' → 확인 클릭")
            time.sleep(2)
            return True
        self._log(f"  ⚠ 알림창 '{msg[:40]}' 확인 버튼 클릭 실패")
        return False

    def _wait_review_form(self, timeout: float = 12) -> bool:
        """리뷰 작성 화면 진입 대기. 알림창이 뜨면 확인 후 계속 대기 (알림창이 떠 있는 동안 화면 요소가 안 잡힘)"""
        end = time.time() + timeout
        while time.time() < end:
            self._check_stop()
            self._handle_alert_ok()
            if _ah().element_exists(self.driver, REVIEW_FORM_XPATH, timeout=1):
                return True
            time.sleep(0.5)
        return False

    def _select_ratings(self):
        """
        [단계 4-1 ~ 4-4] 화면의 모든 RadioGroup 선택 후 조금씩 스크롤.
        별점 그룹은 '5', 그 외 질문 그룹은 항상 마지막 RadioButton (촉촉해요 / 아주좋아요 등).
        리뷰 입력란이 보이고 더 선택할 그룹이 없으면 종료.
        """
        self._handle_alert_ok()   # 알림창이 늦게 뜨는 경우 대비
        self._set_status("별점/만족도 선택")
        done = set()
        prev_sig = None
        idle = 0
        for _ in range(MAX_RATING_ROUNDS):
            self._check_stop()
            try:
                groups, input_visible = parse_radio_groups(self.driver.page_source)
            except ET.ParseError:
                groups, input_visible = [], False

            clicked = 0
            for g in groups:
                if g.key in done or not g.visible:
                    continue
                self._tap(*g.center)
                done.add(g.key)
                clicked += 1
                self._log(f"  ✅ [{g.question[:20] or '질문'}] '{g.label}' 클릭  (선택지: {'/'.join(g.options)})")
                time.sleep(0.7)
            if clicked:
                continue   # 선택 후 새 질문이 펼쳐질 수 있어 같은 위치에서 한 번 더 확인

            if input_visible:
                break
            sig = tuple((g.key, g.center) for g in groups)
            idle = idle + 1 if sig == prev_sig else 0
            if idle >= 2:
                break      # 스크롤해도 화면 변화 없음
            prev_sig = sig
            self._swipe("down", 0.12)
        self._log(f"  ✅ 별점/만족도 {len(done)}개 질문 선택 완료")

    def _input_review_text(self, text: str) -> bool:
        self._set_status("리뷰 내용 입력")
        el = None
        for _ in range(4):
            el = self._find_visible(REVIEW_INPUT_XPATH)
            if el is not None:
                break
            self._swipe("down", 0.14)
        if el is None:
            return False

        try:
            el.click()
            time.sleep(0.8)
            el.send_keys(text)
            time.sleep(1.0)
        except Exception as e:
            self._log(f"  ⚠ send_keys 실패: {str(e).splitlines()[0][:120]}")

        if self._review_text_entered(text):
            return True
        self._log("  ⚠ 입력 미확인 → 클립보드 붙여넣기 재시도")
        self._retype_via_clipboard(text)
        return self._review_text_entered(text)

    def _review_text_entered(self, text: str) -> bool:
        head = _norm(text)[:8]
        try:
            val = self.driver.find_element(By.XPATH, REVIEW_INPUT_XPATH).get_attribute("text") or ""
            if head and head in _norm(val):
                return True
        except Exception:
            pass
        try:
            for e in self.driver.find_elements(By.XPATH, '//*[starts-with(@text,"입력글자수")]'):
                m = re.search(r"입력글자수\s*(\d+)", e.get_attribute("text") or "")
                if m and int(m.group(1)) > 0:
                    return True
        except Exception:
            pass
        return False

    def _retype_via_clipboard(self, text: str):
        try:
            el = self.driver.find_element(By.XPATH, REVIEW_INPUT_XPATH)
            el.click()
            time.sleep(0.5)
            el.clear()
            time.sleep(0.3)
        except Exception:
            pass
        self._paste_clipboard(text)
        time.sleep(1.0)

    def _click_submit(self, text: str) -> bool:
        """등록 버튼 활성화 확인 후 클릭. 비활성이면 키 입력 이벤트/클립보드 재입력으로 활성화 시도."""
        self._set_status("등록 클릭")
        for attempt in range(1, 4):
            self._check_stop()
            self._hide_keyboard()   # 입력 보정/재입력으로 키보드가 다시 열렸을 수 있음
            el = self._find_visible(SUBMIT_BTN_XPATH)
            if el is None:
                self._swipe("down", 0.14)
                continue
            enabled = (el.get_attribute("enabled") or "").lower() == "true"
            if not enabled:
                if attempt == 1:
                    self._log("  ⚠ 등록 버튼 비활성 → 입력 이벤트 보정(스페이스/삭제)")
                    self._nudge_input()
                elif attempt == 2:
                    self._log("  ⚠ 등록 버튼 비활성 → 클립보드 재입력")
                    self._retype_via_clipboard(text)
                    for _ in range(2):
                        self._swipe("down", 0.14)
                time.sleep(1.0)
                continue
            try:
                el.click()
            except Exception:
                b = el.rect
                self._tap(b["x"] + b["width"] // 2, b["y"] + b["height"] // 2, label="등록")
            self._log("  ✅ [단계 4-6] 등록 클릭 (3초 대기)")
            return True
        return False

    def _nudge_input(self):
        """실제 키 이벤트로 웹페이지 입력 상태 갱신 (끝에 공백 입력 후 삭제)"""
        try:
            el = self.driver.find_element(By.XPATH, REVIEW_INPUT_XPATH)
            el.click()
            time.sleep(0.3)
        except Exception:
            pass
        _run_cmd(["adb", "-s", self.device_id, "shell", "input", "keyevent", "123"],   # MOVE_END
                 capture_output=True, timeout=5)
        _run_cmd(["adb", "-s", self.device_id, "shell", "input", "keyevent", "62"],    # SPACE
                 capture_output=True, timeout=5)
        time.sleep(0.3)
        _run_cmd(["adb", "-s", self.device_id, "shell", "input", "keyevent", "67"],    # DEL
                 capture_output=True, timeout=5)
        time.sleep(0.5)

    def _keyboard_shown(self) -> Optional[bool]:
        """키보드 표시 여부. 확인 불가 시 None"""
        try:
            return bool(self.driver.is_keyboard_shown())
        except Exception:
            pass
        try:
            res = _run_cmd(["adb", "-s", self.device_id, "shell", "dumpsys", "input_method"],
                           capture_output=True, text=True, timeout=5)
            out = res.stdout or ""
            if "mInputShown=true" in out:
                return True
            if "mInputShown=false" in out:
                return False
        except Exception:
            pass
        return None

    def _hide_keyboard(self) -> bool:
        """
        댓글 입력 후 키보드 숨김 (ESC 키 우선).
        BACK 키는 키보드가 실제로 떠 있을 때만 사용 (키보드 없이 BACK 하면 리뷰 화면이 닫힘).
        """
        if self._keyboard_shown() is False:
            return True

        # 1) ESC 키
        _run_cmd(["adb", "-s", self.device_id, "shell", "input", "keyevent", "111"],  # ESCAPE
                 capture_output=True, timeout=5)
        time.sleep(0.8)
        shown = self._keyboard_shown()
        if shown is False:
            self._log("  ⌨ 키보드 숨김 (ESC)")
            return True

        # 2) Appium hide_keyboard
        try:
            self.driver.hide_keyboard()
            time.sleep(0.8)
        except Exception:
            pass
        shown = self._keyboard_shown()
        if shown is False:
            self._log("  ⌨ 키보드 숨김 (hide_keyboard)")
            return True

        # 3) BACK 키 — 키보드가 떠 있는 것이 확인된 경우에만
        if shown is True:
            _run_cmd(["adb", "-s", self.device_id, "shell", "input", "keyevent", "4"],
                     capture_output=True, timeout=5)
            time.sleep(0.8)
            if self._keyboard_shown() is False:
                self._log("  ⌨ 키보드 숨김 (BACK)")
                return True
        self._log("  ⚠ 키보드 숨김 확인 실패 → 그대로 진행")
        return False

    def _page_state(self) -> str:
        try:
            return classify_page(self.driver.page_source)
        except Exception:
            return "other"

    def _wait_review_result(self, timeout: float = 10) -> Optional[str]:
        """
        등록 후 결과 확인. '다음에 동의하기' 있으면 클릭.
        - 'success': '리뷰가등록되었어요' 화면 (닫기 필요)
        - 'list'   : 완료 화면 없이 리뷰 목록으로 바로 복귀 (등록 성공, 닫기 불필요)
        - None     : 리뷰 작성 화면에 그대로 남음 등 → 실패
        """
        ah = _ah()
        self._set_status("등록 결과 확인")
        end = time.time() + timeout
        agreed = False
        while time.time() < end:
            self._check_stop()
            if not agreed and ah.element_exists(self.driver, AGREE_LATER_XPATH, timeout=1):
                if self._click_xpath(AGREE_LATER_XPATH):
                    self._log("  📌 [단계 5] '다음에 동의하기' 클릭")
                    agreed = True
                    time.sleep(2)
                    continue
            state = self._page_state()
            if state in ("success", "list"):
                return state
            time.sleep(1)
        return None

    def _close_review_page(self, after_failure: bool):
        ah = _ah()
        self._set_status("닫기")
        if ah.element_exists(self.driver, CLOSE_BTN_XPATH, timeout=5) and self._click_xpath(CLOSE_BTN_XPATH):
            self._log("  📌 [단계 7] 닫기 클릭")
        elif self._page_state() == "list":
            # 리뷰 목록에서 뒤로가기 하면 목록을 벗어나므로 아무것도 하지 않음
            self._log("  ℹ 닫기 버튼 없음 (이미 리뷰 목록 화면) → 닫기 생략")
            return
        else:
            self._log("  ⚠ 닫기 버튼 미발견 → 뒤로가기")
            _run_cmd(["adb", "-s", self.device_id, "shell", "input", "keyevent", "4"],
                     capture_output=True, timeout=5)
        time.sleep(2)
        if after_failure and ah.element_exists(self.driver, LEAVE_CONFIRM_XPATH, timeout=2):
            # 작성 중 나가기 확인 팝업
            if self._click_xpath(LEAVE_CONFIRM_XPATH):
                self._log("  📌 나가기 확인 팝업 → 확인")
                time.sleep(2)

    # ─── 공통 유틸 ───────────────────────────────────────────────────────────

    def _check_stop(self):
        if self._stop_event.is_set():
            raise StopRequested()

    # 자동주문 워커의 _dismiss_popups / _paste_text_via_clipboard 는 버전마다 시그니처가
    # 달라(구버전: 인자 없음 / 미존재) 리뷰 워커는 자체 구현을 사용한다.

    def _close_popups(self, max_count: int = 3):
        """'하루/7일 동안 보지 않기' 팝업 닫기 (ADB 좌표 탭)"""
        union = " | ".join(HIDE_POPUP_XPATHS)
        for _ in range(max_count):
            try:
                els = self.driver.find_elements(By.XPATH, union)
            except Exception:
                return
            if not els:
                return
            try:
                el = els[0]
                txt = el.get_attribute("text") or el.get_attribute("content-desc") or ""
                r = el.rect
                self._log(f"📌 '보지 않기' 팝업 감지 → 탭 ('{txt[:20]}')")
                self._tap(r["x"] + r["width"] // 2, r["y"] + r["height"] // 2)
                time.sleep(1.0)
            except Exception:
                return

    def _paste_clipboard(self, text: str) -> bool:
        """한글 입력용: 클립보드 설정 후 붙여넣기(KEYCODE_PASTE)"""
        try:
            self.driver.set_clipboard_text(text)
        except Exception:
            try:
                self.driver.execute_script("mobile: setClipboard",
                                           {"content": text, "contentType": "plaintext"})
            except Exception as e:
                self._log(f"  ⚠ 클립보드 설정 실패: {str(e).splitlines()[0][:100]}")
                return False
        time.sleep(0.3)
        _run_cmd(["adb", "-s", self.device_id, "shell", "input", "keyevent", "279"],
                 capture_output=True, timeout=5)
        time.sleep(0.6)
        self._log("  ✅ 클립보드 붙여넣기 완료")
        return True

    def _release_current_comment(self):
        row = self.current_comment
        self.current_comment = None
        if not row:
            return
        try:
            if self.review_manager.release_comment(row.row_index):
                self._log(f"  ↩ 미작성 댓글 예약 해제 (row={row.row_index} → 공백)")
        except Exception as e:
            self._log(f"  ⚠ 댓글 예약 해제 실패 (row={row.row_index}): {e}")

    def _content_box(self):
        """WebView 콘텐츠 영역 (상단바/하단 툴바 제외)"""
        try:
            best = None
            for el in self.driver.find_elements(By.XPATH, '//android.webkit.WebView'):
                r = el.rect
                if best is None or r["width"] * r["height"] > best["width"] * best["height"]:
                    best = r
            if best and best["height"] > 200:
                return best["y"], best["y"] + best["height"]
        except Exception:
            pass
        _, h = self._get_window_size()
        return int(h * 0.04), int(h * 0.88)

    def _is_visible_rect(self, rect, box=None) -> bool:
        top, bottom = box or self._content_box()
        cy = rect["y"] + rect["height"] // 2
        return rect["height"] > 10 and top + 10 < cy < bottom - 10

    def _find_visible(self, xpath: str):
        try:
            els = self.driver.find_elements(By.XPATH, xpath)
        except Exception:
            return None
        if not els:
            return None
        box = self._content_box()
        for el in els:
            try:
                if self._is_visible_rect(el.rect, box):
                    return el
            except Exception:
                continue
        return None

    def _click_xpath(self, xpath: str) -> bool:
        try:
            el = self.driver.find_element(By.XPATH, xpath)
        except Exception:
            return False
        try:
            el.click()
            return True
        except Exception:
            try:
                r = el.rect
                return self._tap(r["x"] + r["width"] // 2, r["y"] + r["height"] // 2)
            except Exception:
                return False

    def _tap(self, x: int, y: int, label: str = "") -> bool:
        if label:
            self._log(f"  👉 '{label}' 탭 ({x}, {y})")
        try:
            _run_cmd(["adb", "-s", self.device_id, "shell", "input", "tap", str(x), str(y)],
                     capture_output=True, timeout=5)
            return True
        except Exception:
            return _ah().tap_by_coords(self.driver, x, y, log_callback=self._log)

    def _swipe(self, direction: str, ratio: float):
        """요소 클릭 없는 우측 여백 드래그 스크롤 (down=아래 내용 보기, up=위 내용 보기)"""
        if not self._scroll_gesture(direction, ratio):
            w, h = self._get_window_size()
            x = int(w * 0.88)
            if direction == "down":
                sy = int(h * 0.65)
                ey = max(100, int(sy - h * ratio))
            else:
                sy = int(h * 0.30)
                ey = min(h - 100, int(sy + h * ratio))
            self._adb_swipe(x, sy, x, ey, duration_ms=700)
        time.sleep(0.8)

    # ─── 로그/상태 ────────────────────────────────────────────────────────────

    def _log(self, message: str):
        if self._log_cb:
            self._log_cb(self.device_id, message)
        else:
            print(f"[{self.device_id}] {message}")
        try:
            now = datetime.datetime.now()
            log_dir = os.path.join(_HERE, "logs", now.strftime("%Y%m%d"))
            log_path = os.path.join(log_dir, f"naver_review_기기{self.machine_num}_{self.device_id}.log")
            _LOG_WRITE_QUEUE.put_nowait(
                (log_path, f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] [{self.device_id}] {message}\n")
            )
        except Exception:
            pass
