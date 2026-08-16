"""
main_order.py
네이버 자동 주문 프로그램 - 메인 GUI
naver_address_auto/main.py와 동일한 구조:
  - 좌측: 기기 선택 목록 + ADB 상태 표시 + 체크박스 선택
  - 우측: 선택된 기기별 실시간 로그 패널
  - 상단: 결재목록 현황 (전체/대기/완료/실패)
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import subprocess
import time
import os
import sys
import random
import json
from datetime import datetime

# ─── 경로 설정 ────────────────────────────────────────────────────────────────
_BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
_NAVER_DIR = os.path.join(_BASE_DIR, "naver_address_auto")
if _NAVER_DIR not in sys.path:
    sys.path.insert(0, _NAVER_DIR)

from order_manager import OrderManager
from naver_order_worker import NaverOrderWorker

# ─── 기본 설정 ────────────────────────────────────────────────────────────────
XLSX_PATH           = os.path.join(_BASE_DIR, "개발문서", "결재목록.xlsx")
DEVICES_CONFIG_PATH = os.path.join(_NAVER_DIR, "devices_config.json")
APPIUM_PORT_MIN     = 7723
APPIUM_PORT_MAX     = 8500

# ─── 색상 팔레트 (naver_address_auto/main.py 동일) ───────────────────────────
CLR_BG        = "#0d1117"
CLR_SURFACE   = "#161b22"
CLR_SURFACE2  = "#21262d"
CLR_BORDER    = "#30363d"
CLR_PRIMARY   = "#4ec9b0"
CLR_SUCCESS   = "#3fb950"
CLR_WARNING   = "#d29922"
CLR_ERROR     = "#f85149"
CLR_TEXT      = "#c9d1d9"
CLR_TEXT_MUTE = "#8b949e"
CLR_NAVER     = "#03c75a"


# ─── 기기 로그 패널 ───────────────────────────────────────────────────────────

class DevicePanel(tk.Frame):
    """기기 1대 실시간 로그 패널"""

    def __init__(self, parent, device_id: str, port: int, **kwargs):
        super().__init__(parent, bg=CLR_SURFACE, **kwargs)
        self.device_id = device_id
        self.port = port
        self._build_ui()

    def _build_ui(self):
        hdr = tk.Frame(self, bg=CLR_SURFACE2, padx=10, pady=6)
        hdr.pack(fill=tk.X)

        self.status_dot = tk.Label(hdr, text="●", fg=CLR_TEXT_MUTE,
                                   bg=CLR_SURFACE2, font=("Segoe UI", 11))
        self.status_dot.pack(side=tk.LEFT)

        tk.Label(hdr, text=f"  {self.device_id}",
                 fg=CLR_TEXT, bg=CLR_SURFACE2,
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)

        tk.Label(hdr, text=f"PORT:{self.port}",
                 fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2,
                 font=("Segoe UI", 8)).pack(side=tk.RIGHT)

        self.status_label = tk.Label(
            self, text="대기 중", fg=CLR_TEXT_MUTE,
            bg=CLR_SURFACE, font=("Segoe UI", 9), anchor="w",
            padx=10, pady=4
        )
        self.status_label.pack(fill=tk.X)

        self.log_box = scrolledtext.ScrolledText(
            self, height=12, bg="#0d1117", fg=CLR_TEXT,
            font=("Consolas", 8), relief=tk.FLAT,
            insertbackground=CLR_TEXT, wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.log_box.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        self.log_box.tag_config("success", foreground=CLR_SUCCESS)
        self.log_box.tag_config("error",   foreground=CLR_ERROR)
        self.log_box.tag_config("warning", foreground=CLR_WARNING)
        self.log_box.tag_config("info",    foreground=CLR_PRIMARY)
        self.log_box.tag_config("normal",  foreground=CLR_TEXT)

    def append_log(self, message: str):
        tag = "normal"
        if any(k in message for k in ["✅", "성공", "완료"]):
            tag = "success"
        elif any(k in message for k in ["❌", "실패", "오류"]):
            tag = "error"
        elif any(k in message for k in ["⚠", "타임아웃"]):
            tag = "warning"
        elif any(k in message for k in ["🚀", "📌", "📋", "🔍", "🔐"]):
            tag = "info"

        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END, f"[{ts}] {message}\n", tag)
        self.log_box.see(tk.END)
        self.log_box.config(state=tk.DISABLED)

    def set_status(self, status: str):
        self.status_label.config(text=status)
        if any(k in status for k in ["완료", "성공"]):
            dot_color = CLR_SUCCESS
        elif any(k in status for k in ["실패", "오류"]):
            dot_color = CLR_ERROR
        elif any(k in status for k in ["중", "클릭", "입력", "주문", "선택"]):
            dot_color = CLR_PRIMARY
        elif "대기" in status:
            dot_color = CLR_TEXT_MUTE
        else:
            dot_color = CLR_WARNING
        self.status_dot.config(fg=dot_color)
        self.status_label.config(fg=CLR_TEXT)

    def set_idle(self):
        self.status_dot.config(fg=CLR_TEXT_MUTE)
        self.status_label.config(text="대기 중", fg=CLR_TEXT_MUTE)


# ─── 스크롤 가능 프레임 ───────────────────────────────────────────────────────

class ScrollableFrame(tk.Frame):
    def __init__(self, parent, bg, *args, **kwargs):
        super().__init__(parent, bg=bg, *args, **kwargs)
        self.canvas = tk.Canvas(self, bg=bg, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=bg)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def check_image_exists_on_device(did: str, template_path: str, threshold: float = 0.48) -> bool:
    """ADB 스크린샷 캡처 후 template_path(활성.png)가 화면상에 존재하는지 OpenCV 템플릿 매칭 검사"""
    if not os.path.exists(template_path):
        return False
    try:
        import cv2
        import numpy as np

        res = subprocess.run(["adb", "-s", did, "exec-out", "screencap", "-p"], capture_output=True, timeout=8)
        if not res.stdout or len(res.stdout) < 100:
            return False

        img_array = np.frombuffer(res.stdout, np.uint8)
        screen_img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
        if screen_img is None:
            return False

        # 한글 경로 인코딩 대응 np.fromfile 로드
        template_img = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        if template_img is None:
            template_img = cv2.imread(template_path, cv2.IMREAD_COLOR)
        if template_img is None:
            return False

        result = cv2.matchTemplate(screen_img, template_img, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, _ = cv2.minMaxLoc(result)

        if max_val >= threshold:
            print(f"[ImageMatch:{did}] 템플릿 '{os.path.basename(template_path)}' 인식 성공 (일치율: {max_val:.2f})")
            return True
    except Exception as e:
        print(f"[ImageMatch:{did}] 이미지 인식 예외: {e}")
    return False


def wake_and_keep_screen_on(did: str):
    """ADB 화면 꺼짐/잠금 상태 감시, 화면 자동 깨우기 및 상시 켜짐 유지 (stay_on_while_plugged_in=7)"""
    try:
        # USB 충전 중 화면 상시 켜짐 설정 (7 = AC + USB + Wireless)
        subprocess.run(["adb", "-s", did, "shell", "settings", "put", "global", "stay_on_while_plugged_in", "7"], capture_output=True, timeout=3)
        # 화면 자동 꺼짐 타임아웃 30분(1800000ms)으로 확장
        subprocess.run(["adb", "-s", did, "shell", "settings", "put", "system", "screen_off_timeout", "1800000"], capture_output=True, timeout=3)

        # 화면 꺼짐/어두워짐(Dozing/Asleep) 여부 감시
        res = subprocess.run(["adb", "-s", did, "shell", "dumpsys", "power"], capture_output=True, text=True, timeout=4)
        out = res.stdout or ""
        is_off = ("mWakefulness=Asleep" in out or "mWakefulness=Dozing" in out or 
                  "Display Power: state=OFF" in out or "isScreenOn=false" in out or "mHoldingDisplaySuspendBlocker=false" in out)

        if is_off:
            print(f"[ScreenWake:{did}] ⚡ 화면 꺼짐/어두워짐 감지! 화면 깨우기 및 잠금 해제 전송...")
            subprocess.run(["adb", "-s", did, "shell", "input", "keyevent", "224"], capture_output=True, timeout=3) # WAKEUP
            time.sleep(0.3)
            subprocess.run(["adb", "-s", did, "shell", "input", "keyevent", "82"], capture_output=True, timeout=3)  # UNLOCK
            time.sleep(0.3)
            subprocess.run(["adb", "-s", did, "shell", "input", "keyevent", "3"], capture_output=True, timeout=3)   # HOME
    except Exception as e:
        print(f"[ScreenWake:{did}] 화면 감시 예외: {e}")


# ─── 메인 앱 ──────────────────────────────────────────────────────────────────

class MainApp(tk.Tk):
    """네이버 자동 주문 프로그램 메인 GUI"""

    def __init__(self):
        super().__init__()
        self.title("🛒 네이버 자동 주문 프로그램")
        self.geometry("1400x820")
        self.minsize(1100, 650)
        self.configure(bg=CLR_BG)

        self._xlsx_path = XLSX_PATH
        self.order_manager: OrderManager = None
        self.workers: dict = {}
        self.worker_threads: dict = {}
        self.device_panels: dict = {}
        self.running_ports: set = set()
        self.running: bool = False
        self.worker_semaphore = threading.Semaphore(8)  # CPU 부하 감소를 위해 동시 실행 최대 8대 제한

        self.devices_data = self._load_devices_config()
        self._sync_devices_with_adb()

        self._build_ui()
        self._refresh_summary()

    # ─── 기기 설정 관리 ──────────────────────────────────────────────────────

    def _load_devices_config(self) -> dict:
        if os.path.exists(DEVICES_CONFIG_PATH):
            try:
                with open(DEVICES_CONFIG_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"[MainApp] 설정 로드 오류: {e}")
        return {}

    def _save_devices_config(self):
        try:
            with open(DEVICES_CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self.devices_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[MainApp] 설정 저장 오류: {e}")

    def _sync_devices_with_adb(self):
        """현재 연결된 ADB 기기를 조회하여 로컬 설정과 동기화.

        - 기존 기기: connected 상태만 갱신, selected/remark는 유지
        - 신규 기기: 자동으로 목록에 추가 + selected=True (바로 작업 가능)
        - 연결 끊긴 기기: connected=False, 목록에서 제거하지 않음 (선택 유지)
        """
        connected_devices = []
        try:
            result = subprocess.run(
                ["adb", "devices"],
                capture_output=True, text=True, timeout=5
            )
            lines = result.stdout.strip().splitlines()
            for line in lines[1:]:
                parts = line.split()
                if len(parts) >= 2 and parts[1] == "device":
                    connected_devices.append(parts[0])
        except Exception as e:
            print(f"[MainApp] ADB 조회 실패: {e}")

        # 기존 기기: connected 상태만 False로 초기화 (selected/remark 유지)
        for did in self.devices_data:
            self.devices_data[did]["connected"] = False

        # ADB 연결된 기기 반영
        newly_added = []
        for did in connected_devices:
            if did not in self.devices_data:
                # 신규 기기 → 자동 추가 + 자동 선택
                self.devices_data[did] = {
                    "selected": True,
                    "remark": "",
                    "connected": True
                }
                newly_added.append(did)
            else:
                # 기존 기기 → connected만 갱신, selected/remark는 건드리지 않음
                self.devices_data[did]["connected"] = True

        self._save_devices_config()

        if newly_added:
            print(f"[MainApp] 신규 기기 자동 추가: {newly_added}")

    # ─── 기기 목록 콜백 ──────────────────────────────────────────────────────

    def _update_selected_count_label(self):
        selected_count = sum(1 for d in self.devices_data.values() if d.get("selected", False))
        connected_count = sum(1 for d in self.devices_data.values() if d.get("connected", False))
        if hasattr(self, "dev_header_label"):
            self.dev_header_label.config(
                text=f"📱 기기 선택  ({connected_count}대 연결 / {selected_count}대 선택)"
            )
        if hasattr(self, "select_all_var"):
            all_selected = (
                all(info.get("selected", False) for info in self.devices_data.values())
                if self.devices_data else False
            )
            self.select_all_var.set(all_selected)

    def _on_select_all_toggled(self):
        is_selected = self.select_all_var.get()
        for did in self.devices_data:
            self.devices_data[did]["selected"] = is_selected
            if hasattr(self, "device_check_vars") and did in self.device_check_vars:
                self.device_check_vars[did].set(is_selected)
        self._save_devices_config()
        self._update_selected_count_label()

    def _on_device_select_toggled(self, device_id: str, is_selected: bool):
        if device_id in self.devices_data:
            self.devices_data[device_id]["selected"] = is_selected
        self._save_devices_config()
        self._update_selected_count_label()

    def _toggle_remark_sort(self):
        self.sort_desc = not getattr(self, "sort_desc", False)
        self._draw_device_list()

    def _remark_sort_key(self, item):
        did, info = item
        remark = str(info.get("remark", "")).strip()
        if not remark:
            return (2, 999999, did)
        try:
            return (0, int(remark), did)
        except ValueError:
            return (1, remark.lower(), did)

    def _on_remark_changed(self, device_id: str, remark: str):
        if device_id in self.devices_data:
            if self.devices_data[device_id].get("remark") != remark:
                self.devices_data[device_id]["remark"] = remark
                self._save_devices_config()
                self._draw_device_list()

    def _on_tethering_toggled(self, device_id: str, is_tethering: bool):
        if device_id in self.devices_data:
            self.devices_data[device_id]["tethering"] = is_tethering
        self._save_devices_config()

    def _draw_device_list(self):
        """기기 목록 UI 갱신 (비고 메모 오름차순/내림차순 정렬)"""
        for w in self.scroll_frame.scrollable_frame.winfo_children():
            w.destroy()

        self.device_check_vars = {}

        is_desc = getattr(self, "sort_desc", False)
        sorted_devices = sorted(
            self.devices_data.items(),
            key=self._remark_sort_key,
            reverse=is_desc
        )

        selected_count = sum(1 for d in self.devices_data.values() if d.get("selected", False))
        connected_count = sum(1 for d in self.devices_data.values() if d.get("connected", False))
        self.dev_header_label.config(
            text=f"📱 기기 선택  ({connected_count}대 연결 / {selected_count}대 선택)"
        )

        is_running = self.running

        if hasattr(self, "select_all_var"):
            all_selected = (
                all(info.get("selected", False) for info in self.devices_data.values())
                if self.devices_data else False
            )
            self.select_all_var.set(all_selected)
        if hasattr(self, "select_all_chk"):
            self.select_all_chk.config(state=tk.DISABLED if is_running else tk.NORMAL)

        for idx, (did, info) in enumerate(sorted_devices):
            is_connected = info.get("connected", False)
            is_selected  = info.get("selected",  False)

            row_bg = CLR_SURFACE if idx % 2 == 0 else CLR_BG
            row_frame = tk.Frame(self.scroll_frame.scrollable_frame, bg=row_bg, pady=5)
            row_frame.pack(fill=tk.X)

            # 체크박스
            var = tk.BooleanVar(value=is_selected)
            self.device_check_vars[did] = var
            chk = tk.Checkbutton(
                row_frame, variable=var, bg=row_bg, activebackground=row_bg,
                selectcolor="#ffffff",
                state=tk.DISABLED if is_running else tk.NORMAL,
                command=lambda d=did, v=var: self._on_device_select_toggled(d, v.get())
            )
            chk.pack(side=tk.LEFT, anchor="center", padx=(8, 10))

            # ADB 상태 도트
            dot_color = CLR_SUCCESS if is_connected else CLR_ERROR
            dot_text  = "●" if is_connected else "○"
            tk.Label(row_frame, text=dot_text, fg=dot_color, bg=row_bg,
                     font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(0, 6))

            # 기기 ID
            id_fg = CLR_TEXT if is_connected else CLR_TEXT_MUTE
            id_font_weight = "bold" if is_connected else "normal"
            tk.Label(
                row_frame, text=did,
                fg=id_fg, bg=row_bg,
                font=("Segoe UI", 9, id_font_weight),
                width=14, anchor="w"
            ).pack(side=tk.LEFT)

            # 복사 버튼
            tk.Button(
                row_frame, text="복사",
                command=lambda d=did: self._copy_to_clipboard(d),
                bg=CLR_SURFACE2, fg=CLR_TEXT_MUTE,
                font=("Segoe UI", 8), relief=tk.FLAT, cursor="hand2", padx=4, pady=0
            ).pack(side=tk.LEFT, padx=(0, 6))

            # ADB 상태 텍스트
            status_text  = "연결됨" if is_connected else "미연결"
            status_color = CLR_SUCCESS if is_connected else CLR_TEXT_MUTE
            tk.Label(
                row_frame, text=status_text, fg=status_color, bg=row_bg,
                font=("Segoe UI", 9), width=7, anchor="center"
            ).pack(side=tk.LEFT)

            # 테더링 여부 체크박스 (시안/스카이블루 고대비 색상 적용)
            tether_var = tk.BooleanVar(value=info.get("tethering", True))
            tether_chk = tk.Checkbutton(
                row_frame, text="테더링", variable=tether_var, bg=row_bg, fg="#38bdf8",
                activebackground=row_bg, activeforeground="#38bdf8",
                selectcolor="#ffffff", font=("Segoe UI", 9, "bold"),
                state=tk.DISABLED if is_running else tk.NORMAL,
                command=lambda d=did, v=tether_var: self._on_tethering_toggled(d, v.get())
            )
            tether_chk.pack(side=tk.LEFT, padx=(4, 6))

            # 비고 입력
            remark_var = tk.StringVar(value=info.get("remark", ""))
            entry = tk.Entry(
                row_frame, textvariable=remark_var, bg=CLR_SURFACE2, fg=CLR_TEXT,
                insertbackground=CLR_TEXT, relief=tk.FLAT, font=("Segoe UI", 9),
                state=tk.DISABLED if is_running else tk.NORMAL
            )
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 10))
            entry.bind("<FocusOut>", lambda e, d=did, v=remark_var: self._on_remark_changed(d, v.get()))
            entry.bind("<Return>",   lambda e, d=did, v=remark_var: self._on_remark_changed(d, v.get()))

        self.scroll_frame.canvas.configure(scrollregion=self.scroll_frame.canvas.bbox("all"))

    # ─── UI 구성 ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # 타이틀 바
        title_bar = tk.Frame(self, bg=CLR_SURFACE2, pady=8)
        title_bar.pack(fill=tk.X)

        tk.Label(title_bar, text="🛒  네이버 자동 주문 프로그램",
                 fg=CLR_NAVER, bg=CLR_SURFACE2,
                 font=("Segoe UI", 16, "bold")).pack(side=tk.LEFT, padx=16)

        # 엑셀 경로 표시
        self.xlsx_label = tk.Label(
            title_bar, text=f"📄 {os.path.basename(self._xlsx_path)}",
            fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2, font=("Segoe UI", 9),
            cursor="hand2"
        )
        self.xlsx_label.pack(side=tk.RIGHT, padx=16)
        self.xlsx_label.bind("<Button-1>", lambda e: self._browse_xlsx())

        # 컨트롤 바
        ctrl = tk.Frame(self, bg=CLR_SURFACE, pady=8, padx=12)
        ctrl.pack(fill=tk.X, pady=(2, 0))

        tk.Label(ctrl,
                 text="💡 좌측에서 작업할 기기를 선택하고 시작하세요.",
                 fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
                 font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=10)

        right_ctrl = tk.Frame(ctrl, bg=CLR_SURFACE)
        right_ctrl.pack(side=tk.RIGHT)

        self.adb_btn = self._make_btn(
            right_ctrl, "📱 ADB 조회", self._query_adb_devices,
            fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2
        )
        self.adb_btn.pack(side=tk.LEFT, padx=4)

        self._make_btn(
            right_ctrl, "📂 엑셀 변경", self._browse_xlsx,
            fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2
        ).pack(side=tk.LEFT, padx=4)

        self.reboot_btn = self._make_btn(
            right_ctrl, "🔄 기기 재부팅", self._reboot_selected_devices,
            fg="#ffffff", bg=CLR_WARNING
        )
        self.reboot_btn.pack(side=tk.LEFT, padx=4)

        self.start_btn = self._make_btn(
            right_ctrl, "▶  시작", self._start_all,
            fg="#ffffff", bg=CLR_NAVER
        )
        self.start_btn.pack(side=tk.LEFT, padx=4)

        self.stop_btn = self._make_btn(
            right_ctrl, "⏹  중지", self._stop_all,
            fg="#ffffff", bg=CLR_ERROR, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=4)

        # 요약 바
        summary_bar = tk.Frame(self, bg=CLR_SURFACE2, padx=16, pady=6)
        summary_bar.pack(fill=tk.X)

        self.summary_labels = {}
        for key, label, color in [
            ("total",   "전체",    CLR_TEXT),
            ("pending", "대기",    CLR_WARNING),
            ("done",    "완료(Y)", CLR_SUCCESS),
            ("failed",  "실패(F)", CLR_ERROR),
        ]:
            lf = tk.Frame(summary_bar, bg=CLR_SURFACE2)
            lf.pack(side=tk.LEFT, padx=16)
            tk.Label(lf, text=label + ":", fg=CLR_TEXT_MUTE,
                     bg=CLR_SURFACE2, font=("Segoe UI", 9)).pack(side=tk.LEFT)
            lbl = tk.Label(lf, text="0", fg=color,
                           bg=CLR_SURFACE2, font=("Segoe UI", 11, "bold"))
            lbl.pack(side=tk.LEFT, padx=(3, 0))
            self.summary_labels[key] = lbl

        tk.Button(
            summary_bar, text="↻ 갱신", command=self._refresh_summary,
            bg=CLR_SURFACE, fg=CLR_TEXT_MUTE,
            font=("Segoe UI", 8), relief=tk.FLAT, cursor="hand2",
            padx=6, pady=2
        ).pack(side=tk.RIGHT)

        # 메인 바디 (좌우 분할)
        body_frame = tk.Frame(self, bg=CLR_BG)
        body_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # ── 좌측: 기기 선택 패널 ──────────────────────────────────────────────
        self.dev_list_frame = tk.Frame(
            body_frame, bg=CLR_SURFACE, width=440,
            highlightbackground=CLR_BORDER, highlightthickness=1
        )
        self.dev_list_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        self.dev_list_frame.pack_propagate(False)

        # 기기 목록 헤더
        dev_hdr_frame = tk.Frame(self.dev_list_frame, bg=CLR_SURFACE2, pady=6, padx=10)
        dev_hdr_frame.pack(fill=tk.X)
        self.dev_header_label = tk.Label(
            dev_hdr_frame,
            text="📱 기기 선택  (0대 연결 / 0대 선택)",
            fg=CLR_NAVER, bg=CLR_SURFACE2, font=("Segoe UI", 10, "bold")
        )
        self.dev_header_label.pack(side=tk.LEFT)

        # 컬럼 헤더
        col_hdr = tk.Frame(self.dev_list_frame, bg=CLR_SURFACE, pady=4)
        col_hdr.pack(fill=tk.X)

        self.select_all_var = tk.BooleanVar(value=False)
        self.select_all_chk = tk.Checkbutton(
            col_hdr, variable=self.select_all_var, bg=CLR_SURFACE,
            activebackground=CLR_SURFACE, selectcolor="#ffffff",
            command=self._on_select_all_toggled
        )
        self.select_all_chk.pack(side=tk.LEFT, anchor="center", padx=(8, 10))

        tk.Label(col_hdr, text="ADB", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
                 font=("Segoe UI", 8, "bold"), width=3).pack(side=tk.LEFT)
        tk.Label(col_hdr, text="기기 ID", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
                 font=("Segoe UI", 9, "bold"), width=14, anchor="w").pack(side=tk.LEFT)
        tk.Label(col_hdr, text="상태", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
                 font=("Segoe UI", 9, "bold"), width=7, anchor="center").pack(side=tk.LEFT)
        sort_icon = " ▼" if getattr(self, "sort_desc", False) else " ▲"
        lbl_remark = tk.Label(col_hdr, text=f"비고{sort_icon}", fg=CLR_PRIMARY, bg=CLR_SURFACE,
                              font=("Segoe UI", 9, "bold"), cursor="hand2", anchor="w")
        lbl_remark.pack(side=tk.LEFT, fill=tk.X, expand=True)
        lbl_remark.bind("<Button-1>", lambda e: self._toggle_remark_sort())

        tk.Frame(self.dev_list_frame, bg=CLR_BORDER, height=1).pack(fill=tk.X)

        self.scroll_frame = ScrollableFrame(self.dev_list_frame, bg=CLR_BG)
        self.scroll_frame.pack(fill=tk.BOTH, expand=True)

        # ── 우측: 기기별 실행 패널 ────────────────────────────────────────────
        self.panels_scroll_frame = ScrollableFrame(body_frame, bg=CLR_BG)
        self.panels_scroll_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.panels_frame = self.panels_scroll_frame.scrollable_frame

        self._draw_device_list()
        self._rebuild_device_panels()

        # 하단 상태바
        status_bar = tk.Frame(self, bg=CLR_SURFACE2, pady=3)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.statusbar_label = tk.Label(
            status_bar, text="준비",
            fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2,
            font=("Segoe UI", 8), anchor="w"
        )
        self.statusbar_label.pack(side=tk.LEFT, padx=10)

        tk.Label(
            status_bar, text=f"엑셀: {self._xlsx_path}",
            fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2,
            font=("Segoe UI", 8)
        ).pack(side=tk.RIGHT, padx=10)

    def _make_btn(self, parent, text, command, fg, bg, state=tk.NORMAL):
        btn = tk.Button(
            parent, text=text, command=command,
            fg=fg, bg=bg, activeforeground=fg, activebackground=bg,
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT, cursor="hand2",
            padx=12, pady=6, state=state
        )
        btn.bind("<Enter>", lambda e: btn.config(bg=self._lighten(bg)))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        return btn

    @staticmethod
    def _lighten(hex_color: str) -> str:
        try:
            r = int(hex_color[1:3], 16)
            g = int(hex_color[3:5], 16)
            b = int(hex_color[5:7], 16)
            return f"#{min(255,r+25):02x}{min(255,g+25):02x}{min(255,b+25):02x}"
        except Exception:
            return hex_color

    # ─── 기기 패널 관리 ──────────────────────────────────────────────────────

    def _rebuild_device_panels(self):
        """선택된 기기에 맞게 우측 패널 재구성"""
        for w in self.panels_frame.winfo_children():
            w.destroy()
        self.device_panels.clear()

        selected_devices = self._get_selected_devices()
        count = len(selected_devices)

        if count == 0:
            tk.Label(
                self.panels_frame,
                text="선택된 기기가 없습니다.\n좌측 목록에서 기기를 체크하세요.",
                fg=CLR_TEXT_MUTE, bg=CLR_BG, font=("Segoe UI", 12)
            ).pack(expand=True, pady=60)
            return

        cols = min(count, 3)
        for i, did in enumerate(selected_devices):
            remark = self.devices_data.get(did, {}).get("remark", "")
            display_name = f"{did} ({remark})" if remark else did
            port = self._get_port_for_device(i)

            panel = DevicePanel(
                self.panels_frame, display_name, port,
                relief=tk.FLAT,
                highlightbackground=CLR_BORDER,
                highlightthickness=1
            )
            row = i // cols
            col = i % cols
            panel.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
            self.device_panels[did] = panel

            self.panels_frame.rowconfigure(row, weight=1)
            self.panels_frame.columnconfigure(col, weight=1)

    def _get_selected_devices(self) -> list:
        return [did for did, info in self.devices_data.items()
                if info.get("selected", False)]

    def _get_port_for_device(self, index: int) -> int:
        return APPIUM_PORT_MIN + index

    # ─── 작업 제어 ────────────────────────────────────────────────────────────

    def _start_all(self):
        xlsx = self._xlsx_path
        if not os.path.exists(xlsx):
            messagebox.showerror("오류", f"결재목록.xlsx 파일을 찾을 수 없습니다:\n{xlsx}")
            return

        selected_devices = self._get_selected_devices()
        if not selected_devices:
            messagebox.showwarning("경고", "선택된 기기가 없습니다. 기기를 선택해주세요.")
            return

        # 미연결 기기 경고
        disconnected = [d for d in selected_devices
                        if not self.devices_data[d].get("connected", False)]
        if disconnected:
            confirm = messagebox.askyesno(
                "경고",
                f"다음 기기들은 ADB 미연결 상태입니다:\n" + "\n".join(disconnected) +
                "\n\n계속 진행하시겠습니까?"
            )
            if not confirm:
                return

        self.order_manager = OrderManager(xlsx)
        summary = self.order_manager.get_summary()
        if summary["pending"] == 0:
            messagebox.showinfo("알림", "처리할 미완료 주문이 없습니다.\n결재목록.xlsx를 확인하세요.")
            return

        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self._draw_device_list()
        self._rebuild_device_panels()
        self._log_status("🚀 자동 주문 시작")

        used_ports: set = set()
        for i, did in enumerate(selected_devices):
            port = self._new_random_port(list(used_ports))
            used_ports.add(port)

            worker = NaverOrderWorker(
                device_id      = did,
                appium_port    = port,
                order_manager  = self.order_manager,
                log_callback   = self._on_worker_log,
                status_callback= self._on_worker_status,
            )
            self.workers[did] = worker

            t = threading.Thread(target=self._run_worker, args=(worker,), daemon=True)
            self.worker_threads[did] = t
            t.start()

            if did in self.device_panels:
                self.device_panels[did].append_log(f"🚀 워커 시작 (포트: {port})")

        threading.Thread(target=self._monitor_completion, daemon=True).start()
        self._refresh_summary()

    def _run_worker(self, worker: NaverOrderWorker):
        """워커 실행 래퍼 (스레드에서 호출) - 포트 재시도 포함"""
        did = worker.device_id
        max_retries = 5
        tried_ports: list = []

        self._on_worker_log(did, "⏳ CPU 부하 방지: 실행 대기 중 (최대 8대 동시 실행 제한)")
        with self.worker_semaphore:
            for attempt in range(1, max_retries + 1):
                if worker._stop_event.is_set():
                    self._on_worker_log(did, "⏹ 중지 요청 - 재시도 중단")
                    break

                port = self._new_random_port(tried_ports)
                tried_ports.append(port)
                worker.appium_port = port

                self._on_worker_log(did, f"🔄 [연결 시도 {attempt}/{max_retries}] 포트 {port}")

                try:
                    self._start_appium_server(port)
                    time.sleep(5)

                    if worker.run():
                        break
                    else:
                        self._on_worker_log(did, f"⚠ [{attempt}회] 재시도...")
                except Exception as e:
                    self._on_worker_log(did, f"❌ [{attempt}회] 예외: {str(e)[:120]}")
                finally:
                    self._on_worker_log(did, f"⏹ Appium 종료 (port={port})...")
                    self._kill_process_on_port(port)
                    time.sleep(2)

    def _stop_all(self):
        self._log_status("⏹ 중지 요청 중...")
        for worker in self.workers.values():
            try:
                worker.stop()
            except Exception:
                pass

    def _monitor_completion(self):
        for t in self.worker_threads.values():
            t.join()
        self.after(0, self._on_all_done)

    def _on_all_done(self):
        self.running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.workers.clear()
        self.worker_threads.clear()
        self._draw_device_list()
        self._log_status("✅ 모든 작업 완료")
        self._refresh_summary()
        messagebox.showinfo("완료", "모든 기기의 자동 주문 작업이 완료되었습니다.")

    # ─── 콜백 ─────────────────────────────────────────────────────────────────

    def _on_worker_log(self, device_id: str, message: str):
        def _update():
            if device_id in self.device_panels:
                self.device_panels[device_id].append_log(message)
        self.after(0, _update)

    def _on_worker_status(self, device_id: str, status: str):
        def _update():
            if device_id in self.device_panels:
                self.device_panels[device_id].set_status(status)
            self._refresh_summary()
        self.after(0, _update)

    # ─── ADB 조회 ─────────────────────────────────────────────────────────────

    def _query_adb_devices(self):
        self._log_status("📱 ADB 기기 조회 중...")
        try:
            self._sync_devices_with_adb()
            self._draw_device_list()
            if not self.running:
                self._rebuild_device_panels()

            connected_count = sum(
                1 for d in self.devices_data.values() if d.get("connected", False)
            )
            self._log_status(f"📱 {connected_count}대 기기 연결됨")
            messagebox.showinfo(
                "ADB 기기 조회",
                f"ADB 조회 완료.\n현재 연결된 기기: {connected_count}대"
            )
        except FileNotFoundError:
            messagebox.showerror("오류", "adb 명령을 찾을 수 없습니다.\nADB를 설치하고 PATH에 추가하세요.")
        except subprocess.TimeoutExpired:
            messagebox.showerror("오류", "ADB 명령 타임아웃")

    def _enable_samsung_hotspot_macro(self, did: str) -> bool:
        """갤럭시 S9~S25 설정 앱 매크로 방식으로 모바일 핫스팟 활성화 (부팅 프로그램 대기 후 홈 이동 & 설정 진입)"""
        import xml.etree.ElementTree as ET
        import re

        def get_ui_nodes(filename="ui_hotspot.xml"):
            try:
                subprocess.run(["adb", "-s", did, "shell", "uiautomator", "dump", f"/sdcard/{filename}"],
                               capture_output=True, timeout=10)
                res = subprocess.run(["adb", "-s", did, "shell", "cat", f"/sdcard/{filename}"],
                                     capture_output=True, text=True, timeout=5)
                xml_str = (res.stdout or "").strip()
                if xml_str and "<hierarchy" in xml_str:
                    return ET.fromstring(xml_str)
            except Exception as e:
                self._on_worker_log(did, f"  ⚠ UI 덤프 예외: {e}")
            return None

        def get_center(node):
            bounds = node.attrib.get('bounds', '')
            m = re.match(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', bounds)
            if m:
                x1, y1, x2, y2 = map(int, m.groups())
                return (x1 + x2) // 2, (y1 + y2) // 2
            return None

        def tap_center(cx, cy, label=""):
            self._on_worker_log(did, f"  👆 [매크로 탭] {label} ({cx}, {cy})")
            subprocess.run(["adb", "-s", did, "shell", "input", "tap", str(cx), str(cy)],
                           capture_output=True, timeout=5)

        # 0. 화면 깨우기, HOME 키 이동 및 Wi-Fi 비활성화 (타 시작 앱 이탈)
        self._on_worker_log(did, "🏠 홈 화면 이동, 화면 깨우기 및 Wi-Fi 비활성화...")
        subprocess.run(["adb", "-s", did, "shell", "input", "keyevent", "224"], capture_output=True, timeout=3)
        subprocess.run(["adb", "-s", did, "shell", "input", "keyevent", "82"], capture_output=True, timeout=3)
        subprocess.run(["adb", "-s", did, "shell", "input", "keyevent", "3"], capture_output=True, timeout=3) # HOME 키
        subprocess.run(["adb", "-s", did, "shell", "svc", "wifi", "disable"], capture_output=True, timeout=5)
        time.sleep(2.0)

        # 1. 설정 앱 완전 종료 후 메인 새로 실행
        self._on_worker_log(did, "⚙️ 설정 앱 완전 종료 후 메인 새로 진입...")
        subprocess.run(["adb", "-s", did, "shell", "am", "force-stop", "com.android.settings"], capture_output=True, timeout=3)
        time.sleep(0.5)
        subprocess.run(["adb", "-s", did, "shell", "am", "start", "-n", "com.android.settings/.Settings"],
                       capture_output=True, timeout=5)
        time.sleep(2.5)

        tree = get_ui_nodes("ui_main.xml")

        # Step 1: 설정 메인 화면에서 '연결' 탐색 및 클릭
        if tree is not None:
            conn_target = None
            for node in tree.iter('node'):
                text = node.attrib.get('text', '')
                res_id = node.attrib.get('resource-id', '')
                if (text == "연결" and "title" in res_id) or ("Wi-Fi" in text and "블루투스" in text):
                    conn_target = node
                    break
                if text == "연결":
                    conn_target = node

            if conn_target is not None:
                center = get_center(conn_target)
                if center:
                    tap_center(center[0], center[1], "'연결' 메뉴 클릭")
                    time.sleep(2.0)
                    tree = get_ui_nodes("ui_conn.xml")

        # Step 2: 연결 화면에서 '모바일 핫스팟 및 테더링' 탐색 및 클릭
        if tree is not None:
            tether_menu_target = None
            for node in tree.iter('node'):
                text = node.attrib.get('text', '')
                if text == "모바일 핫스팟 및 테더링" or ("핫스팟" in text and "테더링" in text):
                    tether_menu_target = node
                    break

            if tether_menu_target is not None:
                center = get_center(tether_menu_target)
                if center:
                    tap_center(center[0], center[1], "'모바일 핫스팟 및 테더링' 메뉴 클릭")
                    time.sleep(2.0)
                    tree = get_ui_nodes("ui_tether.xml")

        # 핫스팟 화면 진입 여부 확인 및 보조 인텐트 직행
        has_hotspot_screen = False
        if tree is not None:
            for node in tree.iter('node'):
                if node.attrib.get('text') == "모바일 핫스팟 및 테더링" or node.attrib.get('content-desc') == "모바일 핫스팟":
                    has_hotspot_screen = True
                    break

        if not has_hotspot_screen:
            self._on_worker_log(did, "⚙️ '모바일 핫스팟 및 테더링' 직행 인텐트 시도...")
            subprocess.run(["adb", "-s", did, "shell", "am", "start", "-a", "android.settings.TETHER_SETTINGS"],
                           capture_output=True, timeout=5)
            time.sleep(2.5)
            tree = get_ui_nodes("ui_tether_direct.xml")

        # Step 3: 모바일 핫스팟 스위치 ON 클릭 (활성.png/활성.PNG 이미지 인식 시 즉시 중단 후 다음 단계 진입)
        tpl_path = os.path.join(os.path.dirname(__file__), "활성.PNG")
        if not os.path.exists(tpl_path):
            tpl_path = os.path.join(os.path.dirname(__file__), "활성.png")
        if not os.path.exists(tpl_path):
            tpl_path = os.path.join(os.path.dirname(__file__), "naver_address_auto", "활성.PNG")
        if not os.path.exists(tpl_path):
            tpl_path = os.path.join(os.path.dirname(__file__), "naver_address_auto", "활성.png")

        # 사전 검사: 이미 핫스팟이 켜져있으면 매크로 탭 없이 바로 완료 및 탈출
        if check_image_exists_on_device(did, tpl_path, threshold=0.48):
            self._on_worker_log(did, "✅ [이미지 인식] '활성.PNG' 사전 감지 성공! 핫스팟 이미 활성화됨 -> 반복 탭 중단 및 다음 단계 진행")
            return True

        for attempt in range(1, 11):
            wake_and_keep_screen_on(did)
            # 이미지 인식 검사: 활성.png 이미지가 발견되면 즉시 종료 및 다음 단계 진입
            if check_image_exists_on_device(did, tpl_path, threshold=0.48):
                self._on_worker_log(did, f"✅ [이미지 인식] '활성.PNG' 상태 감지 완료! ({attempt}회차) -> 반복 탭 중단 및 다음 단계 진행")
                return True

            if tree is None or attempt > 1:
                tree = get_ui_nodes(f"ui_tether_retry{attempt}.xml")

            switch_node = None
            title_node = None

            if tree is not None:
                for node in tree.iter('node'):
                    desc = node.attrib.get('content-desc', '')
                    res_id = node.attrib.get('resource-id', '')
                    cls = node.attrib.get('class', '')
                    text = node.attrib.get('text', '')

                    if (desc == "모바일 핫스팟" and ("Switch" in cls or "switch_widget" in res_id)) or \
                       (res_id == "android:id/switch_widget" and desc == "모바일 핫스팟") or \
                       ("Switch" in cls and desc == "모바일 핫스팟"):
                        switch_node = node
                        break
                    if text == "모바일 핫스팟":
                        title_node = node

                if switch_node is None:
                    for node in tree.iter('node'):
                        if "Switch" in node.attrib.get('class', '') or "switch_widget" in node.attrib.get('resource-id', ''):
                            switch_node = node
                            break

            # 이미 활성화(checked="true") 상태인지 체크
            if switch_node is not None and switch_node.attrib.get('checked') == 'true':
                self._on_worker_log(did, f"✅ [XML 검증] 모바일 핫스팟 ON 감지 완료! ({attempt}회차) -> 반복 탭 중단 및 다음 단계 진행")
                return True

            # 스위치 탭 수행
            if switch_node is not None:
                center = get_center(switch_node)
                if center:
                    tap_center(center[0], center[1], f"핫스팟 스위치 ON 탭 ({attempt}/10)")
                else:
                    tap_center(912, 381, f"핫스팟 스위치 우측 좌표 탭 ({attempt}/10)")
            elif title_node is not None:
                center = get_center(title_node)
                cy = center[1] if center else 381
                tap_center(912, cy, f"핫스팟 제목 우측 스위치 영역 탭 ({attempt}/10)")
            else:
                tap_center(912, 381, f"핫스팟 스위치 표준 좌표 탭 ({attempt}/10)")

            # 핫스팟 켜짐 대기 시간 3.5초 (상태바 활성화 아이콘 렌더링 보장)
            time.sleep(3.5)

            # 팝업 (확인 / 켜기 / Turn on) 자동 처리
            pop_tree = get_ui_nodes(f"ui_pop{attempt}.xml")
            if pop_tree is not None:
                for pnode in pop_tree.iter('node'):
                    ptext = pnode.attrib.get('text', '')
                    pdesc = pnode.attrib.get('content-desc', '')
                    if any(k in ptext or k in pdesc for k in ["확인", "켜기", "Turn on", "OK"]):
                        pcenter = get_center(pnode)
                        if pcenter:
                            tap_center(pcenter[0], pcenter[1], f"팝업 '{ptext or pdesc}' 버튼 클릭")
                            time.sleep(2.0)
                            break

            # 탭 수행 후 이미지 인식 재확인
            if check_image_exists_on_device(did, tpl_path, threshold=0.48):
                self._on_worker_log(did, f"✅ [이미지 인식] 탭 후 '활성.PNG' 감지 성공! ({attempt}회차) -> 반복 탭 중단 및 다음 단계 진행")
                return True

            # 탭 후 활성화 여부 다시 재검증
            verify_tree = get_ui_nodes(f"ui_verify{attempt}.xml")
            if verify_tree is not None:
                for node in verify_tree.iter('node'):
                    desc = node.attrib.get('content-desc', '')
                    res_id = node.attrib.get('resource-id', '')
                    checked = node.attrib.get('checked', '')
                    if (desc == "모바일 핫스팟" or "switch_widget" in res_id) and checked == 'true':
                        self._on_worker_log(did, f"✅ [XML 검증] 핫스팟 켜기 성공! ({attempt}회차) -> 반복 탭 중단 및 다음 단계 진행")
                        return True

        self._on_worker_log(did, "⚡ 명령어 보조 실행 (cmd tethering start-tethering wifi)...")
        subprocess.run(["adb", "-s", did, "shell", "cmd", "tethering", "start-tethering", "wifi"],
                       capture_output=True, timeout=5)
        return True

    def _enable_wifi_macro(self, did: str) -> bool:
        """테더링 미체크 설비: 핫스팟 종료 및 Wi-Fi 활성화"""
        self._on_worker_log(did, "🏠 홈 화면 이동 및 테더링 중지...")
        subprocess.run(["adb", "-s", did, "shell", "input", "keyevent", "224"], capture_output=True, timeout=3)
        subprocess.run(["adb", "-s", did, "shell", "input", "keyevent", "82"], capture_output=True, timeout=3)
        subprocess.run(["adb", "-s", did, "shell", "input", "keyevent", "3"], capture_output=True, timeout=3)
        
        # 핫스팟 종료
        subprocess.run(["adb", "-s", did, "shell", "cmd", "tethering", "stop-tethering"], capture_output=True, timeout=5)
        time.sleep(1.0)
        
        # Wi-Fi 활성화
        self._on_worker_log(did, "📶 Wi-Fi 활성화 (svc wifi enable)...")
        subprocess.run(["adb", "-s", did, "shell", "svc", "wifi", "enable"], capture_output=True, timeout=5)
        time.sleep(2.0)
        return True

    def _reboot_and_hotspot(self, did):
        try:
            self._on_worker_log(did, "🔄 ADB 재부팅 명령 전송...")
            self._on_worker_status(did, "재부팅 중")
            subprocess.run(["adb", "-s", did, "reboot"], capture_output=True, timeout=5)
            
            self._on_worker_log(did, "⏳ 기기 부팅 완료 대기 중 (최대 3분)...")
            boot_completed = False
            for _ in range(36):  # 36 * 5초 = 180초 (3분)
                time.sleep(5)
                try:
                    res = subprocess.run(["adb", "-s", did, "shell", "getprop", "sys.boot_completed"], 
                                         capture_output=True, text=True, timeout=3)
                    if "1" in res.stdout:
                        boot_completed = True
                        break
                except Exception:
                    pass
            
            if not boot_completed:
                self._on_worker_log(did, "❌ 부팅 감지 시간 초과")
                self._on_worker_status(did, "부팅 시간초과")
                return
                
            self._on_worker_log(did, "✅ 부팅 완료 감지됨! 타 부팅 앱 대기 (75초, 화면 꺼짐 감시 및 자동 깨우기)...")
            for _ in range(15):  # 15 * 5초 = 75초
                time.sleep(5)
                wake_and_keep_screen_on(did)
            
            is_tethering = self.devices_data.get(did, {}).get("tethering", True)
            if is_tethering:
                self._on_worker_log(did, "📡 [테더링 체크됨] 설정 매크로 방식으로 핫스팟 활성화 시작...")
                self._on_worker_status(did, "핫스팟 매크로 실행")
                self._enable_samsung_hotspot_macro(did)
                self._on_worker_log(did, "✅ 핫스팟 매크로 실행 완료!")
                self._on_worker_status(did, "핫스팟 완료")
            else:
                self._on_worker_log(did, "📶 [테더링 미체크] Wi-Fi 활성화 실행...")
                self._on_worker_status(did, "Wi-Fi 활성화")
                self._enable_wifi_macro(did)
                self._on_worker_log(did, "✅ Wi-Fi 활성화 완료!")
                self._on_worker_status(did, "Wi-Fi 완료")
            
        except Exception as e:
            self._on_worker_log(did, f"❌ 재부팅 예외: {e}")
            self._on_worker_status(did, "예외 발생")

    def _reboot_selected_devices(self):
        selected_devices = self._get_selected_devices()
        if not selected_devices:
            messagebox.showwarning("경고", "선택된 기기가 없습니다. 재부팅할 기기를 선택해주세요.")
            return

        confirm = messagebox.askyesno(
            "재부팅 및 핫스팟 켜기",
            f"선택된 {len(selected_devices)}대의 기기를 재부팅하고 핫스팟을 켜시겠습니까?\n(재부팅 완료 후 자동으로 핫스팟이 실행됩니다.)"
        )
        if not confirm:
            return

        self._log_status(f"🔄 선택된 기기 {len(selected_devices)}대 재부팅 및 핫스팟 연결 작업 시작...")
        
        # 로그 패널 생성을 위해 running 임시 true
        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self._rebuild_device_panels()

        def _reboot_all_task():
            threads = []
            for did in selected_devices:
                t = threading.Thread(target=self._reboot_and_hotspot, args=(did,), daemon=True)
                t.start()
                threads.append(t)
            
            for t in threads:
                t.join()
                
            self.running = False
            self.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self._log_status(f"✅ {len(selected_devices)}대 기기 재부팅 및 핫스팟 처리 완료")
            self.after(0, lambda: messagebox.showinfo("작업 완료", "모든 기기의 재부팅 및 핫스팟 켜기 작업이 완료되었습니다."))

        threading.Thread(target=_reboot_all_task, daemon=True).start()

    # ─── 엑셀 파일 선택 ──────────────────────────────────────────────────────

    def _browse_xlsx(self):
        path = filedialog.askopenfilename(
            title="결재목록.xlsx 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]
        )
        if path:
            self._xlsx_path = path
            self.xlsx_label.config(text=f"📄 {os.path.basename(path)}")
            self._refresh_summary()
            self._log_status(f"📂 엑셀 변경: {os.path.basename(path)}")

    # ─── 현황 요약 갱신 ──────────────────────────────────────────────────────

    def _refresh_summary(self):
        if not os.path.exists(self._xlsx_path):
            return
        try:
            if not self.order_manager or self.order_manager.xlsx_path != self._xlsx_path:
                self.order_manager = OrderManager(self._xlsx_path)
            summary = self.order_manager.get_summary()
            for key, lbl in self.summary_labels.items():
                lbl.config(text=str(summary.get(key, 0)))
        except Exception:
            pass
        # 5초마다 자동 갱신
        self.after(5000, self._refresh_summary)

    def _log_status(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.statusbar_label.config(text=f"[{ts}] {msg}")

    def _copy_to_clipboard(self, text: str):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self._log_status(f"📋 복사됨: {text}")

    # ─── Appium 서버 관리 ────────────────────────────────────────────────────

    def _new_random_port(self, exclude: list = None) -> int:
        exclude = set(exclude or [])
        for _ in range(200):
            port = random.randint(APPIUM_PORT_MIN, APPIUM_PORT_MAX)
            if port not in exclude:
                return port
        return random.randint(APPIUM_PORT_MIN, APPIUM_PORT_MAX)

    def _start_appium_server(self, port: int):
        self._kill_process_on_port(port)
        time.sleep(1)
        self.running_ports.add(port)
        try:
            startupinfo = None
            if os.name == "nt":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0
            subprocess.Popen(
                ["appium", "--port", str(port)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                shell=True, startupinfo=startupinfo
            )
        except Exception as e:
            print(f"Appium 실행 실패 (포트: {port}): {e}")

    def _kill_process_on_port(self, port: int):
        try:
            result = subprocess.run(
                ["cmd", "/c", f"netstat -ano | findstr :{port}"],
                capture_output=True, text=True, timeout=3
            )
            for line in result.stdout.strip().splitlines():
                parts = line.split()
                if len(parts) >= 5:
                    local_addr = parts[1]
                    addr_parts = local_addr.rsplit(":", 1)
                    if len(addr_parts) == 2 and addr_parts[1] == str(port):
                        pid = parts[-1]
                        subprocess.run(["taskkill", "/F", "/PID", pid],
                                       capture_output=True, timeout=3)
        except Exception:
            pass
        finally:
            self.running_ports.discard(port)

    # ─── 종료 ────────────────────────────────────────────────────────────────

    def destroy(self):
        for port in list(self.running_ports):
            self._kill_process_on_port(port)
        super().destroy()


# ─── 진입점 ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = MainApp()
    app.protocol("WM_DELETE_WINDOW", app.destroy)
    app.mainloop()
