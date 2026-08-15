"""
main.py
네이버 앱 자동 주소 등록 프로그램 - 메인 GUI
Tkinter 기반 실시간 모니터링 대시보드
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import subprocess
import time
import os
import sys
import random
from datetime import datetime
import json

from address_manager import AddressManager
from naver_worker import NaverWorker

# ─── 설정 ─────────────────────────────────────────────────────────────────────
XLSX_PATH           = os.path.join(os.path.dirname(__file__), "주소록.xlsx")
DEVICES_CONFIG_PATH = os.path.join(os.path.dirname(__file__), "devices_config.json")
# Appium 포트 랜덤 범위 (7723 포트부터 사용)
APPIUM_PORT_MIN = 7723
APPIUM_PORT_MAX = 8500
# 기능 인수 포트 (systemPort) 범위
SYS_PORT_MIN    = 12000
SYS_PORT_MAX    = 13000


# ─── 색상 팔레트 ──────────────────────────────────────────────────────────────
CLR_BG        = "#0d1117"
CLR_SURFACE   = "#161b22"
CLR_SURFACE2  = "#21262d"
CLR_BORDER    = "#30363d"
CLR_PRIMARY   = "#4ec9b0"     # 네이버 초록 계열 teal
CLR_SUCCESS   = "#3fb950"
CLR_WARNING   = "#d29922"
CLR_ERROR     = "#f85149"
CLR_TEXT      = "#c9d1d9"
CLR_TEXT_MUTE = "#8b949e"
CLR_ACCENT    = "#57ab5a"
CLR_NAVER     = "#03c75a"     # 네이버 시그니처 그린


class DevicePanel(tk.Frame):
    """기기 1대 상태 패널"""

    def __init__(self, parent, device_id: str, port: int, **kwargs):
        super().__init__(parent, bg=CLR_SURFACE, **kwargs)
        self.device_id = device_id
        self.port = port
        self._build_ui()

    def _build_ui(self):
        # 헤더
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

        # 상태 텍스트
        self.status_label = tk.Label(
            self, text="대기 중", fg=CLR_TEXT_MUTE,
            bg=CLR_SURFACE, font=("Segoe UI", 9), anchor="w",
            padx=10, pady=4
        )
        self.status_label.pack(fill=tk.X)

        # 로그 박스
        self.log_box = scrolledtext.ScrolledText(
            self, height=12, bg="#0d1117", fg=CLR_TEXT,
            font=("Consolas", 8), relief=tk.FLAT,
            insertbackground=CLR_TEXT, wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.log_box.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        # 로그 태그
        self.log_box.tag_config("success", foreground=CLR_SUCCESS)
        self.log_box.tag_config("error",   foreground=CLR_ERROR)
        self.log_box.tag_config("warning", foreground=CLR_WARNING)
        self.log_box.tag_config("info",    foreground=CLR_PRIMARY)
        self.log_box.tag_config("normal",  foreground=CLR_TEXT)

    def append_log(self, message: str):
        """로그 메시지 추가"""
        tag = "normal"
        if "✅" in message or "성공" in message or "완료" in message:
            tag = "success"
        elif "❌" in message or "실패" in message or "오류" in message:
            tag = "error"
        elif "⚠" in message or "타임아웃" in message:
            tag = "warning"
        elif "🚀" in message or "📌" in message or "📋" in message:
            tag = "info"

        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END, f"[{ts}] {message}\n", tag)
        self.log_box.see(tk.END)
        self.log_box.config(state=tk.DISABLED)

    def set_status(self, status: str, color: str = None):
        """상태 텍스트 업데이트"""
        self.status_label.config(text=status)
        if color:
            self.status_dot.config(fg=color)
            self.status_label.config(fg=color)
        else:
            if "완료" in status or "성공" in status:
                dot_color = CLR_SUCCESS
            elif "실패" in status or "오류" in status:
                dot_color = CLR_ERROR
            elif "중" in status or "클릭" in status or "입력" in status or "등록" in status:
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


class ScrollableFrame(tk.Frame):
    """Tkinter Canvas + Scrollbar Scrollable Frame"""
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


class MainApp(tk.Tk):
    """메인 애플리케이션 윈도우"""

    def __init__(self):
        super().__init__()
        self.title("🟢 네이버 자동 주소 등록 v1.0")
        self.geometry("1400x820")
        self.minsize(1100, 650)
        self.configure(bg=CLR_BG)

        self.address_manager = AddressManager(XLSX_PATH)
        self.workers: dict = {}
        self.worker_threads: dict = {}
        self.device_panels: dict = {}
        self.running_ports = set()
        self.running = False
        self.worker_semaphore = threading.Semaphore(10)  # CPU 부하 감소를 위해 동시 실행 최대 15개 제한

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
        """현재 연결된 ADB 기기를 조회하여 로컬 설정과 동기화"""
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

        for did in self.devices_data:
            self.devices_data[did]["connected"] = False
            self.devices_data[did]["selected"] = False

        for did in connected_devices:
            if did not in self.devices_data:
                self.devices_data[did] = {
                    "selected": False,
                    "remark": "",
                    "connected": True
                }
            else:
                self.devices_data[did]["connected"] = True
                remark = self.devices_data[did].get("remark", "")
                self.devices_data[did]["selected"] = bool(remark and remark.strip())

        self._save_devices_config()

    # ─── 기기 목록 콜백 ──────────────────────────────────────────────────────

    def _on_select_all_toggled(self):
        is_selected = self.select_all_var.get()
        for did in self.devices_data:
            self.devices_data[did]["selected"] = is_selected
        self._save_devices_config()
        self._draw_device_list()
        if not self.running:
            self._rebuild_device_panels()

    def _on_device_select_toggled(self, device_id: str, is_selected: bool):
        if device_id in self.devices_data:
            self.devices_data[device_id]["selected"] = is_selected
        self._save_devices_config()
        self._draw_device_list()
        if not self.running:
            self._rebuild_device_panels()

    def _on_remark_changed(self, device_id: str, remark: str):
        if device_id in self.devices_data:
            if self.devices_data[device_id].get("remark") != remark:
                self.devices_data[device_id]["remark"] = remark
                self._save_devices_config()
                if not self.running:
                    self._rebuild_device_panels()

    def _draw_device_list(self):
        """기기 목록 UI 갱신"""
        for w in self.scroll_frame.scrollable_frame.winfo_children():
            w.destroy()

        sorted_devices = sorted(
            self.devices_data.items(),
            key=lambda item: (not item[1].get("selected", False), item[0])
        )

        selected_count = sum(1 for d in self.devices_data.values() if d.get("selected", False))
        self.dev_header_label.config(text=f"📱 기기 선택 및 메모 ({selected_count}대 선택됨)")

        is_running = self.running

        if hasattr(self, "select_all_var"):
            all_selected = all(
                info.get("selected", False) for info in self.devices_data.values()
            ) if self.devices_data else False
            self.select_all_var.set(all_selected)
        if hasattr(self, "select_all_chk"):
            self.select_all_chk.config(state=tk.DISABLED if is_running else tk.NORMAL)

        for idx, (did, info) in enumerate(sorted_devices):
            row_bg = CLR_SURFACE if idx % 2 == 0 else CLR_BG
            row_frame = tk.Frame(self.scroll_frame.scrollable_frame, bg=row_bg, pady=6)
            row_frame.pack(fill=tk.X)

            var = tk.BooleanVar(value=info.get("selected", False))
            chk = tk.Checkbutton(
                row_frame, variable=var, bg=row_bg, activebackground=row_bg,
                selectcolor="#ffffff",
                state=tk.DISABLED if is_running else tk.NORMAL,
                command=lambda d=did, v=var: self._on_device_select_toggled(d, v.get())
            )
            chk.pack(side=tk.LEFT, anchor="center", padx=(10, 15))

            id_lbl = tk.Label(
                row_frame, text=did,
                fg=CLR_TEXT if info.get("connected") else CLR_TEXT_MUTE,
                bg=row_bg,
                font=("Segoe UI", 9, "bold" if info.get("connected") else "normal"),
                width=13, anchor="w"
            )
            id_lbl.pack(side=tk.LEFT)

            copy_btn = tk.Button(
                row_frame, text="복사",
                command=lambda d=did: self._copy_to_clipboard(d),
                bg=CLR_SURFACE2, fg=CLR_TEXT_MUTE,
                activebackground=CLR_BORDER, activeforeground=CLR_TEXT,
                font=("Segoe UI", 8), relief=tk.FLAT, cursor="hand2", padx=4, pady=0
            )
            copy_btn.pack(side=tk.LEFT, padx=(0, 6))

            status_text = "연결됨" if info.get("connected") else "미연결"
            status_color = CLR_SUCCESS if info.get("connected") else CLR_TEXT_MUTE
            tk.Label(
                row_frame, text=status_text, fg=status_color, bg=row_bg,
                font=("Segoe UI", 9), width=8, anchor="center"
            ).pack(side=tk.LEFT)

            remark_var = tk.StringVar(value=info.get("remark", ""))
            entry = tk.Entry(
                row_frame, textvariable=remark_var, bg=CLR_SURFACE2, fg=CLR_TEXT,
                insertbackground=CLR_TEXT, relief=tk.FLAT, font=("Segoe UI", 9),
                state=tk.DISABLED if is_running else tk.NORMAL
            )
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
            entry.bind("<FocusOut>", lambda e, d=did, v=remark_var: self._on_remark_changed(d, v.get()))
            entry.bind("<Return>",   lambda e, d=did, v=remark_var: self._on_remark_changed(d, v.get()))

        self.scroll_frame.canvas.configure(scrollregion=self.scroll_frame.canvas.bbox("all"))

    # ─── UI 구성 ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # 타이틀 바
        title_bar = tk.Frame(self, bg=CLR_SURFACE2, pady=8)
        title_bar.pack(fill=tk.X)

        tk.Label(
            title_bar,
            text="🟢  네이버 자동 주소 등록 시스템",
            fg=CLR_NAVER, bg=CLR_SURFACE2,
            font=("Segoe UI", 16, "bold")
        ).pack(side=tk.LEFT, padx=16)

        tk.Label(
            title_bar, text="Powered by Appium",
            fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2,
            font=("Segoe UI", 9)
        ).pack(side=tk.RIGHT, padx=16)

        # 컨트롤 바
        ctrl = tk.Frame(self, bg=CLR_SURFACE, pady=8, padx=12)
        ctrl.pack(fill=tk.X, pady=(2, 0))

        tk.Label(
            ctrl,
            text="💡 왼쪽 기기 목록에서 활성화할 기기를 선택하고 시작하세요.",
            fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
            font=("Segoe UI", 10)
        ).pack(side=tk.LEFT, padx=10)

        right_ctrl = tk.Frame(ctrl, bg=CLR_SURFACE)
        right_ctrl.pack(side=tk.RIGHT)

        self.adb_btn = self._make_btn(
            right_ctrl, "📱 ADB 기기 조회", self._query_adb_devices,
            fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2
        )
        self.adb_btn.pack(side=tk.LEFT, padx=4)

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
            ("done",    "성공(Y)", CLR_SUCCESS),
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
        self.body_frame = tk.Frame(self, bg=CLR_BG)
        self.body_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # 좌측: 기기 설정 패널
        self.dev_list_frame = tk.Frame(
            self.body_frame, bg=CLR_SURFACE, width=430,
            relief=tk.FLAT,
            highlightbackground=CLR_BORDER, highlightthickness=1
        )
        self.dev_list_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        self.dev_list_frame.pack_propagate(False)

        self.dev_header_frame = tk.Frame(self.dev_list_frame, bg=CLR_SURFACE2, pady=6, padx=10)
        self.dev_header_frame.pack(fill=tk.X)

        self.dev_header_label = tk.Label(
            self.dev_header_frame, text="📱 기기 선택 및 메모 (0대 선택됨)",
            fg=CLR_NAVER, bg=CLR_SURFACE2, font=("Segoe UI", 10, "bold")
        )
        self.dev_header_label.pack(side=tk.LEFT)

        headers_frame = tk.Frame(self.dev_list_frame, bg=CLR_SURFACE, pady=4)
        headers_frame.pack(fill=tk.X)

        self.select_all_var = tk.BooleanVar(value=False)
        self.select_all_chk = tk.Checkbutton(
            headers_frame, variable=self.select_all_var, bg=CLR_SURFACE,
            activebackground=CLR_SURFACE, selectcolor="#ffffff",
            command=self._on_select_all_toggled
        )
        self.select_all_chk.pack(side=tk.LEFT, anchor="center", padx=(10, 15))
        tk.Label(headers_frame, text="기기 ID", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
                 font=("Segoe UI", 9, "bold"), width=19, anchor="w").pack(side=tk.LEFT)
        tk.Label(headers_frame, text="상태", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
                 font=("Segoe UI", 9, "bold"), width=8, anchor="center").pack(side=tk.LEFT)
        tk.Label(headers_frame, text="비고 (메모)", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
                 font=("Segoe UI", 9, "bold"), anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Frame(self.dev_list_frame, bg=CLR_BORDER, height=1).pack(fill=tk.X)

        self.scroll_frame = ScrollableFrame(self.dev_list_frame, bg=CLR_BG)
        self.scroll_frame.pack(fill=tk.BOTH, expand=True)

        # 우측: 기기 구동 상태 대시보드
        self.panels_scroll_frame = ScrollableFrame(self.body_frame, bg=CLR_BG)
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
            status_bar, text=f"엑셀: {XLSX_PATH}",
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
            return f"#{min(255, r+25):02x}{min(255, g+25):02x}{min(255, b+25):02x}"
        except Exception:
            return hex_color

    # ─── 기기 패널 관리 ──────────────────────────────────────────────────────

    def _rebuild_device_panels(self):
        """기기 수에 맞게 패널 재구성"""
        for w in self.panels_frame.winfo_children():
            w.destroy()
        self.device_panels.clear()

        selected_devices = self._get_device_ids()
        count = len(selected_devices)

        if count == 0:
            tk.Label(
                self.panels_frame,
                text="선택된 기기가 없습니다.\n좌측 목록에서 기기를 선택하세요.",
                fg=CLR_TEXT_MUTE, bg=CLR_BG, font=("Segoe UI", 12)
            ).pack(expand=True)
            return

        cols = min(count, 3)
        for i, did in enumerate(selected_devices):
            remark = self.devices_data.get(did, {}).get("remark", "")
            display_name = f"{did} ({remark})" if remark else did

            panel = DevicePanel(self.panels_frame, display_name, 0,
                                relief=tk.FLAT,
                                highlightbackground=CLR_BORDER,
                                highlightthickness=1)
            row = i // cols
            col = i % cols
            panel.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
            self.device_panels[did] = panel

            self.panels_frame.rowconfigure(row, weight=1)
            self.panels_frame.columnconfigure(col, weight=1)

    def _get_device_ids(self) -> list:
        return [did for did, info in self.devices_data.items() if info.get("selected", False)]

    # ─── 작업 제어 ────────────────────────────────────────────────────────────

    def _start_all(self):
        if not os.path.exists(XLSX_PATH):
            messagebox.showerror("오류", f"엑셀 파일이 없습니다:\n{XLSX_PATH}")
            return

        selected_devices = self._get_device_ids()
        if not selected_devices:
            messagebox.showwarning("경고", "선택된 기기가 없습니다. 기기를 선택해주세요.")
            return

        disconnected = [d for d in selected_devices if not self.devices_data[d].get("connected")]
        if disconnected:
            confirm = messagebox.askyesno(
                "경고",
                f"다음 기기들은 미연결 상태입니다:\n" + "\n".join(disconnected) +
                "\n\n계속 진행하시겠습니까?"
            )
            if not confirm:
                return

        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self._draw_device_list()
        self._rebuild_device_panels()
        self._log_status("🚀 자동화 시작")

        # 기존 Appium 잔류 서버 정리 (5723~5740 등 이전 포트 다 계속 사용 중일 수 있음)
        self._log_status("🧹 이전 Appium 서버 정리 중...")
        self._cleanup_old_appium_ports()

        # 기기별 제각도 다른 랜덤 포트 할당
        used_ports: set = set()
        for i, did in enumerate(selected_devices):
            port = self._pick_random_appium_port(used_ports)
            used_ports.add(port)
            worker = NaverWorker(
                device_id=did,
                appium_port=port,
                address_manager=self.address_manager,
                log_callback=self._on_worker_log,
                status_callback=self._on_worker_status,
            )
            self.workers[did] = worker
            t = threading.Thread(target=self._run_worker, args=(worker,), daemon=True)
            self.worker_threads[did] = t
            t.start()

            if did in self.device_panels:
                self.device_panels[did].append_log(f"🚀 워커 시작됨 (초기 포트: {port})")

        threading.Thread(target=self._monitor_completion, daemon=True).start()

    def _run_worker(self, worker: NaverWorker):
        """워커 실행 래퍼 (스레드에서 호출)"""
        did = worker.device_id
        max_retries = 5
        success = False
        tried_ports: list = []

        self._on_worker_log(did, "⏳ CPU 부하 방지: 실행 대기 중 (최대 15대 동시 실행 제한)")
        with self.worker_semaphore:
            for attempt in range(1, max_retries + 1):
                if worker._stop_event.is_set():
                    self._on_worker_log(did, "⏹ 중지 요청 감지 - 재시도 중단")
                    break

                # 매번 시도마다 완전히 새로운 랜덤 포트
                port = self._new_random_port(tried_ports)
                tried_ports.append(port)
                worker.appium_port = port

                self._on_worker_log(did, f"🔄 [연결 시도 {attempt}/{max_retries}] 랜덤 포트 {port} 사용")

                try:
                    self._start_appium_server(port)
                    time.sleep(5)  # Appium 완전 기동 대기

                    if worker.run():
                        success = True
                        break
                    else:
                        self._on_worker_log(did, f"⚠ [시도 {attempt}] 다른 포트로 재시도...")
                except Exception as e:
                    self._on_worker_log(did, f"❌ [시도 {attempt}] 예외: {str(e)[:120]}")
                finally:
                    self._on_worker_log(did, f"⏹ Appium 종료 중 (port={port})...")
                    self._kill_process_on_port(port)
                    time.sleep(2)

        if not success and not worker._stop_event.is_set():
            self._on_worker_log(did, f"❌ {max_retries}회 시도 모두 실패")

    def _stop_all(self):
        self._log_status("⏹ 중지 요청 중...")
        for worker in self.workers.values():
            worker.stop()

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
        messagebox.showinfo("완료", "모든 기기의 작업이 완료되었습니다.")

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

            connected_count = sum(1 for d in self.devices_data.values() if d.get("connected"))
            self._log_status(f"📱 {connected_count}대 기기 연결됨")
            messagebox.showinfo("ADB 기기 조회",
                                f"ADB 조회 완료.\n현재 연결된 기기: {connected_count}대")
        except FileNotFoundError:
            messagebox.showerror("오류", "adb 명령이 없습니다.\nADB를 설치하고 PATH에 추가하세요.")
        except subprocess.TimeoutExpired:
            messagebox.showerror("오류", "ADB 명령 타임아웃")

    # ─── 요약 갱신 ────────────────────────────────────────────────────────────

    def _refresh_summary(self):
        try:
            summary = self.address_manager.get_all_rows_summary()
            for key, lbl in self.summary_labels.items():
                lbl.config(text=str(summary.get(key, 0)))
        except Exception:
            pass

    def _log_status(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.statusbar_label.config(text=f"[{ts}] {msg}")

    def _copy_to_clipboard(self, text: str):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self._log_status(f"📋 기기 ID 복사됨: {text}")

    # ─── Appium 서버 관리 ────────────────────────────────────────────────────

    def _new_random_port(self, exclude: list = None) -> int:
        """제외 목록에 없는 새 랜덤 Appium 포트 반환 (6100~6900)"""
        exclude = set(exclude or [])
        for _ in range(200):
            port = random.randint(APPIUM_PORT_MIN, APPIUM_PORT_MAX)
            if port not in exclude:
                return port
        return random.randint(APPIUM_PORT_MIN, APPIUM_PORT_MAX)

    def _pick_random_appium_port(self, exclude: set = None) -> int:
        """사용 중이지 않은 랜덤 Appium 포트 선택 (legacy)"""
        return self._new_random_port(list(exclude) if exclude else [])

    def _cleanup_old_appium_ports(self):
        """이전에 사용했던 알려진 포트 범위의 Appium 프로세스 정리"""
        # 쿠팡 기본 포트 범위 및 이전 실행 잔류 포트 정리
        cleanup_ports = list(range(5720, 5740)) + list(range(9200, 9220))
        for port in cleanup_ports:
            try:
                result = subprocess.run(
                    ["cmd", "/c", f"netstat -ano | findstr :{port}"],
                    capture_output=True, text=True, timeout=2
                )
                for line in result.stdout.strip().splitlines():
                    parts = line.split()
                    if len(parts) >= 5:
                        local_addr = parts[1]
                        addr_parts = local_addr.rsplit(":", 1)
                        if len(addr_parts) == 2 and addr_parts[1] == str(port):
                            pid = parts[-1]
                            subprocess.run(["taskkill", "/F", "/PID", pid],
                                           capture_output=True, timeout=2)
            except Exception:
                pass
        time.sleep(1)

    def _start_appium_server(self, port: int):
        """지정한 포트로 Appium 서버 실행"""
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
        """해당 포트를 사용하는 프로세스 종료"""
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
        except Exception as e:
            print(f"포트 {port} 프로세스 종료 실패: {e}")
        finally:
            self.running_ports.discard(port)

    def destroy(self):
        """창 종료 시 Appium 서버 정리"""
        for port in list(self.running_ports):
            self._kill_process_on_port(port)
        super().destroy()


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
