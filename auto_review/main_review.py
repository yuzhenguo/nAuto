"""
main_review.py
네이버 자동 리뷰 프로그램 - 메인 GUI (main_order.py 와 동일한 구조)
  - 좌측: 기기 선택 목록 + ADB 상태 + 체크박스
  - 우측: 선택된 기기별 실시간 로그 패널 (개별 시작/정지)
  - 기기(폰ID)별 병행 작업, 기기마다 Appium 서버 1개
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
import queue
from datetime import datetime

# ─── 경로 설정 ────────────────────────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
_NAVER_DIR = os.path.join(_ROOT, "naver_address_auto")
for _p in (_HERE, _ROOT, _NAVER_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from review_manager import ReviewManager
from naver_review_worker import NaverReviewWorker

XLSX_PATH           = os.path.join(_HERE, "댓글목록.xlsx")
DEVICES_CONFIG_PATH = os.path.join(_NAVER_DIR, "devices_config.json")  # 비고(remark) 읽기 전용
APPIUM_PORT_MIN     = 7723
APPIUM_PORT_MAX     = 8500

# ─── 색상 팔레트 (main_order.py 동일) ─────────────────────────────────────────
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
CLR_WORKING   = "#38bdf8"


class SysOutQueueWriter:
    """sys.stdout 을 큐에 담아 메인 스레드에서 출력 (스레드 충돌 방지)"""
    def __init__(self, original_out, q):
        self.original_out = original_out
        self.q = q

    def write(self, msg):
        self.q.put((self.original_out, msg))

    def flush(self):
        pass


# ─── 기기 로그 패널 ───────────────────────────────────────────────────────────

class DevicePanel(tk.Frame):
    MAX_LOG_LINES = 500

    def __init__(self, parent, device_id: str, title: str, on_start=None, on_stop=None, **kwargs):
        super().__init__(parent, bg=CLR_SURFACE, **kwargs)
        self.device_id = device_id
        self.on_start = on_start
        self.on_stop = on_stop
        self._running = False

        hdr = tk.Frame(self, bg=CLR_SURFACE2, padx=10, pady=6)
        hdr.pack(fill=tk.X)
        self.status_dot = tk.Label(hdr, text="●", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2, font=("Segoe UI", 11))
        self.status_dot.pack(side=tk.LEFT)
        tk.Label(hdr, text=f"  {title}", fg=CLR_TEXT, bg=CLR_SURFACE2,
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        self.account_label = tk.Label(hdr, text="아이디: -", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
                                      font=("Segoe UI", 8, "bold"), padx=5, pady=1)
        self.account_label.pack(side=tk.LEFT, padx=(8, 0))

        btn_wrap = tk.Frame(hdr, bg=CLR_SURFACE2)
        btn_wrap.pack(side=tk.RIGHT, padx=(4, 0))
        self.stop_btn = tk.Button(
            btn_wrap, text="⏹ 정지", command=self._click_stop,
            bg=CLR_ERROR, fg="#ffffff", font=("Segoe UI", 8, "bold"),
            relief=tk.FLAT, cursor="hand2", padx=6, pady=2, state=tk.DISABLED,
        )
        self.stop_btn.pack(side=tk.RIGHT, padx=2)
        self.start_btn = tk.Button(
            btn_wrap, text="▶ 시작", command=self._click_start,
            bg=CLR_NAVER, fg="#ffffff", font=("Segoe UI", 8, "bold"),
            relief=tk.FLAT, cursor="hand2", padx=6, pady=2,
        )
        self.start_btn.pack(side=tk.RIGHT, padx=2)

        self.status_label = tk.Label(self, text="대기 중", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE,
                                     font=("Segoe UI", 9), anchor="w", padx=10, pady=4)
        self.status_label.pack(fill=tk.X)

        self.log_box = scrolledtext.ScrolledText(
            self, height=12, bg="#0d1117", fg=CLR_TEXT, font=("Consolas", 8),
            relief=tk.FLAT, insertbackground=CLR_TEXT, wrap=tk.WORD, state=tk.DISABLED,
        )
        self.log_box.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        for tag, color in (("success", CLR_SUCCESS), ("error", CLR_ERROR), ("warning", CLR_WARNING),
                           ("info", CLR_PRIMARY), ("normal", CLR_TEXT)):
            self.log_box.tag_config(tag, foreground=color)

    def _click_start(self):
        if self.on_start and not self._running:
            self.on_start(self.device_id)

    def _click_stop(self):
        if self.on_stop and self._running:
            self.on_stop(self.device_id)

    def set_running(self, running: bool):
        self._running = bool(running)
        self.start_btn.config(state=tk.DISABLED if running else tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL if running else tk.DISABLED)

    def append_log(self, message: str):
        tag = "normal"
        if any(k in message for k in ("✅", "성공", "완료")):
            tag = "success"
        elif any(k in message for k in ("❌", "실패", "오류")):
            tag = "error"
        elif any(k in message for k in ("⚠", "타임아웃")):
            tag = "warning"
        elif any(k in message for k in ("🚀", "📌", "📝", "👤", "🔍")):
            tag = "info"
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END, f"[{ts}] {message}\n", tag)
        line_count = int(self.log_box.index(tk.END).split(".")[0]) - 1
        if line_count > self.MAX_LOG_LINES:
            self.log_box.delete("1.0", f"{line_count - self.MAX_LOG_LINES + 1}.0")
        self.log_box.see(tk.END)
        self.log_box.config(state=tk.DISABLED)

    def set_status(self, status: str):
        self.status_label.config(text=status, fg=CLR_TEXT)
        if any(k in status for k in ("완료", "성공")):
            color = CLR_SUCCESS
        elif any(k in status for k in ("실패", "오류")):
            color = CLR_ERROR
        elif "대기" in status:
            color = CLR_TEXT_MUTE
        elif any(k in status for k in ("중", "클릭", "입력", "선택", "탐색", "확인")):
            color = CLR_PRIMARY
        else:
            color = CLR_WARNING
        self.status_dot.config(fg=color)

    def set_account(self, login_id: str):
        v = str(login_id or "").strip()
        if v and v != "-":
            self.account_label.config(text=f"아이디: {v}", fg=CLR_WORKING, bg="#1e293b")
        else:
            self.account_label.config(text="아이디: -", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE)

    def set_idle(self):
        self.status_dot.config(fg=CLR_TEXT_MUTE)
        self.status_label.config(text="대기 중", fg=CLR_TEXT_MUTE)
        self.set_account("-")
        self.set_running(False)


class ScrollableFrame(tk.Frame):
    def __init__(self, parent, bg, *args, **kwargs):
        super().__init__(parent, bg=bg, *args, **kwargs)
        self.canvas = tk.Canvas(self, bg=bg, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=bg)
        self.scrollable_frame.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


# ─── 메인 앱 ──────────────────────────────────────────────────────────────────

class MainApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("📝 네이버 자동 리뷰 프로그램")
        self.geometry("1400x820")
        self.minsize(1100, 650)
        self.configure(bg=CLR_BG)

        self._xlsx_path = XLSX_PATH
        self.review_manager: ReviewManager = None
        self.workers: dict = {}
        self.worker_threads: dict = {}
        self.worker_gen: dict = {}
        self.device_panels: dict = {}
        self.device_check_vars: dict = {}
        self.running_ports: set = set()
        self.devices: dict = {}        # did -> {"selected", "connected", "remark", "seq"}
        self.sheet_devices: dict = {}  # 엑셀 Sheet2 폰ID 정보

        self._log_queue: queue.Queue = queue.Queue()
        self._status_queue: queue.Queue = queue.Queue()
        self._sys_out_queue: queue.Queue = queue.Queue()
        self._orig_stdout = sys.stdout
        self._orig_stderr = sys.stderr
        sys.stdout = SysOutQueueWriter(self._orig_stdout, self._sys_out_queue)
        sys.stderr = SysOutQueueWriter(self._orig_stderr, self._sys_out_queue)
        self._flush_log_queue()

        self._load_sheet_devices(show_error=False)
        self._sync_devices(initial=True)
        self._build_ui()

    # ─── 기기 목록 ────────────────────────────────────────────────────────────

    def _load_sheet_devices(self, show_error: bool = True) -> bool:
        try:
            self.review_manager = ReviewManager(self._xlsx_path)
            self.sheet_devices = self.review_manager.get_device_summary()
            return True
        except Exception as e:
            self.review_manager = None
            self.sheet_devices = {}
            if show_error:
                messagebox.showerror("오류", f"댓글목록 엑셀 로드 실패:\n{self._xlsx_path}\n\n{e}")
            else:
                print(f"[MainApp] 엑셀 로드 실패: {e}")
            return False

    def _adb_connected(self) -> list:
        try:
            res = subprocess.run(["adb", "devices"], capture_output=True, text=True, timeout=5)
            return [ln.split()[0] for ln in res.stdout.strip().splitlines()[1:]
                    if len(ln.split()) >= 2 and ln.split()[1] == "device"]
        except Exception as e:
            print(f"[MainApp] ADB 조회 실패: {e}")
            return []

    def _sync_devices(self, initial: bool = False):
        """ADB 연결 기기 + 엑셀 Sheet2 폰ID 로 기기 목록 구성 (선택 상태 유지)"""
        remarks = {}
        try:
            with open(DEVICES_CONFIG_PATH, "r", encoding="utf-8") as f:
                remarks = {k.upper(): v.get("remark", "") for k, v in json.load(f).items()}
        except Exception:
            pass

        connected = {d.upper(): d for d in self._adb_connected()}
        all_ids = {}
        for key, did in connected.items():
            all_ids[key] = did
        for key in self.sheet_devices:
            all_ids.setdefault(key, key)
        for did in self.devices:
            all_ids.setdefault(did.upper(), did)

        new_devices = {}
        for key, did in all_ids.items():
            prev = self.devices.get(did, {})
            sheet = self.sheet_devices.get(key, {})
            if initial or did not in self.devices:
                selected = bool(sheet.get("has_pending"))
            else:
                selected = prev.get("selected", False)
            new_devices[did] = {
                "selected": selected,
                "connected": key in connected,
                "remark": remarks.get(key, ""),
                "seq": sheet.get("seq", ""),
                "in_sheet": key in self.sheet_devices,
            }
        self.devices = new_devices

    def _sort_key(self, did):
        info = self.devices.get(did, {})
        seq = str(info.get("seq") or "")
        remark = str(info.get("remark") or "")
        return (0 if seq else 1, seq, remark, did)

    def _get_selected_devices(self) -> list:
        return sorted([d for d, i in self.devices.items() if i.get("selected")], key=self._sort_key)

    def _panel_title(self, did: str) -> str:
        info = self.devices.get(did, {})
        parts = [p for p in (info.get("seq"), info.get("remark")) if p]
        return f"{' / '.join(parts)}  {did}" if parts else did

    # ─── UI 구성 ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        title_bar = tk.Frame(self, bg=CLR_SURFACE2, pady=8)
        title_bar.pack(fill=tk.X)
        tk.Label(title_bar, text="📝  네이버 자동 리뷰 프로그램", fg=CLR_NAVER, bg=CLR_SURFACE2,
                 font=("Segoe UI", 16, "bold")).pack(side=tk.LEFT, padx=16)
        self.xlsx_label = tk.Label(title_bar, text=f"📄 {os.path.basename(self._xlsx_path)}",
                                   fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2, font=("Segoe UI", 9), cursor="hand2")
        self.xlsx_label.pack(side=tk.RIGHT, padx=16)
        self.xlsx_label.bind("<Button-1>", lambda e: self._browse_xlsx())

        ctrl = tk.Frame(self, bg=CLR_SURFACE, pady=8, padx=12)
        ctrl.pack(fill=tk.X, pady=(2, 0))
        tk.Label(ctrl, text="💡 좌측에서 기기 선택 후 전체 시작, 또는 각 패널의 ▶시작 / ⏹정지로 개별 제어",
                 fg=CLR_TEXT_MUTE, bg=CLR_SURFACE, font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=10)
        right_ctrl = tk.Frame(ctrl, bg=CLR_SURFACE)
        right_ctrl.pack(side=tk.RIGHT)
        self._make_btn(right_ctrl, "📱 ADB 조회", self._query_adb_devices,
                       fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2).pack(side=tk.LEFT, padx=4)
        self._make_btn(right_ctrl, "📂 엑셀 변경", self._browse_xlsx,
                       fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2).pack(side=tk.LEFT, padx=4)
        self.start_btn = self._make_btn(right_ctrl, "▶  시작", self._start_all, fg="#ffffff", bg=CLR_NAVER)
        self.start_btn.pack(side=tk.LEFT, padx=4)
        self.stop_btn = self._make_btn(right_ctrl, "⏹  중지", self._stop_all,
                                       fg="#ffffff", bg=CLR_ERROR, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=4)

        body = tk.Frame(self, bg=CLR_BG)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.dev_list_frame = tk.Frame(body, bg=CLR_SURFACE, width=470,
                                       highlightbackground=CLR_BORDER, highlightthickness=1)
        self.dev_list_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        self.dev_list_frame.pack_propagate(False)

        hdr = tk.Frame(self.dev_list_frame, bg=CLR_SURFACE2, pady=6, padx=10)
        hdr.pack(fill=tk.X)
        self.dev_header_label = tk.Label(hdr, text="📱 기기 선택", fg=CLR_NAVER, bg=CLR_SURFACE2,
                                         font=("Segoe UI", 10, "bold"))
        self.dev_header_label.pack(side=tk.LEFT)

        col_hdr = tk.Frame(self.dev_list_frame, bg=CLR_SURFACE, pady=4)
        col_hdr.pack(fill=tk.X)
        self.select_all_var = tk.BooleanVar(value=False)
        self.select_all_chk = tk.Checkbutton(col_hdr, variable=self.select_all_var, bg=CLR_SURFACE,
                                             activebackground=CLR_SURFACE, selectcolor="#ffffff",
                                             command=self._on_select_all_toggled)
        self.select_all_chk.pack(side=tk.LEFT, padx=(8, 10))
        for text, width, anchor in (("ADB", 3, "center"), ("기기 ID", 14, "w"), ("상태", 7, "center"),
                                    ("순번", 7, "center"), ("비고", 8, "w")):
            tk.Label(col_hdr, text=text, fg=CLR_TEXT_MUTE, bg=CLR_SURFACE, font=("Segoe UI", 9, "bold"),
                     width=width, anchor=anchor).pack(side=tk.LEFT)
        tk.Frame(self.dev_list_frame, bg=CLR_BORDER, height=1).pack(fill=tk.X)

        self.scroll_frame = ScrollableFrame(self.dev_list_frame, bg=CLR_BG)
        self.scroll_frame.pack(fill=tk.BOTH, expand=True)

        self.panels_scroll = ScrollableFrame(body, bg=CLR_BG)
        self.panels_scroll.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.panels_frame = self.panels_scroll.scrollable_frame

        status_bar = tk.Frame(self, bg=CLR_SURFACE2, pady=3)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        self.statusbar_label = tk.Label(status_bar, text="준비", fg=CLR_TEXT_MUTE, bg=CLR_SURFACE2,
                                        font=("Segoe UI", 8), anchor="w")
        self.statusbar_label.pack(side=tk.LEFT, padx=10)
        self.xlsx_path_label = tk.Label(status_bar, text=f"엑셀: {self._xlsx_path}", fg=CLR_TEXT_MUTE,
                                        bg=CLR_SURFACE2, font=("Segoe UI", 8))
        self.xlsx_path_label.pack(side=tk.RIGHT, padx=10)

        self._draw_device_list()
        self._rebuild_device_panels()

    def _make_btn(self, parent, text, command, fg, bg, state=tk.NORMAL):
        btn = tk.Button(parent, text=text, command=command, fg=fg, bg=bg,
                        activeforeground=fg, activebackground=bg, font=("Segoe UI", 10, "bold"),
                        relief=tk.FLAT, cursor="hand2", padx=12, pady=6, state=state)
        return btn

    def _draw_device_list(self):
        for w in self.scroll_frame.scrollable_frame.winfo_children():
            w.destroy()
        self.device_check_vars = {}
        running = self._any_worker_running()

        sel = sum(1 for i in self.devices.values() if i.get("selected"))
        conn = sum(1 for i in self.devices.values() if i.get("connected"))
        self.dev_header_label.config(text=f"📱 기기 선택  ({conn}대 연결 / {sel}대 선택)")
        self.select_all_var.set(bool(self.devices) and all(i.get("selected") for i in self.devices.values()))
        self.select_all_chk.config(state=tk.DISABLED if running else tk.NORMAL)

        for idx, did in enumerate(sorted(self.devices, key=self._sort_key)):
            info = self.devices[did]
            row_bg = CLR_SURFACE if idx % 2 == 0 else CLR_BG
            row = tk.Frame(self.scroll_frame.scrollable_frame, bg=row_bg, pady=5)
            row.pack(fill=tk.X)

            var = tk.BooleanVar(value=info.get("selected", False))
            self.device_check_vars[did] = var
            tk.Checkbutton(row, variable=var, bg=row_bg, activebackground=row_bg, selectcolor="#ffffff",
                           state=tk.DISABLED if running else tk.NORMAL,
                           command=lambda d=did, v=var: self._on_device_toggled(d, v.get())
                           ).pack(side=tk.LEFT, padx=(8, 10))

            connected = info.get("connected", False)
            tk.Label(row, text="●" if connected else "○", fg=CLR_SUCCESS if connected else CLR_ERROR,
                     bg=row_bg, font=("Segoe UI", 10), width=3).pack(side=tk.LEFT)
            tk.Label(row, text=did, fg=CLR_TEXT if connected else CLR_TEXT_MUTE, bg=row_bg,
                     font=("Segoe UI", 9, "bold" if connected else "normal"),
                     width=14, anchor="w").pack(side=tk.LEFT)
            tk.Label(row, text="연결됨" if connected else "미연결",
                     fg=CLR_SUCCESS if connected else CLR_TEXT_MUTE, bg=row_bg,
                     font=("Segoe UI", 9), width=7).pack(side=tk.LEFT)
            seq_text = info.get("seq") or ("-" if info.get("in_sheet") else "엑셀없음")
            tk.Label(row, text=seq_text, fg=CLR_PRIMARY if info.get("in_sheet") else CLR_TEXT_MUTE,
                     bg=row_bg, font=("Segoe UI", 9, "bold"), width=7).pack(side=tk.LEFT)
            tk.Label(row, text=info.get("remark", ""), fg=CLR_TEXT_MUTE, bg=row_bg,
                     font=("Segoe UI", 9), anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _on_device_toggled(self, did: str, selected: bool):
        if did in self.devices:
            self.devices[did]["selected"] = selected
        self._draw_device_list()
        if not self._any_worker_running():
            self._rebuild_device_panels()

    def _on_select_all_toggled(self):
        v = self.select_all_var.get()
        for info in self.devices.values():
            info["selected"] = v
        self._draw_device_list()
        if not self._any_worker_running():
            self._rebuild_device_panels()

    def _rebuild_device_panels(self):
        for w in self.panels_frame.winfo_children():
            w.destroy()
        self.device_panels.clear()

        selected = self._get_selected_devices()
        if not selected:
            tk.Label(self.panels_frame, text="선택된 기기가 없습니다.\n좌측 목록에서 기기를 체크하세요.",
                     fg=CLR_TEXT_MUTE, bg=CLR_BG, font=("Segoe UI", 12)).pack(expand=True, pady=60)
            return

        cols = min(len(selected), 3)
        for c in range(3):
            self.panels_frame.columnconfigure(c, weight=1 if c < cols else 0)
        for i, did in enumerate(selected):
            panel = DevicePanel(self.panels_frame, did, self._panel_title(did),
                                on_start=self._start_device, on_stop=self._stop_device,
                                relief=tk.FLAT, highlightbackground=CLR_BORDER, highlightthickness=1)
            panel.grid(row=i // cols, column=i % cols, padx=5, pady=5, sticky="nsew")
            self.panels_frame.rowconfigure(i // cols, weight=1)
            self.device_panels[did] = panel
            if self._is_device_running(did):
                panel.set_running(True)
                panel.set_status("실행 중")

    # ─── 작업 제어 ────────────────────────────────────────────────────────────

    def _is_device_running(self, did: str) -> bool:
        t = self.worker_threads.get(did)
        return bool(t and t.is_alive())

    def _any_worker_running(self) -> bool:
        return any(t.is_alive() for t in self.worker_threads.values() if t)

    def _launch_worker(self, did: str, machine_num: int, used_ports: set):
        port = self._new_random_port(used_ports)
        used_ports.add(port)
        gen = self.worker_gen.get(did, 0) + 1
        self.worker_gen[did] = gen

        worker = NaverReviewWorker(
            device_id=did,
            appium_port=port,
            review_manager=self.review_manager,
            log_callback=self._on_worker_log,
            status_callback=self._on_worker_status,
            machine_num=machine_num,
        )
        worker._ui_gen = gen
        self.workers[did] = worker
        t = threading.Thread(target=self._run_worker, args=(worker,), daemon=True)
        self.worker_threads[did] = t
        t.start()
        if did in self.device_panels:
            self.device_panels[did].set_running(True)
            self.device_panels[did].set_status("시작 중")
            self.device_panels[did].append_log(f"🚀 워커 시작 (포트: {port})")

    def _detach_stopped_worker(self, did: str) -> bool:
        """이미 중지 요청된 이전 워커면 분리하고 True. 정상 실행 중이면 False."""
        w = self.workers.get(did)
        if self._is_device_running(did) and not (w is not None and w._stop_event.is_set()):
            return False
        self.workers.pop(did, None)
        self.worker_threads.pop(did, None)
        return True

    def _start_device(self, did: str):
        if not self._detach_stopped_worker(did):
            messagebox.showinfo("알림", f"이미 실행 중입니다:\n{did}")
            return
        if not self.devices.get(did, {}).get("connected"):
            if not messagebox.askyesno("경고", f"{did}\nADB 미연결 상태입니다. 계속할까요?"):
                return
        if not self._load_sheet_devices():
            return
        if not self.sheet_devices.get(did.upper(), {}).get("has_pending"):
            messagebox.showinfo("알림", f"{did}\n엑셀 Sheet2 에 작업할(작업완료여부 공백) 로그인 아이디가 없습니다.")
            return
        selected = self._get_selected_devices()
        machine_num = selected.index(did) + 1 if did in selected else 1
        used = set(self.running_ports) | {w.appium_port for w in self.workers.values()}
        self._launch_worker(did, machine_num, used)
        self._set_running_ui()
        self._log_status(f"▶ 개별 시작: {did}")

    def _start_all(self):
        selected = self._get_selected_devices()
        if not selected:
            messagebox.showwarning("경고", "선택된 기기가 없습니다. 기기를 선택해주세요.")
            return
        if not self._load_sheet_devices():
            return

        targets = [d for d in selected if self.sheet_devices.get(d.upper(), {}).get("has_pending")]
        skipped = [d for d in selected if d not in targets]
        if not targets:
            messagebox.showinfo("알림", "선택된 기기에 작업할(작업완료여부 공백) 로그인 아이디가 없습니다.\n엑셀 Sheet2 를 확인하세요.")
            return
        disconnected = [d for d in targets if not self.devices.get(d, {}).get("connected")]
        if disconnected and not messagebox.askyesno(
                "경고", "다음 기기들은 ADB 미연결 상태입니다:\n" + "\n".join(disconnected) + "\n\n계속 진행하시겠습니까?"):
            return

        self._rebuild_device_panels()
        for d in skipped:
            if d in self.device_panels:
                self.device_panels[d].append_log("ℹ 작업할 로그인 아이디 없음 → 건너뜀")

        used = set(self.running_ports) | {w.appium_port for w in self.workers.values()}
        for i, did in enumerate(selected, start=1):
            if did not in targets:
                continue
            if not self._detach_stopped_worker(did):
                if did in self.device_panels:
                    self.device_panels[did].append_log("ℹ 이미 실행 중 → 건너뜀")
                    self.device_panels[did].set_running(True)
                continue
            self._launch_worker(did, i, used)
            time.sleep(0.05)

        self._set_running_ui()
        self._log_status(f"🚀 자동 리뷰 시작 ({len(targets)}대)")

    def _set_running_ui(self):
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self._draw_device_list()

    def _stop_device(self, did: str):
        worker = self.workers.get(did)
        panel = self.device_panels.get(did)
        if not worker:
            if panel:
                panel.set_idle()
            return
        try:
            worker.stop()
        except Exception:
            pass
        if panel:
            panel.append_log("⏹ 개별 중지 요청됨 (종료 대기 중...)")
            panel.set_status("중지 중...")
            panel.stop_btn.config(state=tk.DISABLED)
        self._log_status(f"⏹ 개별 중지: {did}")
        self.after(15000, lambda d=did, g=worker._ui_gen: self._force_unlock_if_stuck(d, g))

    def _stop_all(self):
        self._log_status("⏹ 전체 중지 요청 중...")
        for did in list(self.workers):
            self._stop_device(did)

    def _force_unlock_if_stuck(self, did: str, gen):
        if self.worker_gen.get(did) != gen:
            return
        if self._is_device_running(did) and did in self.device_panels:
            self.device_panels[did].set_running(False)
            self.device_panels[did].append_log("⚠ 종료 지연 → 시작 버튼 복구 (이전 워커는 백그라운드 정리 중)")
            self.device_panels[did].set_status("대기 중 (재시작 가능)")

    def _sleep_worker(self, worker, seconds: float) -> bool:
        end = time.time() + seconds
        while time.time() < end:
            if worker._stop_event.is_set():
                return False
            time.sleep(0.2)
        return not worker._stop_event.is_set()

    def _run_worker(self, worker: NaverReviewWorker):
        """워커 실행 래퍼 (스레드) - Appium 서버 기동 + 포트 재시도"""
        did = worker.device_id
        gen = worker._ui_gen
        tried_ports = [worker.appium_port]
        port = worker.appium_port
        for attempt in range(1, 6):
            if worker._stop_event.is_set():
                break
            if attempt > 1:
                port = self._new_random_port(tried_ports)
                tried_ports.append(port)
                worker.appium_port = port
            self._on_worker_log(did, f"🔄 [연결 시도 {attempt}/5] 포트 {port}")
            try:
                self._start_appium_server(port)
                if not self._sleep_worker(worker, 5):
                    break
                if worker.run():
                    break
                self._on_worker_log(did, f"⚠ [{attempt}회] 재시도...")
            except Exception as e:
                self._on_worker_log(did, f"❌ [{attempt}회] 예외: {str(e)[:120]}")
            finally:
                self._on_worker_log(did, f"⏹ Appium 종료 (port={port})")
                self._kill_process_on_port(port)
        self.after(0, lambda: self._on_worker_done(did, gen))

    def _on_worker_done(self, did: str, gen):
        if self.worker_gen.get(did) != gen:
            return
        self.workers.pop(did, None)
        self.worker_threads.pop(did, None)
        if did in self.device_panels:
            self.device_panels[did].set_idle()
            self.device_panels[did].append_log("✅ 중지/작업 완료 → 다시 시작 가능")
        if not self._any_worker_running():
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self._load_sheet_devices(show_error=False)
            self._sync_devices()
            self._draw_device_list()
            self._log_status("✅ 실행 중인 기기 없음")

    # ─── 콜백 / 로그 큐 ──────────────────────────────────────────────────────

    def _on_worker_log(self, device_id: str, message: str):
        self._log_queue.put((device_id, message))

    def _on_worker_status(self, device_id: str, status: str, account: str = None):
        self._status_queue.put((device_id, status, account))

    def _flush_log_queue(self):
        try:
            for _ in range(30):
                did, msg = self._log_queue.get_nowait()
                if did in self.device_panels:
                    self.device_panels[did].append_log(msg)
        except queue.Empty:
            pass
        try:
            for _ in range(10):
                did, status, account = self._status_queue.get_nowait()
                panel = self.device_panels.get(did)
                if panel:
                    panel.set_status(status)
                    if account is not None:
                        panel.set_account(account)
        except queue.Empty:
            pass
        n = 0
        try:
            while n < 100:
                out, msg = self._sys_out_queue.get_nowait()
                out.write(msg)
                n += 1
        except queue.Empty:
            pass
        if n:
            self._orig_stdout.flush()
            self._orig_stderr.flush()
        self.after(50, self._flush_log_queue)

    # ─── 버튼 동작 ────────────────────────────────────────────────────────────

    def _query_adb_devices(self):
        self._load_sheet_devices(show_error=False)
        self._sync_devices()
        self._draw_device_list()
        if not self._any_worker_running():
            self._rebuild_device_panels()
        conn = sum(1 for i in self.devices.values() if i.get("connected"))
        self._log_status(f"📱 {conn}대 기기 연결됨")
        messagebox.showinfo("ADB 기기 조회", f"ADB 조회 완료.\n현재 연결된 기기: {conn}대")

    def _browse_xlsx(self):
        if self._any_worker_running():
            messagebox.showwarning("경고", "작업 중에는 엑셀을 변경할 수 없습니다.")
            return
        path = filedialog.askopenfilename(title="댓글목록.xlsx 선택",
                                          filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")])
        if not path:
            return
        self._xlsx_path = path
        self.xlsx_label.config(text=f"📄 {os.path.basename(path)}")
        self.xlsx_path_label.config(text=f"엑셀: {path}")
        self._load_sheet_devices()
        self._sync_devices(initial=True)
        self._draw_device_list()
        self._rebuild_device_panels()
        self._log_status(f"📂 엑셀 변경: {os.path.basename(path)}")

    def _log_status(self, msg: str):
        self.statusbar_label.config(text=f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    # ─── Appium 서버 관리 (main_order.py 동일) ───────────────────────────────

    def _new_random_port(self, exclude=None) -> int:
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
            subprocess.Popen(["appium", "--port", str(port)], stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL, shell=True, startupinfo=startupinfo)
        except Exception as e:
            print(f"Appium 실행 실패 (포트: {port}): {e}")

    def _kill_process_on_port(self, port: int):
        try:
            result = subprocess.run(["cmd", "/c", f"netstat -ano | findstr :{port}"],
                                    capture_output=True, text=True, timeout=3)
            for line in result.stdout.strip().splitlines():
                parts = line.split()
                if len(parts) >= 5:
                    addr = parts[1].rsplit(":", 1)
                    if len(addr) == 2 and addr[1] == str(port):
                        subprocess.run(["taskkill", "/F", "/PID", parts[-1]], capture_output=True, timeout=3)
        except Exception:
            pass
        finally:
            self.running_ports.discard(port)

    def destroy(self):
        for w in list(self.workers.values()):
            try:
                w.stop()
            except Exception:
                pass
        for port in list(self.running_ports):
            self._kill_process_on_port(port)
        super().destroy()


if __name__ == "__main__":
    app = MainApp()
    app.protocol("WM_DELETE_WINDOW", app.destroy)
    app.mainloop()
