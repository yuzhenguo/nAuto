"""
address_manager.py
엑셀(주소록.xlsx) 스레드 안전 읽기/쓰기 관리 모듈
컬럼 구조:
  1열: 수취인 이름
  2열: 주소 검색어
  3열: 우편번호
  4열: 휴대폰 번호 (010-xxxx-xxxx)
  5열: 상세 주소
  6열: 작업 상태 (공백=미작업, Y=성공, F=실패)
  7열: 핸드폰 디바이스 ID
  8열: (예비)
  9열: 주소초기화 여부 (Y=기존 배송지 모두 삭제, 공백=유지)
"""

import threading
import openpyxl
import os
import queue
from typing import Optional


class AddressRow:
    """주소록 단일 행 데이터"""
    def __init__(self, row_index: int, name: str, address_search: str,
                 zipcode: str, phone: str, detail_address: str,
                 status: str, device_id: str, delete_existing: bool = False,
                 naver_id: str = ""):
        self.row_index = row_index            # 엑셀 실제 행 번호 (1-based)
        self.name = name                      # 1열: 수취인
        self.address_search = address_search  # 2열: 주소 검색어
        self.zipcode = str(zipcode).strip() if zipcode else ""   # 3열: 우편번호
        self.phone = str(phone).strip() if phone else ""         # 4열: 휴대폰
        self.detail_address = detail_address  # 5열: 상세주소
        self.status = str(status).strip() if status else ""      # 6열: 상태
        self.device_id = str(device_id).strip() if device_id else ""  # 7열: 기기 ID
        self.delete_existing = delete_existing  # 9열: 기존 주소 삭제 여부
        self.naver_id = str(naver_id).strip() if naver_id else "" # 11열: 네이버아이디

    def get_phone_middle(self) -> str:
        """전화번호 중간 4자리 반환 (앞 3자리 제거)"""
        digits = ''.join(filter(str.isdigit, self.phone))
        if len(digits) >= 11:
            return digits[3:7]   # 01012345678 → 1234
        elif len(digits) >= 8:
            return digits[3:7]
        return ""

    def get_phone_last(self) -> str:
        """전화번호 마지막 4자리 반환"""
        digits = ''.join(filter(str.isdigit, self.phone))
        if len(digits) >= 8:
            return digits[-4:]
        return ""

    def __repr__(self):
        return (f"AddressRow(row={self.row_index}, name={self.name}, "
                f"zipcode={self.zipcode}, device={self.device_id}, status={self.status}, "
                f"delete_existing={self.delete_existing})")


class AddressManager:
    """
    스레드 안전한 엑셀 주소록 관리자 (초고속 인메모리 + 큐 기반 백그라운드 저장).
    여러 워커 스레드가 동시에 접근해도 안전하게 처리.
    """

    def __init__(self, xlsx_path: str):
        self.xlsx_path = xlsx_path
        self._lock = threading.Lock()
        self._memory_rows = []
        self._write_queue = queue.Queue()
        
        self._load_into_memory()
        
        # 엑셀 저장을 전담하는 백그라운드 큐 스레드 시작
        self._writer_thread = threading.Thread(target=self._excel_writer_loop, daemon=True)
        self._writer_thread.start()

    def _load_into_memory(self):
        """프로그램 시작 시 엑셀 데이터를 메모리로 한 번에 로드"""
        with self._lock:
            self._memory_rows.clear()
            try:
                wb = openpyxl.load_workbook(self.xlsx_path, data_only=True)
                ws = wb.active
                for row_idx in range(2, ws.max_row + 1):  # 2행부터 시작
                    name = ws.cell(row_idx, 1).value
                    # 1열(이름)이 비어있으면 데이터 끝
                    if not name or str(name).strip() == "":
                        break

                    status = ws.cell(row_idx, 6).value
                    dev = ws.cell(row_idx, 7).value

                    status_str = str(status).strip() if status else ""
                    dev_str = str(dev).strip() if dev else ""
                    addr_search = ws.cell(row_idx, 2).value
                    zipcode = ws.cell(row_idx, 3).value
                    phone = ws.cell(row_idx, 4).value
                    detail = ws.cell(row_idx, 5).value

                    # 9열 (주소초기화 여부)
                    delete_val = ws.cell(row_idx, 9).value
                    delete_existing = (str(delete_val).strip().upper() == "Y") if delete_val else False
                    naver_id = ws.cell(row_idx, 11).value

                    self._memory_rows.append(AddressRow(
                        row_index=row_idx,
                        name=str(name).strip(),
                        address_search=str(addr_search).strip() if addr_search else "",
                        zipcode=zipcode,
                        phone=phone,
                        detail_address=str(detail).strip() if detail else "",
                        status=status_str,
                        device_id=dev_str,
                        delete_existing=delete_existing,
                        naver_id=str(naver_id).strip() if naver_id else ""
                    ))
            except Exception as e:
                print(f"[AddressManager] 초기 엑셀 로드 오류: {e}")

    def get_pending_rows_for_device(self, device_id: str) -> list:
        """
        특정 기기 ID에 해당하며 6열이 공백인(미처리) 행 목록 반환 (인메모리 초고속 반환)
        device_id가 비어있으면 기기 ID 무관하게 공백 행 반환
        """
        with self._lock:
            rows = []
            for row in self._memory_rows:
                if row.status == "" and (device_id == "" or row.device_id == device_id):
                    rows.append(row)
            return rows

    def get_next_pending_row(self, device_id: str) -> Optional[AddressRow]:
        """기기 ID에 해당하는 미처리 행 중 첫 번째를 원자적으로 예약(W)하여 반환"""
        with self._lock:
            for row in self._memory_rows:
                if row.status == "" and (device_id == "" or row.device_id == device_id):
                    # 즉시 'W'(작업중) 상태로 변경하여 다른 기기/스레드가 선점하지 못하도록 보호 (Lock-free 병행 최적화)
                    row.status = "W"
                    self._write_queue.put((row.row_index, "W"))
                    return row
            return None

    def mark_success(self, row_index: int):
        """해당 행의 6열을 Y로 업데이트"""
        self._update_status(row_index, "Y")

    def mark_failed(self, row_index: int):
        """해당 행의 6열을 F로 업데이트"""
        self._update_status(row_index, "F")

    def _update_status(self, row_index: int, status: str):
        """6열 상태값 업데이트 (메모리 반영 후 백그라운드 큐 전송)"""
        with self._lock:
            for row in self._memory_rows:
                if row.row_index == row_index:
                    row.status = status
                    self._write_queue.put((row_index, status))
                    break

    def has_pending_rows(self, device_id: str) -> bool:
        """미처리 행이 남아있는지 확인"""
        return len(self.get_pending_rows_for_device(device_id)) > 0

    def get_all_rows_summary(self) -> dict:
        """전체 현황 요약 반환 (GUI 표시용 - 메모리 집계)"""
        with self._lock:
            summary = {"total": 0, "done": 0, "failed": 0, "pending": 0}
            for row in self._memory_rows:
                summary["total"] += 1
                st = str(row.status).strip().upper()
                if st == "Y":
                    summary["done"] += 1
                elif st == "F":
                    summary["failed"] += 1
                else:
                    summary["pending"] += 1
            return summary

    def get_all_devices_task_counts(self) -> dict:
        """
        기기ID별 작업 현황 집계 반환 (스레드 안전, 인메모리)
        """
        with self._lock:
            counts = {}
            for row in self._memory_rows:
                dev_key = str(row.device_id).strip().upper() if row.device_id else ""
                if dev_key not in counts:
                    counts[dev_key] = {"total": 0, "pending": 0, "done": 0, "failed": 0}

                counts[dev_key]["total"] += 1
                st = str(row.status).strip().upper()
                if st == "Y":
                    counts[dev_key]["done"] += 1
                elif st == "F":
                    counts[dev_key]["failed"] += 1
                else:
                    counts[dev_key]["pending"] += 1
            return counts

    def get_device_task_counts(self, device_id: str) -> dict:
        """특정 기기ID의 {total, pending, done, failed} 반환"""
        norm_id = str(device_id).strip().upper() if device_id else ""
        all_counts = self.get_all_devices_task_counts()
        if norm_id in all_counts:
            return all_counts[norm_id]
        for k, v in all_counts.items():
            if k == norm_id:
                return v
        return {"total": 0, "pending": 0, "done": 0, "failed": 0}

    def _excel_writer_loop(self):
        """백그라운드에서 큐에 쌓인 상태 업데이트를 일괄(Batch)로 엑셀에 저장"""
        while True:
            updates = []
            updates.append(self._write_queue.get())
            while not self._write_queue.empty():
                try:
                    updates.append(self._write_queue.get_nowait())
                except queue.Empty:
                    break
            
            try:
                # 단 한 번만 엑셀을 열어서 모든 변경사항을 반영
                wb = openpyxl.load_workbook(self.xlsx_path)
                ws = wb.active
                for (row_idx, status) in updates:
                    ws.cell(row_idx, 6).value = status
                wb.save(self.xlsx_path)
            except Exception as e:
                print(f"[AddressManager] 백그라운드 엑셀 저장 오류: {e}")
