import codecs

content = open('naver_worker.py', 'rb').read().decode('utf-8', errors='replace')

correct_block = '''    def _delete_existing_addresses(self):
        """
        [단계 9.1~9.3] 기본배송지 제외 기존 배송지 모두 삭제
        """
        self._set_status("기존 배송지 삭제 중")
        self._log("🗑 기존 배송지 삭제 시작")

        # 대기 추가: 배송지 관리 화면 로드 대기 (최대 10초)
        ah.wait_for_element(self.driver, NEW_ADDRESS_BTN_XPATH, timeout=10, log_callback=self._log)

        for loop_count in range(DELETE_LOOP_MAX):
            if self._stop_event.is_set():
                break

            time.sleep(1)

            # ListView 내 기본배송지가 아닌 삭제 버튼 탐색
            target_delete_btn = self._find_non_default_delete_button()

            if not target_delete_btn:
                # 스크롤 후 재탐색
                self._scroll_down()
                time.sleep(1.5)
                target_delete_btn = self._find_non_default_delete_button()
                if not target_delete_btn:
                    self._log("✅ 삭제 완료 (더 이상 삭제할 주소 없음)")
                    break

            # [9.1] 삭제 버튼 클릭
            try:
                target_delete_btn.click()
                self._log(f"  삭제 버튼 클릭 (loop {loop_count + 1})")
            except WebDriverException as e:
                self._log(f"  삭제 버튼 click() 실패 ({e}) → 좌표 탭 시도")
                if not self._click_element_center_coordinates(target_delete_btn):
                    continue

            # [9.2] 삭제 확인 팝업에서 OK 버튼 클릭
            if ah.element_exists(self.driver, DELETE_CONFIRM_XPATH, timeout=5):
                ah.wait_and_click(self.driver, DELETE_CONFIRM_XPATH, timeout=5, log_callback=self._log)
                self._log("  삭제 확인 OK 클릭 완료")
            else:
                # android:id/button1 이 없으면 '확인' 텍스트 버튼 시도
                ah.wait_and_click(self.driver, '//android.widget.Button[@text="확인"]',
                                  timeout=3, log_callback=self._log)
            # 삭제 완료 후 WebView 재로딩 대기
            self._wait_webview_ready(label="삭제 후", timeout=10)

        self._log("🗑 삭제 루프 종료")
        # 삭제 루프 전체 완료 후 WebView 완전 로딩 대기
        self._wait_webview_ready(label="삭제 루프 완료", timeout=15)

'''

start_idx = content.find('    def _delete_existing_addresses(self):')
end_idx = content.find('    def _find_non_default_delete_button(self):')

if start_idx != -1 and end_idx != -1:
    new_content = content[:start_idx] + correct_block + content[end_idx:]
    with codecs.open('naver_worker.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print('Fix applied')
else:
    print('Failed to find indices', start_idx, end_idx)
