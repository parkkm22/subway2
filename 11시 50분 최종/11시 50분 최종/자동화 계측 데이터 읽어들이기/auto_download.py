from pywinauto.application import Application
from pywinauto.timings import wait_until_passes
import time

# 1. EXE 실행
app = Application(backend="uia").start(r"도림.exe")

# 2. 메인 창 연결
main = app.window(title="DataSniper For SDL1610B(Client)")
main.wait('visible', timeout=15)

# 3. "연결 하기" 버튼 클릭
main.child_window(title="연결 하기", control_type="Button").click()
time.sleep(1)

# 4. "로그인 정보 등록" 팝업 → "로그인" 클릭
login_popup = app.window(title="로그인 정보 등록")
login_popup.wait('visible', timeout=10)
login_popup.child_window(title="로그인", control_type="Button").click()
time.sleep(1)

# 5. "데이터 보기" 버튼 클릭
wait_until_passes(10, 1, lambda: main.child_window(title="데이터 보기", control_type="Button").is_enabled())
main.child_window(title="데이터 보기", control_type="Button").click()
time.sleep(2)  # 여유 시간 줌

# 데이터 보기 팝업 창 접근
data_view_dialog = main.child_window(title="데이터 보기", control_type="Window")
data_view_dialog.wait('visible', timeout=10)

# 엑셀파일로 저장 Pane 클릭 시도
excel_save_pane = data_view_dialog.child_window(title="조회", control_type="Button")
excel_save_pane.click_input()
time.sleep(3)  # 저장 팝업 대기

# 엑셀파일로 저장 Pane 클릭 시도
excel_save_pane = data_view_dialog.child_window(title="엑셀파일로 저장", control_type="Pane")
excel_save_pane.click_input()
time.sleep(3)  # 저장 팝업 대기


# 창 접근
save_dialog = app.window(title="데이터 저장")
save_dialog.wait('visible', timeout=10)


save_button = save_dialog.child_window(title="저장(S)", control_type="Button")
save_button.click_input()