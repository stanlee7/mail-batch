# -*- coding: utf-8 -*-
"""GUI 스모크 테스트 (offscreen, 발송 없음). python tests/smoke_gui.py"""
import os
import sys
import tempfile

os.environ["QT_QPA_PLATFORM"] = "offscreen"
os.environ["APPDATA"] = tempfile.mkdtemp(prefix="mailbatch_gui_")  # 실제 설정 보호

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication  # noqa: E402

from mailbatch.engine import dry_run_row  # noqa: E402
from mailbatch.gui import MainWindow  # noqa: E402
from tests.make_samples import make  # noqa: E402

FAILED = []


def check(name, cond, detail=""):
    if cond:
        print(f"  ok  {name}")
    else:
        print(f"FAIL  {name}  {detail}")
        FAILED.append(name)


def main():
    xlsx, attach_dir = make()
    app = QApplication(sys.argv)
    win = MainWindow()

    check("탭 3개", win.centralWidget().count() == 3)

    # 엑셀 로드 (파일 다이얼로그 우회)
    lt = win.list_tab
    from mailbatch.engine import read_excel
    lt.excel_edit.setText(xlsx)
    lt.headers, lt.rows = read_excel(xlsx)
    lt.email_combo.clear()
    lt.email_combo.addItems(lt.headers)
    from mailbatch.engine import guess_email_column
    lt.email_combo.setCurrentText(guess_email_column(lt.headers, lt.rows))
    lt.refresh_preview()
    win.on_excel_loaded()

    check("이메일 컬럼 자동선택", lt.email_column() == "이메일", lt.email_column())
    check("미리보기 3행", lt.preview.rowCount() == 3)
    check("자리표시 칩 4개", win.compose_tab.chip_row.count() - 1 == 4,
          win.compose_tab.chip_row.count() - 1)

    # 작성 + dry run이 GUI 설정에서 만들어지는지
    ct = win.compose_tab
    ct.subject_edit.setText("{{이름}}님 수료증")
    ct.body_edit.setPlainText("{{이름}}님, {{과정명}} 수료를 축하합니다.")
    ct.attach_dir_edit.setText(attach_dir)
    ct.pattern_edit.setText("{이름}_수료증.pdf")
    win.send_tab.user_edit.setText("me@naver.com")

    cfg = win.build_config()
    check("cfg 이메일 컬럼", cfg.to_column == "이메일")
    d = dry_run_row(cfg, lt.rows[0], 0)
    check("dry run 제목", d["subject"] == "홍길동님 수료증", d["subject"])
    check("dry run 첨부", len(d["attachments"]) == 1 and "홍길동" in d["attachments"][0])

    # 프리셋 적용
    st_tab = win.send_tab
    st_tab.preset_combo.setCurrentText("지메일")
    check("프리셋 호스트", st_tab.host_edit.text() == "smtp.gmail.com")
    check("프리셋 안내문", "앱 비밀번호" in st_tab.note.text())

    # 설정 저장/복원 왕복
    st_tab.sender_edit.setText("홍 강사")
    st_tab.delay_spin.setValue(3.5)
    win.persist_settings()
    win2 = MainWindow()
    check("설정 복원", win2.send_tab.sender_edit.text() == "홍 강사"
          and win2.send_tab.delay_spin.value() == 3.5
          and win2.send_tab.user_edit.text() == "me@naver.com")
    check("비밀번호 기본 미저장", win2.send_tab.pw_edit.text() == "")

    win.close()
    win2.close()
    app.processEvents()

    print()
    if FAILED:
        print(f"FAILED: {FAILED}")
        sys.exit(1)
    print("ALL PASS")


if __name__ == "__main__":
    main()
