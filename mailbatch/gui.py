# -*- coding: utf-8 -*-
"""메일 배치 GUI (PySide6).

탭 1: 명단 — 엑셀 로드, 이메일 컬럼 선택
탭 2: 메일 작성 — 제목/본문 {{컬럼}} 치환, 공통/개인별 첨부
탭 3: 발송 — SMTP 설정, 테스트 발송, 전체 발송, 실패 재발송
"""
import os
import sys

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDoubleSpinBox, QFileDialog,
    QFormLayout, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
    QListWidget, QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar,
    QPushButton, QSpinBox, QTabWidget, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from mailbatch import settings as st  # noqa: E402
from mailbatch.engine import (  # noqa: E402
    AuthError, SendConfig, dry_run_row, export_results,
    find_unknown_placeholders, guess_email_column, read_excel, send_batch,
    validate_email,
)
from mailbatch.providers import PRESETS  # noqa: E402

APP_TITLE = "메일 배치 — 엑셀 명단 개인별 일괄발송"

MISSING_POLICIES = [("실패로 처리", "fail"), ("행 건너뜀", "skip_row"), ("첨부 없이 발송", "send_without")]
SECURITY_ITEMS = [("SSL (465)", "ssl"), ("STARTTLS (587)", "starttls")]


def resource_path(name: str):
    """개발/PyInstaller 양쪽에서 리소스 파일 경로를 찾는다."""
    candidates = []
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(sys.executable), name))
        if hasattr(sys, "_MEIPASS"):
            candidates.append(os.path.join(sys._MEIPASS, name))
    candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", name))
    for c in candidates:
        if os.path.exists(c):
            return os.path.abspath(c)
    return None


def pick_file(parent, label, filt):
    path, _ = QFileDialog.getOpenFileName(parent, label, "", filt)
    return path


def pick_dir(parent, label):
    return QFileDialog.getExistingDirectory(parent, label)


class SendWorker(QThread):
    progressed = Signal(int, int, str)
    done = Signal(list)
    failed = Signal(str)

    def __init__(self, cfg, rows, to_override=None, parent=None):
        super().__init__(parent)
        self.cfg, self.rows, self.to_override = cfg, rows, to_override
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            res = send_batch(
                self.cfg, self.rows, to_override=self.to_override,
                progress=lambda d, t, n: self.progressed.emit(d, t, n),
                cancel=lambda: self._cancel,
            )
            self.done.emit(res)
        except AuthError as e:
            self.failed.emit(str(e))
        except Exception as e:
            self.failed.emit(f"발송 오류: {e}")


class ListTab(QWidget):
    """탭 1: 엑셀 명단 로드 + 이메일 컬럼 선택."""

    def __init__(self, on_loaded=None):
        super().__init__()
        self.headers, self.rows = [], []
        self.on_loaded = on_loaded
        root = QVBoxLayout(self)

        box = QGroupBox("엑셀 명단 (첫 행 = 컬럼명)")
        grid = QGridLayout(box)
        self.excel_edit = QLineEdit()
        self.excel_edit.setPlaceholderText("예: 명단.xlsx — 이름/이메일/과정명 등의 컬럼")
        btn = QPushButton("찾아보기…")
        btn.clicked.connect(self.choose_excel)
        grid.addWidget(QLabel("엑셀"), 0, 0)
        grid.addWidget(self.excel_edit, 0, 1)
        grid.addWidget(btn, 0, 2)
        grid.addWidget(QLabel("이메일 컬럼"), 1, 0)
        self.email_combo = QComboBox()
        grid.addWidget(self.email_combo, 1, 1)
        root.addWidget(box)

        self.preview = QTableWidget()
        self.preview.setEditTriggers(QTableWidget.NoEditTriggers)
        root.addWidget(self.preview, 1)

        self.info = QLabel("엑셀 파일을 선택하세요.")
        root.addWidget(self.info)
        self.email_combo.currentTextChanged.connect(lambda _: self.refresh_preview())

    def email_column(self) -> str:
        return self.email_combo.currentText()

    def choose_excel(self):
        p = pick_file(self, "엑셀 명단 선택", "엑셀 (*.xlsx *.xlsm)")
        if not p:
            return
        self.excel_edit.setText(p)
        try:
            self.headers, self.rows = read_excel(p)
        except Exception as e:
            QMessageBox.warning(self, "엑셀 읽기 실패", str(e))
            return
        self.email_combo.blockSignals(True)
        self.email_combo.clear()
        self.email_combo.addItems(self.headers)
        guess = guess_email_column(self.headers, self.rows)
        if guess:
            self.email_combo.setCurrentText(guess)
        self.email_combo.blockSignals(False)
        self.refresh_preview()
        if self.on_loaded:
            self.on_loaded()

    def refresh_preview(self):
        col = self.email_column()
        self.preview.setColumnCount(len(self.headers))
        self.preview.setHorizontalHeaderLabels(self.headers)
        show = self.rows[:50]
        self.preview.setRowCount(len(show))
        bad = 0
        for r, row in enumerate(show):
            for c, h in enumerate(self.headers):
                item = QTableWidgetItem(row.get(h, ""))
                if h == col and not validate_email(row.get(h, "")):
                    item.setBackground(QColor(255, 200, 200))
                self.preview.setItem(r, c, item)
        bad = sum(1 for row in self.rows if not validate_email(row.get(col, "")))
        msg = f"명단 {len(self.rows)}건 로드."
        if bad:
            msg += f"  ⚠ 이메일 주소 오류 {bad}건 (빨간 칸)"
        self.info.setText(msg)

    def invalid_rows(self):
        col = self.email_column()
        return [i for i, row in enumerate(self.rows) if not validate_email(row.get(col, ""))]


class PreviewDialog(QDialog):
    """행을 골라 렌더된 제목/본문/첨부를 확인하는 미리보기."""

    def __init__(self, parent, build_cfg, rows):
        super().__init__(parent)
        self.build_cfg, self.rows = build_cfg, rows
        self.setWindowTitle("발송 미리보기")
        self.setMinimumSize(560, 480)
        root = QVBoxLayout(self)
        top = QHBoxLayout()
        top.addWidget(QLabel("행 번호:"))
        self.spin = QSpinBox()
        self.spin.setRange(1, len(rows))
        self.spin.valueChanged.connect(self.refresh)
        top.addWidget(self.spin)
        top.addStretch()
        root.addLayout(top)
        self.text = QPlainTextEdit()
        self.text.setReadOnly(True)
        root.addWidget(self.text, 1)
        close = QPushButton("닫기")
        close.clicked.connect(self.accept)
        root.addWidget(close, alignment=Qt.AlignRight)
        self.refresh()

    def refresh(self):
        i = self.spin.value() - 1
        d = dry_run_row(self.build_cfg(), self.rows[i], i)
        lines = [f"받는 사람: {d['email']}", f"제목: {d['subject']}", ""]
        lines.append(d["body"])
        lines.append("")
        if d["attachments"]:
            lines.append("첨부:")
            lines += [f"  - {os.path.basename(p)}" for p in d["attachments"]]
        else:
            lines.append("첨부: 없음")
        for w in d["warnings"]:
            lines.append(f"⚠ {w}")
        self.text.setPlainText("\n".join(lines))


class ComposeTab(QWidget):
    """탭 2: 제목/본문 작성 + 첨부 설정."""

    def __init__(self, list_tab: ListTab):
        super().__init__()
        self.list_tab = list_tab
        root = QVBoxLayout(self)

        chip_box = QGroupBox("자리표시 삽입 (클릭하면 커서 위치에 {{컬럼}} 추가)")
        self.chip_row = QHBoxLayout(chip_box)
        self.chip_row.addStretch()
        root.addWidget(chip_box)

        form_box = QGroupBox("메일 내용")
        form = QFormLayout(form_box)
        self.subject_edit = QLineEdit()
        self.subject_edit.setPlaceholderText("예: {{이름}}님, {{과정명}} 수료증을 보내드립니다")
        form.addRow("제목", self.subject_edit)
        self.body_edit = QPlainTextEdit()
        self.body_edit.setPlaceholderText("예:\n{{이름}}님 안녕하세요.\n{{과정명}} 수료증을 첨부합니다.")
        form.addRow("본문", self.body_edit)
        self.chk_html = QCheckBox("HTML 본문")
        form.addRow("", self.chk_html)
        root.addWidget(form_box, 1)

        fix_box = QGroupBox("공통 첨부 (모든 사람에게)")
        v = QVBoxLayout(fix_box)
        self.fixed_list = QListWidget()
        self.fixed_list.setMaximumHeight(80)
        v.addWidget(self.fixed_list)
        h = QHBoxLayout()
        add_btn = QPushButton("추가…")
        add_btn.clicked.connect(self.add_fixed)
        del_btn = QPushButton("제거")
        del_btn.clicked.connect(self.del_fixed)
        h.addWidget(add_btn)
        h.addWidget(del_btn)
        h.addStretch()
        v.addLayout(h)
        root.addWidget(fix_box)

        per_box = QGroupBox("개인별 첨부 (행마다 다른 파일 — 예: HWP 배치로 만든 수료증 PDF)")
        grid = QGridLayout(per_box)
        self.attach_dir_edit = QLineEdit()
        self.attach_dir_edit.setPlaceholderText("첨부파일이 들어있는 폴더 (비우면 미사용)")
        dir_btn = QPushButton("찾아보기…")
        dir_btn.clicked.connect(self.choose_attach_dir)
        grid.addWidget(QLabel("폴더"), 0, 0)
        grid.addWidget(self.attach_dir_edit, 0, 1)
        grid.addWidget(dir_btn, 0, 2)
        self.pattern_edit = QLineEdit()
        self.pattern_edit.setPlaceholderText("예: {이름}_수료증.pdf  (* 와일드카드 가능)")
        grid.addWidget(QLabel("파일명 패턴"), 1, 0)
        grid.addWidget(self.pattern_edit, 1, 1)
        self.policy_combo = QComboBox()
        for label, _ in MISSING_POLICIES:
            self.policy_combo.addItem(label)
        grid.addWidget(QLabel("파일 누락 시"), 2, 0)
        grid.addWidget(self.policy_combo, 2, 1)
        root.addWidget(per_box)

        self.preview_btn = QPushButton("미리보기…")
        self.preview_btn.clicked.connect(self.show_preview)
        root.addWidget(self.preview_btn, alignment=Qt.AlignRight)

    def refresh_chips(self):
        while self.chip_row.count() > 1:
            item = self.chip_row.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for h in self.list_tab.headers:
            btn = QPushButton("{{%s}}" % h)
            btn.clicked.connect(lambda _=False, name=h: self.insert_chip(name))
            self.chip_row.insertWidget(self.chip_row.count() - 1, btn)

    def insert_chip(self, name):
        text = "{{%s}}" % name
        if self.subject_edit.hasFocus():
            self.subject_edit.insert(text)
        else:
            self.body_edit.insertPlainText(text)

    def add_fixed(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "공통 첨부 선택", "", "모든 파일 (*.*)")
        for p in paths:
            self.fixed_list.addItem(p)

    def del_fixed(self):
        for item in self.fixed_list.selectedItems():
            self.fixed_list.takeItem(self.fixed_list.row(item))

    def choose_attach_dir(self):
        d = pick_dir(self, "개인별 첨부 폴더 선택")
        if d:
            self.attach_dir_edit.setText(d)

    def fixed_attachments(self):
        return [self.fixed_list.item(i).text() for i in range(self.fixed_list.count())]

    def missing_policy(self):
        return MISSING_POLICIES[self.policy_combo.currentIndex()][1]

    def show_preview(self):
        if not self.list_tab.rows:
            QMessageBox.warning(self, "확인", "먼저 [명단] 탭에서 엑셀을 불러오세요.")
            return
        main = self.window()
        PreviewDialog(self, main.build_config, self.list_tab.rows).exec()


class SendTab(QWidget):
    """탭 3: SMTP 설정 + 발송."""

    def __init__(self, list_tab: ListTab, compose_tab: ComposeTab):
        super().__init__()
        self.list_tab, self.compose_tab = list_tab, compose_tab
        self.worker = None
        self.last_results, self.last_rows = [], []
        root = QVBoxLayout(self)

        smtp_box = QGroupBox("보내는 메일 계정")
        grid = QGridLayout(smtp_box)
        self.preset_combo = QComboBox()
        self.preset_combo.addItems(PRESETS.keys())
        self.preset_combo.currentTextChanged.connect(self.apply_preset)
        grid.addWidget(QLabel("서비스"), 0, 0)
        grid.addWidget(self.preset_combo, 0, 1)
        self.host_edit = QLineEdit()
        self.port_spin = QSpinBox()
        self.port_spin.setRange(1, 65535)
        self.security_combo = QComboBox()
        for label, _ in SECURITY_ITEMS:
            self.security_combo.addItem(label)
        hrow = QHBoxLayout()
        hrow.addWidget(self.host_edit, 1)
        hrow.addWidget(self.port_spin)
        hrow.addWidget(self.security_combo)
        grid.addWidget(QLabel("SMTP 서버"), 1, 0)
        grid.addLayout(hrow, 1, 1)
        self.user_edit = QLineEdit()
        self.user_edit.setPlaceholderText("메일 주소 전체 (예: me@naver.com)")
        grid.addWidget(QLabel("아이디"), 2, 0)
        grid.addWidget(self.user_edit, 2, 1)
        self.pw_edit = QLineEdit()
        self.pw_edit.setEchoMode(QLineEdit.Password)
        self.pw_edit.setPlaceholderText("앱 비밀번호")
        self.chk_save_pw = QCheckBox("비밀번호 저장 (이 PC에 암호화 보관)")
        prow = QHBoxLayout()
        prow.addWidget(self.pw_edit, 1)
        prow.addWidget(self.chk_save_pw)
        grid.addWidget(QLabel("비밀번호"), 3, 0)
        grid.addLayout(prow, 3, 1)
        self.sender_edit = QLineEdit()
        self.sender_edit.setPlaceholderText("받는 사람에게 표시될 이름 (예: 홍길동 강사)")
        grid.addWidget(QLabel("보내는 이름"), 4, 0)
        grid.addWidget(self.sender_edit, 4, 1)
        self.note = QLabel()
        self.note.setWordWrap(True)
        self.note.setStyleSheet("color: #666;")
        grid.addWidget(self.note, 5, 0, 1, 2)
        root.addWidget(smtp_box)

        run_box = QGroupBox("발송")
        v = QVBoxLayout(run_box)
        opt = QHBoxLayout()
        opt.addWidget(QLabel("발송 간격(초)"))
        self.delay_spin = QDoubleSpinBox()
        self.delay_spin.setRange(0.0, 60.0)
        self.delay_spin.setSingleStep(0.5)
        self.delay_spin.setValue(2.0)
        opt.addWidget(self.delay_spin)
        opt.addStretch()
        v.addLayout(opt)
        btns = QHBoxLayout()
        self.test_btn = QPushButton("테스트 발송 (1행 → 내 주소)")
        self.test_btn.clicked.connect(self.start_test)
        self.send_btn = QPushButton("전체 발송")
        self.send_btn.clicked.connect(self.start_all)
        self.cancel_btn = QPushButton("중지")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel)
        self.retry_btn = QPushButton("실패만 재발송")
        self.retry_btn.setEnabled(False)
        self.retry_btn.clicked.connect(self.start_retry)
        self.export_btn = QPushButton("결과 내보내기…")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export)
        for b in (self.test_btn, self.send_btn, self.cancel_btn, self.retry_btn, self.export_btn):
            btns.addWidget(b)
        v.addLayout(btns)
        self.progress = QProgressBar()
        v.addWidget(self.progress)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        v.addWidget(self.log, 1)
        root.addWidget(run_box, 1)

        self.apply_preset(self.preset_combo.currentText())

    # --- 설정 ---
    def apply_preset(self, name):
        p = PRESETS.get(name)
        if not p:
            return
        if p["host"]:
            self.host_edit.setText(p["host"])
            self.port_spin.setValue(p["port"])
            self.security_combo.setCurrentIndex(
                next(i for i, (_, v) in enumerate(SECURITY_ITEMS) if v == p["security"]))
        self.note.setText(p["note"])

    def security(self):
        return SECURITY_ITEMS[self.security_combo.currentIndex()][1]

    # --- 발송 ---
    def _validate_common(self) -> bool:
        if not self.list_tab.rows:
            QMessageBox.warning(self, "확인", "[명단] 탭에서 엑셀을 불러오세요.")
            return False
        if not self.host_edit.text().strip() or not self.user_edit.text().strip() \
                or not self.pw_edit.text():
            QMessageBox.warning(self, "확인", "SMTP 서버/아이디/비밀번호를 입력하세요.")
            return False
        cfg = self.window().build_config()
        if not cfg.subject_tpl:
            QMessageBox.warning(self, "확인", "[메일 작성] 탭에서 제목을 입력하세요.")
            return False
        unknown = (find_unknown_placeholders(cfg.subject_tpl, self.list_tab.headers)
                   + find_unknown_placeholders(cfg.body_tpl, self.list_tab.headers))
        if unknown:
            r = QMessageBox.question(
                self, "자리표시 확인",
                "엑셀에 없는 자리표시가 있습니다 (오타?):\n  "
                + ", ".join("{{%s}}" % u for u in unknown)
                + "\n\n빈 값으로 두고 계속할까요?")
            if r != QMessageBox.Yes:
                return False
        return True

    def start_test(self):
        if not self._validate_common():
            return
        me = self.user_edit.text().strip()
        cfg = self.window().build_config()
        cfg.subject_tpl = "[테스트] " + cfg.subject_tpl
        self.log.appendPlainText(f"테스트 발송: 1행 데이터를 {me} (내 주소)로 보냅니다.")
        self._run(cfg, self.list_tab.rows[:1], to_override=me)

    def start_all(self):
        if not self._validate_common():
            return
        rows = self.list_tab.rows
        invalid = self.list_tab.invalid_rows()
        if invalid:
            r = QMessageBox.question(
                self, "주소 오류",
                f"이메일 주소 오류 {len(invalid)}건이 있습니다.\n해당 행을 제외하고 발송할까요?")
            if r != QMessageBox.Yes:
                return
            rows = [row for i, row in enumerate(rows) if i not in set(invalid)]
        if not rows:
            QMessageBox.warning(self, "확인", "발송할 행이 없습니다.")
            return
        delay = self.delay_spin.value()
        est_min = len(rows) * delay / 60
        r = QMessageBox.question(
            self, "전체 발송",
            f"{len(rows)}건을 발송합니다. (간격 {delay:g}초, 예상 약 {est_min:.0f}분)\n계속할까요?")
        if r != QMessageBox.Yes:
            return
        self._run(self.window().build_config(), rows)

    def start_retry(self):
        failed_idx = [r["index"] for r in self.last_results if r["error"]]
        rows = [self.last_rows[i] for i in failed_idx]
        if not rows:
            QMessageBox.information(self, "안내", "재발송할 실패 건이 없습니다.")
            return
        self.log.appendPlainText(f"실패 {len(rows)}건 재발송 시작")
        self._run(self.window().build_config(), rows)

    def _run(self, cfg, rows, to_override=None):
        self.last_rows = rows
        self.progress.setMaximum(len(rows))
        self.progress.setValue(0)
        self._set_running(True)
        self.window().persist_settings()
        self.worker = SendWorker(cfg, rows, to_override=to_override)
        self.worker.progressed.connect(self.on_progress)
        self.worker.done.connect(self.on_done)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def _set_running(self, running):
        for b in (self.test_btn, self.send_btn, self.retry_btn):
            b.setEnabled(not running)
        self.cancel_btn.setEnabled(running)

    def cancel(self):
        if self.worker:
            self.worker.cancel()
            self.log.appendPlainText("중지 요청됨 — 현재 메일까지 보내고 멈춥니다.")

    def on_progress(self, done, total, name):
        self.progress.setValue(done)
        self.log.appendPlainText(f"[{done}/{total}] {name}")

    def on_done(self, results):
        self.last_results = results
        errs = [r for r in results if r["error"]]
        self.log.appendPlainText(f"완료: 성공 {len(results) - len(errs)} / 실패 {len(errs)}")
        self._set_running(False)
        self.retry_btn.setEnabled(bool(errs))
        self.export_btn.setEnabled(bool(results))

    def on_failed(self, msg):
        QMessageBox.critical(self, "발송 중단", msg)
        self.log.appendPlainText(f"중단: {msg}")
        self._set_running(False)

    def export(self):
        path, _ = QFileDialog.getSaveFileName(self, "결과 저장", "발송결과.xlsx",
                                              "엑셀 (*.xlsx);;CSV (*.csv)")
        if not path:
            return
        try:
            export_results(self.last_results, path)
            self.log.appendPlainText(f"결과 저장: {path}")
        except Exception as e:
            QMessageBox.warning(self, "저장 실패", str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        icon = resource_path("app.ico")
        if icon:
            from PySide6.QtGui import QIcon
            self.setWindowIcon(QIcon(icon))
        self.resize(860, 680)

        self.list_tab = ListTab(on_loaded=self.on_excel_loaded)
        self.compose_tab = ComposeTab(self.list_tab)
        self.send_tab = SendTab(self.list_tab, self.compose_tab)
        tabs = QTabWidget()
        tabs.addTab(self.list_tab, "1) 명단")
        tabs.addTab(self.compose_tab, "2) 메일 작성")
        tabs.addTab(self.send_tab, "3) 발송")
        self.setCentralWidget(tabs)
        self.restore_settings()

    def on_excel_loaded(self):
        self.compose_tab.refresh_chips()

    def build_config(self) -> SendConfig:
        s = self.send_tab
        c = self.compose_tab
        return SendConfig(
            host=s.host_edit.text().strip(), port=s.port_spin.value(),
            security=s.security(), user=s.user_edit.text().strip(),
            password=s.pw_edit.text(), sender_name=s.sender_edit.text().strip(),
            to_column=self.list_tab.email_column(),
            subject_tpl=c.subject_edit.text().strip(),
            body_tpl=c.body_edit.toPlainText(),
            body_is_html=c.chk_html.isChecked(),
            fixed_attachments=c.fixed_attachments(),
            attach_dir=c.attach_dir_edit.text().strip(),
            attach_pattern=c.pattern_edit.text().strip(),
            missing_attach_policy=c.missing_policy(),
            delay_sec=s.delay_spin.value(),
        )

    def persist_settings(self):
        s = self.send_tab
        st.save_settings({
            "preset": s.preset_combo.currentText(),
            "host": s.host_edit.text().strip(), "port": s.port_spin.value(),
            "security": s.security(), "user": s.user_edit.text().strip(),
            "sender_name": s.sender_edit.text().strip(),
            "delay_sec": s.delay_spin.value(),
            "password": s.pw_edit.text(),
            "save_password": s.chk_save_pw.isChecked(),
        })

    def restore_settings(self):
        d = st.load_settings()
        if not d:
            return
        s = self.send_tab
        if d.get("preset") in PRESETS:
            s.preset_combo.setCurrentText(d["preset"])
        if d.get("host"):
            s.host_edit.setText(d["host"])
        if d.get("port"):
            s.port_spin.setValue(int(d["port"]))
        if d.get("security"):
            idx = next((i for i, (_, v) in enumerate(SECURITY_ITEMS) if v == d["security"]), 0)
            s.security_combo.setCurrentIndex(idx)
        s.user_edit.setText(d.get("user", ""))
        s.sender_edit.setText(d.get("sender_name", ""))
        if d.get("delay_sec") is not None:
            s.delay_spin.setValue(float(d["delay_sec"]))
        if d.get("save_password"):
            s.chk_save_pw.setChecked(True)
            s.pw_edit.setText(d.get("password", ""))

    def closeEvent(self, event):
        try:
            self.persist_settings()
        except Exception:
            pass
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
