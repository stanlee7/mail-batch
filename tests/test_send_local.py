# -*- coding: utf-8 -*-
"""발송 통합 테스트 — aiosmtpd 로컬 서버 (127.0.0.1, 인증 없음, security='none').

python tests/test_send_local.py 로 실행. dev 의존성: pip install aiosmtpd
"""
import os
import sys
import time
from email import message_from_bytes
from email.policy import default as default_policy

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aiosmtpd.controller import Controller  # noqa: E402

from mailbatch.engine import SendConfig, send_batch  # noqa: E402
from tests.make_samples import DUMMY_PDF, make  # noqa: E402

FAILED = []


def check(name, cond, detail=""):
    if cond:
        print(f"  ok  {name}")
    else:
        print(f"FAIL  {name}  {detail}")
        FAILED.append(name)


class Collector:
    """수신 메일을 모은다. fail_rcpt에 든 주소는 거부해 행별 실패를 흉내낸다."""

    def __init__(self, fail_rcpt=()):
        self.messages = []
        self.fail_rcpt = set(fail_rcpt)

    async def handle_RCPT(self, server, session, envelope, address, rcpt_options):
        if address in self.fail_rcpt:
            return "550 mailbox unavailable"
        envelope.rcpt_tos.append(address)
        return "250 OK"

    async def handle_DATA(self, server, session, envelope):
        self.messages.append(envelope)
        return "250 Message accepted"


def base_cfg(attach_dir, port):
    return SendConfig(
        host="127.0.0.1", port=port, security="none",
        user="", password="",  # 인증 없는 테스트 서버 — login 생략됨
        sender_name="홍 강사", to_column="이메일",
        subject_tpl="{{이름}}님 수료증", body_tpl="{{이름}}님, {{과정명}} 수료를 축하합니다.",
        attach_dir=attach_dir, attach_pattern="{이름}_수료증.pdf",
        missing_attach_policy="send_without", delay_sec=0.0,
    )


def main():
    xlsx, attach_dir = make()
    from mailbatch.engine import read_excel
    _, rows = read_excel(xlsx)

    handler = Collector()
    ctl = Controller(handler, hostname="127.0.0.1", port=8025)
    ctl.start()
    try:
        port = ctl.port

        # 1) 기본 전달: 3행 발송, 이철수는 첨부 누락이지만 send_without 정책으로 발송됨
        cfg = base_cfg(attach_dir, port)
        res = send_batch(cfg, rows)
        check("3건 발송", len(res) == 3 and all(r["error"] is None for r in res), res)
        check("서버 3건 수신", len(handler.messages) == 3)
        msg0 = message_from_bytes(handler.messages[0].content, policy=default_policy)
        check("제목 개인화", "홍길동" in msg0["Subject"])
        atts = list(msg0.iter_attachments())
        check("첨부 전달", len(atts) == 1 and atts[0].get_content() == DUMMY_PDF)
        check("한글 첨부 파일명", atts[0].get_filename() == "홍길동_수료증.pdf", atts[0].get_filename())
        msg2 = message_from_bytes(handler.messages[2].content, policy=default_policy)
        check("누락 행 첨부 없이 발송", len(list(msg2.iter_attachments())) == 0)

        # 2) 누락 정책 skip_row / fail
        handler.messages.clear()
        cfg = base_cfg(attach_dir, port)
        cfg.missing_attach_policy = "skip_row"
        res = send_batch(cfg, rows)
        check("skip_row: 2건 발송 1건 건너뜀",
              len(handler.messages) == 2 and "건너뜀" in (res[2]["error"] or ""), res[2])
        handler.messages.clear()
        cfg.missing_attach_policy = "fail"
        res = send_batch(cfg, rows)
        check("fail: 누락 행 실패 기록", res[2]["error"] is not None and "이철수" in res[2]["error"], res[2])

        # 3) 행별 실패 격리 — 김영희 수신 거부, 나머지는 계속
        handler2 = Collector(fail_rcpt={"kim@example.com"})
        ctl2 = Controller(handler2, hostname="127.0.0.1", port=8026)
        ctl2.start()
        try:
            cfg = base_cfg(attach_dir, ctl2.port)
            res = send_batch(cfg, rows)
            check("실패 행만 에러", res[1]["error"] is not None and res[0]["error"] is None
                  and res[2]["error"] is None, [r["error"] for r in res])
            check("나머지 2건 전달", len(handler2.messages) == 2)
            # 실패만 재발송 시나리오: 실패 인덱스만 다시
            retry_rows = [rows[r["index"]] for r in res if r["error"]]
            check("재발송 대상 1건", len(retry_rows) == 1 and retry_rows[0]["이름"] == "김영희")
        finally:
            ctl2.stop()

        # 4) to_override (테스트 발송) — 모든 메일이 내 주소로
        handler.messages.clear()
        cfg = base_cfg(attach_dir, port)
        res = send_batch(cfg, rows[:1], to_override="me@example.com")
        check("테스트 발송 수신자 교체",
              handler.messages[-1].rcpt_tos == ["me@example.com"], handler.messages[-1].rcpt_tos)

        # 5) 발송 간격 + 취소 반응
        handler.messages.clear()
        cfg = base_cfg(attach_dir, port)
        cfg.delay_sec = 0.5
        t0 = time.monotonic()
        send_batch(cfg, rows)
        elapsed = time.monotonic() - t0
        check("간격 준수 (2회 대기 ≥1s)", elapsed >= 1.0, f"{elapsed:.2f}s")

        handler.messages.clear()
        cfg.delay_sec = 5.0
        cancel_after = {"n": 0}

        def cancel():
            return cancel_after["n"] >= 1  # 첫 행 발송 후 취소

        def progress(done, total, name):
            cancel_after["n"] = done

        t0 = time.monotonic()
        res = send_batch(cfg, rows, progress=progress, cancel=cancel)
        elapsed = time.monotonic() - t0
        check("취소 즉각 반응 (5s 간격 중 <2s 종료)", elapsed < 2.0 and len(res) == 1,
              f"{elapsed:.2f}s, {len(res)}건")
    finally:
        ctl.stop()

    print()
    if FAILED:
        print(f"FAILED: {len(FAILED)}건 — {FAILED}")
        sys.exit(1)
    print("ALL PASS")


if __name__ == "__main__":
    main()
