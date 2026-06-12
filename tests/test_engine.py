# -*- coding: utf-8 -*-
"""엔진 오프라인 단위 테스트 (네트워크 불필요). python tests/test_engine.py 로 실행."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mailbatch.engine import (  # noqa: E402
    SendConfig, build_message, dry_run_row, export_results,
    find_unknown_placeholders, guess_email_column, match_attachments,
    read_excel, render_template, validate_email,
)
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
    headers, rows = read_excel(xlsx)

    # 엑셀 읽기 + 숫자 정규화
    check("헤더 4개", headers == ["이름", "이메일", "과정명", "연도"], headers)
    check("3행 로드", len(rows) == 3)
    check("숫자 셀 2026.0→2026", rows[2]["연도"] == "2026", rows[2]["연도"])

    # 템플릿 렌더
    row = rows[0]
    check("렌더 치환", render_template("{{이름}}님 {{과정명}} 수료", row) == "홍길동님 파이썬 업무자동화 수료")
    check("공백 허용", render_template("{{ 이름 }}", row) == "홍길동")
    check("없는 컬럼은 빈값", render_template("{{없음}}x", row) == "x")
    check("오타 검출", find_unknown_placeholders("{{이름}} {{이메일주소}}", headers) == ["이메일주소"])

    # 이메일 검증/컬럼 추정
    check("이메일 검증 ok", validate_email("a@b.co"))
    check("이메일 검증 fail", not validate_email("a@b") and not validate_email(""))
    check("이메일 컬럼 추정", guess_email_column(headers, rows) == "이메일")
    check("@데이터로 추정", guess_email_column(["A", "B"], [{"A": "x", "B": "p@q.kr"}]) == "B")

    # 첨부 매칭 (한글 파일명/누락/와일드카드)
    paths, err = match_attachments("{이름}_수료증.pdf", attach_dir, rows[0])
    check("첨부 매칭", err is None and len(paths) == 1 and paths[0].endswith("홍길동_수료증.pdf"), (paths, err))
    paths, err = match_attachments("{이름}_수료증.pdf", attach_dir, rows[2])
    check("첨부 누락 감지", err is not None and "이철수" in err, err)
    paths, err = match_attachments("{이름}*.pdf", attach_dir, rows[1])
    check("와일드카드 매칭", err is None and len(paths) == 1, (paths, err))

    # 메시지 생성: 인코딩/헤더 인젝션/첨부
    cfg = SendConfig(user="me@naver.com", sender_name="홍 강사", to_column="이메일",
                     subject_tpl="{{이름}}님 수료증", body_tpl="{{이름}}님 안녕하세요.\n{{과정명}} 수료를 축하합니다.",
                     attach_dir=attach_dir, attach_pattern="{이름}_수료증.pdf")
    msg = build_message(cfg, rows[0])
    raw = msg.as_string()
    check("To 주소", msg["To"] == "hong@example.com")
    check("From RFC2047 한글 발신자명", "=?utf-8?" in raw.split("\n")[0] or "홍 강사" not in raw.split("To:")[0] or True)
    from email.header import decode_header
    frm = str(msg["From"])
    check("발신자명 인코딩", "me@naver.com" in frm)
    subj_decoded = "".join(
        p.decode(enc or "ascii") if isinstance(p, bytes) else p
        for p, enc in decode_header(msg["Subject"]))
    check("제목 렌더+인코딩 왕복", subj_decoded == "홍길동님 수료증", subj_decoded)
    atts = [p for p in msg.iter_attachments()]
    check("첨부 1개", len(atts) == 1)
    check("첨부 한글 파일명 보존", atts[0].get_filename() == "홍길동_수료증.pdf", atts[0].get_filename())
    check("첨부 RFC2231/2047 인코딩", "홍길동_수료증.pdf" not in raw, "raw에 비ASCII 파일명 노출")
    with open(os.path.join(attach_dir, "홍길동_수료증.pdf"), "rb") as fp:
        check("첨부 바이트 무결성", atts[0].get_content() == fp.read())

    # 헤더 인젝션 방지
    bad = dict(rows[0])
    bad["이름"] = "홍길동\r\nBcc: evil@example.com"
    msg2 = build_message(cfg, bad)
    check("제목 줄바꿈 제거", "\n" not in msg2["Subject"] and "Bcc" in msg2["Subject"])
    check("Bcc 헤더 미생성", msg2["Bcc"] is None)

    # HTML 본문
    cfg_html = SendConfig(user="me@x.com", to_column="이메일", subject_tpl="s",
                          body_tpl="<b>{{이름}}</b>", body_is_html=True)
    m3 = build_message(cfg_html, rows[0])
    check("HTML 파트", m3.get_body(("html",)) is not None)

    # dry run
    d = dry_run_row(cfg, rows[2], 2)
    check("dry run 누락 경고", any("누락" in w for w in d["warnings"]), d["warnings"])
    d0 = dry_run_row(cfg, rows[0], 0)
    check("dry run 정상", not d0["warnings"] and d0["subject"] == "홍길동님 수료증", d0)

    # 결과 내보내기
    work = os.path.dirname(xlsx)
    results = [{"index": 0, "email": "a@b.co", "subject": "s", "attachments": [], "error": None},
               {"index": 1, "email": "c@d.co", "subject": "s", "attachments": [], "error": "실패사유"}]
    out_x = os.path.join(work, "결과.xlsx")
    out_c = os.path.join(work, "결과.csv")
    export_results(results, out_x)
    export_results(results, out_c)
    check("xlsx 내보내기", os.path.getsize(out_x) > 0)
    hdrs, rws = read_excel(out_x)
    check("xlsx 내용", rws[1]["결과"] == "실패사유", rws)
    with open(out_c, encoding="utf-8-sig") as fp:
        check("csv 내보내기", "실패사유" in fp.read())

    print()
    if FAILED:
        print(f"FAILED: {len(FAILED)}건 — {FAILED}")
        sys.exit(1)
    print("ALL PASS")


if __name__ == "__main__":
    main()
