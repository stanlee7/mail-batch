# -*- coding: utf-8 -*-
"""메일 배치 진입점.

python app.py            # GUI 실행
python app.py --selftest # 오프라인 엔진 검증 (패키징 exe 확인용, 네트워크 불필요)
"""
import os
import sys
import tempfile


def selftest() -> bool:
    log_path = os.path.join(tempfile.gettempdir(), "mailbatch_selftest.log")
    lines = []
    ok = False
    try:
        from mailbatch.engine import SendConfig, build_message, dry_run_row, render_template
        rows = [{"이름": "홍길동", "이메일": "hong@example.com", "과정명": "셀프테스트"},
                {"이름": "김영희", "이메일": "kim@example.com", "과정명": "셀프테스트"}]
        cfg = SendConfig(user="me@example.com", sender_name="셀프 테스트",
                         to_column="이메일", subject_tpl="{{이름}}님 안내",
                         body_tpl="{{이름}}님, {{과정명}} 메일입니다.")
        for i, row in enumerate(rows):
            msg = build_message(cfg, row)
            d = dry_run_row(cfg, row, i)
            lines.append(f"row {i}: to={msg['To']} subject={d['subject']} "
                         f"warnings={d['warnings']} bytes={len(msg.as_bytes())}")
        ok = (render_template("{{이름}}", rows[0]) == "홍길동"
              and all("@example.com" in ln for ln in lines))
    except Exception:
        import traceback
        lines.append(traceback.format_exc())
    lines.append("SELFTEST: " + ("SUCCESS" if ok else "FAIL"))
    try:
        with open(log_path, "w", encoding="utf-8") as fp:
            fp.write("\n".join(lines))
    except OSError:
        pass
    print("\n".join(lines))
    return ok


def main():
    if "--selftest" in sys.argv:
        sys.exit(0 if selftest() else 1)
    from mailbatch.gui import main as gui_main
    gui_main()


if __name__ == "__main__":
    main()
