# -*- coding: utf-8 -*-
"""settings.py DPAPI 왕복 + 저장/로드 테스트. python tests/test_settings.py"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 실제 %APPDATA% 를 건드리지 않도록 임시 폴더로 돌린다
os.environ["APPDATA"] = tempfile.mkdtemp(prefix="mailbatch_settings_")

from mailbatch import settings  # noqa: E402

FAILED = []


def check(name, cond, detail=""):
    if cond:
        print(f"  ok  {name}")
    else:
        print(f"FAIL  {name}  {detail}")
        FAILED.append(name)


def main():
    enc = settings.dpapi_protect("앱비밀번호123!".encode("utf-8"))
    check("DPAPI 왕복", settings.dpapi_unprotect(enc).decode("utf-8") == "앱비밀번호123!")
    check("암호문에 평문 없음", "앱비밀번호" not in enc)

    # 저장 안 함 (기본): 비밀번호가 파일에 남지 않아야 한다
    settings.save_settings({"host": "smtp.naver.com", "user": "me@naver.com",
                            "password": "secret!", "save_password": False})
    with open(settings.settings_path(), encoding="utf-8") as fp:
        raw = fp.read()
    check("미저장 시 비밀번호 없음", "secret!" not in raw and "password" not in raw, raw)
    d = settings.load_settings()
    check("호스트 로드", d.get("host") == "smtp.naver.com")

    # 저장 함: 암호화돼 저장되고 로드 시 복호화
    settings.save_settings({"host": "smtp.naver.com", "user": "me@naver.com",
                            "password": "secret!", "save_password": True})
    with open(settings.settings_path(), encoding="utf-8") as fp:
        raw = fp.read()
    check("저장 시 평문 없음", "secret!" not in raw, raw)
    d = settings.load_settings()
    check("복호화 로드", d.get("password") == "secret!")

    print()
    if FAILED:
        print(f"FAILED: {FAILED}")
        sys.exit(1)
    print("ALL PASS")


if __name__ == "__main__":
    main()
