# -*- coding: utf-8 -*-
"""한국 주요 메일 서비스 SMTP 프리셋."""

PRESETS = {
    "네이버": {
        "host": "smtp.naver.com", "port": 465, "security": "ssl",
        "note": "네이버 메일 → 환경설정 → POP3/IMAP 설정에서 'POP3/SMTP 사용'을 켜세요.\n"
                "2단계 인증을 쓰면 '앱 비밀번호'를 발급해 입력하세요. 일일 발송 한도 약 500건.",
    },
    "지메일": {
        "host": "smtp.gmail.com", "port": 465, "security": "ssl",
        "note": "Google 계정 → 보안 → 2단계 인증을 켠 뒤 '앱 비밀번호'를 발급해 입력하세요.\n"
                "일반 비밀번호로는 로그인되지 않습니다. 일일 발송 한도 약 500건.",
    },
    "다음(카카오)": {
        "host": "smtp.daum.net", "port": 465, "security": "ssl",
        "note": "다음 메일 → 환경설정 → IMAP/POP3 설정에서 사용을 켜세요.",
    },
    "직접 입력": {
        "host": "", "port": 465, "security": "ssl",
        "note": "회사/기관 메일 서버의 SMTP 주소·포트를 입력하세요. (보안 방식은 보통 SSL 465 또는 STARTTLS 587)",
    },
}
