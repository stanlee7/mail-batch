# -*- coding: utf-8 -*-
"""테스트용 샘플 데이터 생성: 명단 xlsx + 더미 수료증 PDF."""
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))

PEOPLE = [
    ("홍길동", "hong@example.com", "파이썬 업무자동화", "2026"),
    ("김영희", "kim@example.com", "엑셀 데이터 분석", "2026"),
    ("이철수", "lee@example.com", "AI 활용 실무", "2026.0"),  # 숫자형 셀 검증용
]

# %PDF 헤더만 갖춘 최소 더미 — 첨부 바이트 무결성 검증용
DUMMY_PDF = b"%PDF-1.4\n1 0 obj<</Type/Catalog>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF\n"


def make(out_dir: str = None):
    out_dir = out_dir or os.path.join(HERE, "work")
    os.makedirs(out_dir, exist_ok=True)
    xlsx = os.path.join(out_dir, "명단.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["이름", "이메일", "과정명", "연도"])
    for name, email, course, year in PEOPLE:
        ws.append([name, email, course, float(year) if "." in year else year])
    wb.save(xlsx)

    attach_dir = os.path.join(out_dir, "수료증")
    os.makedirs(attach_dir, exist_ok=True)
    for name, _, _, _ in PEOPLE[:2]:  # 이철수는 일부러 누락 (누락 정책 검증용)
        with open(os.path.join(attach_dir, f"{name}_수료증.pdf"), "wb") as fp:
            fp.write(DUMMY_PDF)
    return xlsx, attach_dir


if __name__ == "__main__":
    x, a = make()
    print("xlsx:", x)
    print("attach_dir:", a)
