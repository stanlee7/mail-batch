# -*- coding: utf-8 -*-
"""앱 아이콘 생성: 초록 둥근 사각형 + 봉투 + '@' 배지 (메일 배치 상징)."""
import os

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ICO = os.path.join(ROOT, "app.ico")

SIZE = 256
GREEN = (5, 150, 105, 255)
LIGHT = (209, 250, 229, 255)
WHITE = (255, 255, 255, 255)
DARK = (4, 120, 87, 255)


def make_base():
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([8, 8, SIZE - 8, SIZE - 8], radius=48, fill=GREEN)
    # 겹친 봉투 두 장 (일괄 발송 상징)
    d.rounded_rectangle([52, 86, 168, 186], radius=10, fill=LIGHT)
    d.rounded_rectangle([72, 66, 188, 166], radius=10, fill=WHITE)
    # 봉투 덮개 (V자)
    d.line([76, 72, 130, 122], fill=(167, 243, 208, 255), width=10)
    d.line([184, 72, 130, 122], fill=(167, 243, 208, 255), width=10)
    # '@' 배지
    d.ellipse([142, 138, 222, 218], fill=DARK)
    try:
        font = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", 48)
    except OSError:
        font = ImageFont.truetype(r"C:\Windows\Fonts\malgun.ttf", 48)
    d.text((182, 176), "@", font=font, fill=WHITE, anchor="mm")
    return img


def main():
    base = make_base()
    base.save(ICO, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("icon:", ICO, os.path.getsize(ICO), "bytes")


if __name__ == "__main__":
    main()
