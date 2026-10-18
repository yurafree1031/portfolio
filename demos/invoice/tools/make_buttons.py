"""「操作」シートのボタンの画像を作り、Button.gs に書き出す（開発用。お客様には渡さない）

使い方: python tools/make_buttons.py
見た目は領収書のデモのボタンにそろえる（緑・角丸・▶・白い太字。560×120 で作り、シートでは 280×60 で表示）
"""
import base64
import io
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

DEMO_DIR = Path(__file__).resolve().parent.parent
FONT = "C:/Windows/Fonts/meiryob.ttc"
BUTTONS = [
    ("CREATE_BUTTON_PNG_BASE64", "請求書を作る"),
    ("PAYMENT_BUTTON_PNG_BASE64", "入金を取り込む"),
]
W, H = 560, 120
GREEN = (30, 122, 75)
SHADOW = (20, 90, 55)


def make(label):
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((2, 6, W - 3, H - 2), radius=16, fill=SHADOW)
    d.rounded_rectangle((2, 2, W - 3, H - 6), radius=16, fill=GREEN)
    font = ImageFont.truetype(FONT, 40)
    tw = d.textlength(label, font=font)
    tri = 26
    gap = 28
    x = (W - (tri + gap + tw)) / 2
    cy = (H - 6) / 2 + 2
    d.polygon([(x, cy - 13), (x, cy + 13), (x + tri - 4, cy)], fill="white")
    d.text((x + tri + gap, cy), label, font=font, fill="white", anchor="lm")
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    return base64.b64encode(buf.getvalue()).decode()


def main():
    lines = ["/** 「操作」シートに置くボタンの画像（PNG、280×60で表示）。tools/make_buttons.py で作り直す */"]
    for name, label in BUTTONS:
        lines.append(f"const {name} = '{make(label)}';")
    (DEMO_DIR / "Button.gs").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("Button.gs を書き出しました")


if __name__ == "__main__":
    main()
