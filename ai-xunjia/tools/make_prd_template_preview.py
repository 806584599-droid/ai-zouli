"""Render a clean cover preview for the AI Inquiry PRD artifact template.

LibreOffice on this Mac fails to fall back to a CJK font, so the rendered
docx page shows tofu boxes. Draw the cover directly with an explicit font
file instead so the template preview is legible.
"""

from PIL import Image, ImageDraw, ImageFont

W, H = 1414, 2000
BG = (255, 255, 255)
BLACK = (0, 0, 0)
HEADER_FILL = (217, 234, 247)
BORDER = (60, 60, 60)

CN = "/System/Library/Fonts/Supplemental/Songti.ttc"
CN_BOLD = "/System/Library/Fonts/Supplemental/Songti.ttc"
EN = "/System/Library/Fonts/Supplemental/Times New Roman.ttf"


def font(path, size, index=0):
    return ImageFont.truetype(path, size, index=index)


F_TITLE1 = font(CN_BOLD, 46, index=1)
F_TITLE2 = font(CN_BOLD, 54, index=1)
F_META = font(CN, 24)
F_META_B = font(CN_BOLD, 26, index=1)
F_H = font(CN_BOLD, 32, index=1)
F_TBL = font(CN, 20)
F_TBL_H = font(CN_BOLD, 20, index=1)

canvas = Image.new("RGB", (W, H), BG)
draw = ImageDraw.Draw(canvas)


def center(text, y, f, fill=BLACK):
    w = draw.textlength(text, font=f)
    draw.text(((W - w) / 2, y), text, font=f, fill=fill)
    return y + (f.size if hasattr(f, "size") else 30)


center("造价通 AI 询价", 240, F_TITLE1)
center("需求规格说明书", 300, F_TITLE2)

center("广东中建普联科技股份有限公司", 560, F_META_B)
center("文档密级：普通", 606, F_META)
center("文档状态：[  ] 草案   [√] 正式发布   [  ] 正在修订", 642, F_META)
center("版本号：V1.0", 678, F_META)
center("编写日期：2026-10-09", 714, F_META)

center("变更履历", 800, F_H)

# Change-log table
left, right = 150, W - 150
top = 870
row1_h = 56
row2_h = 210
cols = [0.09, 0.11, 0.60, 0.20]
xs = [left]
for c in cols:
    xs.append(xs[-1] + c * (right - left))

headers = ["序号", "版本", "变更描述", "修订人/日期"]
body = [
    "1",
    "V1.0",
    "建立 AI 询价一期需求规格说明书：首页入口与页面跳转、对话式询价输入与条件识别、"
    "通宝思考状态、历史询价推荐前 3 条规则、系数匹配与建议价口径、询价建议结果与结果操作、"
    "供应商联系方式与价格来源话术、盖章报价预览与导出、历史询价与会话删除、使用次数与日志埋点。",
    "邹力 / 2026-10-09",
]

draw.rectangle((left, top, right, top + row1_h), fill=HEADER_FILL, outline=BORDER, width=2)
for i, h in enumerate(headers):
    tw = draw.textlength(h, font=F_TBL_H)
    draw.text(((xs[i] + xs[i + 1] - tw) / 2, top + (row1_h - 24) / 2), h, font=F_TBL_H, fill=BLACK)
for i in range(1, len(xs) - 1):
    draw.line((xs[i], top, xs[i], top + row1_h + row2_h), fill=BORDER, width=2)

draw.rectangle((left, top + row1_h, right, top + row1_h + row2_h), outline=BORDER, width=2)


def wrap(text, f, max_w):
    lines, cur = [], ""
    for ch in text:
        if ch == "\n":
            lines.append(cur)
            cur = ""
            continue
        if draw.textlength(cur + ch, font=f) <= max_w:
            cur += ch
        else:
            lines.append(cur)
            cur = ch
    lines.append(cur)
    return lines


# First two columns centered
for i, val in enumerate([body[0], body[1]]):
    tw = draw.textlength(val, font=F_TBL)
    draw.text(((xs[i] + xs[i + 1] - tw) / 2, top + row1_h + (row2_h - 24) / 2), val, font=F_TBL, fill=BLACK)

desc_x = xs[2] + 14
desc_lines = wrap(body[2], F_TBL, (xs[3] - xs[2]) - 28)
ty = top + row1_h + 18
for ln in desc_lines:
    draw.text((desc_x, ty), ln, font=F_TBL, fill=BLACK)
    ty += 32

# Recompute col lines for the second row region
for i in range(1, len(xs) - 1):
    draw.line((xs[i], top + row1_h, xs[i], top + row1_h + row2_h), fill=BORDER, width=2)

last = body[3]
tw = draw.textlength(last, font=F_TBL)
draw.text(((xs[3] + xs[4] - tw) / 2, top + row1_h + (row2_h - 24) / 2), last, font=F_TBL, fill=BLACK)

canvas.save("/Users/zouxiaoli/Documents/ChatGPT/ai询价/.qa/template_build/preview.png")
print("preview written")
