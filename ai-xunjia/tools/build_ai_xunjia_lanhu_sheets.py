"""Build Lanhu-ready interaction spec sheets for AI 询价.

Output: generated/lanhu_demo_based_big_sheets/
  00_AI询价业务流程图.png            (rendered by render_ai_inquiry_flow.js)
  01_PC端_基于demo截图_AI询价交互说明.png
  02_PC端_项目阶段多轮问答_交互说明.png
"""

from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path("/Users/zouxiaoli/Documents/ChatGPT/ai询价")
SHOTS = ROOT / ".qa/ai_xunjia_doc"
OUT_DIR = ROOT / "generated/lanhu_demo_based_big_sheets"
OUT_DIR.mkdir(parents=True, exist_ok=True)

CANVAS_W = 3800
BG = "#eef3f8"
CARD_BG = "#ffffff"
CARD_BORDER = "#c8d6e8"
ACCENT = "#1f78ff"
ACCENT_SOFT = "#e8f1ff"
TEXT = "#14243a"
MUTED = "#5b6b80"
TRIGGER_BG = "#f2f7ff"
TRIGGER_BORDER = "#c9dcf7"

FONT_PATH = "/System/Library/Fonts/PingFang.ttc"


def font(size, bold=False, index=None):
    if index is None:
        index = 1 if bold else 0
    try:
        return ImageFont.truetype(FONT_PATH, size, index=index)
    except Exception:
        return ImageFont.truetype(FONT_PATH, size)


F_TITLE = font(64, True)
F_SUBTITLE = font(34)
F_STEP_NO = font(32, True)
F_STEP_TITLE = font(40, True)
F_LABEL = font(28, True)
F_BODY = font(27)
F_SMALL = font(24)


def wrap(draw, text, fnt, max_width):
    lines = []
    for raw in text.split("\n"):
        if not raw:
            lines.append("")
            continue
        current = ""
        for char in raw:
            trial = current + char
            if draw.textlength(trial, font=fnt) <= max_width:
                current = trial
            else:
                lines.append(current)
                current = char
        lines.append(current)
    return lines


def fit_image(image_path: Path, box_w: int, box_h: int, crop_body: bool = False):
    image = Image.open(image_path).convert("RGB")
    if crop_body:
        width, height = image.size
        image = image.crop((0, 0, width, height))
    ratio = min(box_w / image.width, box_h / image.height)
    size = (max(1, int(image.width * ratio)), max(1, int(image.height * ratio)))
    return image.resize(size, Image.LANCZOS)


def draw_shot(canvas, draw, image, x, y, box_w, box_h):
    offset_x = x + (box_w - image.width) // 2
    offset_y = y + (box_h - image.height) // 2
    shadow = Image.new("RGBA", (image.width + 24, image.height + 24), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle(
        (12, 12, image.width + 12, image.height + 12), radius=18, fill=(31, 61, 104, 26)
    )
    canvas.paste(shadow, (offset_x - 12, offset_y - 12), shadow)
    draw.rounded_rectangle(
        (offset_x - 1, offset_y - 1, offset_x + image.width, offset_y + image.height),
        radius=12,
        outline="#d7e2f0",
        width=2,
    )
    canvas.paste(image, (offset_x, offset_y))
    return offset_x, offset_y


def build_pc_sheet(
    steps,
    out_path: Path,
    title="AI 询价 PC 端交互说明",
    subtitle=None,
    card_h=1310,
    shot_box_h=930,
):
    margin = 80
    gap = 60
    card_w = (CANVAS_W - margin * 2 - gap) // 2
    header_h = 300
    pad = 34
    rows = (len(steps) + 1) // 2
    canvas_h = header_h + rows * (card_h + 44) + 60
    canvas = Image.new("RGB", (CANVAS_W, canvas_h), BG)
    draw = ImageDraw.Draw(canvas)

    draw.text((margin, 84), title, font=F_TITLE, fill=TEXT)
    if subtitle is None:
        subtitle = (
            "来源：造价通 AI 询价演示包实际页面截图 · 共 {} 步 · 交互说明与需求文档 3.4 章节一一对应".format(len(steps))
        )
    draw.text(
        (margin, 176),
        subtitle,
        font=F_SUBTITLE,
        fill=MUTED,
    )
    draw.rounded_rectangle(
        (margin, 240, margin + 260, 252), radius=6, fill=ACCENT
    )

    for index, step in enumerate(steps):
        row = index // 2
        col = index % 2
        x = margin + col * (card_w + gap)
        y = header_h + row * (card_h + 44)
        draw.rounded_rectangle(
            (x, y, x + card_w, y + card_h),
            radius=24,
            fill=CARD_BG,
            outline=CARD_BORDER,
            width=3,
        )
        pill = f"{index + 1:02d}"
        pill_w = 108
        draw.rounded_rectangle(
            (x + pad, y + pad, x + pad + pill_w, y + pad + 62), radius=31, fill=ACCENT
        )
        tw = draw.textlength(pill, font=F_STEP_NO)
        draw.text(
            (x + pad + (pill_w - tw) / 2, y + pad + 12), pill, font=F_STEP_NO, fill="#ffffff"
        )
        draw.text((x + pad + pill_w + 22, y + pad + 8), step["title"], font=F_STEP_TITLE, fill=TEXT)

        shot_y = y + pad + 84
        image = fit_image(SHOTS / step["shot"], card_w - pad * 2, shot_box_h)
        draw_shot(canvas, draw, image, x + pad, shot_y, card_w - pad * 2, shot_box_h)

        band_y = shot_y + shot_box_h + 18
        draw.rounded_rectangle(
            (x + pad, band_y, x + card_w - pad, band_y + 54),
            radius=12,
            fill=TRIGGER_BG,
            outline=TRIGGER_BORDER,
            width=2,
        )
        trigger = step["trigger"]
        draw.text((x + pad + 18, band_y + 12), "触发动作", font=F_LABEL, fill=ACCENT)
        draw.text(
            (x + pad + 18 + draw.textlength("触发动作", font=F_LABEL) + 18, band_y + 13),
            trigger,
            font=F_BODY,
            fill=TEXT,
        )

        text_y = band_y + 74
        max_w = card_w - pad * 2 - 10
        for line in step["notes"]:
            wrapped = wrap(draw, line, F_BODY, max_w)
            for sub in wrapped:
                if not sub:
                    text_y += 14
                    continue
                draw.text((x + pad, text_y), sub, font=F_BODY, fill=MUTED)
                text_y += 38
            text_y += 6

    canvas.save(out_path)
    print(out_path)


PC_STEPS = [
    {
        "title": "首页 AI 询价入口",
        "shot": "pc/01_home.png",
        "trigger": "点击首页搜索区「AI 询价」",
        "notes": [
            "入口与既有「AI 查价」并列，使用通宝头像标识。",
            "点击进入对话式询价页，不改动首页其他模块。",
        ],
    },
    {
        "title": "对话页初始状态",
        "shot": "pc/02_workspace_empty.png",
        "trigger": "进入对话页",
        "notes": [
            "展示通宝问候语与输入提示：材料名称、规格型号、数量、单位、报价地区、目标时间与项目阶段。",
            "通宝提示语仅作要素引导，不提供复制入口。",
            "输入区固定展示「快速思考（推荐）」模式与通宝形象；左侧底部展示每日额度：AI 询价 10 条/天、AI 查价 50 条/天。",
            "页面右上角提供「造价通首页」返回入口。",
        ],
    },
    {
        "title": "提交询价条件",
        "shot": "pc/03_thinking_prepare.png",
        "trigger": "输入条件后回车或点击发送",
        "notes": [
            "示例：广东广州，2026年9月，普通硅酸盐水泥 P.O42.5，数量200吨。",
            "提交后进入准备态，通宝提示正在接收询价需求。",
            "本次查询扣减 1 条使用次数。",
        ],
    },
    {
        "title": "通宝思考状态",
        "shot": "pc/04_thinking_steps.png",
        "trigger": "等待识别与匹配",
        "notes": [
            "依次展示理解材料条件、匹配历史询价与市场价等步骤。",
            "通宝形象轻微上下浮动、左右停顿，呈现思考状态。",
            "思考结束先生成询价建议结果，项目阶段提示位于结果之后；减少动态效果时降级为静态展示。",
        ],
    },
    {
        "title": "询价建议结果（3 条）",
        "shot": "pc/05_result_cement.png",
        "trigger": "匹配完成后自动输出",
        "notes": [
            "副标题：通宝给您推荐 3 条最优供应商报价。",
            "展示建议价区间、材料规格、数量单位、品牌、建议价、供应商与回复备注。",
            "结果区固定展示「转人工」引导，不再提供「修改条件」入口。",
            "结果生成后，通宝在结果卡片下方提示确认本次询价所属项目阶段。",
        ],
    },
    {
        "title": "搜索条件不足提示",
        "shot": "pc/21_insufficient_conditions.png",
        "trigger": "输入条件项少于 5 项",
        "notes": [
            "在建议价区间旁展示提示语「由于搜索条件不足，建议价区间或存在偏差」。",
            "结果照常展示，不阻断查询、不提示用户重新询价。",
        ],
    },
    {
        "title": "候选结果不足",
        "shot": "pc/09_result_insufficient.png",
        "trigger": "候选少于 3 条时",
        "notes": [
            "副标题切换为「通宝给您展示全部 2 条候选供应商报价」。",
            "直接展示全部候选，不提示用户重新询价。",
            "建议价区间按实际候选计算。",
        ],
    },
    {
        "title": "结果操作与反馈",
        "shot": "pc/06_result_feedback.png",
        "trigger": "点击复制 / 点赞 / 点踩 / 重新查询",
        "notes": [
            "复制：写入建议价区间与供应商价格并提示已复制。",
            "点赞、点踩互斥，选中后图标颜色加深，再次点击取消。",
            "重新查询：提示「点击将重新生成询价结果，并消耗1条询价使用次数」，重新生成时由后端通知第三方模型重新筛选候选供应商。",
        ],
    },
    {
        "title": "重新查询更换候选供应商",
        "shot": "pc/22_result_refreshed.png",
        "trigger": "点击结果区「重新查询」图标",
        "notes": [
            "重新生成后结果区提示「已通知后端并触发第三方模型重新筛选候选供应商」。",
            "刷新前后候选供应商与报价数据发生更换，建议价区间同步更新。",
            "重新查询扣减 1 条 AI 询价使用次数。",
        ],
    },
    {
        "title": "条件覆盖与二次询价",
        "shot": "pc/24_condition_override.png",
        "trigger": "首轮询价完成后再修改地区或其他条件",
        "notes": [
            "仅修改地区时，保留原材料、规格、数量、单位与目标时间，新地区覆盖旧地区。",
            "二次询价生成新 requestId、扣减 1 次，旧候选与导出引用失效；会话只展示用户真实输入。",
        ],
    },
    {
        "title": "供应商联系方式追问",
        "shot": "pc/07_supplier_contact.png",
        "trigger": "再次输入「供应商联系方式」",
        "notes": [
            "通宝在对话内返回对应供应商的联系方式。",
            "答复内容以结果记录为准，不额外编造联系方式。",
        ],
    },
    {
        "title": "联系方式不可用兜底",
        "shot": "pc/08_supplier_unavailable.png",
        "trigger": "反馈「供应商电话不准确」",
        "notes": [
            "按报价记录时间较久或供应商已更换电话答复。",
            "答复内提供「转人工询价」按钮，由工作人员核验最新联系方式。",
        ],
    },
    {
        "title": "非材料询价输入提示",
        "shot": "pc/12_invalid_input.png",
        "trigger": "输入与材料询价无关的内容",
        "notes": [
            "返回「暂未识别为材料询价需求」提示并给出输入示例。",
            "不进入查询流程，不扣减使用次数。",
        ],
    },
    {
        "title": "缺少规格与补充规格",
        "shot": "pc/11_incomplete_spec.png",
        "trigger": "识别到缺少关键规格",
        "notes": [
            "展示公称直径、压力等级等可选项。",
            "确认补充后按新条件重新查询，或选择转人工询价。",
        ],
    },
    {
        "title": "无可用数据与转人工",
        "shot": "pc/10_no_data.png",
        "trigger": "历史询价与市场价均无结果",
        "notes": [
            "展示暂无法自动查询到可用数据，说明可能超出可查询时间范围或未收录。",
            "提供调整目标时间、扩大报价区域与转人工询价三种出口。",
            "不生成建议价与供应商报价。",
        ],
    },
    {
        "title": "转人工询价统一跳转",
        "shot": "pc/16_manual_inquiry_jump.png",
        "trigger": "点击任意「转人工询价」入口",
        "notes": [
            "结果区、无数据、缺少规格、联系方式不可用等场景统一跳转造价通人工询价页面。",
            "由造价通页面承接登录、联系方式和询价要求填写；demo 不收集联系人、电话，不生成工单或受理状态。",
        ],
    },
    {
        "title": "盖章报价预览与下载",
        "shot": "pc/13_preview_seal.png",
        "trigger": "点击「导出盖章报价文件」",
        "notes": [
            "点击「加盖印章」后叠加固定位置询价专用章，按钮切换为「撤销盖章」。",
            "支持显示或隐藏联系方式，并在盖章后下载 PDF；页面保留造价通标识、水印与 AI 生成提示。",
        ],
    },
    {
        "title": "撤销盖章与下载拦截",
        "shot": "pc/13b_preview_unsealed.png",
        "trigger": "点击「撤销盖章」",
        "notes": [
            "印章立即隐藏，按钮恢复为「加盖印章」。",
            "撤销后再次下载会被拦截；重新盖章后才允许生成 PDF。",
        ],
    },
    {
        "title": "历史询价记录",
        "shot": "pc/14_history.png",
        "trigger": "点击左侧「历史记录」",
        "notes": [
            "展示查询时间、查价条件与操作列，不展示推荐结果与状态列。",
            "操作列提供「查看结果／演示对话」与「删除」入口；历史询价记录保留一年，超过一年自动清理。",
        ],
    },
    {
        "title": "删除会话二次确认",
        "shot": "pc/15_delete_confirm.png",
        "trigger": "会话右侧更多操作 → 删除",
        "notes": [
            "弹出确认框「确定删除对话？」。",
            "确认后移除会话条目并记录删除日志事件。",
        ],
    },
]

PHASE_STEPS = [
    {
        "title": "首次会话先展示通宝思考",
        "shot": "pc/17_phase_thinking.png",
        "trigger": "首次提交材料询价条件",
        "notes": [
            "首次会话先展示通宝思考状态与推理步骤，通宝形象保持轻微上下浮动、左右停顿。",
            "思考期间不提前弹出项目阶段选择，避免打断用户等待与过程反馈。",
            "思考结束后先生成询价建议结果，顺序为“通宝思考 → 询价建议结果 → 项目阶段提示选择 → 按阶段系数更新建议价”。",
        ],
    },
    {
        "title": "结果生成后确认项目阶段",
        "shot": "pc/18_phase_question.png",
        "trigger": "询价建议结果生成后",
        "notes": [
            "项目阶段提示位于询价建议结果之后，在结果卡片下方发起一轮确认。",
            "选项为施工图预算（默认）、投资估算、设计概算、工程竣工结算；用户未选择时默认按施工图预算阶段处理。",
            "阶段提示只在新建会话出现一次，同一会话内修改后再次提交或重新查询沿用已选阶段。",
        ],
    },
    {
        "title": "按所选阶段计算建议价",
        "shot": "pc/19_phase_estimate_result.png",
        "trigger": "用户选择项目阶段后",
        "notes": [
            "后端按建议价乘以对应项目阶段系数重新计算并刷新结果，投资估算示例结果为基准价 × 1.15。",
            "结果区只展示计算后的建议价区间与报价行，不展示系数与计算过程。",
        ],
    },
    {
        "title": "未填写时默认预算阶段",
        "shot": "pc/20_phase_default_result.png",
        "trigger": "跳过 / 未填写项目阶段",
        "notes": [
            "用户未选择项目阶段时默认按施工图预算阶段处理，系数为 1。",
            "建议价按默认阶段计算后直接展示，用户侧不感知阶段系数。",
        ],
    },
]


def build_phase_sheet():
    build_pc_sheet(
        PHASE_STEPS,
        OUT_DIR / "02_PC端_项目阶段多轮问答_交互说明.png",
        title="AI 询价 PC 端 · 项目阶段多轮问答交互说明",
        subtitle="来源：造价通 AI 询价演示包实际页面截图 · 共 {} 步 · 先思考并生成询价建议结果，再采集项目阶段，未填写默认按施工图预算阶段".format(
            len(PHASE_STEPS)
        ),
    )


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "pc"
    if mode == "phase":
        build_phase_sheet()
    else:
        build_pc_sheet(PC_STEPS, OUT_DIR / "01_PC端_基于demo截图_AI询价交互说明.png")
