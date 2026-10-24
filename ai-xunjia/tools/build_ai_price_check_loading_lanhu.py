"""Build one Lanhu sheet for the current AI price-check loading states."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path("/Users/zouxiaoli/Documents/ChatGPT/ai询价")
sys.path.insert(0, str(ROOT / "tools"))

from build_ai_xunjia_lanhu_sheets import OUT_DIR, build_pc_sheet  # noqa: E402


LOADING_STEPS = [
    {
        "title": "AI 查价 · 快速思考中",
        "shot": "../ai_price_check_doc/pc_current/08_quick_loading_current.png",
        "trigger": "快速思考模式下提交查价条件",
        "notes": [
            "加载态与 AI 询价共用轻量思考样式：标题、状态文案左对齐，进度条位于文案下方。",
            "状态依次展示识别材料条件、检索市场价与信息价、核算同规格价格区间。",
            "快速思考不展示步骤列表、预计耗时或取消入口。",
        ],
    },
    {
        "title": "AI 查价 · 深度思考中",
        "shot": "../ai_price_check_doc/pc_current/09_deep_loading_current.png",
        "trigger": "深度思考模式下提交查价条件",
        "notes": [
            "深度思考沿用同一套思考组件，以圆点步骤呈现理解条件、匹配价格、校验地区差异和整理对比依据。",
            "已完成步骤显示绿色圆点，当前步骤高亮；状态文案随步骤推进而变化。",
            "不展示本地深度标签、预计耗时、离开页面提示或取消入口。",
        ],
    },
]


def main() -> None:
    out_path = OUT_DIR / "04_PC端_AI查价思考加载态_交互说明.png"
    build_pc_sheet(
        LOADING_STEPS,
        out_path,
        title="AI 查价 PC 端 · 思考加载态交互说明",
        subtitle="来源：当前 ai询价demo.html 实际截图 · 快速思考 / 深度思考统一样式 · 共 2 个状态",
        card_h=1360,
        shot_box_h=960,
    )


if __name__ == "__main__":
    main()
