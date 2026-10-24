"""Build the standalone Lanhu interaction sheet for AI 查价 copy states."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path("/Users/zouxiaoli/Documents/ChatGPT/ai询价")
sys.path.insert(0, str(ROOT / "tools"))

from build_ai_xunjia_lanhu_sheets import OUT_DIR, build_pc_sheet  # noqa: E402


PRICE_CHECK_STEPS = [
    {
        "title": "AI 查价默认快速思考",
        "shot": "../ai_price_check_doc/pc/01_price_check_default.png",
        "trigger": "切换至「AI 查价」",
        "notes": [
            "默认选中「快速思考」，不弹窗、不增加首次提交成本。",
            "输入区左下角常驻当前模式，用户可直接查看并切换；AI 询价仍仅保留快速思考。",
        ],
    },
    {
        "title": "思考模式菜单",
        "shot": "../ai_price_check_doc/pc/02_mode_menu.png",
        "trigger": "点击输入框左下角「快速思考」",
        "notes": [
            "快速思考：常规材料与地区价格查询，标记「推荐」。",
            "深度思考：复杂规格与多条件分析，标记「本地深度」；菜单在输入框上方展开，不遮挡输入内容。",
        ],
    },
    {
        "title": "首次切换深度思考",
        "shot": "../ai_price_check_doc/pc/03_deep_confirm.png",
        "trigger": "首次选择「深度思考」",
        "notes": [
            "标题：切换为深度思考？",
            "说明：深度思考会使用本地模型进行多步推理，适合复杂规格、多条件比较和需要完整判断依据的查价任务。",
            "按钮：保持快速思考 / 切换并继续。当前会话首次确认后，后续切换不再重复弹窗。",
        ],
    },
    {
        "title": "快速思考中",
        "shot": "../ai_price_check_doc/pc/04_quick_thinking.png",
        "trigger": "快速思考模式下提交查价条件",
        "notes": [
            "查询中仅展示「快速思考中」和进度条，不展示步骤、预计时间或取消入口。",
            "通宝提示：快速模型正在检索最新市场数据，我先帮你拿到一个可用参考价～",
        ],
    },
    {
        "title": "快速思考查价结果",
        "shot": "../ai_price_check_doc/pc/05_quick_result.png",
        "trigger": "快速思考查询完成",
        "notes": [
            "结果标记「快速思考 / 线上快速」，展示市场参考价、广州市均价、合理区间和一句简短结论。",
            "通宝提示：结果出来啦～看下结果是否满足要求，有疑问可以继续问我。",
        ],
    },
    {
        "title": "深度思考中",
        "shot": "../ai_price_check_doc/pc/06_deep_thinking.png",
        "trigger": "深度思考模式下提交查价条件",
        "notes": [
            "展示「深度思考中 / 本地深度」、已运行时长、预计剩余时间、四个处理步骤和「取消查询」。",
            "步骤：理解材料属性与查价条件 → 多维度匹配价格数据与历史区间 → 交叉校验周边地区价格差异 → 整理多维度价格对比依据。",
            "提示：深度思考需要完成多轮推理。提交后可离开页面，完成后会通知你，无需重复提交。",
        ],
    },
    {
        "title": "深度思考查价结果",
        "shot": "../ai_price_check_doc/pc/07_deep_result.png",
        "trigger": "深度思考查询完成",
        "notes": [
            "结果标记「深度思考 / 本地深度」，展示综合参考价、参考下限、参考上限和简短区域差异说明。",
            "结果区不展示底层模型、接口、供应商信息，也不冒用快速思考的结论口径。",
            "通宝提示：结果出来啦～看下结果是否满足要求，有疑问可以继续问我。",
        ],
    },
]


def main() -> None:
    out_path = OUT_DIR / "03_PC端_AI查价双模式文案_交互说明.png"
    build_pc_sheet(
        PRICE_CHECK_STEPS,
        out_path,
        title="AI 查价 PC 端 · 双模式与思考文案交互说明",
        subtitle="来源：AI 查价模型切换项目与当前 demo 实际截图 · 共 7 步 · 独立于 AI 询价长图",
    )


if __name__ == "__main__":
    main()
