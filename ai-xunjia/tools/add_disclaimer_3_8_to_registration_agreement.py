"""Insert 3.8（关于造价通 AI 生成的免责声明）into 注册协议.docx.

The new paragraph is a deep copy of the existing 3.7 paragraph, so font
(PingFang SC / 14pt), first-line indent and line spacing stay identical.
The original file is backed up to .qa before writing.
"""

from __future__ import annotations

import copy
import shutil
from pathlib import Path

from docx import Document
from docx.text.paragraph import Paragraph

ROOT = Path("/Users/zouxiaoli/Documents/ChatGPT/ai询价")
SRC = Path("/Users/zouxiaoli/Desktop/注册协议.docx")
BACKUP = ROOT / ".qa/注册协议_增加3.8前备份.docx"

TEXT = (
    "3.8关于造价通 AI 生成的免责声明：造价通网站提供的 AI 查价、AI 询价及其他由人工智能生成的"
    "结果与文件（含导出的报价文件、询价单等），均由造价通 AI 基于 building DS 模型自动生成，"
    "仅供参考，不构成报价承诺、价格证明、合同依据或任何专业意见，用户应自行核实并以供应商正式报价"
    "及双方签署的书面文件为准；用户使用上述 AI 生成数据所作出的任何判断、决策及由此产生的后果，"
    "由用户自行承担，中建普联不对该等 AI 生成内容及其使用结果承担任何责任。"
)


def main() -> None:
    BACKUP.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SRC, BACKUP)

    doc = Document(str(SRC))
    base = next(p for p in doc.paragraphs if p.text.strip().startswith("3.7"))

    new_element = copy.deepcopy(base._p)
    base._p.addnext(new_element)
    new_paragraph = Paragraph(new_element, base._parent)

    for run in new_paragraph.runs[1:]:
        run._element.getparent().remove(run._element)
    new_paragraph.runs[0].text = TEXT

    doc.save(str(SRC))
    print(f"backup: {BACKUP}")
    print(f"updated: {SRC}")


if __name__ == "__main__":
    main()
