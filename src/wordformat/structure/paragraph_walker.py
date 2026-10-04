#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""文档体段落遍历器：识别（base.DocxBase.parse）与格式化对齐
（pipeline.stages.ParagraphAlignmentStage）共用，保证两处段落序列一一对应。

python-docx 的 document.paragraphs 只含表外段落——表格内容整体不可见
（缺口清单 P0-2）。本遍历器按文档顺序给出全部段落，含表格单元格内
段落（行优先、合并单元格去重、嵌套表格递归），并附带 in_table 标记。
"""

from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

# OMML 公式的 XML 元素名
_OMATH_TAG = qn("m:oMath")


def iter_document_paragraphs(document):
    """按文档顺序遍历文档体内全部段落。

    :returns: list[tuple[Paragraph, bool]] —— (段落对象, 是否位于表格内)。
        表内段落按 行优先、合并单元格去重、嵌套表格递归 的顺序展开。
    """
    result = []

    def _collect_table(table: Table) -> None:
        for row in table.rows:
            seen_tcs = set()
            for cell in row.cells:
                tc = cell._tc
                if tc in seen_tcs:  # 合并单元格在 row.cells 中重复出现
                    continue
                seen_tcs.add(tc)
                for para in cell.paragraphs:
                    result.append((para, True))
                for nested in cell.tables:
                    _collect_table(nested)

    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            result.append((Paragraph(child, document), False))
        elif child.tag == qn("w:tbl"):
            _collect_table(Table(child, document))
    return result


def paragraph_has_omath(paragraph) -> bool:
    """段落是否包含 OMML 公式（m:oMath，含 oMathPara 包裹的情况）。"""
    return any(True for _ in paragraph._p.iter(_OMATH_TAG))
