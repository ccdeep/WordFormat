#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""公式段落排版（需求文档 D4.4 + 基准 08-公式）：

- 带编号（段落非数学文本以（n）/(n) 结尾）→ 制表位法：居中制表位=版心宽/2、
  右制表位=版心宽，公式前/编号前各插一个制表符；oMathPara 先解包为行内 oMath
  （制表位法要求公式为行内级，与 Word 手工"公式+编号同排"的做法一致）
- 无编号 → 段落居中（学位论文惯例；oMathPara 显示公式 Word 本就居中，显式声明无害）

OMML 本体（m:oMath 内部任何节点）不修改；只新增兄弟 run 与 pPr 属性。
"""
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu

# 段尾编号形态：（3-1）/ (2) / （12）等
_NUMBER_TAIL_RE = re.compile(r"[（(]\s*\d[-\d–.]*\s*[)）]\s*$")
_EMU_PER_TWIP = 635


def _make_tab_run():
    r = OxmlElement("w:r")
    r.append(OxmlElement("w:tab"))
    return r


def _is_tab_run(r):
    """run 是否只含制表符（无 w:t）。"""
    return (r.tag == qn("w:r") and r.find(qn("w:tab")) is not None
            and r.find(qn("w:t")) is None)


def _prev_sibling(el):
    prev = el.getprevious()
    return prev


def _content_width_twips(paragraph) -> int:
    """版心宽（twips）= 页宽 − 左右页边距（取文档最后一个分节）。"""
    document = paragraph.part.document
    section = document.sections[-1]
    width_emu = (section.page_width or 0) - (section.left_margin or 0) - (section.right_margin or 0)
    return max(int(width_emu / _EMU_PER_TWIP), 1000)


def _unwrap_omathpara(paragraph):
    """oMathPara 解包为行内 oMath（保留其内部的 oMath，丢弃段落级对齐设置）。"""
    p = paragraph._p
    for omp in p.findall(qn("m:oMathPara")):
        for child in list(omp):
            if child.tag == qn("m:oMath"):
                omp.addprevious(child)
        p.remove(omp)


def format_equation_paragraph(paragraph):
    """按是否有段尾编号套用排版。返回 'tab' / 'center' / 'skip'。"""
    text = (paragraph.text or "").strip()  # python-docx text 不含数学内容
    numbered = bool(_NUMBER_TAIL_RE.search(text))
    if numbered:
        _apply_tab_numbering(paragraph)
        return "tab"
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return "center"


def _apply_tab_numbering(paragraph):
    p = paragraph._p
    width_tw = _content_width_twips(paragraph)

    # 清对齐与首行缩进（制表位法要求段落左对齐、无首行缩进）
    pPr = p.get_or_add_pPr()
    for tag in ("w:jc",):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    ind = pPr.find(qn("w:ind"))
    if ind is not None:
        pPr.remove(ind)

    # 双制表位：居中（版心/2）+ 右（版心）
    tab_stops = paragraph.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Emu(width_tw // 2 * _EMU_PER_TWIP), WD_TAB_ALIGNMENT.CENTER)
    tab_stops.add_tab_stop(Emu(width_tw * _EMU_PER_TWIP), WD_TAB_ALIGNMENT.RIGHT)

    _unwrap_omathpara(paragraph)

    # 公式前插制表符
    first_math = p.find(qn("m:oMath"))
    if first_math is not None:
        prev = _prev_sibling(first_math)
        if prev is None or not _is_tab_run(prev):
            first_math.addprevious(_make_tab_run())

    # 编号前插制表符：找到段尾编号所在的 w:t，在其 run 前插入
    for r in p.findall(qn("w:r")):
        text = "".join(t.text or "" for t in r.findall(qn("w:t")))
        if _NUMBER_TAIL_RE.search(text.strip()):
            prev = _prev_sibling(r)
            if prev is None or not _is_tab_run(prev):
                r.addprevious(_make_tab_run())
            break
