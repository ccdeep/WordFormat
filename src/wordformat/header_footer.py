#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : header_footer.py
"""页眉页脚格式识别与修正（issue#93）。

Word 的页眉页脚存在于各节（section）的 header/footer part 中，不属于正文
段落流，模型分类与 FormatNode 树均不覆盖。本模块通过 hook 机制接入流水线
（见 handlers.py 中注册的 builtin.header_footer 回调，触发于
before_document_save 事件）：

- check 模式：检查每节页眉/页脚的字号、中文字体、英文字体、对齐方式及
  内容缺失，问题作为批注锚定在文档第一段（页眉/页脚 run 无法挂批注——
  python-docx 限制 comments 只能出现在 document part）；
- apply 模式：按配置修正格式，残留差异仍以批注报告。

配置段（YAML 可选，缺省用 DEFAULT_HF_CONFIG）::

    header_footer:
      enabled: true
      header:
        chinese_font_name: '宋体'
        english_font_name: 'Times New Roman'
        font_size: '小五'
        alignment: '居中对齐'
        require_content: false
      footer: ...

仅检查四项属性：字号、中文字体（eastAsia）、英文字体（ascii）、对齐。
不检查加粗/颜色/行距等易误报属性，避免对页码域等特殊内容误伤。
"""

from __future__ import annotations

import copy
from typing import List, Optional

from loguru import logger

from wordformat.style.defs import Alignment, FontSize
from wordformat.style.reader import (
    paragraph_get_alignment,
    run_get_font_name,
    run_get_font_name_en,
    run_get_font_size_pt,
)
from wordformat.utils import has_chinese

# 页眉/页脚缺省规范（本科论文常见：小五号宋体、Times New Roman、居中）
DEFAULT_HF_CONFIG = {
    "enabled": True,
    "header": {
        "chinese_font_name": "宋体",
        "english_font_name": "Times New Roman",
        "font_size": "小五",  # 9pt
        "alignment": "居中对齐",
        "require_content": False,
    },
    "footer": {
        "chinese_font_name": "宋体",
        "english_font_name": "Times New Roman",
        "font_size": "小五",  # 9pt
        "alignment": "居中对齐",
        "require_content": False,
    },
}


def resolve_hf_config(config_model=None) -> dict:
    """从配置模型提取 header_footer 段，与缺省规范深合并（用户配置优先）。

    config_model 可为 NodeConfigRoot / dict / None；禁用（enabled=false）
    返回空 dict，调用方应跳过处理。
    """
    if config_model is None:
        merged = copy.deepcopy(DEFAULT_HF_CONFIG)
    else:
        hf = (
            config_model.get("header_footer")
            if isinstance(config_model, dict)
            else getattr(config_model, "header_footer", None)
        )
        if not isinstance(hf, dict):
            hf = {}
        from wordformat.config.dotdict import deep_merge

        merged = deep_merge(copy.deepcopy(DEFAULT_HF_CONFIG), hf)
    return merged if merged.get("enabled", True) else {}


def _part_label(kind: str, sec_idx: int) -> str:
    """节标签：页眉(第1节)。"""
    return f"{kind}(第{sec_idx}节)"


def _resolve_alignment_value(label):
    """对齐标签 → WD_ALIGN_PARAGRAPH 枚举值；解析失败返回 None。"""
    try:
        return Alignment(label).rel_value
    except (ValueError, TypeError):
        return None


def _resolve_font_size_pt(label) -> Optional[float]:
    """字号标签 → pt；支持 '小五' / 12 / '12pt' 等。"""
    if label is None:
        return None
    try:
        if isinstance(label, str):
            return float(FontSize(label).rel_value)
        return float(label)
    except (ValueError, TypeError):
        return None


def _check_paragraph(paragraph, kind: str, sec_idx: int, cfg: dict, issues: List[str]):
    """检查单个页眉/页脚段落（非空）：对齐 + 字号 + 中英文字体。"""
    pos = _part_label(kind, sec_idx)

    # 对齐（有效值沿继承链解析，None 表示未设置 → 跳过避免误报）
    expected_alignment = _resolve_alignment_value(cfg.get("alignment"))
    if expected_alignment is not None:
        actual = paragraph_get_alignment(paragraph)
        if actual is not None and actual != expected_alignment:
            issues.append(
                f"{pos}-对齐错误：当前非{cfg.get('alignment')}，规范:{cfg.get('alignment')}"
            )

    # 字符属性：字号 + 中/英文字体（每段各检查一次即可，取首个非空 run）
    expected_size = _resolve_font_size_pt(cfg.get("font_size"))
    expected_cn = (cfg.get("chinese_font_name") or "").lower()
    expected_en = (cfg.get("english_font_name") or "").lower()
    if not (expected_size or expected_cn or expected_en):
        return

    checked_cn = checked_en = checked_size = False
    for run in paragraph.runs:
        if not run.text.strip():
            continue
        if expected_size and not checked_size:
            actual_size = run_get_font_size_pt(run)
            if actual_size != expected_size:
                issues.append(
                    f"{pos}-字号错误：当前{actual_size:g}pt，规范:{cfg.get('font_size')}"
                )
            checked_size = True
        if expected_cn and not checked_cn and has_chinese(run.text):
            actual = (run_get_font_name(run) or "").lower()
            if actual != expected_cn:
                issues.append(
                    f"{pos}-中文字体错误：当前{actual or '未设置'}，"
                    f"规范:{cfg.get('chinese_font_name')}"
                )
            checked_cn = True
        if expected_en and not checked_en:
            actual = (run_get_font_name_en(run) or "").lower()
            if actual != expected_en:
                issues.append(
                    f"{pos}-英文字体错误：当前{actual or '未设置'}，"
                    f"规范:{cfg.get('english_font_name')}"
                )
            checked_en = True
        if checked_cn and checked_en and checked_size:
            break


def _check_part(part, kind: str, sec_idx: int, cfg: dict, issues: List[str]):
    """检查某节的一个 header/footer 部件。"""
    # 第 2 节起与前一节链接时沿用上一节格式，跳过；
    # 第 1 节无前驱，linked=True 仅表示未创建独立 part（无内容），
    # 由 require_content 决定是否提示缺失
    if sec_idx > 1 and getattr(part, "is_linked_to_previous", False):
        return
    content_paras = [p for p in part.paragraphs if p.text.strip()]
    if not content_paras:
        if cfg.get("require_content", False):
            issues.append(
                f"{_part_label(kind, sec_idx)}-内容缺失：当前为空，"
                f"规范：应包含{kind}内容"
            )
        return
    for p in content_paras:
        _check_paragraph(p, kind, sec_idx, cfg, issues)


def check_header_footer(document, config_model=None) -> List[str]:
    """检查文档所有节的页眉页脚格式，返回标准批注文案列表。

    Args:
        document: python-docx Document 对象
        config_model: 配置模型（NodeConfigRoot/dict），None 用缺省规范

    Returns:
        问题文案列表（空表示无问题或未启用）
    """
    hf_config = resolve_hf_config(config_model)
    if not hf_config:
        return []
    issues: List[str] = []
    for idx, section in enumerate(document.sections, start=1):
        _check_part(section.header, "页眉", idx, hf_config["header"], issues)
        _check_part(section.footer, "页脚", idx, hf_config["footer"], issues)
    logger.debug(f"页眉页脚检查完成：{len(issues)} 个问题")
    return issues


def _apply_run_format(run, cfg: dict) -> None:
    """修正单个 run 的字号与中/英文字体。"""
    if not run.text.strip():
        return
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    expected_size = _resolve_font_size_pt(cfg.get("font_size"))
    if expected_size:
        run.font.size = Pt(expected_size)

    cn_name = cfg.get("chinese_font_name")
    en_name = cfg.get("english_font_name")
    if cn_name or en_name:
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        if cn_name:
            rFonts.set(qn("w:eastAsia"), cn_name)
        if en_name:
            rFonts.set(qn("w:ascii"), en_name)
            rFonts.set(qn("w:hAnsi"), en_name)


def _apply_paragraph(
    paragraph, kind: str, sec_idx: int, cfg: dict, remaining: List[str]
):
    """修正单个页眉/页脚段落格式；修正后复检，残留问题写入 remaining。"""
    expected_alignment = _resolve_alignment_value(cfg.get("alignment"))
    if expected_alignment is not None:
        paragraph.alignment = expected_alignment
    for run in paragraph.runs:
        _apply_run_format(run, cfg)

    leftover: List[str] = []
    _check_paragraph(paragraph, kind, sec_idx, cfg, leftover)
    remaining.extend(leftover)


def _apply_part(part, kind: str, sec_idx: int, cfg: dict, remaining: List[str]):
    """修正某节 header/footer 部件；残留问题写入 remaining。"""
    # 第 2 节起与前一节链接时沿用上一节格式，跳过（第 1 节无前驱）
    if sec_idx > 1 and getattr(part, "is_linked_to_previous", False):
        return
    content_paras = [p for p in part.paragraphs if p.text.strip()]
    if not content_paras:
        return  # 空内容不自动补内容（页码域等由用户在 Word 中插入）
    for p in content_paras:
        _apply_paragraph(p, kind, sec_idx, cfg, remaining)


def apply_header_footer(document, config_model=None) -> List[str]:
    """按配置修正文档所有节的页眉页脚格式。

    Returns:
        修正后仍残留的问题文案列表（空表示全部修正或未启用）。
    """
    hf_config = resolve_hf_config(config_model)
    if not hf_config:
        return []
    remaining: List[str] = []
    for idx, section in enumerate(document.sections, start=1):
        _apply_part(section.header, "页眉", idx, hf_config["header"], remaining)
        _apply_part(section.footer, "页脚", idx, hf_config["footer"], remaining)
    if remaining:
        logger.debug(f"页眉页脚修正完成，残留 {len(remaining)} 个问题")
    return remaining
