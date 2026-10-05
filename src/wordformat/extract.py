#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""从参考文档的样式表反推 wordformat 格式方案（需求文档 D4.1 / 开发计划 M4 的 fork 实现）。

原理：解析参考 docx 的 styles.xml，按语义命名（如"02-一级标题"）或内置名（heading 1）
把段落样式映射到配置节（headings.level_1 / body.text / references.entry ...），
沿 basedOn 链解析有效属性（中英文字体、字号、加粗、对齐、行距、段前后、缩进），
转换为 wordformat YAML 词表，覆盖到基础预设之上生成新方案。

用法（CLI）：
    wordf extract -d 参考文档.docx -o 新方案.yaml [-b 基础预设.yaml]

未提取到的类别/字段回落基础预设（无基础预设时用内置默认值）。
"""
import re
import zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 语义匹配表：section → (名称关键词列表, styleId 前缀列表)
# 顺序即优先级；一个样式可喂多个 section（09- 同时是参考文献/致谢标题，07- 同图表题）
# 语义匹配表：section → (名称关键词列表, styleId 精确值列表)，顺序即优先级
_SECTIONS = [
    ("document.title", ["论文封面中文标题", "封面标题"], ["11-"]),
    ("abstract.english.title", ["英文摘要标题"], ["20-"]),
    ("abstract.english.keywords", ["英文摘要关键词"], ["21-"]),
    ("abstract.english.body", ["英文摘要正文"], ["22-"]),
    ("abstract.chinese.title", ["中文摘要标题", "摘要标题"], ["18-"]),
    ("abstract.chinese.keywords", ["摘要关键词", "关键词"], ["19-"]),
    ("headings.level_1", ["一级标题", "章标题"], ["02-"]),
    ("headings.level_2", ["二级标题"], ["03-"]),
    ("headings.level_3", ["三级标题"], ["04-"]),
    ("references.title", ["参考文献标题"], ["09-"]),
    ("references.entry", ["参考文献正文"], ["27-"]),
    ("acknowledgements.title", ["致谢标题", "致谢"], []),
    ("tables.text", ["表格内文字"], ["07-1"]),
    ("figures.caption", ["图题", "图注", "题注"], ["07-"]),
    ("tables.caption", ["表题", "表注"], []),
    ("math.block", ["公式"], ["08-"]),
    ("body.text", ["正文"], ["01-"]),
]
# 一个样式喂多个 section（模板里结论/参考文献/附录/致谢共用 09-，图/表题共用 07-）
_SHARED_BY_ID = {
    "09-": ["references.title", "acknowledgements.title"],
    "07-": ["figures.caption", "tables.caption"],
}
# 名称含这些词的样式不作为正文（避免"英文摘要正文"抢到 body.text）
_BODY_EXCLUDE = ("摘要", "英文", "目录", "关键词", "声明", "封面", "表格", "公式", "参考文献", "题注")
# 字号：磅 → 中文标签（wordformat FontSize 词表）
_PT_TO_LABEL = {26: "一号", 24: "小一", 22: "二号", 18: "小二", 16: "三号", 15: "小三",
                14: "四号", 12: "小四", 10.5: "五号", 9: "小五", 7.5: "六号", 5.5: "七号"}
_JC_TO_LABEL = {"both": "两端对齐", "distribute": "两端对齐", "center": "居中对齐",
                "left": "左对齐", "start": "左对齐", "right": "右对齐", "end": "右对齐"}

# 内置名兜底（语义名全未命中时使用）
_BUILTIN_FALLBACK = {
    "heading 1": "headings.level_1",
    "heading 2": "headings.level_2",
    "heading 3": "headings.level_3",
    "caption": "figures.caption",
    "normal": "body.text",
}


def classify(name: str, style_id: str):
    """语义分类：返回 section 路径列表。首个命中优先，共享样式（09-/07-）一对多。"""
    if style_id in _SHARED_BY_ID:
        return list(_SHARED_BY_ID[style_id])
    for section, name_keys, id_values in _SECTIONS:
        name_hit = any(k in (name or "") for k in name_keys)
        id_hit = style_id in id_values
        if not (name_hit or id_hit):
            continue
        if section == "body.text" and any(w in (name or "") for w in _BODY_EXCLUDE):
            continue
        return [section]
    lowered = (name or "").lower()
    if lowered in _BUILTIN_FALLBACK:
        return [_BUILTIN_FALLBACK[lowered]]
    return []


def _g(e, attr):
    v = e.get(W + attr)
    return v


def _parse_bool(val):
    if val is None:
        return None
    return str(val).lower() not in ("0", "false", "none", "off")


def _style_attrs(style_el, styles_by_id):
    """沿 basedOn 链合并出一个样式的有效属性字典（子级覆盖父级）。"""
    chain = []
    seen = set()
    cur = style_el
    while cur is not None and cur.get(W + "styleId") not in seen:
        seen.add(cur.get(W + "styleId"))
        chain.append(cur)
        based = cur.find(W + "basedOn")
        cur = styles_by_id.get(_g(based, "val")) if based is not None else None
        if len(seen) > 10:
            break

    attrs = {"east_asia": None, "ascii": None, "sz_pt": None, "bold": None,
             "jc": None, "line_rule": None, "line": None,
             "before_lines": None, "after_lines": None,
             "first_line_chars": None, "hanging_chars": None}
    for el in reversed(chain):  # 父级先写、子级覆盖
        rpr = el.find(W + "rPr")
        if rpr is not None:
            fonts = rpr.find(W + "rFonts")
            if fonts is not None:
                if _g(fonts, "eastAsia"):
                    attrs["east_asia"] = _g(fonts, "eastAsia")
                if _g(fonts, "ascii"):
                    attrs["ascii"] = _g(fonts, "ascii")
            sz = rpr.find(W + "sz")
            if sz is not None and _g(sz, "val"):
                attrs["sz_pt"] = float(_g(sz, "val")) / 2
            bold = rpr.find(W + "b")
            if bold is not None:
                attrs["bold"] = _parse_bool(_g(bold, "val"))
        ppr = el.find(W + "pPr")
        if ppr is not None:
            jc = ppr.find(W + "jc")
            if jc is not None and _g(jc, "val"):
                attrs["jc"] = _g(jc, "val")
            spacing = ppr.find(W + "spacing")
            if spacing is not None:
                if _g(spacing, "line"):
                    attrs["line"] = float(_g(spacing, "line"))
                if _g(spacing, "lineRule"):
                    attrs["line_rule"] = _g(spacing, "lineRule")
                if _g(spacing, "beforeLines"):
                    attrs["before_lines"] = float(_g(spacing, "beforeLines")) / 100
                if _g(spacing, "afterLines"):
                    attrs["after_lines"] = float(_g(spacing, "afterLines")) / 100
            ind = ppr.find(W + "ind")
            if ind is not None:
                if _g(ind, "firstLineChars"):
                    attrs["first_line_chars"] = float(_g(ind, "firstLineChars")) / 100
                if _g(ind, "hangingChars"):
                    attrs["hanging_chars"] = float(_g(ind, "hangingChars")) / 100
    return attrs


def _sz_label(pt):
    for k, label in _PT_TO_LABEL.items():
        if abs(pt - k) < 0.26:
            return label
    return f"{pt:g}pt"


def _spacing_label(rule, line):
    if line is None:
        return ("单倍行距", "1倍")
    if rule in ("exact", "atLeast"):
        key = "固定值" if rule == "exact" else "最小值"
        return (key, f"{line / 20:g}磅")
    mult = line / 240
    if abs(mult - 1.0) < 0.01:
        return ("单倍行距", "1倍")
    if abs(mult - 1.5) < 0.01:
        return ("1.5倍行距", "1.5倍")
    if abs(mult - 2.0) < 0.01:
        return ("2倍行距", "2倍")
    return ("多倍行距", f"{mult:g}倍")


def _to_section_attrs(a: dict) -> dict:
    """把解析出的有效属性转成 wordformat 配置词表（只含提取到的字段）。"""
    out = {"paragraph": {}, "font": {}}
    if a.get("jc") and a["jc"] in _JC_TO_LABEL:
        out["paragraph"]["alignment"] = _JC_TO_LABEL[a["jc"]]
    if a.get("line") is not None or a.get("line_rule"):
        rule, val = _spacing_label(a.get("line_rule"), a.get("line"))
        out["paragraph"]["line_spacingrule"] = rule
        out["paragraph"]["line_spacing"] = val
    if a.get("before_lines") is not None:
        out["paragraph"]["space_before"] = f"{a['before_lines']:g}行"
    if a.get("after_lines") is not None:
        out["paragraph"]["space_after"] = f"{a['after_lines']:g}行"
    if a.get("first_line_chars") is not None:
        out["paragraph"]["first_line_indent"] = f"{a['first_line_chars']:g}字符"
    if a.get("hanging_chars") is not None:
        out["paragraph"]["first_line_indent"] = f"-{a['hanging_chars']:g}字符"
    if a.get("east_asia"):
        out["font"]["chinese_font_name"] = a["east_asia"]
    if a.get("ascii"):
        out["font"]["english_font_name"] = a["ascii"]
    if a.get("sz_pt") is not None:
        out["font"]["font_size"] = _sz_label(a["sz_pt"])
    if a.get("bold") is not None:
        out["font"]["bold"] = a["bold"]
    return out


def extract_profile(docx_path: str, base_path: str | None = None):
    """从参考文档反推配置字典。

    :returns: (config_dict, report_lines) —— config 可直接 yaml.dump 为方案；
              report 记录命中/未命中/与基础预设的差异。
    """
    import copy
    import yaml

    base = {}
    if base_path:
        with open(base_path, encoding="utf-8") as f:
            base = yaml.safe_load(f) or {}
    if not base:
        raise ValueError("需要基础预设：-b 基础预设.yaml（用于回落未提取字段）")

    with zipfile.ZipFile(docx_path) as z:
        root = ET.fromstring(z.read("word/styles.xml").decode("utf-8"))
    styles_by_id = {e.get(W + "styleId"): e
                    for e in root.findall(W + "style") if e.get(W + "type") == "paragraph"}

    config = copy.deepcopy(base)
    report = ["由参考文档反推的格式方案", f"参考文档: {docx_path}", ""]
    matched_sections = set()
    for sid, el in styles_by_id.items():
        nm_el = el.find(W + "name")
        name = _g(nm_el, "val") if nm_el is not None else ""
        sections = classify(name, sid)
        if not sections:
            continue
        attrs = _style_attrs(el, styles_by_id)
        section_attrs = _to_section_attrs(attrs)
        report.append(f"命中 [{sid}] {name} → {' / '.join(sections)}")
        for section in sections:
            matched_sections.add(section)
            node = config
            for part in section.split("."):
                node = node.setdefault(part, {})
            # paragraph / font 子节合并提取字段（builtin_style_name 等保留基础值）
            for sub in ("paragraph", "font"):
                if sub in section_attrs:
                    node.setdefault(sub, {}).update(section_attrs[sub])

    missing = sorted({s for s, _, _ in _SECTIONS} - matched_sections)
    if missing:
        report.append("")
        report.append("未提取（回落基础预设）: " + ", ".join(missing))
    report.append("")
    report.append("与基础预设的字段差异:")
    for section in sorted(matched_sections):
        node = config
        for part in section.split("."):
            node = node[part]
        base_node = base
        try:
            for part in section.split("."):
                base_node = base_node[part]
        except (KeyError, TypeError):
            base_node = {}
        for sub in ("paragraph", "font"):
            for field, val in node.get(sub, {}).items():
                old = base_node.get(sub, {}).get(field) if isinstance(base_node, dict) else None
                if old != val:
                    report.append(f"  {section}.{sub}.{field}: {old!r} → {val!r}")
    return config, report


def save_yaml(config: dict, out_path: str):
    import yaml
    with open(out_path, "w", encoding="utf-8") as f:
        yaml.safe_dump(config, f, allow_unicode=True, sort_keys=False,
                       default_flow_style=False, width=100)
