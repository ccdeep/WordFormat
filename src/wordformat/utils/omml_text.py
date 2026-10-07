"""OMML → 线性数学文本：供识别预览面板显示公式内容用。

预览场景只需要"可读"，不追求排版还原：m:f 渲染成 a/b，上下标用 _ ^，
未识别的结构回退为递归拼接 m:t 文本，保证任何公式都有输出。
"""

from lxml import etree

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# 这些字符出现在子/母式里就加括号，避免 a+b/c、e^-t 这类歧义
_ATOMIC_BREAK = set(" \t+-*/=;,()[]")


def _q(tag: str) -> str:
    return f"{{{M_NS}}}{tag}"


def _localname(elem) -> str:
    return etree.QName(elem).localname if isinstance(elem.tag, str) else ""


def _text_of_runs(elem) -> str:
    """拼接元素下所有 m:t 的文字（数学 run 直接相连，不加空格）。"""
    return "".join(t.text or "" for t in elem.iter(_q("t")))


def _needs_parens(s: str) -> bool:
    return s == "" or any(c in _ATOMIC_BREAK for c in s)


def _wrap(s: str) -> str:
    return f"({s})" if _needs_parens(s) else s


def _delimiter_chars(elem) -> tuple[str, str]:
    """从 m:dPr 读 begChr/endChr，默认圆括号。"""
    beg, end = "(", ")"
    dpr = elem.find(_q("dPr"))
    if dpr is not None:
        b = dpr.find(_q("begChr"))
        e = dpr.find(_q("endChr"))
        if b is not None:
            beg = b.get(_q("val"), "(")
        if e is not None:
            end = e.get(_q("val"), ")")
    return beg, end


def _render(elem) -> str:
    name = _localname(elem)
    if name in ("ctrlPr", "rPr", "ctrl", "sty", "scr", "naryPr", "dPr", "fPr",
                "sSubPr", "sSupPr", "sSubSupPr", "radPr", "funcPr", "eqArrPr",
                "limLowPr", "limHighPr", "mPr", "accPr", "barPr", "argPr",
                "maxDist", "objDist", "baseJc", "begChr", "endChr", "chr",
                "pos", "vertJc", "degHide", "type", "lim", "sep", "count"):
        return ""
    if name == "t":
        return elem.text or ""
    if name in ("oMathPara", "oMath", "e", "num", "den", "sup", "sub", "fName",
                "deg", "mr", "limLow", "limHigh"):
        return "".join(_render(child) for child in elem)

    if name == "f":
        num = "".join(_render(n) for n in elem.findall(_q("num")))
        den = "".join(_render(n) for n in elem.findall(_q("den")))
        if _needs_parens(num) or _needs_parens(den):
            return f"({num})/({den})"
        return f"{num}/{den}"
    if name == "sSub":
        base = "".join(_render(n) for n in elem.findall(_q("e")))
        sub = "".join(_render(n) for n in elem.findall(_q("sub")))
        return f"{_wrap(base)}_{_wrap(sub)}"
    if name == "sSup":
        base = "".join(_render(n) for n in elem.findall(_q("e")))
        sup = "".join(_render(n) for n in elem.findall(_q("sup")))
        return f"{_wrap(base)}^{_wrap(sup)}"
    if name == "sSubSup":
        base = "".join(_render(n) for n in elem.findall(_q("e")))
        sub = "".join(_render(n) for n in elem.findall(_q("sub")))
        sup = "".join(_render(n) for n in elem.findall(_q("sup")))
        return f"{_wrap(base)}_{_wrap(sub)}^{_wrap(sup)}"
    if name == "rad":
        deg = "".join(_render(n) for n in elem.findall(_q("deg")))
        base = "".join(_render(n) for n in elem.findall(_q("e")))
        idx = deg.strip() if deg.strip() not in ("", "2") else ""
        return f"{idx}√({base})"
    if name == "d":
        beg, end = _delimiter_chars(elem)
        parts = ["".join(_render(n) for n in e) for e in elem.findall(_q("e"))]
        return beg + ", ".join(p for p in parts if p) + end
    if name == "func":
        fname = "".join(_render(n) for n in elem.findall(_q("fName")))
        arg = "".join(_render(n) for n in elem.findall(_q("e")))
        return fname + arg
    if name == "nary":
        npr = elem.find(_q("naryPr"))
        op = "∫"
        if npr is not None:
            ch = npr.find(_q("chr"))
            if ch is not None:
                op = ch.get(_q("val"), "∫")
        sub = "".join(_render(n) for n in elem.findall(_q("sub")))
        sup = "".join(_render(n) for n in elem.findall(_q("sup")))
        base = "".join(_render(n) for n in elem.findall(_q("e")))
        head = op
        if sub:
            head += f"_{_wrap(sub)}"
        if sup:
            head += f"^{_wrap(sup)}"
        return f"{head}{_wrap(base)}"
    if name == "limLow":
        base = "".join(_render(n) for n in elem.findall(_q("e")))
        lim = "".join(_render(n) for n in elem.findall(_q("lim")))
        return f"{base}_{_wrap(lim)}"
    if name == "limHigh":
        base = "".join(_render(n) for n in elem.findall(_q("e")))
        lim = "".join(_render(n) for n in elem.findall(_q("lim")))
        return f"{base}^{_wrap(lim)}"
    if name in ("m", "eqArr"):
        # 矩阵：m:mr 行内 m:e 用 ", " 分隔；方程组：m:e 每个是一条方程
        if name == "m":
            rows = [", ".join(filter(None, (_render(e) for e in mr.findall(_q("e")))))
                    for mr in elem.findall(_q("mr"))]
            return " ; ".join(r for r in rows if r)
        lines = [_render(e) for e in elem.findall(_q("e"))]
        return " ; ".join(ln for ln in lines if ln)

    # 未知结构：递归子节点；叶子则收集其中的 m:t 文本
    children = list(elem)
    if children:
        return "".join(_render(child) for child in children)
    return _text_of_runs(elem)


def omml_to_text(elem) -> str:
    """把 m:oMath / m:oMathPara 元素转成一行线性数学文本。"""
    text = _render(elem).strip()
    # "#(2-1)" 是 OMML 的公式编号标记：去掉 #，编号 (2-1) 保留在行尾
    text = text.replace("#", " ")
    return " ".join(text.split())


def paragraph_math_texts(p_elem) -> list[str]:
    """按文档顺序取出段落里每个公式的线性文本。

    显示型公式是 w:p > m:oMathPara > m:oMath，行内公式是 w:p > m:oMath，
    深度遍历保证两类都收集且不重复。
    """
    texts = []

    def walk(el):
        for child in el:
            tag = child.tag
            if tag == _q("oMathPara"):
                for om in child.findall(_q("oMath")):
                    texts.append(omml_to_text(om))
            elif tag == _q("oMath"):
                texts.append(omml_to_text(child))
            else:
                walk(child)

    walk(p_elem)
    return texts
