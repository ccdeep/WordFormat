"""docx → 高保真 HTML，供右侧"文档视图"（WebView2）显示与即时跳转。

- run 级字体（中西文分槽）、字号、加粗、斜体、颜色，段落对齐/缩进/行距直接
  取自文档 XML，段落样式（pStyle）沿 basedOn 链回溯解析，排版接近 Word；
- 公式 OMML → MathML，WebView2 的 Chromium 内核原生渲染；
- 表格渲染为带边框 HTML 表格（处理 gridSpan/vMerge），图片转 base64 内嵌；
- 每个识别段落一个锚点 id="p-<序号>"，序号与识别结果（iter_document_paragraphs
  顺序）严格一致，左侧列表/目录导航据此毫秒级跳转。
"""
import re
from html import escape

from docx.oxml.ns import qn

from wordformat.utils import get_paragraph_numbering_text

_W = lambda tag: qn("w:" + tag)
_M = lambda tag: qn("m:" + tag)

_TRUE = {"1", "true", "on"}
_OP_CHARS = set("+-*/=<>|()[]{}∑∫∏√±×÷≡≈≤≥≠∂∇∞∴∶,:;.。、，；：")


def _local(el):
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


# ── 样式解析 ─────────────────────────────────────────────
class _StyleResolver:
    """pStyle → basedOn 链回溯，拿到段落/字符样式的 rPr、pPr。"""

    def __init__(self, document):
        self.styles = {}
        try:
            for st in document.styles.element.findall(_W("style")):
                based = st.find(_W("basedOn"))
                self.styles[st.get(_W("styleId"))] = {
                    "based": based.get(_W("val")) if based is not None else None,
                    "rPr": st.find(_W("rPr")),
                    "pPr": st.find(_W("pPr")),
                }
        except Exception:
            self.styles = {}

    def chain(self, p_elem):
        """返回 (样式rPr列表[从基类到派生], 样式pPr列表)，供首者优/近者优查找。"""
        ppr = p_elem.find(_W("pPr"))
        ps = ppr.find(_W("pStyle")) if ppr is not None else None
        sid = ps.get(_W("val")) if ps is not None else None
        rprs, pprs, seen = [], [], set()
        while sid and sid not in seen and len(seen) < 8:
            seen.add(sid)
            st = self.styles.get(sid)
            if not st:
                break
            if st["rPr"] is not None:
                rprs.append(st["rPr"])
            if st["pPr"] is not None:
                pprs.append(st["pPr"])
            sid = st["based"]
        return rprs, pprs


def _find_prop(elems, tag, attr=None):
    """在 rPr/pPr 元素列表里找属性：近处（派生样式/直接格式）优先。"""
    for el in reversed(elems):
        if el is None:
            continue
        sub = el.find(_W(tag))
        if sub is not None:
            if attr is None:
                return sub
            return sub.get(_W(attr))
    return None


def _run_css(style_rprs, rpr):
    """run 级 CSS：字体（中西文分槽）/字号/加粗/斜体/下划线/颜色。"""
    elems = list(style_rprs) + ([rpr] if rpr is not None else [])
    css = []
    fonts = _find_prop(elems, "rFonts")
    if fonts is not None:
        ea = fonts.get(_W("eastAsia"))
        ascii_ = fonts.get(_W("ascii"))
        if ea or ascii_:
            # CSS 字体名必须用单引号：外层 HTML 属性是双引号
            css.append(f"font-family:'{ea or ascii_}','{ascii_ or ea}'")
    sz = _find_prop(elems, "sz", "val")
    if sz and sz.isdigit():
        css.append(f"font-size:{int(sz) / 2}pt")
    if (_find_prop(elems, "b")) is not None:
        v = _find_prop(elems, "b").get(_W("val"))
        css.append("font-weight:400" if v in _TRUE else "font-weight:700")
    if (_find_prop(elems, "i")) is not None:
        v = _find_prop(elems, "i").get(_W("val"))
        css.append("font-style:normal" if v in _TRUE else "font-style:italic")
    if (_find_prop(elems, "u")) is not None:
        css.append("text-decoration:underline")
    color = _find_prop(elems, "color", "val")
    if color and color != "auto":
        css.append(f"color:#{color}")
    return ";".join(css)


def _para_css(style_pprs, ppr, paragraph):
    """段落级 CSS：对齐/首行缩进/行距；返回 (css, 编号前缀文字)。"""
    elems = list(style_pprs) + ([ppr] if ppr is not None else [])
    css = []
    jc = _find_prop(elems, "jc", "val")
    if jc:
        css.append({"both": "text-align:justify", "center": "text-align:center",
                    "right": "text-align:right", "left": "text-align:left"}.get(jc, ""))
    ind = _find_prop(elems, "ind")
    if ind is not None:
        flc = ind.get(_W("firstLineChars"))
        fl = ind.get(_W("firstLine"))
        if flc and flc.isdigit():
            css.append(f"text-indent:{int(flc) / 100}em")
        elif fl and fl.isdigit():
            css.append(f"text-indent:{int(fl) / 20}pt")
    spacing = _find_prop(elems, "spacing")
    if spacing is not None:
        line = spacing.get(_W("line"))
        rule = spacing.get(_W("lineRule"))
        if line and line.isdigit() and int(line) >= 240 and rule in (None, "auto"):
            css.append(f"line-height:{int(line) / 240:.2f}")
    prefix = get_paragraph_numbering_text(paragraph) if paragraph is not None else ""
    return ";".join(c for c in css if c), prefix


# ── 公式 OMML → MathML ───────────────────────────────────
_SKIP = {"ctrlPr", "rPr", "sty", "scr", "naryPr", "dPr", "fPr", "sSubPr", "sSupPr",
         "sSubSupPr", "radPr", "funcPr", "eqArrPr", "limLowPr", "limHighPr", "mPr",
         "accPr", "barPr", "argPr", "maxDist", "objDist", "baseJc", "begChr",
         "endChr", "chr", "pos", "vertJc", "degHide", "type", "lim", "sep", "count"}


def _mtext_to_mathml(text):
    """数学 run 文本 → mi/mn/mo 的粗分类（显示够用）。"""
    out = []
    for tok in re.findall(r"[0-9]+\.?[0-9]*|[a-zA-Zα-ωΑ-Ω]+|\S", text or ""):
        if re.fullmatch(r"[0-9]+\.?[0-9]*", tok):
            out.append(f"<mn>{escape(tok)}</mn>")
        elif len(tok) == 1 and (tok in _OP_CHARS or tok in "=+-"):
            out.append(f"<mo>{escape(tok)}</mo>")
        elif all(c in _OP_CHARS for c in tok):
            out.append(f"<mo>{escape(tok)}</mo>")
        else:
            out.append(f"<mi>{escape(tok)}</mi>")
    return "<mrow>" + "".join(out) + "</mrow>" if out else ""


def _m_children(el):
    return "".join(_mathml_of(c) for c in el)


def _mathml_of(el):
    name = _local(el)
    if name in _SKIP or name == "t":
        return ""
    if name in ("oMathPara", "oMath", "e", "num", "den", "sup", "sub", "fName",
                "deg", "mr"):
        return _m_children(el)
    if name == "r":
        return _mtext_to_mathml("".join(t.text or "" for t in el.findall(_M("t"))))
    if name == "f":
        num = "".join(_mathml_of(c) for c in el.find(_M("num")))
        den = "".join(_mathml_of(c) for c in el.find(_M("den")))
        return f"<mfrac>{num}{den}</mfrac>"
    if name == "sSub":
        return (f"<msub>{_m_children(_e_of(el, 'e'))}"
                f"{_m_children(_e_of(el, 'sub'))}</msub>")
    if name == "sSup":
        return (f"<msup>{_m_children(_e_of(el, 'e'))}"
                f"{_m_children(_e_of(el, 'sup'))}</msup>")
    if name == "sSubSup":
        return (f"<msubsup>{_m_children(_e_of(el, 'e'))}"
                f"{_m_children(_e_of(el, 'sub'))}"
                f"{_m_children(_e_of(el, 'sup'))}</msubsup>")
    if name == "rad":
        deg = el.find(_M("deg"))
        base = _m_children(_e_of(el, "e"))
        hide = el.find(_M("degHide"))
        if deg is None or hide is not None or not "".join(deg.itertext()).strip():
            return f"<msqrt>{base}</msqrt>"
        return f"<mroot>{base}{_m_children(deg)}</mroot>"
    if name == "d":
        dpr = el.find(_M("dPr"))
        beg, end = "(", ")"
        if dpr is not None:
            b = dpr.find(_M("begChr"))
            e = dpr.find(_M("endChr"))
            beg = b.get(_M("val"), "(") if b is not None else "("
            end = e.get(_M("val"), ")") if e is not None else ")"
        inner = "".join(_mathml_of(c) for c in el.findall(_M("e")))
        return (f'<mrow><mo fence="true">{escape(beg)}</mo>{inner}'
                f'<mo fence="true">{escape(end)}</mo></mrow>')
    if name == "func":
        fname = _m_children(_e_of(el, "fName"))
        return fname + _m_children(_e_of(el, "e"))
    if name == "nary":
        npr = el.find(_M("naryPr"))
        op = "∫"
        if npr is not None and npr.find(_M("chr")) is not None:
            op = npr.find(_M("chr")).get(_M("val"), "∫")
        sub = _m_children(_e_of(el, "sub"))
        sup = _m_children(_e_of(el, "sup"))
        head = f"<mo>{escape(op)}</mo>"
        if sub or sup:
            head = f"<munderover>{head}{sub or '<mrow/>'}{sup or '<mrow/>'}</munderover>"
        elif sub:
            head = f"<munder>{head}{sub}</munder>"
        return f"<mrow>{head}{_m_children(_e_of(el, 'e'))}</mrow>"
    if name == "limLow":
        return (f"<munder>{_m_children(_e_of(el, 'e'))}"
                f"{_m_children(_e_of(el, 'lim'))}</munder>")
    if name == "limHigh":
        return (f"<mover>{_m_children(_e_of(el, 'e'))}"
                f"{_m_children(_e_of(el, 'lim'))}</mover>")
    if name == "eqArr":
        rows = "".join(f"<mtr><mtd>{_mathml_of(c)}</mtd></mtr>"
                       for c in el.findall(_M("e")))
        return f"<mtable>{rows}</mtable>"
    if name == "m":
        rows = "".join(
            "<mtr>" + "".join(f"<mtd>{_mathml_of(c)}</mtd>"
                              for c in mr.findall(_M("e"))) + "</mtr>"
            for mr in el.findall(_M("mr")))
        return f"<mtable>{rows}</mtable>"
    children = list(el)
    if children:
        return "".join(_mathml_of(c) for c in children)
    return ""


def _e_of(el, tag):
    """取 OMML 子元素（可能不存在），保证 _m_children 不崩。"""
    sub = el.find(_M(tag))
    return sub if sub is not None else el.makeelement(_M("e"), {})


# ── 段落/表格渲染 ────────────────────────────────────────
def _img_html(r_elem, part, img_dir, counter):
    """把 run 里的图片落盘（Chromium 不认的格式用 Pillow 转 PNG），返回 <img>。"""
    imgs = []
    for blip in r_elem.findall(".//" + qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if not rid:
            continue
        try:
            p = part.related_parts[rid]
            blob = p.blob
            ct = (getattr(p, "content_type", None) or "image/png").lower()
            n = next(counter)
            if ct in ("image/png", "image/jpeg", "image/gif"):
                ext = {"image/png": "png", "image/jpeg": "jpg", "image/gif": "gif"}[ct]
                name = f"img_{n}.{ext}"
                (img_dir / name).write_bytes(blob)
            else:
                # EMF/WMF/BMP 等 Chromium 不支持的格式 → Pillow 转 PNG
                from PIL import Image
                im = Image.open(__import__("io").BytesIO(blob))
                try:
                    im.load(dpi=192)
                except TypeError:
                    im.load()
                if im.width > 1200:
                    im = im.resize((1200, max(1, round(im.height * 1200 / im.width))))
                if im.mode not in ("RGB", "RGBA"):
                    im = im.convert("RGB")
                name = f"img_{n}.png"
                im.save(img_dir / name)
            imgs.append(f'<img src="{name}" style="max-width:100%">')
        except Exception:
            imgs.append('<span style="color:#999">〔图片〕</span>')
    return "".join(imgs)


def _run_html(r_elem, part, style_rprs, img_dir, counter):
    rpr = r_elem.find(_W("rPr"))
    css = _run_css(style_rprs, rpr)
    st = f' style="{css}"' if css else ""
    out = []
    if r_elem.find(".//" + qn("a:blip")) is not None:
        out.append(_img_html(r_elem, part, img_dir, counter))
    for child in r_elem:
        name = _local(child)
        if name == "t":
            txt = child.text or ""
            if txt:
                out.append(escape(txt))
        elif name == "tab":
            out.append("&ensp;&ensp;")
        elif name == "br":
            out.append("<br>")
    if not out:
        return ""
    wrap = f"<span{st}>" + "".join(out) + "</span>"
    return wrap


def _oMath_html(om, part):
    return f'<math display="block">{_mathml_of(om)}</math>'


def _para_html(p_elem, part, resolver, idx, paragraph, img_dir, counter):
    """单个段落 → <p id="p-<idx>">。idx 为 None 时不加锚点（如文本框内容）。"""
    ppr = p_elem.find(_W("pPr"))
    style_rprs, style_pprs = resolver.chain(p_elem)
    css, prefix = _para_css(style_pprs, ppr, paragraph)
    merged_rprs = style_rprs
    if ppr is not None:
        own_rpr = ppr.find(_W("rPr"))
        if own_rpr is not None:
            merged_rprs = style_rprs + [own_rpr]

    pieces = []
    for child in p_elem:
        name = _local(child)
        if name == "r":
            pieces.append(_run_html(child, part, merged_rprs, img_dir, counter))
        elif name == "hyperlink":
            for r in child.findall(_W("r")):
                pieces.append(_run_html(r, part, merged_rprs, img_dir, counter))
        elif name == "oMathPara":
            for om in child.findall(_M("oMath")):
                pieces.append(_oMath_html(om, part))
        elif name == "oMath":
            pieces.append(_oMath_html(child, part))
    if prefix:
        pieces.insert(0, escape(prefix))
    content = "".join(pieces) or "<br>"
    anchor = f' id="p-{idx}"' if idx is not None else ""
    st = f' style="{css}"' if css else ""
    return f"<p{anchor}{st}>{content}</p>"


def _tbl_html(tbl, part, resolver, para_map, img_dir, counter):
    rows = []
    for tr in tbl.findall(_W("tr")):
        cells = []
        for tc in tr.findall(_W("tc")):
            tcpr = tc.find(_W("tcPr"))
            colspan = 1
            vmerge = None
            if tcpr is not None:
                gs = tcpr.find(_W("gridSpan"))
                if gs is not None and gs.get(_W("val"), "1").isdigit():
                    colspan = int(gs.get(_W("val")))
                vm = tcpr.find(_W("vMerge"))
                if vm is not None:
                    vmerge = vm.get(_W("val"), "continue")
            if vmerge == "continue":
                continue  # 被合并的延续单元格不渲染
            inner = []
            for p in tc.findall(_W("p")):
                info = para_map.get(p)
                if info is not None:
                    idx, paragraph = info[0], info[1]
                    inner.append(_para_html(p, part, resolver, idx, paragraph,
                                            img_dir, counter))
                else:
                    inner.append(_para_html(p, part, resolver, None, None,
                                            img_dir, counter))
            rowspan = ""
            if vmerge == "restart":
                rowspan = ' rowspan="1"'  # 纵向合并的精确行数 v1 不展开，视觉可接受
            cells.append(f'<td colspan="{colspan}"{rowspan}>{"".join(inner)}</td>')
        if cells:
            rows.append("<tr>" + "".join(cells) + "</tr>")
    return f'<table style="border-collapse:collapse">{"".join(rows)}</table>'


def build_html(document, para_map, img_dir) -> str:
    """生成整篇高保真 HTML。图片落盘到 img_dir（与 HTML 同目录，相对引用）。

    para_map: {段落XML元素: (识别序号, Paragraph, in_table)}，
    由调用方按 iter_document_paragraphs 的顺序构建，保证锚点序号与识别结果一致。
    """
    import itertools
    from pathlib import Path

    img_dir = Path(img_dir)
    img_dir.mkdir(parents=True, exist_ok=True)
    for old in img_dir.glob("img_*"):
        old.unlink()
    counter = itertools.count()
    resolver = _StyleResolver(document)
    part = document.part
    out = []

    def walk(container):
        for child in container.iterchildren():
            name = _local(child)
            if name == "p":
                info = para_map.get(child)
                if info is not None:
                    idx, paragraph = info[0], info[1]
                    out.append(_para_html(child, part, resolver, idx, paragraph,
                                          img_dir, counter))
                else:
                    out.append(_para_html(child, part, resolver, None, None,
                                          img_dir, counter))
            elif name == "tbl":
                out.append(_tbl_html(child, part, resolver, para_map,
                                     img_dir, counter))
            elif name == "sdt":
                content = child.find(_W("sdtContent"))
                if content is not None:
                    walk(content)

    walk(document.element.body)

    css = ('body{font-family:"宋体",SimSun,serif;font-size:12pt;color:#111;'
           "margin:20px 28px;background:#fff;line-height:1.5;}"
           "p{margin:3px 0;}"
           'table{border-collapse:collapse;margin:6px 0;}'
           'td{border:1px solid #777;padding:4px 8px;vertical-align:top;}'
           "math{display:block;text-align:center;margin:4px 0;}"
           "span{white-space:pre-wrap;}"
           ".wf-hl{background:#FFE69C;}")
    script = ('function wfJump(i){var el=document.getElementById("p-"+i);'
              "if(!el)return false;"
              'var olds=document.querySelectorAll(".wf-hl");'
              "for(var k=0;k<olds.length;k++){olds[k].classList.remove('wf-hl');}"
              "el.classList.add('wf-hl');"
              'el.scrollIntoView({block:"center"});return true;}')
    return ('<html><head><meta charset="utf-8">'
            f"<style>{css}</style></head><body>"
            + "".join(out) + f"<script>{script}</script></body></html>")
