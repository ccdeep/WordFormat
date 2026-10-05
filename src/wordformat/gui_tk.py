#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""Tkinter 识别预览/改判界面（需求文档 D6 形态，双标签页）：

Tab1 识别与套用：
    Word 文件 → [开始识别] → 结构列表预览（类别下拉改判，右侧原文联动）→ [套用] → 新 docx
Tab2 格式方案：
    载入预设/从参考文档提取 → 按类别可视化调参（字体/字号/对齐/行距/缩进）→ 另存
全部本地离线运行。
"""
import copy
import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

import yaml

from wordformat.base import DocxBase
from wordformat.pipeline.orchestrate import auto_format_thesis_document
from wordformat.settings import VOIDNODELIST

# ── 类别全集（模型标签 + 结构规则标签）─────────────────────
MODEL_LABELS = [
    "body_text", "heading_level_1", "heading_level_2", "heading_level_3",
    "heading_mulu", "heading_fulu", "caption_figure", "caption_table",
    "document_title", "other", "references_title", "references_content",
    "acknowledgements_title", "acknowledgements_content",
    "keywords_chinese", "keywords_english",
    "abstract_chinese_title", "abstract_chinese_content",
    "abstract_chinese_title_content", "abstract_english_title",
    "abstract_english_title_content", "abstract_english_content",
]
RULE_LABELS = ["equation_para", "table_text", "toc_line", "footer", "figure_image"]
CATEGORIES = sorted(set(MODEL_LABELS + RULE_LABELS))

CAT_CN = {
    "body_text": "正文", "heading_level_1": "一级标题", "heading_level_2": "二级标题",
    "heading_level_3": "三级标题", "heading_mulu": "目录标题", "heading_fulu": "附录标题",
    "caption_figure": "图注", "caption_table": "表注", "document_title": "文档标题",
    "other": "其他", "references_title": "参考文献标题", "references_content": "参考文献条目",
    "acknowledgements_title": "致谢标题", "acknowledgements_content": "致谢正文",
    "keywords_chinese": "中文关键词", "keywords_english": "英文关键词",
    "abstract_chinese_title": "中文摘要标题", "abstract_chinese_content": "中文摘要正文",
    "abstract_chinese_title_content": "中文摘要标题正文",
    "abstract_english_title": "英文摘要标题", "abstract_english_content": "英文摘要正文",
    "abstract_english_title_content": "英文摘要标题正文",
    "equation_para": "公式段落", "table_text": "表格内文字", "toc_line": "目录行",
    "footer": "页脚", "figure_image": "图片段落",
}
CN_TO_CATEGORY = {v: k for k, v in CAT_CN.items()}

# 下拉框按标准论文内容顺序排列（找不到的类别追加在末尾）
_CATEGORY_ORDER = [
    "document_title",
    "abstract_chinese_title", "abstract_chinese_content", "keywords_chinese",
    "abstract_english_title", "abstract_english_content", "keywords_english",
    "heading_level_1", "heading_level_2", "heading_level_3",
    "body_text", "equation_para",
    "figure_image", "caption_figure", "caption_table", "table_text",
    "heading_mulu", "toc_line", "heading_fulu",
    "references_title", "references_content",
    "acknowledgements_title", "acknowledgements_content",
    "footer", "other",
]
_CATEGORY_ORDER_CN = [CAT_CN.get(c, c) for c in _CATEGORY_ORDER if c in CAT_CN]
_CATEGORY_ORDER_CN += [CAT_CN[c] for c in CATEGORIES if c not in _CATEGORY_ORDER]
CN_LIST_ORDERED = _CATEGORY_ORDER_CN

# ── 可配置节（Tab2 编辑面板）────────────────────────────────
_SECTIONS = [
    ("body.text", "正文"),
    ("headings.level_1", "一级标题"),
    ("headings.level_2", "二级标题"),
    ("headings.level_3", "三级标题"),
    ("abstract.chinese.title", "中文摘要标题"),
    ("abstract.chinese.body", "中文摘要正文"),
    ("abstract.chinese.keywords", "中文关键词"),
    ("abstract.english.title", "英文摘要标题"),
    ("abstract.english.body", "英文摘要正文"),
    ("abstract.english.keywords", "英文关键词"),
    ("references.title", "参考文献标题"),
    ("references.entry", "参考文献条目"),
    ("acknowledgements.title", "致谢标题"),
    ("figures.caption", "图注"),
    ("tables.caption", "表注"),
    ("math.block", "公式段落"),
]
_FONT_FIELDS = ["chinese_font_name", "english_font_name", "font_size", "bold"]
_PARA_FIELDS = ["alignment", "line_spacingrule", "line_spacing",
                "space_before", "space_after", "first_line_indent"]
_FONT_FIELD_CN = {"chinese_font_name": "中文字体", "english_font_name": "西文字体",
                  "font_size": "字号", "bold": "加粗"}
_PARA_FIELD_CN = {"alignment": "对齐", "line_spacingrule": "行距规则",
                  "line_spacing": "行距值", "space_before": "段前",
                  "space_after": "段后", "first_line_indent": "首行缩进"}
_SIZE_CHOICES = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
                 "四号", "小四", "五号", "小五", "六号", "七号"]
_ALIGN_CHOICES = ["两端对齐", "居中对齐", "左对齐", "右对齐"]
_RULE_CHOICES = ["单倍行距", "1.5倍行距", "2倍行距", "多倍行距", "固定值", "最小值"]


def _cat_cn(cat: str) -> str:
    return CAT_CN.get(cat, cat)


def _builtin_presets() -> dict[str, Path]:
    """内置预设：{方案名: 路径}。

    预设 YAML 打包进 wordformat/data/presets/（--collect-all 会带进冻结环境），
    开发环境与冻结环境都通过 gui_tk.py 同级的 data/presets 解析。
    """
    presets: dict[str, Path] = {}
    pkg_dir = Path(__file__).resolve().parent / "data" / "presets"
    if pkg_dir.is_dir():
        for f in sorted(pkg_dir.glob("*.yaml")):
            presets[f.stem] = f
    return presets


def _default_preset() -> str:
    """默认预设路径（北理工毕设报告）；找不到返回空串。"""
    presets = _builtin_presets()
    if "北理工毕设报告" in presets:
        return str(presets["北理工毕设报告"])
    return str(next(iter(presets.values()))) if presets else ""


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Word 格式整理器（wordformat）")
        root.geometry("1280x780")
        root.minsize(1024, 620)

        self.data: list[dict] = []
        self.docx_path: str | None = None
        self._busy = False
        self._cat_editor = None  # 单元格上的改判下拉

        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True)
        self.tab1 = ttk.Frame(nb, padding=(8, 6))
        self.tab2 = ttk.Frame(nb, padding=(8, 6))
        nb.add(self.tab1, text=" 识别与套用 ")
        nb.add(self.tab2, text=" 格式方案 ")
        self._build_tab1(self.tab1)
        self._build_tab2(self.tab2)
        self.status_var = tk.StringVar(value="就绪——选择 Word 文件后点击「开始识别」")
        ttk.Label(root, textvariable=self.status_var, relief="sunken",
                  anchor="w", padding=(6, 3)).pack(fill="x", side="bottom")
        self.yaml_var.set(_default_preset())

    # ══ Tab1：识别与套用 ═════════════════════════════════
    def _build_tab1(self, tab):
        top = ttk.Frame(tab)
        top.pack(fill="x")
        ttk.Label(top, text="Word 文件:").grid(row=0, column=0, sticky="w")
        self.docx_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.docx_var, width=54).grid(row=0, column=1, padx=4)
        ttk.Button(top, text="浏览…", command=self.browse_docx).grid(row=0, column=2)
        ttk.Label(top, text="格式方案:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.yaml_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.yaml_var, width=54).grid(row=1, column=1, padx=4, pady=(4, 0))
        ttk.Button(top, text="浏览…", command=self.browse_yaml).grid(row=1, column=2, pady=(4, 0))
        ttk.Button(top, text="从参考文档提取…", command=self.extract_from_reference).grid(
            row=1, column=3, padx=(6, 0), pady=(4, 0))
        ttk.Button(top, text="去「格式方案」页编辑 ↗", command=lambda: self._select_tab(1)).grid(
            row=1, column=4, padx=(6, 0), pady=(4, 0))
        ttk.Label(top, text="内置方案:").grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.preset_cb = ttk.Combobox(top, width=54, state="readonly")
        self.preset_cb["values"] = list(_builtin_presets()) or ["（未找到内置方案）"]
        if _builtin_presets():
            self.preset_cb.set("北理工毕设报告" if "北理工毕设报告" in _builtin_presets()
                               else next(iter(_builtin_presets())))
        self.preset_cb.grid(row=2, column=1, columnspan=2, sticky="w", padx=4, pady=(4, 0))
        ttk.Label(top, text="选中即载入，也可在「格式方案」页编辑",
                  foreground="#888888").grid(row=2, column=3, columnspan=2, sticky="w", pady=(4, 0))
        self.preset_cb.bind("<<ComboboxSelected>>", self._on_preset_selected)

        bar = ttk.Frame(tab)
        bar.pack(fill="x", pady=6)
        self.detect_btn = ttk.Button(bar, text="🔍 开始识别", command=self.start_detect)
        self.detect_btn.pack(side="left")
        self.apply_btn = ttk.Button(bar, text="📄 按当前方案套用并另存", command=self.start_apply,
                                    state="disabled")
        self.apply_btn.pack(side="left", padx=8)
        self.comments_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="写入审计批注", variable=self.comments_var).pack(side="left", padx=4)
        self.save_btn = ttk.Button(bar, text="💾 保存结构 JSON", command=self.save_json,
                                   state="disabled")
        self.save_btn.pack(side="left", padx=8)

        pane = ttk.Panedwindow(tab, orient="horizontal")
        pane.pack(fill="both", expand=True)

        left = ttk.Frame(pane)
        pane.add(left, weight=3)
        columns = ("idx", "cat", "score", "review", "text")
        headers = ("序号", "类别（双击改判）", "置信度", "需复核", "文本摘录")
        widths = (50, 120, 56, 60, 320)
        self.tree = ttk.Treeview(left, columns=columns, show="headings", height=24)
        for col, head, width in zip(columns, headers, widths):
            self.tree.heading(col, text=head)
            self.tree.column(col, width=width, anchor="w")
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(fill="both", expand=True, side="left")
        vsb.pack(side="right", fill="y")
        self.tree.tag_configure("review", background="#FFF3CD")
        self.tree.tag_configure("void", foreground="#8A8A8A")
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        right = ttk.Frame(pane)
        pane.add(right, weight=2)
        ttk.Label(right, text="原文（点击左侧列表行可定位高亮）",
                  foreground="#666666").pack(anchor="w")
        self.原文 = tk.Text(right, wrap="word", state="disabled",
                            font=("Microsoft YaHei UI", 10), padx=8, pady=6,
                            cursor="arrow")
        zsb = ttk.Scrollbar(right, orient="vertical", command=self.原文.yview)
        self.原文.configure(yscrollcommand=zsb.set)
        self.原文.pack(fill="both", expand=True, side="left")
        zsb.pack(side="right", fill="y")
        self.原文.tag_configure("current", background="#FFE69C")
        self.原文.tag_configure("placeholder", foreground="#999999")

    # ══ Tab2：格式方案 ═══════════════════════════════════
    def _build_tab2(self, tab):
        top = ttk.Frame(tab)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="方案文件:").grid(row=0, column=0, sticky="w")
        self.cfg_path_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.cfg_path_var, width=64).grid(row=0, column=1, padx=4)
        ttk.Button(top, text="打开…", command=self.cfg_open).grid(row=0, column=2)
        ttk.Button(top, text="另存为…", command=self.cfg_save_as).grid(row=0, column=3, padx=4)
        ttk.Button(top, text="从参考文档提取…", command=self.extract_from_reference).grid(row=0, column=4)
        ttk.Button(top, text="保存修改", command=self.cfg_save).grid(row=0, column=5, padx=(8, 0))

        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="元素类别:").grid(row=0, column=0, sticky="nw")
        self.section_list = tk.Listbox(body, width=22, height=20, exportselection=False)
        for _, label in _SECTIONS:
            self.section_list.insert("end", label)
        self.section_list.grid(row=1, column=0, sticky="ns")
        self.section_list.bind("<<ListboxSelect>>", self._on_section_select)
        ttk.Label(body, text="（值示例：行距值 '1.5倍' 或 '22磅'；段前段后 '0.5行' 或 '12磅'；"
                             "缩进 '2字符'，悬挂用 '-2字符'）",
                  foreground="#888888").grid(row=2, column=0, sticky="w")

        panel = ttk.LabelFrame(body, text="参数", padding=10)
        panel.grid(row=1, column=1, sticky="nw", padx=(12, 0))
        self.fields = {}
        rows = [
            ("中文字体", "entry", "font", "chinese_font_name"),
            ("西文字体", "entry", "font", "english_font_name"),
            ("字号", ("combobox", _SIZE_CHOICES), "font", "font_size"),
            ("加粗", ("combobox", ["加粗", "不加粗"]), "font", "bold"),
            ("对齐", ("combobox", _ALIGN_CHOICES), "paragraph", "alignment"),
            ("行距规则", ("combobox", _RULE_CHOICES), "paragraph", "line_spacingrule"),
            ("行距值", "entry", "paragraph", "line_spacing"),
            ("段前", "entry", "paragraph", "space_before"),
            ("段后", "entry", "paragraph", "space_after"),
            ("首行缩进", "entry", "paragraph", "first_line_indent"),
        ]
        for i, (label, kind, sub, field) in enumerate(rows):
            ttk.Label(panel, text=label + ":").grid(row=i, column=0, sticky="w", pady=3)
            if kind == "entry":
                w = ttk.Entry(panel, width=24)
            else:
                w = ttk.Combobox(panel, width=22, values=kind[1], state="readonly")
            w.grid(row=i, column=1, sticky="w", padx=(6, 0))
            self.fields[(sub, field)] = w
        ttk.Label(panel, text="改完记得点上方「保存修改」。未填写的字段保持原值。",
                  foreground="#888888").grid(row=len(rows), column=0, columnspan=2,
                                             sticky="w", pady=(10, 0))
        self.section_list.selection_set(0)
        self._on_section_select()

    # ── Tab2：方案文件读写 ────────────────────────────────
    def _load_cfg_dict(self, path: str) -> dict:
        with open(path, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}

    def _cfg_node(self, section: str) -> dict:
        node = self.cfg_dict
        for part in section.split("."):
            node = node.setdefault(part, {})
        return node

    def _on_section_select(self, _event=None):
        if not getattr(self, "cfg_dict", None):
            self.cfg_dict = {}
        sel = self.section_list.curselection()
        if not sel:
            return
        section = _SECTIONS[sel[0]][0]
        node = self._cfg_node(section)
        font, para = node.get("font") or {}, node.get("paragraph") or {}
        def _set_widget(w, val):
            text = str(val) if val is not None else ""
            if isinstance(w, ttk.Combobox):
                w.set(text)
            else:
                w.delete(0, "end")
                w.insert(0, text)

        for (sub, field), w in self.fields.items():
            src = font if sub == "font" else para
            val = src.get(field)
            if field == "bold":
                _set_widget(w, "加粗" if val else "不加粗")
            else:
                _set_widget(w, val)

    def _collect_section(self, section: str) -> None:
        """把面板字段写回 cfg_dict 对应节（空值不覆盖）。"""
        node = self._cfg_node(section)
        font = node.setdefault("font", {})
        para = node.setdefault("paragraph", {})
        for (sub, field), w in self.fields.items():
            val = str(w.get()).strip()
            if not val:
                continue
            target = font if sub == "font" else para
            if field == "bold":
                target[field] = val == "加粗"
            else:
                target[field] = val

    def cfg_open(self, path: str | None = None):
        if not path:
            path = filedialog.askopenfilename(title="打开格式方案",
                                              filetypes=[("YAML", "*.yaml *.yml")])
            if not path:
                return
        try:
            self.cfg_dict = self._load_cfg_dict(path)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))
            return
        self.cfg_path_var.set(path)
        self.yaml_var.set(path)  # 两页方案路径始终同步
        self.section_list.selection_clear(0, "end")
        self.section_list.selection_set(0)
        self._on_section_select()
        self.status_var.set(f"方案已载入: {path}")

    def cfg_save_as(self):
        path = filedialog.asksaveasfilename(
            title="另存格式方案", defaultextension=".yaml",
            initialfile="我的方案.yaml", filetypes=[("YAML", "*.yaml *.yml")])
        if not path:
            return
        self.cfg_path_var.set(path)
        self.cfg_save()

    def cfg_save(self):
        path = self.cfg_path_var.get().strip()
        if not path:
            self.cfg_save_as()
            return
        if not getattr(self, "cfg_dict", None):
            messagebox.showwarning("提示", "请先打开一个方案文件（或从参考文档提取）")
            return
        sel = self.section_list.curselection()
        if sel:
            self._collect_section(_SECTIONS[sel[0]][0])
        with open(path, "w", encoding="utf-8") as f:
            yaml.safe_dump(self.cfg_dict, f, allow_unicode=True, sort_keys=False, width=100)
        # 状态栏反馈即可，避免频繁弹窗打断编辑
        self.status_var.set(f"方案已保存: {path}")
        self.yaml_var.set(path)  # 保存后立即作为套用方案

    # ── 参考文档反推（M4）─────────────────────────────────
    def extract_from_reference(self):
        ref = filedialog.askopenfilename(title="选择参考文档（符合目标格式的范文）",
                                         filetypes=[("Word 文档", "*.docx")])
        if not ref:
            return
        base = self.yaml_var.get().strip() or self.cfg_path_var.get().strip() or None
        if base and not Path(base).exists():
            base = None
        out = filedialog.asksaveasfilename(
            title="保存提取出的方案", defaultextension=".yaml",
            initialfile=Path(ref).stem + "_方案.yaml",
            filetypes=[("YAML", "*.yaml *.yml")])
        if not out:
            return
        from wordformat.extract import extract_profile, save_yaml
        try:
            config, report = extract_profile(ref, base)
            save_yaml(config, out)
            report_path = Path(out).with_suffix(".提取报告.txt")
            report_path.write_text("\n".join(report), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("提取失败", str(e))
            return
        self.yaml_var.set(out)
        self.cfg_path_var.set(out)
        self.cfg_dict = self._load_cfg_dict(out)
        messagebox.showinfo(
            "提取完成",
            f"方案已生成并载入:\n{out}\n\n提取报告:\n{report_path}")

    def _select_tab(self, index: int):
        nb = None
        for child in self.root.winfo_children():
            if isinstance(child, ttk.Notebook):
                nb = child
                break
        if nb is not None:
            nb.select(index)

    # ── 方案路径 ─────────────────────────────────────────
    def _resolve_yaml_path(self) -> tuple[str, str]:
        """解析套用应使用的方案路径，返回 (路径, 说明)。

        Tab1 方案框为空时自动回落内置预设（修复：冻结版默认方案不可达）。"""
        path = self.yaml_var.get().strip()
        if path and Path(path).exists():
            return path, ""
        fallback = _default_preset()
        if fallback:
            if path:
                return fallback, f"填写的方案不存在，已自动改用内置方案: {fallback}"
            return fallback, "未填写方案，已自动使用内置预设"
        return path, ""

    def _on_preset_selected(self, _event=None):
        name = self.preset_cb.get()
        presets = _builtin_presets()
        if name in presets:
            path = str(presets[name])
            self.yaml_var.set(path)
            self.cfg_path_var.set(path)
            try:
                self.cfg_dict = self._load_cfg_dict(path)
            except Exception:
                self.cfg_dict = {}
            self.status_var.set(f"已载入内置方案: {name}")

    # ── 识别 ─────────────────────────────────────────────
    def browse_docx(self):
        path = filedialog.askopenfilename(title="选择 Word 文件",
                                          filetypes=[("Word 文档", "*.docx")])
        if path:
            self.docx_var.set(path)

    def browse_yaml(self):
        path = filedialog.askopenfilename(title="选择格式方案 YAML",
                                          filetypes=[("YAML", "*.yaml *.yml")])
        if path:
            self.yaml_var.set(path)
            self.cfg_path_var.set(path)
            self.cfg_dict = self._load_cfg_dict(path)

    def start_detect(self):
        docx = self.docx_var.get().strip()
        if not docx or not Path(docx).exists():
            messagebox.showwarning("提示", "请先选择有效的 Word 文件")
            return
        self.docx_path = docx
        self._set_busy(True, "识别中…（公式段/表格/目录行走结构规则，其余走模型）")
        threading.Thread(target=self._detect_worker, args=(docx,), daemon=True).start()

    def _detect_worker(self, docx: str):
        try:
            data = DocxBase(docx, configpath=None).parse()
        except Exception as e:  # 线程内异常必须回主线程报告
            self.root.after(0, self._detect_failed, str(e))
            return
        self.root.after(0, self._populate, data)

    def _populate(self, data: list[dict]):
        self.data = data
        self.tree.delete(*self.tree.get_children())
        for i, item in enumerate(data):
            cat = item["category"]
            tags = []
            if item.get("needs_review"):
                tags.append("review")
            if cat in VOIDNODELIST:
                tags.append("void")
            self.tree.insert("", "end", iid=str(i), values=(
                i, _cat_cn(cat), f"{item.get('score', 0):.2f}",
                "是" if item.get("needs_review") else "",
                (item.get("paragraph") or "")[:80],
            ), tags=tags)
        self._fill_original_text(data)
        n_review = sum(1 for d in data if d.get("needs_review"))
        self._set_busy(False, f"识别完成：{len(data)} 段，其中 {n_review} 段建议人工复核（黄色行）。"
                              "单击行可在右侧原文定位，双击「类别」单元格可改判。")

    def _fill_original_text(self, data: list[dict]):
        """右侧原文窗：一段一行（空段/图片段/公式段显示灰色占位）。"""
        self.原文.configure(state="normal")
        self.原文.delete("1.0", "end")
        for item in data:
            text = (item.get("paragraph") or "").strip()
            if text:
                self.原文.insert("end", text + "\n")
            elif item["category"] == "figure_image":
                self.原文.insert("end", "〔图片段落〕\n", ("placeholder",))
            elif item["category"] == "equation_para":
                self.原文.insert("end", "〔公式段落〕\n", ("placeholder",))
            else:
                self.原文.insert("end", "〔空段落〕\n", ("placeholder",))
        self.原文.configure(state="disabled")

    # ── 列表 ↔ 原文联动 ───────────────────────────────────
    def _on_tree_select(self, _event=None):
        sel = self.tree.selection()
        if not sel or not self.data:
            return
        idx = int(sel[0])
        self.原文.tag_remove("current", "1.0", "end")
        line = idx + 1  # 原文窗一段一行
        self.原文.tag_add("current", f"{line}.0", f"{line}.end")
        self.原文.see(f"{max(line - 1, 1)}.0")

    # ── 改判 ─────────────────────────────────────────────
    def _close_cat_editor(self):
        if self._cat_editor is not None:
            self._cat_editor.destroy()
            self._cat_editor = None

    def _on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id or col != "#2":
            return
        idx = int(row_id)
        self._close_cat_editor()
        x, y, w, h = self.tree.bbox(row_id, col)
        # 编辑器以列表为父容器、依附在单元格原位（坐标即列表内坐标）
        cb = ttk.Combobox(self.tree, values=CN_LIST_ORDERED, state="readonly")
        cb.set(_cat_cn(self.data[idx]["category"]))
        cb.place(in_=self.tree, x=x, y=y, width=w, height=h)
        cb.focus_set()
        self._cat_editor = cb
        try:
            cb.event_generate("<Down>")  # 自动展开下拉列表
        except tk.TclError:
            pass

        def confirm(_event=None):
            cn = cb.get()
            self._close_cat_editor()
            new_en = CN_TO_CATEGORY.get(cn)
            if not new_en or new_en == self.data[idx]["category"]:
                return
            self.data[idx]["category"] = new_en
            self.data[idx]["needs_review"] = False
            self.data[idx]["comment"] = "人工改判"
            values = list(self.tree.item(row_id, "values"))
            values[1] = _cat_cn(new_en)
            values[3] = ""
            tags = ["void"] if new_en in VOIDNODELIST else []
            self.tree.item(row_id, values=values, tags=tags)
            self.status_var.set(f"已改判第 {idx} 段 → {_cat_cn(new_en)}")

        def cancel(_event=None):
            self._close_cat_editor()

        cb.bind("<<ComboboxSelected>>", confirm)
        cb.bind("<Return>", confirm)
        cb.bind("<Escape>", cancel)

    # ── 列表 ↔ 原文联动 ───────────────────────────────────
    def _on_tree_select(self, _event=None):
        sel = self.tree.selection()
        if not sel or not self.data:
            return
        idx = int(sel[0])
        self.原文.tag_remove("current", "1.0", "end")
        line = idx + 1  # 原文窗一段一行
        self.原文.tag_add("current", f"{line}.0", f"{line}.end")
        self.原文.see(f"{max(line - 1, 1)}.0")

    # ── 保存 / 套用 ──────────────────────────────────────
    def save_json(self):
        if not self.data:
            messagebox.showwarning("提示", "请先识别")
            return
        path = filedialog.asksaveasfilename(
            title="保存结构 JSON", defaultextension=".json",
            initialfile=Path(self.docx_path or "结构").stem + "_结构.json",
            filetypes=[("JSON", "*.json")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        self.status_var.set(f"结构 JSON 已保存: {path}")

    def start_apply(self):
        if not self.data or not self.docx_path:
            messagebox.showwarning("提示", "请先识别")
            return
        yaml_path, note = self._resolve_yaml_path()
        if not yaml_path or not Path(yaml_path).exists():
            messagebox.showwarning(
                "提示", "未找到格式方案：请到「格式方案」页打开/提取，或在下方选择内置方案")
            return
        if note:
            self.yaml_var.set(yaml_path)
            self.status_var.set(note)
        json_path = Path(tempfile.gettempdir()) / f"wordformat_gui_{os.getpid()}.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        self._set_busy(True, "套用格式中…")
        threading.Thread(target=self._apply_worker, args=(
            str(json_path), self.docx_path, yaml_path,
            self.comments_var.get()), daemon=True).start()

    def _apply_worker(self, json_path: str, docx: str, yaml_path: str, comments: bool):
        try:
            out = auto_format_thesis_document(
                jsonpath=json_path, docxpath=docx, configpath=yaml_path,
                savepath=str(Path(docx).parent / "格式化输出"),
                check=False, skip_comments=not comments,
            )
        except Exception as e:
            self.root.after(0, self._apply_failed, str(e))
            return
        self.root.after(0, self._apply_done, out)

    def _apply_failed(self, msg: str):
        self._set_busy(False, "套用失败")
        messagebox.showerror("套用失败", msg)

    def _apply_done(self, out: str):
        self._set_busy(False, f"已输出: {out}")
        if messagebox.askyesno("完成", f"已生成:\n{out}\n\n是否立即打开?"):
            os.startfile(out)

    # ── 公共 ─────────────────────────────────────────────
    def _set_busy(self, busy: bool, status: str):
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.detect_btn.configure(state=state)
        self.apply_btn.configure(state="normal" if (not busy and self.data) else "disabled")
        self.save_btn.configure(state="normal" if (not busy and self.data) else "disabled")
        self.status_var.set(status)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
