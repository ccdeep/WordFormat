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
import io
import json
import os
import queue
import shutil
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

import yaml

from docx import Document
from docx.oxml.ns import qn

from wordformat.base import DocxBase
from wordformat.pipeline.orchestrate import auto_format_thesis_document
from wordformat.settings import VOIDNODELIST
from wordformat.structure.paragraph_walker import iter_document_paragraphs
from wordformat.utils.omml_text import omml_to_text
from wordformat.utils.word_view import open_in_word

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

# VML 老式图片（python-docx 的 nsmap 无 v: 前缀，只能手工拼）
_VML_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"

# 目录导航窗格的类别层级（不在表中的类别不进导航）
NAV_LEVELS = {
    "document_title": 0,
    "abstract_chinese_title": 1, "abstract_english_title": 1,
    "heading_level_1": 1, "heading_level_2": 2, "heading_level_3": 3,
    "heading_mulu": 1, "heading_fulu": 1,
    "references_title": 1, "acknowledgements_title": 1,
}

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
# UI 下拉（筛选/改判）不提供的类别：模型合并段的产物，无独立排版与改判意义
_HIDDEN_IN_UI = {"abstract_chinese_title_content", "abstract_english_title_content"}
_CATEGORY_ORDER_CN = [CAT_CN.get(c, c) for c in _CATEGORY_ORDER
                      if c in CAT_CN and c not in _HIDDEN_IN_UI]
_CATEGORY_ORDER_CN += [CAT_CN[c] for c in CATEGORIES
                       if c not in _CATEGORY_ORDER and c not in _HIDDEN_IN_UI]
CN_LIST_ORDERED = _CATEGORY_ORDER_CN

# ── 可配置节（Tab2 编辑面板）────────────────────────────────
# kind: full = 字体+段落全量参数；align = 仅对齐/缩进（图片段落）
_SECTIONS = [
    ("document.title", "文档标题", "full"),
    ("body.text", "正文", "full"),
    ("headings.level_1", "一级标题", "full"),
    ("headings.level_2", "二级标题", "full"),
    ("headings.level_3", "三级标题", "full"),
    ("abstract.chinese.title", "中文摘要标题", "full"),
    ("abstract.chinese.body", "中文摘要正文", "full"),
    ("abstract.chinese.keywords", "中文关键词", "full"),
    ("abstract.english.title", "英文摘要标题", "full"),
    ("abstract.english.body", "英文摘要正文", "full"),
    ("abstract.english.keywords", "英文关键词", "full"),
    ("references.title", "参考文献标题", "full"),
    ("references.entry", "参考文献条目", "full"),
    ("acknowledgements.title", "致谢标题", "full"),
    ("acknowledgements.content", "致谢正文", "full"),
    ("figures.caption", "图注", "full"),
    ("tables.caption", "表注", "full"),
    ("tables.text", "表格内文字", "full"),
    ("math.block", "公式段落", "full"),
    ("figures.image", "图片段落", "align"),
]
# 不通过本页排版的类别（识别预览里可能出现，特此说明）
_NON_LAYOUT_CATS = "目录标题、目录行、附录标题、页脚、其他（封面/声明）"
_SIZE_CHOICES = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
                 "四号", "小四", "五号", "小五", "六号", "七号"]
_ALIGN_CHOICES = ["两端对齐", "居中对齐", "左对齐", "右对齐"]
_RULE_CHOICES = ["单倍行距", "1.5倍行距", "2倍行距", "多倍行距", "固定值", "最小值"]
_BOLD_CHOICES = ["加粗", "不加粗"]
_ITALIC_CHOICES = ["斜体", "不斜体"]
_UNDERLINE_CHOICES = ["下划线", "无下划线"]
_CN_FONT_CHOICES = ["宋体", "仿宋_GB2312", "楷体", "黑体", "华文细黑", "微软雅黑"]
_FONT_ROWS = [
    ("中文字体", ("combobox_edit", _CN_FONT_CHOICES), "chinese_font_name"),
    ("西文字体", "entry", "english_font_name"),
    ("字号", ("combobox", _SIZE_CHOICES), "font_size"),
    ("加粗", ("combobox", _BOLD_CHOICES), "bold"),
    ("斜体", ("combobox", _ITALIC_CHOICES), "italic"),
    ("下划线", ("combobox", _UNDERLINE_CHOICES), "underline"),
    ("字体颜色", "entry", "font_color"),
]
_PARA_ROWS = [
    ("对齐", ("combobox", _ALIGN_CHOICES), "alignment"),
    ("行距规则", ("combobox", _RULE_CHOICES), "line_spacingrule"),
    ("行距值", "entry", "line_spacing"),
    ("段前", "entry", "space_before"),
    ("段后", "entry", "space_after"),
    ("左缩进", "entry", "left_indent"),
    ("右缩进", "entry", "right_indent"),
    ("首行缩进", "entry", "first_line_indent"),
]
_PARA_FIELD_KEYS = [f for _, _, f in _PARA_ROWS]
_BOOL_LABELS = {"bold": ("加粗", "不加粗"), "italic": ("斜体", "不斜体"),
                "underline": ("下划线", "无下划线")}


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


def _user_presets_dir() -> Path:
    """用户方案固定文件夹：%APPDATA%/WordFormat/presets。

    与软件绑定的固定位置，与 exe/快捷方式在哪无关；
    内置预设在首次运行时播种到此处，用户可直接修改自己的副本。"""
    base = os.getenv("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    d = Path(base) / "WordFormat" / "presets"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _seed_user_presets() -> None:
    """把内置预设播种到用户方案文件夹（同名不覆盖，保护用户修改）。"""
    user_dir = _user_presets_dir()
    for name, src in _builtin_presets().items():
        dst = user_dir / (name + ".yaml")
        if not dst.exists():
            try:
                shutil.copy2(src, dst)
            except OSError:
                pass  # 播种失败不影响使用：内置预设仍可从程序包内读取


def _all_presets() -> dict[str, Path]:
    """全部可用方案 = 用户方案文件夹（已播种内置副本）+ 程序包内预设兜底。"""
    _seed_user_presets()
    user_dir = _user_presets_dir()
    presets = {f.stem: f for f in sorted(user_dir.glob("*.yaml"))}
    for name, path in _builtin_presets().items():
        presets.setdefault(name, path)
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
        self._source_docx: str | None = None  # 识别时的原文档（套用后 docx_path 会指向输出）
        self._para_sources: list[dict] = []   # 与 data 对齐的段落对象（富渲染/Word 定位用）
        self._photo_refs: list = []           # 防 PhotoImage 被垃圾回收
        self._para_marks: list[str] = []      # 每段在原文 Text 里的起始索引
        self._tree_group_cache: dict[str, list[int]] = {}  # 树行 iid → 段落成员（表格折叠用）
        self._nav_syncing = False             # 目录↔列表联动时防事件互调死循环
        self._busy = False
        self._progress_queue: queue.Queue | None = None
        self._progress_sink_id = None
        self._worker_thread: threading.Thread | None = None
        self._task_done = False
        self._task_error: str | None = None
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
        all_presets = _all_presets()
        self.preset_cb["values"] = list(all_presets) or ["（未找到内置方案）"]
        if all_presets:
            self.preset_cb.set("北理工毕设报告" if "北理工毕设报告" in all_presets
                               else next(iter(all_presets)))
        self.preset_cb.grid(row=2, column=1, columnspan=2, sticky="w", padx=4, pady=(4, 0))
        ttk.Label(top, text="选中即载入，也可在「格式方案」页编辑",
                  foreground="#888888").grid(row=2, column=3, columnspan=2, sticky="w", pady=(4, 0))
        self.preset_cb.bind("<<ComboboxSelected>>", self._on_preset_selected)
        ttk.Button(top, text="查看/编辑格式…", command=self.open_preset_editor).grid(
            row=2, column=3, padx=(6, 0), pady=(4, 0))
        ttk.Button(top, text="打开方案文件夹", command=self.open_presets_folder).grid(
            row=2, column=4, padx=(6, 0), pady=(4, 0))

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
        self.word_btn = ttk.Button(bar, text="📖 在 Word 中查看选中段", command=self.view_in_word,
                                   state="disabled")
        self.word_btn.pack(side="left", padx=8)
        ttk.Label(bar, text="类别筛选:").pack(side="left", padx=(14, 2))
        # 多选打勾：Menubutton 弹出可勾选菜单，可一次勾选多个类别
        self._filter_sel = {CN_TO_CATEGORY[cn]: tk.BooleanVar(value=False)
                            for cn in CN_LIST_ORDERED}
        self.filter_mb = ttk.Menubutton(bar, text="类别筛选：全部", width=26)
        filter_menu = tk.Menu(self.filter_mb, tearoff=0,
                              font=("Microsoft YaHei UI", 9))
        for cn in CN_LIST_ORDERED:
            filter_menu.add_checkbutton(label=cn,
                                        variable=self._filter_sel[CN_TO_CATEGORY[cn]],
                                        command=self._on_filter_change)
        filter_menu.add_separator()
        filter_menu.add_command(label="全选", command=lambda: self._set_all_filter(True))
        filter_menu.add_command(label="清空筛选", command=lambda: self._set_all_filter(False))
        self.filter_mb.configure(menu=filter_menu)
        self.filter_mb.pack(side="left")
        self.scope_var = tk.BooleanVar(value=False)
        self.scope_cb = ttk.Checkbutton(bar, text="只排版筛选出的类别",
                                        variable=self.scope_var)
        self.scope_cb.pack(side="left", padx=(10, 0))
        # 表格折叠：连续的表内段落折叠为一行（一个表格一个格式要求，整体改判）
        self._collapse_tables = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="表格折叠为一条", variable=self._collapse_tables,
                        command=self._refresh_tree).pack(side="left", padx=(10, 0))

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
        self.原文.tag_configure("math", foreground="#5B2D8E")
        self.原文.tag_configure("in_table", background="#EEF1F5")
        self.原文.tag_configure("table_sep", foreground="#C2C8D2",
                                font=("Microsoft YaHei UI", 7))
        self.原文.tag_raise("current")  # 高亮永远盖过表格底纹

        # ── 目录导航窗格（最右，像 Word 导航窗格）──────────
        navf = ttk.Frame(pane)
        pane.add(navf, weight=1)
        ttk.Label(navf, text="目录导航（点击跳转）", foreground="#666666").pack(anchor="w")
        self.nav_tree = ttk.Treeview(navf, show="tree", height=24, selectmode="browse")
        nvsb = ttk.Scrollbar(navf, orient="vertical", command=self.nav_tree.yview)
        self.nav_tree.configure(yscrollcommand=nvsb.set)
        self.nav_tree.pack(fill="both", expand=True, side="left")
        nvsb.pack(side="right", fill="y")
        self.nav_tree.bind("<<TreeviewSelect>>", self._on_nav_select)

        # 滚轮提速：默认每档 1 行太慢，统一改为 3 行（Word 手感）
        for w in (self.tree, self.原文, self.nav_tree):
            def _on_wheel(event, _w=w, _lines=3):
                _w.yview_scroll(-_lines if event.delta > 0 else _lines, "units")
                return "break"  # 拦下默认的 1 行滚动，避免叠加
            w.bind("<MouseWheel>", _on_wheel)

    # ══ Tab2：格式方案 ═══════════════════════════════════
    def _build_tab2(self, tab):
        top = ttk.Frame(tab)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="方案文件:").grid(row=0, column=0, sticky="w")
        self.cfg_path_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.cfg_path_var, width=58).grid(row=0, column=1, padx=4)
        ttk.Button(top, text="打开…", command=self.cfg_open).grid(row=0, column=2)
        ttk.Button(top, text="另存为…", command=self.cfg_save_as).grid(row=0, column=3, padx=4)
        ttk.Button(top, text="从参考文档提取…", command=self.extract_from_reference).grid(row=0, column=4)
        ttk.Button(top, text="保存修改", command=self.cfg_save).grid(row=0, column=5, padx=(8, 0))
        ttk.Button(top, text="方案总览", command=self._show_overview).grid(row=0, column=6, padx=(6, 0))

        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="元素类别:").grid(row=0, column=0, sticky="nw")
        self.section_list = tk.Listbox(body, width=22, height=21, exportselection=False)
        for _, label, _kind in _SECTIONS:
            self.section_list.insert("end", label)
        self.section_list.grid(row=1, column=0, sticky="ns")
        self.section_list.bind("<<ListboxSelect>>", self._on_section_select)
        note = ("值示例：行距值 '1.5倍' 或 '22磅'；段前段后 '0.5行' 或 '12磅'；"
                "缩进 '2字符'，悬挂用 '-2字符'。未填写的字段保持原值。")
        ttk.Label(body, text=note, foreground="#888888",
                  wraplength=210, justify="left").grid(row=2, column=0, sticky="nw")
        ttk.Label(body, text="以下类别不通过本页排版：\n" + _NON_LAYOUT_CATS
                  + "\n公式段落由软件自动居中/编号。",
                  foreground="#999999", wraplength=210,
                  justify="left").grid(row=3, column=0, sticky="nw", pady=(8, 0))

        panel = ttk.Frame(body)
        panel.grid(row=1, column=1, sticky="nw", padx=(12, 0))

        # 字体组
        self.font_frame = ttk.LabelFrame(panel, text="字体", padding=8)
        self.font_frame.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        self.font_fields = {}
        for i, (label, kind, field) in enumerate(_FONT_ROWS):
            ttk.Label(self.font_frame, text=label + ":").grid(row=i, column=0, sticky="w", pady=3)
            if kind == "entry":
                w = ttk.Entry(self.font_frame, width=22)
            elif kind == "combobox_edit":
                w = ttk.Combobox(self.font_frame, width=20, values=kind[1])
            else:
                w = ttk.Combobox(self.font_frame, width=20, values=kind[1], state="readonly")
            w.grid(row=i, column=1, sticky="w", padx=(6, 0))
            self.font_fields[field] = w
        # 段落组
        self.para_frame = ttk.LabelFrame(panel, text="段落", padding=8)
        self.para_frame.grid(row=0, column=1, sticky="ns")
        self.para_fields = {}
        self.para_labels = {}
        for i, (label, kind, field) in enumerate(_PARA_ROWS):
            lab = ttk.Label(self.para_frame, text=label + ":")
            lab.grid(row=i, column=0, sticky="w", pady=3)
            self.para_labels[field] = lab
            if kind == "entry":
                w = ttk.Entry(self.para_frame, width=22)
            else:
                w = ttk.Combobox(self.para_frame, width=20, values=kind[1], state="readonly")
            w.grid(row=i, column=1, sticky="w", padx=(6, 0))
            self.para_fields[field] = w
        self.section_list.selection_set(0)
        self._on_section_select()

    def _show_overview(self):
        """弹窗展示当前方案全部类别的关键参数（只读总览）。"""
        if not getattr(self, "cfg_dict", None):
            self.cfg_dict = self._load_cfg_dict(self.cfg_path_var.get()) \
                if self.cfg_path_var.get() else {}
        lines = []
        for section, label, kind in _SECTIONS:
            node = self.cfg_dict
            try:
                for part in section.split("."):
                    node = node[part]
            except (KeyError, TypeError):
                node = {}
            font, para = node.get("font") or {}, node.get("paragraph") or {}
            parts = []
            if font.get("chinese_font_name"):
                parts.append(f"{font['chinese_font_name']}")
            if font.get("english_font_name"):
                parts.append(f"/{font['english_font_name']}")
            if font.get("font_size"):
                parts.append(f" {font['font_size']}")
            if font.get("bold"):
                parts.append(" 加粗")
            if para.get("alignment"):
                parts.append(f" {para['alignment']}")
            if para.get("line_spacingrule") or para.get("line_spacing"):
                parts.append(f" 行距:{para.get('line_spacingrule', '')}"
                             f"{para.get('line_spacing', '')}")
            if para.get("first_line_indent"):
                parts.append(f" 缩进:{para['first_line_indent']}")
            lines.append(f"{label}：{' '.join(parts) if parts else '（未设置）'}")
        win = tk.Toplevel(self.root)
        win.title("方案总览（只读）")
        win.geometry("620x520")
        txt = tk.Text(win, wrap="word", font=("Microsoft YaHei UI", 10), padx=10, pady=8)
        sb = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.pack(fill="both", expand=True, side="left")
        sb.pack(side="right", fill="y")
        txt.insert("1.0", "\n".join(lines))
        txt.configure(state="disabled")

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
        section, _label, kind = _SECTIONS[sel[0]]
        node = self._cfg_node(section)
        font = node.get("font") or {}
        # align 类别（figures.image）的对齐/缩进直接在节点上，无 paragraph 包装
        para = node if kind == "align" else (node.get("paragraph") or {})

        def _set_widget(w, val):
            text = str(val) if val is not None else ""
            if isinstance(w, ttk.Combobox):
                w.set(text)
            else:
                w.delete(0, "end")
                w.insert(0, text)

        # 字体组：仅 full 类别显示
        if kind == "full":
            self.font_frame.grid()
        else:
            self.font_frame.grid_remove()
        for field, w in self.font_fields.items():
            if kind != "full":
                _set_widget(w, "")
                continue
            val = font.get(field)
            if field in _BOOL_LABELS and val is not None:
                _set_widget(w, _BOOL_LABELS[field][0] if val else _BOOL_LABELS[field][1])
            else:
                _set_widget(w, val)
        # 段落组：align 类别（图片段落）显示 对齐/行距/首行缩进
        # （图片必须避开固定值行距，否则内联图片会被裁剪得显示不全）
        for field, w in self.para_fields.items():
            lab = self.para_labels[field]
            if kind == "align" and field not in ("alignment", "line_spacingrule",
                                                 "line_spacing", "first_line_indent"):
                w.grid_remove()
                lab.grid_remove()
                continue
            w.grid()
            lab.grid()
            _set_widget(w, para.get(field))

    def _collect_section(self, section: str, kind: str) -> None:
        """把面板字段写回 cfg_dict 对应节（空值不覆盖；align 类别只收对齐/缩进）。"""
        node = self._cfg_node(section)
        if kind == "align":
            font = None
            para = node
        else:
            font = node.setdefault("font", {})
            para = node.setdefault("paragraph", {})
        for field, w in self.font_fields.items():
            if kind != "full":
                break
            val = str(w.get()).strip()
            if not val:
                continue
            if not val:
                continue
            if field in _BOOL_LABELS:
                font[field] = val == _BOOL_LABELS[field][0]
            else:
                font[field] = val
        for field, w in self.para_fields.items():
            if kind == "align" and field not in ("alignment", "line_spacingrule",
                                                 "line_spacing", "first_line_indent"):
                continue
            val = str(w.get()).strip()
            if val:
                para[field] = val

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
            initialdir=str(_user_presets_dir()), initialfile="我的方案.yaml",
            filetypes=[("YAML", "*.yaml *.yml")])
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
            section, _label, kind = _SECTIONS[sel[0]]
            self._collect_section(section, kind)
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

    def open_presets_folder(self):
        """在资源管理器中打开用户方案固定文件夹。"""
        d = _user_presets_dir()
        self.status_var.set(f"方案文件夹: {d}")
        os.startfile(str(d))

    def open_preset_editor(self):
        """载入当前方案（未选则用内置预设）并切换到「格式方案」页，弹总览。"""
        path = self.yaml_var.get().strip() or _default_preset()
        if not path or not Path(path).exists():
            messagebox.showwarning("提示", "未找到格式方案文件")
            return
        self.cfg_open(path)
        self._select_tab(1)
        self._show_overview()

    def _on_preset_selected(self, _event=None):
        name = self.preset_cb.get()
        presets = _all_presets()
        if name in presets:
            path = str(presets[name])
            self.yaml_var.set(path)
            self.cfg_path_var.set(path)
            try:
                self.cfg_dict = self._load_cfg_dict(path)
            except Exception:
                self.cfg_dict = {}
            self.status_var.set(f"已载入内置方案: {name}")

    # ── 后台任务进度反馈 ──────────────────────────────────
    def _start_progress(self, title: str):
        """把 wordformat 管线的日志实时转发到状态栏（经队列，线程安全）。"""
        from loguru import logger
        self._progress_queue = queue.Queue()
        self._progress_sink_id = logger.add(
            lambda msg: self._progress_queue.put(str(msg).rstrip()),
            level="INFO", format="{message}")
        self._set_busy(True, title)
        self.root.configure(cursor="watch")
        self.root.after(120, self._poll)

    def _stop_progress(self):
        from loguru import logger
        if self._progress_sink_id is not None:
            try:
                logger.remove(self._progress_sink_id)
            except Exception:
                pass
            self._progress_sink_id = None
        self.root.configure(cursor="")

    def _poll(self):
        """主线程轮询器：分发后台任务结果 + 转发管线日志 + 线程意外退出兜底。"""
        tq = self._task_queue
        if tq is not None:
            try:
                while True:
                    kind, payload = tq.get_nowait()
                    if kind == "log":
                        self.status_var.set("⏳ " + str(payload)[:90])
                        continue
                    self._stop_progress()
                    if kind == "done":
                        self._done_handler(payload)
                    else:
                        self._fail_handler(payload)
                    return
            except queue.Empty:
                pass
        if not self._busy:
            return
        # 兜底：后台线程意外退出（未走正常完成/失败路径）时给出明确报错，
        # 杜绝"点了没反应"的静默失败
        wt = self._worker_thread
        if wt is not None and not wt.is_alive() and not self._task_done:
            self._stop_progress()
            self._set_busy(False, "后台任务异常退出")
            err = self._task_error or "后台线程意外终止，原因未知"
            messagebox.showerror("后台任务异常退出", err)
            return
        self.root.after(200, self._poll)

    def _task_done_handler(self, result):
        self._stop_progress()
        self._set_busy(False, "")
        if self._task_kind == "detect":
            self._populate(result)
        else:
            self._apply_done(result)

    def _task_fail_handler(self, msg: str):
        self._stop_progress()
        self._set_busy(False, "任务失败")
        title = "识别失败" if self._task_kind == "detect" else "套用失败"
        # 弹窗只给一句话原因，完整 traceback 放分隔线下方供排查
        lines = [ln.strip() for ln in msg.strip().splitlines() if ln.strip()]
        cause = lines[-1] if lines else msg
        if "is not a Word file" in cause:
            cause = ("该文件不是有效的 .docx 文档（可能是 .doc 老格式改了扩展名，"
                     f"或文件已损坏），请用 Word 打开后另存为 .docx 再试。\n{cause}")
        messagebox.showerror(title, f"{cause}\n\n{'—' * 30}\n完整错误信息：\n{msg}")

    # ── 识别 ─────────────────────────────────────────────
    def browse_docx(self):
        path = filedialog.askopenfilename(title="选择 Word 文件",
                                          filetypes=[("Word 文档", "*.docx")])
        if path:
            self.docx_var.set(path)

    def browse_yaml(self):
        path = filedialog.askopenfilename(
            title="选择格式方案 YAML", initialdir=str(_user_presets_dir()),
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
        self._source_docx = docx
        self._start_task("detect", "识别中…（模型加载与推理，进度见状态栏）",
                         self._detect_worker, (docx,))

    def _start_task(self, kind: str, title: str, target, args: tuple,
                    done=None, fail=None):
        """启动后台任务。结果经队列汇报，由主线程轮询分发：
        done(result)/fail(msg) 均在主线程执行；失败/意外退出有明确报错，
        绝不静默，也彻底避免跨线程调用 tkinter。"""
        self._task_kind = kind
        self._task_done = False
        self._task_error = None
        self._progress_queue = queue.Queue()
        self._task_queue = queue.Queue()
        from loguru import logger
        self._progress_sink_id = logger.add(
            lambda msg: self._task_queue.put(("log", str(msg).rstrip())),
            level="INFO", format="{message}")
        self._set_busy(True, title)
        self.root.configure(cursor="watch")
        self._done_handler = done or self._task_done_handler
        self._fail_handler = fail or self._task_fail_handler

        def wrapped():
            try:
                result = target(*args)
                self._task_queue.put(("done", result))
            except BaseException:
                import traceback
                self._task_error = traceback.format_exc()
                self._task_queue.put(("fail", self._task_error))
            finally:
                self._task_done = True

        self._worker_thread = threading.Thread(target=wrapped, daemon=True)
        self._worker_thread.start()
        self.root.after(120, self._poll)

    def _detect_worker(self, docx: str):
        return DocxBase(docx, configpath=None).parse()

    def _populate(self, data: list[dict]):
        self.data = data
        self._build_para_sources()
        self._refresh_tree()
        self._fill_original_text(data)
        self.word_btn.configure(state="normal")
        n_review = sum(1 for d in data if d.get("needs_review"))
        self._set_busy(False, f"识别完成：{len(data)} 段，其中 {n_review} 段建议人工复核（黄色行）。"
                              "单击行可在右侧原文定位，双击「类别」单元格可改判。")

    def _update_scope_ui(self):
        """范围勾选随筛选联动：勾了具体类别=自动勾选（标签写明后果）；
        未勾任何类别=禁用（全量排版，勾选无意义）。"""
        sel = self._selected_categories()
        if not sel:
            self.scope_var.set(False)
            self.scope_cb.configure(state="disabled",
                                    text="只排版筛选出的类别（先在类别筛选里打勾）")
            return
        self.scope_var.set(True)
        self.scope_cb.configure(state="normal",
                                text="只排版筛选出的类别，其余段落不动")

    def _selected_categories(self) -> set:
        """当前勾选的类别集合；空集合 = 全部。"""
        return {cat for cat, var in self._filter_sel.items() if var.get()}

    def _filter_summary(self) -> str:
        sel = self._selected_categories()
        if not sel:
            return "类别筛选：全部"
        if len(sel) == 1:
            return f"类别筛选：{_cat_cn(next(iter(sel)))}"
        return f"类别筛选：已选 {len(sel)} 类"

    def _on_filter_change(self):
        self.filter_mb.configure(text=self._filter_summary())
        self._update_scope_ui()
        self._refresh_tree()

    def _set_all_filter(self, on: bool):
        for var in self._filter_sel.values():
            var.set(on)
        self._on_filter_change()

    def _tree_groups(self) -> list[list[int]]:
        """把 data 切分成树的显示单元：开启折叠时，连续的表内段落合并为一组
        （一个表格一个格式要求）；关闭或无段落对象时每段一组。"""
        if not self._collapse_tables.get() or not self._para_sources:
            return [[i] for i in range(len(self.data))]
        groups = []
        i, n = 0, len(self.data)
        while i < n:
            if self._para_sources[i]["in_table"]:
                j = i
                while j < n and self._para_sources[j]["in_table"]:
                    j += 1
                groups.append(list(range(i, j)))
                i = j
            else:
                groups.append([i])
                i += 1
        return groups

    def _refresh_tree(self):
        """按类别筛选重建列表（iid 恒为该行首段的原始序号，改判/套用不受影响）。
        多选打勾：勾了类别只显示那些类别；一个都没勾 = 全部显示。
        表格组显示为一行，摘录带【表格·N段】标记。"""
        sel = self._selected_categories()
        self.tree.delete(*self.tree.get_children())
        self._tree_group_cache = {}
        for members in self._tree_groups():
            start = members[0]
            self._tree_group_cache[str(start)] = members
            in_table = len(members) > 1 or (self._para_sources and self._para_sources[start]["in_table"])
            if in_table:
                rep = ("table_text"
                       if any(self.data[m]["category"] == "table_text" for m in members)
                       else self.data[start]["category"])
            else:
                rep = self.data[start]["category"]
            if sel and rep not in sel:
                continue
            if in_table:
                first_text = next((self.data[m].get("paragraph") or ""
                                   for m in members if (self.data[m].get("paragraph") or "").strip()),
                                  "")
                excerpt = f"【表格·{len(members)}段】{first_text.strip()[:60]}"
                score = min(self.data[m].get("score", 0) for m in members)
                review = any(self.data[m].get("needs_review") for m in members)
                idx_text = f"{start}–{members[-1]}" if len(members) > 1 else str(start)
            else:
                item = self.data[start]
                excerpt = (item.get("paragraph") or "")[:80]
                score = item.get("score", 0)
                review = item.get("needs_review")
                idx_text = str(start)
            tags = []
            if review:
                tags.append("review")
            if rep in VOIDNODELIST:
                tags.append("void")
            self.tree.insert("", "end", iid=str(start), values=(
                idx_text, _cat_cn(rep), f"{score:.2f}",
                "是" if review else "", excerpt,
            ), tags=tags)
        self._refresh_nav()

    def _build_para_sources(self):
        """重开原文档按识别顺序取段落对象，供富渲染与 Word 定位使用。
        任何失败（文件没了/打开异常/段数对不上）都留空，界面回退纯文本显示，
        识别结果本身不受影响。"""
        self._para_sources = []
        path = self._source_docx
        if not path or not Path(path).exists():
            return
        try:
            doc = Document(path)
            paras = iter_document_paragraphs(doc)
        except Exception:
            return
        if len(paras) != len(self.data):
            return
        self._para_sources = [
            {"para": p, "in_table": in_table,
             "anchor": " ".join((p.text or "").split())}
            for p, in_table in paras]

    def _fill_original_text(self, data: list[dict]):
        """右侧原文窗：有段落对象时富渲染（图片嵌入/公式线性文本/表格底纹），
        否则回退纯文本一段一行。"""
        self._photo_refs = []
        self._para_marks = []
        self.原文.configure(state="normal")
        self.原文.delete("1.0", "end")
        if self._para_sources:
            self._fill_rich(data)
        else:
            self._fill_plain(data)
        self.原文.configure(state="disabled")

    def _fill_plain(self, data: list[dict]):
        """回退显示：与旧版一致，一段一行（空段/图片段/公式段显示灰色占位）。"""
        for item in data:
            self._para_marks.append(self.原文.index("end-1c"))
            text = (item.get("paragraph") or "").strip()
            if text:
                self.原文.insert("end", text + "\n")
            elif item["category"] == "figure_image":
                self.原文.insert("end", "〔图片段落〕\n", ("placeholder",))
            elif item["category"] == "equation_para":
                self.原文.insert("end", "〔公式段落〕\n", ("placeholder",))
            else:
                self.原文.insert("end", "〔空段落〕\n", ("placeholder",))

    def _fill_rich(self, data: list[dict]):
        prev_tr = None
        for i, (item, src) in enumerate(zip(data, self._para_sources)):
            self._para_marks.append(self.原文.index("end-1c"))
            para = src["para"]
            if src["in_table"]:
                tr = para._p.getparent().getparent()  # w:p → w:tc → w:tr
                if tr is not prev_tr and i > 0:
                    self.原文.insert("end", "· " * 26 + "\n", ("table_sep",))
                prev_tr = tr
                self.原文.insert("end", "▎ ", ("in_table",))
            wrote = False
            for kind, payload in self._para_segments(para):
                if kind == "text":
                    if payload.strip():
                        self.原文.insert("end", payload,
                                        ("in_table",) if src["in_table"] else ())
                        wrote = True
                elif kind == "math":
                    if payload:
                        self.原文.insert("end", f" ⟨{payload}⟩", ("math",))
                        wrote = True
                elif kind == "image":
                    if wrote:
                        self.原文.insert("end", "\n")
                    self._insert_image(para, payload)
                    wrote = True
            if not wrote:
                label = {"figure_image": "〔图片段落〕", "equation_para": "〔公式段落〕"}.get(
                    item["category"], "〔空段落〕")
                self.原文.insert("end", label, ("placeholder",))
            self.原文.insert("end", "\n")

    def _para_segments(self, para):
        """按文档顺序产出段落内容片段：(kind, payload)。
        kind=text 普通文字 / math 公式线性文本 / image 含图 run 元素。"""
        for child in para._p:
            tag = child.tag
            if tag == qn("m:oMathPara"):
                for om in child.findall(qn("m:oMath")):
                    yield "math", omml_to_text(om)
            elif tag == qn("m:oMath"):
                yield "math", omml_to_text(child)
            elif tag == qn("w:r"):
                yield from self._run_segments(child)
            elif tag == qn("w:hyperlink"):
                for r in child.findall(qn("w:r")):
                    yield from self._run_segments(r)

    @staticmethod
    def _run_segments(r_elem):
        for t in r_elem.findall(qn("w:t")):
            if t.text:
                yield "text", t.text
        # 兼容 DrawingML（常规图片）与 VML（老格式）两种图
        if (r_elem.find(".//" + qn("a:blip")) is not None
                or r_elem.find(".//" + _VML_IMAGEDATA) is not None):
            yield "image", r_elem

    def _insert_image(self, para, run_elem):
        """把 run 里的图片渲染进文本框（等比缩到面板宽），失败回退灰色占位。"""
        rids = [b.get(qn("r:embed")) for b in run_elem.findall(".//" + qn("a:blip"))]
        rids += [v.get(qn("r:id")) for v in run_elem.findall(".//" + _VML_IMAGEDATA)]
        for rid in filter(None, rids):
            try:
                blob = para.part.related_parts[rid].blob
                photo = self._blob_to_photo(blob)
            except Exception:
                photo = None
            if photo is not None:
                self.原文.image_create("end-1c", image=photo)
                self.原文.insert("end", "\n")
            else:
                self.原文.insert("end", "〔图片无法预览〕\n", ("placeholder",))
        if not rids:
            self.原文.insert("end", "〔图片段落〕\n", ("placeholder",))

    def _blob_to_photo(self, blob: bytes):
        from PIL import Image, ImageTk
        img = Image.open(io.BytesIO(blob))
        try:
            img.load(dpi=192)  # WMF/EMF 矢量图按 192dpi 点阵化
        except TypeError:
            img.load()
        max_w = self._preview_width()
        if img.width > max_w:
            img = img.resize((max_w, max(1, round(img.height * max_w / img.width))))
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGB")
        photo = ImageTk.PhotoImage(img)
        self._photo_refs.append(photo)
        return photo

    def _preview_width(self) -> int:
        w = self.原文.winfo_width()
        max_w = (w - 30) if w > 80 else 480  # 扣除 padding 与滚动条
        return max(120, min(max_w, 640))
        self.原文.configure(state="disabled")

    # ── 列表 ↔ 原文联动 ───────────────────────────────────
    def _on_tree_select(self, _event=None):
        sel = self.tree.selection()
        if not sel or not self.data:
            return
        idx = int(sel[0])
        members = self._tree_group_cache.get(sel[0], [idx])
        self.原文.tag_remove("current", "1.0", "end")
        if self._para_marks and members[0] < len(self._para_marks):
            start = self._para_marks[members[0]]
            end_at = members[-1] + 1  # 折叠组：高亮整张表格
            end = (self._para_marks[end_at] if end_at < len(self._para_marks)
                   else "end-1c")
            self.原文.tag_add("current", start, end)
            self.原文.see(start)
        if not self._nav_syncing and self.nav_tree.exists(str(idx)):
            self._nav_syncing = True
            self.nav_tree.selection_set(str(idx))
            self.nav_tree.see(str(idx))
            self._nav_syncing = False

    def _refresh_nav(self):
        """目录导航窗格：按标题类层级建树（文档标题/摘要标题为顶层，
        一二三级标题逐级缩进，附录/参考文献/致谢等专项标题为一级）。"""
        self.nav_tree.delete(*self.nav_tree.get_children())
        last_at: dict[int, str] = {}
        for i, item in enumerate(self.data):
            lv = NAV_LEVELS.get(item["category"])
            if lv is None:
                continue
            shallower = [l for l in last_at if l < lv]
            parent = last_at[max(shallower)] if shallower else ""
            text = ((item.get("paragraph") or "").strip() or _cat_cn(item["category"]))[:48]
            self.nav_tree.insert(parent, "end", iid=str(i), text=text, open=True)
            last_at = {k: v for k, v in last_at.items() if k <= lv}
            last_at[lv] = str(i)

    def _on_nav_select(self, _event=None):
        if self._nav_syncing:
            return
        sel = self.nav_tree.selection()
        if not sel:
            return
        iid = sel[0]
        if self.tree.exists(iid):
            self.tree.see(iid)
            self.tree.selection_set(iid)  # 触发 _on_tree_select 完成右侧高亮
        else:
            # 该段被筛选隐藏：临时直接高亮原文，不改左侧列表
            idx = int(iid)
            self.原文.tag_remove("current", "1.0", "end")
            if self._para_marks and idx < len(self._para_marks):
                start = self._para_marks[idx]
                end_at = idx + 1
                end = (self._para_marks[end_at] if end_at < len(self._para_marks)
                       else "end-1c")
                self.原文.tag_add("current", start, end)
                self.原文.see(start)

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
        current_cn = self.tree.item(row_id, "values")[1]  # 显示的类别（表格组为代表类别）
        cb.set(current_cn)
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
            members = self._tree_group_cache.get(row_id, [idx])
            for m in members:
                # 表格整体改判：图内图片段保持图片类，不跟着降级
                if (self.data[m]["category"] == "figure_image"
                        and new_en != "figure_image"):
                    continue
                self.data[m]["category"] = new_en
                self.data[m]["needs_review"] = False
                self.data[m]["comment"] = "人工改判"
            values = list(self.tree.item(row_id, "values"))
            values[1] = _cat_cn(new_en)
            values[3] = ""
            tags = ["void"] if new_en in VOIDNODELIST else []
            self.tree.item(row_id, values=values, tags=tags)
            self._refresh_nav()
            if len(members) > 1:
                self.status_var.set(
                    f"已整体改判第 {members[0]}–{members[-1]} 段（{len(members)} 段）→ {_cat_cn(new_en)}")
            else:
                self.status_var.set(f"已改判第 {idx} 段 → {_cat_cn(new_en)}")

        def cancel(_event=None):
            self._close_cat_editor()

        cb.bind("<<ComboboxSelected>>", confirm)
        cb.bind("<Return>", confirm)
        cb.bind("<Escape>", cancel)

    # ── Word 联动 ────────────────────────────────────────
    def view_in_word(self):
        if not self.data or not self._source_docx:
            messagebox.showwarning("提示", "请先识别")
            return
        anchor = ""
        sel = self.tree.selection()
        if sel:
            idx = int(sel[0])
            if idx < len(self._para_sources) and self._para_sources[idx]:
                anchor = self._para_sources[idx]["anchor"]
            else:
                anchor = " ".join((self.data[idx].get("paragraph") or "").split())
        try:
            msg = open_in_word(self._source_docx, anchor)
        except Exception as e:
            messagebox.showwarning("Word 联动", f"无法调用 Word：{e}")
            return
        self.status_var.set(msg)

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
        # 套用范围：勾选"只排版筛选出的类别"时，仅勾选的类别参与排版
        only_categories = None
        if self.scope_var.get():
            only_categories = self._selected_categories()
            if not only_categories:
                messagebox.showwarning("提示", "未勾选任何筛选类别，无法限定套用")
                return
        json_path = Path(tempfile.gettempdir()) / f"wordformat_gui_{os.getpid()}.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        self._start_task("apply", "套用格式中…（样式修正/公式排版/保存，进度见状态栏）",
                         self._apply_worker,
                         (str(json_path), self.docx_path, yaml_path,
                          self.comments_var.get(), only_categories))

    def _apply_worker(self, json_path: str, docx: str, yaml_path: str,
                      comments: bool, only_categories=None):
        return auto_format_thesis_document(
            jsonpath=json_path, docxpath=docx, configpath=yaml_path,
            savepath=str(Path(docx).parent / "格式化输出"),
            check=False, skip_comments=not comments,
            only_categories=only_categories,
        )

    def _apply_failed(self, msg: str):
        self._set_busy(False, "套用失败")
        messagebox.showerror("套用失败", msg)

    def _apply_done(self, out: str):
        self._set_busy(False, f"已输出: {out}")
        # 限定范围的套用：当前文档自动切换为输出文件，可改勾选继续套用（分段累积）
        if self.scope_var.get() and self._selected_categories():
            self.docx_var.set(out)
            self.docx_path = out
            self.status_var.set(f"已输出并切换当前文档: {out}——可换一个筛选类别继续套用")
            return
        if messagebox.askyesno("完成", f"已生成:\n{out}\n\n是否立即打开?"):
            os.startfile(out)

    # ── 公共 ─────────────────────────────────────────────
    def _set_busy(self, busy: bool, status: str):
        self._busy = busy
        if not busy:
            self._stop_progress()
        state = "disabled" if busy else "normal"
        self.detect_btn.configure(state=state)
        self.apply_btn.configure(state="normal" if (not busy and self.data) else "disabled")
        self.save_btn.configure(state="normal" if (not busy and self.data) else "disabled")
        self.word_btn.configure(state="normal" if (not busy and self.data) else "disabled")
        self.status_var.set(status)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
