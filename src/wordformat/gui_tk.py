#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""Tkinter 识别预览/改判界面（需求文档 D6 形态）：

    Word 文件 → [开始识别] → 结构列表预览（类别下拉改判）→ [套用格式方案] → 新 docx

- 识别：复用 base.DocxBase.parse（含公式段/表内段落/目录行/题注等结构规则）
- 改判：双击"类别"单元格出现下拉框（中文显示），改判写入识别数据
- 联动：右侧原文窗口按段落逐行排布整篇文本，点击列表行自动滚动定位并高亮该段
- 套用：auto_format_thesis_document（check=False），可选不写审计批注
全部本地离线运行。
"""
import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from wordformat.base import DocxBase
from wordformat.pipeline.orchestrate import auto_format_thesis_document
from wordformat.settings import VOIDNODELIST

# 类别全集 = 模型标签 + 结构规则标签
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
CATEGORY_CN_LIST = [CAT_CN.get(c, c) for c in CATEGORIES]


def _cat_cn(cat: str) -> str:
    return CAT_CN.get(cat, cat)


def _default_preset() -> str:
    """随仓库分发的北理工预设；不存在则留空让用户自选。"""
    candidate = Path(__file__).resolve().parents[2] / "example" / "北理工毕设报告.yaml"
    return str(candidate) if candidate.exists() else ""


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Word 格式整理器（wordformat）")
        root.geometry("1280x760")
        root.minsize(1024, 620)

        self.data: list[dict] = []
        self.docx_path: str | None = None
        self._busy = False

        self._build_widgets()
        self.yaml_var.set(_default_preset())

    # ── 界面 ─────────────────────────────────────────────
    def _build_widgets(self):
        top = ttk.Frame(self.root, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="Word 文件:").grid(row=0, column=0, sticky="w")
        self.docx_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.docx_var, width=56).grid(row=0, column=1, padx=4)
        ttk.Button(top, text="浏览…", command=self.browse_docx).grid(row=0, column=2)
        ttk.Label(top, text="格式方案:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.yaml_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.yaml_var, width=56).grid(row=1, column=1, padx=4, pady=(4, 0))
        ttk.Button(top, text="浏览…", command=self.browse_yaml).grid(row=1, column=2, pady=(4, 0))
        ttk.Button(top, text="从参考文档提取…", command=self.extract_from_reference).grid(
            row=1, column=3, padx=(6, 0), pady=(4, 0))

        bar = ttk.Frame(self.root, padding=(10, 6))
        bar.pack(fill="x")
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

        # 左：识别列表 ｜ 右：原文窗口（可拖动分隔条调宽）
        pane = ttk.Panedwindow(self.root, orient="horizontal")
        pane.pack(fill="both", expand=True, padx=10, pady=(0, 4))

        left = ttk.Frame(pane)
        pane.add(left, weight=3)
        columns = ("idx", "cat", "score", "review", "text")
        headers = ("序号", "类别（双击改判）", "置信度", "需复核", "文本摘录")
        widths = (50, 130, 58, 60, 330)
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

        self.status_var = tk.StringVar(value="就绪——选择 Word 文件后点击「开始识别」")
        ttk.Label(self.root, textvariable=self.status_var, relief="sunken",
                  anchor="w", padding=(6, 3)).pack(fill="x", side="bottom")

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
    def _on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id or col != "#2":  # 只允许改"类别"列
            return
        idx = int(row_id)
        x, y, w, h = self.tree.bbox(row_id, col)
        current_en = self.data[idx]["category"]
        cb = ttk.Combobox(self.root, values=CATEGORY_CN_LIST, state="readonly")
        cb.set(_cat_cn(current_en))
        cb.place(x=x + self.tree.winfo_x(), y=y + self.tree.winfo_y(), width=w, height=h)
        cb.focus_set()

        def confirm(_event=None):
            cn = cb.get()
            cb.destroy()
            new_en = CN_TO_CATEGORY.get(cn)
            if not new_en or new_en == current_en:
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
            cb.destroy()

        cb.bind("<<ComboboxSelected>>", confirm)
        cb.bind("<Escape>", cancel)
        cb.bind("<FocusOut>", cancel)

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
        yaml_path = self.yaml_var.get().strip()
        if not yaml_path or not Path(yaml_path).exists():
            messagebox.showwarning("提示", "格式方案 YAML 不存在")
            return
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

    # 参考文档反推（M4）
    def extract_from_reference(self):
        ref = filedialog.askopenfilename(title="选择参考文档（符合目标格式的范文）",
                                         filetypes=[("Word 文档", "*.docx")])
        if not ref:
            return
        base = self.yaml_var.get().strip() or None
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
        messagebox.showinfo(
            "提取完成",
            f"方案已生成:\n{out}\n\n提取报告:\n{report_path}")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
