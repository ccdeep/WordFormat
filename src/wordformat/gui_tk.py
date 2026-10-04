#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""Tkinter 识别预览/改判界面（需求文档 D6 形态）：

    Word 文件 → [开始识别] → 结构列表预览（类别下拉改判）→ [套用格式方案] → 新 docx

- 识别：复用 base.DocxBase.parse（含公式段/表内段落/目录行/题注等结构规则）
- 改判：双击"类别"单元格出现下拉框，改判写入识别数据（needs_review 清零，标记人工改判）
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


def _default_preset() -> str:
    """随仓库分发的北理工预设；不存在则留空让用户自选。"""
    candidate = Path(__file__).resolve().parents[2] / "example" / "北理工毕设报告.yaml"
    return str(candidate) if candidate.exists() else ""


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Word 格式整理器（wordformat）")
        root.geometry("1080x720")
        root.minsize(920, 600)

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
        ttk.Entry(top, textvariable=self.docx_var, width=58).grid(row=0, column=1, padx=4)
        ttk.Button(top, text="浏览…", command=self.browse_docx).grid(row=0, column=2)
        ttk.Label(top, text="格式方案:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.yaml_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.yaml_var, width=58).grid(row=1, column=1, padx=4, pady=(4, 0))
        ttk.Button(top, text="浏览…", command=self.browse_yaml).grid(row=1, column=2, pady=(4, 0))

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

        columns = ("idx", "cat", "score", "review", "text")
        headers = ("序号", "类别（双击改判）", "置信度", "需复核", "文本摘录")
        widths = (56, 150, 64, 64, 560)
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=22)
        for col, head, width in zip(columns, headers, widths):
            self.tree.heading(col, text=head)
            self.tree.column(col, width=width, anchor="w")
        vsb = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(fill="both", expand=True, padx=(10, 0))
        vsb.pack(side="right", fill="y", pady=(0, 0))
        self.tree.tag_configure("review", background="#FFF3CD")
        self.tree.tag_configure("void", foreground="#8A8A8A")
        self.tree.bind("<Double-1>", self._on_double_click)

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
            data = DocxBase(docx).parse()
        except Exception as e:  # 线程内异常必须回主线程报告
            self.root.after(0, self._detect_failed, str(e))
            return
        self.root.after(0, self._populate, data)

    def _detect_failed(self, msg: str):
        self._set_busy(False, "识别失败")
        messagebox.showerror("识别失败", msg)

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
                i, cat, f"{item.get('score', 0):.2f}",
                "是" if item.get("needs_review") else "",
                (item.get("paragraph") or "")[:80],
            ), tags=tags)
        n_review = sum(1 for d in data if d.get("needs_review"))
        self._set_busy(False, f"识别完成：{len(data)} 段，其中 {n_review} 段建议人工复核（黄色行）。"
                              "双击「类别」单元格可改判。")

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
        current = self.data[idx]["category"]
        cb = ttk.Combobox(self.root, values=CATEGORIES, state="readonly")
        cb.set(current)
        cb.place(x=x + self.tree.winfo_x(), y=y + self.tree.winfo_y(), width=w, height=h)
        cb.focus_set()

        def confirm(event=None):
            new_cat = cb.get()
            cb.destroy()
            if new_cat and new_cat != self.data[idx]["category"]:
                self.data[idx]["category"] = new_cat
                self.data[idx]["needs_review"] = False
                self.data[idx]["comment"] = "人工改判"
                tags = ["void"] if new_cat in VOIDNODELIST else []
                values = list(self.tree.item(row_id, "values"))
                values[1] = new_cat
                values[3] = ""
                self.tree.item(row_id, values=values, tags=tags)
                self.status_var.set(f"已改判第 {idx} 段 → {CAT_CN.get(new_cat, new_cat)}")

        def cancel(event=None):
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


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
