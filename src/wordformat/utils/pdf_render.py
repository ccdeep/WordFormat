"""PDF 页面渲染与文本定位（右侧"页面版式"预览用），基于 pypdfium2。

页面来自 Word 的 SaveAs2 导出（见 word_view.export_pdf），因此版式即 Word
的分页打印版式。渲染返回 PIL 图像（单页约 30~60ms），文本定位把段落锚文本
映射到页码（忽略 PDF 提取文本的空白差异）。
"""
import re


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


class PdfPages:
    """打开一个 PDF，提供整页渲染与锚文本定位页码。"""

    def __init__(self, path: str):
        import pypdfium2 as pdfium
        self._pdf = pdfium.PdfDocument(path)
        self.count = len(self._pdf)
        self.page_w, self.page_h = self._pdf[0].get_size()
        self._texts: list[str] | None = None

    def render(self, index: int, target_w: int):
        """渲染第 index 页为 PIL 图像，宽度缩放到 target_w 像素。"""
        page = self._pdf[index]
        return page.render(scale=target_w / self.page_w).to_pil()

    def _page_texts(self) -> list[str]:
        if self._texts is None:
            out = []
            for i in range(self.count):
                tp = self._pdf[i].get_textpage()
                out.append(tp.get_text_bounded())
            self._texts = out
        return self._texts

    def find_page(self, anchor: str, start: int = 0) -> int | None:
        """找锚文本所在页码（忽略空白差异）。命中多页时取离 start 最近的一页，
        同距取靠前；无命中返回 None。"""
        needle = _norm(anchor)[:24]
        if not needle:
            return None
        texts = self._page_texts()
        hits = [i for i in range(self.count) if needle in _norm(texts[i])]
        if not hits:
            return None
        return min(hits, key=lambda p: (abs(p - start), p))

    def close(self):
        try:
            self._pdf.close()
        except Exception:
            pass
