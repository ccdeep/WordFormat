"""Word 联动预览：调用本机 Word 打开原文档并定位到选中段落。

作为预览面板的补充——所见即 Word 100% 原样（公式/图片/表格全真）。
依赖 pywin32；装有 Word（或注册了 Word.Application 兼容接口的 WPS）时可用，
调不动时抛 RuntimeError 由 GUI 弹提示，不影响面板本身使用。
"""
import os


def open_in_word(docx_path: str, anchor_text: str = "") -> str:
    """在 Word 中打开（或复用已打开的）docx，定位并选中 anchor_text 所在段落。

    返回状态栏说明文字；无法调用 Word 时抛 RuntimeError。
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError("缺少 pywin32 依赖") from e

    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = True  # COM 新建的 Word 实例默认不可见

        target = os.path.normcase(os.path.abspath(docx_path))
        doc = None
        for d in app.Documents:
            try:
                if os.path.normcase(d.FullName) == target:
                    doc = d
                    break
            except Exception:
                continue
        if doc is None:
            doc = app.Documents.Open(FileName=os.path.abspath(docx_path),
                                     ReadOnly=True, AddToRecentFiles=False)

        doc.Activate()
        app.Activate()
        anchor = " ".join((anchor_text or "").split())[:80]
        if not anchor:
            return "已在 Word 中打开文档"
        rng = doc.Content
        find = rng.Find
        find.ClearFormatting()
        find.Text = anchor.replace("^", "^^")  # ^ 在 Word 查找里是转义符
        find.Forward = True
        find.Wrap = 0  # wdFindStop
        find.MatchCase = False
        find.MatchWildcards = False
        if find.Execute():
            rng.Select()
            return "已在 Word 中定位并选中该段（只读查看）"
        return "已在 Word 中打开文档（该段未能精确定位，可能是自动编号或空段）"
    finally:
        pythoncom.CoUninitialize()


def export_pdf(docx_path: str, pdf_path: str) -> str:
    """用隐藏 Word 实例把文档导出为 PDF（供"页面版式"预览渲染）。

    注意用 SaveAs2(FileFormat=17) 而不是 ExportAsFixedFormat——后者在部分
    机器上是病理性慢路径（50s+），SaveAs2 只需 2~4s。失败抛 RuntimeError。
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError("缺少 pywin32 依赖") from e

    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            doc = app.Documents.Open(FileName=os.path.abspath(docx_path),
                                     ReadOnly=True, AddToRecentFiles=False)
            try:
                doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=17)  # wdFormatPDF
            finally:
                doc.Close(False)
        finally:
            app.Quit()
        return pdf_path
    finally:
        pythoncom.CoUninitialize()
