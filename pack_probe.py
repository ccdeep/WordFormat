"""打包探针：onefile exe 内运行 tkwebview2，结果写入文件供外部检查。"""
import sys
import time
import tkinter as tk

import pythoncom

RESULT = r"F:\Word格式整理工具\过程文件\pack_probe_result.txt"
out = open(RESULT, "w", buffering=1)

try:
    pythoncom.CoInitialize()
    from tkwebview2.tkwebview2 import WebView2, have_runtime

    print(f"runtime: {have_runtime()}", file=out)
    root = tk.Tk()
    root.geometry("800x600")
    wv = WebView2(root, 780, 580)
    wv.pack(fill="both", expand=True)
    wv.load_html('<html><head><meta charset="utf-8"></head><body>'
                 '<h1>打包测试</h1><p id="t">目标段落</p>'
                 + "<p>填充</p>" * 100 + "</body></html>")
    root.update()
    deadline = time.time() + 30
    while wv.core is None and time.time() < deadline:
        root.update()
        time.sleep(0.1)
    out.write(f"core: {wv.core is not None}\n")
    time.sleep(1)
    res = {}
    wv.evaluate_js('document.getElementById("t").scrollIntoView();'
                   'JSON.stringify({sy: window.scrollY})',
                   lambda r: res.update(v=r))
    t0 = time.time()
    while "v" not in res and time.time() - t0 < 15:
        root.update()
        time.sleep(0.05)
    out.write(f"js: {res.get('v')}\n")
    ok = bool(res.get("v")) and wv.core is not None
    out.write("PASS\n" if ok else "FAIL\n")
    root.destroy()
except Exception as e:
    import traceback
    out.write(f"EXC: {traceback.format_exc()}\nFAIL\n")
    sys.exit(1)
