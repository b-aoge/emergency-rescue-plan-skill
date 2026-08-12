# -*- coding: utf-8 -*-
"""用 Word COM 更新域，并读取每个 _TocN 书签的真实页码，写入 toc_pages.json。

用法：
  python compute_pages.py <path_to_docx>
依赖：Windows + 已安装 Microsoft Word + pywin32。若无参数，默认读取同目录的
施工现场应急救援预案_生成版.docx。
"""
import os, json, sys, time
import win32com.client as win32

base = os.path.dirname(os.path.abspath(__file__))
if len(sys.argv) > 1:
    out = os.path.abspath(sys.argv[1])
else:
    out = os.path.join(base, "施工现场应急救援预案_生成版.docx")
out = os.path.abspath(out)
if not os.path.exists(out):
    raise SystemExit("文档不存在: " + out)

word = win32.Dispatch("Word.Application")
word.Visible = False
word.ScreenUpdating = False
doc = word.Documents.Open(out)
try:
    doc.Fields.Update()
    time.sleep(0.5)
    pages = {}
    for i in range(15):          # 与 build_docx.py 的 TOC 条目数一致
        name = "_Toc%d" % i
        try:
            bk = doc.Bookmarks(name)
            pg = bk.Range.Information(3)   # wdActiveEndPageNumber
            pages[str(i)] = int(pg)
        except Exception:
            pages[str(i)] = ""
    doc.Save()
finally:
    doc.Close()
    word.Quit()

with open(os.path.join(base, "toc_pages.json"), "w", encoding="utf-8") as f:
    json.dump(pages, f, ensure_ascii=False, indent=2)
print("TOC pages:", pages)
