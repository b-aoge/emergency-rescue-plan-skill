# -*- coding: utf-8 -*-
"""导出 docx 为 PDF，逐页抽取文本，核对目录 15 条标题的真实页号与 PDF 实际位置是否对应。

用法：
  python verify_toc.py <path_to_docx> [pdf_path]
依赖：win32com (Word) 导出 PDF + PyMuPDF(fitz) 读文本。无参数则默认同目录生成版 docx。
结果写入与脚本同目录的 verify_toc_out.txt。
"""
import os, sys
base = os.path.dirname(os.path.abspath(__file__))
if len(sys.argv) > 1:
    out = os.path.abspath(sys.argv[1])
else:
    out = os.path.join(base, "施工现场应急救援预案_生成版.docx")
pdf = os.path.join(base, "verify_tmp.pdf") if len(sys.argv) < 3 else os.path.abspath(sys.argv[2])
log = os.path.join(base, "verify_toc_out.txt")

lines = []
try:
    import win32com.client as win32
    word = win32.Dispatch("Word.Application"); word.Visible = False
    doc = word.Documents.Open(os.path.abspath(out))
    doc.ExportAsFixedFormat(pdf, 17)   # wdExportFormatPDF
    doc.Close(); word.Quit()
    lines.append("PDF exported: " + pdf)
except Exception as e:
    lines.append("PDF export error: " + repr(e))

try:
    import fitz
    d = fitz.open(pdf)
    pages_text = [d[i].get_text() for i in range(len(d))]
    titles = [
        "一、编制依据","二、基本原则与方针","三、工程项目基本情况",
        "四、可能发生事故的确定及主要危险源","五、应急机构的组成、责任和分工",
        "六、报警信号与通讯","七、应急响应与救援","八、有关规定和要求","九、附件",
        "附件1 常见事故急救常识","附件2 火灾逃生自救知识","附件3 灭火器的使用常识",
        "附件4 各类事故的预防措施","附件5 施工总平面布置图","附件6 应急救援路线图",
    ]
    for t in titles:
        found = []
        for pi, tx in enumerate(pages_text, 1):
            if t in tx or t.replace(" ", "　") in tx or t.replace(" ", "  ") in tx:
                found.append(pi)
        real = [p for p in found if p != 2]   # 排除第2页（目录页本身含这些标题）
        lines.append(t + " -> real heading page " + (str(real[0]) if real else "NOT FOUND") + " (all:"+str(found)+")")
    lines.append("total PDF pages: " + str(len(d)))
except Exception as e:
    lines.append("fitz error: " + repr(e))

with open(log, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print("done ->", log)
