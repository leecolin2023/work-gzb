#!/usr/bin/env python3
"""Normalize a valuation report into a numbered evidence index.

The index is the structural-independence layer: reviewers locate evidence by
semantic anchors against these IDs, never by chapter number or heading.

Formats:
  .docx  zero-dependency, reads word/document.xml with ElementTree
  .doc   legacy OLE, best-effort UTF-16LE/CP936 plaintext fallback (low confidence)
  .pdf   optional fitz / pdfplumber / pdftotext; scanned PDFs are flagged needs_ocr

Usage:
  python extract_report.py --input REPORT.docx --out OUTDIR [--key 07] [--md-truncate 240]

Outputs:
  evidence_{key}.json   full text, authoritative
  evidence_{key}.md     human-readable index (truncated text, IDs intact)
"""
import argparse
import datetime
import html
import json
import os
import re
import subprocess
import sys
import zipfile
from xml.etree import ElementTree as ET

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
HEAD_RE = re.compile(r"^\s*(?:[一二三四五六七八九十]+[、.]|\d+[、.]\s|第[一二三四五六七八九十]+[章节])")


def para_text(p):
    out = []
    for node in p.iter():
        if node.tag in (W + "t", "{http://schemas.openxmlformats.org/officeDocument/2006/math}t"):
            out.append(node.text or "")
        elif node.tag == W + "tab":
            out.append("\t")
        elif node.tag == W + "br":
            out.append(" ")
    return "".join(out).strip()


def extract_docx(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    if body is None:
        return [], "docx"
    items = []
    pno = tno = 0
    for child in body:
        if child.tag == W + "p":
            txt = para_text(child)
            if not txt:
                continue
            pno += 1
            items.append({"id": "P%04d" % pno, "type": "para", "text": txt})
        elif child.tag == W + "tbl":
            tno += 1
            rno = 0
            for tr in child.findall(W + "tr"):
                rno += 1
                cno = 0
                for tc in tr.findall(W + "tc"):
                    cno += 1
                    txt = " ".join(para_text(p) for p in tc.findall(W + "p")).strip()
                    if not txt:
                        continue
                    items.append({
                        "id": "T%02d-R%02d-C%02d" % (tno, rno, cno),
                        "type": "table_cell", "table": tno, "row": rno, "col": cno,
                        "text": txt,
                    })
    return items, "docx"


def extract_doc(path):
    with open(path, "rb") as f:
        data = f.read()
    best = ""
    for enc in ("utf-16-le", "gb18030", "cp936", "latin-1"):
        try:
            s = data.decode(enc, errors="ignore")
        except Exception:
            continue
        chunks = re.findall(r"[\u4e00-\u9fff0-9A-Za-z,，。、；：（）()%．\.\-/￥元\s]{6,}", s)
        cand = "\n".join(c.strip() for c in chunks if len(c.strip()) >= 6)
        score = len(re.findall(r"[\u4e00-\u9fff]", cand))
        if score > len(re.findall(r"[\u4e00-\u9fff]", best)):
            best = cand
    items = []
    for i, line in enumerate(best.splitlines(), 1):
        line = line.strip()
        if line:
            items.append({"id": "D%04d" % i, "type": "line", "text": line})
    return items, "doc"


def _pdf_pages(path):
    try:
        import fitz  # PyMuPDF
        with fitz.open(path) as doc:
            return [p.get_text() for p in doc]
    except ImportError:
        pass
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return [p.extract_text() or "" for p in pdf.pages]
    except ImportError:
        pass
    try:
        out = subprocess.run(["pdftotext", "-layout", path, "-"],
                             capture_output=True, timeout=180)
        if out.returncode == 0:
            return out.stdout.decode("utf-8", errors="ignore").split("\f")
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return None


def extract_pdf(path):
    pages = _pdf_pages(path)
    if pages is None:
        return [], "pdf", "NO_TEXT_LAYER_EXTRACTOR (install PyMuPDF or pdfplumber, or pdftotext)"
    items = []
    for pno, page in enumerate(pages, 1):
        for j, line in enumerate(str(page).splitlines(), 1):
            line = line.strip()
            if line:
                items.append({"id": "L%04d-%03d" % (pno, j), "type": "line",
                              "page": pno, "text": line})
    return items, "pdf", ""


def add_sections(items):
    cur = ""
    for it in items:
        if it["type"] == "para" and len(it["text"]) <= 40 and HEAD_RE.match(it["text"]):
            cur = it["text"]
        it["section_hint"] = cur
    return items


def build_md(meta, items, truncate):
    lines = [
        "# 证据索引 · %s" % os.path.basename(meta["source"]),
        "",
        "> 格式：%s ｜ 条目数：%d ｜ 生成时间：%s" % (meta["format"], meta["count"], meta["generated"]),
        "> 编号规则：P#### 段落 ｜ T##-R##-C## 表格单元 ｜ L####-### PDF 行 ｜ D#### 旧版 doc 行",
        "> 本文件文本为截断展示，**完整文本以同名 .json 为准**；复核引用只使用编号。",
    ]
    if meta.get("needs_ocr"):
        lines.append("> **需 OCR**：%s" % meta.get("warning", "无文本层"))
    if meta.get("low_confidence"):
        lines.append("> **低置信度**：%s" % meta.get("warning", ""))
    for gap in meta.get("extraction_gaps", []):
        lines.append("> **完整性待核**：%s；%s" % (gap["kind"], gap["action"]))
    lines += ["", "| 编号 | 类型 | 章节提示 | 文本 |", "|---|---|---|---|"]
    for it in items:
        txt = it["text"]
        if truncate and len(txt) > truncate:
            txt = txt[:truncate] + "…"
        txt = txt.replace("|", "\\|").replace("\n", " ")
        sec = it.get("section_hint", "").replace("|", "\\|")
        lines.append("| %s | %s | %s | %s |" % (it["id"], it["type"], sec, txt))
    return "\n".join(lines) + "\n"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--key", default=None, help="output suffix, e.g. 07")
    ap.add_argument("--md-truncate", type=int, default=240)
    args = ap.parse_args()

    src = os.path.abspath(args.input)
    if not os.path.isfile(src):
        sys.exit("input not found: " + src)
    ext = os.path.splitext(src)[1].lower()
    warning = ""
    if ext == ".docx":
        items, fmt = extract_docx(src)
    elif ext == ".doc":
        items, fmt = extract_doc(src)
        warning = "旧版 .doc 为二进制回退抽取，建议另存为 .docx 后重跑"
    elif ext == ".pdf":
        items, fmt, warning = extract_pdf(src)
    else:
        sys.exit("unsupported format: %s (use .docx/.doc/.pdf)" % ext)

    items = add_sections(items)
    extraction_gaps = []
    if fmt == "docx":
        with zipfile.ZipFile(src) as archive:
            xml = ET.fromstring(archive.read("word/document.xml"))
            for tag, label in ((W + "pict", "VML图片/公式"), (W + "drawing", "绘图/图片"),
                               (W + "object", "嵌入对象"), (W + "tbl", "表格")):
                count = len(list(xml.iter(tag)))
                if count and tag != W + "tbl":
                    extraction_gaps.append({"kind": label, "count": count, "action": "回原文检查是否包含公式或估值证据；不能因文字提取成功忽略"})
            if len(xml.findall(".//" + W + "tbl")) > len(xml.findall("./" + W + "body/" + W + "tbl")):
                extraction_gaps.append({"kind": "嵌套表格", "action": "当前索引不保证完整，回原文核验"})
    low_conf = fmt == "doc"
    needs_ocr = fmt == "pdf" and not items
    key = args.key or (re.findall(r"\d{2}", os.path.basename(src)) or ["report"])[0]
    meta = {
        "source": src,
        "format": fmt,
        "generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "count": len(items),
        "needs_ocr": needs_ocr,
        "low_confidence": low_conf,
        "warning": warning,
        "extraction_gaps": extraction_gaps,
        "completeness_verified": False,
    }
    os.makedirs(args.out, exist_ok=True)
    jpath = os.path.join(args.out, "evidence_%s.json" % key)
    mpath = os.path.join(args.out, "evidence_%s.md" % key)
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump({**meta, "items": items}, f, ensure_ascii=False, indent=1)
    with open(mpath, "w", encoding="utf-8") as f:
        f.write(build_md(meta, items, args.md_truncate))
    print("format=%s items=%d needs_ocr=%s low_confidence=%s" % (fmt, len(items), needs_ocr, low_conf))
    print("json:", jpath)
    print("md  :", mpath)
    if warning:
        print("warning:", warning)
    if extraction_gaps:
        print("completeness pending:", ", ".join(g["kind"] for g in extraction_gaps))
    if needs_ocr:
        print("ACTION: scanned PDF has no text layer; mark 待定位 / 需 OCR, do not infer absence from failed extraction")


if __name__ == "__main__":
    main()
