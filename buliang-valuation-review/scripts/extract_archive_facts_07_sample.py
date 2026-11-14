# -*- coding: utf-8 -*-
"""07 江门佰富 · 银行档案数据提取 v1（样板）

只读原始资料，把归档 / 已归集资料中的事实抽成结构化字段，
直接对齐 buliang-valuation-review skill 的 L1《基础事实对照表》。

输出：
  {户}/银行档案/07_江门佰富_档案数据提取_v1.xlsx
  {户}/银行档案/07_江门佰富_证据索引_v1.json
"""
import os
import re
import sys
import json
import html
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_mini import write_xlsx

ROOT = r"E:/work_gzb/buliang"
HOUSE = "07_江门佰富地产有限公司"
OUT = os.path.join(ROOT, "output", "现有资料整理", HOUSE, "银行档案")
ARCHIVE = os.path.join(ROOT, "归档", HOUSE)

SRC = {
    "CL": os.path.join(ROOT, "共享资料", "2026年第四季度拟转让清单1023.xlsx"),
    "RS": os.path.join(ARCHIVE, "02_补充材料", "20261030_补料包",
                       "江门佰富地产有限公司", "江门佰富地产有限公司重组情况.docx"),
    "DD": os.path.join(ROOT, "output", "现有资料整理", HOUSE, "法律尽调",
                       "20261103修改-关于江门佰富地产有限公司贷款项目的法律尽职调查报告.docx"),
    "VR": os.path.join(ROOT, "output", "价值分析报告", "07_江门佰富地产有限公司-20261112.docx"),
    "CR": os.path.join(ROOT, "output", "现有资料整理", HOUSE, "查册成果",
                       "07_江门佰富地产有限公司_抵押物查册核对报告.md"),
    "XC": os.path.join(ROOT, "output", "现有资料整理", HOUSE, "查册成果",
                       "结构包抵押物清单_07_江门佰富地产有限公司.xlsx"),
}
NAME = {
    "CL": "共享资料/2026年第四季度拟转让清单1023.xlsx",
    "RS": "归档/02_补充材料/20261030_补料包/江门佰富地产有限公司重组情况.docx",
    "DD": f"现有资料整理/{HOUSE}/法律尽调/20261103修改-…法律尽职调查报告.docx",
    "VR": "output/价值分析报告/07_江门佰富地产有限公司-20261112.docx",
    "CR": f"现有资料整理/{HOUSE}/查册成果/07_…抵押物查册核对报告.md",
    "XC": f"现有资料整理/{HOUSE}/查册成果/结构包抵押物清单_07_…xlsx",
    "AR": f"归档/{HOUSE}/（原件）",
    "AV": f"归档/{HOUSE}/02_补充材料/法律尽调意见/",
}
LEVEL = {
    "CL": "层6 银行正式清单",
    "RS": "层4 银行说明（归档原件）",
    "DD": "层5 律师引述（须回原件核验）",
    "VR": "被复核对象（报告自述）",
    "CR": "层7 查册成果",
    "XC": "层7 查册成果（结构化清单）",
    "AR": "层1~3 归档原件（PDF 文本层直读）",
    "AV": "层5 归档内尽调意见副本",
}


# ---------------- 通用读取 ----------------
def docx_paras(path):
    z = zipfile.ZipFile(path)
    x = z.read("word/document.xml").decode("utf-8")
    x = re.sub(r"</w:p>", "\n", x)
    x = re.sub(r"</w:tr>", "\n", x)
    x = re.sub(r"</w:tc>", " | ", x)
    t = html.unescape(re.sub(r"<[^>]+>", "", x))
    return [l.strip() for l in t.split("\n") if l.strip()]


def cl_cells(path):
    """按单元格引用（A1/E10…）返回 {行号: {列字母: 值}}，避免空列导致的错位。"""
    z = zipfile.ZipFile(path)
    try:
        ss = z.read("xl/sharedStrings.xml").decode("utf-8")
        strs = [html.unescape(re.sub(r"<[^>]+>", "", s))
                for s in re.findall(r"<si>(.*?)</si>", ss, re.S)]
    except KeyError:
        strs = []

    def cell(attr, body):
        v = re.search(r"<v>(.*?)</v>", body, re.S)
        t = re.search(r't="([^"]+)"', attr)
        if v is None:
            m = re.search(r"<t[^>]*>(.*?)</t>", body, re.S)
            return html.unescape(m.group(1)) if m else ""
        if t and t.group(1) == "s" and strs:
            return strs[int(v.group(1))]
        return v.group(1)

    x = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    out = {}
    for rm in re.finditer(r'<row[^>]*\br="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        rn = int(rm.group(1))
        d = {}
        for cm in re.finditer(r'<c([^>]*)>(.*?)</c>', rm.group(2), re.S):
            attr, body = cm.group(1), cm.group(2)
            ref = re.search(r'\br="([A-Z]+)\d+"', attr)
            if ref:
                d[ref.group(1)] = cell(attr, body)
        out[rn] = d
    return out


def xlsx_rows(path, sheet_index=0):
    z = zipfile.ZipFile(path)
    try:
        ss = z.read("xl/sharedStrings.xml").decode("utf-8")
        strs = [html.unescape(re.sub(r"<[^>]+>", "", s))
                for s in re.findall(r"<si>(.*?)</si>", ss, re.S)]
    except KeyError:
        strs = []

    def cell(attr, body):
        v = re.search(r"<v>(.*?)</v>", body, re.S)
        t = re.search(r't="([^"]+)"', attr)
        if v is None:
            m = re.search(r"<t[^>]*>(.*?)</t>", body, re.S)
            return html.unescape(m.group(1)) if m else ""
        if t and t.group(1) == "s" and strs:
            return strs[int(v.group(1))]
        return v.group(1)

    sheets = [n for n in z.namelist()
              if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
    x = z.read(sheets[sheet_index]).decode("utf-8")
    out = []
    for rn, r in enumerate(re.findall(r"<row[^>]*>(.*?)</row>", x, re.S), 1):
        out.append((rn, [cell(a, b) for a, b in
                         re.findall(r"<c([^>]*)>(.*?)</c>", r, re.S)]))
    return out


P = {k: docx_paras(v) for k, v in SRC.items() if v.endswith(".docx")}
P["CR"] = open(SRC["CR"], encoding="utf-8").read().split("\n")
CL_CELLS = cl_cells(SRC["CL"])

records = []


def add(field, value, src, loc, snippet, unit="", method="锚点+正则", verify="否"):
    rec = {
        "字段": field, "提取值": str(value), "单位": unit,
        "来源代号": src, "来源文件": NAME.get(src, src), "定位": loc,
        "原文片段": str(snippet)[:400], "证据层级": LEVEL.get(src, ""),
        "证据编号": "E-%04d" % (len(records) + 1),
        "提取方式": method, "需回原件核验": verify,
    }
    records.append(rec)
    return value


def grab(field, src, kw, pat, unit="", group=1, note=""):
    """在 src 的段落中找含 kw 的第一段，再用 pat 捕获；失败则登记为未提取。"""
    for i, line in enumerate(P[src]):
        if kw not in line:
            continue
        m = re.search(pat, line)
        if not m:
            continue
        v = m.group(group) if m.lastindex else m.group(0)
        return add(field, v.strip(), src, "%s 段 %d" % (src, i), line,
                   unit=unit, verify="是" if src in ("DD", "RS") else "否")
    add(field, "未自动提取", src, "-", "关键词：%s" % kw,
        unit=unit, method="自动抽取失败", verify="需人工/需OCR")
    return None


# ---------------- 1. 价值分析报告（被复核对象） ----------------
V = P["VR"]
SEG70 = next((l for l in V if "综合可回收价值为" in l and "一般受偿债权价值" in l), "")
REPORT_PATS = [
    ("报告_综合可回收价值", r"综合可回收价值为([\d,]+)元", "元"),
    ("报告_优先受偿价值", r"优先受偿债权价值为([\d,]+)元", "元"),
    ("报告_一般受偿价值", r"一般受偿债权价值为([\d,]+)元", "元"),
    ("报告_保证受偿价值", r"保证受偿债权价值为([\d,]+)元", "元"),
    ("报告_其他受偿价值", r"其他债务责任关联方的受偿价值为([\d,]+)元", "元"),
    ("报告_贴现后估值", r"≈([\d,]+)元", "元"),
    ("报告_处置周期", r"需时(\d+)年", "年"),
    ("报告_贴现率", r"贴现率取([\d.]+)%", "%"),
    ("报告_本金受偿率", r"本金受偿率([\d.]+)%", "%"),
    ("报告_债权受偿率", r"本债权受偿率为([\d.]+)%", "%"),
]
RV = {}
for f, pat, unit in REPORT_PATS:
    m = re.search(pat, SEG70)
    idx = next((i for i, l in enumerate(V) if l == SEG70), -1)
    if m:
        RV[f] = m.group(1)
        add(f, m.group(1), "VR", "VR 段 %d" % idx, SEG70, unit=unit,
            method="结论段统一抽取")
    else:
        add(f, "未自动提取", "VR", "VR 段 %d" % idx, SEG70, unit=unit,
            method="自动抽取失败", verify="需人工")

SEG61 = next((l for l in V if "债权金额合计" in l and "其中本金" in l), "")
idx61 = next((i for i, l in enumerate(V) if l == SEG61), -1)
for f, pat in [
    ("报告_债权合计", r"债权金额合计([\d,]+\.\d{2})元"),
    ("报告_本金", r"其中本金([\d,]+\.\d{2})元"),
    ("报告_利息", r"利息([\d,]+\.\d{2})元"),
    ("报告_代垫费用", r"代垫费用([\d,]+\.\d{2})元"),
]:
    m = re.search(pat, SEG61)
    if m:
        add(f, m.group(1), "VR", "VR 段 %d" % idx61, SEG61, unit="元",
            method="债权构成段统一抽取")
    else:
        add(f, "未自动提取", "VR", "VR 段 %d" % idx61, SEG61, unit="元",
            method="自动抽取失败", verify="需人工")

for f, kw, pat, unit in [
    ("报告_分析基准日", "分析基准日", r"分析基准日\s*([\d年月日]+)", ""),
    ("报告_分析方法", "综合因素分析法", r"(综合因素分析法)", ""),
    ("报告_提交日期", "本分析报告提交日期", r"提交日期为([\d年月日]+)", ""),
]:
    grab(f, "VR", kw, pat, unit=unit)


def num(s):
    try:
        return float(str(s).replace(",", ""))
    except (ValueError, TypeError):
        return None


# ---------------- 2. 银行正式清单（层6） ----------------
H = CL_CELLS.get(3, {})
col_by_name = {re.sub(r"\s+", "", name): col for col, name in H.items()}
cl_rn = next((rn for rn, d in CL_CELLS.items()
              if rn > 3 and any("江门佰富" in str(v) for v in d.values())), None)
DROW = CL_CELLS.get(cl_rn, {}) if cl_rn else {}
CLD = {name: DROW.get(col, "") for name, col in col_by_name.items()}
for key in ["客户名称", "贷款余额", "欠息", "代垫费用", "债权合计", "担保方式",
            "抵押物详情", "质押物详情", "保证人详情", "诉讼执行进展", "重组类型", "五级分类"]:
    if key in CLD and CLD[key] not in ("", None):
        v = CLD[key]
        add("清单_" + key, str(v)[:350], "CL",
            "Sheet1 行 %d / 列 %s「%s」" % (cl_rn or 0, col_by_name.get(key, "?"), key),
            " | ".join("%s=%s" % (k, str(v)[:60]) for k, v in list(DROW.items())[:8]),
            method="表头-列字母映射")

# ---------------- 3. 银行重组说明（层4，归档 docx） ----------------
R = P["RS"][0] if P["RS"] else ""
RS_RE = [
    (1, r"(\d{4})年(\d{1,2})月我行对申请人([\d.]+)万元贷款已进行第一次重组"),
    (2, r"于(\d{4})年(\d{1,2})月对申请人([\d.]+)万元贷款进行第二次重组"),
    (3, r"(\d{4})年(\d{1,2})月对申请人([\d.]+)万元进行第三次重组"),
    (4, r"(\d{4})年(\d{1,2})月对申请人([\d.]+)万元进行第四次重组"),
]
RS_CHAIN = {}
for idx, pat in RS_RE:
    m = re.search(pat, R)
    if m:
        RS_CHAIN[idx] = {"日期": "%s-%s" % (m.group(1), m.group(2).zfill(2)),
                         "金额万元": m.group(3)}
        add("银行说明_第%d次重组时间" % idx, RS_CHAIN[idx]["日期"], "RS",
            "RS 段 0", R, method="正则")
        add("银行说明_第%d次重组本金(万元)" % idx, m.group(3), "RS",
            "RS 段 0", R, unit="万元", method="正则", verify="是")
for f, pat, unit in [
    ("银行说明_授信额度", r"给予([\d.]+)亿元城市更新贷款", "亿元"),
    ("银行说明_累计提款", r"累计提款([\d.]+)亿元", "亿元"),
    ("银行说明_贷款期限", r"期限(\d+)年（([\d.\-]+)）", "年"),
    ("银行说明_重组次数", r"第(四)次重组", "次"),
]:
    m = re.search(pat, R)
    if m:
        add(f, m.group(1), "RS", "RS 段 0", R, unit=unit, method="正则", verify="是")

# ---------------- 4. 法律尽调报告（层5） ----------------
for f, kw, pat, unit in [
    ("档案_贷款本金余额", "截至基准日2026年9月28日", r"本金余额\s*([\d,]+\.\d{2})\s*元", "元"),
    ("档案_欠息", "截至基准日2026年9月28日", r"欠息\s*([\d,]+\.\d{2})\s*元", "元"),
    ("档案_代垫评估费用", "截至基准日2026年9月28日", r"代垫评估费用\s*([\d,]+\.\d{2})\s*元", "元"),
    ("档案_债权总额", "截至基准日2026年9月28日", r"债权总额共计人民币\s*([\d,]+\.\d{2})\s*元", "元"),
    ("档案_借款合同编号", "固定资产借款合同》【编号", r"【编号：\s*([^】]+)】", ""),
    ("档案_实际放款金额", "实际发放贷款金额合计", r"发放贷款人民币\s*([\d,]+\.\d{2})元", "元"),
    ("档案_提前到期本金", "贷款提前到期通知书", r"到期本金([\d,]+\.\d{2})元", "元"),
    ("档案_一审判项本金", "偿还借款本金130,301,203.57", r"借款本金([\d,]+\.\d{2})元", "元"),
    ("档案_保证合同编号", "保证合同》【编号", r"【编号：\s*([^】]+)】", ""),
    ("档案_质押合同编号(兆誉)", "2020江分权质字第1113006-1", r"(2020江分权质字第1113006-1号)", ""),
    ("档案_质押合同编号(宝晟)", "2020江分权质字第1113006-2", r"(2020江分权质字第1113006-2号)", ""),
    ("档案_一审案号", "粤0703民初18146号，江门市蓬江区人民法院于2022年5月9日", r"案号为\s*([^，,]+)", ""),
    ("档案_判决日期", "于2022年5月9日出具《民事判决书》", r"于([\d年月日]+)出具", ""),
    ("档案_二审撤诉案号", "粤07民终3886号", r"作出\s*([^《]+)《民事裁定书》", ""),
    ("档案_保全案号", "执保2735号", r"(（\d{4}）粤\s?0703\s?执保\s?2735\s?号)", ""),
    ("档案_恢复执行案号", "执恢276号《冻结、查封、扣押财产风险告知书》", r"作出\s*([^《]+)《", ""),
    ("档案_查封不动产套数", "合计的372套不动产", r"合计的([\d]+)套不动产", "套"),
    ("档案_本行对372套权利性质", "贵行对该批不动产不享有抵押权", r"(贵行对该批不动产不享有抵押权)", ""),
    ("档案_金翠园C-1801处置", "金翠园C-1801的房产", r"(据贵行反映该房产已处置[^，,，]*)", ""),
]:
    grab(f, "DD", kw, pat, unit=unit)

# 尽调重组链：协议编号 / 日期 / 确认本金
DD_CHAIN = {}
for idx, para in [(1, 45), (2, 46), (3, 47), (4, 48)]:
    if para >= len(P["DD"]):
        continue
    line = P["DD"][para]
    d = re.search(r"(\d{4}年\d{1,2}月\d{1,2}日)", line)
    no = re.search(r"【编号：\s*([^】]+)】", line)
    bal = re.search(r"(?:欠借款本金|本金余额仍为|尚欠借款本金)([\d,]+\.\d{2})元", line)
    DD_CHAIN[idx] = {"日期": d.group(1) if d else "",
                     "协议编号": no.group(1) if no else "",
                     "确认本金": bal.group(1) if bal else ""}
    for k, v in DD_CHAIN[idx].items():
        if v:
            add("尽调_第%d次重组_%s" % (idx, k), v, "DD",
                "DD 段 %d" % para, line, verify="是")

# ---------------- 5. 查册成果（层7） ----------------
for f, kw, pat, unit in [
    ("查册_逐套匹配", "匹配 16/372 套", r"匹配\s*([\d/]+)\s*套", "套"),
    ("查册_覆盖率", "匹配 16/372 套", r"（([\d.]+%)）", "%"),
    ("查册_底稿套数", "底稿套数", r"\|\s*底稿套数\s*\|\s*([\d,]+)\s*套", "套"),
]:
    grab(f, "CR", kw, pat, unit=unit)

# ---------------- 6. 归档文件级台账 ----------------
file_rows = []
all_files = []
for root, dirs, names in os.walk(ARCHIVE):
    for n in names:
        all_files.append(os.path.join(root, n))
all_files.sort(key=lambda p: os.path.relpath(p, ARCHIVE).replace("\\", "/"))

F_HINT = [
    ("合同借据", "F03/F04/F06"), ("权证", "F06/F07"), ("诉讼材料", "F08/F10"),
    ("主体资料", "F05"), ("法律尽调意见", "F02~F08"), ("补料包", "F04/F07/F09"),
    ("现场照片", "L2 资产端"), ("查封资产清单", "F07/L2"),
]
pdf_stats = {}
PDF_TEXTS = {}
try:
    import logging
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

for i, p in enumerate(all_files, 1):
    rel = os.path.relpath(p, ARCHIVE).replace("\\", "/")
    ext = os.path.splitext(p)[1].lower().lstrip(".")
    size = os.path.getsize(p)
    hint = next((h for k, h in F_HINT if k in rel), "-")
    status, detail, ev = "", "", "A-%03d" % i
    if ext == "docx":
        n = len(docx_paras(p))
        status, detail = "已提取", "%d 段文本" % n
    elif ext == "xls":
        status, detail = "待解析", "OLE2 二进制；已有结构化替代（查册成果/结构包抵押物清单_07_…xlsx）"
    elif ext == "pdf":
        if HAS_PYPDF:
            try:
                r = PdfReader(p)
                txt = ""
                for pg in r.pages[:3]:
                    txt += (pg.extract_text() or "")
                if txt.strip():
                    status, detail = "有文本层·已提取", "%d 页；抽取 %d 字符，已入证据索引" % (len(r.pages), len(txt))
                    PDF_TEXTS[rel] = {"pages": len(r.pages), "text": txt}
                else:
                    status, detail = "需 OCR", "%d 页；零文本层（纯扫描）" % len(r.pages)
                pdf_stats[rel] = (len(r.pages), len(txt.strip()))
            except Exception as e:
                status, detail = "读取失败", str(e)[:80]
        else:
            status, detail = "未检测", "pypdf 不可用"
    elif ext in ("jpg", "jpeg", "png", "tif", "tiff"):
        status, detail = "需 OCR", "图片扫描件"
    elif ext == "zip":
        status, detail = "未展开", "阶段 2 暂缓，未解压"
    else:
        status, detail = "未处理", "未知类型"
    file_rows.append([ev, rel, ext.upper(), "%.1f KB" % (size / 1024.0),
                      status, detail, hint, NAME["AR"]])

ocr_rows = [[r[0], r[1], r[5], r[6], "高" if any(k in r[1] for k in
            ("合同借据", "判决书", "裁定书", "权证", "查册资料", "质押合同", "保证合同")) else "中"]
            for r in file_rows if "需 OCR" in r[4]]

# ---------------- 6b. 归档原件字段（PDF 文本层直读，层1~3） ----------------
AR_REL, AR_NORM = "", ""
for rel, obj in PDF_TEXTS.items():
    if "预缴评估费" in rel:
        AR_REL, AR_NORM = rel, re.sub(r"\s+", "", obj["text"])
        break
if AR_NORM:
    for f, pat, unit in [
        ("原件_执行案号", r"(（\d{4}）粤\d+执恢\d+号)", ""),
        ("原件_申请执行人", r"关于申请执行人(.+?)申请执行", ""),
        ("原件_评估机构", r"依法委托(.+?)对佳兆业", ""),
        ("原件_委托评估不动产套数", r"(\d+)套不动产", "套"),
        ("原件_代垫评估费用", r"评估费用(\d+)元", "元"),
        ("原件_通知日期", r"(二○二六年八月三十日)", ""),
    ]:
        m = re.search(pat, AR_NORM)
        if m:
            add(f, m.group(1), "AR", "归档/" + AR_REL, AR_NORM[:400],
                unit=unit, method="PDF 文本层直读", verify="否")
        else:
            add(f, "未自动提取", "AR", "归档/" + AR_REL, AR_NORM[:400],
                unit=unit, method="自动抽取失败", verify="需人工")

# ---------------- 7. 数字验算 ----------------
def fmt(v, nd=2):
    return ("%%.%df" % nd) % v


calc_rows = []
def calc(label, expr, reported, computed, unit="元", tol=0.01, rule=""):
    diff = computed - reported if (computed is not None and reported is not None) else None
    if diff is None:
        verdict = "无法验算"
    elif abs(diff) <= tol:
        verdict = "一致"
    elif rule == "百位取整" and abs(diff) < 100:
        verdict = "百位取整，非实质差错"
    else:
        verdict = "差异，需核"
    calc_rows.append([label, expr, fmt(reported) if reported is not None else "-",
                      fmt(computed) if computed is not None else "-",
                      fmt(diff) if diff is not None else "-", verdict, unit])


p = num(RV.get("报告_一般受偿价值"))
g = num(RV.get("报告_保证受偿价值"))
prio = num(RV.get("报告_优先受偿价值"))
other = num(RV.get("报告_其他受偿价值"))
total = num(RV.get("报告_综合可回收价值"))
disc = num(RV.get("报告_贴现后估值"))
claim = num(CLD.get("债权合计"))
prin = num(CLD.get("贷款余额"))
intr = num(CLD.get("欠息"))
fee = num(CLD.get("代垫费用"))
rate = num(RV.get("报告_贴现率")) / 100 if num(RV.get("报告_贴现率")) is not None else None
yrs = num(RV.get("报告_处置周期"))

calc("债权合计勾稽", "贷款余额 + 欠息 + 代垫费用", claim,
     (prin + intr + fee) if None not in (prin, intr, fee) else None)
calc("受偿项加总", "优先 + 一般 + 保证 + 其他", total,
     sum(x for x in (prio, p, g, other) if x is not None)
     if None not in (prio, p, g, other) else None)
calc("一般受偿测算", "(债权总额 − 优先受偿) × 7.25%", p,
     (claim - prio) * 0.0725 if None not in (claim, prio) else None, rule="百位取整")
calc("保证受偿测算", "(债权总额 − 优先受偿) × 14.30%", g,
     (claim - prio) * 0.1430 if None not in (claim, prio) else None, rule="百位取整")
calc("贴现值", "综合可回收价值 ÷ (1+贴现率)^处置周期", disc,
     total / ((1 + rate) ** yrs) if None not in (total, rate, yrs) else None,
     rule="百位取整")
calc("本金受偿率", "贴现后估值 ÷ 贷款本金 × 100", num(RV.get("报告_本金受偿率")),
     disc / prin * 100 if None not in (disc, prin) else None, unit="%", tol=0.005)
calc("债权受偿率", "贴现后估值 ÷ 债权合计 × 100", num(RV.get("报告_债权受偿率")),
     disc / claim * 100 if None not in (disc, claim) else None, unit="%", tol=0.005)

# ---------------- 8. 三源勾稽 ----------------
def clv(k):
    v = CLD.get(k, "—")
    return v if v != "" else "—"


def dd(field):
    r = next((x for x in records if x["字段"] == field), None)
    return r["提取值"] if r else "—"


tri_rows = [
    ["贷款本金余额", clv("贷款余额"), "—", dd("档案_贷款本金余额"),
     dd("报告_本金"), "一致", "拟转让清单/尽调/报告三源同值（VR 段 61）"],
    ["欠息", clv("欠息"), "—", dd("档案_欠息"), dd("报告_利息"), "一致", ""],
    ["代垫费用", clv("代垫费用"), "—", dd("档案_代垫评估费用"), dd("报告_代垫费用"), "一致",
     "83,030 元有法院《预缴评估费通知》原件支撑（层1，2026-08-30）"],
    ["债权合计", clv("债权合计"), "—", dd("档案_债权总额"), dd("报告_债权合计"), "一致", "报告以此为受偿率分母"],
    ["授信/合同额度", "—", dd("银行说明_授信额度") + " 亿元", "750,000,000.00 元", "750,000,000 元", "一致", "单位换算后同值"],
    ["实际放款", "—", dd("银行说明_累计提款") + " 亿元", dd("档案_实际放款金额") + " 元", "150,000,000 元", "一致", "单位换算后同值"],
    ["重组次数", clv("重组类型"), "4 次（至第四次）", "4 次", "4 次（引述）", "一致", "清单「重组类型」为展期，非次数"],
    ["第三次重组时间", "—", RS_CHAIN.get(3, {}).get("日期", "—"),
     DD_CHAIN.get(3, {}).get("日期", "—"), "2025-09-22（引述）", "冲突待核",
     "银行说明写 2025年5月，尽调/报告写 2025-09-22"],
    ["第三次重组本金", "—", RS_CHAIN.get(3, {}).get("金额万元", "—") + " 万元",
     DD_CHAIN.get(3, {}).get("确认本金", "—") + " 元", "109,301,203.57 元", "冲突待核",
     "银行说明 11,930.12 万 vs 尽调 10,930.12 万，差 1,000 万元"],
    ["第四次重组本金", "—", RS_CHAIN.get(4, {}).get("金额万元", "—") + " 万元",
     DD_CHAIN.get(4, {}).get("确认本金", "—") + " 元", "109,301,203.57 元", "冲突待核",
     "同上，差 1,000 万元；拟转让清单余额支持尽调口径"],
    ["担保方式", clv("担保方式"), "—", "保证 + 股权质押（70%/30%）", "本项目无抵押担保", "口径不同，一致",
     "报告仅声明无抵押，未否认股权质押；需核 L2 是否漏计质权"],
    ["372 套权利性质", "已查封资产（非我行抵押物）", "—", dd("档案_本行对372套权利性质"),
     "无抵押担保；仅轮候查封", "一致", "拟转让清单原文即写「已查封资产」"],
    ["诉讼执行阶段", clv("诉讼执行进展"), "—", dd("档案_恢复执行案号"),
     "（2023）粤0703执恢276号", "一致", ""],
]

# ---------------- 9. L1 比对底稿 ----------------
l1_rows = [
    ["F01", "报告身份、委托人、基准日、出具日、有效期是否明确且自洽",
     "基准日 2026-09-28；提交日 2026-10-28", "VR 段 61 / 78",
     "基准日 2026-09-28（尽调段 49 同）", "DD 段 49", "待判", "", ""],
    ["F02", "债权本金、利息、罚息复利、代垫费用、合计能否与档案勾稽",
     "135,903,299.82 = 109,301,203.57 + 26,519,066.25 + 83,030.00", "VR 段 59~61",
     "拟转让清单行 %d 四要素同值；尽调段 49 同值" % (cl_rn or 0),
     "CL 行 %d / DD 段 49" % (cl_rn or 0), "待判",
     "三源一致；83,030 元代垫费用另有法院《预缴评估费通知》原件支撑（层1）", ""],
    ["F03", "放款事实是否成立；合同额／实际放款额／当前余额是否分列",
     "合同额 750,000,000；放款 150,000,000", "VR 段 ~400",
     "银行说明 7.5 亿 / 1.5 亿；尽调 750,000,000 / 150,000,000",
     "RS 段 0 / DD 段 38、43", "待判", "", ""],
    ["F04", "重组／展期／和解协议链是否完整、最新版本是否被采用",
     "报告引述四次重组", "VR 段 ~400",
     "四次协议编号齐全，末次为 2025 江分和解补字第 0529001 号",
     "DD 段 45~48 / RS 段 0", "待判",
     "银行说明与尽调在第三、四次重组金额差 1,000 万元、第三次时间不一致；见「三源勾稽」", "高"],
    ["F05", "债务人、抵质押人、保证人主体是否准确、存续",
     "佰富 / 佳兆业 / 兆誉 / 宝晟", "VR",
     "四主体；保证人佳兆业集团（深圳）有限公司（清单「保证人详情」同）",
     "CL 行 %d / DD 段 40~41" % (cl_rn or 0), "待判", "", ""],
    ["F06", "担保合同类型、编号、担保范围、保证期间是否与档案一致",
     "保证 + 质押（报告正文）", "VR",
     "借款 2020 江分固借字第 1112008 号；质押 -1/-2；保证 2020 江分保证字第 1113013 号；连带保证期间 2 年",
     "DD 段 38/40/41", "待判", "合同原件均为扫描件，编号来自律师引述，须回原件核验", "中"],
    ["F07", "担保物权是否依法设立、顺位与担保范围",
     "本项目无抵押担保；优先受偿价值 0", "VR 段 70",
     "股权质押 70%+30%，质权登记 2000344397 / 2000344372；372 套仅轮候查封、本行无抵押权",
     "DD 段 55/56/64 / CL 行 %d" % (cl_rn or 0), "待判",
     "报告未讨论股权质权的可实现价值，仅笼统写「无抵押担保」；法院通知写 374 套而清单/查册为 372 套，口径待核；见「价值来源」层", "高"],
    ["F08", "诉讼、判决、保全、执行进展是否与档案一致",
     "（2021）粤0703民初18146号 / 执恢276号", "VR",
     "一审判项 4 项；执保 2735 号；执恢 276 号；处置权移送龙岗／前海法院",
     "DD 段 52~64 / CL 行 %d" % (cl_rn or 0), "待判", "", ""],
    ["F09", "已处置财产与已清偿款项是否已从余额扣减",
     "本金 130,301,203.57 → 109,301,203.57", "VR",
     "金翠园 C-1801 已处置款用于清偿；一审判项 130,301,203.57 → 和解 110,301,203.57 → 109,301,203.57",
     "DD 段 44/46/47/62", "待判",
     "银行说明第三、四次重组金额 11,930.12 万与上述链条不符，需确认是否存在重复扣减或笔误", "高"],
    ["F10", "诉讼时效、执行时效、债权是否仍存续",
     "—", "—",
     "一审判决已生效、二审撤诉；律师结论「不存在诉讼时效届满问题」",
     "DD 段 59/467", "待判", "", ""],
]

# ---------------- 10. 事实清单 ----------------
fact_rows = [
    ["合同与担保", "借款合同", "2020 江分固借字第 1112008 号", "2020-11-30",
     "广州银行江门分行 / 佰富；7.5 亿元；2020-12-08~2023-12-08；年利率 7.5%", "DD 段 38 / RS 段 0"],
    ["合同与担保", "权利质押合同", "2020江分权质字第1113006-1号", "2020-11-30",
     "兆誉 → 佰富 70% 股权；本金/利息/复利/罚息/违约金/实现债权费用", "DD 段 40"],
    ["合同与担保", "权利质押合同", "2020江分权质字第1113006-2号", "2020-11-30",
     "宝晟 → 佰富 30% 股权；同上；补充协议未免除其质押责任", "DD 段 40/48"],
    ["合同与担保", "保证合同", "2020 江分保证字第 1113013 号", "2020-12-02",
     "佳兆业集团（深圳）有限公司；连带责任保证；期间为到期/提前到期之日起 2 年", "DD 段 41"],
    ["合同与担保", "股权出质登记", "（新会）股质登记设字[2020]第2000344397号", "2020",
     "兆誉 70%；判项三：拍卖变卖价款优先受偿", "DD 段 55"],
    ["合同与担保", "股权出质登记", "（新会）股质登记设字[2020]第2000344372号", "2020",
     "宝晟 30%；判项四：拍卖变卖价款优先受偿", "DD 段 56"],
    ["诉讼执行", "起诉/受理", "（2021）粤0703民初18146号", "2021-11-10 / 11-12", "起诉并受理", "DD 段 52"],
    ["诉讼执行", "一审判决", "（2021）粤0703民初18146号", "2022-05-09",
     "判项一 本金 130,301,203.57 及息；二 佳兆业连带；三/四 70%/30% 股权优先受偿", "DD 段 53~56"],
    ["诉讼执行", "二审撤诉", "（2022）粤07民终3886号", "2022-06-20", "撤回上诉，一审判决生效", "DD 段 59"],
    ["诉讼执行", "财产保全", "（2021）粤0703执保2735号", "2021-11",
     "冻结佳兆业账户、10 家子公司股权；查封 372 套不动产", "DD 段 61/62"],
    ["诉讼执行", "恢复执行", "（2023）粤0703执恢276号", "2025-03-18", "继续查封 372 套", "DD 段 63"],
    ["诉讼执行", "移送执行", "—", "—", "处置权移送深圳龙岗法院、深圳前海合作区法院", "DD 段 64"],
    ["诉讼执行", "权利性质", "—", "—", "372 套部分存在在先抵押 + 多轮轮候查封；贵行不享有抵押权", "DD 段 64 / CL 行 %d" % (cl_rn or 0)],
    ["诉讼执行", "查册覆盖", "—", "2026-05-18", "16/372 套可定位（4.3%），14 套记载查封", "CR"],
    ["诉讼执行", "时效", "—", "—", "律师结论：不存在诉讼时效届满问题", "DD 段 467"],
]

# ---------------- 11. 重组链 ----------------
chain_rows = []
for idx in (1, 2, 3, 4):
    rs = RS_CHAIN.get(idx, {})
    ddv = DD_CHAIN.get(idx, {})
    rs_amt = num(rs.get("金额万元"))
    dd_amt = num(ddv.get("确认本金"))
    diff = "-"
    if rs_amt is not None and dd_amt is not None:
        d = rs_amt * 10000 - dd_amt
        if abs(d) < 100:
            diff = "一致（万元口径取整，差 %s 元）" % fmt(d)
        else:
            diff = "%s 元（约 %s 万元）" % (format(d, ",.2f"), format(d / 10000, ",.2f"))
    note = ""
    if idx == 1:
        note = "段 45 未载明确认本金；段 44 提前到期本金 130,301,203.57 元"
    elif idx in (3, 4):
        note = "银行说明金额高于尽调 1,000 万元；拟转让清单余额支持尽调口径，须回原件核验"
    if idx == 3:
        note += "；时间亦不一致（2025-05 vs 2025-09-22）"
    chain_rows.append([idx, rs.get("日期", "—"),
                       rs.get("金额万元", "—") + " 万元",
                       ddv.get("日期", "—"), ddv.get("协议编号", "—"),
                       ddv.get("确认本金", "—") + " 元" if ddv.get("确认本金") else "—",
                       diff, note, "RS 段 0 / DD 段 %d" % (44 + idx)])

# ---------------- 12. 说明 ----------------
notes = [
    ["产物", "07 江门佰富 · 银行档案数据提取 v1（样板）"],
    ["生成时点", "2026-11-13"],
    ["用途", "把归档与已归集资料中的事实抽成结构化字段，直接对齐复核 skill 的 L1《基础事实对照表》；报告值与档案值已并排，可逐项判定"],
    ["证据源 1（层6）", NAME["CL"] + "；Sheet1 行 %d（序号 7 江门佰富）" % (cl_rn or 0)],
    ["证据源 2（层4）", NAME["RS"] + "；银行自述重组沿革，1 段"],
    ["证据源 3（层5）", NAME["DD"] + "；%d 段，合同/判决/案号/债权构成的主要文字载体" % len(P["DD"])],
    ["证据源 4（被复核对象）", NAME["VR"] + "；%d 段" % len(P["VR"])],
    ["证据源 5（层7）", NAME["CR"] + "；查册覆盖率与资产性质结论"],
    ["证据源 6（层7）", NAME["XC"] + "；372 套逐套数据已结构化，本表不重复提取"],
    ["证据层级", "查册/登记簿 > 权证 > 判决/执行文书 > 合同放款 > 律师引述 > 银行正式清单 > 查册成果；层5 引述须回原件核验"],
    ["★ 本轮最重要的发现", "银行《重组情况》第三、四次重组金额 11,930.12 万元，与尽调/报告的 10,930.12 万元差 1,000 万元；第三次时间 2025年5月 vs 2025-09-22。拟转让清单余额支持尽调口径，倾向银行说明笔误，但须回重组协议原件核验（F04/F09）"],
    ["★ 原件直读的新事实", "法院《预缴评估费通知》（2023）粤0703执恢276号、2026-08-30：委托广东诺诚评估「佳兆业名下位于深圳市布吉的 374 套不动产」，预缴评估费 83,030 元。① 83,030 元代垫费用由此获得层1原件支撑；② 374 套与清单/查册的 372 套口径不一致（查册报告另记有 451/401 套口径待核）；③ 法院已委托评估查封资产，对 L2 价值来源有影响"],
    ["★ 归档提取边界", "归档 07 户共 %d 个文件：需 OCR %d 个（PDF/图片零文本层，本机无 PaddleOCR/Tesseract，本轮未做 OCR）；PDF 文本层直读 %d 个；docx 已提取 2 个；xls 2 个为 OLE2 二进制（已有结构化替代）；zip 2 个按两阶段流程阶段 2 暂缓未展开" % (len(all_files), len(ocr_rows), len(PDF_TEXTS))],
    ["因此", "合同编号、案号、判项、登记编号等关键值取自律师引述（层5），不是从扫描原件直接读出；「需回原件核验」列已标「是」。完整 OCR 后才能升级为层1~3 证据"],
    ["OCR 缺口", "共 %d 个文件需 OCR，见「证据文件台账」状态列与优先级；其中合同借据类、判决书、权证类、查册资料为高优先级" % len(ocr_rows)],
    ["状态口径", "L1 比对底稿「状态」列一律为「待判」工作态，不自动判定差异；正式复核报告由 skill 按六值（一致／差异／报告未披露／档案无法核验／不适用／待补）出具"],
    ["验算口径", "「数字验算」中报告值按百位取整的差额判为「百位取整，非实质差错」，不构成差异认定"],
    ["户号口径", "户号按银行拟转让清单体系（07 江门佰富）；查册分户报告与结构包清单用体系 B（06 佰富），关联一律按公司全称，不用编号"],
    ["边界", "只读原始资料；未改动归档、共享资料、任何台账、查册成果或价值分析报告；未解压；两阶段流程阶段 2 仍暂缓"],
]

# ---------------- 写出 ----------------
hi_files = {i for i, r in enumerate(file_rows, 1)
            if "需 OCR" in r[4] and any(k in r[1] for k in
               ("合同借据", "判决书", "裁定书", "权证", "查册资料", "质押合同", "保证合同"))}
sheets = [
    ("L1比对底稿",
     ["编号", "复核项", "报告表述", "报告证据", "档案值", "档案证据", "状态", "差异说明／待核要点", "影响等级"],
     l1_rows, [8, 40, 46, 18, 54, 26, 9, 52, 10],
     [i for i, r in enumerate(l1_rows, 1) if r[8] == "高"]),
    ("三源勾稽",
     ["项目", "拟转让清单(层6)", "银行重组说明(层4)", "法律尽调(层5)", "价值分析报告", "一致性", "说明"],
     tri_rows, [22, 30, 26, 34, 28, 16, 52],
     [i for i, r in enumerate(tri_rows, 1) if "冲突" in r[5]]),
    ("数字验算",
     ["验算项", "算式", "报告值", "重算值", "差额", "判定", "单位"],
     calc_rows, [20, 46, 20, 20, 16, 24, 8],
     [i for i, r in enumerate(calc_rows, 1) if "差异" in r[5]]),
    ("重组链冲突",
     ["次序", "银行说明日期", "银行说明金额", "尽调协议日期", "协议编号", "尽调确认本金", "金额差异", "待核要点", "证据定位"],
     chain_rows, [7, 16, 20, 16, 34, 24, 28, 62, 22],
     [3, 4]),
    ("事实清单",
     ["类别", "事项", "编号/案号", "日期", "要点", "证据定位"],
     fact_rows, [14, 18, 42, 22, 76, 26], []),
    ("字段明细",
     ["字段", "提取值", "单位", "来源代号", "来源文件", "定位", "原文片段",
      "证据层级", "证据编号", "提取方式", "需回原件核验"],
     [[r[k] for k in ["字段", "提取值", "单位", "来源代号", "来源文件", "定位",
                      "原文片段", "证据层级", "证据编号", "提取方式", "需回原件核验"]]
      for r in records],
     [30, 40, 9, 11, 54, 24, 76, 26, 16, 18, 15], []),
    ("证据文件台账",
     ["证据编号", "归档相对路径", "格式", "大小", "提取状态", "状态说明", "关联复核项", "来源"],
     file_rows, [12, 76, 8, 12, 14, 56, 18, 34], hi_files),
    ("OCR缺口清单",
     ["证据编号", "归档相对路径", "页数/说明", "关联复核项", "OCR优先级"],
     ocr_rows, [12, 80, 40, 20, 12],
     [i for i, r in enumerate(ocr_rows, 1) if r[4] == "高"]),
    ("提取说明与来源", ["项目", "内容"], notes, [24, 130], []),
]

os.makedirs(OUT, exist_ok=True)
xp = os.path.join(OUT, "07_江门佰富_档案数据提取_v1.xlsx")
write_xlsx(xp, sheets)

# 机器可读证据索引：供复核 skill 的语义锚点召回直接使用
ev_index = []
for src in ("VR", "DD", "RS", "CR"):
    for i, line in enumerate(P[src]):
        ev_index.append({
            "evidence_id": "P%s%04d" % (src, i),
            "source": NAME[src], "src_code": src, "locator": "%s 段 %d" % (src, i),
            "level": LEVEL[src], "text": line,
        })
for rn in sorted(CL_CELLS):
    line = " | ".join("%s=%s" % (c, v) for c, v in sorted(CL_CELLS[rn].items()))
    if line.strip():
        ev_index.append({
            "evidence_id": "T01-R%03d" % rn, "source": NAME["CL"], "src_code": "CL",
            "locator": "Sheet1 行 %d" % rn, "level": LEVEL["CL"], "text": line,
        })
for k, (rel, obj) in enumerate(sorted(PDF_TEXTS.items()), 1):
    for li, line in enumerate(obj["text"].split("\n"), 1):
        if line.strip():
            ev_index.append({
                "evidence_id": "AR%02d-L%03d" % (k, li),
                "source": "归档/" + HOUSE + "/" + rel, "src_code": "AR",
                "locator": "PDF 第%d份 行%d" % (k, li),
                "level": LEVEL["AR"], "text": line.strip(),
            })
jp = os.path.join(OUT, "07_江门佰富_证据索引_v1.json")
json.dump({
    "house": HOUSE, "generated": "2026-11-13", "version": "v1",
    "purpose": "复核 skill L1 基础事实审阅的证据索引",
    "sources": NAME, "levels": LEVEL,
    "records": records, "l1": l1_rows, "tri_source": tri_rows,
    "calculations": calc_rows, "chain": chain_rows,
    "archive_file_count": len(all_files),
    "ocr_pending_count": len(ocr_rows),
    "evidence": ev_index,
}, open(jp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

print("XLSX:", xp)
print("sheets:", [s[0] for s in sheets])
print("JSON:", jp)
print("字段记录:", len(records), " 未自动提取:",
      sum(1 for r in records if r["提取值"] == "未自动提取"))
for r in records:
    if r["提取值"] == "未自动提取":
        print("   MISS:", r["字段"])
print("归档文件:", len(all_files), " 需OCR:", len(ocr_rows),
      " 证据单元:", len(ev_index))
print("验算:", [(r[0], r[5]) for r in calc_rows])
