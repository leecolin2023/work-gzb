# -*- coding: utf-8 -*-
"""零依赖最小 XLSX 写出器（inlineStr，无 sharedStrings）。

用法：
    from xlsx_mini import write_xlsx
    write_xlsx(path, [(sheet_name, header, rows, widths, highlight_rows), ...])

highlight_rows: 需要标黄的数据行序号（从 1 开始，不含表头）的可迭代对象。
仅依赖 Python 标准库；所有单元格按字符串写出，避免类型/编码问题。
"""
import os
import re
import zipfile

_INVALID = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _col(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _esc(s):
    if s is None:
        s = ""
    s = _INVALID.sub("", str(s))
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def _safe_sheet_name(name, used):
    n = re.sub(r"[\[\]:*?/\\]", "_", str(name))[:31] or "Sheet"
    base, i = n, 2
    while n in used:
        suffix = "_%d" % i
        n = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(n)
    return n


_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
%s
</Types>"""

_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="10"/><name val="DengXian"/></font>
<font><b/><sz val="10"/><color rgb="FF0C447C"/><name val="DengXian"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFE6F1FB"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def _sheet_xml(header, rows, widths, highlight_rows):
    ncol = max([len(header)] + [len(r) for r in rows] + [1])
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
             '<dimension ref="A1:%s%d"/>' % (_col(ncol), len(rows) + 1),
             '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
             '<sheetFormatPr defaultRowHeight="15"/>']
    if widths:
        parts.append('<cols>')
        for i, w in enumerate(widths, 1):
            parts.append('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i, i, w))
        parts.append('</cols>')
    parts.append('<sheetData>')

    def row_xml(rnum, values, style):
        cells = []
        for j in range(ncol):
            v = values[j] if j < len(values) else ""
            ref = "%s%d" % (_col(j + 1), rnum)
            if v is None or v == "":
                cells.append('<c r="%s" s="%d"/>' % (ref, style))
            else:
                cells.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                             % (ref, style, _esc(v)))
        return '<row r="%d">%s</row>' % (rnum, "".join(cells))

    parts.append(row_xml(1, list(header), 1))
    for i, r in enumerate(rows, 1):
        style = 2 if i in highlight_rows else 0
        parts.append(row_xml(i + 1, list(r), style))
    parts.append('</sheetData>')
    if header:
        parts.append('<autoFilter ref="A1:%s%d"/>' % (_col(ncol), len(rows) + 1))
    parts.append('</worksheet>')
    return "".join(parts)


def write_xlsx(path, sheets):
    """sheets: [(name, header, rows, widths, highlight_rows), ...]"""
    used = set()
    names = [_safe_sheet_name(s[0], used) for s in sheets]
    n = len(sheets)

    ct_overrides = "".join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
        for i in range(n))

    wb_sheets = "".join(
        '<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (_esc(names[i]), i + 1, i + 1)
        for i in range(n))
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets>%s</sheets></workbook>' % wb_sheets)

    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1) for i in range(n))
               + '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (n + 1)
               + '</Relationships>')

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CONTENT_TYPES % ct_overrides)
        z.writestr("_rels/.rels", _ROOT_RELS)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", _STYLES)
        for i, s in enumerate(sheets, 1):
            header = s[1]
            rows = s[2]
            widths = s[3] if len(s) > 3 else []
            hi = set(s[4] if len(s) > 4 else [])
            z.writestr("xl/worksheets/sheet%d.xml" % i,
                       _sheet_xml(header, rows, widths, hi))
    return path
