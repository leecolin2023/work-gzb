# -*- coding: utf-8 -*-
"""
export_docx.py — 财务分析报告 Markdown → 公文标准 DOCX 渲染器

用法:
    python export_docx.py <报告.md> <输出.docx> [选项]
    python export_docx.py --selftest

常用选项:
    --template <模板.docx>       基于现有官方模板写入，保留模板页眉页脚与页面设置
    --recipient <主送单位>       在标题下方插入主送单位（左对齐顶格，自动补全角冒号）
    --issuer <落款单位>          在正文末尾插入落款单位
    --date <日期>                在落款单位下方插入日期（如 2026年11月20日，月日不补零）
    --keep-numbering             保留 Markdown 原始标题序号（默认将数字标题规范化为 一、/（一）/1./（1））
    --no-landscape               禁用超宽表（≥8列）自动 A4 横向分节
    --no-first-page-num          首页不显示页码（适用于信函式等特殊格式）
    --include-json               保留文末 Machine-Readable Summary JSON 代码块（默认正式公文版省略）

排版口径（对齐《广州银行公文格式标准》）:
    - 页面设置：A4（210mm × 297mm），上 37mm / 下 35mm / 左 28mm / 右 26mm（版心 156mm × 225mm）。
    - 页码：奇偶页不同，4 号宋体半角阿拉伯数字，左右各一字线（— 1 —），奇数页右空一字，偶数页左空一字。
    - 字体字号：标题 2 号方正小标宋简体（固定行距 31pt）；一级标题 3 号黑体；二级标题 3 号楷体；
      三/四级标题及正文 3 号仿宋体（固定行距 28pt，首行缩进 2 字符）；数字随所在语境字体。
    - 表格契约：正文表格用仿宋体、附件表格用宋体、表首用黑体（跨页重复表头）、表内一级分类用楷体、
      表尾注用楷体小一号；常规密度用四号（14pt），密集表自动降为小四（12pt）；统计数字右对齐、
      短标签/状态/序号居中、说明性长文本左对齐且多句段落首行缩进 2 字符；外框 1pt 深色、内框 0.5pt 浅色。
"""
import sys, os, re, io, argparse, tempfile, datetime
import docx
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls


# ---------------------------------------------------------------- 字体与常量
FONT_TITLE = '方正小标宋简体'
FONT_HEI = '黑体'
FONT_KAI = '楷体'
FONT_FANG = '仿宋'
FONT_SONG = '宋体'

SIZE_ERHAO = Pt(22)      # 二号
SIZE_SANHAO = Pt(16)     # 三号
SIZE_SIHAO = Pt(14)      # 四号
SIZE_XIAOSI = Pt(12)     # 小四
SIZE_WUHAO = Pt(10.5)    # 五号

LINE_BODY = Pt(28)       # 正文固定行距
LINE_TITLE = Pt(31)      # 标题固定行距

MARGIN_TOP = Mm(37)
MARGIN_BOTTOM = Mm(35)
MARGIN_LEFT = Mm(28)
MARGIN_RIGHT = Mm(26)
FOOTER_DIST = Mm(23)     # 底边距 35mm 下，页码上沿距版心下边缘约 7mm

TEXT_WIDTH_PORTRAIT_MM = 156.0
TEXT_WIDTH_LANDSCAPE_MM = 243.0  # 297 - 28 - 26

BORDER_OUTER_SZ = '8'    # 1 pt (8 eighths of a point)
BORDER_INNER_SZ = '4'    # 0.5 pt (4 eighths of a point)
BORDER_OUTER_COLOR = '1A1A1A'
BORDER_INNER_COLOR = '7F7F7F'
HEADER_SHADING_FILL = 'F2F2F2'

CN_NUMS = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十']


def int_to_cn(n):
    if n <= 0:
        return str(n)
    if n <= 10:
        return CN_NUMS[n]
    if n < 20:
        return '十' + CN_NUMS[n - 10]
    if n < 100:
        tens, ones = divmod(n, 10)
        return CN_NUMS[tens] + '十' + (CN_NUMS[ones] if ones else '')
    return str(n)


# ---------------------------------------------------------------- XML 辅助函数
def set_run_font(run, font_name, size_pt, bold=False, color_rgb=None):
    """统一设置中西文字体，确保正文数字沿用所在公文字体。"""
    run.font.name = font_name
    run.font.size = size_pt
    run.bold = bold
    if color_rgb is not None:
        run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)


def set_first_line_chars(p, chars=200):
    """以字符为单位设置首行缩进（200 = 2 个汉字字符），兼容 WPS 与 Word。"""
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLineChars'), str(chars))


def set_left_right_chars(p, left_chars=None, right_chars=None):
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    if left_chars is not None:
        ind.set(qn('w:leftChars'), str(left_chars))
    if right_chars is not None:
        ind.set(qn('w:rightChars'), str(right_chars))


def set_hanging_chars(p, left_chars=500, hanging_chars=300):
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:leftChars'), str(left_chars))
    ind.set(qn('w:hangingChars'), str(hanging_chars))


def add_page_field(p, font_name=FONT_SONG, size_pt=SIZE_SIHAO):
    """插入公文标准页码：— <PAGE> —"""
    r1 = p.add_run('— ')
    set_run_font(r1, font_name, size_pt)

    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')

    r_pg = p.add_run()
    set_run_font(r_pg, font_name, size_pt)
    r_pg._r.append(fld_begin)
    r_pg._r.append(instr)
    r_pg._r.append(fld_sep)
    r_pg._r.append(fld_end)

    r2 = p.add_run(' —')
    set_run_font(r2, font_name, size_pt)


def configure_section_geometry(section, landscape=False, no_first_page_num=False, preserve_template=False):
    if not preserve_template:
        if landscape:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Mm(297)
            section.page_height = Mm(210)
            section.top_margin = Mm(28)
            section.bottom_margin = Mm(26)
            section.left_margin = Mm(28)
            section.right_margin = Mm(26)
        else:
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Mm(210)
            section.page_height = Mm(297)
            section.top_margin = MARGIN_TOP
            section.bottom_margin = MARGIN_BOTTOM
            section.left_margin = MARGIN_LEFT
            section.right_margin = MARGIN_RIGHT
        section.footer_distance = FOOTER_DIST
        section.header_distance = Mm(15)
        section.different_first_page_header_footer = bool(no_first_page_num)

        # 奇数页页脚：右对齐，右空一字
        section.footer.is_linked_to_previous = False
        fp_odd = section.footer.paragraphs[0]
        fp_odd.text = ''
        fp_odd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fp_odd.paragraph_format.space_before = Pt(0)
        fp_odd.paragraph_format.space_after = Pt(0)
        fp_odd.paragraph_format.right_indent = Pt(14)
        set_left_right_chars(fp_odd, right_chars=100)
        add_page_field(fp_odd, FONT_SONG, SIZE_SIHAO)

        # 偶数页页脚：左对齐，左空一字
        section.even_page_footer.is_linked_to_previous = False
        fp_even = section.even_page_footer.paragraphs[0]
        fp_even.text = ''
        fp_even.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fp_even.paragraph_format.space_before = Pt(0)
        fp_even.paragraph_format.space_after = Pt(0)
        fp_even.paragraph_format.left_indent = Pt(14)
        set_left_right_chars(fp_even, left_chars=100)
        add_page_field(fp_even, FONT_SONG, SIZE_SIHAO)

        if no_first_page_num:
            section.first_page_footer.is_linked_to_previous = False
            section.first_page_footer.paragraphs[0].text = ''


def set_table_borders(table):
    """外框 1pt 深色，内框 0.5pt 浅色。"""
    tblPr = table._tbl.tblPr
    borders = tblPr.find(qn('w:tblBorders'))
    if borders is not None:
        tblPr.remove(borders)
    borders_xml = (
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="{BORDER_OUTER_SZ}" w:space="0" w:color="{BORDER_OUTER_COLOR}"/>'
        f'  <w:bottom w:val="single" w:sz="{BORDER_OUTER_SZ}" w:space="0" w:color="{BORDER_OUTER_COLOR}"/>'
        f'  <w:left w:val="single" w:sz="{BORDER_OUTER_SZ}" w:space="0" w:color="{BORDER_OUTER_COLOR}"/>'
        f'  <w:right w:val="single" w:sz="{BORDER_OUTER_SZ}" w:space="0" w:color="{BORDER_OUTER_COLOR}"/>'
        f'  <w:insideH w:val="single" w:sz="{BORDER_INNER_SZ}" w:space="0" w:color="{BORDER_INNER_COLOR}"/>'
        f'  <w:insideV w:val="single" w:sz="{BORDER_INNER_SZ}" w:space="0" w:color="{BORDER_INNER_COLOR}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(parse_xml(borders_xml))


def set_cell_shading(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=60, bottom=60, left=90, right=90):
    """单元格内边距（单位 twips，60≈3pt，90≈4.5pt）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side, val in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_row_repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:tblHeader'))


def set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))


def set_table_fixed_layout(table, total_width_mm, col_widths_mm):
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is not None:
        tblPr.remove(tblW)
    tblPr.append(parse_xml(
        f'<w:tblW {nsdecls("w")} w:w="{int(round(total_width_mm * 56.6929))}" w:type="dxa"/>'
    ))
    tblLayout = tblPr.find(qn('w:tblLayout'))
    if tblLayout is not None:
        tblPr.remove(tblLayout)
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

    tblGrid = table._tbl.tblGrid
    if tblGrid is not None:
        table._tbl.remove(tblGrid)
    new_grid = OxmlElement('w:tblGrid')
    for w_mm in col_widths_mm:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(int(round(w_mm * 56.6929))))
        new_grid.append(gc)
    table._tbl.insert(1, new_grid)

    for row in table.rows:
        for idx, w_mm in enumerate(col_widths_mm):
            if idx < len(row.cells):
                tc = row.cells[idx]._tc
                tcPr = tc.get_or_add_tcPr()
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is not None:
                    tcPr.remove(tcW)
                tcPr.append(parse_xml(
                    f'<w:tcW {nsdecls("w")} w:w="{int(round(w_mm * 56.6929))}" w:type="dxa"/>'
                ))


# ---------------------------------------------------------------- Markdown 解析器
def strip_inline_md(s):
    s = re.sub(r'\*\*(.+?)\*\*', r'\1', s)
    s = re.sub(r'`([^`]+)`', r'\1', s)
    s = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'\1', s)
    return s.strip()


def add_inline_runs(p, text, font_name, size_pt, base_bold=False):
    """解析行内 **加粗** 与 `代码` 标记，保留清晰加粗但不输出原始 Markdown 符号。"""
    text = text.replace('<br/>', '\n').replace('<br>', '\n')
    tokens = re.split(r'(\*\*.*?\*\*|`[^`]+`)', text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith('**') and tok.endswith('**') and len(tok) > 4:
            r = p.add_run(tok[2:-2])
            set_run_font(r, font_name, size_pt, bold=True)
        elif tok.startswith('`') and tok.endswith('`') and len(tok) > 2:
            r = p.add_run(tok[1:-1])
            set_run_font(r, font_name, size_pt, bold=base_bold)
        else:
            cleaned = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'\1', tok)
            r = p.add_run(cleaned)
            set_run_font(r, font_name, size_pt, bold=base_bold)


def is_table_sep(line):
    s = line.strip()
    if not s.startswith('|'):
        return False
    parts = [c.strip() for c in s.strip('|').split('|')]
    return bool(parts) and all(re.fullmatch(r':?-{2,}:?', p) for p in parts)


def split_md_row(line):
    s = line.strip().strip('|')
    # 兼容单元格内转义 \|
    placeholder = '\x00PIPE\x00'
    s = s.replace(r'\|', placeholder)
    return [c.replace(placeholder, '|').strip() for c in s.split('|')]


def parse_markdown_blocks(md_text):
    """将 Markdown 报告解析为结构化块列表。"""
    lines = md_text.replace('\r\n', '\n').split('\n')
    blocks = []
    i = 0
    n = len(lines)
    while i < n:
        ln = lines[i]
        s = ln.strip()

        if not s:
            i += 1
            continue

        # 水平分隔线（公文正文不使用 Markdown 横线）
        if re.fullmatch(r'[-*_]{3,}', s):
            i += 1
            continue

        # 围栏代码块
        if s.startswith('```'):
            lang = s[3:].strip().lower()
            code_lines = []
            i += 1
            while i < n and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1
            blocks.append({'type': 'code', 'lang': lang, 'text': '\n'.join(code_lines)})
            continue

        # 标题
        m_h = re.match(r'^(#{1,4})\s+(.+?)\s*$', s)
        if m_h:
            blocks.append({'type': 'heading', 'level': len(m_h.group(1)), 'text': m_h.group(2).strip()})
            i += 1
            continue

        # Markdown 表格
        if s.startswith('|') and i + 1 < n and is_table_sep(lines[i + 1]):
            header = split_md_row(lines[i])
            i += 2
            rows = []
            while i < n and lines[i].strip().startswith('|'):
                r = split_md_row(lines[i])
                if len(r) < len(header):
                    r += [''] * (len(header) - len(r))
                elif len(r) > len(header):
                    r = r[:len(header)]
                rows.append(r)
                i += 1
            blocks.append({'type': 'table', 'header': header, 'rows': rows})
            continue

        # 引用块 >
        if s.startswith('>'):
            q_lines = []
            while i < n and lines[i].strip().startswith('>'):
                q_lines.append(re.sub(r'^>\s?', '', lines[i].strip()))
                i += 1
            blocks.append({'type': 'quote', 'lines': [x for x in q_lines if x]})
            continue

        # 列表项（逐条输出为规范公文段落，避免混成一团）
        m_ul = re.match(r'^[-*+]\s+(?:\[[ xX]\]\s+)?(.+)$', s)
        if m_ul:
            blocks.append({'type': 'list_item', 'ordered': False, 'text': m_ul.group(1).strip()})
            i += 1
            continue

        m_ol = re.match(r'^(\d+[.、])\s+(.+)$', s)
        if m_ol:
            blocks.append({'type': 'list_item', 'ordered': True, 'marker': m_ol.group(1), 'text': m_ol.group(2).strip()})
            i += 1
            continue

        # 普通段落（合并连续普通行）
        p_lines = [s]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (not nxt or nxt.startswith('#') or nxt.startswith('|') or nxt.startswith('>')
                    or nxt.startswith('```') or re.fullmatch(r'[-*_]{3,}', nxt)
                    or re.match(r'^[-*+]\s+', nxt) or re.match(r'^\d+[.、]\s+', nxt)):
                break
            p_lines.append(nxt)
            i += 1
        blocks.append({'type': 'paragraph', 'text': ' '.join(p_lines)})

    return blocks


# ---------------------------------------------------------------- 单元格分类与列宽算法
NUM_CELL_RE = re.compile(
    r'^[+\-−–]?\s*\d[\d,]*(?:\.\d+)?\s*(?:%|pp|倍|x|天|万元|亿元|元|年|期|个|处)?$'
)
SHORT_LABEL_RE = re.compile(
    r'^(?:Green|Yellow|Orange|Red|NR|PASS|FAIL|FULL-PASS|NO-MISMATCH|APPROVE|REVISE|REJECT|'
    r'L[123]|FM-\d+|DQ\d+|RR\d+|OR\d+|F\d+|D\d+|O-\d+|Z-\d+|\[(?:事实|判断|待核)\]|'
    r'✅.*|❌.*|触发|未触发|通过|不通过|数据缺失|未计算|NOT_COMPUTABLE|—|-|–|无|是|否|'
    r'20\d\d[-/年]?\d{0,2}[HQhq]?\d?|[+−\-]{3})$',
    re.I
)


def classify_cell(raw_text):
    """按《公文格式标准》表格契约判定单元格语义类型与对齐方式。"""
    plain = strip_inline_md(raw_text).replace('\n', ' ').strip()
    if not plain or plain in ('—', '-', '–'):
        return 'label'
    if NUM_CELL_RE.fullmatch(plain):
        return 'numeric'
    if SHORT_LABEL_RE.fullmatch(plain) or (len(plain) <= 6 and not re.search(r'[，。；：]', plain)):
        return 'label'
    return 'text'


def is_long_prose_cell(raw_text):
    plain = strip_inline_md(raw_text).strip()
    if '<br' in raw_text or '\n' in raw_text:
        return True
    if len(plain) >= 24 and re.search(r'[。；]', plain):
        return True
    if len(plain) >= 34 and re.search(r'[，。；]', plain):
        return True
    return False


def display_len(s):
    """估算中文全角=2、ASCII=1 的视觉宽度。"""
    plain = strip_inline_md(s)
    return sum(2 if ord(ch) > 127 else 1 for ch in plain)


def compute_col_widths(header, rows, total_mm):
    ncols = len(header)
    if ncols == 0:
        return []
    weights = []
    for c in range(ncols):
        vals = [display_len(header[c])] + [display_len(r[c]) for r in rows if c < len(r)]
        vals_sorted = sorted(vals)
        p85 = vals_sorted[int(0.85 * (len(vals_sorted) - 1))]
        # 上下限裁剪防止极短列过窄或超长列挤死其余列
        w = max(8.0, min(float(p85), 48.0))
        weights.append(w)
    s = sum(weights) or 1.0
    min_mm = 11.0 if ncols >= 8 else 14.0
    raw = [max(min_mm, total_mm * (w / s)) for w in weights]
    scale = total_mm / sum(raw)
    return [r * scale for r in raw]


# ---------------------------------------------------------------- 标题规范化
def normalize_heading_text(level, raw_text, counters, keep_numbering=False):
    text = strip_inline_md(raw_text)
    if keep_numbering:
        return text

    # 判断是否已是公文序号
    if level == 2:
        if re.match(r'^[一二三四五六七八九十]+、', text) or re.match(r'^附件\s*\d*', text):
            counters['h3'] = 0
            counters['h4'] = 0
            return text
        m = re.match(r'^(\d+)\s*[.、:：]\s*(.+)$', text)
        if m:
            counters['h2'] += 1
            counters['h3'] = 0
            counters['h4'] = 0
            return f'{int_to_cn(counters["h2"])}、{m.group(2).strip()}'
        counters['h2'] += 1
        counters['h3'] = 0
        counters['h4'] = 0
        return f'{int_to_cn(counters["h2"])}、{text}'

    if level == 3:
        if re.match(r'^（[一二三四五六七八九十]+）', text):
            counters['h4'] = 0
            return text
        m = re.match(r'^(?:\d+(?:\.\d+)*|\([一二三四五六七八九十]+\))\s*[.、:：]?\s*(.+)$', text)
        body = m.group(1).strip() if m else text
        counters['h3'] += 1
        counters['h4'] = 0
        return f'（{int_to_cn(counters["h3"])}）{body}'

    if level == 4:
        if re.match(r'^\d+\.', text) or re.match(r'^（\d+）', text):
            return text
        counters['h4'] += 1
        return f'{counters["h4"]}. {text}'

    return text


def is_machine_summary_heading(text):
    t = text.lower()
    return ('machine-readable' in t) or ('machine readable' in t) or ('机器可读' in t)


def is_attachment_heading(text):
    return bool(re.match(r'^附件(?:\s*\d+)?(?:[：:.\s]|$)', strip_inline_md(text)))


def is_table_note_paragraph(text):
    plain = strip_inline_md(text)
    return bool(re.match(r'^(?:注|说明|口径说明|注\d+)[：:]', plain))


# ---------------------------------------------------------------- 主渲染函数
def render_markdown_to_docx(
    md_text,
    out_path,
    template_path=None,
    recipient=None,
    issuer=None,
    doc_date=None,
    keep_numbering=False,
    landscape_wide_tables=True,
    no_first_page_num=False,
    include_json=False,
):
    preserve_template = bool(template_path)
    doc = Document(template_path) if template_path else Document()

    # 启用奇偶页不同页脚
    if not preserve_template:
        doc.settings.odd_and_even_pages_header_footer = True

    # 设置正文默认样式
    normal_style = doc.styles['Normal']
    normal_style.font.name = FONT_FANG
    normal_style.font.size = SIZE_SANHAO
    rPr = normal_style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rFonts.set(qn(attr), FONT_FANG)

    configure_section_geometry(
        doc.sections[0],
        landscape=False,
        no_first_page_num=no_first_page_num,
        preserve_template=preserve_template,
    )

    blocks = parse_markdown_blocks(md_text)
    counters = {'h2': 0, 'h3': 0, 'h4': 0}
    in_attachment = False
    skip_until_next_h2 = False
    last_table_font_size = SIZE_SIHAO
    prev_was_table = False
    title_rendered = False

    def add_body_para(text, font_name=FONT_FANG, size_pt=SIZE_SANHAO, indent_2=True,
                      align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_pt=LINE_BODY, bold=False, keep_next=False):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = line_pt
        pf.keep_with_next = keep_next
        if indent_2:
            pf.first_line_indent = Pt(size_pt.pt * 2)
            set_first_line_chars(p, 200)
        else:
            pf.first_line_indent = Pt(0)
        add_inline_runs(p, text, font_name, size_pt, base_bold=bold)
        return p

    for idx, blk in enumerate(blocks):
        btype = blk['type']

        # 跳过 Machine-Readable Summary 章节（默认正式公文版省略）
        if btype == 'heading' and blk['level'] <= 2:
            if not include_json and is_machine_summary_heading(blk['text']):
                skip_until_next_h2 = True
                continue
            skip_until_next_h2 = False

        if skip_until_next_h2:
            continue

        # 1. 文档主标题（H1）
        if btype == 'heading' and blk['level'] == 1:
            t_txt = strip_inline_md(blk['text']).rstrip('。.')
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = LINE_TITLE
            pf.keep_with_next = True
            r = p.add_run(t_txt)
            set_run_font(r, FONT_TITLE, SIZE_ERHAO, bold=False)

            # 标题后空一行
            sp = doc.add_paragraph()
            sp.paragraph_format.space_before = Pt(0)
            sp.paragraph_format.space_after = Pt(0)
            sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            sp.paragraph_format.line_spacing = LINE_BODY

            # 可选主送单位（左对齐顶格，末加全角冒号）
            if recipient and not title_rendered:
                rec_txt = recipient.strip()
                if not rec_txt.endswith('：'):
                    rec_txt = rec_txt.rstrip(':') + '：'
                add_body_para(rec_txt, FONT_FANG, SIZE_SANHAO, indent_2=False, align=WD_ALIGN_PARAGRAPH.LEFT)
            title_rendered = True
            prev_was_table = False
            continue

        # 2. 附件起页处理（H2 以“附件”开头）
        if btype == 'heading' and blk['level'] == 2 and is_attachment_heading(blk['text']):
            in_attachment = True
            doc.add_page_break()
            raw_att = strip_inline_md(blk['text'])
            m_att = re.match(r'^(附件\s*\d*)[：:.\s]*(.*)$', raw_att)
            att_mark = (m_att.group(1).replace(' ', '') if m_att else '附件')
            att_title = (m_att.group(2).strip() if m_att else '')

            # 左上角：附件N（3号黑体，顶格）
            add_body_para(att_mark, FONT_HEI, SIZE_SANHAO, indent_2=False, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
            # 空一行后第三行居中排附件标题（2号小标宋体）
            sp = doc.add_paragraph()
            sp.paragraph_format.space_before = Pt(0)
            sp.paragraph_format.space_after = Pt(0)
            sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            sp.paragraph_format.line_spacing = LINE_BODY
            if att_title:
                pt = doc.add_paragraph()
                pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pt.paragraph_format.space_before = Pt(0)
                pt.paragraph_format.space_after = Pt(0)
                pt.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                pt.paragraph_format.line_spacing = LINE_TITLE
                pt.paragraph_format.keep_with_next = True
                rt = pt.add_run(att_title)
                set_run_font(rt, FONT_TITLE, SIZE_ERHAO)
            counters['h2'] = counters['h3'] = counters['h4'] = 0
            prev_was_table = False
            continue

        # 3. 各级标题（H2 黑体 / H3 楷体 / H4 仿宋体）
        if btype == 'heading':
            lvl = blk['level']
            htext = normalize_heading_text(lvl, blk['text'], counters, keep_numbering=keep_numbering)
            if lvl == 2:
                add_body_para(htext, FONT_HEI, SIZE_SANHAO, indent_2=True, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
            elif lvl == 3:
                add_body_para(htext, FONT_KAI, SIZE_SANHAO, indent_2=True, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
            else:
                add_body_para(htext, FONT_FANG, SIZE_SANHAO, indent_2=True, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
            prev_was_table = False
            continue

        # 4. Markdown 表格（完整落实表格格式契约）
        if btype == 'table':
            header = blk['header']
            rows = blk['rows']
            ncols = len(header)
            is_wide = landscape_wide_tables and (ncols >= 8) and (not preserve_template)

            # 超宽表（≥8列，如 10 列主张-证据台账）自动切入 A4 横向分节
            if is_wide:
                sec_land = doc.add_section(WD_SECTION.NEW_PAGE)
                configure_section_geometry(sec_land, landscape=True, no_first_page_num=False, preserve_template=False)

            total_mm = TEXT_WIDTH_LANDSCAPE_MM if is_wide else TEXT_WIDTH_PORTRAIT_MM
            col_widths_mm = compute_col_widths(header, rows, total_mm)

            # 字号策略：常规密度用四号（14pt），密集表（≥6列或长文本多）降为小四（12pt）
            avg_cell_len = (
                sum(display_len(c) for r in rows for c in r) / max(1, len(rows) * ncols)
            )
            dense = (ncols >= 6) or (avg_cell_len >= 26)
            body_size = SIZE_XIAOSI if dense else SIZE_SIHAO
            last_table_font_size = body_size
            body_font = FONT_SONG if in_attachment else FONT_FANG
            cell_line_pt = Pt(18) if dense else Pt(21)

            tbl = doc.add_table(rows=1 + len(rows), cols=ncols)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(tbl)
            set_table_fixed_layout(tbl, total_mm, col_widths_mm)

            # 表首行：黑体、居中、跨页重复表头
            hdr_row = tbl.rows[0]
            set_row_repeat_header(hdr_row)
            set_row_cant_split(hdr_row)
            for c_idx, h_text in enumerate(header):
                cell = hdr_row.cells[c_idx]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                set_cell_shading(cell, HEADER_SHADING_FILL)
                set_cell_margins(cell, top=60, bottom=60, left=75, right=75)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                p.paragraph_format.line_spacing = cell_line_pt
                add_inline_runs(p, strip_inline_md(h_text), FONT_HEI, body_size, base_bold=False)

            # 预判各列主体是否为数字列（使同一数字列中的 "—" 与数值保持视觉整齐）
            col_is_num = []
            for c_idx in range(ncols):
                kinds = [classify_cell(r[c_idx]) for r in rows if c_idx < len(r) and strip_inline_md(r[c_idx]) not in ('', '—', '-', '–')]
                col_is_num.append(bool(kinds) and sum(1 for k in kinds if k == 'numeric') >= 0.6 * len(kinds))

            # 数据行
            for r_idx, r_data in enumerate(rows):
                row_obj = tbl.rows[1 + r_idx]
                set_row_cant_split(row_obj)
                first_plain = strip_inline_md(r_data[0]) if r_data else ''
                is_category_row = bool(re.match(r'^[一二三四五六七八九十]+、', first_plain))

                for c_idx in range(ncols):
                    raw_val = r_data[c_idx] if c_idx < len(r_data) else ''
                    cell = row_obj.cells[c_idx]
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    set_cell_margins(cell, top=55, bottom=55, left=75, right=75)

                    ctype = classify_cell(raw_val)
                    use_font = FONT_KAI if (is_category_row and c_idx == 0) else body_font

                    if ctype == 'numeric' or (col_is_num[c_idx] and strip_inline_md(raw_val) in ('—', '-', '–')):
                        align = WD_ALIGN_PARAGRAPH.RIGHT
                    elif ctype == 'label':
                        align = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        align = WD_ALIGN_PARAGRAPH.LEFT

                    # 长文本单元格分自然段 + 首行缩进 2 字符
                    if ctype == 'text' and is_long_prose_cell(raw_val):
                        sub_paras = [sp.strip() for sp in re.split(r'<br\s*/?>|\n', raw_val) if sp.strip()]
                        for sp_i, sp_txt in enumerate(sub_paras):
                            p = cell.paragraphs[0] if sp_i == 0 else cell.add_paragraph()
                            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            p.paragraph_format.space_before = Pt(0)
                            p.paragraph_format.space_after = Pt(0)
                            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                            p.paragraph_format.line_spacing = cell_line_pt
                            p.paragraph_format.first_line_indent = Pt(body_size.pt * 2)
                            set_first_line_chars(p, 200)
                            add_inline_runs(p, sp_txt, use_font, body_size)
                    else:
                        p = cell.paragraphs[0]
                        p.alignment = align
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)
                        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                        p.paragraph_format.line_spacing = cell_line_pt
                        add_inline_runs(p, raw_val, use_font, body_size)

            # 超宽表结束后切回 A4 纵向分节
            if is_wide:
                sec_port = doc.add_section(WD_SECTION.NEW_PAGE)
                configure_section_geometry(sec_port, landscape=False, no_first_page_num=False, preserve_template=False)

            prev_was_table = True
            continue

        # 5. 引用块（表尾注用楷体小一号，其余用标准仿宋体说明段）
        if btype == 'quote':
            for q_ln in blk['lines']:
                if prev_was_table and is_table_note_paragraph(q_ln):
                    note_size = SIZE_WUHAO if last_table_font_size == SIZE_XIAOSI else SIZE_XIAOSI
                    add_body_para(q_ln, FONT_KAI, note_size, indent_2=True, line_pt=Pt(22))
                else:
                    add_body_para(q_ln, FONT_FANG, SIZE_SANHAO, indent_2=True)
            prev_was_table = False
            continue

        # 6. 代码块（如 VERDICT 校验判词或显式保留的 JSON）
        if btype == 'code':
            if blk['lang'] == 'json' and not include_json:
                continue
            for cl in blk['text'].splitlines():
                if cl.strip():
                    add_body_para(cl, FONT_FANG, SIZE_XIAOSI, indent_2=False, align=WD_ALIGN_PARAGRAPH.LEFT, line_pt=Pt(20))
            prev_was_table = False
            continue

        # 7. 列表项（转为标准公文首行缩进段落）
        if btype == 'list_item':
            txt = (f'{blk["marker"]} {blk["text"]}' if blk['ordered'] else blk['text'])
            add_body_para(txt, FONT_FANG, SIZE_SANHAO, indent_2=True)
            prev_was_table = False
            continue

        # 8. 普通段落（识别表尾注、附件说明与正文段落）
        if btype == 'paragraph':
            plain_p = strip_inline_md(blk['text'])
            if prev_was_table and is_table_note_paragraph(blk['text']):
                note_size = SIZE_WUHAO if last_table_font_size == SIZE_XIAOSI else SIZE_XIAOSI
                add_body_para(blk['text'], FONT_KAI, note_size, indent_2=True, line_pt=Pt(22))
            elif re.match(r'^附件[：:]', plain_p):
                # 附件说明：正文后空一行，左空二字，回行与首个附件名称齐平，末尾不加标点
                sp = doc.add_paragraph()
                sp.paragraph_format.space_before = Pt(0)
                sp.paragraph_format.space_after = Pt(0)
                sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                sp.paragraph_format.line_spacing = LINE_BODY
                att_desc = plain_p.rstrip('。.;；')
                p_att = add_body_para(att_desc, FONT_FANG, SIZE_SANHAO, indent_2=False, align=WD_ALIGN_PARAGRAPH.LEFT)
                set_hanging_chars(p_att, left_chars=500, hanging_chars=300)
            else:
                add_body_para(blk['text'], FONT_FANG, SIZE_SANHAO, indent_2=True)
            prev_was_table = False

    # 9. 可选落款单位与日期（绝不凭空编造，仅在显式传参时输出）
    if issuer or doc_date:
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        sp.paragraph_format.line_spacing = LINE_BODY
        if issuer:
            p_iss = add_body_para(issuer.strip(), FONT_FANG, SIZE_SANHAO, indent_2=False, align=WD_ALIGN_PARAGRAPH.RIGHT)
            p_iss.paragraph_format.right_indent = Pt(64)
            set_left_right_chars(p_iss, right_chars=400)
        if doc_date:
            # 规范日期：月、日不补零
            d_clean = re.sub(r'(\d{4})[-/.年]0?(\d{1,2})[-/.月]0?(\d{1,2})日?', r'\1年\2月\3日', doc_date.strip())
            p_dt = add_body_para(d_clean, FONT_FANG, SIZE_SANHAO, indent_2=False, align=WD_ALIGN_PARAGRAPH.RIGHT)
            p_dt.paragraph_format.right_indent = Pt(64)
            set_left_right_chars(p_dt, right_chars=400)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------- 自检测试
def run_selftest():
    sample_md = """# 某制造企业 2023—2025Q3 财务分析报告

> 主体名称：测试制造有限公司 | 行业：通用设备制造 | 金额单位：万元

## 0. 输入边界与深度规格卡

本报告基于资产负债表、利润表与现金流量表原始底稿编制，所有派生指标均由确定性复算脚本生成。

### 0.1 派生指标口径登记表

| 编号 | 指标名称 | 计算公式 | 适用期间 | 口径说明 |
|---|---|---|---|---|
| O-01 | 资产负债率 | 负债合计 ÷ 资产总计 | 全部 | 通用参考口径 |
| O-02 | 已获利息倍数 | （利润总额 + 利息费用）÷ 利息费用 | 全部 | 触发 DQ005 时严禁替代 |

注：本表为报告全篇统一口径基准。

## 1. 核心财务指标（6 期密集表测试）

| 指标 | 2021-12 | 2022-12 | 2023-12 | 2024-12 | 2025-12 | 2026-06 |
|---|---|---|---|---|---|---|
| 一、偿债指标 | — | — | — | — | — | — |
| 资产负债率 | 58.3% | 62.1% | 65.4% | 67.0% | 68.0% | 66.8% |
| 流动比率 | 1.45 | 1.38 | 1.29 | 1.18 | 1.15 | 1.12 |

## 2. 主张-一手证据-替代解释-翻盘条件核心台账

| ID | 域 | 性质 | 核心主张 | 精确定位 | 证据边界 | 证据链数 | 正常业务替代解释 | 翻盘条件 | 验证层 |
|---|---|---|---|---|---|---|---|---|---|
| C-01 | A域 | [待核] | 在建工程余额持续高位，需核查转固进度 | 资产负债表第12行 | 期末存量无法证明报告期内是否发生转固；仅能说明期末余额维持高位。 | 1 | 产线分期技改持续投入，旧项目转固与新项目投入同步抵消。 | 查工程台账与验收单证实正常滚动转固。 | L2 |

## 附件1：补充核查明细表

| 优先级 | 资料名称 | 验证目的 |
|---|---|---|
| 1 | 银行授信与到期台账 | 核验短期有息负债集中度与续贷条件 |

## Machine-Readable Summary

```json
{"domain_a": "NR", "domain_b": "Yellow"}
```
"""
    with tempfile.TemporaryDirectory() as td:
        out_docx = os.path.join(td, 'selftest.docx')
        render_markdown_to_docx(
            sample_md,
            out_docx,
            recipient='授信审批委员会',
            issuer='风险管理部',
            doc_date='2026-11-05',
        )
        d = Document(out_docx)

        # 1. 验证页面几何与奇偶页码
        s0 = d.sections[0]
        assert round(s0.page_width.mm) == 210 and round(s0.page_height.mm) == 297
        assert round(s0.top_margin.mm) == 37 and round(s0.bottom_margin.mm) == 35
        assert round(s0.left_margin.mm) == 28 and round(s0.right_margin.mm) == 26
        assert d.settings.odd_and_even_pages_header_footer is True
        assert s0.footer.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
        assert s0.even_page_footer.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.LEFT
        assert 'PAGE' in s0.footer.paragraphs[0]._p.xml

        # 2. 验证标题与正文排版
        p_title = d.paragraphs[0]
        assert p_title.alignment == WD_ALIGN_PARAGRAPH.CENTER
        assert p_title.runs[0].font.size == SIZE_ERHAO
        assert p_title.runs[0].font.name == FONT_TITLE

        # 3. 验证表格契约（正文表仿宋、表首黑体、密集表小四、超宽表横向、附件表宋体）
        assert len(d.tables) == 4
        # 表1（5列常规密度 -> 四号仿宋，表首黑体）
        t1 = d.tables[0]
        assert t1.rows[0].cells[0].paragraphs[0].runs[0].font.name == FONT_HEI
        assert t1.rows[0].cells[0].paragraphs[0].runs[0].font.size == SIZE_SIHAO
        assert t1.rows[1].cells[1].paragraphs[0].runs[0].font.name == FONT_FANG
        assert 'w:tblHeader' in t1.rows[0]._tr.xml
        assert 'w:sz="8"' in t1._tbl.xml and 'w:sz="4"' in t1._tbl.xml

        # 表2（7列密集表 -> 小四，表内一级分类楷体，数字右对齐）
        t2 = d.tables[1]
        assert t2.rows[1].cells[0].paragraphs[0].runs[0].font.name == FONT_KAI
        assert t2.rows[2].cells[1].paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
        assert t2.rows[2].cells[1].paragraphs[0].runs[0].font.size == SIZE_XIAOSI

        # 表3（10列超宽表 -> 自动切 A4 横向）
        assert any(sec.orientation == WD_ORIENT.LANDSCAPE for sec in d.sections)

        # 表4（附件表 -> 宋体）
        t4 = d.tables[3]
        assert t4.rows[1].cells[1].paragraphs[0].runs[0].font.name == FONT_SONG

        # 3b. 验证表尾注（楷体小一号）与长文本单元格首行缩进 2 字符
        note_paras = [p for p in d.paragraphs if p.text.startswith('注：')]
        assert note_paras and note_paras[0].runs[0].font.name == FONT_KAI
        assert note_paras[0].runs[0].font.size == SIZE_XIAOSI
        long_cell_p = d.tables[2].rows[1].cells[5].paragraphs[0]
        assert 'firstLineChars="200"' in long_cell_p._p.xml

        # 4. 验证 JSON 块默认不在正式公文正文出现，且落款日期去前导零
        all_txt = '\n'.join(p.text for p in d.paragraphs)
        assert '"domain_a"' not in all_txt
        assert '2026年11月5日' in all_txt
        assert '授信审批委员会：' in all_txt

        print('=== export_docx.py 自检通过 ===')
        print('  [OK] A4 页面几何（37/35/28/26mm）与奇偶页 — PAGE — 页码')
        print('  [OK] 公文字体体系（2号小标宋标题 / 3号黑体 / 3号楷体 / 3号仿宋正文）')
        print('  [OK] 表格格式契约（正文仿宋 / 附件宋体 / 表首黑体重复 / 四号↔小四自适应 / 外1pt内0.5pt / 数字右对齐 / 超宽表自动横向）')
        print('  [OK] 正式公文过滤 JSON 与主送/落款日期规范化')


def main():
    parser = argparse.ArgumentParser(description='财务分析报告 Markdown 转公文标准 DOCX')
    parser.add_argument('input_md', nargs='?', help='输入 Markdown 报告路径')
    parser.add_argument('output_docx', nargs='?', help='输出 DOCX 路径')
    parser.add_argument('--template', help='官方 Word 模板路径（保留模板版式）')
    parser.add_argument('--recipient', help='主送单位名称')
    parser.add_argument('--issuer', help='落款单位名称')
    parser.add_argument('--date', help='落款日期（如 2026-11-20 或 2026年11月20日）')
    parser.add_argument('--keep-numbering', action='store_true', help='保留原始标题序号')
    parser.add_argument('--no-landscape', action='store_true', help='禁用超宽表自动横向分节')
    parser.add_argument('--no-first-page-num', action='store_true', help='首页不显示页码')
    parser.add_argument('--include-json', action='store_true', help='在文档末尾保留 Machine-Readable JSON')
    parser.add_argument('--selftest', action='store_true', help='运行内置渲染与 OOXML 自检')
    args = parser.parse_args()

    if args.selftest:
        run_selftest()
        return 0
    if not args.input_md or not args.output_docx:
        parser.print_help()
        return 1

    with io.open(args.input_md, 'r', encoding='utf-8') as f:
        md_text = f.read()

    out = render_markdown_to_docx(
        md_text,
        args.output_docx,
        template_path=args.template,
        recipient=args.recipient,
        issuer=args.issuer,
        doc_date=args.date,
        keep_numbering=args.keep_numbering,
        landscape_wide_tables=not args.no_landscape,
        no_first_page_num=args.no_first_page_num,
        include_json=args.include_json,
    )
    print('OK ->', out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
