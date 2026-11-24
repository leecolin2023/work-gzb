"""
verify_report.py — 全覆盖回归校验器

用法:
    python verify_report.py <报告.md> <metrics.json> [--src <源.xlsx>] [--out 校验记录.md]
    python verify_report.py --selftest [<报告.md> <metrics.json>]   # 校验器参数自检与表格审计

核心机制:
  1. 按展示精度比较：从报告单元格提取小数位数 d，判据为
     `报告值 == round_half_up(全精度值, d)`，动态容差 = `0.5 × 10^(-d)`。
  2. 全表计数与不变式校验：所有含数值的表格单元格一律计入分母，方向未识别的表格
     计入 `uncovered` 并逐条列出；`--selftest` 核验 `INV-1` 与 `INV-2`。
  3. 双基准面与方向覆盖：同时比对 `metrics.json` 与源表直读 `--src`，并通过 `COLREG`
     支持“期间作行”与“单期标签行”转置表。
  4. 自指统计核对：核验正文散文自述的数值点数、覆盖率与未覆盖占比是否与实测一致。

判词语义:
    FAIL          —— 存在不一致、跨章矛盾或自指统计不符，禁止交付。
    NO-MISMATCH   —— 查到的都对，但仍有未覆盖；只能写“查到的都对”并列未覆盖清单。
    FULL-PASS     —— 不一致 = 0 且 未覆盖 = 0，允许写“全部一致”。

退出码: 0 = 无不一致；1 = 存在不一致、跨章矛盾或自指统计不符。
"""
import sys, io, os, json, re
from decimal import Decimal, ROUND_HALF_UP
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)
from recalc import parse_period_header, find_header_row, find_sheet

# ─────────────────────────────────────────────────────────────────────────────
# 行标签 → 指标登记表
#   (键, [报告中的标签别名], 换算, 占位字段(未启用), 单位)
#   单位: 万元 / % / pp / x(倍数) / d(天数)
#   说明：比较时一律按报告展示小数位动态设定容差 0.5×10^(-d)，不使用第 4 列固定值。
# ─────────────────────────────────────────────────────────────────────────────
REG = [
    ('zzc',            ['资产总计', '总资产'],                        1e4, 0.02, '万元'),
    ('fz',             ['负债合计', '负债总额'],                       1e4, 0.02, '万元'),
    ('qy',             ['所有者权益合计', '所有者权益', '权益合计'],     1e4, 0.02, '万元'),
    ('ldzc',           ['流动资产'],                                  1e4, 0.02, '万元'),
    ('fldzc',          ['非流动资产'],                                1e4, 0.02, '万元'),
    ('ldfz',           ['流动负债'],                                  1e4, 0.02, '万元'),
    ('fldfz',          ['非流动负债'],                                1e4, 0.02, '万元'),
    ('yx_xy',          ['狭义有息负债', '狭义有息负债合计'],            1e4, 0.02, '万元'),
    ('yx_gy',          ['广义有息负债'],                              1e4, 0.02, '万元'),
    ('dc_yx',          ['短期有息负债'],                              1e4, 0.02, '万元'),
    ('cq_yx',          ['长期借款'],                                  1e4, 0.02, '万元'),
    ('cqyp',           ['长期应付款'],                                1e4, 0.02, '万元'),
    ('wx_sum',         ['无息负债'],                                  1e4, 0.02, '万元'),
    ('hbzj_cover',     ['货币资金覆盖率', '货币资金/短期有息负债'],        1, 0.051, '%'),
    ('zcfzl',          ['资产负债率', '账面资产负债率'],               1, 0.051, '%'),
    ('zcfzl_adj',      ['剔除虚资产后资产负债率', '调整后资产负债率'],    1, 0.051, '%'),
    ('jrfl',           ['资产金融负债率', '资产金融负债率（狭义）'],      1, 0.051, '%'),
    ('ldbl',           ['流动比率'],                                   1, 0.03, 'x'),
    ('sdbl',           ['速动比率', '速动比率（从严）'],                 1, 0.03, 'x'),
    ('yhlx',           ['已获利息倍数', '基准已获利息倍数'],             1, 0.03, 'x'),
    ('stress_yhlx',    ['情景已获利息倍数（收入-20%）'],                1, 0.03, 'x'),
    ('stress_core',    ['情景核心利润（收入-20%）'],                    1e4, 0.02, '万元'),
    ('dcb',            ['短长比'],                                     1, 0.03, 'x'),
    ('yxzbl',          ['有息负债占比', '有息负债/负债合计'],            1, 0.051, '%'),
    ('yyzb',           ['营运资本', '净营运资本', '营运资本（流动资产-流动负债）'], 1e4, 0.02, '万元'),
    ('cash_ocf_over_st', ['（货币资金+经营活动净额）/短期有息负债',
                          '（货币资金 + 经营活动净额）/ 短期有息负债'],  1, 0.03, 'x'),
    ('gap_refi',       ['续贷中断缺口', '缺口（= 短期有息 - 货币资金）',
                          '缺口（= 短期有息负债 - 货币资金）'],            1e4, 0.02, '万元'),
    ('cqzj_cover',     ['长期资金/非流动资产', '长期资金覆盖倍数'],        1, 0.03, 'x'),
    ('hxlr',           ['核心利润', '基准核心利润'],                     1e4, 0.02, '万元'),
    ('mll',            ['毛利率'],                                     1, 0.051, '%'),
    ('jll',            ['净利率'],                                     1, 0.051, '%'),
    ('hxlrl',          ['核心利润率'],                                 1, 0.051, '%'),
    ('roe',            ['ROE', '净资产收益率'],                         1, 0.051, '%'),
    ('roe_ann',        ['ROE（年化）', 'ROE 年化'],                     1, 0.051, '%'),
    ('zzczzl',         ['总资产周转率'],                               1, 0.03, 'x'),
    ('qycs',           ['权益乘数'],                                   1, 0.03, 'x'),
    ('fjy_ni',         ['非经常性损益/净利润', '非经常性损益占比'],       1, 0.051, '%'),
    ('jyxj',           ['经营活动现金流量净额', '经营净额'],             1e4, 0.02, '万元'),
    ('jxb',            ['净现比'],                                     1, 0.02, 'x'),
    ('hxlr_hxb',       ['核心利润获现比'],                             1, 0.02, 'x'),
    ('sxxjl',          ['销售收现比'],                                 1, 0.03, 'x'),
    ('fcf',            ['自由现金流'],                                 1e4, 0.02, '万元'),
    ('xjdfl',          ['现金流动负债比'],                             1, 0.03, 'x'),
    ('dio',            ['存货周转天数'],                               1, 0.51, 'd'),
    ('dso',            ['应收账款周转天数', '应收周转天数'],             1, 0.51, 'd'),
    ('dpo',            ['应付账款周转天数', '应付周转天数'],             1, 0.51, 'd'),
    ('dpp',            ['预付周转天数', '预付款项周转天数'],             1, 0.51, 'd'),
    ('dad',            ['预收周转天数', '预收款项周转天数'],             1, 0.51, 'd'),
    ('ccc',            ['现金转换周期', 'CCC', '现金转换周期（CCC）', 'CCC（现金转换周期）'], 1, 0.51, 'd'),
    ('tata',           ['TATA'],                                      1, 0.051, '%'),
    ('zscore',         ["Z'-Score", "Z'"],                            1, 0.03, 'x'),
    ('zscore_ann',     ["Z'-Score（年化）", "Z' 年化"],                 1, 0.03, 'x'),
    ('mscore',         ['M-Score'],                                   1, 0.03, 'x'),
    ('cagr4_rev',      ['营业收入 CAGR'],                              1, 0.051, '%'),
    ('cagr4_ni',       ['净利润 CAGR'],                                1, 0.051, '%'),
    ('cagr4_ta',       ['总资产 CAGR'],                                1, 0.051, '%'),
    ('cagr4_eq',       ['所有者权益 CAGR'],                            1, 0.051, '%'),
    ('cagr4_tl',       ['负债 CAGR'],                                  1, 0.051, '%'),
    ('cagr15_ta',      ['总资产 CAGR（2024-12 → 2026-06）'],           1, 0.051, '%'),
    ('cagr15_eq',      ['所有者权益 CAGR（2024-12 → 2026-06）'],       1, 0.051, '%'),
    ('cagr15_tl',      ['负债 CAGR（2024-12 → 2026-06）'],             1, 0.051, '%'),
    ('yoy_rev',        ['营业收入同比'],                               1, 0.051, '%'),
    ('yoy_ni',         ['净利润同比'],                                 1, 0.051, '%'),
    ('yoy_ta',         ['总资产同比'],                                 1, 0.051, '%'),
    ('yoy_eq',         ['所有者权益同比'],                             1, 0.051, '%'),
    ('yoy_tl',         ['负债同比'],                                   1, 0.051, '%'),
    # 应收账款同比（按展示精度动态容差比较）
    ('yoy_ar',         ['应收账款同比'],                               1, 0.60, '%'),
    ('yoy_inv',        ['存货同比'],                                   1, 0.051, '%'),
    ('yoy_prepay',     ['预付款项同比'],                               1, 0.051, '%'),
    ('yoy_wca',        ['营运资产同比'],                               1, 0.051, '%'),
    ('wca_per_rev_delta', ['营运资产增量 / 收入增量（倍）'],             1, 0.03, 'x'),
    ('core_elasticity', ['核心利润对收入的弹性（倍）'],                  1, 0.03, 'x'),
    ('yyzj_turns',     ['营运资金周转次数'],                           1, 0.03, 'x'),
    ('yyzj_amt',       ['营运资金量'],                                 1e4, 0.02, '万元'),
    ('sgr',            ['可持续增长率', 'SGR'],                         1, 0.051, '%'),
    # ---- 报表结构、费用与现金流扩展指标 ----
    ('qjfy',           ['期间费用合计'],                               1e4, 0.02, '万元'),
    ('wca',            ['营运资产合计', '营运资产'],                     1e4, 0.02, '万元'),
    ('wca_ratio',      ['营运资产占总资产'],                            1, 0.051, '%'),
    ('wca_less_cl',    ['营运资产-流动负债', '营运资产扣除流动负债'],      1e4, 0.02, '万元'),
    ('ca_ratio',       ['流动资产占比'],                               1, 0.051, '%'),
    ('nca_ratio',      ['非流动资产占比'],                             1, 0.051, '%'),
    ('non_op_net',     ['营业外收支净额', '其他（营业外收支净额）'],       1e4, 0.02, '万元'),
    ('gy_tzsy',        ['广义投资收益（利息+投资+公允价值）', '广义投资收益'],
                                                                       1e4, 0.02, '万元'),
    ('ar_imp5_pct',    ['5%计提减值/净利润', '应收5%减值/净利润'],        1, 0.051, '%'),
    ('cum_divint5',    ['5年累计偿付利息现金', '累计偿付利息现金'],        1e4, 0.02, '万元'),
    ('cip_over_fa',    ['在建工程/固定资产（倍）', '在建/固定资产'],       1, 0.05, 'x'),
    ('wx_deduct',      ['倒减残差', '倒减口径无息负债'],                  1e4, 0.02, '万元'),
    ('wx_gap',         ['加总与倒减差额'],                              1e4, 0.02, '万元'),
    ('eq_change',      ['权益变动'],                                   1e4, 0.02, '万元'),
    ('cap_add',        ['实收资本增'],                                 1e4, 0.02, '万元'),
    ('res_add',        ['专项储备增'],                                 1e4, 0.02, '万元'),
    ('eq_roll_sum',    ['权益变动合计'],                               1e4, 0.02, '万元'),
    ('eq_roll_diff',   ['权益变动差额'],                               1e4, 0.02, '万元'),
    ('dc_yx_ratio',    ['短期有息负债/狭义有息负债'],                     1, 0.051, '%'),
    ('st_over_nca',    ['短期有息负债/非流动资产'],                       1, 0.03, 'x'),
    ('ocf_over_int',   ['经营活动净额/偿付利息现金', '经营活动净额 / 偿付利息现金'],
                                                                       1, 0.03, 'x'),
    ('hbzj',           ['货币资金（未扣受限）'],                          1e4, 0.02, '万元'),
    ('tzxj',           ['投资活动现金流量净额'],                          1e4, 0.02, '万元'),
    ('czxj',           ['筹资活动现金流量净额'],                          1e4, 0.02, '万元'),
    ('cf_fx',          ['汇率变动影响'],                                 1e4, 0.02, '万元'),
    ('cf_cash_net',    ['现金及现金等价物净增加'],                        1e4, 0.02, '万元'),
    ('growth_gap',     ['增长缺口'],                                    1, 0.051, 'pp'),
    ('sgr',            ['SGR = ROE × 留存收益率', 'SGR=ROE×留存收益率'],  1, 0.051, '%'),
    ('retention',      ['留存收益率'],                                    1, 0.051, '%'),
    ('yx_xy_yoy',      ['狭义有息负债同比'],                              1, 0.051, '%'),
    ('yx_gy_yoy',      ['广义有息负债同比'],                              1, 0.051, '%'),
    ('st_less_wc',     ['短期有息负债-营运资本'],                          1e4, 0.02, '万元'),
    ('eq_in',          ['吸收投资收到的现金'],                             1e4, 0.02, '万元'),
]
# 额外的显式别名（语义上下文化，避免"去括号后撞名"而被迫放弃匹配）
EXTRA = {
    'roe': ['ROE（平均权益口径）'],
    'yoy_ta': ['实际总资产增长率'],
    'dc_yx': ['其中：短期有息负债'],
    'cqyp': ['长期应付款（性质待核）'],
    'cash_ocf_over_st': ['（货币资金 + 经营净额）/ 短期有息负债',
                         '（货币资金+经营净额）/短期有息负债'],
}
# 行向转置表与严格口径指标的换算与单位（用于 COLREG 映射）
NEWKIND = {
    'ni':                 ('万元', 1e4),
    'adj_ta':             ('万元', 1e4),
    'lt_funds':           ('万元', 1e4),
    'lt_funds_surplus':   ('万元', 1e4),
    'adj_neg_ltd':        ('万元', 1e4),
    'adj_neg_cip':        ('万元', 1e4),
    'prepay_over_ta':     ('%',    1),
    'ar_over_rev':        ('%',    1),
    'cip_over_ta':        ('%',    1),
    'lt_funds_over_nca':  ('x',    1),
    'prepay_over_ap':     ('x',    1),
    'ocf_over_divint5':   ('x',    1),
    'adj_up':             ('pp',   1),
    'ar_rev_gap5':        ('pp',   1),
    'rd_blank':           ('x',    1),
    # 非全年期变动与增量键
    'wca_delta':          ('万元', 1e4),
    'prepay_chg5':        ('%',    1),
    'ar_chg5':            ('%',    1),
    'yhlx_strict':        ('x',    1),
    'hxlr_strict':        ('万元', 1e4),
}
SKIP = ('', '—', '-', '–', '数据缺失', '未计算', '无法计算', 'N/A')
# ─────────────────────────────────────────────────────────────────────────────
# COLREG —— 按「表头签名」登记表格的列 → 指标映射
#   支持两类非常规方向表格：
#     ① 期间作行的转置表（第一列是期间，其余列是指标）
#     ② 单期表（整表只有一个期间，期间信息在表名/语境里，不在单元格里）
#   签名 = 表头各单元格去掉 ** 、` 与所有空白后，用 | 连接。
#   列规格：None = 该列是标签/文本列，跳过；'src:科目名' = 直读源表；
#           其他字符串 = metrics.json 的键。
# ─────────────────────────────────────────────────────────────────────────────
COLREG = {
    # 表9 权益变动滚动（期间作行）
    '期间|权益变动|本期净利润|实收资本增|专项储备增|合计|差额':
        dict(mode='row', cols=[None, 'eq_change', 'ni', 'cap_add', 'res_add',
                               'eq_roll_sum', 'eq_roll_diff']),
    # 表14 预付款项。末列 2026-06 是**半年变动**（vs 2025 年末），不是同比 —— 用 ALT 分流。
    '期间|预付款项|占总资产|同比|预付周转天数':
        dict(mode='row', cols=[None, 'src:预付款项', 'prepay_over_ta',
                               ('ALT', 'yoy_prepay', 'prepay_chg5'), 'dpp']),
    # 表15 应收账款
    '期间|应收账款|占营业收入|应收周转天数|同比|同期收入同比':
        dict(mode='row', cols=[None, 'src:应收账款', 'ar_over_rev', 'dso',
                               ('ALT', 'yoy_ar', 'ar_chg5'), 'yoy_rev']),
    # 表16 在建工程
    '期间|在建工程|占总资产|固定资产|在建/固定资产':
        dict(mode='row', cols=[None, 'src:在建工程', 'cip_over_ta',
                               'src:固定资产', 'cip_over_fa']),
    # 表19 无息负债双口径
    '期间|加总口径无息负债|倒减残差（负债合计-狭义有息）|两者差额|差额恰等于':
        dict(mode='row', cols=[None, 'wx_sum', 'wx_deduct', 'wx_gap', None]),
    # 表21 剔除虚资产后的资产负债率
    '期间|资产总计|-长期待摊费用|-50%×在建工程|=调整后资产|负债合计|调整后资产负债率|较①上升':
        dict(mode='row', cols=[None, 'src:资产总计', 'adj_neg_ltd', 'adj_neg_cip', 'adj_ta',
                               'src:负债合计', 'zcfzl_adj', 'adj_up']),
    # 表23 长期资金
    '期间|所有者权益|非流动负债|长期资金|非流动资产|长期资金/非流动资产|富余（=营运资本）':
        dict(mode='row', cols=[None, 'src:所有者权益合计', 'src:非流动负债合计', 'lt_funds',
                               'src:非流动资产合计', 'lt_funds_over_nca', 'lt_funds_surplus']),
    # 表33 营运资产占用
    '期间|营运资产合计（应收+预付+存货）|占总资产|同比增量|同期经营活动净额':
        dict(mode='row', cols=[None, 'wca', 'wca_ratio', 'wca_delta',
                               'src:经营活动产生的现金流量净额']),
    # 表47 应收与收入背离（单期 2025 + **按行标签选指标**）
    '比对项|2025年同比|收入同比|背离幅度|判断':
        dict(mode='labelrow', period=4,
             cols=[None, '#ROW', 'yoy_rev', ('GAP', 'yoy_rev'), None],
             rowmap={'应收账款同比': 'yoy_ar', '存货同比': 'yoy_inv',
                     '预付款项同比': 'yoy_prepay', '总资产同比': 'yoy_ta',
                     '净利润同比': 'yoy_ni'}),
}

MINUS = ('\u2212', '\u2013', '\u2014', '\uff0d', '\ufe63', '\u2010', '\u2011')
NUM_RE = re.compile(r'-?\d+(?:\.\d+)?')
PERIOD_RE = re.compile(r'(20\d\d)\s*(?:年)?\s*(?:[-/]\s*(\d{1,2}))?')


def norm(s):
    for ch in MINUS:
        s = s.replace(ch, '-')
    return s


SRC_STRIP = [
    r'^[一二三四五六七八九十]+[、.．]\s*',
    r'^[（(]\s*[一二三四五六七八九十]+\s*[)）]\s*',
    r'^(减|加|其中)\s*[：:]\s*',
    r'^其中\s*',
]


def strip_all(s):
    outs, cur, changed = [s], s, True
    while changed:
        changed = False
        for p in SRC_STRIP:
            new = re.sub(p, '', cur)
            if new != cur:
                cur = new
                outs.append(cur)
                changed = True
    return outs


def build_alias(reg):
    exact, cands = {}, {}
    for k, names, sc, tol, kind in reg:
        for n in names:
            exact.setdefault(norm(n).strip().lower(), (k, sc, tol, kind))
            exact.setdefault(norm(n).strip().lower().replace(' ', ''), (k, sc, tol, kind))
            base = re.sub(r'[（(][^）)]*[)）]\s*$', '', norm(n)).strip().lower()
            if base and base != norm(n).strip().lower():
                cands.setdefault(base, set()).add(k)
    for base, ks in cands.items():
        if len(ks) != 1:
            continue
        k = list(ks)[0]
        cur = exact.get(base)
        if cur is None or cur[0] == k:
            for k2, names, sc, tol, kind in reg:
                if k2 == k:
                    exact[base] = (k, sc, tol, kind)
                    break
    return exact


ALIAS = build_alias(REG)
_KMETA = {k: (sc, tol, kind) for k, n, sc, tol, kind in REG}
for _k, _names in EXTRA.items():
    if _k not in _KMETA:
        continue
    _sc, _tol, _kind = _KMETA[_k]
    for _n in _names:
        ALIAS.setdefault(norm(_n).strip().lower(), (_k, _sc, _tol, _kind))
        ALIAS.setdefault(norm(_n).strip().lower().replace(' ', ''), (_k, _sc, _tol, _kind))


def scale_kind(key):
    """指标键 → (换算, 单位)。优先 metrics 登记表，其次 NEWKIND。"""
    if key in _KMETA:
        sc, _tol, kind = _KMETA[key]
        return kind, sc
    if key in NEWKIND:
        kind, sc = NEWKIND[key]
        return kind, sc
    return '?', 1


def to_exact(key, raw):
    """metrics 原值 → 报告单位的全精度值。"""
    kind, sc = scale_kind(key)
    if kind == '万元':
        return raw / sc
    if kind in ('%', 'pp'):
        return raw * 100.0
    return raw


def period_index(cell, periods=None):
    """动态匹配期间索引（对齐 metrics.json['_periods']）。"""
    s = cell.replace('**', '').replace('`', '').strip()
    if not s or not periods:
        return None
    parsed = parse_period_header(s)
    if parsed is not None:
        lab, y, mo, kind = parsed
        if lab in periods:
            return periods.index(lab)
        # 若单元格仅含年份（如 '2026' 或 '2026年'），且 periods 中该年份仅有一期（如 '2026-06'），唯一匹配
        same_year = [idx for idx, p in enumerate(periods) if p.startswith('%04d-' % y)]
        if len(same_year) == 1 and not re.search(r'(?:[-/.年]\s*\d{1,2}|H[12]|Q[1-4]|季|半年|中报)', s, re.I):
            return same_year[0]
    return None


def cells(line):
    return [c.strip() for c in line.strip().strip('|').split('|')]


def clean_label(s):
    s = s.replace('**', '').replace('`', '').replace('*', '')
    s = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\d]+[、.．:：]?\s*', '', s)
    return norm(s).strip().strip('：:')


def signature(head):
    """生成表头签名（自动归一化各类负号与空白字符）。"""
    return '|'.join(norm(re.sub(r'\*\*|`|\s', '', c)) for c in head)


def decimals_of(raw):
    """报告单元格里那个数字的小数位数。"""
    s = raw.replace('**', '').replace(',', '')
    m = re.search(r'-?\d+\.(\d+)', s)
    return len(m.group(1)) if m else 0


def rh(x, d):
    if x is None:
        return None
    return float(Decimal(repr(float(x))).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP))


def parse_one(raw):
    """单元格 → 单一数字；不是单值就返回 None。"""
    txt = norm(raw.replace('**', '').replace(',', '').replace('%', ''))
    nums = NUM_RE.findall(txt)
    if len(nums) != 1:
        return None
    t = re.sub(r'[^\d.\-]', '', txt)
    if not t or t in ('.', '-'):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def load_src(path, periods):
    """源 xlsx → "科目 → 动态期间值（元）" 索引，按 periods 精确对齐，作第二基准面。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    p_map = {p: idx for idx, p in enumerate(periods)}
    n_p = len(periods)
    d = {}
    for cands in (['资产负债表'], ['利润表', '损益表'], ['现金流量表']):
        sh = find_sheet(wb, cands)
        if not sh:
            continue
        ws = wb[sh]
        hr, cols = find_header_row(ws)
        if hr is None or not cols:
            continue
        for r in range(hr + 1, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if not isinstance(a, str):
                continue
            a = a.strip()
            if not a or a in d:
                continue
            row = [None] * n_p
            for c, p in cols.items():
                if p[0] in p_map:
                    v = ws.cell(r, c).value
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        row[p_map[p[0]]] = float(v)
            d[a] = row
    idx = {}
    for lab, row in d.items():
        for v in strip_all(lab):
            idx.setdefault(norm(v).strip().lower(), row)
            idx.setdefault(norm(v).strip().lower().replace(' ', ''), row)
    return d, idx


def selftest_params():
    """校验器参数自检：打印动态容差表、登记规模与核心不变式。"""
    print('=== verify_report.py 自检 · 判据与参数 ===')
    print('判据：报告展示值 == round_half_up(全精度值, 展示小数位)')
    print('     容差 = 0.5 × 10^(−展示小数位)   —— 不使用任何固定绝对容差')
    print('     （按展示小数位动态约束，确保 0.1pp 级舍入错误必然暴露）')
    print()
    print('%-8s %-14s %s' % ('展示小数位', '容差', '示例：全精度 20.750656% 的合法展示值'))
    for d in range(0, 5):
        print('%-8d %-14s %.*f' % (d, ('%g' % (0.5 * 10 ** (-d))), d, rh(20.750656, d)))
    print()
    print('登记规模：REG=%d  NEWKIND=%d  COLREG=%d  SKIP=%d'
          % (len(REG), len(NEWKIND), len(COLREG), len(SKIP)))
    modes = {}
    for v in COLREG.values():
        k = '%s' % v['mode']
        modes[k] = modes.get(k, 0) + 1
    print('COLREG 方向分布：%s' % ('  '.join('%s=%d' % kv for kv in sorted(modes.items())) or '—'))
    print()
    print('不变式：')
    print('  INV-1 每张含数值的表格都必须为覆盖率分母贡献 >0 个单元格 —— 防「方向盲区」复发')
    print('  INV-2 分子分母自洽：verified + failed + uncovered == ncell')


def selftest_tables(blocks, recs, uncovered, ncell, periods):
    """表格级审计：独立扫描全部表格，确保每张含数值的表都进入覆盖率分母（核验 INV-1）。"""
    rc, uc = {}, {}
    for r in recs:
        rc[r[5]] = rc.get(r[5], 0) + 1
    for u in uncovered:
        uc[u[3]] = uc.get(u[3], 0) + 1

    rows, warn = [], []
    for bi, blk in enumerate(blocks):
        where = '表%d' % (bi + 1)
        head = cells(blk[0])
        spec = COLREG.get(signature(head))
        if spec is not None:
            orient = 'COLREG:%s' % spec['mode']
        else:
            cm = set(period_index(h, periods) for h in head)
            cm.discard(None)
            orient = '期间作列' if cm else '**方向未识别**'
        has_num = 0                       # 独立重数：本表是否含可解析的单值数字
        for ln in blk[2:]:
            cs = cells(ln)
            for ci in range(1, len(cs)):
                raw = cs[ci].strip()
                if raw in SKIP:
                    continue
                if parse_one(raw) is not None:
                    has_num += 1
        got = rc.get(where, 0) + uc.get(where, 0)
        rows.append((where, orient, has_num, got, rc.get(where, 0), uc.get(where, 0)))
        if has_num > 0 and got == 0:
            warn.append(where)

    print('=== 表格级审计（共 %d 张表）===' % len(blocks))
    print('%-6s %-18s %8s %8s %8s %8s' % ('位置', '方向', '含数字', '进分母', '已验证', '未覆盖'))
    for w in rows:
        flag = '   <== 整表被跳过！' if (w[2] > 0 and w[3] == 0) else ''
        print('%-6s %-18s %8d %8d %8d %8d%s' % (w[0], w[1], w[2], w[3], w[4], w[5], flag))
    print()
    print('=== 不变式检查 ===')
    if warn:
        print('  x INV-1 失败：以下表格含可解析数字，却没有任何单元格进入覆盖率分母 -> %s'
              % ', '.join(warn))
        print('    存在方向盲区：分母被缩小、覆盖率不可信。')
        return False
    print('  v INV-1 通过：所有含数值的表格都为分母贡献了单元格。')
    return True


COV_ENUM = {
    'scan':      r'扫描[^0-9]{0,10}([0-9][0-9,]*)\s*个',
    'cells':     r'数值点\s*([0-9][0-9,]*)\s*个',
    'verified':  r'已验证\s*([0-9][0-9,]*)\s*个',
    'fail':      r'不一致\s*([0-9][0-9,]*)\s*个',
    'uncovered': r'未覆盖\s*([0-9][0-9,]*)\s*个',
    'coverage':  r'覆盖率\s*([0-9]{1,3}(?:\.[0-9]+)?)\s*%',
}
# 仅跳过明确中文历史引述，避免误伤正文正常覆盖率声明。
# 引用历史数值的表格行由 `body.startswith('|')` 拦截，无需在此列版本号。
HIST_MARK = ('上一版', '曾称', '更正史', '当时', '曾发布')
VAL_CTX = ('校验', '数值点', '已验证', '未覆盖', '分母', '样本')
_PCT = r'([0-9]{1,3}(?:\.[0-9]+)?)\s*%'


def _num(s):
    return float(s.replace(',', ''))


def check_declared_stats(lines, cells, npass, nfail, unc, cov, n_scan=None):
    """自指统计核对：校验报告正文对自身校验覆盖情况的描述是否与实测一致。

    通过四条精确规则拦截正文自指矛盾并判 FAIL：

      R1 显式等式   `覆盖率 = A ÷ B = C%` → 校验 C 的算术，且 A/B 须等于实测分子分母
      R2 枚举式声明 同一行出现 ≥2 个自指统计量 → 逐个与实测比对
      R3 覆盖到/已覆盖 邻近的百分比 → 须等于实测覆盖率
      R4 未覆盖 邻近的百分比 → 须等于实测未覆盖占比

    **已知边界**：「覆盖率」一词同时被"财务覆盖倍数"（如"货币资金覆盖率"）使用，
    因此 R2 把「覆盖率 N%」当作标记时要求该行含校验语境词；表格行与历史引述一律跳过。
    """
    unc_pct = round(unc / cells * 100, 1) if cells else 0.0
    truth = {'cells': float(cells), 'verified': float(npass),
             'fail': float(nfail), 'uncovered': float(unc), 'coverage': round(cov, 1)}
    if n_scan is not None:
        truth['scan'] = float(n_scan)
    bad = []
    for i, ln in enumerate(lines, 1):
        s = ln.strip()
        body = s.lstrip('>').strip()
        if not body or body.startswith('|') or body.startswith('#') or body.startswith('```'):
            continue
        if any(m in s for m in HIST_MARK):
            continue
        sn = s.replace('**', '')          # 先剥离粗体标记，否则 `**1,366** 个` 会打断匹配

        # R1 显式等式
        for m in re.finditer(r'覆盖率\s*=\s*([0-9][0-9,]*)\s*÷\s*([0-9][0-9,]*)\s*=\s*' + _PCT, sn):
            a, b, c = _num(m.group(1)), _num(m.group(2)), _num(m.group(3))
            if b <= 0 or abs(round(a / b * 100, 1) - c) > 0.051:
                bad.append((i, c, '等式本身不自洽：%s ÷ %s ≠ %s%%' % (m.group(1), m.group(2), m.group(3))))
            if b != cells or a != npass:
                bad.append((i, c, '等式用的分子分母与实测不符（实测 %d ÷ %d）' % (npass, cells)))

        # R2 枚举式声明
        found = {}
        for k, pat in COV_ENUM.items():
            mm = re.search(pat, sn)
            if mm:
                found[k] = _num(mm.group(1))
        if 'coverage' in found and not any(t in sn for t in VAL_CTX):
            del found['coverage']          # 避开"货币资金覆盖率"这类财务口径
        if len(found) >= 2:
            for k, v in found.items():
                tv = truth[k]
                if abs(v - tv) > (0.051 if k == 'coverage' else 0.5):
                    bad.append((i, v, '声明的「%s」与实测不符（实测 %g）' % (k, tv)))

        # R3 覆盖到 / 已覆盖
        for m in re.finditer(r'(?:覆盖到|已覆盖)[^0-9%]{0,14}' + _PCT, sn):
            v = _num(m.group(1))
            if abs(v - round(cov, 1)) > 0.051:
                bad.append((i, v, '「覆盖到」的百分比与实测覆盖率不符（实测 %.1f%%）' % cov))

        # R4 未覆盖 邻近占比
        for m in re.finditer(r'(?:' + _PCT + r'[^0-9%]{0,6}未覆盖|未覆盖[^0-9%]{0,6}' + _PCT + r')', sn):
            v = _num(m.group(1) or m.group(2))
            if abs(v - unc_pct) > 0.051:
                bad.append((i, v, '「未覆盖」占比与实测不符（实测 %.1f%%）' % unc_pct))

    return unc_pct, bad


def main():
    selftest = '--selftest' in sys.argv
    argv = [a for a in sys.argv[1:] if a != '--selftest']
    if selftest:
        selftest_params()
    if len(argv) < 2:
        if selftest:
            print()
            print('（未提供报告与 metrics.json —— 仅输出参数自检；'
                  '带上报告可追加「表格级审计」与不变式检查。）')
            return
        print(__doc__)
        return

    rep, mj = argv[0], argv[1]
    out = srcpath = None
    if '--out' in sys.argv:
        out = sys.argv[sys.argv.index('--out') + 1]
    if '--src' in sys.argv:
        srcpath = sys.argv[sys.argv.index('--src') + 1]

    M = json.load(io.open(mj, encoding='utf-8'))
    periods = M['_periods']
    period_kinds = M.get('_period_kinds', [parse_period_header(p)[3] if parse_period_header(p) else 'FY' for p in periods])
    lines = io.open(rep, encoding='utf-8').read().split('\n')

    SRCD, SRCIDX = ({}, {})
    if srcpath:
        SRCD, SRCIDX = load_src(srcpath, periods)

    blocks, cur = [], []
    for ln in lines:
        if ln.strip().startswith('|'):
            cur.append(ln)
        else:
            if len(cur) >= 3:
                blocks.append(cur)
            cur = []
    if len(cur) >= 3:
        blocks.append(cur)

    recs = []        # (指标, 期间, 报告值, 正确展示值, 通过?, 位置, 基准面)
    uncovered = []   # (标签, 期间, 报告值, 位置)
    multi = []       # 单元格含多个数
    prose = []       # 正文散文里的数值点（不在机器可回代范围）
    ncell = 0        # 全部表格中的单一数字单元格
    n_scan = 0       # 全部表格中的所有单元格

    def verify(key, pi, raw, got, where, label):
        """核心判据：报告值必须等于全精度值按展示位数正确舍入后的值。

        recs 第 8 位是**跨章比对的去重键**（指标键），刻意区别于展示标签——
        不同表里的「占总资产」「同比」是不同指标，按标签分组会造出成片误报。
        """
        d = decimals_of(raw)
        if key.startswith('src:'):
            lab = key[4:]
            row = SRCD.get(lab)
            if row is None:
                for v in strip_all(lab):
                    row = SRCIDX.get(norm(v).strip().lower())
                    if row is not None:
                        break
            if row is None:
                uncovered.append((label + ' [源表科目未找到:%s]' % lab, periods[pi], raw, where))
                return
            v = row[pi] if pi is not None else None
            if v is None:
                recs.append((label + ' <源表留空>', periods[pi], got, None, False, where, 'src',
                             'src:' + lab))
                return
            exp_full = v / 1e4
            base, dkey = 'src', 'src:' + lab
        else:
            if key not in M:
                uncovered.append((label + ' [键缺失:%s]' % key, periods[pi], raw, where))
                return
            v = M[key][pi] if pi is not None else None
            if v is None or isinstance(v, str):
                recs.append((label + ' <基准为空>', periods[pi], got, None, False, where, 'metrics', key))
                return
            exp_full = to_exact(key, v)
            base, dkey = 'metrics', key
        exp = rh(exp_full, d)
        tol = 0.5 * 10 ** (-d) + 1e-9
        recs.append((label, periods[pi] if pi is not None else '—', got, exp,
                     abs(got - exp) <= tol, where, base, dkey))

    for bi, blk in enumerate(blocks):
        where = '表%d' % (bi + 1)
        head = cells(blk[0])
        sig = signature(head)
        spec = COLREG.get(sig)
        if spec is None:
            m_lr = re.fullmatch(r'比对项\|(20\d\d)年同比\|收入同比\|背离幅度\|判断', sig)
            if m_lr:
                dyn_pi = period_index(m_lr.group(1) + '-12', periods)
                if dyn_pi is not None:
                    spec = dict(mode='labelrow', period=dyn_pi,
                                cols=[None, '#ROW', 'yoy_rev', ('GAP', 'yoy_rev'), None],
                                rowmap={'应收账款同比': 'yoy_ar', '存货同比': 'yoy_inv',
                                        '预付款项同比': 'yoy_prepay', '总资产同比': 'yoy_ta',
                                        '净利润同比': 'yoy_ni'})

        if spec is not None:
            mode, cols = spec['mode'], spec['cols']
            fixed_pi = spec.get('period')
            if sig == '比对项|2025年同比|收入同比|背离幅度|判断':
                fixed_pi = period_index('2025-12', periods)
            rowmap = spec.get('rowmap', {})
            for ln in blk[2:]:
                cs = cells(ln)
                if not cs:
                    continue
                row_lab = clean_label(cs[0])
                if mode == 'row':
                    pi = period_index(cs[0], periods)
                    rowkey = None
                else:
                    pi = fixed_pi
                    rowkey = rowmap.get(row_lab)
                    if rowkey is None:                      # 容忍标点/空格差异
                        for rv, rk in rowmap.items():
                            if norm(rv) == norm(row_lab):
                                rowkey = rk
                                break
                for ci in range(1, len(cs)):
                    raw = cs[ci].strip()
                    n_scan += 1
                    if raw in SKIP:
                        continue
                    hdr = head[ci] if ci < len(head) else row_lab
                    got = parse_one(raw)
                    if got is None:
                        if len(NUM_RE.findall(norm(raw.replace('**', '').replace(',', '')))) > 1:
                            multi.append((row_lab, periods[pi] if pi is not None else '—', raw, where))
                        continue
                    ncell += 1
                    key = cols[ci] if ci < len(cols) else None
                    if mode == 'labelrow':
                        if key == '#ROW':
                            key = rowkey
                        elif isinstance(key, tuple) and key[0] == 'GAP':
                            key = ('GAPKEY', rowkey, key[1])
                    if key is None:
                        uncovered.append(('%s / %s' % (row_lab, hdr),
                                          periods[pi] if pi is not None else '—', raw, where))
                    elif pi is None:
                        uncovered.append(('%s [期间未能识别]' % hdr, '—', raw, where))
                    elif isinstance(key, tuple) and key[0] == 'ALT':
                        verify(key[2] if (pi is not None and period_kinds[pi] != 'FY') else key[1], pi, raw, got, where, hdr)
                    elif isinstance(key, tuple) and key[0] == 'GAPKEY':
                        rk, bk = key[1], key[2]
                        if rk is None or rk not in M or bk not in M:
                            uncovered.append(('%s [行标签未登记:%s]' % (row_lab, rk), periods[pi], raw, where))
                            continue
                        rv, bv = M[rk][pi], M[bk][pi]
                        if rv is None or bv is None:
                            uncovered.append((hdr, periods[pi], raw, where))
                            continue
                        exp_full = to_exact(rk, rv) - to_exact(bk, bv)
                        d = decimals_of(raw)
                        exp = rh(exp_full, d)
                        recs.append((hdr + '（%s − %s）' % (rk, bk), periods[pi], got, exp,
                                     abs(got - exp) <= 0.5 * 10 ** (-d) + 1e-9, where, 'metrics',
                                     'GAP:%s-%s' % (rk, bk)))
                    else:
                        verify(key, pi, raw, got, where, hdr)
            continue

        # 期间作列的标准表格路径
        cmap, seen = {}, set()
        for ci, h in enumerate(head):
            pi = period_index(h, periods)
            if pi is not None and pi not in seen:
                cmap[ci] = pi
                seen.add(pi)

        if not cmap:
            # 方向识别不出 → 全部计入 uncovered，**不再静默跳过**
            for ln in blk[2:]:
                cs = cells(ln)
                if not cs:
                    continue
                lab = clean_label(cs[0])
                for ci in range(1, len(cs)):
                    raw = cs[ci].strip()
                    n_scan += 1
                    if raw in ('', '—', '-', '–', '数据缺失', '未计算', '无法计算'):
                        continue
                    got = parse_one(raw)
                    if got is None:
                        if len(NUM_RE.findall(norm(raw.replace('**', '').replace(',', '')))) > 1:
                            multi.append((lab, '—', raw, where))
                        continue
                    ncell += 1
                    uncovered.append((lab + ' [表无期间列，方向未识别]', '—', raw, where))
            continue

        for ln in blk[2:]:
            cs = cells(ln)
            if not cs:
                continue
            lab = clean_label(cs[0])
            if not lab:
                continue
            key = ALIAS.get(lab.lower()) or ALIAS.get(lab.lower().replace(' ', ''))
            if key is None:
                base_lab = re.sub(r'[（(][^）)]*[)）]\s*$', '', lab).strip()
                if base_lab and base_lab != lab:
                    key = ALIAS.get(base_lab.lower()) or ALIAS.get(base_lab.lower().replace(' ', ''))
            srcrow, srclab = None, None
            if key is None and SRCIDX:
                for v in strip_all(lab):
                    hit = SRCIDX.get(norm(v).strip().lower()) or \
                          SRCIDX.get(norm(v).strip().lower().replace(' ', ''))
                    if hit is not None:
                        srcrow, srclab = hit, v
                        break
            for ci, pi in cmap.items():
                if ci >= len(cs):
                    continue
                raw = cs[ci].strip()
                n_scan += 1
                if raw in ('', '—', '-', '–', '数据缺失', '未计算', '无法计算'):
                    continue
                got = parse_one(raw)
                if got is None:
                    if len(NUM_RE.findall(norm(raw.replace('**', '').replace(',', '')))) > 1:
                        multi.append((lab, periods[pi], raw, where))
                    continue
                ncell += 1
                k = key[0] if key is not None else None
                if k is not None:
                    if k.startswith('_') or k not in M:
                        uncovered.append((lab + ' [键缺失:%s]' % k, periods[pi], raw, where))
                        continue
                    v = M[k][pi]
                    if v is None or isinstance(v, str):
                        recs.append((lab + ' <基准为空>', periods[pi], got, None, False, where,
                                     'metrics', k))
                        continue
                    exp_full = to_exact(k, v)
                    d = decimals_of(raw)
                    exp = rh(exp_full, d)
                    recs.append((lab, periods[pi], got, exp,
                                 abs(got - exp) <= 0.5 * 10 ** (-d) + 1e-9, where, 'metrics', k))
                elif srcrow is not None:
                    v = srcrow[pi]
                    if v is None:
                        recs.append((srclab + ' <源表留空>', periods[pi], got, None, False, where,
                                     'src', 'src:' + srclab))
                        continue
                    exp_full = v / 1e4
                    d = decimals_of(raw)
                    exp = rh(exp_full, d)
                    recs.append((srclab + ' [源表]', periods[pi], got, exp,
                                 abs(got - exp) <= 0.5 * 10 ** (-d) + 1e-9, where, 'src',
                                 'src:' + srclab))
                else:
                    uncovered.append((lab, periods[pi], raw, where))

    # 正文散文里的数值点：不在表格中，机器不可回代。只统计规模并列为盲区。
    for ln in lines:
        if ln.strip().startswith('|'):
            continue
        if not ln.strip() or ln.strip().startswith('#'):
            continue
        cand = ln.replace('**', '').replace(',', '')
        for m in re.finditer(r'\d+\.\d+', cand):
            prose.append((ln.strip()[:60], m.group(0)))

    npass = sum(1 for r in recs if r[4])
    nfail = len(recs) - npass
    mism = [r for r in recs if not r[4]]
    cov = npass / ncell * 100 if ncell else 0.0

    if selftest:
        print()
        selftest_tables(blocks, recs, uncovered, ncell, periods)
        ok2 = (npass + nfail + len(uncovered) == ncell)
        print('  %s INV-2 %s：verified(%d) + failed(%d) + uncovered(%d) == ncell(%d)'
              % ('v' if ok2 else 'x', '通过' if ok2 else '失败',
                 npass, nfail, len(uncovered), ncell))
        if not ok2:
            print('    说明有单元格既未验证也未登记为未覆盖 —— 覆盖率不可信。')
        print()

    byk = {}
    for r in recs:
        byk.setdefault((r[7], r[1]), []).append((r[2], r[5]))
    dup = {kp: v for kp, v in byk.items() if len(v) > 1}
    selfbad = []
    for (k, p), v in dup.items():
        vals = [x[0] for x in v]
        if max(vals) - min(vals) > 0.055:
            selfbad.append((k, p, v))

    unc_pct, decl_bad = check_declared_stats(lines, ncell, npass, nfail,
                                                len(uncovered), cov, n_scan)

    L = ['# 回归校验记录（报告 ↔ 复算结果全覆盖回代）\n']
    L.append('- 报告：`%s`' % os.path.basename(rep))
    L.append('- 基准①：`metrics.json`（脚本复算的派生指标，含行向表列头基准）')
    L.append('- 基准②：`%s`（源表科目逐行直读）' % (os.path.basename(srcpath) if srcpath else '未提供'))
    L.append('')
    L.append('> **判据**：报告展示值必须等于「全精度值按报告展示位数四舍五入」后的值，')
    L.append('> 容差 = 半个末位单位（0.5×10^(-d)），按展示精度动态设定。')
    L.append('')
    L.append('| 口径 | 数量 |')
    L.append('|---|---|')
    L.append('| 表格中的单元格总数（扫描） | %d |' % n_scan)
    L.append('| 其中「可解析为单一数字」的数值点（**覆盖率的分母**） | **%d** |' % ncell)
    L.append('| ✔ 已验证（有基准且一致） | **%d** |' % npass)
    L.append('| ✘ **不一致** | **%d** |' % nfail)
    L.append('| ? **未覆盖（有数字但无基准可比）** | **%d** |' % len(uncovered))
    L.append('| 单元格含多个数（未纳入单值回代） | %d |' % len(multi))
    L.append('| **覆盖率 = 已验证 / 数值点** | **%.1f%%** |' % cov)
    L.append('')
    L.append('| 机器覆盖不到的正文散文数值点（`d.d` 形态） | %d |' % len(prose))
    L.append('')
    L.append('**"未覆盖" 与 "不一致" 是两件事**：未覆盖 = 这个数字没有基准可比，')
    L.append('不等于它错；但在它被覆盖之前，**不得写"全部一致"**。')
    L.append('')

    if mism:
        L.append('## 一、不一致明细\n')
        L.append('| 指标 | 期间 | 报告值 | 正确展示值 | 差异 | 位置 | 基准 |')
        L.append('|---|---|---|---|---|---|---|')
        for r in mism:
            k, p, got, exp, _, where, base = r[0], r[1], r[2], r[3], r[4], r[5], r[6]
            L.append('| %s | %s | %s | %s | %s | %s | %s |' % (
                k, p, fmt(got), fmt(exp), fmt(got - exp) if exp is not None else '基准留空', where, base))
        L.append('')
    else:
        L.append('## 一、不一致明细\n\n无。\n')

    L.append('## 二、未覆盖清单（有数字、无基准可比 → 必须人工确认）\n')
    if uncovered:
        L.append('| 行标签 | 期间 | 报告值 | 位置 |')
        L.append('|---|---|---|---|')
        for lab, p, raw, where in uncovered[:120]:
            L.append('| %s | %s | %s | %s |' % (lab, p, raw, where))
        if len(uncovered) > 120:
            L.append('| … | … | 共 %d 处 | |' % len(uncovered))
        L.append('')
    else:
        L.append('无。每一个数值点都有基准可比。\n')

    if multi:
        L.append('## 三、含多个数的单元格（未纳入单值回代）\n')
        L.append('| 行标签 | 期间 | 内容 | 位置 |')
        L.append('|---|---|---|---|')
        for lab, p, raw, where in multi[:40]:
            L.append('| %s | %s | %s | %s |' % (lab, p, raw, where))
        if len(multi) > 40:
            L.append('| … | … | 共 %d 处 | |' % len(multi))
        L.append('')

    L.append('## 四、跨章一致性\n')
    L.append('- 同一指标在两个及以上表格中重复出现的「指标 × 期间」组合：**%d** 个' % len(dup))
    L.append('- 其中报告自身前后不一致的：**%d** 个' % len(selfbad))
    if selfbad:
        L.append('')
        for k, p, v in selfbad:
            L.append('- ❌ %s @ %s：%s' % (k, p, ' vs '.join('%s(%s)' % (fmt(a), b) for a, b in v)))
    L.append('')

    L.append('## 五、机器覆盖不到的正文散文数值点（盲区声明）\n')
    L.append('正文（非表格）段落中形态为 `d.d` 的数字共 **%d** 处。' % len(prose))
    L.append('这些数字**不在本校验器的回代范围内**（校验器只扫表格单元格）。')
    L.append('它们由"正文每一个派生数字都必须能在 `metrics.json` 找到键"这条**流程规则**约束，')
    L.append('而不是由程序强制拦截。**这是残余风险，不是已消除的风险。**')
    L.append('')

    L.append('## 六、自指统计核对（报告对自身覆盖情况的描述）\n')
    L.append('校验器实测：数值点 **%d**、已验证 **%d**、未覆盖 **%d**（占 **%.1f%%**）、覆盖率 **%.1f%%**。'
             % (ncell, npass, len(uncovered), unc_pct, cov))
    L.append('')
    if decl_bad:
        L.append('**✘ 报告散文里的覆盖率描述与本校验器实测不一致：**')
        L.append('')
        L.append('| 行号 | 报告写的数 | 问题 |')
        L.append('|---|---|---|')
        for i, v, why in decl_bad:
            L.append('| %d | **%s** | %s |' % (i, ('%.4g' % v), why))
        L.append('')
        L.append('**报告不得写一个校验器没有报过的覆盖率。** 请把散文里的自指数字改为实测值。')
    else:
        L.append('✔ 报告散文里出现的覆盖率类百分比，与本校验器实测一致。')
        L.append('（表格内的历史引用不在此检查范围内。）')
    L.append('')

    if nfail or selfbad or decl_bad:
        verdict = 'FAIL'
    elif uncovered:
        verdict = 'NO-MISMATCH'
    else:
        verdict = 'FULL-PASS'
    L.append('## 七、判词\n')
    L.append('```')
    L.append('VERDICT: %s  (scan=%d, cells=%d, verified=%d, fail=%d, uncovered=%d, '
             'coverage=%.1f%%, self_inconsistent=%d, self_declared_bad=%d, prose=%d)'
             % (verdict, n_scan, ncell, npass, nfail, len(uncovered), cov, len(selfbad),
                len(decl_bad), len(prose)))
    L.append('```')
    L.append('')
    L.append('| 判词 | 含义 | 允许怎么写 |')
    L.append('|---|---|---|')
    L.append('| `FULL-PASS` | 不一致=0 **且** 未覆盖=0 | 可以写"全部一致" |')
    L.append('| `NO-MISMATCH` | 不一致=0，**但仍有未覆盖** | **只能写**"已验证 N 个、覆盖率 X%、未覆盖 M 个（已列出）" |')
    L.append('| `FAIL` | 存在不一致、跨章矛盾，或**自指统计与实测不符** | 先修，再谈覆盖率 |')
    L.append('')

    if out:
        with io.open(out, 'w', encoding='utf-8') as f:
            f.write('\n'.join(L))
    print('\n'.join(L[:24]))
    print('...')
    print('VERDICT: %s  (scan=%d, cells=%d, verified=%d, fail=%d, uncovered=%d, '
          'coverage=%.1f%%, self_inconsistent=%d, self_declared_bad=%d, prose=%d)'
          % (verdict, n_scan, ncell, npass, nfail, len(uncovered), cov, len(selfbad),
             len(decl_bad), len(prose)))
    if verdict != 'FULL-PASS':
        print('注：未覆盖 ≠ 不一致；但**覆盖率未到 100% 之前，不得写"全部一致"**。')
    return 0 if not (nfail or selfbad or decl_bad) else 1


def fmt(x):
    if x is None:
        return '—'
    if abs(x) >= 100:
        return format(x, ',.2f')
    return '%.4g' % x


if __name__ == '__main__':
    sys.exit(main())
