# -*- coding: utf-8 -*-
"""
recalc.py — 通用派生指标复算引擎与数据质量校验

用法:
    python recalc.py <xlsx路径> <输出目录> [--company 公司名]

产出:
    <输出目录>/metrics.json  扁平键值（含 _periods / _period_kinds / _period_days，供 verify_report.py 回归校验）
    <输出目录>/metrics.md    人读版：公式 / 代入值 / 结果 / 门禁明细
    <输出目录>/checks.md     C1–C6 校验与 DQ 数据质量门禁记录

核心特性:
    1. 自动扫描资产负债表/利润表/现金流量表表头，提取期间列并按时间升序跨表对齐。
    2. 自动识别 FY(360天) / H1(180天) / Q1(90天) / Q3(270天) / OTHER 期间性质。
    3. 同比增速严格匹配上年同期间性质，无上年同期时返回 None（NOT_COMPUTABLE）。
    4. 存量平均余额优先匹配上年末（FY）期初；累计与 CAGR 按实际完整年度动态计算。
    5. 内置 DQ005 利息口径门禁标记，同步输出严格口径（_strict）与替代参考口径。
"""
import sys, io, os, json, re, datetime
import openpyxl


# ---------------------------------------------------------------- 动态期间解析
PERIOD_DAYS_MAP = {'FY': 360.0, 'H1': 180.0, 'Q1': 90.0, 'Q3': 270.0}
ANN_FACTOR_MAP = {'FY': 1.0, 'H1': 2.0, 'Q1': 4.0, 'Q3': 4.0 / 3.0}


def parse_period_header(val):
    """把表头单元格解析为标准期间元组: (label, year, month, kind) 或 None。

    支持格式:
      - Excel datetime/date: 2025-12-31 -> ('2025-12', 2025, 12, 'FY')
      - '2025-12', '2025/12', '2025.12', '2025年12月', '2025年12月31日'
      - '2025年', '2025年度', '2025FY', 'FY2025', '2025' -> ('2025-12', 2025, 12, 'FY')
      - '2026H1', '2026-06', '2026年6月', '2026年半年度', '2026中报', '2026年1-6月', '2026Q2' -> ('2026-06', 2026, 6, 'H1')
      - '2026Q1', '2026-03', '2026年一季度', '2026年1-3月' -> ('2026-03', 2026, 3, 'Q1')
      - '2026Q3', '2026-09', '2026年三季度', '2026年1-9月' -> ('2026-09', 2026, 9, 'Q3')
    """
    if val is None:
        return None
    if isinstance(val, (datetime.datetime, datetime.date)):
        y, m = val.year, val.month
        if 1990 <= y <= 2100 and 1 <= m <= 12:
            kind = _month_to_kind(m)
            return ('%04d-%02d' % (y, m), y, m, kind)
        return None
    if isinstance(val, (int, float)):
        if float(val).is_integer():
            iv = int(val)
            if 1990 <= iv <= 2100:
                return ('%04d-12' % iv, iv, 12, 'FY')
            if 199001 <= iv <= 210012:
                y, m = iv // 100, iv % 100
                if 1 <= m <= 12:
                    return ('%04d-%02d' % (y, m), y, m, _month_to_kind(m))
        return None

    s = str(val).strip().replace('**', '').replace('`', '')
    if not s:
        return None
    su = s.upper()

    # 1. 明确 Q1 / Q2 / Q3 / Q4 / H1 / H2 / FY
    m = re.search(r'(19\d\d|20\d\d)\s*(?:年)?\s*(H[12]|Q[1-4]|FY)', su)
    if not m:
        m_rev = re.search(r'(H[12]|Q[1-4]|FY)\s*(19\d\d|20\d\d)', su)
        if m_rev:
            y = int(m_rev.group(2))
            tag = m_rev.group(1)
            return _tag_to_period(y, tag)
    else:
        return _tag_to_period(int(m.group(1)), m.group(2))

    # 2. 中文季度 / 半年度 / 1-N月
    m_yr = re.search(r'(19\d\d|20\d\d)', s)
    if not m_yr:
        return None
    y = int(m_yr.group(1))

    if re.search(r'半年度|中报|上半年|1\s*[-~至—]\s*6\s*月', s):
        return ('%04d-06' % y, y, 6, 'H1')
    if re.search(r'一季|第一季|1\s*[-~至—]\s*3\s*月', s):
        return ('%04d-03' % y, y, 3, 'Q1')
    if re.search(r'三季|第三季|1\s*[-~至—]\s*9\s*月', s):
        return ('%04d-09' % y, y, 9, 'Q3')
    if re.search(r'全年|年度|年报|1\s*[-~至—]\s*12\s*月', s):
        return ('%04d-12' % y, y, 12, 'FY')

    # 3. YYYY-MM / YYYY/MM / YYYY.MM / YYYY年MM月
    m_ym = re.search(r'(19\d\d|20\d\d)\s*(?:[-/.年])\s*(\d{1,2})(?:\s*(?:[-/.月])|\b|$)', s)
    if m_ym:
        mo = int(m_ym.group(2))
        if 1 <= mo <= 12:
            return ('%04d-%02d' % (y, mo), y, mo, _month_to_kind(mo))

    # 4. 纯年份（如 "2024" 或 "2024年"）
    if re.fullmatch(r'(?:19\d\d|20\d\d)\s*年?', s):
        return ('%04d-12' % y, y, 12, 'FY')

    return None


def _month_to_kind(m):
    return {12: 'FY', 6: 'H1', 3: 'Q1', 9: 'Q3'}.get(m, 'OTHER')


def _tag_to_period(y, tag):
    if tag in ('FY', 'H2', 'Q4'):
        return ('%04d-12' % y, y, 12, 'FY')
    if tag in ('H1', 'Q2'):
        return ('%04d-06' % y, y, 6, 'H1')
    if tag == 'Q1':
        return ('%04d-03' % y, y, 3, 'Q1')
    if tag == 'Q3':
        return ('%04d-09' % y, y, 9, 'Q3')
    return ('%04d-12' % y, y, 12, 'FY')


def period_days(pmeta):
    kind, month = pmeta['kind'], pmeta['month']
    if kind in PERIOD_DAYS_MAP:
        return PERIOD_DAYS_MAP[kind]
    return float(month * 30)


def ann_factor(pmeta):
    kind, month = pmeta['kind'], pmeta['month']
    if kind in ANN_FACTOR_MAP:
        return ANN_FACTOR_MAP[kind]
    return 12.0 / month if month else 1.0


def find_header_row(ws, max_scan_rows=15):
    """自动寻找包含期间列的表头行，返回 (header_row_idx, {col_idx: period_tuple})。"""
    best_row, best_cols = None, {}
    for r in range(1, min(ws.max_row, max_scan_rows) + 1):
        cols = {}
        for c in range(2, ws.max_column + 1):
            p = parse_period_header(ws.cell(r, c).value)
            if p is not None:
                cols[c] = p
        if len(cols) > len(best_cols):
            best_row, best_cols = r, cols
    return best_row, best_cols


def find_sheet(wb, candidates):
    for name in candidates:
        if name in wb.sheetnames:
            return name
    for s in wb.sheetnames:
        for name in candidates:
            if name in s:
                return s
    return None


# ---------------------------------------------------------------- 读取与跨表对齐
def load(path):
    wbf = openpyxl.load_workbook(path, data_only=False)
    wbv = openpyxl.load_workbook(path, data_only=True)

    sheet_map = {
        'bs': find_sheet(wbf, ['资产负债表']),
        'is': find_sheet(wbf, ['利润表', '损益表']),
        'cf': find_sheet(wbf, ['现金流量表']),
        'note': find_sheet(wbf, ['使用说明']),
        'src': find_sheet(wbf, ['数据来源']),
    }

    # 第一遍：收集三表全部期间并按时间升序统一编号
    raw_headers = {}
    all_periods = {}
    for key in ('bs', 'is', 'cf'):
        sname = sheet_map[key]
        if not sname:
            raw_headers[key] = (None, {})
            continue
        hr, cols = find_header_row(wbv[sname])
        raw_headers[key] = (hr, cols)
        for p in cols.values():
            all_periods[p[0]] = {'label': p[0], 'year': p[1], 'month': p[2], 'kind': p[3]}

    if not all_periods:
        raise ValueError('未在资产负债表/利润表/现金流量表表头中识别出任何期间列，请检查表头格式。')

    pmetas = sorted(all_periods.values(), key=lambda x: (x['year'], x['month']))
    p_index = {pm['label']: idx for idx, pm in enumerate(pmetas)}
    n_periods = len(pmetas)

    # 第二遍：按统一期间轴对齐读取每张表（即使列顺序相反或期数不同也绝不错位）
    def grab_aligned(key):
        sname = sheet_map[key]
        if not sname:
            return {}
        hr, cols = raw_headers[key]
        if hr is None or not cols:
            return {}
        ws, wv = wbf[sname], wbv[sname]
        d = {}
        for r in range(hr + 1, ws.max_row + 1):
            lab = ws.cell(r, 1).value
            if not isinstance(lab, str):
                continue
            lab = lab.strip()
            if not lab or lab in d:
                continue
            row = [None] * n_periods
            for c, p in cols.items():
                v = wv.cell(r, c).value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    row[p_index[p[0]]] = float(v)
            d[lab] = row
        return d

    sheets = {
        'bs': grab_aligned('bs'),
        'is': grab_aligned('is'),
        'cf': grab_aligned('cf'),
        '_names': wbf.sheetnames,
        '_sheet_map': sheet_map,
        '_pmetas': pmetas,
        '_periods': [pm['label'] for pm in pmetas],
        '_period_kinds': [pm['kind'] for pm in pmetas],
        '_period_days': [period_days(pm) for pm in pmetas],
        '_sheet_periods': {
            k: sorted([p[0] for p in raw_headers[k][1].values()])
            for k in ('bs', 'is', 'cf')
        },
    }
    return sheets


def G(d, *names):
    """按候选名取一整行（按统一期间对齐）。找不到返回 None。"""
    for n in names:
        if n in d:
            return d[n]
    return None


def V(row, i, default=0.0):
    """取第 i 期数值，空白视为 default（默认 0，仅用于明确可加总字段）。"""
    if row is None or i is None or i < 0 or i >= len(row):
        return default
    v = row[i]
    return default if v is None else v


def has(row, i):
    return row is not None and i is not None and 0 <= i < len(row) and row[i] is not None


def sdiv(a, b):
    if a is None or b is None or b == 0:
        return None
    return a / b


def find_prior_same_kind(pmetas, i):
    """寻找上年同期间性质的索引（用于严格同比）。"""
    cur = pmetas[i]
    for j in range(i - 1, -1, -1):
        if pmetas[j]['year'] == cur['year'] - 1 and pmetas[j]['kind'] == cur['kind'] and pmetas[j]['month'] == cur['month']:
            return j
    return None


def find_opening_bs_index(pmetas, i):
    """寻找第 i 期的期初资产负债表索引：优先匹配上年末（year-1, FY），否则退回相邻前一期 FY。"""
    cur = pmetas[i]
    for j in range(i - 1, -1, -1):
        if pmetas[j]['year'] == cur['year'] - 1 and pmetas[j]['kind'] == 'FY':
            return j
    if i >= 1 and cur['kind'] == 'FY' and pmetas[i - 1]['kind'] == 'FY' and pmetas[i - 1]['year'] == cur['year'] - 1:
        return i - 1
    return None


# ---------------------------------------------------------------- 主计算
def compute(S):
    bs, isr, cf = S['bs'], S['is'], S['cf']
    pmetas = S['_pmetas']
    PLAB = S['_periods']
    PKIND = S['_period_kinds']
    N = len(pmetas)

    fy_indices = [i for i, pm in enumerate(pmetas) if pm['kind'] == 'FY']
    last_fy = fy_indices[-1] if fy_indices else None
    first_fy = fy_indices[0] if fy_indices else None
    last_i = N - 1

    # 资产负债表
    A = {
        'cash': G(bs, '货币资金'),
        'ar': G(bs, '应收账款'),
        'notes_r': G(bs, '应收票据'),
        'prepay': G(bs, '预付款项'),
        'oar': G(bs, '其他应收款'),
        'inv': G(bs, '存货'),
        'cur_nca': G(bs, '一年内到期的非流动资产'),
        'oca': G(bs, '其他流动资产'),
        'ca': G(bs, '流动资产合计'),
        'lt_eq': G(bs, '长期股权投资'),
        'fa': G(bs, '固定资产'),
        'cip': G(bs, '在建工程'),
        'ia': G(bs, '无形资产'),
        'ltd': G(bs, '长期待摊费用'),
        'nca': G(bs, '非流动资产合计'),
        'ta': G(bs, '资产总计'),
        'st_loan': G(bs, '短期借款'),
        'notes_p': G(bs, '应付票据'),
        'ap': G(bs, '应付账款'),
        'adv': G(bs, '预收款项', '合同负债'),
        'payroll': G(bs, '应付职工薪酬'),
        'tax': G(bs, '应交税费'),
        'oap': G(bs, '其他应付款'),
        'cur_ncl': G(bs, '一年内到期的非流动负债'),
        'cl': G(bs, '流动负债合计'),
        'lt_loan': G(bs, '长期借款'),
        'bonds_p': G(bs, '应付债券'),
        'lt_pay': G(bs, '长期应付款'),
        'def_inc': G(bs, '递延收益'),
        'ncl': G(bs, '非流动负债合计'),
        'tl': G(bs, '负债合计'),
        'capital': G(bs, '实收资本（或股本）', '实收资本'),
        'reserve': G(bs, '专项储备'),
        're': G(bs, '未分配利润'),
        'eq': G(bs, '所有者权益合计'),
        'tle': G(bs, '负债和所有者权益总计'),
    }
    # 利润表
    P = {
        'rev': G(isr, '一、营业收入', '营业收入'),
        'cost': G(isr, '减：营业成本', '营业成本'),
        'tax_sur': G(isr, '税金及附加'),
        'sell': G(isr, '销售费用'),
        'admin': G(isr, '管理费用'),
        'rd': G(isr, '研发费用'),
        'fin': G(isr, '财务费用'),
        'int_exp': G(isr, '其中：利息费用', '利息费用'),
        'int_inc': G(isr, '利息收入'),
        'other_inc': G(isr, '加：其他收益', '其他收益'),
        'inv_inc': G(isr, '投资收益'),
        'fv_inc': G(isr, '公允价值变动收益'),
        'op_profit': G(isr, '二、营业利润', '营业利润'),
        'non_op_in': G(isr, '加：营业外收入', '营业外收入'),
        'non_op_out': G(isr, '减：营业外支出', '营业外支出'),
        'tp': G(isr, '三、利润总额', '利润总额'),
        'itax': G(isr, '减：所得税费用', '所得税费用'),
        'ni': G(isr, '四、净利润', '净利润'),
        'impair': G(isr, '信用减值损失', '资产减值损失'),
    }
    # 现金流量表
    C = {
        'sale_cash': G(cf, '销售商品、提供劳务收到的现金'),
        'op_in': G(cf, '经营活动现金流入小计'),
        'op_out': G(cf, '经营活动现金流出小计'),
        'ocf': G(cf, '经营活动产生的现金流量净额'),
        'capex': G(cf, '购建固定资产、无形资产和其他长期资产支付的现金'),
        'inv_pay': G(cf, '投资支付的现金'),
        'icf': G(cf, '投资活动产生的现金流量净额'),
        'eq_in': G(cf, '吸收投资收到的现金'),
        'borrow': G(cf, '取得借款收到的现金'),
        'repay': G(cf, '偿还债务支付的现金'),
        'div_int': G(cf, '分配股利、利润或偿付利息支付的现金'),
        'fcf_net': G(cf, '筹资活动产生的现金流量净额'),
        'fx': G(cf, '四、汇率变动对现金及现金等价物的影响'),
        'net_inc': G(cf, '五、现金及现金等价物净增加额'),
        'open_cash': G(cf, '加：期初现金及现金等价物余额'),
        'close_cash': G(cf, '六、期末现金及现金等价物余额'),
    }

    M, F, BOOL, T = {}, {}, {}, {}

    def put(key, vals, formula=''):
        M[key] = vals
        F[key] = formula

    # --- 1. 资产负债表主项 ---
    put('zzc', [V(A['ta'], i) if has(A['ta'], i) else None for i in range(N)], '资产总计')
    put('fz', [V(A['tl'], i) if has(A['tl'], i) else None for i in range(N)], '负债合计')
    put('qy', [V(A['eq'], i) if has(A['eq'], i) else None for i in range(N)], '所有者权益合计')
    put('ldzc', [V(A['ca'], i) if has(A['ca'], i) else None for i in range(N)], '流动资产合计')
    put('fldzc', [V(A['nca'], i) if has(A['nca'], i) else None for i in range(N)], '非流动资产合计')
    put('ldfz', [V(A['cl'], i) if has(A['cl'], i) else None for i in range(N)], '流动负债合计')
    put('fldfz', [V(A['ncl'], i) if has(A['ncl'], i) else None for i in range(N)], '非流动负债合计')

    # --- 2. 有息 / 无息负债 ---
    st = [V(A['st_loan'], i) + V(A['notes_p'], i) + V(A['cur_ncl'], i) for i in range(N)]
    lt = [V(A['lt_loan'], i) + V(A['bonds_p'], i) for i in range(N)]
    xy = [st[i] + lt[i] for i in range(N)]
    gy = [xy[i] + V(A['lt_pay'], i) for i in range(N)]
    put('dc_yx', st, '短期有息负债 = 短期借款 + 应付票据 + 一年内到期非流动负债')
    put('cq_yx', lt, '长期有息负债 = 长期借款 + 应付债券')
    put('yx_xy', xy, '狭义有息 = 短期有息负债 + 长期有息负债')
    put('yx_gy', gy, '广义有息 = 狭义有息 + 长期应付款')
    put('cqyp', [V(A['lt_pay'], i) for i in range(N)], '长期应付款（性质待核）')

    wx_sum = [V(A['ap'], i) + V(A['adv'], i) + V(A['payroll'], i) + V(A['tax'], i)
              + V(A['oap'], i) + V(A['def_inc'], i) for i in range(N)]
    wx_res = [V(A['tl'], i) - xy[i] if has(A['tl'], i) else None for i in range(N)]
    put('wx_sum', wx_sum, '无息负债(加总口径) = 应付账款+预收/合同负债+应付职工薪酬+应交税费+其他应付款+递延收益')
    put('wx_res', wx_res, '无息负债(倒减口径) = 负债合计 − 狭义有息（可能混入长期应付款，口径不纯）')
    put('wx_pure', [V(A['tl'], i) - gy[i] if has(A['tl'], i) else None for i in range(N)],
        '无息负债(纯粹口径) = 负债合计 − 狭义有息 − 长期应付款')

    put('yxzbl', [sdiv(xy[i], M['fz'][i]) for i in range(N)], '狭义有息 / 负债合计')
    put('gyxzbl', [sdiv(gy[i], M['fz'][i]) for i in range(N)], '广义有息 / 负债合计')
    put('wxzbl', [sdiv(wx_sum[i], M['fz'][i]) for i in range(N)], '加总口径无息负债 / 负债合计')
    put('dcb', [sdiv(st[i], lt[i]) for i in range(N)], '短长比 = 短期有息 / 长期有息')
    put('dqhbl', [sdiv(st[i], xy[i]) for i in range(N)], '短期有息 / 狭义有息')

    # --- 3. 偿债与资本结构比率 ---
    put('zcfzl', [sdiv(M['fz'][i], M['zzc'][i]) for i in range(N)], '资产负债率 = 负债合计 / 资产总计')
    adj_ta = [(V(A['ta'], i) - V(A['ltd'], i) - 0.5 * V(A['cip'], i)) if has(A['ta'], i) else None for i in range(N)]
    put('adj_ta', adj_ta, '调整后资产 = 资产总计 − 长期待摊费用 − 50% × 在建工程')
    put('zcfzl_adj', [sdiv(M['fz'][i], adj_ta[i]) for i in range(N)],
        '剔除虚资产后资产负债率 = 负债合计 / (资产总计 − 长期待摊费用 − 50%×在建工程)')
    put('jrfl', [sdiv(xy[i], M['zzc'][i]) for i in range(N)], '资产金融负债率 = 狭义有息 / 资产总计')

    put('ldbl', [sdiv(M['ldzc'][i], M['ldfz'][i]) for i in range(N)], '流动比率 = 流动资产 / 流动负债')
    qa = [(V(A['ca'], i) - V(A['inv'], i) - V(A['prepay'], i) - V(A['cur_nca'], i) - V(A['oca'], i))
          if has(A['ca'], i) else None for i in range(N)]
    put('qa', qa, '从严速动资产 = 流动资产 − 存货 − 预付款项 − 一年内到期非流动资产 − 其他流动资产')
    put('sdbl', [sdiv(qa[i], M['ldfz'][i]) for i in range(N)], '速动比率(从严) = 从严速动资产 / 流动负债')

    # DQ005 门禁：存在重大有息负债且财务费用>0，但利息费用缺失或<=0
    dq005 = []
    ie_legacy, ie_strict, yhlx_basis = [], [], []
    for i in range(N):
        has_int = has(P['int_exp'], i) and V(P['int_exp'], i) > 0
        fin_pos = has(P['fin'], i) and V(P['fin'], i) > 0
        triggered_dq5 = (xy[i] > 0) and (not has_int) and fin_pos
        dq005.append(triggered_dq5)
        if has_int and has(P['tp'], i):
            val = (V(P['tp'], i) + V(P['int_exp'], i)) / V(P['int_exp'], i)
            ie_strict.append(val)
            ie_legacy.append(val)
            yhlx_basis.append('利息费用')
        elif fin_pos and has(P['tp'], i):
            ie_strict.append(None)  # 触发 DQ005 时严格口径返回 NOT_COMPUTABLE
            ie_legacy.append((V(P['tp'], i) + V(P['fin'], i)) / V(P['fin'], i))
            yhlx_basis.append('DQ005:财务费用替代(偏低/审慎，严格口径禁用)')
        else:
            ie_strict.append(None)
            ie_legacy.append(None)
            yhlx_basis.append('数据缺失')

    put('yhlx', ie_legacy, '已获利息倍数（财务费用替代参考口径；触发 DQ005 时严禁直接作结论，参见 yhlx_strict）')
    put('yhlx_strict', ie_strict, '已获利息倍数（严格口径；触发 DQ005 时返回 None/NOT_COMPUTABLE）')
    put('yhlx_basis', yhlx_basis, '已获利息倍数分母取值来源与 DQ005 状态')

    put('yyzb', [(M['ldzc'][i] - M['ldfz'][i]) if (M['ldzc'][i] is not None and M['ldfz'][i] is not None) else None for i in range(N)],
        '营运资本 = 流动资产 − 流动负债')
    put('hbzj_cover', [sdiv(V(A['cash'], i) if has(A['cash'], i) else None, st[i]) for i in range(N)],
        '货币资金覆盖率 = 货币资金 / 短期有息负债')
    put('aqfz', [(M['zzc'][i] - M['fz'][i]) if (M['zzc'][i] is not None and M['fz'][i] is not None) else None for i in range(N)],
        '净资产 = 资产总计 − 负债合计')
    cqzj = [(M['qy'][i] + M['fldfz'][i]) if (M['qy'][i] is not None and M['fldfz'][i] is not None) else None for i in range(N)]
    put('cqzj', cqzj, '长期资金 = 所有者权益 + 非流动负债')
    put('cqzj_cover', [sdiv(cqzj[i], M['fldzc'][i]) for i in range(N)], '长期资金 / 非流动资产')

    # --- 4. 盈利质量 ---
    core = []
    core_strict = []
    for i in range(N):
        if not has(P['rev'], i) or not has(P['cost'], i):
            core.append(None)
            core_strict.append(None)
            continue
        base_core = (V(P['rev'], i) - V(P['cost'], i) - V(P['tax_sur'], i)
                     - V(P['sell'], i) - V(P['admin'], i) - V(P['rd'], i))
        core.append(base_core - V(P['fin'], i))
        if dq005[i]:
            core_strict.append(None)
        else:
            int_val = V(P['int_exp'], i) if has(P['int_exp'], i) else V(P['fin'], i)
            core_strict.append(base_core - int_val)

    put('hxlr', core, '核心利润（减财务费用参考口径；研发费用留空期按0代入属上限假设，见 rd_blank）')
    put('hxlr_strict', core_strict, '核心利润（严格口径：减可靠利息费用，触发 DQ005 时返回 None）')
    put('mll', [sdiv(V(P['rev'], i) - V(P['cost'], i), V(P['rev'], i)) if (has(P['rev'], i) and has(P['cost'], i)) else None for i in range(N)],
        '毛利率 =（营业收入 − 营业成本）/ 营业收入')
    put('jll', [sdiv(V(P['ni'], i) if has(P['ni'], i) else None, V(P['rev'], i) if has(P['rev'], i) else None) for i in range(N)],
        '净利率 = 净利润 / 营业收入')
    put('hxlrl', [sdiv(core[i], V(P['rev'], i) if has(P['rev'], i) else None) for i in range(N)],
        '核心利润率 = 核心利润 / 营业收入')

    roe, roe_ann, zzczzl, qycs = [], [], [], []
    for i in range(N):
        op_i = find_opening_bs_index(pmetas, i)
        if op_i is None or not has(A['eq'], i) or not has(A['eq'], op_i) or not has(A['ta'], i) or not has(A['ta'], op_i):
            roe.append(None); roe_ann.append(None); zzczzl.append(None); qycs.append(None)
            continue
        avg_eq = (V(A['eq'], i) + V(A['eq'], op_i)) / 2.0
        avg_ta = (V(A['ta'], i) + V(A['ta'], op_i)) / 2.0
        af = ann_factor(pmetas[i])
        ni_i = V(P['ni'], i) if has(P['ni'], i) else None
        rev_i = V(P['rev'], i) if has(P['rev'], i) else None
        roe.append(sdiv(ni_i, avg_eq))
        roe_ann.append(sdiv(ni_i * af, avg_eq) if ni_i is not None else None)
        zzczzl.append(sdiv(rev_i, avg_ta))
        qycs.append(sdiv(avg_ta, avg_eq))

    put('roe', roe, 'ROE = 净利润 / 平均所有者权益（非全年期不可直接与全年同比）')
    put('roe_end', [sdiv(V(P['ni'], i) if has(P['ni'], i) else None, M['qy'][i]) for i in range(N)], 'ROE(期末权益口径)')
    put('roe_ann', roe_ann, 'ROE(年化) = 净利润 × 年化倍数(FY×1/H1×2/Q1×4/Q3×4/3) / 平均所有者权益')
    put('zzczzl', zzczzl, '总资产周转率 = 营业收入 / 平均资产总计')
    put('qycs', qycs, '权益乘数 = 平均资产总计 / 平均所有者权益')
    fjy = [(V(P['non_op_in'], i) - V(P['non_op_out'], i)) if (has(P['non_op_in'], i) or has(P['non_op_out'], i)) else 0.0 for i in range(N)]
    put('fjy', fjy, '非经常性损益(简化) = 营业外收入 − 营业外支出')
    put('fjy_ni', [sdiv(fjy[i], V(P['ni'], i) if has(P['ni'], i) else None) for i in range(N)], '非经常性损益 / 净利润')

    # --- 5. 现金流 ---
    put('jyxj', [V(C['ocf'], i) if has(C['ocf'], i) else None for i in range(N)], '经营活动现金流量净额')
    put('tzxj', [V(C['icf'], i) if has(C['icf'], i) else None for i in range(N)], '投资活动现金流量净额')
    put('czxj', [V(C['fcf_net'], i) if has(C['fcf_net'], i) else None for i in range(N)], '筹资活动现金流量净额')
    put('jxb', [sdiv(M['jyxj'][i], V(P['ni'], i) if has(P['ni'], i) else None) for i in range(N)], '净现比 = 经营净额 / 净利润')
    put('hxlr_hxb', [sdiv(M['jyxj'][i], core[i]) for i in range(N)], '核心利润获现比 = 经营净额 / 核心利润')
    put('sxxjl', [sdiv(V(C['sale_cash'], i) if has(C['sale_cash'], i) else None, V(P['rev'], i) if has(P['rev'], i) else None) for i in range(N)],
        '销售收现比 = 销售商品提供劳务收到的现金 / 营业收入')
    put('fcf', [(M['jyxj'][i] - V(C['capex'], i)) if M['jyxj'][i] is not None else None for i in range(N)],
        '自由现金流(近似) = 经营活动净额 − 购建长期资产支付的现金（注：未区分维持性与扩张性资本开支）')
    put('fcf_int', [sdiv(M['fcf'][i], V(C['div_int'], i) if has(C['div_int'], i) else None) for i in range(N)], '自由现金流 / 分配股利或偿付利息现金')
    put('xjdfl', [sdiv(M['jyxj'][i], M['ldfz'][i]) for i in range(N)], '现金流动负债比 = 经营活动净额 / 流动负债')
    comb = []
    for i in range(N):
        if M['jyxj'][i] is None or M['tzxj'][i] is None or M['czxj'][i] is None:
            comb.append(None)
        else:
            comb.append('%s%s%s' % (
                '+' if M['jyxj'][i] >= 0 else '−',
                '+' if M['tzxj'][i] >= 0 else '−',
                '+' if M['czxj'][i] >= 0 else '−',
            ))
    put('comb', comb, '现金流三类符号组合（仅描述资金来源结构，+++不得直接等同造假）')

    # --- 5b. 营运资产与累计口径（全动态期间） ---
    wca = [V(A['ar'], i) + V(A['prepay'], i) + V(A['inv'], i) for i in range(N)]
    put('wca', wca, '营运资产合计 = 应收账款 + 预付款项 + 存货')
    put('wca_ratio', [sdiv(wca[i], M['zzc'][i]) for i in range(N)], '营运资产合计 / 资产总计')
    put('wca_delta', [None if i == 0 else wca[i] - wca[i - 1] for i in range(N)], '营运资产较上期增量')

    yoy_wca, wca_per_rev = [], []
    for i in range(N):
        pj = find_prior_same_kind(pmetas, i)
        if pj is None:
            yoy_wca.append(None); wca_per_rev.append(None)
        else:
            yoy_wca.append(sdiv(wca[i], wca[pj]) - 1.0 if wca[pj] else None)
            drev = (V(P['rev'], i) - V(P['rev'], pj)) if (has(P['rev'], i) and has(P['rev'], pj)) else None
            wca_per_rev.append(sdiv(wca[i] - wca[pj], drev))
    put('yoy_wca', yoy_wca, '营运资产同比增速（严格同期间性质比对）')
    put('wca_per_rev_delta', wca_per_rev, '营运资产增量 / 收入增量（倍）')
    put('core_elasticity', [sdiv(V(P['rev'], i) - V(P['cost'], i), core[i]) if (has(P['rev'], i) and has(P['cost'], i)) else None for i in range(N)],
        '核心利润对收入弹性 = 毛利额 / 核心利润')
    put('st_over_nca', [sdiv(st[i], M['fldzc'][i]) for i in range(N)], '短期有息负债 / 非流动资产')
    put('ocf_over_int', [sdiv(M['jyxj'][i], V(C['div_int'], i) if has(C['div_int'], i) else None) for i in range(N)],
        '经营活动净额 / 分配股利或偿付利息现金')
    put('cash_ocf_over_st', [sdiv((V(A['cash'], i) + M['jyxj'][i]) if (has(A['cash'], i) and M['jyxj'][i] is not None) else None, st[i]) for i in range(N)],
        '（货币资金 + 经营活动净额）/ 短期有息负债')

    # 动态全年度累计（自动覆盖所有含完整数据的 FY 期）
    fy_cf_indices = [i for i in fy_indices if has(C['ocf'], i)]
    fy_span_label = ('%s—%s(%d个完整年度)' % (PLAB[fy_indices[0]], PLAB[fy_indices[-1]], len(fy_indices))) if fy_indices else '无完整年度'
    fy_cf_span_label = ('%s—%s(%d个完整年度)' % (PLAB[fy_cf_indices[0]], PLAB[fy_cf_indices[-1]], len(fy_cf_indices))) if fy_cf_indices else '无现金流年度'

    def cum_on(row, idxs):
        if not idxs or row is None:
            return None
        vals = [row[j] for j in idxs if row[j] is not None]
        return sum(vals) if len(vals) == len(idxs) else None

    c_rev = cum_on(P['rev'], fy_indices)
    c_ni = cum_on(P['ni'], fy_indices)
    c_ocf = cum_on(C['ocf'], fy_cf_indices)
    c_int = cum_on(C['div_int'], fy_cf_indices)
    c_capex = cum_on(C['capex'], fy_cf_indices)
    c_fcf = (c_ocf - c_capex) if (c_ocf is not None and c_capex is not None) else None
    c_cov = sdiv(c_ocf, c_int)

    anchor_fy = last_fy
    for cname, val, desc in [
        ('cum_rev5', c_rev, fy_span_label + ' 累计营业收入'),
        ('cum_ni5', c_ni, fy_span_label + ' 累计净利润'),
        ('cum_ocf5', c_ocf, fy_cf_span_label + ' 累计经营活动净额'),
        ('cum_int5', c_int, fy_cf_span_label + ' 累计「分配股利、利润或偿付利息支付的现金」'),
        ('cum_capex5', c_capex, fy_cf_span_label + ' 累计购建长期资产支付的现金'),
        ('cum_fcf5', c_fcf, fy_cf_span_label + ' 累计自由现金流'),
        ('cum_ocf_cov_int5', c_cov, fy_cf_span_label + ' 累计经营净额 / 累计分配股利或偿付利息现金'),
    ]:
        put(cname, [val if (anchor_fy is not None and i == anchor_fy) else None for i in range(N)], desc)

    # --- 5c. 严格同比增速（自动匹配上年同期间性质，FY 与 H1/Q 绝不混比） ---
    for k, row in {'rev': P['rev'], 'ni': P['ni'], 'ta': A['ta'], 'eq': A['eq'], 'tl': A['tl'],
                   'ar': A['ar'], 'inv': A['inv'], 'prepay': A['prepay'], 'ap': A['ap']}.items():
        vals = []
        for i in range(N):
            pj = find_prior_same_kind(pmetas, i)
            if pj is None or not has(row, i) or not has(row, pj) or V(row, pj) == 0:
                vals.append(None)
            else:
                vals.append(V(row, i) / V(row, pj) - 1.0)
        put('yoy_' + k, vals, '同比增速(%s) = 本期 / 上年同期间性质 − 1（无上年同期则 NOT_COMPUTABLE）' % k)

    # --- 5d. 压力测试 ---
    gross = [(V(P['rev'], i) - V(P['cost'], i)) if (has(P['rev'], i) and has(P['cost'], i)) else None for i in range(N)]
    put('gross', gross, '毛利额 = 营业收入 − 营业成本')
    put('gap_refi', [(st[i] - V(A['cash'], i)) if has(A['cash'], i) else None for i in range(N)],
        '续贷中断情景缺口 = 短期有息负债 − 货币资金（未扣受限）')
    st_core = [(core[i] - 0.20 * gross[i]) if (core[i] is not None and gross[i] is not None) else None for i in range(N)]
    st_tp = [(st_core[i] + fjy[i]) if st_core[i] is not None else None for i in range(N)]
    st_yhlx = []
    for i in range(N):
        denom = V(P['int_exp'], i) if (has(P['int_exp'], i) and V(P['int_exp'], i) > 0) else (V(P['fin'], i) if has(P['fin'], i) else None)
        st_yhlx.append(sdiv(st_tp[i] + denom, denom) if (st_tp[i] is not None and denom) else None)
    put('stress_core', st_core, '情景①核心利润 = 核心利润 − 20%×毛利额（假设毛利率与期间费用不变）')
    put('stress_yhlx', st_yhlx, '情景①已获利息倍数 =（情景利润总额 + 利息/财务费用）/ 利息/财务费用')

    # --- 6. 周转天数（动态期间天数 + 正确匹配期初余额） ---
    def turn(flow_row, stock_row, i):
        op_i = find_opening_bs_index(pmetas, i)
        if op_i is None or not has(stock_row, i) or not has(stock_row, op_i) or not has(flow_row, i):
            return None
        avg_bal = (V(stock_row, i) + V(stock_row, op_i)) / 2.0
        fl = V(flow_row, i)
        if not fl:
            return None
        return period_days(pmetas[i]) * avg_bal / fl

    put('dio', [turn(P['cost'], A['inv'], i) for i in range(N)], '存货周转天数 = 实际期间天数(FY360/H1 180/Q1 90/Q3 270) × 平均存货 / 营业成本')
    put('dso', [turn(P['rev'], A['ar'], i) for i in range(N)], '应收周转天数 = 实际期间天数 × 平均应收账款 / 营业收入')
    put('dpo', [turn(P['cost'], A['ap'], i) for i in range(N)], '应付周转天数 = 实际期间天数 × 平均应付账款 / 营业成本（不含应付票据）')
    put('dpp', [turn(P['cost'], A['prepay'], i) for i in range(N)], '预付周转天数 = 实际期间天数 × 平均预付款项 / 营业成本')
    put('dad', [turn(P['rev'], A['adv'], i) for i in range(N)], '预收周转天数 = 实际期间天数 × 平均预收/合同负债 / 营业收入')
    put('ccc', [
        (M['dio'][i] + M['dso'][i] - M['dpo'][i])
        if (M['dio'][i] is not None and M['dso'][i] is not None and M['dpo'][i] is not None) else None
        for i in range(N)
    ], 'CCC = 存货周转天数 + 应收周转天数 − 应付周转天数（核心口径）')

    # 营运资金量（在所有可计算的完整年度上计算）
    yyzj_cycle, yyzj_turns, yyzj_amt = [None] * N, [None] * N, [None] * N
    for i in fy_indices:
        if None not in (M['dio'][i], M['dso'][i], M['dpo'][i], M['dpp'][i], M['dad'][i]):
            cyc = M['dio'][i] + M['dso'][i] - M['dpo'][i] + M['dpp'][i] - M['dad'][i]
            yyzj_cycle[i] = cyc
            if cyc > 0 and has(P['rev'], i) and V(P['rev'], i) > 0 and has(P['ni'], i):
                turns = 360.0 / cyc
                yyzj_turns[i] = turns
                yyzj_amt[i] = V(P['rev'], i) * (1.0 - V(P['ni'], i) / V(P['rev'], i)) / turns
    put('yyzj_cycle', yyzj_cycle, '营运资金周转周期 = DIO + DSO − DPO + 预付天数 − 预收天数')
    put('yyzj_turns', yyzj_turns, '营运资金周转次数 = 360 / 营运资金周转周期')
    put('yyzj_amt', yyzj_amt, '营运资金量 = 上年营收 × (1 − 销售利润率) ÷ 营运资金周转次数（监管参考测算）')

    # --- 7. 动态 CAGR 与 SGR ---
    def calc_cagr(a, b, yrs):
        if a is None or b is None or a <= 0 or b <= 0 or yrs <= 0:
            return None
        return (b / a) ** (1.0 / yrs) - 1.0

    cagr_keys = {'rev': P['rev'], 'ni': P['ni'], 'ta': A['ta'], 'eq': A['eq'], 'tl': A['tl']}
    fy_years = (pmetas[last_fy]['year'] - pmetas[first_fy]['year']) if (first_fy is not None and last_fy is not None and last_fy > first_fy) else 0
    # 最近期非完整年度的跨期区间计算（如上上年末至最近期中报/季报）
    recent_start = None
    recent_years = 0.0
    if last_i is not None and pmetas[last_i]['kind'] != 'FY' and len(fy_indices) >= 2:
        recent_start = fy_indices[-2]
        recent_years = (pmetas[last_i]['year'] - pmetas[recent_start]['year']) + (pmetas[last_i]['month'] - pmetas[recent_start]['month']) / 12.0
    elif last_i is not None and len(fy_indices) >= 1 and last_i > fy_indices[-1]:
        recent_start = fy_indices[-1]
        recent_years = (pmetas[last_i]['year'] - pmetas[recent_start]['year']) + (pmetas[last_i]['month'] - pmetas[recent_start]['month']) / 12.0

    for k, row in cagr_keys.items():
        v_fy = calc_cagr(V(row, first_fy) if (first_fy is not None and has(row, first_fy)) else None,
                         V(row, last_fy) if (last_fy is not None and has(row, last_fy)) else None,
                         float(fy_years)) if fy_years > 0 else None
        put('cagr4_' + k, [v_fy if i == 0 else None for i in range(N)],
            '完整年度 CAGR(%s)：%s → %s（%g 年）' % (k, PLAB[first_fy] if first_fy is not None else '—', PLAB[last_fy] if last_fy is not None else '—', fy_years))
        v_rec = calc_cagr(V(row, recent_start) if (recent_start is not None and has(row, recent_start)) else None,
                          V(row, last_i) if has(row, last_i) else None,
                          recent_years) if recent_years > 0 else None
        put('cagr15_' + k, [v_rec if i == 0 else None for i in range(N)],
            '近期跨期 CAGR(%s)：%s → %s（%g 年）' % (k, PLAB[recent_start] if recent_start is not None else '—', PLAB[last_i], recent_years))

    sgr_vals, ret_vals = [None] * N, [None] * N
    for i in fy_indices:
        op_i = find_opening_bs_index(pmetas, i)
        if op_i is not None and has(P['ni'], i) and V(P['ni'], i) != 0 and has(A['re'], i) and has(A['re'], op_i) and roe[i] is not None:
            ret = (V(A['re'], i) - V(A['re'], op_i)) / V(P['ni'], i)
            ret_vals[i] = ret
            sgr_vals[i] = roe[i] * ret
    put('sgr', sgr_vals, 'SGR ≈ ROE(平均权益口径) × 留存收益率')
    put('retention', ret_vals, '留存收益率 = 未分配利润增量 / 本期净利润')

    # --- 8. Z'-Score & M-Score ---
    z, zt = [], []
    for i in range(N):
        ta = M['zzc'][i]
        tl = M['fz'][i]
        if not ta or not tl or not has(P['rev'], i) or not has(P['tp'], i):
            z.append(None); zt.append(None); continue
        x1 = (V(A['ca'], i) - V(A['cl'], i)) / ta
        x2 = (V(A['re'], i) + V(A['reserve'], i)) / ta
        x3 = (V(P['tp'], i) + V(P['fin'], i)) / ta
        x4 = V(A['eq'], i) / tl
        x5 = V(P['rev'], i) / ta
        z.append(0.717 * x1 + 0.847 * x2 + 3.107 * x3 + 0.420 * x4 + 0.998 * x5)
        af = ann_factor(pmetas[i])
        x3_a = (V(P['tp'], i) * af + V(P['fin'], i) * af) / ta
        x5_a = (V(P['rev'], i) * af) / ta
        zt.append(0.717 * x1 + 0.847 * x2 + 3.107 * x3_a + 0.420 * x4 + 0.998 * x5_a)
    put('zscore', z, "Z' = 0.717X1+0.847X2+3.107X3+0.420X4+0.998X5（仅适用于非上市非金融企业筛查）")
    put('zscore_ann', zt, "Z'(年化) = 非全年期对收入与 EBIT 按期间年化倍数换算后计算")

    ms = []
    for i in range(N):
        pj = find_prior_same_kind(pmetas, i)
        if pj is None or PKIND[i] != 'FY':
            ms.append(None); continue
        def r_ratio(row, j, base):
            return sdiv(V(row, j) if has(row, j) else None, V(base, j) if has(base, j) else None)
        dsri = sdiv(r_ratio(A['ar'], i, P['rev']), r_ratio(A['ar'], pj, P['rev']))
        gmi = sdiv(M['mll'][pj], M['mll'][i])
        def aq(j):
            ta = M['zzc'][j]
            if not ta: return None
            return 1.0 - (V(A['ca'], j) + V(A['fa'], j) + V(A['lt_eq'], j)) / ta
        aqi = sdiv(aq(i), aq(pj))
        sgi = sdiv(V(P['rev'], i) if has(P['rev'], i) else None, V(P['rev'], pj) if has(P['rev'], pj) else None)
        depi = 1.0
        sga_i = sdiv(V(P['sell'], i) + V(P['admin'], i), V(P['rev'], i) if has(P['rev'], i) else None)
        sga_p = sdiv(V(P['sell'], pj) + V(P['admin'], pj), V(P['rev'], pj) if has(P['rev'], pj) else None)
        sgai = sdiv(sga_i, sga_p)
        lvgi = sdiv(M['zcfzl'][i], M['zcfzl'][pj])
        tata = sdiv((V(P['ni'], i) - M['jyxj'][i]) if (has(P['ni'], i) and M['jyxj'][i] is not None) else None, M['zzc'][i])
        if None in (dsri, gmi, aqi, sgi, sgai, lvgi, tata):
            ms.append(None)
        else:
            ms.append(-4.84 + 0.920 * dsri + 0.528 * gmi + 0.404 * aqi + 0.892 * sgi
                      + 0.115 * depi - 0.172 * sgai + 4.679 * tata - 0.327 * lvgi)
    put('mscore', ms, "M-Score（仅 FY 连续可比年度筛查；缺累计折旧时 DEPI 置 1 并标注）")
    put('tata', [sdiv((V(P['ni'], i) - M['jyxj'][i]) if (has(P['ni'], i) and M['jyxj'][i] is not None) else None, M['zzc'][i]) for i in range(N)],
        'TATA =（净利润 − 经营活动净额）/ 资产总计')

    # --- 9. 规则信号与 DQ 门禁 ---
    BOOL['DQ005'] = dq005
    BOOL['F1'] = [(sdiv(V(A['cash'], i), M['zzc'][i]) or 0) >= 0.20 and (sdiv(xy[i], M['zzc'][i]) or 0) >= 0.25 for i in range(N)]
    BOOL['F2_oar_ta'] = [sdiv(V(A['oar'], i), M['zzc'][i]) or 0.0 for i in range(N)]
    BOOL['F7'] = [(has(P['ni'], i) and V(P['ni'], i) > 0 and M['jyxj'][i] is not None and M['jyxj'][i] < 0) for i in range(N)]
    BOOL['F9'] = [((sdiv(abs(V(P['impair'], i)), M['zzc'][i]) or 0) > 0.05 and (sdiv(V(A['inv'], i), M['zzc'][i]) or 0) > 0.30) for i in range(N)]
    BOOL['F10'] = [((M['mll'][i] or 0) > 0.50 and M['jxb'][i] is not None and M['jxb'][i] < 0.8) for i in range(N)]
    BOOL['D1'] = [(M['zcfzl'][i] or 0) > 0.85 for i in range(N)]
    BOOL['D2'] = [(M['ldbl'][i] is not None and M['ldbl'][i] < 1.0) for i in range(N)]
    BOOL['D4'] = [(zt[i] is not None and zt[i] <= 1.23) for i in range(N)]
    BOOL['D5'] = [(has(P['ni'], i) and V(P['ni'], i) < 0) for i in range(N)]
    BOOL['D6'] = [(M['sdbl'][i] is not None and M['sdbl'][i] < 0.5) for i in range(N)]

    # --- 10. 全覆盖与行向表补强（动态化） ---
    put('qjfy', [(V(P['sell'], i) + V(P['admin'], i) + V(P['fin'], i)) if has(P['rev'], i) else None for i in range(N)],
        '期间费用合计 = 销售费用 + 管理费用 + 财务费用（不含留空研发费用）')
    put('ca_ratio', [sdiv(M['ldzc'][i], M['zzc'][i]) for i in range(N)], '流动资产 / 资产总计')
    put('nca_ratio', [sdiv(M['fldzc'][i], M['zzc'][i]) for i in range(N)], '非流动资产 / 资产总计')
    put('dc_yx_ratio', [sdiv(st[i], xy[i]) for i in range(N)], '短期有息负债 / 狭义有息负债')
    put('gy_tzsy', [(V(P['other_inc'], i) + V(P['inv_inc'], i) + V(P['fv_inc'], i)) for i in range(N)],
        '广义投资收益 = 其他收益 + 投资收益 + 公允价值变动收益（不含利息收入）')
    put('non_op_net', fjy, '营业外收支净额')
    put('hbzj', [V(A['cash'], i) if has(A['cash'], i) else None for i in range(N)], '货币资金（未扣受限）')
    put('cf_fx', [V(C['fx'], i) if has(C['fx'], i) else None for i in range(N)], '汇率变动对现金影响')
    put('cf_cash_net', [V(C['net_inc'], i) if has(C['net_inc'], i) else None for i in range(N)], '现金及现金等价物净增加额')
    put('wca_less_cl', [(wca[i] - M['ldfz'][i]) if M['ldfz'][i] is not None else None for i in range(N)], '营运资产 − 流动负债')
    put('ar_imp5_pct', [sdiv(V(A['ar'], i) * 0.05, V(P['ni'], i) if has(P['ni'], i) else None) for i in range(N)], '应收按5%计提减值 / 同期净利润')
    put('cip_over_fa', [sdiv(V(A['cip'], i), V(A['fa'], i) if has(A['fa'], i) else None) for i in range(N)], '在建工程 / 固定资产')
    put('wx_deduct', wx_res, '倒减口径无息负债 = 负债合计 − 狭义有息')
    put('wx_gap', [(wx_sum[i] - wx_res[i]) if wx_res[i] is not None else None for i in range(N)], '加总无息 − 倒减无息（差额检验被错桶科目）')
    put('cum_divint5', [c_int if (anchor_fy is not None and i == anchor_fy) else None for i in range(N)], fy_cf_span_label + ' 累计分配股利或偿付利息现金')
    put('growth_gap', [(M['yoy_ta'][i] - M['sgr'][i]) if (M['yoy_ta'][i] is not None and M['sgr'][i] is not None) else None for i in range(N)],
        '增长缺口 = 实际总资产增长率 − SGR')
    put('yx_xy_yoy', [sdiv(xy[i], xy[find_prior_same_kind(pmetas, i)]) - 1.0 if (find_prior_same_kind(pmetas, i) is not None and xy[find_prior_same_kind(pmetas, i)]) else None for i in range(N)], '狭义有息负债同比')
    put('yx_gy_yoy', [sdiv(gy[i], gy[find_prior_same_kind(pmetas, i)]) - 1.0 if (find_prior_same_kind(pmetas, i) is not None and gy[find_prior_same_kind(pmetas, i)]) else None for i in range(N)], '广义有息负债同比')
    put('st_less_wc', [(st[i] - M['yyzb'][i]) if M['yyzb'][i] is not None else None for i in range(N)], '短期有息负债 − 营运资本')
    put('eq_in', [V(C['eq_in'], i) if has(C['eq_in'], i) else None for i in range(N)], '吸收投资收到的现金')

    eqchg, cap_add, res_add, eq_sum, eq_diff = [], [], [], [], []
    for i in range(N):
        op_i = find_opening_bs_index(pmetas, i)
        if op_i is None or M['qy'][i] is None or M['qy'][op_i] is None:
            eqchg.append(None); cap_add.append(None); res_add.append(None); eq_sum.append(None); eq_diff.append(None)
            continue
        de = M['qy'][i] - M['qy'][op_i]
        dc = V(A['capital'], i) - V(A['capital'], op_i)
        dr = V(A['reserve'], i) - V(A['reserve'], op_i)
        s = V(P['ni'], i) + dc + dr
        eqchg.append(de); cap_add.append(dc); res_add.append(dr); eq_sum.append(s); eq_diff.append(de - s)
    put('eq_change', eqchg, '权益变动 = 期末权益 − 期初权益(上年末)')
    put('cap_add', cap_add, '实收资本增加')
    put('res_add', res_add, '专项储备增加')
    put('eq_roll_sum', eq_sum, '权益变动合计 = 本期净利润 + 实收资本增 + 专项储备增')
    put('eq_roll_diff', eq_diff, '权益变动差额 = 权益变动 − 合计')

    put('ni', [V(P['ni'], i) if has(P['ni'], i) else None for i in range(N)], '本期净利润')
    put('prepay_over_ta', [sdiv(V(A['prepay'], i), M['zzc'][i]) for i in range(N)], '预付款项 / 资产总计')
    put('ar_over_rev', [sdiv(V(A['ar'], i), V(P['rev'], i) if has(P['rev'], i) else None) for i in range(N)], '应收账款 / 同期营业收入')
    put('cip_over_ta', [sdiv(V(A['cip'], i), M['zzc'][i]) for i in range(N)], '在建工程 / 资产总计')
    put('lt_funds', cqzj, '长期资金 = 所有者权益 + 非流动负债')
    put('lt_funds_over_nca', [sdiv(cqzj[i], M['fldzc'][i]) for i in range(N)], '长期资金 / 非流动资产')
    put('lt_funds_surplus', [(cqzj[i] - M['fldzc'][i]) if (cqzj[i] is not None and M['fldzc'][i] is not None) else None for i in range(N)], '长期资金富余')
    put('adj_neg_ltd', [V(A['ltd'], i) for i in range(N)], '长期待摊费用扣减幅值')
    put('adj_neg_cip', [0.5 * V(A['cip'], i) for i in range(N)], '50%×在建工程扣减幅值')
    put('adj_up', [(M['zcfzl_adj'][i] - M['zcfzl'][i]) if (M['zcfzl_adj'][i] is not None and M['zcfzl'][i] is not None) else None for i in range(N)],
        '调整后负债率 − 账面负债率（pp）')
    put('ar_rev_gap5', [(M['yoy_ar'][i] - M['yoy_rev'][i]) if (M['yoy_ar'][i] is not None and M['yoy_rev'][i] is not None) else None for i in range(N)],
        '应收同比 − 收入同比（pp，全可比年度动态计算）')
    put('prepay_over_ap', [sdiv(V(A['prepay'], i), V(A['ap'], i) if has(A['ap'], i) else None) for i in range(N)], '预付款项 / 应付账款（倍）')

    # 非全年期科目较上年末变动（非同比）
    prepay_chg_fy, ar_chg_fy = [None] * N, [None] * N
    for i in range(N):
        if PKIND[i] != 'FY':
            op_i = find_opening_bs_index(pmetas, i)
            if op_i is not None:
                prepay_chg_fy[i] = (V(A['prepay'], i) / V(A['prepay'], op_i) - 1.0) if V(A['prepay'], op_i) else None
                ar_chg_fy[i] = (V(A['ar'], i) / V(A['ar'], op_i) - 1.0) if V(A['ar'], op_i) else None
    put('prepay_chg5', prepay_chg_fy, '非全年期预付款项较上年末变动（非同比）')
    put('ar_chg5', ar_chg_fy, '非全年期应收账款较上年末变动（非同比）')
    put('rd_blank', [0 if has(P['rd'], i) else 1 for i in range(N)], '研发费用留空标记（1=源表留空非0）')
    put('ocf_over_divint5', [c_cov if (anchor_fy is not None and i == anchor_fy) else None for i in range(N)],
        fy_cf_span_label + ' 累计经营净额 / 累计分配股利或偿付利息现金')

    T['periods'] = PLAB
    T['period_kinds'] = PKIND
    T['period_days'] = S['_period_days']
    T['fy_indices'] = fy_indices
    T['last_fy'] = last_fy
    return M, F, BOOL, T, A, P, C


# ---------------------------------------------------------------- 六类校验 + DQ 门禁
def f0(x):
    return '—' if x is None else format(x, ',.2f')


def checks(S, M, BOOL, T, A, P, C):
    PLAB = T['periods']
    PKIND = T['period_kinds']
    pmetas = S['_pmetas']
    N = len(PLAB)
    L = ['# 六类校验与数据质量门禁记录（脚本自动生成）\n']
    L.append('- 识别期间（%d 期）：%s' % (N, ', '.join('%s(%s/%d天)' % (PLAB[i], PKIND[i], int(T['period_days'][i])) for i in range(N))))
    L.append('- 各表实际覆盖期间：资产负债表 %d 期 / 利润表 %d 期 / 现金流量表 %d 期\n' % (
        len(S['_sheet_periods']['bs']), len(S['_sheet_periods']['is']), len(S['_sheet_periods']['cf'])))
    ok = lambda b: '✅ 通过' if b else '❌ 不通过'

    # DQ005 门禁提示
    dq5_periods = [PLAB[i] for i in range(N) if BOOL['DQ005'][i]]
    if dq5_periods:
        L.append('> ⚠️ **DQ005 阻断警告**：期间 `%s` 存在有息负债且财务费用 > 0，但“利息费用”明细缺失或为 0。按严格口径，`yhlx_strict` / `hxlr_strict` 已置 `NOT_COMPUTABLE`，严禁在报告中无条件以财务费用替代利息费用下结论。\n' % ', '.join(dq5_periods))

    # C1 原表勾稽
    L.append('## C1 原表勾稽校验')
    L.append('\n判定标准：差异绝对值 < 0.005 元。\n')
    L.append('| 期间 | 资产=负债+权益 | 流动+非流动=资产 | 负债+权益=总计 | 差异(元) |')
    L.append('|---|---|---|---|---|')
    allok = True
    for i in range(N):
        if M['zzc'][i] is None:
            L.append('| %s | 缺失 | 缺失 | 缺失 | — |' % PLAB[i])
            continue
        d1 = V(A['ta'], i) - (V(A['tl'], i) + V(A['eq'], i))
        d2 = V(A['ta'], i) - (V(A['ca'], i) + V(A['nca'], i))
        d3 = (V(A['tle'], i) - V(A['ta'], i)) if has(A['tle'], i) else 0.0
        z = lambda x: abs(x) < 0.005
        b = z(d1) and z(d2) and z(d3)
        allok &= b
        L.append('| %s | %s | %s | %s | %s / %s / %s |' % (PLAB[i], ok(z(d1)), ok(z(d2)), ok(z(d3)), f0(d1), f0(d2), f0(d3)))
    L.append('\n**结论**：%s\n' % ('全部精确闭合（差异 < 0.005 元）' if allok else '❌ 存在不闭合，必须先解决再分析'))

    # C2 口径完整性
    L.append('## C2 口径完整性校验（加总 vs 倒减，防 FM-02）')
    L.append('\n| 期间 | 加总无息 | 长期应付款 | 狭义有息 | 三者之和 | 负债合计 | 差额 | 倒减残差 | 加总vs倒减差 |')
    L.append('|---|---|---|---|---|---|---|---|---|')
    for i in range(N):
        if M['fz'][i] is None:
            continue
        tot = M['wx_sum'][i] + V(A['lt_pay'], i) + M['yx_xy'][i]
        diff = tot - M['fz'][i]
        gap = M['wx_sum'][i] - M['wx_res'][i]
        L.append('| %s | %s | %s | %s | %s | %s | %.2f | %s | %s |' % (
            PLAB[i], f0(M['wx_sum'][i]), f0(V(A['lt_pay'], i)), f0(M['yx_xy'][i]), f0(tot), f0(M['fz'][i]),
            diff, f0(M['wx_res'][i]), f0(gap)))
    L.append('')

    # C3 数量级合理性
    L.append('## C3 数量级合理性校验（调整类指标，防 FM-03）')
    L.append('\n| 期间 | 资产总计 | 应扣减额 | 占资产比 | 调整后资产 | 调整前负债率 | 调整后负债率 | 上升(pp) |')
    L.append('|---|---|---|---|---|---|---|---|')
    for i in range(N):
        if M['zzc'][i] is None or M['zcfzl'][i] is None or M['zcfzl_adj'][i] is None:
            continue
        cut = V(A['ltd'], i) + 0.5 * V(A['cip'], i)
        L.append('| %s | %s | %s | %.2f%% | %s | %.2f%% | %.2f%% | %.2f |' % (
            PLAB[i], f0(M['zzc'][i]), f0(cut), cut / M['zzc'][i] * 100, f0(M['adj_ta'][i]),
            M['zcfzl'][i] * 100, M['zcfzl_adj'][i] * 100, (M['zcfzl_adj'][i] - M['zcfzl'][i]) * 100))
    L.append('')

    # C4 方向一致性
    L.append('## C4 方向一致性校验（防 FM-01 / FM-04）')
    L.append('\n| 指标 | ' + ' | '.join(PLAB) + ' | 方向 |')
    L.append('|---|' + '---|' * N + '---|')
    zvals = ['—' if M['zcfzl'][i] is None else '%.2f%%' % (M['zcfzl'][i] * 100) for i in range(N)]
    valid_z = [v for v in M['zcfzl'] if v is not None]
    dir_txt = ('终点低于起点' if valid_z[-1] < valid_z[0] else '终点高于/等于起点') if len(valid_z) >= 2 else '期数不足'
    L.append('| 资产负债率 | %s | %s |' % (' | '.join(zvals), dir_txt))
    c4, c4e, c4l = M['cagr4_ta'][0], M['cagr4_eq'][0], M['cagr4_tl'][0]
    if None not in (c4, c4e, c4l):
        rank = sorted([('负债', c4l), ('权益', c4e), ('资产', c4)], key=lambda x: -x[1])
        L.append('\n完整年度 CAGR：负债 %.2f%%、权益 %.2f%%、资产 %.2f%% → 排序 **%s > %s > %s**。\n'
                 % (c4l * 100, c4e * 100, c4 * 100, rank[0][0], rank[1][0], rank[2][0]))
    else:
        L.append('\n完整年度不足 2 期，跳过跨年 CAGR 排序比对。\n')

    # C5 现金流勾稽
    L.append('## C5 现金流量表勾稽校验')
    L.append('\n| 期间 | 期初+净增加 与 期末差额 | 期末现金 与 货币资金差额 | 三流+汇率 与 净增加差额 |')
    L.append('|---|---|---|---|')
    for i in range(N):
        if not has(C['close_cash'], i):
            L.append('| %s | 数据缺失 | 数据缺失 | 数据缺失 |' % PLAB[i])
            continue
        a = V(C['open_cash'], i) + V(C['net_inc'], i) - V(C['close_cash'], i)
        b = V(C['close_cash'], i) - V(A['cash'], i)
        c = (V(C['ocf'], i) + V(C['icf'], i) + V(C['fcf_net'], i) + V(C['fx'], i)) - V(C['net_inc'], i)
        L.append('| %s | %.2f | %.2f | %.2f |' % (PLAB[i], a, b, c))
    L.append('')

    # C6 权益变动勾稽
    L.append('## C6 权益变动勾稽校验（对比上年末期初权益）')
    L.append('\n| 期间 | 期间性质 | 权益变动 | 本期净利润 | 实收资本增 | 专项储备增 | 合计 | 差额 |')
    L.append('|---|---|---|---|---|---|---|---|')
    for i in range(N):
        if M['eq_change'][i] is None:
            continue
        L.append('| %s | %s | %s | %s | %s | %s | %s | %s |' % (
            PLAB[i], PKIND[i], f0(M['eq_change'][i]), f0(M['ni'][i]),
            f0(M['cap_add'][i]), f0(M['res_add'][i]), f0(M['eq_roll_sum'][i]), f0(M['eq_roll_diff'][i])))
    L.append('')
    return '\n'.join(L)


def out_metrics_md(M, F, BOOL, T):
    PLAB = T['periods']
    N = len(PLAB)
    L = ['# 派生指标复算结果（全精度中间值）\n',
         '> 金额单位与原表一致（报告展示时按口径换算并由全精度一次舍入）；比率保留 4 位小数。\n',
         '> 本文件及 `metrics.json` 是报告的**唯一数值来源**。\n']
    L.append('## 一、指标一览（%d 期）\n' % N)
    L.append('| 指标 | ' + ' | '.join(PLAB) + ' |')
    L.append('|---|' + '---|' * N)
    for k, vals_raw in M.items():
        vals = []
        for v in vals_raw:
            if v is None:
                vals.append('—')
            elif isinstance(v, str):
                vals.append(v)
            elif abs(v) >= 1000:
                vals.append(format(v, ',.2f'))
            else:
                vals.append('%.4f' % v)
        L.append('| %s | %s |' % (k, ' | '.join(vals)))
    L.append('\n## 二、公式与口径说明\n')
    for k in M:
        L.append('- **%s**：%s' % (k, F.get(k, '')))
    L.append('\n## 三、门禁与信号明细\n')
    L.append('| 门禁/规则 | ' + ' | '.join(PLAB) + ' |')
    L.append('|---|' + '---|' * N)
    for k, v in BOOL.items():
        if k == 'F2_oar_ta':
            L.append('| F2 其他应收款/总资产 | %s |' % ' | '.join('%.2f%%' % (x * 100) for x in v))
        else:
            L.append('| %s | %s |' % (k, ' | '.join('触发' if x else '未触发' for x in v)))
    L.append('')
    return '\n'.join(L)


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    xlsx = sys.argv[1]
    outdir = sys.argv[2]
    company = '未提供'
    if '--company' in sys.argv:
        company = sys.argv[sys.argv.index('--company') + 1]
    os.makedirs(outdir, exist_ok=True)
    S = load(xlsx)
    M, F, BOOL, T, A, P, C = compute(S)
    json_out = {}
    for k, v in M.items():
        json_out[k] = [None if x is None else x for x in v]
    json_out['_periods'] = T['periods']
    json_out['_period_kinds'] = T['period_kinds']
    json_out['_period_days'] = T['period_days']
    json_out['_sheet_periods'] = S['_sheet_periods']
    json_out['_company'] = company
    with io.open(os.path.join(outdir, 'metrics.json'), 'w', encoding='utf-8') as f:
        json.dump(json_out, f, ensure_ascii=False, indent=1)
    with io.open(os.path.join(outdir, 'metrics.md'), 'w', encoding='utf-8') as f:
        f.write(out_metrics_md(M, F, BOOL, T))
    with io.open(os.path.join(outdir, 'checks.md'), 'w', encoding='utf-8') as f:
        f.write(checks(S, M, BOOL, T, A, P, C))
    print('OK ->', outdir)
    print('Periods:', list(zip(T['periods'], T['period_kinds'])))
    print('DQ005 triggered:', [T['periods'][i] for i, b in enumerate(BOOL['DQ005']) if b] or '无')


if __name__ == '__main__':
    main()
