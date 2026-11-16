#!/usr/bin/env python3
"""Review-item registry and semantic anchor recall.

CHECKS and EXTENDED_CHECKS are the machine-readable form of the three-layer review framework in
references/review-checklist.md. Reports with different chapter structures are
handled by locating evidence through these anchors against a flat evidence
index produced by extract_report.py.

Zero hits => 待定位, never a disclosure or factual judgment. Read full context and external sources.

Usage:
  python anchors.py --list
  python anchors.py --evidence evidence_07.json --out recall_07.json [--layer L2] [--limit 8]
"""
import argparse
import json
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


def C(cid, layer, question, anchors, regex=None, external=False):
    """external=True marks items whose decisive evidence lives OUTSIDE the report
    (查册成果 / 银行清单 / 归档). Zero recall for these must not be treated as
    报告未披露; it means the external evidence still has to be read."""
    return {"id": cid, "layer": layer, "question": question,
            "anchors": anchors, "regex": regex or [], "external": external}


CHECKS = [
    # ---------------- L1 基础事实 ----------------
    C("F01", "L1", "报告身份、委托人、基准日、出具日、有效期是否明确且自洽",
      ["价值分析报告", "评估报告", "报告编号", "评字", "委托人", "债权持有人", "分析基准日", "评估基准日",
       "基准日", "提交日期", "出具日期", "使用有效期", "评估报告使用有效期", "资产评估机构"]),
    C("F02", "L1", "债权本金、利息、罚息复利、代垫费用、合计能否与档案勾稽",
      ["债权本金", "贷款本金", "本金余额", "欠息", "利息", "罚息", "复利", "代垫费用", "代垫评估费",
       "债权总额", "债权金额", "债权合计", "合计", "余额"],
      [r"\d{1,3}(?:,\d{3})+(?:\.\d+)?"]),
    C("F03", "L1", "放款事实是否成立，合同额/实际放款额/当前余额是否分列",
      ["借款合同", "固定资产借款合同", "放款凭证", "实际发放", "发放贷款", "贷款金额", "贷款期限",
       "合同编号", "签署", "借据"]),
    C("F04", "L1", "重组/展期/和解协议链是否完整、最新版本是否被采用",
      ["债务重组", "重组", "展期", "和解协议", "协议书", "补充协议", "还款安排", "提前到期",
       "贷款提前到期通知书", "一次性结清", "分期还本"]),
    C("F05", "L1", "债务人、抵质押人、保证人主体是否准确、存续",
      ["债务人概况", "抵押人概况", "质押人", "保证人", "工商登记", "统一社会信用代码", "企业登记状态",
       "法定代表人", "股东名称", "存续", "注销", "失信被执行人", "限制高消费", "被执行"]),
    C("F06", "L1", "担保合同类型、编号、担保范围、保证期间是否与档案一致",
      ["担保方式", "保证合同", "抵押合同", "质押合同", "权利质押合同", "连带责任保证", "保证期间",
       "担保范围", "保证担保范围", "担保权利行使顺序", "质押+保证", "抵押+保证"]),
    C("F07", "L1", "担保物权是否依法设立、顺位与担保范围",
      ["质权登记", "股权出质", "登记编号", "抵押登记", "他项权证", "不动产权证", "不动产证明号",
       "优先受偿权", "抵押顺位", "在先抵押", "最高额抵押", "登记设立"]),
    C("F08", "L1", "诉讼、判决、保全、执行进展是否与档案一致",
      ["判决书", "裁定书", "案号", "民初", "执恢", "执保", "诉讼", "执行", "财产保全", "查封", "冻结",
       "移送执行函", "恢复执行", "判项", "申请强制执行"]),
    C("F09", "L1", "已处置财产与已清偿款项是否已从余额扣减",
      ["已处置", "拍卖", "变卖", "成交", "分配", "清偿", "偿还本金", "归还本金", "已用于清偿",
       "款项已用于", "还款", "本金递减"]),
    C("F10", "L1", "诉讼时效、执行时效、债权是否仍存续",
      ["诉讼时效", "执行时效", "申请执行", "时效中断", "债权存续", "结清", "展期", "借新还旧",
       "续封", "续行保全", "债务履行期限"], ),

    # ---------------- L2 价值来源 ----------------
    C("V01", "L2", "四项受偿是否各有明确金额与来源，加总是否等于综合可回收价值",
      ["优先受偿债权价值", "一般受偿债权价值", "保证受偿债权价值", "其他受偿", "综合可回收价值",
       "债权价值＝", "债权价值=", "加总", "合计"]),
    C("V02", "L2", "优先受偿来自哪些抵质押物，逐项估值与变现依据",
      ["抵押物优先受偿", "质押物优先受偿", "抵质押物", "变现价值", "评估值", "市场询价",
       "快速变现", "处置税费", "变现率", "逐套"]),
    C("V03", "L2", "优先受偿归零的理由是否成立",
      ["本项目无抵押", "无抵押担保", "无抵押物", "按0元考虑", "按0元计", "变现价值按", "优先受偿债权价值为0",
       "停工", "无可售物业", "无实际经营性现金流"]),
    C("V04", "L2", "一般受偿来自债务人哪部分偿债能力",
      ["一般偿债能力", "一般受偿债权价值", "债务人偿债能力", "名下有效资产", "财务报表", "经营状况",
       "无法获取", "现场走访", "未找到"]),
    C("V05", "L2", "保证受偿来自哪些保证人、哪部分可执行财产",
      ["保证偿债能力", "保证受偿债权价值", "保证人偿债能力", "保证人1", "保证人2", "连带责任",
       "名下有效资产", "资不抵债", "被执行标的"]),
    C("V06", "L2", "权利状态与受偿路径是否匹配（查封≠抵押）",
      ["不享有抵押权", "轮候查封", "在先抵押", "财产保全", "查封", "优先受偿", "顺位", "非我行抵押物",
       "查封资产清单", "协助执行通知书"]),
    C("V07", "L2", "同一财产是否被两条受偿路径重复计入",
      ["同时", "亦作为", "并作为", "该批", "上述财产", "名下有效资产", "优先受偿债权价值为0", "查封"]),
    C("V08", "L2", "跨户共用资产／共同保证人是否被多户重复引用",
      ["共同保证", "跨户", "共用", "同一保证人", "多户", "重复", "分配", "整包", "资产包"],
      external=True),
    C("V09", "L2", "报告与银行清单的资产性质是否冲突",
      ["抵押物清单", "非我行抵押物", "资产性质", "口径", "以登记簿为准", "清单", "需确认"],
      external=True),
    C("V10", "L2", "查册覆盖率是否支撑资产端结论",
      ["查册", "覆盖率", "登记信息查询结果", "未匹配", "待补", "逐套", "全量结论", "扫描件"],
      external=True),

    # ---------------- L3 参数与测算 ----------------
    C("P01", "L3", "采用何种估值方法、适用条件是否满足",
      ["分析方法", "估值方法", "综合因素分析法", "综合因素分析", "现金流", "市场法", "清算价值",
       "分析方法是指", "采用综合因素"]),
    C("P02", "L3", "价值类型与后续参数是否自洽",
      ["价值类型", "清算价值", "清算价值是指", "市场价值", "投资价值", "可变现价值", "可回收价值",
       "被迫出售", "快速变现"]),
    C("P03", "L3", "评分表维度、满分、赋分是否完整且满分合计 100",
      ["分析项目", "分类", "实际情况", "最高分值", "赋分值", "小计", "100", "偿债能力分析"]),
    C("P04", "L3", "每项赋分是否有可定位依据",
      ["本息偿还情况", "工商年检情况", "经营地址查询情况", "经审计财务报告", "所属行业", "企业性质",
       "企业规模", "目前经营状况", "历史信用情况", "所处地域", "资产保全情况"]),
    C("P05", "L3", "最大权重项的判断痕迹",
      ["目前经营状况", "经营状况", "30", "现场走访", "未发现", "无法核实", "实地走访"]),
    C("P06", "L3", "影响程度／权重拆分的依据",
      ["偿债能力影响程度", "影响程度", "50%", "综合分析本债权项目", "注1", "注2", "权重"]),
    C("P07", "L3", "加总逻辑是否把连带保证降格为按份责任",
      ["债权价值＝", "债权价值=", "一般受偿债权价值＝", "保证人一般受偿债权价值＝", "连带", "保证责任",
       "不提出抗辩", "选择行使担保权利顺序"]),
    C("P08", "L3", "四项加总、贴现、两个受偿率是否算对",
      ["取整至百位", "贴现", "折现", "预期收益率", "本金受偿率", "本债权受偿率", "受偿率",
       "预计债权回收价值", "^2", "（1＋5%）", "(1+5%)"],
      [r"\d{1,3}(?:,\d{3})+(?:\.\d+)?"]),
    C("P09", "L3", "贴现率与处置周期的依据",
      ["预计处置", "需时", "处置需时", "回收期", "历年处置经验", "投资预期收益率", "贴现率", "折现率",
       "2年", "两年", "执行程序"]),
    C("P10", "L3", "受偿率分母口径",
      ["本金受偿率", "本债权受偿率", "债权受偿率", "可受偿金额÷贷款本金", "可受偿金额÷债权金额", "贷款本金"]),
    C("P11", "L3", "回收率与贴现率是否双松",
      ["市场投资预期收益率", "贴现率取", "综合可回收", "偿债能力评分", "风险", "折价"]),
    C("P12", "L3", "关键参数是否有敏感性分析",
      ["敏感性", "情景", "假设变动", "如发生变化", "分析结论将失效", "条件发生变化"]),

    # ---------------- L4 披露与结论一致性 ----------------
    C("D01", "L4", "摘要与正文的数额、价值类型、措辞是否一致",
      ["价值分析报告摘要", "摘要", "谨提请", "可变现价值", "综合可回收价值", "清算价值"]),
    C("D02", "L4", "特别事项说明是否与测算相互矛盾",
      ["特别事项说明", "未考虑", "受索取资料的限制", "无法核实", "无法获取", "未对", "本报告不具有鉴证作用"]),
    C("D03", "L4", "价值分析与价值评估措辞是否混用",
      ["价值分析结论", "价值评估结论", "价值分析业务", "价值评估业务", "适用性不同于", "专业意见"]),
    C("D04", "L4", "假设前提与资料限制是否充分披露",
      ["假设前提", "真实性、完整性", "资料有限性", "分析判断的主观性", "由委托方提供", "予以核对"]),
    C("D05", "L4", "使用限制、有效期、期后重大事项是否完整",
      ["使用限制", "报告使用有效期", "期后重大事项", "分析报告期后", "不能直接使用本分析结论", "重新确定"]),
    C("D06", "L4", "报告未披露项是否单独成表",
      ["未提供", "未见", "未取得", "需补充", "待核实", "未披露"]),
    C("D07", "L4", "结论边界是否声明",
      ["不等同于", "可实现价格", "不具有强制约束力", "不应当被认为", "不作保证", "仅供"]),
    C("D08", "L4", "报告使用人与用途是否限定",
      ["分析目的", "其他分析报告使用人", "报告使用人", "不能成为", "仅供委托人", "资产处置"]),
]


# Preserve 40 mother-item IDs; select applicable subchecks rather than forcing one model.
EXTENDED_CHECKS = [
    C("A01", "L1", "经办、银行原件、尽调与报告的债权和主体事实是否一致", ["经办", "反馈", "本金", "抵押人", "出质人", "保证人", "债务人"], external=True),
    C("A02", "L1", "资产权属、数量面积、用途及现状是否与经办和尽调一致", ["面积", "权属", "权利人", "坐落", "商铺", "在建", "出租", "占用"], external=True),
    C("A03", "L1", "欠付工程款等负担是否与无欠付假设冲突", ["工程款", "欠付", "拖欠", "承包人", "优先款", "未作定额扣减", "无欠款", "无负担"], external=True),
    C("A04", "L4", "经办/尽调事实与估值假设冲突是否查明依据和传导", ["假设", "前提", "未考虑", "未取得", "未扣", "不存在", "确权"], external=True),
    C("A05", "L4", "有依据的情况是否进入建议修订及影响分析", ["调整", "核实", "补充", "特别事项", "影响"], external=True),
    C("A06", "L1", "外部证据质量、版本、未知值及基准日资格是否核清", ["查册", "查询时间", "截止", "基准日", "未核实"], external=True),
    C("C01", "L3", "可比案例来源、时点、成交挂牌属性是否可核验", ["可比", "案例", "调查表", "信息来源", "三级市场", "挂牌", "成交", "询价"]),
    C("C02", "L3", "远距/跨区域案例选取理由和市场可比性是否充分", ["周边", "同类型", "商圈", "距离", "区位", "跨区域", "可比", "坐落"]),
    C("C03", "L3", "案例实物、权利、交易和时点差异是否合理修正", ["修正", "调整系数", "楼层", "临街", "含税", "免租", "新旧", "交易情况"]),
    C("C04", "L3", "案例修正、权重与最终租金/单价能否复算", ["租金", "平均", "权重", "单价", "综合确定", "取整", "调整后"]),
    C("C05", "L3", "案例、参数、估值公式与最终结果是否相互印证", ["收益价值", "报酬率", "增长率", "变现率", "测算表", "汇总", "净收益"]),
]
UPDATES = {
    "V03": ("抵质押物按零或未计值的理由是否逐项成立", ["0.00元", "0.00", "股权质押", "应收账款", "确权", "可受偿金额为0"]),
    "V07": ("同一财产是否重复贡献价值，剩余净值是否核实", ["有效资产查册", "不动产档", "剩余净值", "重复计值"]),
    "P03": ("实际评分、基本受偿率及因素调整表是否完整，合计是否按本模型成立", ["基本受偿率", "因素调整系数", "调整取值", "分值"]),
    "P04": ("评分、参数与案例取值是否有可定位依据", ["租金", "空置", "收益期", "报酬率", "净收益", "三级市场"]),
    "P07": ("多主体计算基数、权重、汇总及回收上限是否合理", ["剩余", "单名", "法人保证人", "自然人保证人", "以剩余", "测算基数"]),
    "P08": ("公式、符号单位、代入、逐步计算及最终汇总能否独立验证", ["公式", "代入", "处置费率", "债权回收金额", "收益价值", "税费", "计算说明"]),
}
for check in CHECKS:
    if check["id"] in UPDATES:
        check["question"], extras = UPDATES[check["id"]]
        check["anchors"].extend(extras)


def recall(items, check, limit=8):
    """Keyword/regex navigation with row context; all references remain original IDs."""
    rows, headers = {}, {}
    for it in items:
        if it.get("type") == "table_cell":
            rows.setdefault((it["table"], it["row"]), []).append(it)
            if it["row"] <= 2:
                headers.setdefault(it["table"], []).append(it)
    scored, seen = [], set()
    regs = [re.compile(r) for r in check.get("regex", [])]
    for it in items:
        key = (it.get("table"), it.get("row")) if it.get("type") == "table_cell" else it["id"]
        if key in seen:
            continue
        seen.add(key)
        context = rows[key] if isinstance(key, tuple) else [it]
        txt = " | ".join(x.get("text", "") for x in context)
        heading = " | ".join(x["text"] for x in headers.get(it.get("table"), []) if x["row"] == 1)
        search_text = txt + " | " + heading if isinstance(key, tuple) else txt
        matched = [a for a in check["anchors"] if a in search_text]
        matched.extend("REGEX:" + r.pattern for r in regs if r.search(txt))
        if not matched:
            continue
        scored.append({"id": it["id"], "evidence_ids": [x["id"] for x in context],
                       "type": "table_row_context" if isinstance(key, tuple) else it.get("type"),
                       "section_hint": it.get("section_hint", ""), "score": len(matched),
                       "matched": matched, "text": txt[:1200],
                       "table_evidence_ids": [x["id"] for x in items
                                              if it.get("table") is not None and x.get("table") == it["table"]],
                       "table_header_context": [{"id": x["id"], "text": x["text"]}
                                                for x in headers.get(it.get("table"), [])]})
    scored.sort(key=lambda x: (-x["score"], x["id"]))
    return scored[:limit]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base-only", action="store_true", help="only the original 40 mother checks")
    ap.add_argument("--list", action="store_true", help="print the review-item registry")
    ap.add_argument("--evidence", help="evidence_XX.json from extract_report.py")
    ap.add_argument("--out", help="output recall json")
    ap.add_argument("--layer", choices=["L1", "L2", "L3", "L4"])
    ap.add_argument("--limit", type=int, default=8)
    args = ap.parse_args()

    registry = CHECKS if args.base_only else CHECKS + EXTENDED_CHECKS
    if args.limit < 1:
        ap.error("--limit must be positive")
    if args.list:
        for c in registry:
            print("%-4s %-3s %s" % (c["id"], c["layer"], c["question"]))
        print("total:", len(registry))
        return
    if not args.evidence:
        sys.exit("--evidence required (or use --list)")

    with open(args.evidence, encoding="utf-8") as f:
        data = json.load(f)
    items = data.get("items", [])
    checks = [c for c in registry if not args.layer or c["layer"] == args.layer]
    out_checks = []
    uncovered = []
    external_pending = []
    for c in checks:
        cand = recall(items, c, args.limit)
        if c.get("external"):
            external_pending.append(c["id"])
        elif not cand:
            uncovered.append(c["id"])
        out_checks.append({**c, "candidates": cand})
    result = {
        "evidence": data.get("source"),
        "evidence_count": len(items),
        "layer": args.layer or "ALL",
        "checks": out_checks,
        "uncovered": uncovered,
        "location_pending": uncovered,
        "base_check_count": len([c for c in checks if c["id"][0] in "F V P D".split()]),
        "subcheck_count": len([c for c in checks if c["id"][0] in ("A", "C")]),
        "external_pending": external_pending,
        "note": "uncovered => 待定位; inspect full report, tables and extraction gaps before assigning a disclosure status. "
                "external_pending => decisive evidence is outside the report (查册分户报告 / audit_all.json / 银行清单); "
                "read those before concluding; zero recall here is NOT 报告未披露.",
    }
    text = json.dumps(result, ensure_ascii=False, indent=1)
    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
        print("checks=%d uncovered=%s external_pending=%s -> %s" % (
            len(checks), ",".join(uncovered) or "-", ",".join(external_pending) or "-", args.out))
    else:
        print(text)


if __name__ == "__main__":
    main()
