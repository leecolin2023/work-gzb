#!/usr/bin/env python3
"""Recompute a comprehensive-factor valuation and produce a sensitivity matrix.

Deterministic arithmetic only. Feed it a params JSON assembled from the
evidence index (never from memory of the report's conclusion).

params.json schema:
{
  "claim_amount": 135903299.82,
  "principal": 109301203.57,
  "priority_value": 0,
  "components": [
    {"name": "一般受偿", "score": 14.5, "weight": 0.5, "reported": 9853000},
    {"name": "保证受偿", "score": 28.6, "weight": 0.5, "reported": 19434200}
  ],
  "other_value": 0,
  "reported_total": 29287200,
  "years": 2,
  "discount_rate": 0.05,
  "reported_discounted": 26564400,
  "reported_principal_recovery": 0.2430,
  "reported_claim_recovery": 0.1955,
  "round_to": 100
}

Usage:
  python verify_calc.py --params params_07.json --out verify_07.json [--sensitivity] [--delta 5] [--print]
"""
import argparse
import copy
import json
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


def rnd(v, unit):
    return round(v / unit) * unit if unit else v


def compute(p):
    unit = p.get("round_to", 0) or 0
    base = p["claim_amount"] - p.get("priority_value", 0)
    rows = []
    for c in p.get("components", []):
        rate = c["score"] / 100.0 * c["weight"]
        raw = base * rate
        rows.append({"name": c["name"], "score": c["score"], "weight": c["weight"],
                     "rate": rate, "raw": raw, "rounded": rnd(raw, unit),
                     "reported": c.get("reported")})
    total_rounded = p.get("priority_value", 0) + sum(r["rounded"] for r in rows) + p.get("other_value", 0)
    total_raw = p.get("priority_value", 0) + sum(r["raw"] for r in rows) + p.get("other_value", 0)
    factor = (1.0 + p["discount_rate"]) ** p["years"]
    discounted = total_rounded / factor
    return {
        "base": base, "rows": rows, "total_rounded": total_rounded, "total_raw": total_raw,
        "factor": factor, "discounted": discounted,
        "principal_recovery": discounted / p["principal"] if p.get("principal") else None,
        "claim_recovery": discounted / p["claim_amount"] if p.get("claim_amount") else None,
    }


def check(step, formula, recomputed, reported, tolerance):
    diff = None if reported is None else recomputed - reported
    material = None if diff is None else abs(diff) > tolerance
    return {"step": step, "formula": formula, "recomputed": recomputed,
            "reported": reported, "diff": diff, "material": material,
            "exceeds_tolerance": material}


def build_checks(p, r):
    tol = max(p.get("round_to", 0) or 1, 1)
    out = []
    for row in r["rows"]:
        out.append(check(row["name"] + "率",
                         "%.4g / 100 × %.4g" % (row["score"], row["weight"]),
                         row["rate"], None, tol))
        out.append(check(row["name"] + "额",
                         "%.2f × %.6f" % (r["base"], row["rate"]),
                         row["rounded"], row["reported"], tol))
    out.append(check("四项加总", "优先 + 各受偿项 + 其他",
                     r["total_rounded"], p.get("reported_total"), tol))
    out.append(check("贴现", "%.2f ÷ (1+%.4g)^%g" % (r["total_rounded"], p["discount_rate"], p["years"]),
                     r["discounted"], p.get("reported_discounted"), tol))
    for key, label, formula in (
        ("principal_recovery", "本金受偿率", "贴现后估值 ÷ 贷款本金"),
        ("claim_recovery", "债权受偿率", "贴现后估值 ÷ 债权金额"),
    ):
        reported = p.get("reported_" + key)
        out.append(check(label, formula, r[key],
                         reported, 0.0001 if reported is not None and abs(reported) < 1 else tol))
    return out


def sensitivity(p, base_r, delta):
    scenarios = []
    for i, c in enumerate(p.get("components", [])):
        for sign in (+1, -1):
            q = copy.deepcopy(p)
            q["components"][i]["score"] = max(0, min(100, c["score"] + sign * delta))
            q["components"][i]["reported"] = None
            scenarios.append(("%s评分 %+.4g (%.4g→%.4g)" % (c["name"], sign * delta, c["score"], q["components"][i]["score"]), q))
    for i, c in enumerate(p.get("components", [])):
        for sign in (+1, -1):
            q = copy.deepcopy(p)
            nw = max(0, min(1, c["weight"] + sign * 0.1))
            q["components"][i]["weight"] = nw
            for row in q["components"]:
                row["reported"] = None
            scenarios.append(("%s影响程度 %+.2g (%.2g→%.2g)" % (c["name"], sign * 0.1, c["weight"], nw), q))
    for sign in (+1, -1):
        q = copy.deepcopy(p)
        q["years"] = max(0, p["years"] + sign)
        scenarios.append(("处置周期 %+d 年 (%g→%g)" % (sign, p["years"], q["years"]), q))
    for sign in (+1, -1):
        q = copy.deepcopy(p)
        q["discount_rate"] = max(0, p["discount_rate"] + sign * 0.05)
        scenarios.append(("贴现率 %+.2f%% (%.2f%%→%.2f%%)" % (sign * 5, p["discount_rate"] * 100, q["discount_rate"] * 100), q))

    out = []
    for label, q in scenarios:
        rr = compute(q)
        out.append({"scenario": label,
                    "total": rr["total_rounded"],
                    "discounted": rr["discounted"],
                    "delta_total": rr["total_rounded"] - base_r["total_rounded"],
                    "delta_discounted": rr["discounted"] - base_r["discounted"]})
    return out


def fmt(v, nd=2):
    if v is None:
        return "—"
    if isinstance(v, float):
        return ("%%.%df" % nd) % v
    return str(v)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--params", required=True)
    ap.add_argument("--out")
    ap.add_argument("--sensitivity", action="store_true")
    ap.add_argument("--delta", type=float, default=5)
    ap.add_argument("--print", action="store_true")
    args = ap.parse_args()

    with open(args.params, encoding="utf-8") as f:
        p = json.load(f)
    if p.get("model", "legacy_score_discount") != "legacy_score_discount":
        sys.exit("unsupported model: use verify_formulas.py for formula_graph; do not force legacy formulas")
    for key in ("claim_amount", "principal", "years", "discount_rate"):
        if key not in p:
            sys.exit("missing required param: " + key)

    r = compute(p)
    checks = build_checks(p, r)
    result = {"params": p, "recomputed": r, "checks": checks,
              "tolerance_note": "exceeds_tolerance = |diff| > round_to (或率值 > 0.0001)；material为旧字段兼容别名，仅表示数值超容差，不判业务影响。",
              "scope": "只验算已确认适用的评分权重及统一折现模型；参数依据、代入适用性与业务影响需另行核验。"}
    if args.sensitivity:
        result["sensitivity"] = sensitivity(p, r, args.delta)

    text = json.dumps(result, ensure_ascii=False, indent=1)
    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
        print("written:", args.out)

    if args.print or not args.out:
        print("| 步骤 | 复核结果 | 报告值 | 差异 | 超数值容差 |")
        print("|---|---:|---:|---:|---|")
        for c in checks:
            nd = 6 if "率" in c["step"] else 2
            print("| %s | %s | %s | %s | %s |" % (
                c["step"], fmt(c["recomputed"], nd), fmt(c["reported"], nd),
                fmt(c["diff"], nd), "—" if c["material"] is None else ("是" if c["material"] else "否")))
        if args.sensitivity:
            print("\n| 情景 | 贴现后估值 | 变化 |")
            print("|---|---:|---:|")
            for s in result["sensitivity"]:
                print("| %s | %s | %s |" % (s["scenario"], fmt(s["discounted"]), fmt(s["delta_discounted"])))


if __name__ == "__main__":
    main()
