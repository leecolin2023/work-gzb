# -*- coding: utf-8 -*-
"""校验器回归单测：自指统计核对 + 动态期间解析与对齐。"""
import os, sys, datetime

SCRIPTS = os.path.dirname(os.path.abspath(__file__))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

import recalc
import verify_report as vr

cells, npass, nfail, unc, cov, n_scan = 872, 733, 0, 139, 84.1, 1366

CASES = [
    ("修复前的自指统计（旧分母 849/116/86.3%）",
     "**执行方式声明**：本报告按《通用财报分析提示词包 **v1.2**》执行。**真实结果：数值点 849 个，已验证 733 个，不一致 0 个，未覆盖 116 个，覆盖率 86.3%，跨章矛盾 0 处。**",
     True),
    ("修复前 ① 节的散文化覆盖矛盾",
     '> **"未覆盖" ≠ "不一致"**：机器覆盖到的那 86.3% 全部一致；剩下 13.7% 是**已知边界**。',
     True),
    ("修复前 ① 节的未覆盖占比",
     "> 那 **13.7% 未覆盖**的部分**不是零星遗漏**。",
     True),
    ("修复后同一句（正确值）",
     "**真实结果：数值点 872 个，已验证 733 个，不一致 0 个，未覆盖 139 个，覆盖率 84.1%，跨章矛盾 0 处。**",
     False),
    ("财务口径的「货币资金覆盖率」",
     "结合第四章 §4（短期有息占 80.9%、货币资金覆盖率 29.8%）与第五章（净利率 2.2%）。",
     False),
    ("显式等式（正确）",
     "> - **覆盖率 = 733 ÷ 872 = 84.1%**，跨章自相矛盾 **0 处**",
     False),
    ("显式等式（分母写错）",
     "> - **覆盖率 = 733 ÷ 849 = 84.1%**，跨章自相矛盾 **0 处**",
     True),
    ("历史引述（应跳过）",
     '> 上一版宣称"502 个数值点、覆盖率 100.0%、不一致 0 个"，',
     False),
    ("引用历史数值的表格行（应跳过）",
     '> | v2 | "502 个数值点、覆盖率 100.0%、不一致 0 个" | **真实覆盖率上限 57.8%（502 ÷ 868）** |',
     False),
    ("声明的扫描数与实测不符（1,362 vs 1,366）",
     "> - 扫描表格单元格 **1,362** 个，其中可解析为单一数字的数值点 **872** 个（**这是覆盖率的分母**）",
     True),
    ("声明的扫描数与实测一致（1,366）",
     "> - 扫描表格单元格 **1,366** 个，其中可解析为单一数字的数值点 **872** 个（**这是覆盖率的分母**）",
     False),
]

bad_all = 0
print("=== 一、自指统计核对回归单测 ===")
for name, line, should_catch in CASES:
    _, bad = vr.check_declared_stats([line], cells, npass, nfail, unc, cov, n_scan)
    caught = len(bad) > 0
    ok = (caught == should_catch)
    bad_all += 0 if ok else 1
    print("%-6s %-40s 期望%s / 实际%s" % ("[ OK ]" if ok else "[ !! ]", name,
                                          "拦下" if should_catch else "放行",
                                          "拦下" if caught else "放行"))

print("\n=== 二、动态期间解析与对齐单测 ===")
PERIOD_CASES = [
    ("2024-12", ("2024-12", 2024, 12, "FY")),
    ("2026H1", ("2026-06", 2026, 6, "H1")),
    ("2025年三季度", ("2025-09", 2025, 9, "Q3")),
    ("2025Q1", ("2025-03", 2025, 3, "Q1")),
    ("2023年度", ("2023-12", 2023, 12, "FY")),
    (datetime.date(2024, 6, 30), ("2024-06", 2024, 6, "H1")),
]
for raw, exp in PERIOD_CASES:
    got = recalc.parse_period_header(raw)
    ok = (got == exp)
    bad_all += 0 if ok else 1
    print("%-6s %-24r -> %r (期望 %r)" % ("[ OK ]" if ok else "[ !! ]", raw, got, exp))

dyn_periods = ["2023-12", "2024-12", "2025-09"]
for cell_txt, exp_idx in [("2023年", 0), ("2025Q3", 2), ("2025-09", 2), ("2024-12", 1), ("2022-12", None)]:
    got_idx = vr.period_index(cell_txt, dyn_periods)
    ok = (got_idx == exp_idx)
    bad_all += 0 if ok else 1
    print("%-6s period_index(%r, %r) -> %r (期望 %r)" % ("[ OK ]" if ok else "[ !! ]", cell_txt, dyn_periods, got_idx, exp_idx))

total = len(CASES) + len(PERIOD_CASES) + 5
print("\n全部单测结果：%d/%d 通过" % (total - bad_all, total))
sys.exit(1 if bad_all else 0)
