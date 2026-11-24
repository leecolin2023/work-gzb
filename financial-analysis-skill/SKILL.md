---
name: financial-analysis-skill
version: 2.1.0
display_name: "深度穿透式财报分析流水线"
description: "面向非金融企业信贷审批与尽调的财务分析 Skill，覆盖一手原表导出、确定性指标复算、穿透式风险分析、全覆盖回归校验、独立盲审裁决与公文标准 DOCX 导出。"
---

# 深度穿透式财报分析流水线

## 1. 目录结构与文件职责

全包按职责划分，每类知识仅维护一份正本：

| 文件                         | 职责                                                                             |
| -------------------------- | ------------------------------------------------------------------------------ |
| `SKILL.md`                 | 执行入口：六阶段流水线、模式路由表、脚本命令与交付闸门                                                    |
| `CHANGELOG.md`             | 版本变更记录与历史实测缺陷明细档案                                                              |
| `references/提示词包.md`       | 指令层：S0 规格重写、P1 完整版、P2 简版、P3-Blind 独立盲审、P4/P5/P6（支持脱离 Skill 单独复制使用）             |
| `references/规则与方法学正本.md`   | 业务层：双域 `NR` 门禁、DQ001–DQ012、指标口径、F1–F10/D1–D6 规则、五大 Playbook、压力测试、贷前贷中贷后映射与数据契约 |
| `references/计算约束与质检手册.md`  | 防错层：12 条全局硬约束、FM-01~FM-13 失效模式目录、交付前检查清单与口径登记表模板                               |
| `references/回归与验收基准.md`    | 回归与验收层：`CASE-001` 回归基线与 A/B 盲评验收协议                                             |
| `scripts/dump_xlsx.py`     | 一手原表逐单元格导出（规避转述误差）                                                             |
| `scripts/recalc.py`        | 确定性复算引擎（自动识别表头期间、列顺序与 `FY/H1/Q1/Q3` 期间性质，输出 C1–C6 与 DQ005 门禁）                  |
| `scripts/verify_report.py` | 全覆盖回归校验器（期间动态对齐、全表计入分母、动态展示精度容差 `0.5×10^-d`、`--selftest` 不变式、自指统计核对）           |
| `scripts/test_selfdecl.py` | 校验器与期间解析回归单测                                                                   |
| `scripts/export_docx.py`   | 公文标准 DOCX 渲染器（A4 版心、奇偶页码、公文字体层级、表格格式契约与超宽表自动横向）                                |

---

## 2. 六阶段执行流水线（严格按序，不得跳段）

设 `$SK` 为本 Skill 目录，`$XLSX` 为源报表，`$WD` 为工作目录。

### Stage 0 · 深度规格重写（高风险/上会案卷必跑）

调用 `references/提示词包.md` 中的 **`S0`**，基于企业元数据与原表概览完成四项前置定义：

1. 暴露 2–4 个隐含前提，锁定 **Top 1–3 最可能改变评级的翻盘检查项**；
2. 为每项强制构造**正常业务替代解释**与**翻盘阈值**；
3. 按具体主张指定**一手证据精确入口**及“能证什么 / 不能证什么”边界；
4. 锁定 L1 / L2 / L3 独立验收闸门，产出 `depth_spec.md`。

### Stage 1 · 一手原表导出与数据质量门禁

```bash
python "$SK/scripts/dump_xlsx.py" "$XLSX" "$WD/raw_dump.md"
```

逐表核对实际期数、期间性质、货币单位、合并/单体口径、审计意见与「留空 vs 明确为 0」（防 FM-10）。按 `references/规则与方法学正本.md` 执行 DQ001–DQ012 与双域可分析性门禁（不满足最小证据条件则直接锁定 `NR`）。

### Stage 2 · 确定性指标复算（拿走模型算术权）

```bash
python "$SK/scripts/recalc.py" "$XLSX" "$WD/analysis" --company "公司名"
```

产出 `metrics.json`（报告唯一数值基准，含 `_periods` / `_period_kinds` / `_period_days`）、`metrics.md`、`checks.md`（C1–C6 校验与 DQ005 阻断提示）。

- **硬纪律**：`checks.md` 中 C1 有 ❌ 必须先解决原表闭合问题；无法计算或触发 DQ005 的严格口径指标一律标 `NOT_COMPUTABLE`，严禁心算或填写近似值。

### Stage 3 · 穿透分析与证伪撰写

按模式路由选取 `references/提示词包.md` 中对应提示词（如 `P1` / `P2`），并严格执行 `references/规则与方法学正本.md` 与 `references/计算约束与质检手册.md`：

- 优先攻克 Stage 0 锁定的 Top 1–3 翻盘检查项；
- 构建「主张-一手证据-替代解释-翻盘条件」核心台账，标注 `[事实] / [判断] / [待核]` 并执行独立证据链去重；
- 正文段落中出现的每一个派生数字，必须落入 `metrics.json` 基准键；所有展示值由全精度原值一次舍入。

### Stage 4 · 三层独立裁决（先审校验器，再盲审逻辑）

1. **L1 · 算术与表内全覆盖校验**：
   ```bash
   python "$SK/scripts/test_selfdecl.py"
   python "$SK/scripts/verify_report.py" --selftest "$REPORT" "$WD/analysis/metrics.json"        --src "$XLSX" --out "$WD/analysis/verify.md"
   ```
   确认 `INV-1` / `INV-2` 通过，并按三档判词如实披露：
   - `FULL-PASS`（`fail=0, self_inconsistent=0, self_declared_bad=0, uncovered=0`）：允许写“全部一致”；
   - `NO-MISMATCH`（前三项为 0，但 `uncovered>0`）：**只能写“查到的都对”并逐条列出未覆盖清单**；
   - `FAIL`（存在不一致、跨章矛盾或自指覆盖率不符）：**不得交付**。
2. **L2 · 规则与口径门禁校验**：核查 DQ001–DQ012、C1–C6、`NR` 门禁、根因去重与方法学口径是否严格落实。
3. **L3 · 独立盲审对抗裁决**：**必须新开无历史对话上下文的独立窗口**，运行 `references/提示词包.md` 中的 **`P3-Blind`**，仅提供原始底稿 + 规则正本 + 待审报告草稿，输出 `APPROVE / REVISE / REJECT`。存在任一硬伤（FM-10/11/12/13、越界断言、强行评级替代 `NR`、覆盖率造假）一律 `REJECT` 退回重写。

### Stage 5 · 公文标准 DOCX 导出（通过确定性脚本渲染，不靠模型手写排版）

L1 与 L3 裁决通过后，由 `export_docx.py` 直接将已定稿的 Markdown 报告渲染为公文标准 Word 文档：

```bash
python "$SK/scripts/export_docx.py" --selftest
python "$SK/scripts/export_docx.py" "$REPORT" "$WD/财务分析报告_<主体>_<期间>.docx" \
    --recipient "授信审批委员会" --issuer "风险管理部" --date "YYYY-MM-DD"
```

- **内置排版契约**：A4（`210mm × 297mm`，边距上 37 / 下 35 / 左 28 / 右 26mm）；奇偶页 `— PAGE —` 四号宋体页码；标题 2 号方正小标宋简体、一级标题 3 号黑体、二级标题 3 号楷体、正文 3 号仿宋体（28pt 固定行距、首行缩进 2 字符）；正文表仿宋 / 附件表宋体、表首黑体跨页重复、四号↔小四自适应密度、数字右对齐 / 标签居中 / 长文本左对齐缩进 2 字符、外框 1pt 深色内框 0.5pt 浅色、≥8 列超宽表自动切换 A4 横向分节、默认省略文末机器可读 JSON 块。
- **官方模板边界**：如需红头、发文字号、签发人、印章、版记或正式信函/纪要版式，必须传入 `--template <官方模板.docx>` 保留原有模板版式，禁止凭空伪造发文要素。

---

## 3. 模式路由表

| 模式                     | 适用场景          | 推荐提示词链                      | 核心门禁                                |
| ---------------------- | ------------- | --------------------------- | ----------------------------------- |
| `full_analysis`        | 正式贷前尽调 / 上会审批 | `S0 → P1 → P3-Blind`        | DQ + NR + C1–C6 + L1/L2/L3          |
| `quick_scan`           | 5 分钟一页纸初筛体检   | `P2`                        | DQ + NR + 核心指标脚本复算                  |
| `data_sufficiency`     | 材料不全时的可分析性诊断  | `P6`（禁做评级与结论）               | 仅盘点完整度与缺口                           |
| `reliability_review`   | 报表可信度 / 粉饰专项  | `S0 → P1(A域) → P3-Blind`    | DQ + A域NR + F1–F10 + Fraud Playbook |
| `statement_drilldown`  | 三表核心科目逐项穿透    | `P1(核心科目六问)`                | Core Accounts Playbook + FM-12防越界   |
| `ratio_review`         | 多期比率与显著变动专项   | `recalc.py` + 变动>20%/5pp解释  | 期间对齐 + 动态容差校验                       |
| `profit_growth_review` | 盈利质量与成长性专项    | 利润桥 + 三支柱 + 杜邦 + CAGR/SGR   | DQ005 + 可比期间 CAGR                   |
| `cashflow_review`      | 现金流与回款真实性专项   | 现金流质量 + 流水交叉核验              | 禁伪倒推 + 禁 `+++` 污名化                  |
| `solvency_review`      | 偿债、隐性负债与压力测试  | 偿债 + 隐性负债清单 + 三情景压测         | C2/C3 + 情景非预测声明                     |
| `fraud_scan`           | 七类舞弊线索与红旗排查   | F1–F10 + RQ001–RQ005 + 替代解释 | 根因去重 + 持续性加权                        |
| `credit_limit`         | 流动资金贷款需求参考测算  | `P5`                        | 适用性阻断（>720天/项目制禁用）+ 全精度链            |
| `final_synthesis`      | 结论汇总与贷前贷中贷后动作 | 双域评级 + 风险动作一一映射             | 证据映射 + 人工复核声明                       |



> 非上市/中小微企业在上述任一模式末尾附加 **`P4`** 合规适配补丁。

---

## 4. 交付物清单（缺一不可）

1. `财务分析报告_<主体>_<期间>.md` 与 `财务分析报告_<主体>_<期间>.docx`（正式公文排版成稿，含深度规格卡、主张-证据-证伪台账、三层验证边界声明）
2. `analysis_raw_dump.md`（原表逐单元格导出留痕）
3. `analysis_metrics.json` + `analysis_checks.md`（确定性复算基准与 C1–C6 / DQ 门禁轨迹）
4. `analysis_verify.md`（L1 全覆盖校验记录与判词）
5. `analysis_blind_review.md`（L3 独立盲审裁决记录）
