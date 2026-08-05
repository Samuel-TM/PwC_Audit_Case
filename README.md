# A 公司模拟财务报表审计案例

本项目依据案例 PDF 与 Version 4 优化要求，完成 A 公司 2020—2022 年财务分析、风险评估、审计资料清单、四周审计计划、导师问答及 15 页 LaTeX 汇报 PDF。最终汇报只使用 Beamer + XeLaTeX，不生成或维护 PPTX。

## 一键构建

```bash
source /Users/sunjiashan/MiniConda/miniconda3/etc/profile.d/conda.sh
conda activate hyt
python scripts/build_all.py
python -m pytest -q
```

V4 最终文件位于 `output/output_V4/`。唯一 LaTeX 源文件为 `report/A公司审计案例汇报.tex`，由 Git 管理版本，不再复制到 output。金额单位统一为万元；原表中的 `-` 录入为缺失值。风险判断是模拟审计的计划阶段判断，不代表已经发现错报、舞弊或形成正式审计意见。

## 主要输出

- `A公司审计案例汇报_V4.pdf`
- `审计资料清单_V4.xlsx`
- `审计计划_V4.xlsx`
- `导师问答手册_V4.md`
- `analysis_summary_V4.md`
- `charts_V4/` 九张 1600×900 图表
- `build_report_V4.txt`
- `validation_report_V4.json`
- `audit_risk_results_V4.json`
- `workbook_verification_V4.json`
- `pdf_verification_V4.json`

## 数据口径

- 数据来源：`source/1-2 Assignment for PwC You Plus.pdf`
- 汇报版式：16:9 Beamer，统一字号宏与橙红灰视觉体系
- 报告期间：2020—2022 年
- 金额单位：万元（原材料单价和产品单位价格按源表注明单位）
- 合计与比例校验容差：0.02 万元或 0.02 个百分点
