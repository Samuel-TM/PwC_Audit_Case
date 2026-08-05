# A 公司模拟财务报表审计案例

本项目依据 `Plan.md` 与案例 PDF，完成 A 公司 2020—2022 年财务分析、风险评估、审计资料清单、四周审计计划、导师问答及 15 页汇报材料。

## 一键构建

```bash
python scripts/build_all.py
PYTHONPATH=vendor/python python -m pytest -q
```

最终文件位于 `output/`。金额单位统一为万元；原表中的 `-` 录入为缺失值。风险判断是模拟审计的计划阶段判断，不代表已经发现错报、舞弊或形成正式审计意见。

## 主要输出

- `A公司审计案例汇报.pptx` / `.pdf`
- `审计资料清单.xlsx`
- `审计计划.xlsx`
- `导师问答手册.md`
- `analysis_summary.md`
- `charts/` 九张 1600×900 图表
- `build_report.txt`

## 数据口径

- 数据来源：`source/1-2 Assignment for PwC You Plus.pdf`
- 报告期间：2020—2022 年
- 金额单位：万元（原材料单价和产品单位价格按源表注明单位）
- 合计与比例校验容差：0.02 万元或 0.02 个百分点
