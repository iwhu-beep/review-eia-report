# review-eia-report · 建设项目环境影响报告表审核 Skill

> ⚠️ **风险声明**：本工具为 AI 辅助审核，不能替代注册环境影响评价工程师的专业判断和法定审批程序；
> 所有审核意见（含缺陷分级与总评结论）必须经人工复核确认后方可作为整改或报批依据；
> 法规标准状态以生态环境部、市场监管总局官网现行版本为准。

审核建设项目环境影响报告表（环评报告表），输出结构化审核意见清单。
兼容 Qoder / Claude Code / Codex / OpenCode。

## 文件结构

```text
SKILL.md            # 主流程（入口，先读这个）
checklist.md        # 10维度加权审核清单（第零节指南符合性+设备三对照+★必查/⛔否决/验算公式/行业因子与设备矩阵/速查）
regulations.md      # 法规标准参考（截至2026-09，附现行状态与实施日期+第十二节联网核验操作）
scripts/
  extract_docx.py   # docx/pdf 文本提取脚本
  verify_standards.py # 标准时效初筛脚本（本地库，需官网复核）
  export_pdf.py     # 审核意见md转PDF（中文，需fpdf2）
  standards_db.json # 标准库（42项，更新日期见_meta.updated）
  requirements.txt  # Python 依赖
examples/
  audit_example.md  # 输出示例
CHANGELOG.md        # 变更记录
```

## 快速开始

```bash
pip install -r scripts/requirements.txt

# PDF 先做扫描预检（退出码4=疑似扫描件，先OCR）
python scripts/extract_docx.py 报告.pdf --scan-check

# 提取报告文本
python scripts/extract_docx.py "建设项目环境影响报告表.docx"

# 长文档按章节切分（分块审核，表格章单独完整载入）
python scripts/extract_docx.py "报告.docx" --split-chapters --split-out chaps/

# JSON 输出（对接系统）
python scripts/extract_docx.py "报告.docx" --format json --out extracted.json

# 标准时效初筛（疑似过期/未知须官网复核）
python scripts/verify_standards.py extracted.md --out std_check.md

# 审核意见导出 PDF（存档/送审）
python scripts/export_pdf.py 审核意见.md --out 审核意见.pdf
```

然后把提取文本 + `SKILL.md` 工作流一起交给模型，按第 0→5 步执行审核。
详细流程见 [SKILL.md](SKILL.md)。

## 审核顺序建议

法规符合性(8) → 基本信息(1) → 工程分析(2) → 一致性验算(9) → 预测(4) → 措施(5) → 总量(6) → 现状(3) → 编制质量(10) → 公参(7)

## 版本

v2.6.0（2026-09）：生态环境法典施行，10部单行法废止，编制依据与过渡规则更新。
v2.5.2（2026-09）：全量复核标准政策（新增施工噪声/污水厂修改单/排污许可新规）。
v2.5.1（2026-09）：危废名录2025与GB3095-2026纠错。
v2.5.0（2026-09）：PDF归档升级+四级缺陷制与总评结论+扫描预检与生态分支+三处强声明+脚本异常加固与章节切分。
v2.4.0（2026-09）：设备防遗漏核查（三对照+行业设备矩阵）。详见 CHANGELOG。
