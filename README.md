# review-eia-report · 建设项目环境影响报告表审核 Skill

审核建设项目环境影响报告表（环评报告表），输出结构化审核意见清单。
兼容 Qoder / Claude Code / Codex / OpenCode。

## 文件结构

```text
SKILL.md            # 主流程（入口，先读这个）
checklist.md        # 10维度加权审核清单（第零节指南符合性+★必查/⛔否决/验算公式/行业因子矩阵/Top21）
regulations.md      # 法规标准参考（截至2026-09，附现行状态与实施日期+第十二节联网核验操作）
scripts/
  extract_docx.py   # docx/pdf 文本提取脚本
  verify_standards.py # 标准时效初筛脚本（本地库，需官网复核）
  standards_db.json # 标准库（34项，更新日期见_meta.updated）
  requirements.txt  # Python 依赖
examples/
  audit_example.md  # 输出示例
CHANGELOG.md        # 变更记录
```

## 快速开始

```bash
pip install -r scripts/requirements.txt

# 提取报告文本
python scripts/extract_docx.py "建设项目环境影响报告表.docx"

# JSON 输出（对接系统）
python scripts/extract_docx.py "报告.docx" --format json --out extracted.json

# 标准时效初筛（疑似过期/未知须官网复核）
python scripts/verify_standards.py extracted.md --out std_check.md
```

然后把提取文本 + `SKILL.md` 工作流一起交给模型，按第 0→5 步执行审核。
详细流程见 [SKILL.md](SKILL.md)。

## 审核顺序建议

法规符合性(8) → 基本信息(1) → 工程分析(2) → 一致性验算(9) → 预测(4) → 措施(5) → 总量(6) → 现状(3) → 编制质量(10) → 公参(7)

## 版本

v2.2.0（2026-09）：按污染影响类编制技术指南审核（第零节+专项判定）、标准时效初筛脚本+联网核验流程、第4节专项/非专项分流。详见 CHANGELOG。
