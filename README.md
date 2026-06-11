# Chemical Formula Subscript Formatter — Hermes Agent Skill

自动识别 Word/docx / Markdown / 纯文本中的化学式，将数字下标统一改为真实下标格式。

## 功能

| 模块 | 内容 |
|------|------|
| docx 处理 | Word 真实下标 (w:vertAlign)，保留原文格式 |
| md/txt 处理 | Unicode 下标字符 (₀₁₂₃₄₅₆₇₈₉₊₋) |
| 元素识别 | 118 个 IUPAC 元素符号，双字母优先匹配 |
| 支持类型 | 氧化物、钙钛矿、尖晶石、复合氧化物、载氧体、掺杂材料 |
| 特殊处理 | 括号化学式 Ca(OH)₂、小数掺杂 La₀.₈Sr₀.₂FeO₃、希腊字母缺陷 O₃₋δ |
| 智能排除 | 章节/图表/引用编号、年份、单位、反应系数、离子价态 |
| 输出报告 | 统计 + 示例 + 人工复核清单 |

## 安装

```bash
git clone https://github.com/<your-username>/chemical-formula-subscript-formatter.git \
  "$HOME/AppData/Local/hermes/skills/research/chemical-formula-subscript-formatter"
```

## 使用

```bash
# 预览
python scripts/format_subscripts.py thesis.docx --dry-run

# 处理
python scripts/format_subscripts.py thesis.docx
python scripts/format_subscripts.py notes.md
```

## 依赖

- `python-docx` (仅 docx 模式需要)

## 触发关键词

化学式下标, 下标格式化, Al2O3下标, Fe2O3格式, 钙钛矿下标, chemical formula subscript

## 作者

Aim1ya
