# Chemical Formula Subscript Formatter Skill

用于审查并受控修订 DOCX、Markdown 和纯文本中的化学式、材料式及非整比表达下标。

本 Skill 采用平台无关的 `SKILL.md` 指令，可供任何能够读取 Skill 文件、访问用户文档并输出报告或修改副本的 AI Agent 使用。

## 功能

| 模块 | 内容 |
|---|---|
| 保守审计 | 默认只生成候选清单，不直接修改文件 |
| DOCX 处理 | 对安全的普通 run 使用 Word 真实下标；跳过跨 run 和复杂结构 |
| Markdown/TXT | 经明确允许后使用 Unicode 下标字符 |
| 支持表达 | `Al2O3`、`Fe3O4`、`O2`、`NOx`、`CaMnTiFeOx`、`CeO2-x` 等 |
| 智能排除 | 图表/章节/公式/引用编号、年份、单位、反应系数和离子价态 |
| 风险控制 | 不覆盖源文件；不改变正文语义；歧义项进入人工复核 |

## 安装

将整个仓库放入目标 Agent 支持的 Skill 目录，或让 Agent 直接读取本仓库中的 `SKILL.md`。

```bash
git clone https://github.com/miaozw97-ship-it/chemical-formula-subscript-formatter.git
```

不同 Agent 的 Skill 安装目录和加载方式不同，请遵循对应平台的说明。

## 脚本使用

默认只审计：

```bash
python scripts/format_subscripts.py thesis.docx
python scripts/format_subscripts.py notes.md
```

审阅候选后，显式应用高置信度修改：

```bash
python scripts/format_subscripts.py thesis.docx --apply --output thesis_subscript_fixed.docx
```

DOCX 处理需要 `python-docx`。

## 测试

```bash
python -m unittest discover -s tests -v
```

测试场景位于 [`references/test-cases.md`](references/test-cases.md)。

## 作者

Aim1ya
