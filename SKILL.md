---
name: chemical-formula-subscript-formatter
description: >-
  Automatically detect chemical formulas in .docx/.md/.txt documents and format
  numerical subscripts as Word true subscript (python-docx OxmlElement) or Unicode
  subscript characters. Handles oxides, perovskites, spinels, doped materials, parenthetical
  groups, decimal doping ratios, and Greek-letter nonstoichiometric suffixes — while
  skipping chapter/table/figure numbers, reference brackets, years, units, reaction
  coefficients, and ion charges. Use when the user asks to format chemical subscripts,
  fix chemical formulas, 化学式下标, 下标格式化, Al2O3→Al₂O₃, Fe2O3 formatting.
version: 1.0.0
author: Aim1ya
metadata:
  hermes:
    tags: [chemistry, subscript, docx, formula, formatting, academic]
    related_skills: [document-editing, seu-phd-thesis]
---

# 化学式下标格式统一处理

自动识别全文中的化学式，将数字下标统一改为真实下标格式（Word `w:vertAlign` 或 Unicode 下标字符）。

## When to Use

- Word/docx 论文、博士论文中化学式下标格式不统一
- Markdown/纯文本中需要规范化化学式下标
- 送审前格式化检查：确保 Al2O3 → Al₂O₃、CaMn₀.₅Ti₀.₅FeO₃ 等
- 关键词：化学式下标, 下标格式化, Al2O3下标, 钙钛矿下标, 载氧体下标, Fe2O3格式

## Don't Use For

- 化学反应方程的配平系数格式化（系数不应为下标）
- 离子价态/电荷上标格式化（Fe³⁺、O²⁻ 是上标，本 skill 不处理）
- 数学公式中的上下标
- LaTeX 数学模式内的化学式

---

## 一、核心规则：什么应该改为下标

### ✅ 应改为下标

| 场景 | 示例 | 下标对象 |
|------|------|----------|
| 元素后整数 | Al2O3 → Al₂O₃ | 2, 3 |
| 小数掺杂比 | La0.8Sr0.2FeO3 → La₀.₈Sr₀.₂FeO₃ | 0.8, 0.2 |
| 括号后数字 | Ca(OH)2 → Ca(OH)₂ | 2 |
| 括号内元素+数字 | Fe(NO3)3 → Fe(NO₃)₃ | 括号内3, 括号外3 |
| 希腊字母缺陷项 | O3-δ → O₃₋δ | 3, -δ |
| 非整比化学计量 | CeO2-x → CeO₂₋ₓ | 2, -x |
| 复合氧化物 | CaMn0.5Ti0.5FeO3 | 0.5, 0.5, 3 |
| 钙钛矿 | La0.8Sr0.2FeO3 | 0.8, 0.2, 3 |
| 尖晶石 | CuFe2O4 → CuFe₂O₄ | 2, 4 |
| 载氧体材料 | Fe2O3/Al2O3 → Fe₂O₃/Al₂O₃ | 2, 3, 2, 3 |

### ❌ 不得修改

| 类别 | 示例 | 原因 |
|------|------|------|
| 章节编号 | 1.1、2.3.4、第3章 | 文档结构 |
| 图表编号 | 图2-1、表3.2、Figure 4 | 文档结构 |
| 参考文献编号 | [1]、[2-5]、(Smith et al., 2020) | 引用 |
| 年份 | 2024、2025 | 时间 |
| 温度/时间/质量 | 900 ℃、2 h、10 wt%、50 mL、4 MW | 单位 |
| 反应式系数 | **2**Fe + O₂ → **2**FeO（前置系数不改） | 化学计量系数 |
| 离子价态 | Fe3+、Cu2+、O2−（此处为**上标**） | 价态/电荷 |
| 英文缩写 | CEJ、XRD、SEM、4MW | 非化学式 |
| DOI / URL | 10.1016/j.cej.2023.139224 | 标识符 |
| 文件名 | fig_2_3.png | 文件路径 |

---

## 二、化学式识别逻辑

### 元素符号列表

使用完整的 118 个 IUPAC 元素符号（He, Li, Be, ..., Og），按长到短排序以优先匹配双字母元素（如 He 优先于 H，Fe 优先于 F）。

### 化学式判定条件

一个文本片段被判定为化学式，需同时满足：

1. 包含至少 **2 个** 元素符号（大写字母开头的 1–2 字母组合）
2. 元素符号与其后数字之间**无空格**
3. 至少有一个元素符号后紧跟数字（整数或小数）
4. 不被任何排除规则命中

### 排除规则（按优先级）

1. **前置上下文排除**：前面紧邻 "图"、"表"、"Table"、"Figure"、"第"、"节"、"Section"、"参考文献"、"Reference"、"式("、"公式(" → 不是化学式
2. **后置上下文排除**：后面紧跟 "℃"、"h"、"min"、"s"、"mL"、"g"、"kg"、"wt%"、"MW"、"eV"、"ppm" → 不是化学式
3. **引用格式排除**：被 `[...]`、`(...)` 包围且含 `et al.` 或年份 → 不是化学式
4. **纯数字排除**：全部由数字和符号构成，无任何元素符号 → 不是化学式
5. **离子电荷排除**：以 `\d+[+-]$` 结尾且前面紧邻元素符号 → 标记为价态，不处理
6. **反应式系数排除**：化学式前的独立数字（如 `2Fe` 中的 2）→ 系数不改

---

## 三、Word/docx 真实下标实现

优先使用 Word 原生 `w:vertAlign` 属性，**不使用** Unicode 下标字符。

### docx 处理要点

1. 必须先拼接段落完整文本（`para.text`），在完整文本上检测化学式位置
2. 通过字符偏移映射回具体的 `<w:r>` 和 `<w:t>` 元素
3. 对下标字符：将其从原 run 中拆分出来，包入新的 `<w:r>` 并设置 `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`
4. 保留原文的字体、字号、粗体、斜体等属性

### OxmlElement 下标设置

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def make_subscript_run(text):
    """Create a new w:r element with subscript formatting."""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    vertAlign = OxmlElement('w:vertAlign')
    vertAlign.set(qn('w:val'), 'subscript')
    rPr.append(vertAlign)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    return r
```

### 跨 run 处理

Word 常将化学式分散在多个 `<w:r>` 中（如 `"Al"` 和 `"2O3"` 分属两个 run）。脚本需：
1. 先提取段落级文本，在段落级检测化学式
2. 定位每个下标字符在哪个 run 的哪个字符位置
3. 拆分 run 并重组为标准/下标交替结构

---

## 四、Markdown / 纯文本 Unicode 下标

对于 .md / .txt 文件，使用 Unicode 下标字符替换：

```python
SUBSCRIPT_UNICODE = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
    '+': '₊', '-': '₋',
}
# 注意：Unicode 无下标小数点 '·'，小数在下标上下文中通常保留 '.' 原样
```

---

## 五、操作工作流

### Step 1 — 备份

```python
import shutil
shutil.copy2(src, src + '.bak')
```

### Step 2 — 执行格式化

```bash
python scripts/format_subscripts.py <文件路径.docx/.md/.txt>
```

可选参数：
- `--dry-run` — 仅检测不修改，输出预览报告
- `--output <路径>` — 指定输出路径（默认：原文件名_subscript_fixed.ext）
- `--format {docx,unicode}` — 强制指定格式模式

### Step 3 — 验证

打开输出文件，抽查以下几类化学式：
- 简单氧化物：Al2O3、Fe2O3、CO2
- 括号复合物：Ca(OH)2、Fe(NO3)3
- 小数掺杂：La0.8Sr0.2FeO3
- 希腊字母缺陷：O3-δ、CeO2-x

### Step 4 — 输出报告

脚本自动输出修改报告（见第六节）。

---

## 六、修改报告格式

```
══════════════════════════════════════
化学式下标格式化报告
══════════════════════════════════════

源文件: thesis_ch3.docx
输出文件: thesis_ch3_subscript_fixed.docx
格式模式: Word 真实下标 (w:vertAlign)

──────────────────────────────────────
修改统计
──────────────────────────────────────
共识别化学式: 247 处
已修改下标: 247 处
跳过（非化学式数字）: 892 处

──────────────────────────────────────
修改示例（前 20 条）
──────────────────────────────────────
Al2O3     → Al₂O₃
Fe2O3     → Fe₂O₃
CO2       → CO₂
Ca(OH)2   → Ca(OH)₂
La0.8Sr0.2FeO3 → La₀.₈Sr₀.₂FeO₃
O3-δ      → O₃₋δ
...

──────────────────────────────────────
需人工复核（价态/电荷/不确定项）
──────────────────────────────────────
⚠ Fe3+     — 可能为离子价态（应上标），未修改
⚠ Cu2+     — 可能为离子价态（应上标），未修改
⚠ O2−      — 可能为氧离子，未修改
? Ce0.8Zr0.2O2/Al2O3 — 确认 / 是否分隔两个独立化学式
...

──────────────────────────────────────
未修改的数字类型
──────────────────────────────────────
章节编号: 42 处（如 "1.1"、"2.3.4"）
图表编号: 18 处（如 "图 2-1"、"表 3.2"）
参考文献: 197 处（如 "[1]"、"[2-5]"）
年份/单位: 156 处（如 "2024"、"900 ℃"）
反应式系数: 12 处（如 "2Fe + O2 → 2FeO" 中的系数 2）
离子价态: 8 处（需人工确认上标）
══════════════════════════════════════
```

---

## Common Pitfalls

1. **跨 run 拆分** → Word 将 `Al2O3` 存为 `"Al"` 和 `"2O3"` 两个 run，逐 run 正则不匹配。必须先 `para.text` 再映射回 run 位置。
2. **反应式系数误改** → `2Fe` 的 2 是系数不是下标。脚本通过"元素符号前的数字不改"规则规避。
3. **离子价态误改** → `Fe3+` 的 3+ 是上标。脚本检测 `\d+[+-]` 结尾模式并主动跳过，放入人工复核清单。
4. **化学式/缩写混淆** → `CO2` 是化学式，但 `CO2 emissions` 的上下文是环境叙述。脚本依赖元素序列特征判断，短序列（单个元素+数字如 `N2`）可能误判 — 报告会列出所有短序列供复核。
5. **Unicode vs Word 下标混用** → 对 docx 必须用 Word 真实下标，否则在 Word 中无法被 TOC/搜索/格式刷正确识别。
6. **小数下标的点号处理** → `0.5` 作为下标时，Unicode 没有下标点号。Word 模式中直接将 `.` 设为下标属性即可；Unicode 模式保留原样 `.` 并在报告中注明。
7. **含斜杠的复合化学式** → `Fe2O3/Al2O3` 中的 `/` 不是化学式部分。脚本将 `/` 视为分隔符，两侧分别处理。
8. **希腊字母编码** → `δ` 可能是 Unicode U+03B4，也可能是 `&delta;` HTML 实体。脚本统一按 Unicode 处理。

---

## Verification Checklist

- [ ] 原文件已备份（`.bak`）
- [ ] 输出文件另存，未覆盖原文件
- [ ] Al2O3、Fe2O3、CO2、H2O 等基础氧化物下标正确
- [ ] Ca(OH)2、Fe(NO3)3 等括号化学式下标正确
- [ ] La0.8Sr0.2FeO3 小数掺杂下标正确
- [ ] O3-δ、CeO2-x 希腊字母缺陷项下标正确
- [ ] 章节/图表/参考文献编号未被修改
- [ ] 年份、温度、时间、单位数字未被修改
- [ ] 反应式系数未被修改（如 2Fe 的 2）
- [ ] 离子价态（Fe3+、Cu2+）未被误改为下标
- [ ] Word 文本格式（字体/字号/粗体/斜体）保留
- [ ] 修改报告已生成并审阅
- [ ] 人工复核清单中的项目已逐项确认
