#!/usr/bin/env python3
"""
Chemical Formula Subscript Formatter

Detects chemical formulas in .docx/.md/.txt files and formats numerical
subscripts as Word true subscript (python-docx OxmlElement) or Unicode
subscript characters.

Usage:
    python format_subscripts.py thesis.docx
    python format_subscripts.py thesis.docx --dry-run
    python format_subscripts.py thesis.docx --output fixed.docx
    python format_subscripts.py notes.md --format unicode
"""

import re
import sys
import shutil
import argparse
from pathlib import Path
from collections import defaultdict
from copy import deepcopy

# ── Element Symbols (IUPAC, sorted by length descending) ──────────────
ELEMENTS = sorted([
    'H',  'He', 'Li', 'Be', 'B',  'C',  'N',  'O',  'F',  'Ne',
    'Na', 'Mg', 'Al', 'Si', 'P',  'S',  'Cl', 'Ar', 'K',  'Ca',
    'Sc', 'Ti', 'V',  'Cr', 'Mn', 'Fe', 'Co', 'Ni', 'Cu', 'Zn',
    'Ga', 'Ge', 'As', 'Se', 'Br', 'Kr', 'Rb', 'Sr', 'Y',  'Zr',
    'Nb', 'Mo', 'Tc', 'Ru', 'Rh', 'Pd', 'Ag', 'Cd', 'In', 'Sn',
    'Sb', 'Te', 'I',  'Xe', 'Cs', 'Ba', 'La', 'Ce', 'Pr', 'Nd',
    'Pm', 'Sm', 'Eu', 'Gd', 'Tb', 'Dy', 'Ho', 'Er', 'Tm', 'Yb',
    'Lu', 'Hf', 'Ta', 'W',  'Re', 'Os', 'Ir', 'Pt', 'Au', 'Hg',
    'Tl', 'Pb', 'Bi', 'Po', 'At', 'Rn', 'Fr', 'Ra', 'Ac', 'Th',
    'Pa', 'U',  'Np', 'Pu', 'Am', 'Cm', 'Bk', 'Cf', 'Es', 'Fm',
    'Md', 'No', 'Lr', 'Rf', 'Db', 'Sg', 'Bh', 'Hs', 'Mt', 'Ds',
    'Rg', 'Cn', 'Nh', 'Fl', 'Mc', 'Lv', 'Ts', 'Og',
], key=len, reverse=True)

ELEM_ALT = '|'.join(ELEMENTS)  # He|Hg|Hf|...|H

# ── Unicode Subscript Map ─────────────────────────────────────────────
SUBSCRIPT_UNICODE = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
    '+': '₊', '-': '₋',
}
# Note: no Unicode subscript for '.' — keep dot as-is in unicode mode

# ── Exclusion Patterns ────────────────────────────────────────────────

# Pre-context: if the match is preceded by these, it's not a chemical formula
PRE_EXCLUDE = re.compile(
    r'(图|表|Table|Figure|Fig\.?|第|节|Section|式\s*\(|公式\s*\(|'
    r'参考文献|Reference|方程\s*\()\s*$',
    re.IGNORECASE
)

# Post-context: if the match is followed by these units, it's not a chemical formula
POST_EXCLUDE = re.compile(
    r'^\s*(℃|°C|°F|K|h|min|s|mL|L|g|kg|mg|μg|wt%|%|MW|kW|eV|ppm|ppb|'
    r'mol|M|N|Pa|bar|atm|W|J|cal|V|A|Ω|Hz|m|cm|mm|μm|nm)',
    re.IGNORECASE
)

# Reference bracket patterns — skip content inside these
REF_PATTERN = re.compile(r'\[\d+(?:[-,]\d+)*\]')

# Ion charge pattern: element symbol followed by digits and +/-
ION_CHARGE = re.compile(rf'({ELEM_ALT})(\d+[+\-])')

# Greek / nonstoichiometric suffix characters
GREEK_SUFFIX_CHARS = 'δxαβγΔ'

# ── Chemical Formula Detection ────────────────────────────────────────

def build_formula_pattern():
    """Build the regex pattern for chemical formula detection."""
    # A single formula unit: element symbol + optional digits/decimal
    unit = rf'(?:{ELEM_ALT})(?:\d+(?:\.\d+)?)?'

    # A parenthesized group: (element+num element+num ...) followed by optional num
    paren = rf'\((?:{unit})+\)\d*'

    # Greek/nonstoichiometric suffix
    greek = rf'[-–][{GREEK_SUFFIX_CHARS}]\d*'

    # Full formula: at least 2 units, possibly with paren groups, optional greek suffix
    # Must be bounded by non-alpha-numeric chars (or start/end of string)
    formula = (
        r'(?<![a-zA-Z\d.])'
        rf'((?:{unit}|{paren}){{2,}}'
        rf'(?:{greek})?)'
        r'(?![a-zA-Z\d.])'
    )
    return re.compile(formula)

FORMULA_RE = build_formula_pattern()

# ── Exclusion Checks ──────────────────────────────────────────────────

def is_excluded(match, full_text, match_start):
    """Check if a match should be excluded based on context."""
    matched = match.group(0)

    # Skip pure numbers
    if re.match(r'^[\d.\-+]+$', matched):
        return True, "纯数字"

    # Skip if only one element symbol (likely not a formula)
    elem_count = len(re.findall(rf'\b({ELEM_ALT})\b', matched))
    if elem_count < 2:
        return True, "元素符号不足2个"

    # Check pre-context (50 chars before)
    pre_start = max(0, match_start - 50)
    pre_text = full_text[pre_start:match_start]
    if PRE_EXCLUDE.search(pre_text):
        return True, "命中前置排除规则"

    # Check post-context (20 chars after)
    post_end = min(len(full_text), match_end(match))
    post_text = full_text[post_end:post_end + 20]
    if POST_EXCLUDE.match(post_text):
        return True, "命中后置排除规则"

    # Check if inside reference brackets
    line_start = max(0, full_text.rfind('\n', 0, match_start))
    line = full_text[line_start:full_text.find('\n', match_end(match))] if '\n' in full_text[match_end(match):] else full_text[line_start:]
    for ref_match in REF_PATTERN.finditer(line):
        if ref_match.start() <= match_start - line_start < ref_match.end():
            return True, "位于引用编号内"

    # Check for ion charge pattern
    if ION_CHARGE.fullmatch(matched.strip()):
        return True, "离子价态（应上标）"

    return False, ""

def match_end(match):
    """Get match end position, handling both re.Match and span tuples."""
    if hasattr(match, 'end'):
        return match.end()
    return match[1]

# ── Subscript Position Detection ──────────────────────────────────────

def find_subscript_ranges(formula_text):
    """
    Given a chemical formula string, return list of (start, end) character
    positions that should be subscripted.
    """
    ranges = []

    # Pattern 1: Element symbol followed by digits/decimal → digits are subscript
    # e.g., Al2 → '2' is subscript; La0.8 → '0.8' is subscript
    for m in re.finditer(rf'(?:{ELEM_ALT})(\d+(?:\.\d+)?)', formula_text):
        ranges.append((m.start(1), m.end(1)))

    # Pattern 2: ')' followed by digits → digits are subscript
    # e.g., Ca(OH)2 → '2' is subscript
    for m in re.finditer(r'\)(\d+)', formula_text):
        ranges.append((m.start(1), m.end(1)))

    # Pattern 3: '-' or '–' followed by Greek letter + optional digit → subscript
    # e.g., O3-δ → '-δ' is subscript; CeO2-x → '-x' is subscript
    for m in re.finditer(rf'([-–][{GREEK_SUFFIX_CHARS}]\d*)', formula_text):
        ranges.append((m.start(1), m.end(1)))

    # Merge overlapping ranges
    if not ranges:
        return []
    ranges.sort()
    merged = [ranges[0]]
    for r in ranges[1:]:
        if r[0] <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], r[1]))
        else:
            merged.append(r)
    return merged


def apply_unicode_subscripts(text, ranges):
    """Apply Unicode subscript characters to text at given ranges."""
    result = list(text)
    for start, end in reversed(ranges):  # reverse to preserve indices
        segment = text[start:end]
        subbed = ''.join(SUBSCRIPT_UNICODE.get(c, c) for c in segment)
        result[start:end] = subbed
    return ''.join(result)


# ── DOCX Processing ───────────────────────────────────────────────────

def process_docx(filepath, dry_run=False, output_path=None):
    """Process a .docx file: detect and format chemical formula subscripts."""
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    doc = Document(filepath)
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    stats = {
        'formulas_found': 0,
        'formulas_modified': 0,
        'digits_skipped': 0,
        'excluded': defaultdict(int),
        'modifications': [],
        'to_review': [],
    }

    def make_subscript_element(text):
        """Create a w:r element with subscript formatting."""
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        va = OxmlElement('w:vertAlign')
        va.set(qn('w:val'), 'subscript')
        rPr.append(va)
        r.append(rPr)
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = text
        r.append(t)
        return r

    def make_normal_element(text):
        """Create a w:r element with no special formatting."""
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = text
        r.append(t)
        return r

    def split_and_format_runs(paragraph, formula_spans):
        """Split paragraph runs to apply subscript at given character spans."""
        if not formula_spans:
            return

        # Collect all runs and their text
        runs = list(paragraph._element.findall(f'{{{ns}}}r'))
        run_texts = []
        run_lengths = []
        for r in runs:
            t_elems = r.findall(f'{{{ns}}}t')
            text = ''.join(t.text or '' for t in t_elems)
            run_texts.append(text)
            run_lengths.append(len(text))

        if not run_texts:
            return

        # Build cumulative offsets
        cum_lens = [0]
        for l in run_lengths:
            cum_lens.append(cum_lens[-1] + l)

        # For each subscript range, find which run(s) it spans
        sub_positions = defaultdict(set)  # run_index -> set of char positions to subscript
        for s_start, s_end in formula_spans:
            for i in range(len(runs)):
                r_start = cum_lens[i]
                r_end = cum_lens[i + 1]
                overlap_start = max(s_start, r_start)
                overlap_end = min(s_end, r_end)
                if overlap_start < overlap_end:
                    for pos in range(overlap_start, overlap_end):
                        sub_positions[i].add(pos - r_start)

        if not sub_positions:
            return

        # Rebuild runs with subscript formatting
        parent = paragraph._element
        new_children = []

        for i, r_elem in enumerate(runs):
            if i not in sub_positions:
                new_children.append(r_elem)
                continue

            text = run_texts[i]
            sub_set = sub_positions[i]
            if not sub_set or not text:
                new_children.append(r_elem)
                continue

            # Split this run into normal/subscript segments
            # Copy rPr from original run for normal segments
            orig_rPr = r_elem.find(f'{{{ns}}}rPr')
            segments = []
            current = []
            current_is_sub = False

            for pos, ch in enumerate(text):
                is_sub = pos in sub_set
                if is_sub != current_is_sub and current:
                    segments.append((''.join(current), current_is_sub))
                    current = []
                current_is_sub = is_sub
                current.append(ch)

            if current:
                segments.append((''.join(current), current_is_sub))

            for seg_text, is_sub in segments:
                if is_sub:
                    new_r = make_subscript_element(seg_text)
                    if orig_rPr is not None:
                        rPr_copy = deepcopy(orig_rPr)
                        va = rPr_copy.find(f'{{{ns}}}vertAlign')
                        if va is None:
                            va = OxmlElement('w:vertAlign')
                            rPr_copy.append(va)
                        va.set(qn('w:val'), 'subscript')
                        new_r.remove(new_r[0])  # remove the rPr we added
                        new_r.insert(0, rPr_copy)
                    new_children.append(new_r)
                else:
                    new_r = make_normal_element(seg_text)
                    if orig_rPr is not None:
                        new_r.remove(new_r[0])
                        new_r.insert(0, deepcopy(orig_rPr))
                    new_children.append(new_r)

            # Remove original run
            parent.remove(r_elem)

        # Insert new children at the correct position
        # We need to find where to insert — use the position of the first removed run
        # For simplicity, append all and reorder
        for child in new_children:
            if child.getparent() is None:
                parent.append(child)

    # Process paragraphs
    for pi, para in enumerate(doc.paragraphs):
        text = para.text
        if not text or len(text) < 3:
            continue

        # Find formula candidates
        formula_spans = []
        for m in FORMULA_RE.finditer(text):
            excluded, reason = is_excluded(m, text, m.start())
            if excluded:
                stats['excluded'][reason] += 1
                if reason == "离子价态（应上标）":
                    stats['to_review'].append(f"  ⚠ {m.group(0):30} — 离子价态，需人工确认上标")
                continue

            stats['formulas_found'] += 1
            formula_text = m.group(0)

            # Find subscript ranges within this formula
            sub_ranges = find_subscript_ranges(formula_text)
            if not sub_ranges:
                continue

            # Convert to absolute positions in paragraph text
            abs_ranges = [(m.start() + rs, m.start() + re) for rs, re in sub_ranges]
            formula_spans.extend(abs_ranges)

            # Record modification
            modified = formula_text
            if REBUILD_FROM_RANGES:
                result = list(formula_text)
                for rs, re in reversed(sub_ranges):
                    seg = formula_text[rs:re]
                    subbed = ''.join(SUBSCRIPT_UNICODE.get(c, c) for c in seg)
                    result[rs:re] = subbed
                modified = ''.join(result)

            if len(stats['modifications']) < 20:
                stats['modifications'].append(f"  {formula_text:35} → {modified}")
            stats['formulas_modified'] += 1

        # Apply formatting to paragraph
        if formula_spans and not dry_run:
            split_and_format_runs(para, formula_spans)

    # Also process table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text = para.text
                    if not text or len(text) < 3:
                        continue
                    formula_spans = []
                    for m in FORMULA_RE.finditer(text):
                        excluded, reason = is_excluded(m, text, m.start())
                        if excluded:
                            stats['excluded'][reason] += 1
                            continue
                        stats['formulas_found'] += 1
                        sub_ranges = find_subscript_ranges(m.group(0))
                        if sub_ranges:
                            abs_ranges = [(m.start() + rs, m.start() + re) for rs, re in sub_ranges]
                            formula_spans.extend(abs_ranges)
                            stats['formulas_modified'] += 1
                    if formula_spans and not dry_run:
                        split_and_format_runs(para, formula_spans)

    # Save
    if not dry_run:
        out = output_path or str(Path(filepath).with_stem(Path(filepath).stem + '_subscript_fixed'))
        doc.save(out)
        stats['output_path'] = out
    else:
        stats['output_path'] = '(dry-run, 未保存)'

    return stats


# ── Plain Text / Markdown Processing ───────────────────────────────────

def process_text(filepath, dry_run=False, output_path=None):
    """Process .md/.txt file with Unicode subscript substitution."""
    text = Path(filepath).read_text(encoding='utf-8')
    lines = text.split('\n')

    stats = {
        'formulas_found': 0,
        'formulas_modified': 0,
        'digits_skipped': 0,
        'excluded': defaultdict(int),
        'modifications': [],
        'to_review': [],
    }

    modified_lines = []
    for line in lines:
        if not line.strip():
            modified_lines.append(line)
            continue

        # Find formulas
        matches = list(FORMULA_RE.finditer(line))
        if not matches:
            modified_lines.append(line)
            continue

        # Process matches from right to left to preserve indices
        line_result = list(line)
        for m in reversed(matches):
            excluded, reason = is_excluded(m, line, m.start())
            if excluded:
                stats['excluded'][reason] += 1
                if reason == "离子价态（应上标）":
                    stats['to_review'].append(f"  ⚠ {m.group(0):30} — 离子价态，需人工确认上标")
                continue

            stats['formulas_found'] += 1
            formula_text = m.group(0)
            sub_ranges = find_subscript_ranges(formula_text)
            if not sub_ranges:
                continue

            # Apply Unicode subscripts
            result = list(formula_text)
            for rs, re in reversed(sub_ranges):
                seg = formula_text[rs:re]
                subbed = ''.join(SUBSCRIPT_UNICODE.get(c, c) for c in seg)
                result[rs:re] = subbed
            modified = ''.join(result)

            # Replace in line
            line_result[m.start():m.end()] = modified

            if len(stats['modifications']) < 20:
                stats['modifications'].append(f"  {formula_text:35} → {modified}")
            stats['formulas_modified'] += 1

        modified_lines.append(''.join(line_result))

    result_text = '\n'.join(modified_lines)

    if not dry_run:
        out = output_path or str(Path(filepath).with_stem(Path(filepath).stem + '_subscript_fixed'))
        Path(out).write_text(result_text, encoding='utf-8')
        stats['output_path'] = out
    else:
        stats['output_path'] = '(dry-run, 未保存)'

    return stats


# ── Report Generation ──────────────────────────────────────────────────

# Used in find_subscript_ranges for display purposes
REBUILD_FROM_RANGES = True

def print_report(stats, filepath, mode):
    """Print the modification report."""
    print()
    print("═" * 54)
    print("化学式下标格式化报告")
    print("═" * 54)
    print()
    print(f"  源文件:     {filepath}")
    print(f"  输出文件:   {stats.get('output_path', 'N/A')}")
    print(f"  格式模式:   {mode}")
    print()
    print("─" * 54)
    print("修改统计")
    print("─" * 54)
    print(f"  共识别化学式:     {stats['formulas_found']} 处")
    print(f"  已修改下标:       {stats['formulas_modified']} 处")
    total_excluded = sum(stats['excluded'].values())
    print(f"  跳过（非化学式数字）: {total_excluded} 处")

    if stats['modifications']:
        print()
        print("─" * 54)
        print(f"修改示例（前 {min(20, len(stats['modifications']))} 条）")
        print("─" * 54)
        for mod in stats['modifications'][:20]:
            print(mod)

    if stats['to_review']:
        print()
        print("─" * 54)
        print("需人工复核（价态/电荷/不确定项）")
        print("─" * 54)
        for item in stats['to_review'][:30]:
            print(item)

    if stats['excluded']:
        print()
        print("─" * 54)
        print("未修改的数字类型")
        print("─" * 54)
        for reason, count in sorted(stats['excluded'].items()):
            print(f"  {reason}: {count} 处")

    print("═" * 54)


# ── Main ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='化学式下标格式统一处理工具'
    )
    parser.add_argument('file', help='输入文件路径 (.docx / .md / .txt)')
    parser.add_argument('--dry-run', action='store_true', help='仅检测，不修改')
    parser.add_argument('--output', '-o', help='输出文件路径')
    parser.add_argument('--format', choices=['docx', 'unicode'], help='强制格式模式')
    args = parser.parse_args()

    filepath = Path(args.file)
    if not filepath.exists():
        print(f"错误: 文件不存在 — {args.file}")
        sys.exit(1)

    suffix = filepath.suffix.lower()

    # Auto-detect mode
    if suffix == '.docx':
        mode = 'Word 真实下标 (w:vertAlign)'
    else:
        mode = 'Unicode 下标字符'

    if args.format:
        mode = {'docx': 'Word 真实下标 (w:vertAlign)', 'unicode': 'Unicode 下标字符'}[args.format]

    # Backup
    if not args.dry_run:
        backup = str(filepath) + '.bak'
        shutil.copy2(str(filepath), backup)
        print(f"已备份: {backup}")

    # Process
    if suffix == '.docx':
        stats = process_docx(str(filepath), dry_run=args.dry_run, output_path=args.output)
    else:
        stats = process_text(str(filepath), dry_run=args.dry_run, output_path=args.output)

    # Report
    print_report(stats, args.file, mode)

    if not args.dry_run:
        print(f"\n✅ 已保存: {stats['output_path']}")


if __name__ == '__main__':
    main()
