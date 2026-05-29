#!/usr/bin/env python3
"""
sanitize_i18n_xlsx.py
清理 i18n xlsx 文案列中的格式问题，规则如下：

  1. 删除不可见字符：零宽空格 U+200B、BOM U+FEFF、
     行分隔符 U+2028、段落分隔符 U+2029。
     这些字符不应出现在任何文案中，无论位置。

  2. 首尾换行符处理（分三种情况）：
     a. 首/尾只有换行符           → 自动去除
     b. 首/尾换行符与空格混合存在  → 仅警告，不自动处理
     c. 首/尾只有空格（无换行）    → 仅警告，不自动处理

  3. 字符串内部内容完全不修改：
     真实换行符、字面量 \\n、空格均原样保留。

  4. [可选，加 --escape 参数启用]
     将行内真实换行符转为字面量 \\n（反斜杠 + n），
     使表格内不再存在真实换行符。

用法：
  python3 sanitize_i18n_xlsx.py input.xlsx               # 生成 input_sanitized.xlsx
  python3 sanitize_i18n_xlsx.py input.xlsx -o out.xlsx   # 指定输出路径
  python3 sanitize_i18n_xlsx.py input.xlsx --inplace     # 直接修改原文件（自动备份）
  python3 sanitize_i18n_xlsx.py input.xlsx --escape      # 同时将行内换行符转为字面量 \\n
"""

import argparse
import shutil
import sys
from collections import defaultdict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("依赖缺失，请先运行：pip3 install openpyxl")


# ── 常量 ────────────────────────────────────────────────────────────────────

# 规则 1：需要从文案中完全删除的不可见字符
INVISIBLE_CHARS = (
    "\u200b",  # Zero Width Space  — 零宽空格，复制粘贴常带入
    "\ufeff",  # BOM               — 字节顺序标记，文件开头易混入
    "\u2028",  # Line Separator    — Unicode 行分隔符，渲染行为不一致
    "\u2029",  # Paragraph Sep.    — Unicode 段落分隔符，同上
)

# 变更类型标签，用于最终汇总统计
TAG_INVISIBLE = "不可见字符删除"
TAG_NEWLINE   = "首尾换行符去除"
TAG_ESCAPE    = "行内换行符转义"

# 警告类型标签
WARN_MIXED    = "首尾换行与空格混合"
WARN_SPACE    = "首尾仅含空格"


# ── 格式化辅助 ───────────────────────────────────────────────────────────────

def fmt(val: str) -> str:
    """将值格式化为可读字符串，含特殊字符时用 repr 显示以便看清差异。"""
    has_special = "\n" in val or "\t" in val or any(c in val for c in INVISIBLE_CHARS)
    return repr(val) if has_special else val


def classify_edge(chars: str) -> str:
    """
    判断首/尾空白字符的类型：
      'clean'        — 无空白
      'newline_only' — 只有换行符
      'space_only'   — 只有空格/制表符（无换行）
      'mixed'        — 换行符与空格/制表符混合
    """
    if not chars:
        return "clean"
    has_newline = "\n" in chars
    has_space   = bool(set(chars) - {"\n"})   # 除换行外的其他空白
    if has_newline and has_space:
        return "mixed"
    if has_newline:
        return "newline_only"
    return "space_only"


# ── 核心处理逻辑 ─────────────────────────────────────────────────────────────

def process_cell(val: str, escape: bool) -> tuple[str, set[str], list[str]]:
    """
    对单个文案值执行清理。
    escape=True 时额外执行步骤 4（行内换行符转义）。
    返回 (处理后的值, 触发的变更类型集合, 警告信息列表)。
    """
    tags:     set[str]             = set()
    warnings: list[tuple[str,str]] = []   # (类型标签, 位置详情)

    # ── 步骤 1：删除不可见字符 ───────────────────────────────
    # 遍历常量列表，将每个不可见字符从字符串中完全移除。
    # 不限位置：行首、行尾、行内均删除。
    for ch in INVISIBLE_CHARS:
        if ch in val:
            val = val.replace(ch, "")
            tags.add(TAG_INVISIBLE)

    # ── 步骤 2：首尾空白分析与处理 ──────────────────────────
    # 提取首部和尾部的空白字符，分别判断类型后决定操作。
    #
    # 注意：先单独提取首/尾，再判断，避免两端互相干扰。
    leading  = val[: len(val) - len(val.lstrip("\n \t"))]
    trailing = val[len(val.rstrip("\n \t")):]

    # 若首尾重叠（整个字符串全为空白），只报一次警告，跳过后续处理
    if len(leading) + len(trailing) > len(val):
        warnings.append((WARN_SPACE, f"整个值均为空白字符：{repr(val)}"))
    else:
        lead_type  = classify_edge(leading)
        trail_type = classify_edge(trailing)

        # 首部处理
        if lead_type == "newline_only":
            # 只有换行 → 自动去除
            val = val.lstrip("\n")
            tags.add(TAG_NEWLINE)
        elif lead_type == "mixed":
            # 换行与空格混合 → 警告，不动
            warnings.append((WARN_MIXED, f"行首：{repr(leading)}"))
        elif lead_type == "space_only":
            # 只有空格 → 警告，不动
            warnings.append((WARN_SPACE, f"行首：{repr(leading)}"))

        # 尾部处理（在首部已处理后重新提取，防止错位）
        trailing = val[len(val.rstrip("\n \t")):]
        trail_type = classify_edge(trailing)

        if trail_type == "newline_only":
            # 只有换行 → 自动去除
            val = val.rstrip("\n")
            tags.add(TAG_NEWLINE)
        elif trail_type == "mixed":
            # 换行与空格混合 → 警告，不动
            warnings.append((WARN_MIXED, f"行尾：{repr(trailing)}"))
        elif trail_type == "space_only":
            # 只有空格 → 警告，不动
            warnings.append((WARN_SPACE, f"行尾：{repr(trailing)}"))

    # ── 步骤 3：行内内容保持不变 ─────────────────────────────
    # 字符串内部的真实换行符、字面量 \n、空格均不做任何处理。
    # 本步骤无代码，仅作说明。

    # ── 步骤 4（可选）：行内换行符转义 ──────────────────────
    # 仅在 --escape 参数启用时执行。
    # 将字符串内部残余的真实换行符（\n）替换为字面量 \n，
    # 使表格单元格内不再出现真实换行，便于统一文档格式。
    #
    # 若步骤 2 已产生警告（首尾存在混合或空格问题），
    # 则跳过本步骤，避免自动修改掩盖需人工确认的问题。
    if escape and "\n" in val and not warnings:
        val = val.replace("\n", r"\n")
        tags.add(TAG_ESCAPE)

    return val, tags, warnings


# ── 核心处理：workbook 级别（供 CLI 和 GUI 共用）───────────────────────────

Change   = tuple[str, int, str, str, str, str, set[str]]   # sheet,row,key,col,old,new,tags
WarnItem = tuple[str, int, str, str, str, list[tuple[str,str]]]  # sheet,row,key,col,val,msgs

def process_workbook(wb, escape: bool) -> tuple[list[Change], list[WarnItem]]:
    """
    遍历 workbook 所有工作表，对每个文案单元格执行 process_cell。
    直接修改传入的 workbook 对象，同时返回变更列表和警告列表。
    """
    changes:    list[Change]   = []
    warn_items: list[WarnItem] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # 读取第 1 行作为列标题，用于日志显示
        header_row = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

        for row in ws.iter_rows(min_row=2):
            # 第 2 列是 Key，仅用于日志，不参与修改
            key = str(row[1].value) if len(row) > 1 else ""

            # 第 3 列起（索引 2+）是各语言文案列
            for cell in row[2:]:
                val = cell.value
                if not isinstance(val, str):
                    continue

                new_val, tags, cell_warnings = process_cell(val, escape)

                # 取列标题，清理标题自身可能含有的换行
                col_header = header_row[cell.column - 1] or f"列{cell.column}"
                col_header = str(col_header).replace("\n", " ").strip()

                if new_val != val:
                    changes.append((sheet_name, cell.row, key, col_header, val, new_val, tags))
                    cell.value = new_val

                if cell_warnings:
                    warn_items.append((sheet_name, cell.row, key, col_header, new_val, cell_warnings))

    return changes, warn_items


# ── 主处理流程（CLI 入口）────────────────────────────────────────────────────

def sanitize(src: Path, dst: Path, escape: bool) -> None:
    if src != dst:
        shutil.copy2(src, dst)

    wb = openpyxl.load_workbook(dst)
    changes, warn_items = process_workbook(wb, escape)
    wb.save(dst)

    # ── 输出变更详情 ─────────────────────────────────────────
    print(f"\n输出文件：{dst}\n")

    if changes:
        print("【变更】")
        cur_sheet = None
        for sheet, row, key, col, old, new, tags in changes:
            if sheet != cur_sheet:
                print(f"┌─ {sheet}")
                cur_sheet = sheet
            reason = "、".join(sorted(tags))
            print(f"│  第{row:>4}行  [{col}]  key={key}")
            print(f"│    原因：{reason}")
            print(f"│    旧：{fmt(old)}")
            print(f"│    新：{fmt(new)}")
    else:
        print("无变更。")

    # ── 输出警告详情 ─────────────────────────────────────────
    if warn_items:
        print()
        print("【警告】（需人工确认，未自动处理）")
        cur_sheet = None
        for sheet, row, key, col, val, msgs in warn_items:
            if sheet != cur_sheet:
                print(f"┌─ {sheet}")
                cur_sheet = sheet
            print(f"│  第{row:>4}行  [{col}]  key={key}")
            print(f"│    值：{fmt(val)}")
            for warn_type, detail in msgs:
                print(f"│    ⚠ 原因：{warn_type}  {detail}")

    # ── 输出总结 ──────────────────────────────────────────────
    sheet_change_counts: dict[str, int] = defaultdict(int)
    sheet_warn_counts:   dict[str, int] = defaultdict(int)
    tag_counts:          dict[str, int] = defaultdict(int)
    warn_type_counts:    dict[str, int] = defaultdict(int)

    for _, _, _, _, _, _, tags in changes:
        for tag in tags:
            tag_counts[tag] += 1
    for sheet, *_ in changes:
        sheet_change_counts[sheet] += 1

    for sheet, _, _, _, _, msgs in warn_items:
        sheet_warn_counts[sheet] += 1
        for warn_type, _ in msgs:
            warn_type_counts[warn_type] += 1

    all_tags = [TAG_INVISIBLE, TAG_NEWLINE]
    if escape:
        all_tags.append(TAG_ESCAPE)

    print()
    print("─" * 44)
    print("  处理总结")
    print("─" * 44)
    print(f"  修改单元格  {len(changes):>6} 个")
    print(f"  警告单元格  {len(warn_items):>6} 个")
    print()
    print("  变更类型：")
    for tag in all_tags:
        print(f"    {tag:<14}  {tag_counts.get(tag, 0):>6} 个")
    if warn_type_counts:
        print()
        print("  警告类型：")
        for kind in (WARN_MIXED, WARN_SPACE):
            count = warn_type_counts.get(kind, 0)
            if count:
                print(f"    {kind:<14}  {count:>6} 个")
    print()
    print("  按工作表（修改 / 警告）：")
    for sheet_name in wb.sheetnames:
        c = sheet_change_counts.get(sheet_name, 0)
        w = sheet_warn_counts.get(sheet_name, 0)
        if c or w:
            print(f"    {sheet_name:<22}  {c:>4} 改  {w:>4} 警")
    print("─" * 44)


# ── 入口 ─────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="清理 i18n xlsx 文案列中的不可见字符和首尾换行符",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input", type=Path, help="输入 xlsx 文件路径")
    parser.add_argument("-o", "--output", type=Path, default=None,
                        help="输出路径（默认在原文件名后加 _sanitized）")
    parser.add_argument("--inplace", action="store_true",
                        help="直接修改原文件（自动备份为 .bak.xlsx）")
    parser.add_argument("--escape", action="store_true",
                        help="将行内真实换行符转为字面量 \\n（步骤 4，默认不启用）")
    args = parser.parse_args()

    src: Path = args.input.resolve()
    if not src.exists():
        sys.exit(f"文件不存在：{src}")
    if src.suffix.lower() != ".xlsx":
        sys.exit(f"仅支持 .xlsx 文件：{src}")

    if args.inplace:
        bak = src.with_suffix(".bak.xlsx")
        shutil.copy2(src, bak)
        print(f"备份已保存：{bak}")
        dst = src
    elif args.output:
        dst = args.output.resolve()
    else:
        dst = src.with_stem(src.stem + "_sanitized")

    sanitize(src, dst, escape=args.escape)


if __name__ == "__main__":
    main()
