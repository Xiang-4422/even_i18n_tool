#!/usr/bin/env python3
"""
sanitize_i18n_xlsx.py
清理 i18n xlsx 文案列中的格式问题，规则如下：

  步骤 1：统一为真实换行符
    a. 删除不可见字符：零宽空格 U+200B、BOM U+FEFF、
       行分隔符 U+2028、段落分隔符 U+2029。
    b. 统一行尾序列：\\r\\n / \\r → 真实换行符。
    c. 将字面量 \\n（反斜杠 + n）转为真实换行符（0x0A）。
    经过步骤 1 后，所有换行均为真实换行符，不再有字面量 \\n。

  步骤 2：处理首尾（基于真实换行符）
    a. 首/尾只有换行符           → 自动去除
    b. 首/尾换行符与空格混合     → 警告，不自动处理
    c. 首/尾只有空格（无换行）   → 警告，不自动处理

用法：
  python3 sanitize_i18n_xlsx.py input.xlsx              # 生成 input_sanitized.xlsx
  python3 sanitize_i18n_xlsx.py input.xlsx -o out.xlsx  # 指定输出路径
  python3 sanitize_i18n_xlsx.py input.xlsx --inplace    # 直接修改原文件（自动备份）
"""

import argparse
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("依赖缺失，请先运行：pip3 install openpyxl")


# ── 常量 ────────────────────────────────────────────────────────────────────

# 步骤 1a：需要从文案中完全删除的不可见字符
INVISIBLE_CHARS = (
    "\u200b",  # Zero Width Space  — 零宽空格，复制粘贴常带入
    "\ufeff",  # BOM               — 字节顺序标记，文件开头易混入
    "\u2028",  # Line Separator    — Unicode 行分隔符，渲染行为不一致
    "\u2029",  # Paragraph Sep.    — Unicode 段落分隔符，同上
)

# 变更类型标签
TAG_INVISIBLE = "不可见字符删除"
TAG_NORMALIZE = "换行符统一为真实换行"
TAG_NEWLINE   = "首尾换行符去除"
TAG_MIXED     = "首尾换行与空格去除"

# 警告类型标签
WARN_MIXED       = "首尾换行与空格混合"   # 保留常量供历史兼容，不再由核心产生
WARN_SPACE       = "首尾仅含空格"
WARN_KEY_FORMAT  = "Key 格式错误"
WARN_EMPTY_VALUE = "语言值为空"

# key 合法格式：仅小写字母与下划线
_KEY_RE = re.compile(r'^[a-z0-9_]+$')


# ── 格式化辅助 ───────────────────────────────────────────────────────────────

def fmt(val: str) -> str:
    """将值格式化为可读字符串，含特殊字符时用 repr 显示。"""
    has_special = "\n" in val or "\t" in val or any(c in val for c in INVISIBLE_CHARS)
    return repr(val) if has_special else val


# ── 首尾分析 ─────────────────────────────────────────────────────────────────

def _extract_edges(val: str) -> tuple[str, str]:
    """
    提取字符串首部和尾部的空白类内容（真实换行符和空格/制表符）。
    返回 (leading, trailing)；若整个字符串均为空白，返回 (val, "")。
    """
    n = len(val)

    # 从头扫描
    i = 0
    while i < n and val[i] in "\n \t":
        i += 1

    # 从尾扫描（只扫 i 之后部分，防止与首部重叠）
    j = n
    while j > i and val[j - 1] in "\n \t":
        j -= 1

    if i >= j:
        return val, ""

    return val[:i], val[j:]


def classify_edge(chars: str) -> str:
    """
    判断首/尾空白内容的类型：
      'clean'        — 无空白
      'newline_only' — 只有换行符
      'space_only'   — 只有空格/制表符（无换行）
      'mixed'        — 换行符与空格/制表符混合
    """
    if not chars:
        return "clean"
    has_newline = "\n" in chars
    has_space   = bool(set(chars) - {"\n"})
    if has_newline and has_space:
        return "mixed"
    if has_newline:
        return "newline_only"
    return "space_only"


# ── 核心处理逻辑 ─────────────────────────────────────────────────────────────

def process_cell(val: str) -> tuple[str, set[str], list[tuple[str, str]]]:
    """
    对单个文案值执行两步清理。
    返回 (处理后的值, 触发的变更类型集合, 警告信息列表)。
    """
    tags:     set[str]              = set()
    warnings: list[tuple[str, str]] = []   # (类型标签, 位置详情)

    # ── 步骤 1：统一为真实换行符 ─────────────────────────────
    #
    # 1a. 删除不可见字符（任意位置）
    for ch in INVISIBLE_CHARS:
        if ch in val:
            val = val.replace(ch, "")
            tags.add(TAG_INVISIBLE)

    # 1b. 统一行尾序列：\r\n / \r → \n
    val = val.replace("\r\n", "\n").replace("\r", "\n")

    # 1c. 字面量 \n（反斜杠 + n）→ 真实换行符（0x0A）
    # 经过此步，字符串里所有换行均为真实换行符，不再有字面量 \n。
    if r"\n" in val:
        val = val.replace(r"\n", "\n")
        tags.add(TAG_NORMALIZE)

    # ── 步骤 2：处理首尾（基于真实换行符）──────────────────
    #
    # 提取首尾内容
    leading, trailing = _extract_edges(val)

    # 若整个字符串均为空白，报一次警告后跳过
    if leading == val and not trailing:
        warnings.append((WARN_SPACE, f"整个值均为空白：{repr(val)}"))
        return val, tags, warnings

    # 首部处理
    lead_type = classify_edge(leading)
    if lead_type == "newline_only":
        val = val[len(leading):]
        tags.add(TAG_NEWLINE)
    elif lead_type == "mixed":
        # 换行与空格混合 → 自动去除所有首部空白
        val = val.lstrip("\n \t")
        tags.add(TAG_MIXED)
    elif lead_type == "space_only":
        warnings.append((WARN_SPACE, f"行首：{repr(leading)}"))

    # 尾部处理（首部处理后重新提取，防止偏移）
    _, trailing = _extract_edges(val)
    trail_type = classify_edge(trailing)
    if trail_type == "newline_only":
        val = val[:len(val) - len(trailing)]
        tags.add(TAG_NEWLINE)
    elif trail_type == "mixed":
        # 换行与空格混合 → 自动去除所有尾部空白
        val = val.rstrip("\n \t")
        tags.add(TAG_MIXED)
    elif trail_type == "space_only":
        warnings.append((WARN_SPACE, f"行尾：{repr(trailing)}"))

    return val, tags, warnings


# ── 核心处理：workbook 级别（供 CLI 和 GUI 共用）───────────────────────────

Change   = tuple[str, int, int, str, str, str, str, set[str]]   # +col_idx at [2]
WarnItem = tuple[str, int, int, str, str, str, list[tuple[str, str]]]  # +col_idx at [2]


def process_workbook(wb) -> tuple[list[Change], list[WarnItem]]:
    """
    遍历 workbook 所有工作表，对每个文案单元格执行 process_cell。
    直接修改传入的 workbook 对象，同时返回变更列表和警告列表。
    """
    changes:    list[Change]   = []
    warn_items: list[WarnItem] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        header_row = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

        for row in ws.iter_rows(min_row=2):
            raw_key = row[1].value if len(row) > 1 else None
            key = str(raw_key) if raw_key is not None else ""
            row_num = row[0].row

            # ── Key 格式校验 ─────────────────────────────────────
            key_col = row[1].column if len(row) > 1 else 2
            if not key:
                warn_items.append((sheet_name, row_num, key_col, key, "Key", key,
                                   [(WARN_KEY_FORMAT, "Key 为空")]))
            elif not _KEY_RE.match(key):
                warn_items.append((sheet_name, row_num, key_col, key, "Key", key,
                                   [(WARN_KEY_FORMAT, "只允许小写字母、数字和下划线")]))

            for cell in row[2:]:
                # ── 空值校验（处理前检查原始值）──────────────────
                if cell.value is None or (isinstance(cell.value, str)
                                          and cell.value.strip() == ""):
                    col_h = str(header_row[cell.column - 1] or f"列{cell.column}").replace("\n", " ").strip()
                    warn_items.append((sheet_name, cell.row, cell.column, key, col_h, "",
                                       [(WARN_EMPTY_VALUE, "翻译值为空")]))
                    continue

                val = cell.value
                if not isinstance(val, str):
                    continue

                new_val, tags, cell_warnings = process_cell(val)

                col_header = header_row[cell.column - 1] or f"列{cell.column}"
                col_header = str(col_header).replace("\n", " ").strip()

                if new_val != val:
                    changes.append((sheet_name, cell.row, cell.column, key, col_header, val, new_val, tags))
                    cell.value = new_val

                if cell_warnings:
                    warn_items.append((sheet_name, cell.row, cell.column, key, col_header, new_val, cell_warnings))

    return changes, warn_items


# ── 主处理流程（CLI 入口）────────────────────────────────────────────────────

def sanitize(src: Path, dst: Path) -> None:
    if src != dst:
        shutil.copy2(src, dst)

    wb = openpyxl.load_workbook(dst)
    changes, warn_items = process_workbook(wb)
    wb.save(dst)

    print(f"\n输出文件：{dst}\n")

    if changes:
        print("【变更】")
        cur_sheet = None
        for sheet, row, col_idx, key, col, old, new, tags in changes:
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

    if warn_items:
        print()
        print("【警告】（需人工确认，未自动处理）")
        cur_sheet = None
        for sheet, row, col_idx, key, col, val, msgs in warn_items:
            if sheet != cur_sheet:
                print(f"┌─ {sheet}")
                cur_sheet = sheet
            print(f"│  第{row:>4}行  [{col}]  key={key}")
            print(f"│    值：{fmt(val)}")
            for warn_type, detail in msgs:
                print(f"│    ⚠ 原因：{warn_type}  {detail}")

    sheet_change_counts: dict[str, int] = defaultdict(int)
    sheet_warn_counts:   dict[str, int] = defaultdict(int)
    tag_counts:          dict[str, int] = defaultdict(int)
    warn_type_counts:    dict[str, int] = defaultdict(int)

    for _, _, _, _, _, _, _, tags in changes:
        for tag in tags:
            tag_counts[tag] += 1
    for sheet, *_ in changes:
        sheet_change_counts[sheet] += 1
    for sheet, _, _, _, _, _, msgs in warn_items:
        sheet_warn_counts[sheet] += 1
        for warn_type, _ in msgs:
            warn_type_counts[warn_type] += 1

    print()
    print("─" * 44)
    print("  处理总结")
    print("─" * 44)
    print(f"  修改单元格  {len(changes):>6} 个")
    print(f"  警告单元格  {len(warn_items):>6} 个")
    print()
    print("  变更类型：")
    for tag in (TAG_INVISIBLE, TAG_NORMALIZE, TAG_NEWLINE, TAG_MIXED):
        count = tag_counts.get(tag, 0)
        if count:
            print(f"    {tag:<16}  {count:>6} 个")
    if warn_type_counts:
        print()
        print("  警告类型：")
        for kind in (WARN_MIXED, WARN_SPACE, WARN_KEY_FORMAT, WARN_EMPTY_VALUE):
            count = warn_type_counts.get(kind, 0)
            if count:
                print(f"    {kind:<16}  {count:>6} 个")
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
        description="清理 i18n xlsx 文案列",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input", type=Path, help="输入 xlsx 文件路径")
    parser.add_argument("-o", "--output", type=Path, default=None,
                        help="输出路径（默认加 _sanitized 后缀）")
    parser.add_argument("--inplace", action="store_true",
                        help="直接修改原文件（自动备份为 .bak.xlsx）")
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

    sanitize(src, dst)


if __name__ == "__main__":
    main()
