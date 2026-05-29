"""
sanitize_i18n_app.py
i18n XLSX 清理工具 — Streamlit 图形界面

启动方式：
  python3 -m streamlit run sanitize_i18n_app.py
"""

import difflib
import io
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
import streamlit as st

# 从同目录的脚本文件导入核心逻辑和常量
sys.path.insert(0, str(Path(__file__).parent))
from sanitize_i18n_xlsx import (
    TAG_INVISIBLE, TAG_NORMALIZE, TAG_NEWLINE,
    WARN_MIXED, WARN_SPACE,
    INVISIBLE_CHARS,
    process_workbook,
)


# ── 显示用映射 ───────────────────────────────────────────────────────────────

# 不可见字符 → 人类可读名称（GUI 提示用；核心未提供则在此维护）
INVISIBLE_NAMES = {
    "​": "零宽空格 U+200B",
    "﻿": "BOM U+FEFF",
    " ": "行分隔符 U+2028",
    " ": "段落分隔符 U+2029",
}

# 变更/警告类型 → 筛选条上的短标签
TAG_SHORT = {
    TAG_INVISIBLE: "不可见字符",
    TAG_NORMALIZE: "换行统一",
    TAG_NEWLINE:   "首尾换行",
}
WARN_SHORT = {
    WARN_MIXED: "换行+空格",
    WARN_SPACE: "纯空格",
}
# 固定展示顺序
TAG_ORDER  = (TAG_INVISIBLE, TAG_NORMALIZE, TAG_NEWLINE)
WARN_ORDER = (WARN_MIXED, WARN_SPACE)


# ── 页面配置 ─────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="i18n XLSX 清理工具",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ── 全局样式 ─────────────────────────────────────────────────────────────────

st.markdown("""
<style>
/* 卡片 */
.change-row { background:#f1f8f2; border-left:4px solid #2e7d32;
              padding:10px 14px; border-radius:6px; margin-bottom:10px; }
.warn-row   { background:#fffaf0; border-left:4px solid #f9a825;
              padding:10px 14px; border-radius:6px; margin-bottom:10px; }
.card-head  { font-size:0.82rem; color:#444; font-weight:600; margin-bottom:6px; }
.card-head .sep     { color:#bbb; margin:0 4px; }
.card-head .col-tag { color:#1565c0; }
.label      { font-size:0.74rem; color:#888; margin:4px 0 1px; }
.mono       { font-family:'SF Mono',Menlo,Consolas,monospace;
              font-size:0.9rem; word-break:break-all; line-height:1.7; }
.diff-row   { display:flex; align-items:baseline; gap:6px; margin:3px 0; }
.diff-label { font-size:0.76rem; color:#888; flex-shrink:0; white-space:nowrap; }
.str-pill   { display:inline-block; padding:3px 12px; border-radius:13px;
              background:#eceff1; max-width:100%; word-break:break-all; }
/* 标签 chips */
.chips      { margin-top:7px; }
.chip       { display:inline-block; background:#e8f5e9; color:#2e7d32;
              border:1px solid #c8e6c9; border-radius:11px;
              padding:1px 10px; font-size:0.74rem; margin:2px 5px 0 0; }
.warn-row .chip { background:#fff3e0; color:#e65100; border-color:#ffe0b2; }
/* 特殊字符上色 */
.real-nl    { color:#1565c0; font-weight:bold; }   /* 真实换行 ↵ */
.lit-n      { color:#e65100; font-weight:600; }     /* 字面量 \\n */
.invis-char { color:#9e9e9e; font-weight:bold; }    /* 不可见 ⬚ */
.space-char { color:#c0c0c0; }                       /* 空格 ␣ */
/* 图例 */
.legend     { background:#fafafa; border:1px solid #eee; border-radius:6px;
              padding:7px 14px; font-size:0.8rem; color:#666; margin-bottom:14px; }
.legend b   { color:#444; }
.legend .real-nl,.legend .lit-n,.legend .invis-char,.legend .space-char
            { font-size:1rem; margin:0 2px; }
/* 总结比例条 */
.ratio-row  { margin-bottom:9px; }
.ratio-head { display:flex; justify-content:space-between;
              font-size:0.85rem; margin-bottom:2px; }
.ratio-track{ background:#eee; border-radius:4px; height:14px; overflow:hidden; }
.ratio-fill { height:100%; border-radius:4px; }
/* 让 radio 横排看起来像 Tab */
div[data-testid="stRadio"] > div { display:flex; gap:8px; flex-wrap:wrap; }
div[data-testid="stRadio"] label {
    border:1px solid #ddd; border-radius:6px; padding:4px 14px;
    cursor:pointer; font-size:0.9rem;
}
div[data-testid="stRadio"] label:has(input:checked) {
    background:#f0f2f6; border-color:#aaa; font-weight:600;
}
</style>
""", unsafe_allow_html=True)


# ── HTML 显示格式化 ───────────────────────────────────────────────────────────

def _esc(s) -> str:
    """HTML 转义（用于普通文本片段，如 sheet / key / 列名）。"""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _fmt_html(val: str, show_spaces: bool = False) -> str:
    """
    将值转为 HTML 安全的可读字符串，用颜色区分特殊字符：
      · 真实换行符 (0x0A)    → 蓝色 ↵
      · 字面量 \\n（反斜杠+n） → 橙色 \\n
      · 不可见字符           → 灰色 ⬚（悬停显示具体名称）
      · 空格（show_spaces）   → 浅灰 ␣
    """
    segments = val.split("\n")          # 按真实换行切分
    processed = []
    for seg in segments:
        # 1. HTML 转义
        s = seg.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        # 2. 空格先占位，避免污染随后插入的标签属性（如 title 内的空格）
        if show_spaces:
            s = s.replace(" ", "\x00")
        # 3. 不可见字符 → 灰色 ⬚ + 悬停名称
        for ch in INVISIBLE_CHARS:
            name = INVISIBLE_NAMES.get(ch, f"U+{ord(ch):04X}")
            s = s.replace(ch, f'<span class="invis-char" title="{name}">⬚</span>')
        # 4. 字面量 \n → 橙色
        s = s.replace("\\n", '<span class="lit-n">\\n</span>')
        # 5. 占位还原为 ␣
        if show_spaces:
            s = s.replace("\x00", '<span class="space-char">␣</span>')
        processed.append(s)
    # 6. 用蓝色 ↵ 连接各段，代表真实换行符
    return '<span class="real-nl">↵</span>'.join(processed)


def _diff_html(old: str, new: str) -> tuple[str, str]:
    """逐字符比较，差异部分加背景高亮（红=删除/旧，绿=新增/新）。"""
    matcher = difflib.SequenceMatcher(None, old, new, autojunk=False)
    old_parts: list[str] = []
    new_parts: list[str] = []
    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        old_seg = _fmt_html(old[i1:i2])
        new_seg = _fmt_html(new[j1:j2])
        if op == "equal":
            old_parts.append(old_seg)
            new_parts.append(new_seg)
        elif op == "replace":
            old_parts.append(f'<mark style="background:#ffcdd2;border-radius:2px">{old_seg}</mark>')
            new_parts.append(f'<mark style="background:#c8e6c9;border-radius:2px">{new_seg}</mark>')
        elif op == "delete":
            old_parts.append(f'<mark style="background:#ffcdd2;border-radius:2px">{old_seg}</mark>')
        elif op == "insert":
            new_parts.append(f'<mark style="background:#c8e6c9;border-radius:2px">{new_seg}</mark>')
    return "".join(old_parts), "".join(new_parts)


def render_legend() -> None:
    """图例条：解释颜色与符号。"""
    st.markdown(
        '<div class="legend"><b>图例</b>　'
        '<span class="real-nl">↵</span>真实换行　'
        '<span class="lit-n">\\n</span>字面量　'
        '<span class="space-char">␣</span>空格　'
        '<span class="invis-char">⬚</span>不可见字符（悬停看详情）　'
        '<mark style="background:#c8e6c9;border-radius:2px">绿底</mark>新增　'
        '<mark style="background:#ffcdd2;border-radius:2px">红底</mark>删除'
        '</div>',
        unsafe_allow_html=True,
    )


# ── 文案辅助 ─────────────────────────────────────────────────────────────────

def _change_reasons(old: str, tags: set) -> list:
    """根据变更类型生成人话标签（不可见字符列出具体名称）。"""
    out = []
    if TAG_INVISIBLE in tags:
        names = [INVISIBLE_NAMES.get(ch, f"U+{ord(ch):04X}")
                 for ch in INVISIBLE_CHARS if ch in old]
        out.append("删除 " + "、".join(names) if names else "删除不可见字符")
    if TAG_NORMALIZE in tags:
        out.append("字面量 \\n → 真实换行")
    if TAG_NEWLINE in tags:
        out.append("去除首尾换行")
    return out


def _humanize_warn(val: str, detail: str) -> str:
    """把核心的 repr 警告详情（如 行尾：'\\n  '）转成人话计数。"""
    stripped_l = val.lstrip("\n \t")
    stripped_r = val.rstrip("\n \t")
    leading  = val[: len(val) - len(stripped_l)]
    trailing = val[len(stripped_r):]

    if detail.startswith("整个值") or not stripped_l:
        seg, loc = val, "整个单元格"
    elif detail.startswith("行首"):
        seg, loc = leading, "行首"
    elif detail.startswith("行尾"):
        seg, loc = trailing, "行尾"
    else:
        seg, loc = val, ""

    n_nl = seg.count("\n")
    n_sp = len(seg) - n_nl
    bits = []
    if n_nl:
        bits.append(f"换行符 ×{n_nl}")
    if n_sp:
        bits.append(f"空格 ×{n_sp}")
    desc = " + ".join(bits) if bits else "空白"

    if loc == "整个单元格":
        return f"整个单元格仅由空白组成（{desc}）"
    return f"{loc}有 {desc}"


# ── 筛选条 ───────────────────────────────────────────────────────────────────

def _filter_bar(prefix: str, type_opts: list,
                sheets: list, cols: list) -> tuple:
    """渲染筛选控件，返回 (sel_type, query, sel_sheet, sel_col)。"""
    c1, c2 = st.columns([3, 2])
    with c1:
        sel_type = st.segmented_control(
            "类型", ["全部"] + type_opts, default="全部",
            key=f"flt_{prefix}_type", label_visibility="collapsed",
        ) or "全部"
    with c2:
        query = st.text_input(
            "搜索", placeholder="🔍 搜索 key 或内容",
            key=f"flt_{prefix}_q", label_visibility="collapsed",
        )

    sel_sheet = sel_col = "全部"
    if len(sheets) > 1 or len(cols) > 1:
        c3, c4 = st.columns(2)
        if len(sheets) > 1:
            sel_sheet = c3.selectbox("工作表", ["全部"] + sheets,
                                     key=f"flt_{prefix}_sheet")
        if len(cols) > 1:
            sel_col = c4.selectbox("语言列", ["全部"] + cols,
                                   key=f"flt_{prefix}_col")
    return sel_type, query, sel_sheet, sel_col


def _reset_filters() -> None:
    """处理新文件时清除所有筛选状态。"""
    for k in list(st.session_state.keys()):
        if k.startswith("flt_"):
            del st.session_state[k]


# ── 渲染：变更 ───────────────────────────────────────────────────────────────

def render_changes(changes: list) -> None:
    render_legend()

    present = set().union(*(tags for *_, tags in changes)) if changes else set()
    type_opts = [TAG_SHORT[t] for t in TAG_ORDER if t in present]
    sheets = list(dict.fromkeys(s for s, *_ in changes))
    cols   = list(dict.fromkeys(c for _, _, _, c, *_ in changes))

    sel_type, query, sel_sheet, sel_col = _filter_bar("c", type_opts, sheets, cols)
    q = query.lower().strip()

    # 先按筛选条件过滤
    filtered = []
    for item in changes:
        sheet, row, key, col, old, new, tags = item
        if sel_type != "全部" and sel_type not in {TAG_SHORT.get(t, t) for t in tags}:
            continue
        if sel_sheet != "全部" and sheet != sel_sheet:
            continue
        if sel_col != "全部" and col != sel_col:
            continue
        if q and q not in f"{key}\n{old}\n{new}".lower():
            continue
        filtered.append(item)

    st.caption(f"显示 {len(filtered)} / {len(changes)} 处变更"
               + ("（无匹配项）" if not filtered else ""))

    # 按 key 分组（保持出现顺序）：同一 key 的多处变更归为一组、可折叠
    groups: dict = {}
    for item in filtered:
        groups.setdefault(item[2], []).append(item)

    for key, items in groups.items():
        label = key if key else "（无 key）"
        with st.expander(f"🔑 {label}　{len(items)} 处", expanded=True):
            for sheet, row, _key, col, old, new, tags in items:
                old_d, new_d = _diff_html(old, new)
                chips = "".join(f'<span class="chip">{_esc(r)}</span>'
                                for r in _change_reasons(old, tags))
                st.markdown(
                    f'<div class="change-row">'
                    f'<div class="card-head">📄 {_esc(sheet)}<span class="sep">·</span>'
                    f'第 {row} 行 <span class="col-tag">[{_esc(col)}]</span></div>'
                    f'<div class="diff-row"><span class="diff-label">原始字符：</span>'
                    f'<span class="str-pill mono">{old_d}</span></div>'
                    f'<div class="diff-row"><span class="diff-label">修正字符：</span>'
                    f'<span class="str-pill mono">{new_d}</span></div>'
                    f'<div class="chips">{chips}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )


# ── 渲染：警告 ───────────────────────────────────────────────────────────────

def render_warnings(warn_items: list) -> None:
    render_legend()

    present = set()
    for *_, msgs in warn_items:
        present |= {wt for wt, _ in msgs}
    type_opts = [WARN_SHORT[t] for t in WARN_ORDER if t in present]
    sheets = list(dict.fromkeys(s for s, *_ in warn_items))
    cols   = list(dict.fromkeys(c for _, _, _, c, *_ in warn_items))

    sel_type, query, sel_sheet, sel_col = _filter_bar("w", type_opts, sheets, cols)
    q = query.lower().strip()

    shown = 0
    for sheet, row, key, col, val, msgs in warn_items:
        types_here = {WARN_SHORT.get(wt, wt) for wt, _ in msgs}
        if sel_type != "全部" and sel_type not in types_here:
            continue
        if sel_sheet != "全部" and sheet != sel_sheet:
            continue
        if sel_col != "全部" and col != sel_col:
            continue
        if q and q not in f"{key}\n{val}".lower():
            continue

        shown += 1
        val_d = _fmt_html(val, show_spaces=True)
        chips = "".join(f'<span class="chip">⚠ {_esc(_humanize_warn(val, d))}</span>'
                        for _, d in msgs)
        st.markdown(
            f'<div class="warn-row">'
            f'<div class="card-head">📄 {_esc(sheet)}<span class="sep">·</span>'
            f'🔑 {_esc(key)}<span class="sep">·</span>第 {row} 行 '
            f'<span class="col-tag">[{_esc(col)}]</span></div>'
            f'<div class="label">当前值</div>'
            f'<div class="mono">{val_d}</div>'
            f'<div class="chips">{chips}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.caption(f"显示 {shown} / {len(warn_items)} 处警告"
               + ("（无匹配项）" if shown == 0 else ""))


# ── 渲染：总结 ───────────────────────────────────────────────────────────────

def _ratio_bar(label: str, count: int, total: int, color: str) -> None:
    pct = (count / total * 100) if total else 0
    st.markdown(
        f'<div class="ratio-row">'
        f'<div class="ratio-head"><span>{_esc(label)}</span><span><b>{count}</b></span></div>'
        f'<div class="ratio-track"><div class="ratio-fill" '
        f'style="width:{pct:.1f}%;background:{color}"></div></div></div>',
        unsafe_allow_html=True,
    )


def render_summary(changes: list, warn_items: list) -> None:
    tag_counts:       dict = defaultdict(int)
    warn_type_counts: dict = defaultdict(int)
    sheet_c:          dict = defaultdict(int)
    sheet_w:          dict = defaultdict(int)

    for *_, tags in changes:
        for tag in tags:
            tag_counts[tag] += 1
    for sheet, *_ in changes:
        sheet_c[sheet] += 1
    for sheet, _, _, _, _, msgs in warn_items:
        sheet_w[sheet] += 1
        for wt, _ in msgs:
            warn_type_counts[wt] += 1

    c1, c2 = st.columns(2)
    c1.metric("✅ 修改单元格", len(changes))
    c2.metric("⚠️ 警告单元格", len(warn_items))
    st.divider()

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**变更类型**")
        tc_total = max(sum(tag_counts.values()), 1)
        if any(tag_counts.get(t) for t in TAG_ORDER):
            for tag in TAG_ORDER:
                n = tag_counts.get(tag, 0)
                if n:
                    _ratio_bar(tag, n, tc_total, "#66bb6a")
        else:
            st.caption("无变更")
    with col_b:
        st.markdown("**警告类型**")
        wc_total = max(sum(warn_type_counts.values()), 1)
        if any(warn_type_counts.get(t) for t in WARN_ORDER):
            for wt in WARN_ORDER:
                n = warn_type_counts.get(wt, 0)
                if n:
                    _ratio_bar(wt, n, wc_total, "#ffa726")
        else:
            st.caption("无警告")

    st.divider()
    st.markdown("**按工作表（修改 / 警告）**")
    all_sheets = list(dict.fromkeys(
        [s for s, *_ in changes] + [s for s, *_ in warn_items]
    ))
    if all_sheets:
        st.markdown(
            '<table style="width:100%;border-collapse:collapse;font-size:0.9rem">'
            '<thead><tr style="background:#f0f2f6">'
            '<th style="text-align:left;padding:6px 10px">工作表</th>'
            '<th style="text-align:center;padding:6px 10px">修改</th>'
            '<th style="text-align:center;padding:6px 10px">警告</th>'
            '</tr></thead><tbody>'
            + "".join(
                f'<tr style="border-top:1px solid #e0e0e0">'
                f'<td style="padding:5px 10px">{_esc(s)}</td>'
                f'<td style="text-align:center;padding:5px 10px;color:#2e7d32">'
                f'<b>{sheet_c.get(s, 0)}</b></td>'
                f'<td style="text-align:center;padding:5px 10px;color:#e65100">'
                f'<b>{sheet_w.get(s, 0)}</b></td></tr>'
                for s in all_sheets
            )
            + "</tbody></table>",
            unsafe_allow_html=True,
        )
    else:
        st.caption("没有任何修改或警告。")


# ── 主界面 ───────────────────────────────────────────────────────────────────

def main() -> None:

    # ── 侧边栏 ────────────────────────────────────────────────
    with st.sidebar:
        st.header("📋 处理规则")
        st.markdown("""
**步骤 1　统一为真实换行**
- 删除不可见字符
  （零宽空格、BOM、行/段分隔符）
- 行尾 `\\r\\n` / `\\r` → 换行
- 字面量 `\\n` → 真实换行

**步骤 2　处理首尾**
- 仅换行 → 自动去除
- 换行＋空格混合 → ⚠ 警告
- 仅空格 → ⚠ 警告

> 处理后表格内所有换行均为真实换行符，
> 字符串内部内容不增删。
""")

    # ── 主区域 ────────────────────────────────────────────────
    st.title("🧹 i18n XLSX 清理工具")
    st.caption("上传翻译文件，自动清理格式问题，下载修复后的版本")

    uploaded = st.file_uploader(
        "将 xlsx 文件拖放到此处，或点击选择",
        type=["xlsx"],
        label_visibility="collapsed",
    )

    if uploaded is None:
        st.session_state.pop("results", None)
        st.info("📂 请上传一个 xlsx 文件开始处理")
        return

    st.success(f"已选择：**{uploaded.name}**（{len(uploaded.getvalue()) / 1024:.1f} KB）")

    if st.button("🚀 开始处理", type="primary", use_container_width=True):
        with st.spinner("处理中，请稍候…"):
            wb = openpyxl.load_workbook(io.BytesIO(uploaded.getvalue()))
            changes, warn_items = process_workbook(wb)
            output = io.BytesIO()
            wb.save(output)

        st.session_state.results = {
            "changes":    changes,
            "warn_items": warn_items,
            "output":     output.getvalue(),
            "filename":   uploaded.name,
        }
        _reset_filters()   # 新结果：重置筛选条

    # ── 结果区域 ───────────────────────────────────────────────
    if "results" not in st.session_state:
        return

    r          = st.session_state.results
    changes    = r["changes"]
    warn_items = r["warn_items"]

    st.download_button(
        label="⬇️ 下载处理结果",
        data=r["output"],
        file_name=Path(r["filename"]).stem + "_sanitized.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()

    # ── Tab 导航（用 radio 替代 st.tabs，保证重跑后不跳回第一项）──
    TAB_SUMMARY  = "📊 总结"
    TAB_CHANGES  = f"✅ 变更  {len(changes)} 处"
    TAB_WARNINGS = f"⚠️ 警告  {len(warn_items)} 处"

    active_tab = st.radio(
        "tab_nav",
        options=[TAB_SUMMARY, TAB_CHANGES, TAB_WARNINGS],
        horizontal=True,
        label_visibility="collapsed",
        key="active_tab",
    )

    st.divider()

    if active_tab == TAB_SUMMARY:
        render_summary(changes, warn_items)
    elif active_tab == TAB_CHANGES:
        if changes:
            render_changes(changes)
        else:
            st.success("无需修改，文件已是规范格式。")
    else:  # TAB_WARNINGS
        if warn_items:
            st.info("以下条目需人工确认，脚本未自动处理。")
            render_warnings(warn_items)
        else:
            st.success("没有需要人工确认的警告。")


if __name__ == "__main__":
    main()
