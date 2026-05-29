"""
sanitize_i18n_app.py
i18n XLSX 清理工具 — Streamlit 图形界面

启动方式：
  python3 -m streamlit run sanitize_i18n_app.py
"""

import difflib
import hashlib
import io
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
import streamlit as st

# 从同目录的脚本文件导入核心逻辑和常量
sys.path.insert(0, str(Path(__file__).parent))
from sanitize_i18n_xlsx import (
    TAG_ESCAPE, TAG_INVISIBLE, TAG_NEWLINE,
    WARN_MIXED, WARN_SPACE,
    INVISIBLE_CHARS,
    fmt, process_workbook,
)


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
.change-row { background:#f0faf0; border-left:3px solid #2e7d32;
              padding:10px 14px; border-radius:4px; margin-bottom:8px; }
.warn-row   { background:#fffde7; border-left:3px solid #f9a825;
              padding:10px 14px; border-radius:4px; margin-bottom:8px; }
.label      { font-size:0.78rem; color:#666; margin-bottom:2px; }
.mono       { font-family:monospace; font-size:0.88rem; word-break:break-all; }
/* 真实换行符：蓝色 ↵ */
.real-nl    { color:#1565c0; font-weight:bold; }
/* 字面量 \n：橙色 */
.lit-n      { color:#e65100; }
/* 不可见字符：灰色 */
.invis-char { color:#9e9e9e; }
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

def _fmt_html(val: str) -> str:
    """
    将值转为 HTML 安全的可读字符串，不加外层引号。
    用不同颜色区分三种特殊字符：
      · 真实换行符 (0x0A)  → 蓝色 ↵
      · 字面量 \\n（反斜杠+n） → 橙色 \\n
      · 不可见字符          → 灰色 \\uXXXX
    """
    # 按真实换行符切分，保留每段内容
    # 这样可以在 HTML escape 之后安全地插入换行标记
    segments = val.split("\n")

    processed = []
    for seg in segments:
        # 1. HTML 转义（& < > 必须最先处理）
        s = seg.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        # 2. 不可见字符 → 灰色 \uXXXX
        for ch in INVISIBLE_CHARS:
            s = s.replace(ch, f'<span class="invis-char">\\u{ord(ch):04x}</span>')
        # 3. 字面量 \n（反斜杠+n）→ 橙色 \n
        #    此时字符串里 \ 已经是普通字符，直接替换两字符序列 \n
        s = s.replace("\\n", '<span class="lit-n">\\n</span>')
        processed.append(s)

    # 4. 用蓝色 ↵ 连接各段，代表原始的真实换行符
    return '<span class="real-nl">↵</span>'.join(processed)


def _diff_html(old: str, new: str) -> tuple[str, str]:
    """
    逐字符比较 old 和 new，返回 (old_html, new_html)。
    相同部分正常渲染，差异部分加背景色高亮：
      · 删除/替换的旧内容 → 浅红背景
      · 新增/替换的新内容 → 浅绿背景
    文字颜色统一使用默认深色，不再区分红/绿文字。
    """
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
            old_parts.append(
                f'<mark style="background:#ffcdd2;border-radius:2px">{old_seg}</mark>'
            )
            new_parts.append(
                f'<mark style="background:#c8e6c9;border-radius:2px">{new_seg}</mark>'
            )
        elif op == "delete":
            old_parts.append(
                f'<mark style="background:#ffcdd2;border-radius:2px">{old_seg}</mark>'
            )
        elif op == "insert":
            new_parts.append(
                f'<mark style="background:#c8e6c9;border-radius:2px">{new_seg}</mark>'
            )

    return "".join(old_parts), "".join(new_parts)


# ── session state 键管理 ─────────────────────────────────────────────────────

def _safe(s: str) -> str:
    """将字符串转为合法的 session state 键片段。
    保留 ASCII 字母数字，其余字符替换为下划线，
    再附加 6 位 MD5 hash 保证不同原始字符串生成唯一键。"""
    ascii_part = re.sub(r"[^a-zA-Z0-9]", "_", str(s))
    hash_part  = hashlib.md5(str(s).encode()).hexdigest()[:6]
    return f"{ascii_part}_{hash_part}"


def _ek(*parts) -> str:
    """生成 expander 的 session state 键。"""
    return "exp__" + "__".join(_safe(p) for p in parts)


def _ev(*parts) -> bool:
    """读取 expander 展开状态，未设置时回落到全局默认。"""
    return st.session_state.get(_ek(*parts),
           st.session_state.get("ui_expanded", False))


def _set_children(value: bool, *parent_parts, key_dict: dict) -> None:
    """
    将 parent_parts 对应节点的所有子 expander 设为 value。
    key_dict 是该节点下一层的 {name: children} 字典。
    同时保持父节点自身为展开（按钮在父节点内部，点击时父节点必然是打开的）。
    """
    # 父节点保持展开
    st.session_state[_ek(*parent_parts)] = True
    # 遍历子节点并向下递归
    for child_name, grandchildren in key_dict.items():
        child_parts = (*parent_parts, child_name)
        st.session_state[_ek(*child_parts)] = value
        if isinstance(grandchildren, dict):
            for gc_name in grandchildren:
                st.session_state[_ek(*child_parts, gc_name)] = value


def _clear_all_exp_keys() -> None:
    """删除所有 exp__ 前缀的 session state，让 expander 回落到全局状态。"""
    for k in list(st.session_state.keys()):
        if k.startswith("exp__"):
            del st.session_state[k]


def _expand_collapse_buttons(btn_key: str, *parent_parts, key_dict: dict) -> None:
    """在当前位置渲染一对展开/折叠子项按钮。
    按钮点击本身已触发 Streamlit 重跑，session state 在同一次重跑里生效，
    无需额外调用 st.rerun()，避免二次刷新导致 Tab 重置。"""
    c1, c2, _ = st.columns([0.8, 0.8, 6])
    if c1.button("⊞ 展开", key=f"btn_exp_{btn_key}", use_container_width=True):
        _set_children(True, *parent_parts, key_dict=key_dict)
    if c2.button("⊟ 折叠", key=f"btn_col_{btn_key}", use_container_width=True):
        _set_children(False, *parent_parts, key_dict=key_dict)


# ── 数据层级构建 ─────────────────────────────────────────────────────────────

def _build_change_hierarchy(changes: list) -> dict:
    """tag → sheet → key → [(row, col, old, new)]"""
    h: dict = {}
    for sheet, row, key, col, old, new, tags in changes:
        for tag in sorted(tags):
            h.setdefault(tag, {})
            h[tag].setdefault(sheet, {})
            h[tag][sheet].setdefault(key, [])
            h[tag][sheet][key].append((row, col, old, new))
    return h


def _build_warn_hierarchy(warn_items: list) -> dict:
    """warn_type → sheet → key → [(row, col, val, [details])]"""
    h: dict = {}
    for sheet, row, key, col, val, msgs in warn_items:
        type_msgs: dict[str, list] = defaultdict(list)
        for wt, detail in msgs:
            type_msgs[wt].append(detail)
        for wt, details in type_msgs.items():
            h.setdefault(wt, {})
            h[wt].setdefault(sheet, {})
            h[wt][sheet].setdefault(key, [])
            h[wt][sheet][key].append((row, col, val, details))
    return h


# ── 渲染函数 ─────────────────────────────────────────────────────────────────

def render_changes(changes: list) -> None:
    """三层结构：变更类型 → 工作表 → key → 条目"""
    h = _build_change_hierarchy(changes)

    for tag, sheet_dict in h.items():
        tag_total = sum(len(k) for s in sheet_dict.values() for k in s.values())

        # ── 第一层：变更类型 ────────────────────────────────────
        with st.expander(f"🏷 {tag}　{tag_total} 处", expanded=_ev("c", tag)):

            # 类型级展开/折叠按钮（控制下面的工作表和 key）
            _expand_collapse_buttons(f"c__{_safe(tag)}", "c", tag,
                                     key_dict=sheet_dict)

            for sheet, key_dict in sheet_dict.items():
                sheet_total = sum(len(v) for v in key_dict.values())

                # ── 第二层：工作表 ──────────────────────────────
                with st.expander(f"📄 {sheet}　{sheet_total} 处",
                                 expanded=_ev("c", tag, sheet)):

                    # 工作表级展开/折叠按钮（控制下面的 key）
                    _expand_collapse_buttons(f"c__{_safe(tag)}__{_safe(sheet)}",
                                             "c", tag, sheet,
                                             key_dict=key_dict)

                    for key_name, entries in key_dict.items():

                        # ── 第三层：key ─────────────────────────
                        with st.expander(f"🔑 key = {key_name}　{len(entries)} 处",
                                         expanded=_ev("c", tag, sheet, key_name)):
                            for row, col, old, new in entries:
                                old_d, new_d = _diff_html(old, new)
                                st.markdown(
                                    f'<div class="change-row">'
                                    f'<div class="label">第 {row} 行 · [{col}]</div>'
                                    f'<div class="label" style="margin-top:4px">修改前</div>'
                                    f'<div class="mono">\'{old_d}\'</div>'
                                    f'<div class="label" style="margin-top:6px">修改后</div>'
                                    f'<div class="mono">\'{new_d}\'</div>'
                                    f'</div>',
                                    unsafe_allow_html=True,
                                )


def render_warnings(warn_items: list) -> None:
    """三层结构：警告类型 → 工作表 → key → 条目"""
    h = _build_warn_hierarchy(warn_items)

    for warn_type, sheet_dict in h.items():
        type_total = sum(len(k) for s in sheet_dict.values() for k in s.values())

        # ── 第一层：警告类型 ────────────────────────────────────
        with st.expander(f"🏷 {warn_type}　{type_total} 处",
                         expanded=_ev("w", warn_type)):

            # 类型级展开/折叠按钮
            _expand_collapse_buttons(f"w__{_safe(warn_type)}", "w", warn_type,
                                     key_dict=sheet_dict)

            for sheet, key_dict in sheet_dict.items():
                sheet_total = sum(len(v) for v in key_dict.values())

                # ── 第二层：工作表 ──────────────────────────────
                with st.expander(f"📄 {sheet}　{sheet_total} 处",
                                 expanded=_ev("w", warn_type, sheet)):

                    # 工作表级展开/折叠按钮
                    _expand_collapse_buttons(
                        f"w__{_safe(warn_type)}__{_safe(sheet)}",
                        "w", warn_type, sheet,
                        key_dict=key_dict,
                    )

                    for key_name, entries in key_dict.items():

                        # ── 第三层：key ─────────────────────────
                        with st.expander(f"🔑 key = {key_name}　{len(entries)} 处",
                                         expanded=_ev("w", warn_type, sheet, key_name)):
                            for row, col, val, details in entries:
                                val_d = _fmt_html(val)
                                details_html = "".join(
                                    f'<div class="mono" style="color:#e65100">'
                                    f'{d.replace("<","&lt;")}</div>'
                                    for d in details
                                )
                                st.markdown(
                                    f'<div class="warn-row">'
                                    f'<div class="label">第 {row} 行 · [{col}]</div>'
                                    f'<div class="label" style="margin-top:4px">当前值</div>'
                                    f'<div class="mono">\'{val_d}\'</div>'
                                    f'<div style="margin-top:6px">{details_html}</div>'
                                    f'</div>',
                                    unsafe_allow_html=True,
                                )


def render_summary(changes: list, warn_items: list, escape: bool) -> None:
    """处理总结 Tab。"""
    tag_counts:       dict[str, int] = defaultdict(int)
    warn_type_counts: dict[str, int] = defaultdict(int)
    sheet_c:          dict[str, int] = defaultdict(int)
    sheet_w:          dict[str, int] = defaultdict(int)

    for _, _, _, _, _, _, tags in changes:
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
        for tag in [TAG_INVISIBLE, TAG_NEWLINE] + ([TAG_ESCAPE] if escape else []):
            st.markdown(f"- {tag}：**{tag_counts.get(tag, 0)}** 个")
    with col_b:
        st.markdown("**警告类型**")
        for wt in (WARN_MIXED, WARN_SPACE):
            count = warn_type_counts.get(wt, 0)
            if count:
                st.markdown(f"- {wt}：**{count}** 个")

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
                f'<td style="padding:5px 10px">{s}</td>'
                f'<td style="text-align:center;padding:5px 10px;color:#2e7d32">'
                f'<b>{sheet_c.get(s,0)}</b></td>'
                f'<td style="text-align:center;padding:5px 10px;color:#e65100">'
                f'<b>{sheet_w.get(s,0)}</b></td></tr>'
                for s in all_sheets
            )
            + "</tbody></table>",
            unsafe_allow_html=True,
        )


# ── 主界面 ───────────────────────────────────────────────────────────────────

def main() -> None:

    # ── 侧边栏 ────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ 处理选项")
        escape = st.checkbox(
            "将行内换行符转为字面量 `\\n`",
            help=(
                "启用后，字符串内部的真实换行符会被替换为两个字符 \\n，"
                "表格内不再存在真实换行符。\n\n"
                "注意：若首尾存在换行与空格混合的问题，"
                "该单元格会跳过此步骤并归入警告。"
            ),
        )
        st.divider()
        st.markdown("""
**处理规则**

1. 删除不可见字符
   零宽空格、BOM、行分隔符

2. 首尾换行符
   · 仅换行 → 自动去除
   · 换行＋空格混合 → 警告
   · 仅空格 → 警告

3. 行内内容不修改

4. *(可选)* 行内换行转义
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
            changes, warn_items = process_workbook(wb, escape)
            output = io.BytesIO()
            wb.save(output)

        st.session_state.results = {
            "changes":    changes,
            "warn_items": warn_items,
            "output":     output.getvalue(),
            "filename":   uploaded.name,
            "escape":     escape,
        }
        # 新结果：清除所有 expander 状态，重置为全部折叠
        _clear_all_exp_keys()
        st.session_state["ui_expanded"] = False

    # ── 结果区域 ───────────────────────────────────────────────
    if "results" not in st.session_state:
        return

    r          = st.session_state.results
    changes    = r["changes"]
    warn_items = r["warn_items"]
    escape     = r["escape"]

    # 下载按钮
    st.download_button(
        label="⬇️ 下载处理结果",
        data=r["output"],
        file_name=Path(r["filename"]).stem + "_sanitized.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()

    # ── Tab 导航（用 radio 替代 st.tabs，保证重跑后不跳回第一项）──
    TAB_SUMMARY  = f"📊 总结"
    TAB_CHANGES  = f"✅ 变更  {len(changes)} 处"
    TAB_WARNINGS = f"⚠️ 警告  {len(warn_items)} 处"

    nav_col, exp_col = st.columns([3, 2])

    with nav_col:
        active_tab = st.radio(
            "tab_nav",
            options=[TAB_SUMMARY, TAB_CHANGES, TAB_WARNINGS],
            horizontal=True,
            label_visibility="collapsed",
            key="active_tab",
        )

    # 全局展开/折叠按钮（仅在变更/警告页时有意义）
    if active_tab != TAB_SUMMARY:
        with exp_col:
            gc1, gc2, _ = st.columns([1, 1, 2])
            if gc1.button("⊞ 全部展开", use_container_width=True):
                _clear_all_exp_keys()
                st.session_state["ui_expanded"] = True
            if gc2.button("⊟ 全部折叠", use_container_width=True):
                _clear_all_exp_keys()
                st.session_state["ui_expanded"] = False

    st.divider()

    # ── 页面内容 ───────────────────────────────────────────────
    if active_tab == TAB_SUMMARY:
        render_summary(changes, warn_items, escape)

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
