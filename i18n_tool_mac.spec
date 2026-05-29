# -*- mode: python ; coding: utf-8 -*-
"""
i18n_tool.spec — PyInstaller 打包配置
"""

from PyInstaller.utils.hooks import collect_data_files, collect_all

# ── 收集 streamlit 的所有数据文件（静态资源、模板等）────────────────────────
streamlit_datas, streamlit_binaries, streamlit_hidden = collect_all("streamlit")
openpyxl_datas, _, _ = collect_all("openpyxl")

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=streamlit_binaries,
    datas=[
        # Streamlit 静态资源
        *streamlit_datas,
        # openpyxl 数据（模板、默认样式等）
        *openpyxl_datas,
        # 应用脚本作为数据文件随包一起发布
        ("sanitize_i18n_app.py",  "."),
        ("sanitize_i18n_xlsx.py", "."),
    ],
    hiddenimports=[
        *streamlit_hidden,
        # Streamlit 运行时
        "streamlit.web.cli",
        "streamlit.web.server",
        "streamlit.runtime",
        "streamlit.runtime.scriptrunner",
        "streamlit.components.v1",
        # 依赖
        "openpyxl",
        "openpyxl.cell._writer",
        "tornado",
        "tornado.platform.asyncio",
        "click",
        "packaging",
        "packaging.version",
        "packaging.requirements",
        "importlib_metadata",
        "gitpython",
        "pydeck",
        "validators",
        "rich",
        "tenacity",
        "altair",
        "pandas",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 排除不需要的大包，减小体积
        "matplotlib",
        "scipy",
        "sklearn",
        "notebook",
        "IPython",
        "PIL",
        "cv2",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="i18n清理工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # macOS 上 UPX 兼容性差，关闭
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

app = BUNDLE(
    exe,
    name="i18n清理工具.app",
    icon=None,
    bundle_identifier="com.even.i18n-sanitizer",
    info_plist={
        "CFBundleDisplayName": "i18n清理工具",
        "CFBundleShortVersionString": "1.0.0",
        "NSHighResolutionCapable": True,
        "LSUIElement": False,
    },
)
