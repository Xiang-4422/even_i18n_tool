# -*- mode: python ; coding: utf-8 -*-
"""
i18n_tool_win.spec — Windows 打包配置
在 Windows 机器上运行：pyinstaller i18n_tool_win.spec
"""

from PyInstaller.utils.hooks import collect_all

streamlit_datas, streamlit_binaries, streamlit_hidden = collect_all("streamlit")
openpyxl_datas, _, _ = collect_all("openpyxl")

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=streamlit_binaries,
    datas=[
        *streamlit_datas,
        *openpyxl_datas,
        ("sanitize_i18n_app.py",  "."),
        ("sanitize_i18n_xlsx.py", "."),
    ],
    hiddenimports=[
        *streamlit_hidden,
        "streamlit.web.cli",
        "streamlit.web.server",
        "streamlit.runtime",
        "streamlit.runtime.scriptrunner",
        "streamlit.components.v1",
        "openpyxl",
        "openpyxl.cell._writer",
        "tornado",
        "tornado.platform.asyncio",
        "click",
        "packaging",
        "packaging.version",
        "packaging.requirements",
        "importlib_metadata",
        "validators",
        "rich",
        "tenacity",
        "altair",
        "pandas",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["matplotlib", "scipy", "sklearn", "notebook", "IPython", "PIL", "cv2"],
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
    strip=False,
    upx=False,
    console=False,       # 不弹出命令行窗口
    icon=None,
)
