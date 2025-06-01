# -*- mode: python ; coding: utf-8 -*-
# 说明：本文件是使用pyinstaller打包的配置文件
# 使用方法：
# pip install pyinstaller
# pyinstaller word2md.spec

a = Analysis(
    ['gui\\gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('gui/style.py', '.'),
        ('parsers', 'parsers'),
        ('converters', 'converters'),
        ('generators', 'generators'),
        ('processors', 'processors'),
        ('selector', 'selector'),
        ('gui/logo.ico','.'),
    ],
    hiddenimports=['docx'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='word2md',  # 可执行文件名称
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # 没有控制台窗口
    icon='gui/logo.ico'
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,  # 数据文件（包括 styles.py 和其他依赖）
    strip=False,
    upx=True,
    upx_exclude=[],
    name='word2md'  # 输出目录名称
)