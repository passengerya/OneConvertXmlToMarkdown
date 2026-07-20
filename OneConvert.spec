# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = [('E:\\Code\\OneConvertXmlToMarkdown\\convert_onenote_xml.py', '.'), ('E:\\Code\\OneConvertXmlToMarkdown\\Convert-OneNoteSectionToXml.ps1', '.'), ('E:\\Code\\OneConvertXmlToMarkdown\\Convert-OneNoteToMarkdownPipeline.ps1', '.')]
datas += collect_data_files('flet')
datas += collect_data_files('flet_core')
datas += collect_data_files('flet_desktop')


a = Analysis(
    ['E:\\Code\\OneConvertXmlToMarkdown\\OneConvertGUI.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name='OneConvert',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
