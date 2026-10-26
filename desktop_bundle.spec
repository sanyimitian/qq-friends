# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules

datas, binaries, hiddenimports = [], [], []
for package in ('rapidocr_onnxruntime', 'onnxruntime', 'pywinauto', 'comtypes'):
    package_data, package_binaries, package_imports = collect_all(package)
    datas += package_data
    binaries += package_binaries
    hiddenimports += package_imports
hiddenimports += ['zq_qq', 'open_qq_add_friend', 'task_control', 'gui_worker', 'gui_config',
                  'prepare_numbers', 'win32timezone', 'win32job', 'win32event', 'cv2', 'numpy']
hiddenimports += collect_submodules('odf')
hiddenimports += collect_submodules('openpyxl')

a = Analysis(['capture_app.py', 'add_app.py'], pathex=[SPECPATH], binaries=binaries,
             datas=datas, hiddenimports=hiddenimports, hookspath=[], hooksconfig={},
             runtime_hooks=[], excludes=['pytest', 'IPython', 'matplotlib', 'torch', 'tensorflow'],
             noarchive=False, optimize=0)
pyz = PYZ(a.pure)
support = [item for item in a.scripts if item[0] not in ('capture_app', 'add_app')]
capture_script = next(item for item in a.scripts if item[0] == 'capture_app')
add_script = next(item for item in a.scripts if item[0] == 'add_app')
capture = EXE(pyz, support + [capture_script], [], exclude_binaries=True,
              name='QQ群成员采集', debug=False, bootloader_ignore_signals=False,
              strip=False, upx=False, console=False, disable_windowed_traceback=False)
adding = EXE(pyz, support + [add_script], [], exclude_binaries=True,
             name='QQ好友添加', debug=False, bootloader_ignore_signals=False,
             strip=False, upx=False, console=False, disable_windowed_traceback=False)
bundle = COLLECT(capture, adding, a.binaries, a.datas, strip=False, upx=False,
                 name='QQ桌面工具')
