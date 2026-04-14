# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — LicitaMonitor Soldesp
# Compilar con: pyinstaller LicitaMonitor.spec --noconfirm

from PyInstaller.utils.hooks import collect_all
import os as _os

datas_sel,     binaries_sel,     hiddenimports_sel     = collect_all('selenium')
datas_wdm,     binaries_wdm,     hiddenimports_wdm     = collect_all('webdriver_manager')
datas_opxl,    binaries_opxl,    hiddenimports_opxl    = collect_all('openpyxl')
datas_certifi, binaries_certifi, hiddenimports_certifi = collect_all('certifi')
datas_urllib3, binaries_urllib3, hiddenimports_urllib3 = collect_all('urllib3')

try:
    datas_pil, binaries_pil, hiddenimports_pil = collect_all('PIL')
except Exception:
    datas_pil, binaries_pil, hiddenimports_pil = [], [], []

_logo = [('logo.png', '.')] if _os.path.exists('logo.png') else []

a = Analysis(
    ['LicitaMonitor_GUI.py'],
    pathex=[],
    binaries=(binaries_sel + binaries_wdm + binaries_opxl
              + binaries_certifi + binaries_urllib3 + binaries_pil),
    datas=(
        datas_sel + datas_wdm + datas_opxl + datas_certifi + datas_urllib3
        + datas_pil + _logo
        # config.json NO se incluye aquí — debe estar junto al .exe
    ),
    hiddenimports=(
        hiddenimports_sel + hiddenimports_wdm + hiddenimports_opxl
        + hiddenimports_certifi + hiddenimports_urllib3 + hiddenimports_pil
        + [
            'tkinter', 'tkinter.ttk', 'tkinter.filedialog',
            'tkinter.messagebox', 'tkinter.scrolledtext', 'tkinter.font',
            '_tkinter',
            'email.mime.multipart', 'email.mime.text',
            'email.mime.base', 'email.encoders', 'smtplib',
            'logging.handlers', 'pathlib', 'json', 'queue',
            'threading', 'subprocess', 'unicodedata',
            'requests', 'requests.adapters', 'requests.auth',
            'charset_normalizer',
        ]
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy', 'PIL.ImageTk',
        'IPython', 'notebook', 'pytest', 'pydoc', 'xmlrpc',
        'ftplib', 'imaplib', 'poplib', 'turtle', 'turtledemo',
    ],
    noarchive=False,
    optimize=2,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='LicitaMonitor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[
        'vcruntime140.dll', 'msvcp140.dll', 'ucrtbase.dll',
        'api-ms-win-*.dll', '_ssl.pyd', '_hashlib.pyd',
        'libssl*.dll', 'libcrypto*.dll',
        'tcl*.dll', 'tk*.dll', 'python3*.dll',
    ],
    runtime_tmpdir=None,
    console=False,
    icon=None,              # Cambiar a 'icon.ico' si existe
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    onefile=True,
    uac_admin=False,
)
