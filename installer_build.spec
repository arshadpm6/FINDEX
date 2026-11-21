# -*- mode: python ; coding: utf-8 -*-
datas = [
    ('resources', 'resources'),
    ('dist/app_payload.zip', '.'),
    ('dist/app_payload.manifest.json', '.'),
    ('app_icon.ico', '.'),
]

hiddenimports = [
    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
]

excludes = [
    'torch',
    'torchaudio',
    'torchvision',
    'tensorflow',
    'keras',
    'tensorboard',
    'scipy',
    'spacy',
    'transformers',
    'matplotlib',
    'nltk',
    'TTS',
    'soundfile',
    'pydub',
    'cv2',
    'Cython',
    'IPython',
    'pandas',
    'openpyxl',
    'PIL',
]

a = Analysis(
    ['installer_wizard.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    noarchive=False,
)

# Use Windows ICU rather than accidentally bundling a foreign PATH copy.
a.binaries = [entry for entry in a.binaries
              if entry[0].lower() not in {'icuuc.dll', 'icuin.dll'}
              and not (entry[0].lower().startswith('icudt') and entry[0].lower().endswith('.dll'))]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FINDEX_v3.0.1_Setup',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources/branding/icon.ico',
    version='resources/branding/installer_version_info.txt',
)
