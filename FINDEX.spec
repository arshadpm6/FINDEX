# -*- mode: python ; coding: utf-8 -*-
datas = [
    ('resources', 'resources'),
    ('assets', 'assets'),
    ('external', 'external'),
    ('logo.png', '.'),
    ('icon.png', '.'),
    ('process.png', '.'),
    ('app_icon.ico', '.'),
]

hiddenimports = [
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.utils',
    'PIL',
    'PIL.Image',
    'PIL.ImageQt',
    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
    'PySide6.QtSvg',
]

excludes = [
    'numpy',  # Optional to openpyxl; FINDEX exports use builtin Python values.
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
    'jupyter',
    'notebook',
    'pandas',
    'sklearn',
    'scikit-learn',
]

a = Analysis(
    ['FINDEX.py'],
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

# Qt's Windows build uses the OS ICU API. Do not shadow it with a differently
# versioned ICU from unrelated tools (e.g. Poppler) on the developer's PATH.
a.binaries = [entry for entry in a.binaries
              if entry[0].lower() not in {'icuuc.dll', 'icuin.dll'}
              and not (entry[0].lower().startswith('icudt') and entry[0].lower().endswith('.dll'))]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='FINDEX',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources/branding/icon.ico',
    version='resources/branding/version_info.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='FINDEX',
)
