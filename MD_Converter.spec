# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['convert_to_md.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas', 'numpy', 'scipy', 'torch', 'torchvision', 'torchaudio', 'tensorflow', 'tensorboard', 'sklearn', 'transformers', 'sentence_transformers', 'IPython', 'jedi', 'parso', 'zmq', 'pyarrow', 'onnxruntime', 'jinja2', 'sympy', 'psutil', 'anyio', 'orjson', 'rich', 'pygments', 'wcwidth', 'fsspec', 'bcrypt', 'win32com', 'pythoncom', 'pywintypes', 'tkinter', 'matplotlib', 'gi', 'nltk', 'spacy', 'gensim', 'pydantic', 'pyarrow'],
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
    name='MD_Converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
