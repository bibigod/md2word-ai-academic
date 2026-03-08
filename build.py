"""
PyInstaller 打包脚本
用法: python build.py
"""
import subprocess
import sys
import os

ROOT = os.path.dirname(os.path.abspath(__file__))

# 数据文件：(源路径, 打包后相对路径)
datas = [
    (os.path.join(ROOT, 'theme_breeze.json'), '.'),
    (os.path.join(ROOT, 'sample.md'), '.'),
    (os.path.join(ROOT, 'md2docx.yaml'), '.'),
]

# 隐式导入（PyInstaller 可能扫描不到的模块）
hidden_imports = [
    'yaml',
    'docx',
    'PIL',
    'fitz',
    'win32com',
    'win32com.client',
    'pythoncom',
    'customtkinter',
]

# 排除不需要的大型包（PyInstaller 可能通过依赖链错误拉入）
excludes = [
    'torch', 'tensorflow', 'scipy', 'pandas', 'matplotlib',
    'numpy.testing', 'pytest', 'IPython', 'notebook', 'jupyter',
    'openpyxl', 'xlsxwriter', 'h5py', 'sympy', 'sklearn',
    'cv2', 'boto3', 'botocore', 'setuptools', 'pkg_resources',
]

cmd = [
    sys.executable, '-m', 'PyInstaller',
    '--noconfirm',
    '--onedir',
    '--windowed',
    '--name', 'Md2word-ai-academic',
    '--clean',
]

for mod in excludes:
    cmd.extend(['--exclude-module', mod])

for src, dst in datas:
    cmd.extend(['--add-data', f'{src}{os.pathsep}{dst}'])

for mod in hidden_imports:
    cmd.extend(['--hidden-import', mod])

# customtkinter 需要整个包数据
try:
    import customtkinter
    ctk_path = os.path.dirname(customtkinter.__file__)
    cmd.extend(['--add-data', f'{ctk_path}{os.pathsep}customtkinter'])
except ImportError:
    pass

cmd.append(os.path.join(ROOT, 'main.py'))

print('执行命令:')
print(' '.join(cmd))
print()

subprocess.run(cmd, cwd=ROOT)
