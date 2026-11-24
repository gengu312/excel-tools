#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具集合 - 打包脚本
使用PyInstaller将程序打包成exe文件
"""

import os
import subprocess
import sys
from pathlib import Path

def check_pyinstaller():
    """检查PyInstaller是否已安装"""
    try:
        import PyInstaller
        print("✓ PyInstaller已安装")
        return True
    except ImportError:
        print("✗ PyInstaller未安装")
        return False

def install_pyinstaller():
    """安装PyInstaller"""
    print("正在安装PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller安装成功")
        return True
    except subprocess.CalledProcessError:
        print("✗ PyInstaller安装失败")
        return False

def create_spec_file():
    """创建PyInstaller规格文件"""
    spec_content = '''
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel_tools_launcher.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'xlrd',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel工具集合',
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
    icon=None,
)
'''
    
    with open('excel_tools.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    print("✓ 规格文件已创建")

def build_exe():
    """构建exe文件"""
    print("开始构建exe文件...")
    try:
        # 使用规格文件构建
        subprocess.check_call(["pyinstaller", "excel_tools.spec", "--clean"])
        print("✓ exe文件构建成功")
        
        # 检查输出文件
        exe_path = Path("dist/Excel工具集合.exe")
        if exe_path.exists():
            print(f"✓ exe文件位置: {exe_path.absolute()}")
            print(f"✓ 文件大小: {exe_path.stat().st_size / 1024 / 1024:.1f} MB")
        else:
            print("✗ 未找到生成的exe文件")
            
    except subprocess.CalledProcessError as e:
        print(f"✗ 构建失败: {e}")
        return False
    
    return True

def clean_build_files():
    """清理构建文件"""
    import shutil
    
    dirs_to_remove = ['build', '__pycache__']
    files_to_remove = ['excel_tools.spec']
    
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"✓ 已清理: {dir_name}")
    
    for file_name in files_to_remove:
        if os.path.exists(file_name):
            os.remove(file_name)
            print(f"✓ 已清理: {file_name}")

def main():
    """主函数"""
    print("=" * 50)
    print("Excel工具集合 - 打包脚本")
    print("=" * 50)
    
    # 检查当前目录
    if not os.path.exists('excel_tools_launcher.py'):
        print("✗ 未找到excel_tools_launcher.py文件")
        print("请确保在正确的目录中运行此脚本")
        return
    
    # 检查并安装PyInstaller
    if not check_pyinstaller():
        if not install_pyinstaller():
            print("请手动安装PyInstaller: pip install pyinstaller")
            return
    
    # 创建规格文件
    create_spec_file()
    
    # 构建exe
    if build_exe():
        print("\n" + "=" * 50)
        print("✓ 打包完成！")
        print("exe文件位置: dist/Excel工具集合.exe")
        print("=" * 50)
        
        # 询问是否清理构建文件
        response = input("\n是否清理构建文件？(y/n): ")
        if response.lower() in ['y', 'yes', '是']:
            clean_build_files()
    else:
        print("\n" + "=" * 50)
        print("✗ 打包失败")
        print("=" * 50)

if __name__ == "__main__":
    main()