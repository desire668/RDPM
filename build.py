import os
import glob
import shutil
import subprocess
import sys

def clean_build():
    """清理构建文件夹"""
    dirs_to_clean = ['build', 'dist']
    files_to_clean = ['*.spec']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            
    for pattern in files_to_clean:
        for file in glob.glob(pattern):
            os.remove(file)

def build_exe():
    """构建可执行文件"""
    print("开始构建可执行文件...")
    
    # PyInstaller命令行参数
    pyinstaller_args = [
        'pyinstaller',
        '--noconfirm',
        '--onefile',
        '--windowed',
        '--icon=app.ico',  # 如果有图标的话
        '--name=RDP管理器',
        '--add-data=README.md;.',
        '--hidden-import=win32com.shell.shell',
        '--hidden-import=win32api',
        '--hidden-import=win32con',
        '--hidden-import=win32security',
        'rdp_gui.py'
    ]
    
    # 运行PyInstaller
    try:
        subprocess.run(pyinstaller_args, check=True)
        print("\n构建成功！可执行文件位于 dist 文件夹中。")
    except subprocess.CalledProcessError as e:
        print(f"\n构建失败：{str(e)}")
        sys.exit(1)

def main():
    # 清理旧的构建文件
    clean_build()
    
    # 检查是否已安装所需依赖
    print("检查依赖...")
    subprocess.run([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'])
    
    # 构建可执行文件
    build_exe()
    
    print("\n打包完成！")
    print("你可以在 dist 文件夹中找到打包好的程序。")
    print("提示：")
    print("1. 程序需要管理员权限才能运行")
    print("2. 第一次运行可能会被杀毒软件拦截，需要添加信任")
    print("3. 配置文件会保存在用户目录的 .rdp_manager 文件夹中")

if __name__ == '__main__':
    main() 