# RDP管理器

一个用于管理Windows远程桌面连接的图形界面工具。
![img](https://github.com/desire668/RDPM/raw/main/test.png)

## 功能特点

- 快速启用/禁用远程桌面
- 自定义远程桌面端口
- 管理多个远程桌面连接
- 批量连接功能
- 安全的密码管理
- 自动配置防火墙规则

## 使用方法

### 打包好的程序

1. 下载 `RDP管理器.exe`
2. 右键以管理员身份运行
3. 在界面中添加和管理远程桌面连接

注意事项：
- 程序需要管理员权限才能运行
- 第一次运行时可能会被杀毒软件拦截，需要添加信任
- 配置文件保存在用户目录的 `.rdp_manager` 文件夹中

### 开发环境

如果你想从源代码运行或打包程序：

1. 确保安装了 Python 3.8 或更高版本
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

3. 运行程序：
   ```bash
   python rdp_gui.py
   ```

4. 打包程序：
   ```bash
   python build.py
   ```
   打包后的程序在 `dist` 文件夹中。

## 技术细节

- 使用 PyQt6 构建图形界面
- 使用 Windows Registry 管理远程桌面设置
- 使用 Fernet 加密保存敏感信息
- 自动管理 Windows 防火墙规则

## 注意事项

1. 修改远程桌面设置需要管理员权限
2. 请确保远程桌面端口没有被其他程序占用
3. 如果使用非默认端口，请确保目标计算机的防火墙允许该端口
4. 建议定期备份 `.rdp_manager` 文件夹中的配置文件 
