#!/usr/bin/env python3
"""
Windows远程桌面批量管理工具
用于快速管理和连接Windows远程桌面
"""

import os
import sys
import json
import click
import winreg
import subprocess
from pathlib import Path
from typing import Dict, Optional
from rich.console import Console
from rich.table import Table
from cryptography.fernet import Fernet
from win32com.shell import shell, shellcon
import win32security
import win32api
import win32con
import win32process
import win32event

console = Console()
DEFAULT_PORT = 3389
DEFAULT_USERNAME = "administrator"

# 添加进程创建标志
CREATE_NO_WINDOW = 0x08000000

class RDPManager:
    """远程桌面管理器类"""
    
    def __init__(self):
        self.config_dir = Path.home() / '.rdp_manager'
        self.config_file = self.config_dir / 'config.json'
        self.key_file = self.config_dir / '.key'
        self._init_config()
        
    def _init_config(self) -> None:
        """初始化配置目录和文件"""
        self.config_dir.mkdir(exist_ok=True)
        if not self.key_file.exists():
            key = Fernet.generate_key()
            self.key_file.write_bytes(key)
        if not self.config_file.exists():
            self.config_file.write_text('{}')
            
    def _get_cipher(self) -> Fernet:
        """获取加密器"""
        key = self.key_file.read_bytes()
        return Fernet(key)
        
    def _is_admin(self) -> bool:
        """检查是否具有管理员权限"""
        try:
            return shell.IsUserAnAdmin()
        except:
            return False
            
    def _require_admin(self) -> None:
        """要求管理员权限"""
        if not self._is_admin():
            console.print("[red]此操作需要管理员权限！[/red]")
            if sys.platform == 'win32':
                script = os.path.abspath(sys.argv[0])
                params = ' '.join([script] + sys.argv[1:])
                shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
            sys.exit(0)

    def _run_command(self, cmd, check=True, capture_output=True):
        """运行命令并隐藏控制台窗口"""
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        
        try:
            return subprocess.run(
                cmd,
                check=check,
                capture_output=capture_output,
                text=True,
                creationflags=CREATE_NO_WINDOW,
                startupinfo=startupinfo
            )
        except subprocess.CalledProcessError as e:
            # 特殊处理服务已经启动的情况
            if cmd[0] == 'net' and cmd[1] == 'start' and "服务已经启动" in (e.stderr or ""):
                return subprocess.CompletedProcess(cmd, 0, "", "")
            elif check:
                raise
            return subprocess.CompletedProcess(cmd, e.returncode, e.stdout, e.stderr)
        
    def _wait_for_service_status(self, desired_status, timeout=30):
        """等待服务达到期望状态"""
        import time
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                result = self._run_command(['sc', 'query', 'TermService'])
                current_status = result.stdout
                if desired_status == "STOPPED" and "STOPPED" in current_status:
                    return True
                if desired_status == "RUNNING" and "RUNNING" in current_status:
                    return True
                time.sleep(1)
            except:
                pass
        return False

    def change_rdp_port(self, port: int = DEFAULT_PORT) -> None:
        """修改远程桌面端口"""
        import time
        self._require_admin()
        try:
            # 1. 检查服务当前状态
            status_check = self._run_command(['sc', 'query', 'TermService'])
            initial_status = "RUNNING" in status_check.stdout
            
            # 2. 停止服务
            if initial_status:
                try:
                    self._run_command(['net', 'stop', 'TermService', '/y'])
                except subprocess.CalledProcessError:
                    pass  # 忽略停止服务的错误，继续执行
                if not self._wait_for_service_status("STOPPED"):
                    console.print("[yellow]警告：无法完全停止服务，将继续尝试修改端口...[/yellow]")
            
            # 3. 修改注册表中的端口设置
            reg_path = r"SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, 
                              winreg.KEY_ALL_ACCESS) as key:
                winreg.SetValueEx(key, "PortNumber", 0, winreg.REG_DWORD, port)
            
            # 4. 配置防火墙规则
            try:
                # 删除现有规则
                self._run_command(['netsh', 'advfirewall', 'firewall', 'delete', 'rule',
                             'name="RDP Manager"'], check=False)
                self._run_command(['netsh', 'advfirewall', 'firewall', 'delete', 'rule',
                             f'name="RDP Port {DEFAULT_PORT}"'], check=False)
                self._run_command(['netsh', 'advfirewall', 'firewall', 'delete', 'rule',
                             f'name="RDP Port {port}"'], check=False)
            except:
                pass

            # 添加新的入站规则
            self._run_command([
                'netsh', 'advfirewall', 'firewall', 'add', 'rule',
                f'name=RDP Port {port}',
                'dir=in',
                'action=allow',
                'protocol=TCP',
                f'localport={port}',
                'program=%SystemRoot%\\system32\\svchost.exe',
                'service=TermService',
                'description=Remote Desktop Port'
            ])
            
            # 5. 确保服务配置正确
            self._run_command(['sc', 'config', 'TermService', 'start=auto'])
            
            # 6. 启动服务
            try:
                self._run_command(['net', 'start', 'TermService'])
            except subprocess.CalledProcessError:
                pass  # 忽略启动错误，检查实际状态
                
            # 等待一会儿让服务有时间启动
            time.sleep(2)
            
            # 7. 验证最终状态
            final_status = self._run_command(['sc', 'query', 'TermService'])
            if "RUNNING" in final_status.stdout:
                console.print("[green]远程桌面端口修改成功！[/green]")
            else:
                console.print("[yellow]警告：服务可能未正常启动，但端口已经修改。请手动检查服务状态。[/yellow]")
            
        except Exception as e:
            error_msg = str(e)
            console.print(f"[red]修改端口时出错：{error_msg}[/red]")
            raise

    def enable_rdp(self, port: int = DEFAULT_PORT) -> None:
        """启用远程桌面"""
        import time
        self._require_admin()
        try:
            # 1. 先停止远程桌面服务
            try:
                self._run_command(['net', 'stop', 'TermService', '/y'], check=False)
            except:
                pass  # 忽略停止服务的错误
            
            # 2. 修改注册表启用远程桌面
            reg_path = r"System\CurrentControlSet\Control\Terminal Server"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, 
                              winreg.KEY_ALL_ACCESS) as key:
                winreg.SetValueEx(key, "fDenyTSConnections", 0, 
                                winreg.REG_DWORD, 0)
                
            # 3. 启用Network Level Authentication (NLA)
            reg_path = r"System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, 
                              winreg.KEY_ALL_ACCESS) as key:
                winreg.SetValueEx(key, "UserAuthentication", 0, 
                                winreg.REG_DWORD, 1)
            
            # 4. 设置端口
            self.change_rdp_port(port)
            
            # 5. 配置远程桌面服务
            self._run_command(['sc', 'config', 'TermService', 'start=auto'])
            
            # 6. 启动服务
            try:
                self._run_command(['net', 'start', 'TermService'])
            except subprocess.CalledProcessError:
                pass  # 忽略启动错误，检查实际状态
            
            # 等待一会儿让服务有时间启动
            time.sleep(2)
            
            # 7. 验证最终状态
            final_status = self._run_command(['sc', 'query', 'TermService'])
            if "RUNNING" in final_status.stdout:
                console.print("[green]远程桌面已成功启用！[/green]")
            else:
                console.print("[yellow]警告：服务可能未正常启动，但远程桌面已启用。请手动检查服务状态。[/yellow]")
            
            # 8. 配置防火墙规则
            try:
                # 删除现有规则
                self._run_command(['netsh', 'advfirewall', 'firewall', 'delete', 'rule',
                             'name="RDP Manager"'], check=False)
            except:
                pass

            # 添加新的入站规则
            self._run_command([
                'netsh', 'advfirewall', 'firewall', 'add', 'rule',
                'name=RDP Manager',
                'dir=in',
                'action=allow',
                'protocol=TCP',
                f'localport={port}',
                'program=%SystemRoot%\\system32\\svchost.exe',
                'service=TermService',
                'description=Remote Desktop Port'
            ])
            
            # 启用远程桌面防火墙规则组
            self._run_command([
                'netsh', 'advfirewall', 'firewall', 'set', 'rule',
                'group="远程桌面"',
                'new', 'enable=Yes'
            ])
            
        except Exception as e:
            console.print(f"[red]启用远程桌面时出错：{str(e)}[/red]")
            # 尝试恢复服务
            try:
                self._run_command(['net', 'start', 'TermService'], check=False)
            except:
                pass
            raise
            
    def disable_rdp(self) -> None:
        """禁用远程桌面"""
        self._require_admin()
        try:
            reg_path = r"System\CurrentControlSet\Control\Terminal Server"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, 
                              winreg.KEY_ALL_ACCESS) as key:
                winreg.SetValueEx(key, "fDenyTSConnections", 0, 
                                winreg.REG_DWORD, 1)
                                
            subprocess.run(['netsh', 'advfirewall', 'firewall', 'set', 'rule',
                          'group="远程桌面"', 'new', 'enable=No'],
                         check=True, capture_output=True)
            
            console.print("[green]远程桌面已成功禁用！[/green]")
        except Exception as e:
            console.print(f"[red]禁用远程桌面时出错：{str(e)}[/red]")
            raise
            
    def add_connection(self, name: str, host: str, username: str = DEFAULT_USERNAME, 
                      password: Optional[str] = None, port: int = DEFAULT_PORT) -> None:
        """添加新的远程桌面连接配置"""
        config = json.loads(self.config_file.read_text())
        cipher = self._get_cipher()
        
        connection = {
            "host": host,
            "port": port,
            "username": username,
            "password": cipher.encrypt(password.encode()).decode() if password else None
        }
        
        config[name] = connection
        self.config_file.write_text(json.dumps(config, indent=2))
        console.print(f"[green]已添加远程桌面配置：{name}[/green]")
        
    def list_connections(self) -> None:
        """列出所有保存的远程桌面连接"""
        config = json.loads(self.config_file.read_text())
        
        if not config:
            console.print("[yellow]没有保存的远程桌面配置[/yellow]")
            return
            
        table = Table(show_header=True, header_style="bold magenta")
        table.add_column("名称")
        table.add_column("主机地址")
        table.add_column("端口")
        table.add_column("用户名")
        
        for name, details in config.items():
            table.add_row(
                name,
                details["host"],
                str(details.get("port", DEFAULT_PORT)),
                details["username"]
            )
            
        console.print(table)
        
    def connect(self, name: str) -> None:
        """连接到指定的远程桌面"""
        config = json.loads(self.config_file.read_text())
        
        if name not in config:
            console.print(f"[red]未找到名为 {name} 的远程桌面配置[/red]")
            return
            
        connection = config[name]
        rdp_file = Path.home() / 'temp.rdp'
        
        # 生成RDP文件
        rdp_content = f"""
screen mode id:i:2
use multimon:i:0
desktopwidth:i:1920
desktopheight:i:1080
session bpp:i:32
winposstr:s:0,1,0,0,800,600
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
connection type:i:7
networkautodetect:i:1
bandwidthautodetect:i:1
displayconnectionbar:i:1
username:s:{connection["username"]}
full address:s:{connection["host"]}:{connection.get("port", DEFAULT_PORT)}
prompt for credentials:i:{"1" if not connection["password"] else "0"}
"""
        rdp_file.write_text(rdp_content)
        
        try:
            subprocess.Popen(['mstsc', str(rdp_file)])
            console.print(f"[green]正在连接到 {name}...[/green]")
        except Exception as e:
            console.print(f"[red]连接失败：{str(e)}[/red]")
        finally:
            # 延迟删除RDP文件
            import threading
            import time
            def delete_file():
                time.sleep(5)
                rdp_file.unlink(missing_ok=True)
            threading.Thread(target=delete_file).start()

    def get_rdp_status(self) -> tuple[bool, int]:
        """
        获取远程桌面状态
        返回: (是否启用, 当前端口号)
        """
        try:
            # 检查远程桌面服务状态
            result = subprocess.run(['sc', 'query', 'TermService'], 
                                 capture_output=True, text=True)
            service_running = "RUNNING" in result.stdout

            # 检查远程桌面设置
            reg_path = r"System\CurrentControlSet\Control\Terminal Server"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, 
                              winreg.KEY_READ) as key:
                rdp_enabled = not bool(winreg.QueryValueEx(key, "fDenyTSConnections")[0])

            # 获取当前端口
            reg_path = r"SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, 
                              winreg.KEY_READ) as key:
                current_port = int(winreg.QueryValueEx(key, "PortNumber")[0])

            return (service_running and rdp_enabled, current_port)
        except Exception:
            return (False, DEFAULT_PORT)

@click.group()
def cli():
    """Windows远程桌面批量管理工具"""
    pass

@cli.command()
@click.option('--port', '-p', default=DEFAULT_PORT, help='远程桌面端口号')
def enable(port):
    """启用远程桌面"""
    RDPManager().enable_rdp(port)

@cli.command()
def disable():
    """禁用远程桌面"""
    RDPManager().disable_rdp()

@cli.command()
@click.option('--port', '-p', default=DEFAULT_PORT, help='设置远程桌面端口号')
def set_port(port):
    """设置远程桌面端口"""
    RDPManager().change_rdp_port(port)

@cli.command()
@click.option('--name', '-n', required=True, help='连接名称')
@click.option('--host', '-h', required=True, help='主机地址')
@click.option('--username', '-u', default=DEFAULT_USERNAME, help='用户名')
@click.option('--password', '-p', help='密码（可选）')
@click.option('--port', '-P', default=DEFAULT_PORT, help='端口号')
def add(name, host, username, password, port):
    """添加远程桌面连接配置"""
    RDPManager().add_connection(name, host, username, password, port)

@cli.command()
def list():
    """列出所有保存的远程桌面连接"""
    RDPManager().list_connections()

@cli.command()
@click.argument('name')
def connect(name):
    """连接到指定的远程桌面"""
    RDPManager().connect(name)

if __name__ == '__main__':
    cli() 