#!/usr/bin/env python3
"""
Windows远程桌面批量管理工具 - 图形界面版本
"""

import sys
import json
import subprocess
import time
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QLabel, QLineEdit, 
                           QMessageBox, QTableWidget, QTableWidgetItem, 
                           QHeaderView, QDialog, QFormLayout, QSpinBox,
                           QGroupBox, QToolBar)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QFont
import rdp_manager

DEFAULT_PORT = 3389
DEFAULT_USERNAME = "administrator"

class AddConnectionDialog(QDialog):
    """添加连接对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("添加远程桌面连接")
        self.setMinimumWidth(400)
        
        layout = QFormLayout()
        
        # 创建输入框
        self.name_edit = QLineEdit()
        self.host_edit = QLineEdit()
        self.port_spinbox = QSpinBox()
        self.port_spinbox.setRange(1, 65535)
        self.port_spinbox.setValue(DEFAULT_PORT)
        self.username_edit = QLineEdit()
        self.username_edit.setText(DEFAULT_USERNAME)
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        
        # 添加到布局
        layout.addRow("连接名称:", self.name_edit)
        layout.addRow("主机地址:", self.host_edit)
        layout.addRow("端口:", self.port_spinbox)
        layout.addRow("用户名:", self.username_edit)
        layout.addRow("密码:", self.password_edit)
        
        # 按钮
        buttons = QHBoxLayout()
        save_btn = QPushButton("保存")
        cancel_btn = QPushButton("取消")
        
        save_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        
        buttons.addWidget(save_btn)
        buttons.addWidget(cancel_btn)
        layout.addRow(buttons)
        
        self.setLayout(layout)
    
    def get_data(self):
        """获取输入的数据"""
        return {
            "name": self.name_edit.text(),
            "host": self.host_edit.text(),
            "port": self.port_spinbox.value(),
            "username": self.username_edit.text(),
            "password": self.password_edit.text()
        }

class PasswordTableItem(QTableWidgetItem):
    """密码单元格项，用于加密显示密码"""
    def __init__(self, cipher):
        super().__init__()
        self.cipher = cipher
        self.encrypted_password = None
        self.setFlags(self.flags() | Qt.ItemFlag.ItemIsEditable)
        
    def set_encrypted_password(self, encrypted_password):
        """设置加密密码"""
        self.encrypted_password = encrypted_password
        self.update_display()
        
    def get_decrypted_password(self):
        """获取解密后的密码"""
        if not self.encrypted_password:
            return ""
        try:
            return self.cipher.decrypt(self.encrypted_password.encode()).decode()
        except:
            return ""
            
    def update_display(self, show_password=False):
        """更新显示的文本"""
        if not self.encrypted_password:
            self.setText("")
        elif show_password:
            self.setText(self.get_decrypted_password())
        else:
            self.setText("●●●●●●")

class RDPManagerGUI(QMainWindow):
    """远程桌面管理器主窗口"""
    def __init__(self):
        super().__init__()
        self.rdp = rdp_manager.RDPManager()
        self.show_passwords = False  # 添加密码显示状态标志
        self.init_ui()
        # 启动时检查状态
        self.update_rdp_status()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("远程桌面管理器")
        self.setMinimumSize(800, 600)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        layout = QVBoxLayout(central_widget)
        
        # 状态部分
        status_group = QGroupBox("远程桌面状态")
        status_layout = QVBoxLayout(status_group)
        
        # 状态显示和基本控制
        basic_controls = QHBoxLayout()
        self.status_label = QLabel("正在检查远程桌面状态...")
        self.status_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        
        enable_btn = QPushButton("启用远程桌面")
        enable_btn.clicked.connect(self.enable_rdp)
        
        disable_btn = QPushButton("禁用远程桌面")
        disable_btn.clicked.connect(self.disable_rdp)
        
        refresh_btn = QPushButton("刷新状态")
        refresh_btn.clicked.connect(self.update_rdp_status)
        
        basic_controls.addWidget(self.status_label)
        basic_controls.addWidget(enable_btn)
        basic_controls.addWidget(disable_btn)
        basic_controls.addWidget(refresh_btn)
        
        # 端口设置
        port_controls = QHBoxLayout()
        port_label = QLabel("远程桌面端口:")
        self.port_spinbox = QSpinBox()
        self.port_spinbox.setRange(1, 65535)
        self.port_spinbox.setValue(DEFAULT_PORT)
        apply_port_btn = QPushButton("应用端口设置")
        apply_port_btn.clicked.connect(self.apply_port_settings)
        
        port_controls.addWidget(port_label)
        port_controls.addWidget(self.port_spinbox)
        port_controls.addWidget(apply_port_btn)
        port_controls.addStretch()
        
        status_layout.addLayout(basic_controls)
        status_layout.addLayout(port_controls)
        
        layout.addWidget(status_group)
        
        # 连接列表
        self.table = QTableWidget()
        self.table.setColumnCount(6)  # 增加复选框列
        self.table.setHorizontalHeaderLabels(["选择", "连接名称", "主机地址", "端口", "用户名", "密码"])
        # 设置第一列宽度较小
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(0, 50)
        # 其他列自适应宽度
        for i in range(1, 6):
            self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        
        # 允许编辑单元格
        self.table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked | 
                                 QTableWidget.EditTrigger.EditKeyPressed)
        # 连接编辑完成信号
        self.table.itemChanged.connect(self.on_item_changed)
        layout.addWidget(self.table)
        
        # 按钮组
        button_group = QWidget()
        button_layout = QHBoxLayout(button_group)
        
        add_btn = QPushButton("添加连接")
        add_btn.clicked.connect(self.add_connection)
        
        connect_btn = QPushButton("连接选中")
        connect_btn.clicked.connect(self.connect_selected)
        
        delete_btn = QPushButton("删除")
        delete_btn.clicked.connect(self.delete_selected)
        
        select_all_btn = QPushButton("全选")
        select_all_btn.clicked.connect(self.select_all)
        
        deselect_all_btn = QPushButton("取消全选")
        deselect_all_btn.clicked.connect(self.deselect_all)
        
        button_layout.addWidget(select_all_btn)
        button_layout.addWidget(deselect_all_btn)
        button_layout.addWidget(add_btn)
        button_layout.addWidget(connect_btn)
        button_layout.addWidget(delete_btn)
        
        layout.addWidget(button_group)
        
        # 添加显示密码按钮
        self.show_password_btn = QPushButton('显示密码', self)
        self.show_password_btn.clicked.connect(self.toggle_password_display)
        
        # 修改工具栏布局
        toolbar = QToolBar()
        toolbar.addWidget(port_label)
        toolbar.addWidget(self.port_spinbox)
        toolbar.addWidget(apply_port_btn)
        toolbar.addWidget(enable_btn)
        toolbar.addWidget(disable_btn)
        toolbar.addWidget(self.show_password_btn)  # 添加显示密码按钮
        self.addToolBar(toolbar)
        
        # 更新连接列表
        self.refresh_connections()
        
    def create_checkbox_item(self):
        """创建居中的复选框单元格"""
        item = QTableWidgetItem()
        item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
        item.setCheckState(Qt.CheckState.Unchecked)
        return item

    def refresh_connections(self):
        """刷新连接列表"""
        # 暂时断开信号连接，防止触发itemChanged
        self.table.itemChanged.disconnect(self.on_item_changed)
        try:
            self.table.setRowCount(0)
            config = self.rdp.config_file.read_text()
            if not config:
                return
                
            connections = json.loads(config)
            self.table.setRowCount(len(connections))
            
            for row, (name, details) in enumerate(connections.items()):
                # 添加复选框
                self.table.setItem(row, 0, self.create_checkbox_item())
                self.table.setItem(row, 1, QTableWidgetItem(name))
                self.table.setItem(row, 2, QTableWidgetItem(details["host"]))
                self.table.setItem(row, 3, QTableWidgetItem(str(details.get("port", DEFAULT_PORT))))
                self.table.setItem(row, 4, QTableWidgetItem(details["username"]))
                
                # 添加密码列，解密显示
                password_item = PasswordTableItem(self.rdp._get_cipher())
                if details.get("password"):
                    password_item.set_encrypted_password(details["password"])
                self.table.setItem(row, 5, password_item)
        finally:
            # 重新连接信号
            self.table.itemChanged.connect(self.on_item_changed)

    def on_item_changed(self, item):
        """处理表格项编辑完成事件"""
        try:
            # 暂时断开信号连接，防止递归触发
            self.table.itemChanged.disconnect(self.on_item_changed)
            
            row = item.row()
            col = item.column()
            
            # 如果是复选框列，不需要处理
            if col == 0:
                return
                
            new_value = item.text()
            
            # 获取连接名称（第二列）
            name = self.table.item(row, 1).text()
            
            # 读取当前配置
            config = json.loads(self.rdp.config_file.read_text())
            if name not in config:
                return
            
            connection = config[name]
            
            # 根据列更新相应的值
            if col == 1:  # 连接名称
                if new_value != name and new_value in config:
                    QMessageBox.warning(self, "警告", "连接名称已存在！")
                    item.setText(name)  # 恢复原值
                    return
                if new_value != name:
                    # 更新连接名称
                    config[new_value] = config.pop(name)
            elif col == 2:  # 主机地址
                connection["host"] = new_value
            elif col == 3:  # 端口
                try:
                    port = int(new_value)
                    if 1 <= port <= 65535:
                        connection["port"] = port
                    else:
                        raise ValueError("端口范围无效")
                except ValueError:
                    QMessageBox.warning(self, "警告", "端口必须是1-65535之间的数字！")
                    item.setText(str(connection.get("port", DEFAULT_PORT)))
                    return
            elif col == 4:  # 用户名
                connection["username"] = new_value
            elif col == 5:  # 密码
                # 获取密码项
                password_item = self.table.item(row, col)
                if isinstance(password_item, PasswordTableItem):
                    # 更新密码
                    if new_value and new_value != "●●●●●●":
                        cipher = self.rdp._get_cipher()
                        encrypted_password = cipher.encrypt(new_value.encode()).decode()
                        connection["password"] = encrypted_password
                        password_item.set_encrypted_password(encrypted_password)
                    elif not new_value:
                        connection["password"] = None
                        password_item.set_encrypted_password(None)
                    # 恢复显示为掩码
                    password_item.update_display(self.show_passwords)
            
            # 保存更新后的配置
            self.rdp.config_file.write_text(json.dumps(config, indent=2))
            
        finally:
            # 重新连接信号
            self.table.itemChanged.connect(self.on_item_changed)

    def select_all(self):
        """全选"""
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item:
                item.setCheckState(Qt.CheckState.Checked)

    def deselect_all(self):
        """取消全选"""
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item:
                item.setCheckState(Qt.CheckState.Unchecked)

    def get_checked_connections(self):
        """获取所有勾选的连接名称"""
        checked = []
        for row in range(self.table.rowCount()):
            checkbox = self.table.item(row, 0)
            if checkbox and checkbox.checkState() == Qt.CheckState.Checked:
                name = self.table.item(row, 1).text()
                checked.append(name)
        return checked

    def connect_selected(self):
        """连接选中的远程桌面"""
        checked = self.get_checked_connections()
        if not checked:
            # 如果没有勾选的连接，则连接当前选中的行
            current_row = self.table.currentRow()
            if current_row >= 0:
                name = self.table.item(current_row, 1).text()
                try:
                    self.rdp.connect(name)
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"连接失败：{str(e)}")
            else:
                QMessageBox.warning(self, "警告", "请先选择要连接的远程桌面！")
            return

        # 批量连接
        for name in checked:
            try:
                self.rdp.connect(name)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"连接 {name} 失败：{str(e)}")
    
    def update_rdp_status(self):
        """更新远程桌面状态显示"""
        enabled, current_port = self.rdp.get_rdp_status()
        if enabled:
            self.status_label.setText(f"远程桌面状态: 已启用 (端口: {current_port})")
            self.status_label.setStyleSheet("color: green")
        else:
            self.status_label.setText("远程桌面状态: 已禁用")
            self.status_label.setStyleSheet("color: red")
        # 更新端口显示
        self.port_spinbox.setValue(current_port)
    
    def enable_rdp(self):
        """启用远程桌面"""
        try:
            port = self.port_spinbox.value()
            self.rdp.enable_rdp(port)
            self.update_rdp_status()
            QMessageBox.information(self, "成功", f"远程桌面已成功启用！端口: {port}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"启用远程桌面时出错：{str(e)}")
    
    def disable_rdp(self):
        """禁用远程桌面"""
        try:
            self.rdp.disable_rdp()
            self.update_rdp_status()
            QMessageBox.information(self, "成功", "远程桌面已成功禁用！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"禁用远程桌面时出错：{str(e)}")
    
    def apply_port_settings(self):
        """应用端口设置"""
        try:
            port = self.port_spinbox.value()
            self.rdp.change_rdp_port(port)
            self.update_rdp_status()
            
            # 直接重启远程桌面服务
            try:
                subprocess.run(['net', 'stop', 'TermService', '/y'], 
                             capture_output=True)
                time.sleep(2)
                subprocess.run(['net', 'start', 'TermService'], 
                             capture_output=True)
            except Exception as e:
                QMessageBox.warning(self, "警告", f"重启服务时出错：{str(e)}")
                return
            
            # 显示成功消息和后续步骤提示
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setWindowTitle("端口修改成功")
            msg.setText(f"远程桌面端口已更改为：{port}\n远程桌面服务已重启！")
            
            # 添加详细信息
            details = f"""
重要提示：
1. 请确保在防火墙中允许新端口 {port} 的入站连接
2. 连接时需要在地址后指定新端口，例如：
   192.168.0.102:{port}
3. 如果远程连接仍然失败，请尝试重启计算机
            """
            msg.setInformativeText(details)
            msg.exec()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"更改端口时出错：{str(e)}")
    
    def add_connection(self):
        """添加新连接"""
        dialog = AddConnectionDialog(self)
        if dialog.exec():
            data = dialog.get_data()
            try:
                self.rdp.add_connection(
                    data["name"], 
                    data["host"], 
                    data["username"], 
                    data["password"],
                    data["port"]
                )
                self.refresh_connections()
                QMessageBox.information(self, "成功", "连接已成功添加！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"添加连接时出错：{str(e)}")
    
    def delete_selected(self):
        """删除选中的连接"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "警告", "请先选择一个连接！")
            return
            
        name = self.table.item(current_row, 1).text()
        reply = QMessageBox.question(
            self, "确认删除", 
            f"确定要删除连接 '{name}' 吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            config = eval(self.rdp.config_file.read_text())
            if name in config:
                del config[name]
                self.rdp.config_file.write_text(str(config))
                self.refresh_connections()
                QMessageBox.information(self, "成功", "连接已成功删除！")

    def toggle_password_display(self):
        """切换密码显示状态"""
        self.show_passwords = not self.show_passwords
        self.show_password_btn.setText('隐藏密码' if self.show_passwords else '显示密码')
        self.update_password_display()
        
    def update_password_display(self):
        """更新密码显示"""
        for row in range(self.table.rowCount()):
            password_item = self.table.item(row, 5)
            if isinstance(password_item, PasswordTableItem):
                password_item.update_display(self.show_passwords)

def main():
    app = QApplication(sys.argv)
    window = RDPManagerGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 