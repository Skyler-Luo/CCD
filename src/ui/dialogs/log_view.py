# -*- coding: utf-8 -*-
"""
实时日志查看对话框组件
"""
import os
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QWidget, QMessageBox
)
from qfluentwidgets import (
    PushButton, ComboBox, TextEdit, CaptionLabel, StrongBodyLabel, TitleLabel, FluentStyleSheet, isDarkTheme
)
from src.utils.theme_utils import set_window_dark_title_bar
from src.utils.logger import Logger

def format_file_size(size_bytes):
    """格式化文件大小"""
    if size_bytes < 1024:
        return '{} B'.format(size_bytes)
    elif size_bytes < 1024 * 1024:
        return '{:.1f} KB'.format(size_bytes / 1024)
    elif size_bytes < 1024 * 1024 * 1024:
        return '{:.1f} MB'.format(size_bytes / (1024 * 1024))
    else:
        return '{:.1f} GB'.format(size_bytes / (1024 * 1024 * 1024))

class LogViewDialog(QDialog):
    """日志查看对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('查看日志')
        # 去掉右上角的问号
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.resize(900, 600)
        self.setAttribute(Qt.WA_StyledBackground, True)
        FluentStyleSheet.DIALOG.apply(self)
        set_window_dark_title_bar(self, isDarkTheme(), 0x002B2B2B if isDarkTheme() else 0x00F3F3F3)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        # 标题
        header = TitleLabel('日志查看')
        layout.addWidget(header)
        
        # 日志类型选择
        type_row = QWidget()
        type_layout = QHBoxLayout(type_row)
        type_layout.setContentsMargins(0, 0, 0, 0)
        type_layout.setSpacing(8)
        
        type_label = CaptionLabel('日志类型:')
        self.log_type_combo = ComboBox()
        self.log_type_combo.addItems(['全部日志', '错误日志'])
        self.log_type_combo.currentTextChanged.connect(self._load_log)
        type_layout.addWidget(type_label)
        type_layout.addWidget(self.log_type_combo)
        type_layout.addStretch(1)
        
        # 按钮组
        self.refresh_btn = PushButton('刷新')
        self.open_dir_btn = PushButton('打开目录')
        self.clear_btn = PushButton('清空')
        type_layout.addWidget(self.refresh_btn)
        type_layout.addWidget(self.open_dir_btn)
        type_layout.addWidget(self.clear_btn)
        
        layout.addWidget(type_row)
        
        # 日志文本框
        self.log_text = TextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText('暂无日志')
        layout.addWidget(self.log_text)
        
        # 底部状态信息
        self.log_info_label = CaptionLabel('')
        self.log_info_label.setStyleSheet('font-size: 12px;')
        layout.addWidget(self.log_info_label)
        
        # 绑定事件
        self.refresh_btn.clicked.connect(self._load_log)
        self.open_dir_btn.clicked.connect(self._open_log_dir)
        self.clear_btn.clicked.connect(self._clear_log)
        
        # 自动加载日志
        self._load_log()

    def _get_current_log_file(self):
        """获取当前选中的日志文件路径"""
        logger_instance = Logger()
        log_type = self.log_type_combo.currentText()
        if log_type == '错误日志':
            return logger_instance.error_log_file
        else:
            return logger_instance.log_file

    def _show_msg(self, title, text, icon=QMessageBox.Information):
        """公用原生提示框"""
        w = QMessageBox(self)
        w.setWindowTitle(title)
        w.setText(text)
        w.setIcon(icon)
        if isDarkTheme():
            w.setStyleSheet("""
                QMessageBox { background-color: rgb(32, 32, 32); color: white; }
                QLabel { color: white; font-family: "Microsoft YaHei"; font-size: 13px; }
                QPushButton { background-color: rgb(45, 45, 45); color: white; border: 1px solid rgb(60, 60, 60); border-radius: 4px; min-width: 80px; min-height: 28px; }
                QPushButton:hover { background-color: rgb(60, 60, 60); }
            """)
        else:
            w.setStyleSheet("""
                QMessageBox { background-color: rgb(243, 243, 243); color: black; }
                QLabel { color: black; font-family: "Microsoft YaHei"; font-size: 13px; }
                QPushButton { background-color: white; color: black; border: 1px solid rgb(200, 200, 200); border-radius: 4px; min-width: 80px; min-height: 28px; }
                QPushButton:hover { background-color: rgb(240, 240, 240); }
            """)
        w.exec_()

    def _load_log(self):
        """加载日志内容"""
        log_file = self._get_current_log_file()
        if not os.path.exists(log_file):
            self.log_text.setText('日志文件不存在')
            self.log_info_label.setText('加载失败')
            return
            
        try:
            size = os.path.getsize(log_file)
            size_str = format_file_size(size)
            
            with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
                
            total_lines = len(lines)
            max_lines = 10000
            
            if total_lines > max_lines:
                display_lines = lines[-max_lines:]
                content = ''.join(display_lines)
                content = f'... (省略前 {total_lines - max_lines} 行)\n\n' + content
            else:
                content = ''.join(lines)
            
            self.log_text.setText(content)
            
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.End)
            self.log_text.setTextCursor(cursor)
            
            self.log_info_label.setText(f'文件: {os.path.basename(log_file)} | 大小: {size_str} | 总行数: {total_lines}')
            
        except Exception as e:
            self.log_text.setText(f'读取日志失败:\n{str(e)}')
            self.log_info_label.setText('读取失败')

    def _open_log_dir(self):
        """打开日志目录"""
        logger_instance = Logger()
        log_dir = logger_instance.log_dir
        if os.path.exists(log_dir):
            QDesktopServices.openUrl(QUrl.fromLocalFile(log_dir))
        else:
            self._show_msg('目录不存在', f'日志目录不存在:\n{log_dir}', QMessageBox.Warning)

    def _clear_log(self):
        """清空当前日志文件"""
        w = QMessageBox(self)
        w.setWindowTitle('确认清空')
        w.setText('确定要清空当前日志文件吗？此操作不可恢复。')
        w.setIcon(QMessageBox.Question)
        
        if isDarkTheme():
            w.setStyleSheet("""
                QMessageBox { background-color: rgb(32, 32, 32); color: white; }
                QLabel { color: white; font-family: "Microsoft YaHei"; font-size: 13px; }
                QPushButton { background-color: rgb(45, 45, 45); color: white; border: 1px solid rgb(60, 60, 60); border-radius: 4px; min-width: 80px; min-height: 28px; }
                QPushButton:hover { background-color: rgb(60, 60, 60); }
            """)
        else:
            w.setStyleSheet("""
                QMessageBox { background-color: rgb(243, 243, 243); color: black; }
                QLabel { color: black; font-family: "Microsoft YaHei"; font-size: 13px; }
                QPushButton { background-color: white; color: black; border: 1px solid rgb(200, 200, 200); border-radius: 4px; min-width: 80px; min-height: 28px; }
                QPushButton:hover { background-color: rgb(240, 240, 240); }
            """)
            
        yes_btn = w.addButton('确定', QMessageBox.YesRole)
        cancel_btn = w.addButton('取消', QMessageBox.NoRole)
        
        w.exec_()
        
        if w.clickedButton() == yes_btn:
            log_file = self._get_current_log_file()
            try:
                with open(log_file, 'w', encoding='utf-8') as f:
                    f.write('')
                self._load_log()
                self._show_msg('成功', '日志已清空', QMessageBox.Information)
            except Exception as e:
                self._show_msg('失败', f'清空日志失败:\n{str(e)}', QMessageBox.Critical)
