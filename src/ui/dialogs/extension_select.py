# -*- coding: utf-8 -*-
"""
文件后缀选择对话框组件
"""
from PyQt5.QtCore import QEvent, Qt
from PyQt5.QtWidgets import (
    QDialog, QListWidgetItem, QAbstractItemView, QDialogButtonBox, QStyle, QStyleOptionViewItem, QVBoxLayout, QHBoxLayout, QWidget
)
from qfluentwidgets import (
    PushButton, ListWidget, CaptionLabel, StrongBodyLabel, FluentStyleSheet, isDarkTheme
)
from src.utils.theme_utils import set_window_dark_title_bar

class ExtensionSelectDialog(QDialog):
    """文件后缀选择对话框"""
    def __init__(self, exts, selected, parent=None):
        super().__init__(parent)
        self.setWindowTitle('选择文件后缀')
        self.resize(420, 480)
        self.setAttribute(Qt.WA_StyledBackground, True)
        FluentStyleSheet.DIALOG.apply(self)
        set_window_dark_title_bar(self, isDarkTheme(), 0x002B2B2B if isDarkTheme() else 0x00F3F3F3)
        self.setWindowFlags((self.windowFlags() | Qt.Window) & ~Qt.WindowContextHelpButtonHint)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        header = StrongBodyLabel('可选后缀')
        header.setStyleSheet('font-size: 13px;')
        layout.addWidget(header)

        hint = CaptionLabel('请勾选需要提取的后缀')
        hint.setStyleSheet('font-size: 12px;')
        layout.addWidget(hint)

        self.list_widget = ListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.NoSelection)
        self.list_widget.setAlternatingRowColors(True)
        for ext in exts:
            item = QListWidgetItem(ext)
            if ext in selected:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget, 1)

        action_row = QWidget()
        action_layout = QHBoxLayout(action_row)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(8)
        self.select_all_btn = PushButton('全选')
        self.clear_btn = PushButton('清空')
        action_layout.addWidget(self.select_all_btn)
        action_layout.addWidget(self.clear_btn)
        action_layout.addStretch(1)
        self.count_label = CaptionLabel('')
        self.count_label.setStyleSheet('font-size: 12px;')
        action_layout.addWidget(self.count_label)
        layout.addWidget(action_row)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.select_all_btn.clicked.connect(self._select_all)
        self.clear_btn.clicked.connect(self._clear_all)
        self.list_widget.itemChanged.connect(self._update_count)
        self.list_widget.viewport().installEventFilter(self)
        self._update_count()

    def eventFilter(self, obj, event):
        if obj is self.list_widget.viewport() and event.type() == QEvent.MouseButtonPress:
            item = self.list_widget.itemAt(event.pos())
            if item:
                option = QStyleOptionViewItem()
                option.rect = self.list_widget.visualItemRect(item)
                option.state = QStyle.State_Enabled
                if item.checkState() == Qt.Checked:
                    option.state |= QStyle.State_On
                else:
                    option.state |= QStyle.State_Off
                check_rect = self.list_widget.style().subElementRect(
                    QStyle.SE_ItemViewItemCheckIndicator,
                    option,
                    self.list_widget
                )
                if not check_rect.contains(event.pos()):
                    item.setCheckState(Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked)
                    return True
        return super().eventFilter(obj, event)

    def _select_all(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if not item.isHidden():
                item.setCheckState(Qt.Checked)

    def _clear_all(self):
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(Qt.Unchecked)

    def _update_count(self):
        total = self.list_widget.count()
        checked = sum(1 for i in range(total) if self.list_widget.item(i).checkState() == Qt.Checked)
        self.count_label.setText('已选 {} / {}'.format(checked, total))

    def get_selected(self):
        selected = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                selected.append(item.text())
        return sorted(selected)
