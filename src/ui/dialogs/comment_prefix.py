# -*- coding: utf-8 -*-
"""
注释前缀选择对话框组件
"""
from PyQt5.QtCore import QEvent, Qt
from PyQt5.QtWidgets import (
    QDialog, QListWidgetItem, QAbstractItemView, QDialogButtonBox, QStyle, QStyleOptionViewItem, QVBoxLayout, QHBoxLayout, QWidget
)
from qfluentwidgets import (
    PushButton, ListWidget, CaptionLabel, StrongBodyLabel, FluentStyleSheet, isDarkTheme
)
from src.utils.theme_utils import set_window_dark_title_bar

# 注释前缀映射
COMMENT_PREFIX_BY_LANG = {
    'python': ['#', '"""', "'''"],
    'javascript': ['//', '/*', '*/'],
    'typescript': ['//', '/*', '*/'],
    'go': ['//', '/*', '*/'],
    'php': ['//', '#', '/*', '*/'],
    'csharp': ['//', '/*', '*/'],
    'kotlin': ['//', '/*', '*/'],
    'swift': ['//', '/*', '*/'],
    'rust': ['//', '/*', '*/'],
    'dart': ['//', '/*', '*/'],
    'scala': ['//', '/*', '*/'],
    'sql': ['--', '/*', '*/'],
    'r': ['#'],
    'lua': ['--', '--[[', ']]'],
    'powershell': ['#', '<#', '#>'],
    'yaml': ['#'],
    'java': ['//', '/*', '*/'],
    'c': ['//', '/*', '*/'],
    'cpp': ['//', '/*', '*/'],
    'html': ['<!--', '-->'],
    'xml': ['<!--', '-->'],
    'css': ['/*', '*/'],
    'shellscript': ['#'],
    'ruby': ['#', '=begin', '=end'],
    'perl': ['#', '=begin', '=end']
}

class CommentPrefixDialog(QDialog):
    """注释前缀选择对话框"""
    def __init__(self, langs, selected_langs, parent=None):
        super().__init__(parent)
        self.setWindowTitle('选择注释前缀')
        self.resize(420, 460)
        self.setAttribute(Qt.WA_StyledBackground, True)
        FluentStyleSheet.DIALOG.apply(self)
        set_window_dark_title_bar(self, isDarkTheme(), 0x002B2B2B if isDarkTheme() else 0x00F3F3F3)
        self.setWindowFlags((self.windowFlags() | Qt.Window) & ~Qt.WindowContextHelpButtonHint)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        header = StrongBodyLabel('按语言选择注释前缀')
        header.setStyleSheet('font-size: 13px;')
        layout.addWidget(header)

        hint = CaptionLabel('勾选语言后将应用对应的注释前缀')
        hint.setStyleSheet('font-size: 12px;')
        layout.addWidget(hint)

        self.list_widget = ListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.NoSelection)
        self.list_widget.setAlternatingRowColors(True)
        for lang in langs:
            prefixes = list(COMMENT_PREFIX_BY_LANG.get(lang, []))
            if lang in ('c', 'cpp', 'java', 'javascript', 'typescript') and '//' not in prefixes:
                prefixes.insert(0, '//')
            label = '{}  ({})'.format(lang, ', '.join(prefixes)) if prefixes else lang
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, lang)
            if lang in selected_langs:
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
            self.list_widget.item(i).setCheckState(Qt.Checked)

    def _clear_all(self):
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(Qt.Unchecked)

    def _update_count(self):
        total = self.list_widget.count()
        checked = sum(1 for i in range(total) if self.list_widget.item(i).checkState() == Qt.Checked)
        self.count_label.setText('已选 {} / {}'.format(checked, total))

    def get_selected_langs(self):
        selected = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                selected.append(item.data(Qt.UserRole))
        return selected

    def get_selected_prefixes(self):
        prefixes = []
        for lang in self.get_selected_langs():
            for prefix in COMMENT_PREFIX_BY_LANG.get(lang, []):
                if prefix not in prefixes:
                    prefixes.append(prefix)
        return prefixes
