# -*- coding: utf-8 -*-
"""
UI对话框模块
包含所有GUI对话框组件
"""
import os
from PyQt5.QtCore import QEvent, Qt, QUrl, QTimer, QRegExp
from PyQt5.QtGui import QDesktopServices, QPalette, QSyntaxHighlighter, QTextCharFormat, QColor, QFont
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QDialog, QListWidgetItem,
    QAbstractItemView, QDialogButtonBox, QStyle,
    QStyleOptionViewItem, QLabel
)
from qfluentwidgets import (
    PushButton, LineEdit, BodyLabel, PrimaryPushButton, TextEdit,
    ListWidget, ComboBox, CaptionLabel, StrongBodyLabel, MessageBox,
    isDarkTheme, TitleLabel, FluentStyleSheet, CardWidget
)

from src.core.code_processor import LANGUAGE_BY_EXT, decode_content, get_language_by_extension, filter_lines
from src.utils.logger import Logger
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


class CodeHighlighter(QSyntaxHighlighter):
    """用于代码预览区域的轻量级语法高亮器，支持深浅色模式自动适配"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.highlighting_rules = []

        is_dark = isDarkTheme()
        
        # 关键字高亮规则
        keyword_format = QTextCharFormat()
        keyword_format.setForeground(QColor("#569CD6" if is_dark else "#007ACC"))
        keyword_format.setFontWeight(QFont.Bold)
        keywords = [
            "class", "def", "function", "var", "let", "const", "if", "else", "for", "while",
            "return", "import", "from", "as", "public", "private", "protected", "interface",
            "package", "namespace", "using", "include", "struct", "void", "int", "float",
            "double", "bool", "char", "string", "func", "fn", "mut", "use", "mod", "pub", 
            "impl", "trait", "self", "this"
        ]
        for word in keywords:
            rule = (QRegExp(r"\b" + word + r"\b"), keyword_format)
            self.highlighting_rules.append(rule)

        # 字符串高亮规则
        string_format = QTextCharFormat()
        string_format.setForeground(QColor("#CE9178" if is_dark else "#A31515"))
        self.highlighting_rules.append((QRegExp(r'"[^"\\]*(\\.[^"\\]*)*"'), string_format))
        self.highlighting_rules.append((QRegExp(r"'[^'\\]*(\\.[^'\\]*)*'"), string_format))

        # 注释高亮规则（绿色）
        comment_format = QTextCharFormat()
        comment_format.setForeground(QColor("#6A9955" if is_dark else "#008000"))
        self.highlighting_rules.append((QRegExp(r"//[^\n]*"), comment_format))
        self.highlighting_rules.append((QRegExp(r"#[^\n]*"), comment_format))
        self.highlighting_rules.append((QRegExp(r"--[^\n]*"), comment_format))

    def highlightBlock(self, text):
        for pattern, format in self.highlighting_rules:
            expression = QRegExp(pattern)
            index = expression.indexIn(text)
            while index >= 0:
                length = expression.matchedLength()
                self.setFormat(index, length, format)
                index = expression.indexIn(text, index + length)


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



class FileSelectDialog(QDialog):
    """文件多选对话框，左右两栏布局，支持多选、顺序调整和文件预览"""
    def __init__(self, files, selected_files, skip_blank_lines=True, skip_comment_lines=True, comment_chars=None, encoding='auto', parent=None):
        super().__init__(parent)
        self.setWindowTitle('选择源文件并排序')
        self.resize(900, 700)
        self.setAttribute(Qt.WA_StyledBackground, True)
        FluentStyleSheet.DIALOG.apply(self)
        set_window_dark_title_bar(self, isDarkTheme(), 0x002B2B2B if isDarkTheme() else 0x00F3F3F3)
        self.setWindowFlags((self.windowFlags() | Qt.Window) & ~Qt.WindowContextHelpButtonHint)
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)
        
        # 保存过滤配置
        self.skip_blank_lines = skip_blank_lines
        self.skip_comment_lines = skip_comment_lines
        self.comment_chars = comment_chars if comment_chars else []
        self.encoding = encoding

        # 标题和提示
        header = StrongBodyLabel('选择源文件并排序')
        header.setStyleSheet('font-size: 14px;')
        main_layout.addWidget(header)

        hint = CaptionLabel('从左侧选择文件添加到右侧，可调整右侧文件顺序，支持多选操作')
        hint.setStyleSheet('font-size: 12px;')
        main_layout.addWidget(hint)

        # 主要内容区域：左右两栏 + 中间按钮
        content_layout = QHBoxLayout()
        content_layout.setSpacing(12)
        
        # 左栏：可选文件
        left_panel = CardWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(12, 12, 12, 12)
        left_layout.setSpacing(8)
        
        left_header = StrongBodyLabel('可选文件')
        left_header.setStyleSheet('font-size: 13px;')
        left_layout.addWidget(left_header)
        
        # 排序选项行
        sort_row = QWidget()
        sort_layout = QHBoxLayout(sort_row)
        sort_layout.setContentsMargins(0, 0, 0, 0)
        sort_layout.setSpacing(8)
        
        sort_label = CaptionLabel('排序:')
        sort_label.setStyleSheet('font-size: 11px;')
        sort_layout.addWidget(sort_label)
        
        self.sort_combo = ComboBox()
        self.sort_combo.addItems(['默认顺序', '按文件名', '按代码行数↓', '按代码行数↑'])
        self.sort_combo.setFixedWidth(120)
        self.sort_combo.currentTextChanged.connect(self._sort_available_files)
        sort_layout.addWidget(self.sort_combo)
        sort_layout.addStretch(1)
        
        left_layout.addWidget(sort_row)
        
        self.search_box = LineEdit()
        self.search_box.setPlaceholderText('搜索文件...')
        self.search_box.textChanged.connect(self._filter_available_files)
        left_layout.addWidget(self.search_box)
        
        self.available_list = ListWidget()
        self.available_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.available_list.setAlternatingRowColors(True)
        # 禁用自动滚动到选中项
        self.available_list.setAutoScroll(False)
        left_layout.addWidget(self.available_list, 1)
        
        self.available_count_label = CaptionLabel('')
        self.available_count_label.setStyleSheet('font-size: 11px;')
        left_layout.addWidget(self.available_count_label)
        
        content_layout.addWidget(left_panel, 1)
        
        # 中间按钮区域
        middle_panel = QWidget()
        middle_layout = QVBoxLayout(middle_panel)
        middle_layout.setContentsMargins(0, 0, 0, 0)
        middle_layout.setSpacing(8)
        middle_layout.addStretch(1)
        
        self.add_btn = PrimaryPushButton('添加 >>')
        self.add_btn.setMinimumWidth(100)
        self.add_btn.setMinimumHeight(36)
        middle_layout.addWidget(self.add_btn)
        
        self.add_all_btn = PushButton('全部添加')
        self.add_all_btn.setMinimumWidth(100)
        middle_layout.addWidget(self.add_all_btn)
        
        middle_layout.addSpacing(20)
        
        self.remove_btn = PushButton('<< 移除')
        self.remove_btn.setMinimumWidth(100)
        self.remove_btn.setMinimumHeight(36)
        middle_layout.addWidget(self.remove_btn)
        
        self.remove_all_btn = PushButton('全部移除')
        self.remove_all_btn.setMinimumWidth(100)
        middle_layout.addWidget(self.remove_all_btn)
        
        middle_layout.addStretch(1)
        content_layout.addWidget(middle_panel)
        
        # 右栏：已选文件
        right_panel = CardWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(12, 12, 12, 12)
        right_layout.setSpacing(8)
        
        right_header = StrongBodyLabel('已选文件（按顺序生成）')
        right_header.setStyleSheet('font-size: 13px;')
        right_layout.addWidget(right_header)
        
        self.selected_list = ListWidget()
        self.selected_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.selected_list.setAlternatingRowColors(True)
        # 禁用自动滚动到选中项
        self.selected_list.setAutoScroll(False)
        right_layout.addWidget(self.selected_list, 1)
        
        # 排序按钮
        order_buttons = QHBoxLayout()
        order_buttons.setSpacing(8)
        
        self.move_top_btn = PushButton('置顶')
        self.move_up_btn = PushButton('上移')
        self.move_down_btn = PushButton('下移')
        self.move_bottom_btn = PushButton('置底')
        
        order_buttons.addWidget(self.move_top_btn)
        order_buttons.addWidget(self.move_up_btn)
        order_buttons.addWidget(self.move_down_btn)
        order_buttons.addWidget(self.move_bottom_btn)
        order_buttons.addStretch(1)
        
        right_layout.addLayout(order_buttons)
        
        self.selected_count_label = CaptionLabel('')
        self.selected_count_label.setStyleSheet('font-size: 11px;')
        right_layout.addWidget(self.selected_count_label)
        
        content_layout.addWidget(right_panel, 1)
        
        main_layout.addLayout(content_layout, 1)
        
        # 预览面板
        preview_panel = CardWidget()
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(12, 12, 12, 12)
        preview_layout.setSpacing(8)
        
        preview_header_layout = QHBoxLayout()
        preview_header = StrongBodyLabel('文件预览')
        preview_header.setStyleSheet('font-size: 13px;')
        preview_header_layout.addWidget(preview_header)
        
        self.preview_lines_label = CaptionLabel('预览行数:')
        self.preview_lines_label.setStyleSheet('font-size: 11px;')
        self.preview_lines_combo = ComboBox()
        self.preview_lines_combo.addItems(['20', '50', '100', '200'])
        self.preview_lines_combo.setCurrentText('50')
        self.preview_lines_combo.setFixedWidth(80)
        self.preview_lines_combo.currentTextChanged.connect(self._refresh_preview)
        preview_header_layout.addWidget(self.preview_lines_label)
        preview_header_layout.addWidget(self.preview_lines_combo)
        preview_header_layout.addStretch(1)
        
        self.preview_info_label = CaptionLabel('')
        self.preview_info_label.setStyleSheet('font-size: 11px;')
        preview_header_layout.addWidget(self.preview_info_label)
        
        preview_layout.addLayout(preview_header_layout)
        
        self.preview_text = TextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setPlaceholderText('点击文件名查看预览...')
        self.preview_text.setMinimumHeight(150)
        self.preview_text.setMaximumHeight(200)
        from PyQt5.QtGui import QFont
        preview_font = QFont('Consolas', 9)
        if not preview_font.exactMatch():
            preview_font = QFont('Courier New', 9)
        self.preview_text.setFont(preview_font)
        preview_layout.addWidget(self.preview_text)
        
        # 绑定语法高亮器
        self.preview_highlighter = CodeHighlighter(self.preview_text.document())
        
        main_layout.addWidget(preview_panel)
        
        # 底部按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        main_layout.addWidget(buttons)
        
        # 初始化数据
        self.all_files = files
        self._current_preview_file = None
        self._file_cache = {}  # 缓存文件信息（行数、大小等）
        self._init_file_lists(files, selected_files)
        
        # 节流更新定时器（避免后台更新过于频繁造成卡顿）
        self.update_timer = QTimer(self)
        self.update_timer.setSingleShot(True)
        self.update_timer.timeout.connect(self._update_counts)

        # 启动后台线程统计文件信息
        from src.ui.ui_tasks import FileInfoWorker
        self.loader_worker = FileInfoWorker(
            files,
            self.skip_blank_lines,
            self.skip_comment_lines,
            self.comment_chars,
            self.encoding
        )
        self.loader_worker.file_processed.connect(self._on_file_processed)
        self.loader_worker.finished.connect(self._on_loader_finished)
        self.loader_worker.start()

        # 连接信号
        self.add_btn.clicked.connect(self._add_selected)
        self.add_all_btn.clicked.connect(self._add_all)
        self.remove_btn.clicked.connect(self._remove_selected)
        self.remove_all_btn.clicked.connect(self._remove_all)
        self.move_top_btn.clicked.connect(self._move_to_top)
        self.move_up_btn.clicked.connect(self._move_up)
        self.move_down_btn.clicked.connect(self._move_down)
        self.move_bottom_btn.clicked.connect(self._move_to_bottom)
        
        self.available_list.itemDoubleClicked.connect(self._on_available_double_click)
        self.selected_list.itemDoubleClicked.connect(self._on_selected_double_click)
        self.available_list.itemClicked.connect(self._on_file_clicked)
        self.selected_list.itemClicked.connect(self._on_file_clicked)
        
        self._update_counts()
    
    def accept(self):
        self._stop_worker()
        super().accept()

    def reject(self):
        self._stop_worker()
        super().reject()

    def closeEvent(self, event):
        self._stop_worker()
        super().closeEvent(event)

    def _stop_worker(self):
        if hasattr(self, 'loader_worker') and self.loader_worker and self.loader_worker.isRunning():
            self.loader_worker.requestInterruption()
            self.loader_worker.wait()

    def _on_file_processed(self, file_path, info):
        self._file_cache[file_path] = info
        self.update_timer.start(100) # 100ms 刷新一次

    def _on_loader_finished(self, cache):
        self._file_cache.update(cache)
        self._update_counts()
        if self.sort_combo.currentText() != '默认顺序':
            self._sort_available_files()
    
    def _get_file_info(self, file_path):
        """获取文件信息（缓存），若未命中缓存则同步计算（兜底用）"""
        if file_path in self._file_cache:
            return self._file_cache[file_path]
        
        info = {
            'path': file_path,
            'name': os.path.basename(file_path),
            'size': 0,
            'lines': 0,
            'filtered_lines': 0  # 过滤后的行数
        }
        
        try:
            if os.path.exists(file_path):
                info['size'] = os.path.getsize(file_path)
                try:
                    content = decode_content(file_path, self.encoding)
                    info['lines'] = len(content.splitlines())
                    language = get_language_by_extension(file_path)
                    filtered = filter_lines(
                        content, 
                        language, 
                        self.skip_blank_lines, 
                        self.skip_comment_lines, 
                        self.comment_chars
                    )
                    info['filtered_lines'] = len(filtered)
                except:
                    info['lines'] = 0
                    info['filtered_lines'] = 0
        except:
            pass
        
        self._file_cache[file_path] = info
        return info

    def _get_file_info_fast(self, file_path):
        """快速获取文件信息（无磁盘I/O，若无缓存返回默认空值）"""
        if file_path in self._file_cache:
            return self._file_cache[file_path]
        return {
            'path': file_path,
            'name': os.path.basename(file_path),
            'size': 0,
            'lines': 0,
            'filtered_lines': 0
        }
    
    def _sort_available_files(self):
        """根据选择的排序方式对可选文件列表进行排序"""
        sort_mode = self.sort_combo.currentText()
        
        # 收集所有可选的文件项 (包含隐藏的)
        items = []
        for i in range(self.available_list.count()):
            items.append(self.available_list.item(i).text())
        
        if not items:
            return
        
        # 根据排序模式排序
        if sort_mode == '默认顺序':
            file_indices = {file: idx for idx, file in enumerate(self.all_files)}
            items.sort(key=lambda x: file_indices.get(x, 999999))
        elif sort_mode == '按文件名':
            items.sort(key=lambda x: os.path.basename(x).lower())
        elif sort_mode == '按代码行数↓':
            # 降序：代码行数多的在前，若相同则看总行数，再按文件名排序
            items.sort(key=lambda x: (self._get_file_info_fast(x)['filtered_lines'], self._get_file_info_fast(x)['lines'], os.path.basename(x).lower()), reverse=True)
        elif sort_mode == '按代码行数↑':
            # 升序：代码行数少的在前，若相同则看总行数，再按文件名排序
            items.sort(key=lambda x: (self._get_file_info_fast(x)['filtered_lines'], self._get_file_info_fast(x)['lines'], os.path.basename(x).lower()))
        
        # 清空列表并重新添加排序后的项
        self.available_list.clear()
        for file_path in items:
            self.available_list.addItem(file_path)
        
        # 应用当前搜索过滤
        search_text = self.search_box.text()
        self._filter_available_files(search_text)
    
    def _init_file_lists(self, files, selected_files):
        """初始化左右两栏的文件列表"""
        selected_set = set(selected_files) if selected_files else set()
        
        for file in selected_files:
            if file in files:
                self.selected_list.addItem(file)
        
        for file in files:
            if file not in selected_set:
                self.available_list.addItem(file)
    
    def _add_selected(self):
        """将左栏选中的文件添加到右栏"""
        # 保存左栏滚动位置
        left_scroll_pos = self.available_list.verticalScrollBar().value()
        
        selected_items = self.available_list.selectedItems()
        for item in selected_items:
            row = self.available_list.row(item)
            self.available_list.takeItem(row)
            self.selected_list.addItem(item.text())
        
        self._update_counts()
        
        # 恢复左栏滚动位置(调整到合适的范围)
        max_scroll = self.available_list.verticalScrollBar().maximum()
        self.available_list.verticalScrollBar().setValue(min(left_scroll_pos, max_scroll))
    
    def _add_all(self):
        """将所有可见的可选文件添加到右栏"""
        items_to_add = [self.available_list.item(i).text() 
                       for i in range(self.available_list.count()) 
                       if not self.available_list.item(i).isHidden()]
        
        i = self.available_list.count() - 1
        while i >= 0:
            if not self.available_list.item(i).isHidden():
                self.available_list.takeItem(i)
            i -= 1
        
        for file_path in items_to_add:
            self.selected_list.addItem(file_path)
        
        self._update_counts()
        
        # 重置左栏滚动位置到顶部
        self.available_list.verticalScrollBar().setValue(0)
    
    def _remove_selected(self):
        """将右栏选中的文件移回左栏"""
        # 保存右栏滚动位置
        right_scroll_pos = self.selected_list.verticalScrollBar().value()
        
        selected_items = self.selected_list.selectedItems()
        items_to_add = [item.text() for item in selected_items]
        
        for item in selected_items:
            row = self.selected_list.row(item)
            self.selected_list.takeItem(row)
        
        # 添加到左栏并重新排序
        for file_path in items_to_add:
            self.available_list.addItem(file_path)
        
        # 应用当前排序
        self._sort_available_files()
        
        self._update_counts()
        
        # 恢复右栏滚动位置(调整到合适的范围)
        max_scroll = self.selected_list.verticalScrollBar().maximum()
        self.selected_list.verticalScrollBar().setValue(min(right_scroll_pos, max_scroll))
    
    def _remove_all(self):
        """将所有已选文件移回左栏"""
        items_to_add = []
        while self.selected_list.count() > 0:
            item = self.selected_list.takeItem(0)
            items_to_add.append(item.text())
        
        # 添加到左栏
        for file_path in items_to_add:
            self.available_list.addItem(file_path)
        
        # 应用当前排序
        self._sort_available_files()
        
        self._update_counts()
        
        # 重置右栏滚动位置到顶部
        self.selected_list.verticalScrollBar().setValue(0)
    
    def _move_to_top(self):
        """将选中的项移到顶部"""
        selected_items = self.selected_list.selectedItems()
        if not selected_items:
            return
        items_text = [item.text() for item in selected_items]
        for item in selected_items:
            self.selected_list.takeItem(self.selected_list.row(item))
        for i, text in enumerate(items_text):
            self.selected_list.insertItem(i, text)
            self.selected_list.item(i).setSelected(True)
        
        # 滚动到顶部
        self.selected_list.scrollToTop()
    
    def _move_up(self):
        """将选中的项上移一位"""
        selected_items = self.selected_list.selectedItems()
        if not selected_items:
            return
        
        # 保存滚动位置
        scroll_pos = self.selected_list.verticalScrollBar().value()
        
        rows = sorted([self.selected_list.row(item) for item in selected_items])
        for row in rows:
            if row > 0:
                item = self.selected_list.takeItem(row)
                self.selected_list.insertItem(row - 1, item)
                item.setSelected(True)
        
        # 恢复滚动位置
        self.selected_list.verticalScrollBar().setValue(scroll_pos)
    
    def _move_down(self):
        """将选中的项下移一位"""
        selected_items = self.selected_list.selectedItems()
        if not selected_items:
            return
        
        # 保存滚动位置
        scroll_pos = self.selected_list.verticalScrollBar().value()
        
        rows = sorted([self.selected_list.row(item) for item in selected_items], reverse=True)
        for row in rows:
            if row < self.selected_list.count() - 1:
                item = self.selected_list.takeItem(row)
                self.selected_list.insertItem(row + 1, item)
                item.setSelected(True)
        
        # 恢复滚动位置
        self.selected_list.verticalScrollBar().setValue(scroll_pos)
    
    def _move_to_bottom(self):
        """将选中的项移到底部"""
        selected_items = self.selected_list.selectedItems()
        if not selected_items:
            return
        items_text = [item.text() for item in selected_items]
        for item in selected_items:
            self.selected_list.takeItem(self.selected_list.row(item))
        for text in items_text:
            self.selected_list.addItem(text)
            self.selected_list.item(self.selected_list.count() - 1).setSelected(True)
        
        # 滚动到底部
        self.selected_list.scrollToBottom()
    
    def _filter_available_files(self, text):
        """根据搜索文本过滤左栏文件列表"""
        # 保存滚动位置
        scroll_pos = self.available_list.verticalScrollBar().value()
        
        for i in range(self.available_list.count()):
            item = self.available_list.item(i)
            item.setHidden(text.lower() not in item.text().lower())
        
        self._update_counts()
        
        # 如果有过滤文本,滚动到顶部;否则恢复滚动位置
        if text:
            self.available_list.verticalScrollBar().setValue(0)
        else:
            max_scroll = self.available_list.verticalScrollBar().maximum()
            self.available_list.verticalScrollBar().setValue(min(scroll_pos, max_scroll))
    
    def _on_available_double_click(self, item):
        """双击左栏文件，添加到右栏"""
        # 保存滚动位置
        scroll_pos = self.available_list.verticalScrollBar().value()
        
        row = self.available_list.row(item)
        self.available_list.takeItem(row)
        self.selected_list.addItem(item.text())
        self._update_counts()
        
        # 恢复滚动位置
        max_scroll = self.available_list.verticalScrollBar().maximum()
        self.available_list.verticalScrollBar().setValue(min(scroll_pos, max_scroll))
    
    def _on_selected_double_click(self, item):
        """双击右栏文件，移回左栏"""
        # 保存滚动位置
        scroll_pos = self.selected_list.verticalScrollBar().value()
        
        row = self.selected_list.row(item)
        file_path = item.text()
        self.selected_list.takeItem(row)
        
        # 添加到左栏并重新排序
        self.available_list.addItem(file_path)
        self._sort_available_files()
        
        self._update_counts()
        
        # 恢复滚动位置
        max_scroll = self.selected_list.verticalScrollBar().maximum()
        self.selected_list.verticalScrollBar().setValue(min(scroll_pos, max_scroll))
    
    def _on_file_clicked(self, item):
        """单击文件时显示预览"""
        if item:
            self._preview_file(item.text())
    
    def _refresh_preview(self):
        """刷新当前预览"""
        if self._current_preview_file:
            self._preview_file(self._current_preview_file)
    
    def _preview_file(self, file_path):
        """预览文件内容"""
        self._current_preview_file = file_path
        
        if not os.path.exists(file_path):
            self.preview_text.setText('文件不存在: {}'.format(file_path))
            self.preview_info_label.setText('')
            return
        
        try:
            # 获取文件信息
            file_info = self._get_file_info(file_path)
            file_size = file_info['size']
            total_lines = file_info['lines']
            size_str = format_file_size(file_size)
            
            max_lines = int(self.preview_lines_combo.currentText())
            content = decode_content(file_path, self.encoding)
            
            all_lines = content.splitlines()
            
            preview_lines = all_lines[:max_lines]
            preview_content = '\n'.join(preview_lines)
            
            if total_lines > max_lines:
                preview_content += '\n\n... (省略 {} 行，共 {} 行)'.format(
                    total_lines - max_lines, total_lines
                )
            
            self.preview_text.setText(preview_content)
            
            file_name = os.path.basename(file_path)
            language = get_language_by_extension(file_path)
            lang_str = ' | 语言: {}'.format(language) if language else ''
            
            # 显示文件名、大小、行数、语言
            self.preview_info_label.setText('文件: {} | 大小: {} | 行数: {}{}'.format(
                file_name, size_str, total_lines, lang_str
            ))
            
        except Exception as e:
            self.preview_text.setText('无法预览文件:\n{}'.format(str(e)))
            self.preview_info_label.setText('预览失败')
    
    def _update_counts(self):
        """更新统计信息"""
        available_visible = sum(1 for i in range(self.available_list.count()) 
                               if not self.available_list.item(i).isHidden())
        self.available_count_label.setText('可选: {} 项'.format(available_visible))
        
        selected_count = self.selected_list.count()
        
        # 计算已选文件的总行数和过滤后行数
        total_lines = 0
        total_filtered_lines = 0
        for i in range(selected_count):
            file_path = self.selected_list.item(i).text()
            info = self._get_file_info_fast(file_path)
            total_lines += info['lines']
            total_filtered_lines += info['filtered_lines']
        
        if total_lines > 0:
            # 显示：已选文件数 | 总行数 | 过滤后行数
            self.selected_count_label.setText(
                '已选: {} 项 | 总计: {} 行 | 过滤后: {} 行'.format(
                    selected_count, total_lines, total_filtered_lines
                )
            )
        else:
            self.selected_count_label.setText('已选: {} 项'.format(selected_count))
    
    def get_selected(self):
        """返回选中的文件列表，按照右栏顺序"""
        return [self.selected_list.item(i).text() for i in range(self.selected_list.count())]
    
    def get_all_ordered(self):
        """返回所有文件的当前顺序（已选文件在前）"""
        ordered = self.get_selected()
        ordered.extend([self.available_list.item(i).text() for i in range(self.available_list.count())])
        return ordered



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
        
        type_label = BodyLabel('日志类型:')
        self.log_type_combo = ComboBox()
        self.log_type_combo.addItems(['全部日志', '错误日志'])
        self.log_type_combo.currentTextChanged.connect(self._load_log)
        type_layout.addWidget(type_label)
        type_layout.addWidget(self.log_type_combo)
        type_layout.addStretch(1)
        
        self.refresh_btn = PushButton('刷新')
        self.refresh_btn.clicked.connect(self._load_log)
        type_layout.addWidget(self.refresh_btn)
        
        self.open_dir_btn = PushButton('打开日志目录')
        self.open_dir_btn.clicked.connect(self._open_log_dir)
        type_layout.addWidget(self.open_dir_btn)
        
        self.clear_btn = PushButton('清空当前日志')
        self.clear_btn.clicked.connect(self._clear_log)
        type_layout.addWidget(self.clear_btn)
        
        layout.addWidget(type_row)
        
        # 日志信息
        self.log_info_label = CaptionLabel('')
        self.log_info_label.setStyleSheet('font-size: 11px;')
        layout.addWidget(self.log_info_label)
        
        # 日志内容
        self.log_text = TextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText('暂无日志...')
        
        from PyQt5.QtGui import QFont
        log_font = QFont('Consolas', 9)
        if not log_font.exactMatch():
            log_font = QFont('Courier New', 9)
        self.log_text.setFont(log_font)
        layout.addWidget(self.log_text, 1)
        
        # 底部按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.accept)
        layout.addWidget(buttons)
        
        self._load_log()
    
    def _get_current_log_file(self):
        """获取当前选择的日志文件路径"""
        logger_instance = Logger()
        if self.log_type_combo.currentText() == '错误日志':
            return logger_instance.error_log_file
        else:
            return logger_instance.log_file
    
    def _load_log(self):
        """加载日志内容"""
        log_file = self._get_current_log_file()
        
        if not os.path.exists(log_file):
            self.log_text.setText('日志文件不存在')
            self.log_info_label.setText(f'文件: {log_file}')
            return
        
        try:
            file_size = os.path.getsize(log_file)
            size_str = format_file_size(file_size)
            
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
            w = MessageBox('目录不存在', f'日志目录不存在:\n{log_dir}', self)
            w.cancelButton.hide()
            w.exec_()
            w.hide()
            w.close()
    
    def _clear_log(self):
        """清空当前日志文件"""
        w = MessageBox('确认清空', '确定要清空当前日志文件吗？此操作不可恢复。', self)
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        
        res = w.exec_()
        w.hide()
        w.close()
        
        if res:
            log_file = self._get_current_log_file()
            try:
                with open(log_file, 'w', encoding='utf-8') as f:
                    f.write('')
                self._load_log()
                w_info = MessageBox('成功', '日志已清空', self)
                w_info.cancelButton.hide()
                w_info.exec_()
                w_info.hide()
                w_info.close()
            except Exception as e:
                w_err = MessageBox('失败', f'清空日志失败:\n{str(e)}', self)
                w_err.cancelButton.hide()
                w_err.exec_()
                w_err.hide()
                w_err.close()
