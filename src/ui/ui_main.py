# -*- coding: utf-8 -*-
"""
UI主窗口模块
包含主界面窗口和启动函数
"""
import os
import sys

# 抑制 libpng 警告
os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.qpa.*=false'

from PyQt5.QtCore import QEvent, Qt, QTimer, QUrl
from PyQt5.QtGui import QDesktopServices, QIcon, QFont
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QFileDialog, QMainWindow, QDialog, QMessageBox
)
from qfluentwidgets import (
    PrimaryPushButton, PushButton, LineEdit, TextEdit,
    DoubleSpinBox, ComboBox, CheckBox, InfoBar, InfoBarPosition,
    TitleLabel, BodyLabel, CaptionLabel, StrongBodyLabel, CardWidget,
    setTheme, Theme, isDarkTheme, ScrollArea, MessageBox
)

from src.core.code_processor import (
    normalize_items, normalize_exts, normalize_paths, read_gitignore_excludes,
    DEFAULT_SKIP_DIRS, DEFAULT_SKIP_FILES, LANGUAGE_BY_EXT
)
from src.ui.ui_tasks import GenerateWorker, ExtensionScanWorker
from src.ui.ui_dialogs import ExtensionSelectDialog, CommentPrefixDialog, FileSelectDialog, LogViewDialog, COMMENT_PREFIX_BY_LANG
from src.utils.logger import get_logger
from src.utils.theme_utils import set_window_dark_title_bar

class GeneratorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('软著源代码文档生成器')
        self.resize(1020, 720)
        
        # 设置窗口图标
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        icon_path = os.path.join(project_root, 'assets', 'logo.ico')
        if os.path.exists(icon_path):
            from PyQt5.QtGui import QPixmap, QPainter, QPainterPath
            pixmap = QPixmap(icon_path)
            if not pixmap.isNull():
                size = min(pixmap.width(), pixmap.height())
                # 创建一个支持透明通道的目标图像
                circular_pixmap = QPixmap(size, size)
                circular_pixmap.fill(Qt.transparent)
                
                # 绘制圆形裁剪后的图像
                painter = QPainter(circular_pixmap)
                painter.setRenderHint(QPainter.Antialiasing)
                
                # 边缘收缩 2%，切除外围的白色边缘
                margin = max(2, int(size * 0.02))
                path = QPainterPath()
                path.addEllipse(margin, margin, size - 2 * margin, size - 2 * margin)
                
                painter.setClipPath(path)
                painter.drawPixmap(0, 0, pixmap)
                painter.end()
                
                self.setWindowIcon(QIcon(circular_pixmap))
            else:
                self.setWindowIcon(QIcon(icon_path))
        self.worker = None
        self.ext_worker = None
        self.pending_ext_scan = False
        self.available_exts = []
        self.available_files = []  # 存储可用文件列表
        self.selected_files = []   # 存储用户选择的文件列表
        self.file_order = []       # 存储文件顺序
        self.last_scan_count = 0
        self.last_stats = None  # 存储上次统计的详细信息
        self.ext_scan_timer = QTimer(self)
        self.ext_scan_timer.setSingleShot(True)
        self.ext_scan_timer.timeout.connect(self.start_extension_scan)
        
        # 初始化日志
        self.logger = get_logger('GUI')
        self.logger.info('=' * 50)
        self.logger.info('软著源代码文档生成器启动')
        self.logger.info('=' * 50)
        
        self._build_ui()
        setTheme(Theme.AUTO)
        self.apply_theme_stylesheet()

    def _build_ui(self):
        wrapper = QWidget()
        wrapper.setObjectName("centralWrapper")
        wrapper_layout = QHBoxLayout(wrapper)
        wrapper_layout.setContentsMargins(16, 16, 16, 16)
        wrapper_layout.setSpacing(18)
        self.setCentralWidget(wrapper)

        scroll = ScrollArea()
        scroll.setWidgetResizable(True)
        scroll.enableTransparentBackground()
        self.scroll_area = scroll
        container = QWidget()
        container.setObjectName("scrollContainer")
        scroll.setWidget(container)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(16)

        title = TitleLabel('软件著作权源代码文档生成器')
        subtitle = CaptionLabel('支持多目录扫描、注释过滤、模板与排版参数配置')
        title.setFont(QFont('Microsoft YaHei', 18, QFont.Bold))
        subtitle.setStyleSheet('font-size: 12px;')
        subtitle.setWordWrap(True)
        
        self.theme_label = BodyLabel('主题模式')
        self.theme_combo = ComboBox()
        self.theme_combo.addItems(['系统自适应', '浅色模式', '深色模式'])
        self.theme_combo.blockSignals(True)
        self.theme_combo.setCurrentText('系统自适应')
        self.theme_combo.blockSignals(False)
        self.theme_combo.currentTextChanged.connect(self._on_theme_changed)

        header_row = QWidget()
        header_layout = QHBoxLayout(header_row)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.addWidget(title)
        header_layout.addStretch(1)
        header_layout.addWidget(self.theme_label)
        header_layout.addWidget(self.theme_combo)
        layout.addWidget(header_row)
        layout.addWidget(subtitle)

        right_column = QWidget()
        right_column.setMinimumWidth(340)
        right_layout = QVBoxLayout(right_column)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(16)

        source_group, source_layout = self._create_group('源码设置')
        source_grid = QGridLayout()
        source_grid.setHorizontalSpacing(12)
        source_grid.setVerticalSpacing(12)
        source_grid.setColumnStretch(1, 1)
        source_grid.setColumnMinimumWidth(0, 110)
        source_layout.addLayout(source_grid)

        self.indirs_edit = TextEdit()
        self.indirs_edit.setPlaceholderText('每行一个目录')
        self.indirs_edit.setText('')
        self.indirs_edit.setMinimumHeight(90)
        indir_buttons = QHBoxLayout()
        indir_buttons.setSpacing(8)
        self.add_indir_btn = PushButton('添加目录')
        self.clear_indir_btn = PushButton('清空')
        self.add_indir_btn.setMinimumHeight(32)
        self.clear_indir_btn.setMinimumHeight(32)
        indir_buttons.addWidget(self.add_indir_btn)
        indir_buttons.addWidget(self.clear_indir_btn)
        indir_buttons.addStretch(1)
        indir_container = QWidget()
        indir_container_layout = QVBoxLayout(indir_container)
        indir_container_layout.setContentsMargins(0, 0, 0, 0)
        indir_container_layout.addWidget(self.indirs_edit)
        indir_container_layout.addLayout(indir_buttons)
        source_grid.addWidget(BodyLabel('源码目录'), 0, 0)
        source_grid.addWidget(indir_container, 0, 1)

        self.exts_edit = LineEdit()
        self.exts_edit.setPlaceholderText('如 py, java, cpp')
        self.exts_edit.setText('')
        self.exts_edit.setMinimumHeight(32)
        self.exts_select_btn = PushButton('选择')
        self.exts_select_btn.setMinimumHeight(32)
        exts_row = QWidget()
        exts_layout = QHBoxLayout(exts_row)
        exts_layout.setContentsMargins(0, 0, 0, 0)
        exts_layout.setSpacing(8)
        exts_layout.addWidget(self.exts_edit, 1)
        exts_layout.addWidget(self.exts_select_btn)
        source_grid.addWidget(BodyLabel('文件后缀'), 1, 0)
        source_grid.addWidget(exts_row, 1, 1)

        self.comment_chars_edit = LineEdit()
        self.comment_chars_edit.setPlaceholderText('如 #, //')
        self.comment_chars_edit.setText('')
        self.comment_chars_edit.setMinimumHeight(32)
        self.comment_select_btn = PushButton('选择')
        self.comment_select_btn.setMinimumHeight(32)
        comment_row = QWidget()
        comment_layout = QHBoxLayout(comment_row)
        comment_layout.setContentsMargins(0, 0, 0, 0)
        comment_layout.setSpacing(8)
        comment_layout.addWidget(self.comment_chars_edit, 1)
        comment_layout.addWidget(self.comment_select_btn)
        source_grid.addWidget(BodyLabel('注释前缀'), 2, 0)
        source_grid.addWidget(comment_row, 2, 1)

        self.excludes_edit = TextEdit()
        self.excludes_edit.setPlaceholderText('每行一个排除路径')
        self.excludes_edit.setMinimumHeight(90)
        exclude_buttons = QHBoxLayout()
        exclude_buttons.setSpacing(8)
        self.add_exclude_dir_btn = PushButton('添加排除目录')
        self.add_exclude_file_btn = PushButton('添加排除文件')
        self.read_gitignore_btn = PushButton('读取.gitignore')
        self.clear_exclude_btn = PushButton('清空')
        self.add_exclude_dir_btn.setMinimumHeight(32)
        self.add_exclude_file_btn.setMinimumHeight(32)
        self.read_gitignore_btn.setMinimumHeight(32)
        self.clear_exclude_btn.setMinimumHeight(32)
        exclude_buttons.addWidget(self.add_exclude_dir_btn)
        exclude_buttons.addWidget(self.add_exclude_file_btn)
        exclude_buttons.addWidget(self.read_gitignore_btn)
        exclude_buttons.addWidget(self.clear_exclude_btn)
        exclude_buttons.addStretch(1)
        exclude_container = QWidget()
        exclude_container_layout = QVBoxLayout(exclude_container)
        exclude_container_layout.setContentsMargins(0, 0, 0, 0)
        exclude_container_layout.addWidget(self.excludes_edit)
        exclude_container_layout.addLayout(exclude_buttons)
        source_grid.addWidget(BodyLabel('排除路径'), 3, 0)
        source_grid.addWidget(exclude_container, 3, 1)

        self.encoding_combo = ComboBox()
        self.encoding_combo.addItems(['自动', 'utf-8', 'utf-8-sig', 'gbk', 'gb18030'])
        self.encoding_combo.setMinimumHeight(32)
        source_grid.addWidget(BodyLabel('文件编码'), 4, 0)
        source_grid.addWidget(self.encoding_combo, 4, 1)

        self.skip_blank_check = CheckBox('过滤空行')
        self.skip_blank_check.setChecked(True)
        self.skip_comment_check = CheckBox('过滤注释')
        self.skip_comment_check.setChecked(True)
        check_row = QWidget()
        check_layout = QHBoxLayout(check_row)
        check_layout.setContentsMargins(0, 0, 0, 0)
        check_layout.addWidget(self.skip_blank_check)
        check_layout.addWidget(self.skip_comment_check)
        check_layout.addStretch(1)
        source_grid.addWidget(BodyLabel('过滤规则'), 5, 0)
        source_grid.addWidget(check_row, 5, 1)

        layout.addWidget(source_group)

        output_group, output_layout = self._create_group('文档设置')
        output_grid = QGridLayout()
        output_grid.setHorizontalSpacing(12)
        output_grid.setVerticalSpacing(12)
        output_grid.setColumnStretch(1, 1)
        output_grid.setColumnMinimumWidth(0, 110)
        output_layout.addLayout(output_grid)

        self.title_edit = LineEdit()
        self.title_edit.setText('软件著作权程序鉴别材料生成器V1.0')
        self.title_edit.setMinimumHeight(32)
        
        self.title_align_combo = ComboBox()
        self.title_align_combo.addItems(['左对齐', '居中', '右对齐'])
        self.title_align_combo.setCurrentText('左对齐')
        self.title_align_combo.setMinimumHeight(32)
        self.title_align_combo.setMinimumWidth(100)
        
        title_row = QWidget()
        title_layout = QHBoxLayout(title_row)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(8)
        title_layout.addWidget(self.title_edit, 1)
        title_layout.addWidget(self.title_align_combo)
        
        output_grid.addWidget(BodyLabel('页眉标题'), 0, 0)
        output_grid.addWidget(title_row, 0, 1)

        self.outfile_edit = LineEdit()
        self.outfile_edit.setText(os.path.abspath('code.docx'))
        self.outfile_edit.setPlaceholderText('如 D:\\输出\\code.docx')
        self.outfile_edit.setMinimumHeight(32)
        self.outfile_btn = PushButton('选择输出文件')
        self.outfile_btn.setMinimumHeight(32)
        outfile_row = QWidget()
        outfile_layout = QHBoxLayout(outfile_row)
        outfile_layout.setContentsMargins(0, 0, 0, 0)
        outfile_layout.setSpacing(8)
        outfile_layout.addWidget(self.outfile_edit)
        outfile_layout.addWidget(self.outfile_btn)
        output_grid.addWidget(BodyLabel('输出路径'), 1, 0)
        output_grid.addWidget(outfile_row, 1, 1)

        self.template_edit = LineEdit()
        self.template_edit.setPlaceholderText('可选，留空使用默认模板')
        self.template_edit.setMinimumHeight(32)
        self.template_btn = PushButton('选择模板')
        self.template_clear_btn = PushButton('清空')
        self.template_btn.setMinimumHeight(32)
        self.template_clear_btn.setMinimumHeight(32)
        template_row = QWidget()
        template_layout = QHBoxLayout(template_row)
        template_layout.setContentsMargins(0, 0, 0, 0)
        template_layout.setSpacing(8)
        template_layout.addWidget(self.template_edit)
        template_layout.addWidget(self.template_btn)
        template_layout.addWidget(self.template_clear_btn)
        output_grid.addWidget(BodyLabel('模板文件'), 2, 0)
        output_grid.addWidget(template_row, 2, 1)

        layout.addWidget(output_group)

        # 排版设置 - 使用分组优化布局
        style_group, style_layout = self._create_group('排版设置')
        
        # === 字体设置 ===
        font_section = StrongBodyLabel('字体设置')
        font_section.setStyleSheet('font-size: 12px; margin-top: 4px;')
        style_layout.addWidget(font_section)
        
        font_grid = QGridLayout()
        font_grid.setHorizontalSpacing(12)
        font_grid.setVerticalSpacing(10)
        font_grid.setColumnStretch(1, 1)
        font_grid.setColumnMinimumWidth(0, 120)
        style_layout.addLayout(font_grid)

        self.font_name_edit = ComboBox()
        self.font_name_edit.addItems(['Arial', 'Consolas', 'Courier New', 'Times New Roman', '仿宋', '黑体', '楷体', '宋体', '微软雅黑'])
        self.font_name_edit.setCurrentText('Consolas')
        self.font_name_edit.setMinimumHeight(32)
        font_grid.addWidget(BodyLabel('正文字体'), 0, 0)
        font_grid.addWidget(self.font_name_edit, 0, 1)

        self.font_size_spin = DoubleSpinBox()
        self.font_size_spin.setRange(1.0, 72.0)
        self.font_size_spin.setValue(10.5)
        self.font_size_spin.setMinimumHeight(32)
        font_grid.addWidget(BodyLabel('正文字号'), 1, 0)
        font_grid.addWidget(self.font_size_spin, 1, 1)

        self.header_font_combo = ComboBox()
        self.header_font_combo.addItems(['Arial', 'Times New Roman', '仿宋', '黑体', '楷体', '宋体', '微软雅黑'])
        self.header_font_combo.setCurrentText('宋体')
        self.header_font_combo.setMinimumHeight(32)
        font_grid.addWidget(BodyLabel('页眉页脚字体'), 2, 0)
        font_grid.addWidget(self.header_font_combo, 2, 1)

        self.header_font_size_spin = DoubleSpinBox()
        self.header_font_size_spin.setRange(1.0, 72.0)
        self.header_font_size_spin.setValue(10.5)
        self.header_font_size_spin.setMinimumHeight(32)
        font_grid.addWidget(BodyLabel('页眉页脚字号'), 3, 0)
        font_grid.addWidget(self.header_font_size_spin, 3, 1)

        # 分隔线
        from qfluentwidgets import HorizontalSeparator
        style_layout.addSpacing(8)
        style_layout.addWidget(HorizontalSeparator())
        style_layout.addSpacing(8)

        # === 段落设置 ===
        para_section = StrongBodyLabel('段落设置')
        para_section.setStyleSheet('font-size: 12px;')
        style_layout.addWidget(para_section)
        
        para_grid = QGridLayout()
        para_grid.setHorizontalSpacing(12)
        para_grid.setVerticalSpacing(10)
        para_grid.setColumnStretch(1, 1)
        para_grid.setColumnMinimumWidth(0, 120)
        style_layout.addLayout(para_grid)

        self.space_before_spin = DoubleSpinBox()
        self.space_before_spin.setRange(0.0, 72.0)
        self.space_before_spin.setValue(0.0)
        self.space_before_spin.setMinimumHeight(32)
        para_grid.addWidget(BodyLabel('段前间距(pt)'), 0, 0)
        para_grid.addWidget(self.space_before_spin, 0, 1)

        self.space_after_spin = DoubleSpinBox()
        self.space_after_spin.setRange(0.0, 72.0)
        self.space_after_spin.setValue(2.3)
        self.space_after_spin.setMinimumHeight(32)
        para_grid.addWidget(BodyLabel('段后间距(pt)'), 1, 0)
        para_grid.addWidget(self.space_after_spin, 1, 1)

        self.line_spacing_spin = DoubleSpinBox()
        self.line_spacing_spin.setRange(0.0, 72.0)
        self.line_spacing_spin.setValue(10.5)
        self.line_spacing_spin.setMinimumHeight(32)
        para_grid.addWidget(BodyLabel('行距(pt)'), 2, 0)
        para_grid.addWidget(self.line_spacing_spin, 2, 1)

        # 分隔线
        style_layout.addSpacing(8)
        style_layout.addWidget(HorizontalSeparator())
        style_layout.addSpacing(8)

        # === 页面边距 ===
        margin_section = StrongBodyLabel('页面边距')
        margin_section.setStyleSheet('font-size: 12px;')
        style_layout.addWidget(margin_section)
        
        margin_grid = QGridLayout()
        margin_grid.setHorizontalSpacing(12)
        margin_grid.setVerticalSpacing(10)
        margin_grid.setColumnStretch(1, 1)
        margin_grid.setColumnStretch(3, 1)
        margin_grid.setColumnMinimumWidth(0, 60)
        margin_grid.setColumnMinimumWidth(2, 60)
        style_layout.addLayout(margin_grid)

        self.top_margin_spin = DoubleSpinBox()
        self.top_margin_spin.setRange(0.0, 10.0)
        self.top_margin_spin.setValue(2.5)
        self.top_margin_spin.setSingleStep(0.1)
        self.top_margin_spin.setDecimals(2)
        self.top_margin_spin.setSuffix(' cm')
        self.top_margin_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('上边距'), 0, 0)
        margin_grid.addWidget(self.top_margin_spin, 0, 1)

        self.bottom_margin_spin = DoubleSpinBox()
        self.bottom_margin_spin.setRange(0.0, 10.0)
        self.bottom_margin_spin.setValue(2.5)
        self.bottom_margin_spin.setSingleStep(0.1)
        self.bottom_margin_spin.setDecimals(2)
        self.bottom_margin_spin.setSuffix(' cm')
        self.bottom_margin_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('下边距'), 0, 2)
        margin_grid.addWidget(self.bottom_margin_spin, 0, 3)

        self.left_margin_spin = DoubleSpinBox()
        self.left_margin_spin.setRange(0.0, 10.0)
        self.left_margin_spin.setValue(2.5)
        self.left_margin_spin.setSingleStep(0.1)
        self.left_margin_spin.setDecimals(2)
        self.left_margin_spin.setSuffix(' cm')
        self.left_margin_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('左边距'), 1, 0)
        margin_grid.addWidget(self.left_margin_spin, 1, 1)

        self.right_margin_spin = DoubleSpinBox()
        self.right_margin_spin.setRange(0.0, 10.0)
        self.right_margin_spin.setValue(2.5)
        self.right_margin_spin.setSingleStep(0.1)
        self.right_margin_spin.setDecimals(2)
        self.right_margin_spin.setSuffix(' cm')
        self.right_margin_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('右边距'), 1, 2)
        margin_grid.addWidget(self.right_margin_spin, 1, 3)

        self.header_distance_spin = DoubleSpinBox()
        self.header_distance_spin.setRange(0.0, 10.0)
        self.header_distance_spin.setValue(1.5)
        self.header_distance_spin.setSingleStep(0.1)
        self.header_distance_spin.setDecimals(2)
        self.header_distance_spin.setSuffix(' cm')
        self.header_distance_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('页眉距离'), 2, 0)
        margin_grid.addWidget(self.header_distance_spin, 2, 1)

        self.footer_distance_spin = DoubleSpinBox()
        self.footer_distance_spin.setRange(0.0, 10.0)
        self.footer_distance_spin.setValue(1.75)
        self.footer_distance_spin.setSingleStep(0.1)
        self.footer_distance_spin.setDecimals(2)
        self.footer_distance_spin.setSuffix(' cm')
        self.footer_distance_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('页脚距离'), 2, 2)
        margin_grid.addWidget(self.footer_distance_spin, 2, 3)

        # 分隔线
        style_layout.addSpacing(8)
        style_layout.addWidget(HorizontalSeparator())
        style_layout.addSpacing(8)

        # === 页码格式 ===
        page_section = StrongBodyLabel('页码格式')
        page_section.setStyleSheet('font-size: 12px;')
        style_layout.addWidget(page_section)
        
        page_grid = QGridLayout()
        page_grid.setHorizontalSpacing(12)
        page_grid.setVerticalSpacing(10)
        page_grid.setColumnStretch(1, 1)
        page_grid.setColumnMinimumWidth(0, 120)
        style_layout.addLayout(page_grid)

        # 页码格式
        self.page_number_combo = ComboBox()
        self.page_number_combo.addItems([
            '{page}',
            '第 {page} 页 共 {total} 页',
            '{page} / {total}',
            'Page {page} of {total}',
            '自定义...'
        ])
        self.page_number_combo.setMinimumHeight(32)
        self.page_number_combo.currentTextChanged.connect(self._on_page_number_format_changed)
        page_grid.addWidget(BodyLabel('格式'), 0, 0)
        page_grid.addWidget(self.page_number_combo, 0, 1)

        # 自定义页码格式
        self.custom_page_number_edit = LineEdit()
        self.custom_page_number_edit.setPlaceholderText('使用 {page} 和 {total} 作为占位符')
        self.custom_page_number_edit.setMinimumHeight(32)
        self.custom_page_number_edit.setVisible(False)
        page_grid.addWidget(self.custom_page_number_edit, 1, 1)

        # 页码位置
        self.page_loc_combo = ComboBox()
        self.page_loc_combo.addItems(['页眉', '页脚'])
        self.page_loc_combo.setCurrentText('页眉')
        self.page_loc_combo.setMinimumHeight(32)
        page_grid.addWidget(BodyLabel('页码位置'), 2, 0)
        page_grid.addWidget(self.page_loc_combo, 2, 1)

        # 页码对齐
        self.page_align_combo = ComboBox()
        self.page_align_combo.addItems(['左对齐', '居中', '右对齐'])
        self.page_align_combo.setCurrentText('右对齐')
        self.page_align_combo.setMinimumHeight(32)
        page_grid.addWidget(BodyLabel('页码对齐'), 3, 0)
        page_grid.addWidget(self.page_align_combo, 3, 1)

        # 恢复默认按钮
        style_layout.addSpacing(8)
        style_btn_row = QWidget()
        style_btn_layout = QHBoxLayout(style_btn_row)
        style_btn_layout.setContentsMargins(0, 0, 0, 0)
        style_btn_layout.setSpacing(8)
        self.reset_style_btn = PushButton('恢复默认')
        self.reset_style_btn.setMinimumHeight(32)
        style_btn_layout.addWidget(self.reset_style_btn)
        style_btn_layout.addStretch(1)
        style_layout.addWidget(style_btn_row)

        layout.addWidget(style_group)

        summary_group, summary_layout = self._create_group('运行概览')
        self.summary_title = StrongBodyLabel('尚未执行任务')
        self.summary_title.setStyleSheet('font-size: 12px;')
        self.summary_body = CaptionLabel('')
        self.summary_body.setWordWrap(True)
        self.summary_body.setStyleSheet('font-size: 12px;')
        summary_layout.addWidget(self.summary_title)
        summary_layout.addWidget(self.summary_body)
        right_layout.addWidget(summary_group)

        action_group, action_layout = self._create_group('操作区')
        action_row = QWidget()
        action_row_layout = QVBoxLayout(action_row)
        action_row_layout.setContentsMargins(0, 0, 0, 0)
        action_row_layout.setSpacing(8)
        
        # 60页模式开关
        self.pages_60_check = CheckBox('启用60页模式（前30+后30）')
        self.pages_60_check.setChecked(False)
        action_row_layout.addWidget(self.pages_60_check)
        
        # 导出PDF开关
        self.export_pdf_check = CheckBox('同时导出 PDF')
        self.export_pdf_check.setChecked(False)
        action_row_layout.addWidget(self.export_pdf_check)
        
        self.generate_btn = PrimaryPushButton('生成文档')
        self.scan_btn = PushButton('扫描文件数')
        self.select_files_btn = PushButton('选择文件')
        self.count_lines_btn = PushButton('统计代码行数')
        self.open_output_btn = PushButton('打开输出目录')
        self.view_log_btn = PushButton('查看日志')
        self.generate_btn.setMinimumHeight(36)
        self.scan_btn.setMinimumHeight(34)
        self.select_files_btn.setMinimumHeight(34)
        self.count_lines_btn.setMinimumHeight(34)
        self.open_output_btn.setMinimumHeight(34)
        self.view_log_btn.setMinimumHeight(34)
        action_row_layout.addWidget(self.generate_btn)
        action_row_layout.addWidget(self.scan_btn)
        action_row_layout.addWidget(self.select_files_btn)
        action_row_layout.addWidget(self.count_lines_btn)
        action_row_layout.addWidget(self.open_output_btn)
        action_row_layout.addWidget(self.view_log_btn)
        action_layout.addWidget(action_row)
        self.status_label = CaptionLabel('')
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet('font-size: 12px;')
        action_layout.addWidget(self.status_label)
        right_layout.addWidget(action_group)
        right_layout.addStretch(1)

        layout.addStretch(1)
        wrapper_layout.addWidget(scroll, 1)
        wrapper_layout.addWidget(right_column)

        self.add_indir_btn.clicked.connect(self.add_indir)
        self.clear_indir_btn.clicked.connect(lambda: self.indirs_edit.setText(''))
        self.add_exclude_dir_btn.clicked.connect(self.add_exclude_dir)
        self.add_exclude_file_btn.clicked.connect(self.add_exclude_file)
        self.read_gitignore_btn.clicked.connect(self.load_gitignore_excludes)
        self.clear_exclude_btn.clicked.connect(lambda: self.excludes_edit.setText(''))
        self.outfile_btn.clicked.connect(self.choose_outfile)
        self.template_btn.clicked.connect(self.choose_template)
        self.template_clear_btn.clicked.connect(lambda: self.template_edit.setText(''))
        self.reset_style_btn.clicked.connect(self.reset_style_defaults)
        self.generate_btn.clicked.connect(lambda: self.start_worker('generate'))
        self.scan_btn.clicked.connect(lambda: self.start_worker('scan'))
        self.select_files_btn.clicked.connect(self.open_file_select_dialog)
        self.count_lines_btn.clicked.connect(lambda: self.start_worker('count'))
        self.open_output_btn.clicked.connect(self.open_output_dir)
        self.view_log_btn.clicked.connect(self.view_logs)
        self.outfile_edit.textChanged.connect(self._update_open_output_enabled)
        self.outfile_edit.textChanged.connect(self._update_summary)
        self.indirs_edit.textChanged.connect(self._update_summary)
        self.indirs_edit.textChanged.connect(self.schedule_extension_scan)
        self.exts_edit.textChanged.connect(self._update_summary)
        self.comment_chars_edit.textChanged.connect(self._update_summary)
        self.excludes_edit.textChanged.connect(self._update_summary)
        self.encoding_combo.currentTextChanged.connect(self._update_summary)
        self.skip_blank_check.toggled.connect(self._update_summary)
        self.skip_comment_check.toggled.connect(self._update_summary)
        self.pages_60_check.toggled.connect(self._update_summary)
        self.title_align_combo.currentTextChanged.connect(lambda: self._validate_header_footer_positions('title_align'))
        self.title_align_combo.currentTextChanged.connect(self._update_summary)
        self.page_loc_combo.currentTextChanged.connect(lambda: self._validate_header_footer_positions('page_loc'))
        self.page_loc_combo.currentTextChanged.connect(self._update_summary)
        self.page_align_combo.currentTextChanged.connect(lambda: self._validate_header_footer_positions('page_align'))
        self.page_align_combo.currentTextChanged.connect(self._update_summary)
        self.exts_select_btn.clicked.connect(self.open_extension_dialog)
        self.comment_select_btn.clicked.connect(self.open_comment_prefix_dialog)
        self._update_open_output_enabled()
        self._update_summary()
        for widget in (right_column, summary_group, action_group, self.status_label, self.indirs_edit, self.excludes_edit):
            widget.installEventFilter(self)
    
    def _on_page_number_format_changed(self, text):
        """当页码格式改变时显示/隐藏自定义输入框"""
        if text == '自定义...':
            self.custom_page_number_edit.setVisible(True)
        else:
            self.custom_page_number_edit.setVisible(False)

    def _validate_header_footer_positions(self, changed_source):
        """
        验证和自动调整页眉标题和页码位置，防止重叠
        changed_source 可以是 'title_align', 'page_loc', 'page_align'
        """
        if self.page_loc_combo.currentText() == '页脚':
            return
            
        t_align = self.title_align_combo.currentText()
        p_align = self.page_align_combo.currentText()
        
        if t_align == p_align:
            # 冲突了，需要调整
            aligns = ['左对齐', '居中', '右对齐']
            aligns.remove(t_align)
            
            if changed_source == 'title_align':
                # 用户改了标题，调整页码
                next_align = '右对齐' if '右对齐' in aligns else aligns[0]
                self.page_align_combo.blockSignals(True)
                self.page_align_combo.setCurrentText(next_align)
                self.page_align_combo.blockSignals(False)
                self._notify('warning', '对齐方式冲突', f'标题和页码不能在同一位置，已自动将页码调整为【{next_align}】')
            else:
                # 用户改了页码，调整标题
                next_align = '左对齐' if '左对齐' in aligns else aligns[0]
                self.title_align_combo.blockSignals(True)
                self.title_align_combo.setCurrentText(next_align)
                self.title_align_combo.blockSignals(False)
                self._notify('warning', '对齐方式冲突', f'标题和页码不能在同一位置，已自动将标题调整为【{next_align}】')

    def _create_group(self, title):
        group = CardWidget()
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        header = StrongBodyLabel(title)
        header.setStyleSheet('font-size: 13px;')
        layout.addWidget(header)
        return group, layout

    def add_indir(self):
        directory = QFileDialog.getExistingDirectory(self, '选择源码目录', os.getcwd())
        if directory:
            self._append_line(self.indirs_edit, directory)

    def add_exclude_dir(self):
        directory = QFileDialog.getExistingDirectory(self, '选择排除目录', os.getcwd())
        if directory:
            self._append_line(self.excludes_edit, directory)

    def add_exclude_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择排除文件', os.getcwd())
        if file_path:
            self._append_line(self.excludes_edit, file_path)

    def load_gitignore_excludes(self):
        indirs = normalize_items(self.indirs_edit.toPlainText())
        if not indirs:
            self._notify('warning', '未选择源码目录', '请先添加源码目录')
            return
        excludes = read_gitignore_excludes(indirs)
        if not excludes:
            self._notify('warning', '未读取到排除项', '请确认 .gitignore 是否存在或包含可解析路径')
            return
        self._append_lines(self.excludes_edit, excludes)
        self._notify('success', '读取完成', '已添加 {} 条排除路径'.format(len(excludes)))

    def choose_outfile(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, '选择输出文件', self.outfile_edit.text() or os.getcwd(), 'Word 文档 (*.docx)'
        )
        if file_path:
            file_path = os.path.normpath(file_path)
            if not file_path.lower().endswith('.docx'):
                file_path += '.docx'
            self.outfile_edit.setText(file_path)

    def choose_template(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择模板文件', os.getcwd(), 'Word 文档 (*.docx)'
        )
        if file_path:
            file_path = os.path.normpath(file_path)
            self.template_edit.setText(file_path)

    def open_output_dir(self):
        path = self.outfile_edit.text().strip()
        if not path:
            self._notify('warning', '输出路径为空', '请先选择或填写输出文件路径')
            return
        directory = os.path.dirname(path) or os.getcwd()
        if not os.path.isdir(directory):
            self._notify('warning', '输出目录不存在', directory)
            return
        self.logger.info(f'打开输出目录: {directory}')
        QDesktopServices.openUrl(QUrl.fromLocalFile(directory))
    
    def view_logs(self):
        """打开日志查看对话框"""
        self.logger.info('打开日志查看对话框')
        dialog = LogViewDialog(self)
        dialog.exec_()

    def _append_line(self, widget, text):
        """追加单行文本到widget（格式化为本地路径格式并去重）"""
        norm_text = os.path.normpath(text.strip())
        current = widget.toPlainText().strip()
        if current:
            lines = [os.path.normpath(line.strip()) for line in current.splitlines() if line.strip()]
            if norm_text not in lines:
                lines.append(norm_text)
            widget.setText('\n'.join(lines))
        else:
            widget.setText(norm_text)

    def _append_lines(self, widget, lines):
        """追加多行文本到widget（格式化为本地路径格式并去重）"""
        current = [os.path.normpath(item.strip()) for item in normalize_items(widget.toPlainText()) if item.strip()]
        norm_lines = [os.path.normpath(item.strip()) for item in lines if item.strip()]
        
        merged = []
        for item in current + norm_lines:
            if item not in merged:
                merged.append(item)
        widget.setText('\n'.join(merged))

    def build_config(self):
        indirs = normalize_items(self.indirs_edit.toPlainText())
        exts = normalize_exts(normalize_items(self.exts_edit.text()))
        comment_chars = normalize_items(self.comment_chars_edit.text())
        excludes = normalize_items(self.excludes_edit.toPlainText())
        title = self.title_edit.text().strip() or '软件著作权程序鉴别材料生成器V1.0'
        outfile = self.outfile_edit.text().strip() or os.path.abspath('code.docx')
        if not outfile.lower().endswith('.docx'):
            outfile += '.docx'
        template_path = self.template_edit.text().strip() or None
        encoding = self.encoding_combo.currentText().strip() or 'utf-8'
        if encoding == '自动':
            encoding = 'auto'
        
        # 获取页码格式
        page_number_format = self.page_number_combo.currentText()
        if page_number_format == '自定义...':
            page_number_format = self.custom_page_number_edit.text().strip() or '{page}'
        
        return {
            'title': title,
            'indirs': indirs,
            'exts': exts,
            'comment_chars': comment_chars,
            'font_name': self.font_name_edit.currentText().strip() or 'Consolas',
            'font_size': self.font_size_spin.value(),
            'space_before': self.space_before_spin.value(),
            'space_after': self.space_after_spin.value(),
            'line_spacing': self.line_spacing_spin.value(),
            'excludes': excludes,
            'outfile': outfile,
            'template_path': template_path,
            'skip_blank_lines': self.skip_blank_check.isChecked(),
            'skip_comment_lines': self.skip_comment_check.isChecked(),
            'encoding': encoding,
            'header_font': self.header_font_combo.currentText().strip() or '宋体',
            'header_font_size': self.header_font_size_spin.value(),
            'top_margin': self.top_margin_spin.value(),
            'bottom_margin': self.bottom_margin_spin.value(),
            'left_margin': self.left_margin_spin.value(),
            'right_margin': self.right_margin_spin.value(),
            'header_distance': self.header_distance_spin.value(),
            'footer_distance': self.footer_distance_spin.value(),
            'page_number_format': page_number_format,
            'selected_files': self.selected_files if self.selected_files else None,
            'pages_60_mode': self.pages_60_check.isChecked(),
            'export_pdf': self.export_pdf_check.isChecked(),
            'header_title_align': self.title_align_combo.currentText().strip() or '左对齐',
            'page_number_position': self.page_loc_combo.currentText().strip() or '页眉',
            'page_number_align': self.page_align_combo.currentText().strip() or '右对齐'
        }

    def schedule_extension_scan(self):
        if self.ext_scan_timer.isActive():
            self.ext_scan_timer.stop()
        self.ext_scan_timer.start(300)

    def start_extension_scan(self):
        if self.ext_worker and self.ext_worker.isRunning():
            self.pending_ext_scan = True
            return
        config = self.build_config()
        if not config['indirs']:
            self.available_exts = []
            return
        excludes = normalize_paths(config['excludes'])
        self.ext_worker = ExtensionScanWorker(config['indirs'], excludes)
        self.ext_worker.finished.connect(self.handle_extension_scan_finished)
        self.ext_worker.failed.connect(self.handle_extension_scan_failed)
        self.ext_worker.start()

    def handle_extension_scan_finished(self, exts):
        self.available_exts = exts
        if self.pending_ext_scan:
            self.pending_ext_scan = False
            self.start_extension_scan()

    def handle_extension_scan_failed(self, message):
        self._notify('warning', '后缀扫描失败', message)
        if self.pending_ext_scan:
            self.pending_ext_scan = False
            self.start_extension_scan()

    def open_extension_dialog(self):
        if not self.available_exts:
            self._notify('warning', '暂无可选后缀', '请先添加源码目录以自动收集后缀')
            return
        current = set(normalize_exts(normalize_items(self.exts_edit.text())))
        dialog = ExtensionSelectDialog(self.available_exts, current, self)
        if dialog.exec_() == QDialog.Accepted:
            selected = dialog.get_selected()
            self.exts_edit.setText(', '.join(selected))

    def _get_available_languages(self):
        langs = []
        for ext in self.available_exts:
            lang = LANGUAGE_BY_EXT.get(ext)
            if lang and lang in COMMENT_PREFIX_BY_LANG and lang not in langs:
                langs.append(lang)
        if not langs:
            langs = sorted(COMMENT_PREFIX_BY_LANG.keys())
        return langs

    def _get_selected_languages_from_prefixes(self):
        current_prefixes = normalize_items(self.comment_chars_edit.text())
        selected_langs = []
        for lang, prefixes in COMMENT_PREFIX_BY_LANG.items():
            if any(prefix in current_prefixes for prefix in prefixes):
                selected_langs.append(lang)
        return selected_langs

    def open_comment_prefix_dialog(self):
        langs = self._get_available_languages()
        dialog = CommentPrefixDialog(langs, self._get_selected_languages_from_prefixes(), self)
        if dialog.exec_() == QDialog.Accepted:
            prefixes = dialog.get_selected_prefixes()
            self.comment_chars_edit.setText(', '.join(prefixes))
    
    def open_file_select_dialog(self):
        """打开文件选择对话框"""
        if not self.available_files:
            self._notify('warning', '无可选文件', '请先扫描文件以获取可选文件列表')
            return
        
        # 获取当前配置
        encoding = self.encoding_combo.currentText().strip() or 'utf-8'
        if encoding == '自动':
            encoding = 'auto'
        
        comment_chars = normalize_items(self.comment_chars_edit.text())
        skip_blank_lines = self.skip_blank_check.isChecked()
        skip_comment_lines = self.skip_comment_check.isChecked()
        
        # 创建对话框并传递过滤配置
        dialog = FileSelectDialog(
            self.available_files, 
            self.selected_files,
            skip_blank_lines=skip_blank_lines,
            skip_comment_lines=skip_comment_lines,
            comment_chars=comment_chars,
            encoding=encoding,
            parent=self
        )
        
        if dialog.exec_() == QDialog.Accepted:
            self.selected_files = dialog.get_selected()
            # 保存文件顺序
            self.file_order = dialog.get_all_ordered()
            self._notify('success', '文件选择完成', f'已选择 {len(self.selected_files)} 个文件')
            self._update_summary()

    def _update_summary(self):
        config = self.build_config()
        indir_count = len(config['indirs'])
        ext_count = len(config['exts'])
        exclude_count = len(config['excludes'])
        filters = []
        if config['skip_blank_lines']:
            filters.append('空行')
        if config['skip_comment_lines']:
            filters.append('注释')
        filter_text = '、'.join(filters) if filters else '无'
        output_path = config['outfile']
        if output_path:
            output_path = os.path.abspath(output_path)
        else:
            output_path = '未设置'
        
        summary = [
            '源码目录：{} 个'.format(indir_count),
            '文件后缀：{} 个'.format(ext_count) if ext_count else '文件后缀：未填写',
            '排除路径：{} 条'.format(exclude_count),
            '过滤规则：{}'.format(filter_text),
            '输出路径：{}'.format(output_path),
            '上次扫描：{} 个文件'.format(self.last_scan_count)
        ]
        
        if self.selected_files:
            summary.append('已选文件：{} 个'.format(len(self.selected_files)))
        
        if self.last_stats:
            summary.append('原始总行数：{} 行'.format(self.last_stats.get('raw_lines', 0)))
            summary.append('代码行数：{} 行'.format(self.last_stats.get('code_lines', 0)))
            summary.append('注释行数：{} 行'.format(self.last_stats.get('comment_lines', 0)))
            summary.append('空行数：{} 行'.format(self.last_stats.get('blank_lines', 0)))
            summary.append('过滤后行数：{} 行'.format(self.last_stats.get('total_lines', 0)))
        
        if config['pages_60_mode']:
            summary.append('60页模式：已启用')
        
        self.summary_body.setText('\n'.join(summary))

    def eventFilter(self, obj, event):
        if obj in (self.indirs_edit, self.excludes_edit):
            if event.type() == QEvent.DragEnter:
                if event.mimeData().hasUrls():
                    event.acceptProposedAction()
                    return True
            elif event.type() == QEvent.DragMove:
                event.acceptProposedAction()
                return True
            elif event.type() == QEvent.Drop:
                urls = event.mimeData().urls()
                paths = []
                for url in urls:
                    path = os.path.normpath(url.toLocalFile())
                    if path:
                        paths.append(path)
                if paths:
                    self._append_lines(obj, paths)
                    self._update_summary()
                    event.accept()
                    return True

        if event.type() == QEvent.Wheel and hasattr(self, 'scroll_area'):
            bar = self.scroll_area.verticalScrollBar()
            if bar:
                delta = event.angleDelta().y()
                if delta:
                    bar.setValue(bar.value() - delta)
                    return True
        return super().eventFilter(obj, event)

    def closeEvent(self, event):
        """窗口关闭事件：确保后台工作线程被正确终止，防止残留后台进程"""
        self.logger.info('正在关闭窗口，正在清理后台线程...')
        # 强制停止扫描和生成线程
        if self.worker and self.worker.isRunning():
            self.logger.info('终止文档生成后台线程...')
            self.worker.terminate()
            self.worker.wait(1000)  # 最多等待 1 秒
        if self.ext_worker and self.ext_worker.isRunning():
            self.logger.info('终止后缀扫描后台线程...')
            self.ext_worker.terminate()
            self.ext_worker.wait(1000)
        event.accept()

    def _check_word_conflict(self, config):
        """（私有辅助方法）检查并处理 Word 进程冲突"""
        # 仅在启用 60 页模式或导出 PDF 时需要 Word COM
        if not (config.get('pages_60_mode') or config.get('export_pdf')):
            return True

        import subprocess
        
        # 仅检查 winword.exe 是否运行
        word_running = False
        try:
            output_word = subprocess.check_output(
                'tasklist /FI "IMAGENAME eq winword.exe" /FO CSV /NH',
                shell=True,
                stderr=subprocess.DEVNULL,
                creationflags=0x08000000
            )
            word_running = b"winword.exe" in output_word.lower()
        except:
            pass

        if word_running:
            w = QMessageBox(self)
            w.setWindowTitle('检测到 Word 正在运行')
            w.setText('系统检测到后台/前台有正在运行的 Word 进程。\n'
                      '由于 60 页模式及 PDF 导出依赖 Word COM 自动化，此时运行可能会产生独占访问冲突并导致失败。\n\n'
                      '建议您先保存当前所有 Word 文档。\n'
                      ' - 点击【强杀进程】将自动终止后台 Word 进程并继续生成。\n'
                      ' - 点击【取消生成】以退出，建议您保存文件并手动关闭后重试。')
            w.setIcon(QMessageBox.Warning)
            
            # 适配暗黑/浅色模式的样式，保持与 Fluent 风格大体一致
            if isDarkTheme():
                w.setStyleSheet("""
                    QMessageBox { background-color: rgb(32, 32, 32); color: white; }
                    QLabel { color: white; font-family: "Microsoft YaHei"; font-size: 13px; }
                    QPushButton { background-color: rgb(45, 45, 45); color: white; border: 1px solid rgb(60, 60, 60); border-radius: 4px; min-width: 80px; min-height: 28px; font-family: "Microsoft YaHei"; }
                    QPushButton:hover { background-color: rgb(60, 60, 60); }
                    QPushButton:pressed { background-color: rgb(35, 35, 35); }
                """)
            else:
                w.setStyleSheet("""
                    QMessageBox { background-color: rgb(243, 243, 243); color: black; }
                    QLabel { color: black; font-family: "Microsoft YaHei"; font-size: 13px; }
                    QPushButton { background-color: white; color: black; border: 1px solid rgb(200, 200, 200); border-radius: 4px; min-width: 80px; min-height: 28px; font-family: "Microsoft YaHei"; }
                    QPushButton:hover { background-color: rgb(240, 240, 240); }
                    QPushButton:pressed { background-color: rgb(220, 220, 220); }
                """)
                
            yes_btn = w.addButton('强杀进程', QMessageBox.YesRole)
            cancel_btn = w.addButton('取消生成', QMessageBox.NoRole)
            
            w.exec_()
            
            if w.clickedButton() == yes_btn:
                killed_ok = True
                try:
                    subprocess.check_call(
                        'taskkill /f /im winword.exe', 
                        shell=True, 
                        stdout=subprocess.DEVNULL, 
                        stderr=subprocess.DEVNULL,
                        creationflags=0x08000000
                    )
                except Exception as e:
                    killed_ok = False
                
                if killed_ok:
                    self._notify('success', '清理完成', '已终止所有后台 Word 进程')
                    return True
                else:
                    self._notify('warning', '清理失败', 'Word 进程无法自动终止。为避免软件卡死，请手动关闭后重试。')
                    return False  # 强杀失败则不允许继续，防止卡死
            else:
                return False  # 取消生成
        return True

    def start_worker(self, mode):
        if self.worker and self.worker.isRunning():
            self._notify('warning', '任务进行中', '请等待当前任务完成后再操作')
            return
        config = self.build_config()
        valid, message = self._validate_inputs(config, mode)
        if not valid:
            self._notify('warning', '请检查输入', message)
            return
        if not config['exts']:
            self._notify('warning', '未选择后缀', '请点击“选择”添加文件后缀')
            return
        
        # Word 进程冲突前置检查
        if mode == 'generate':
            if not self._check_word_conflict(config):
                self._notify('info', '已取消', '已取消文档生成')
                return

        self.set_buttons_enabled(False)
        if mode == 'scan':
            self.status_label.setText('正在扫描文件，请稍候...')
            self.summary_title.setText('正在扫描文件')
        elif mode == 'count':
            self.status_label.setText('正在统计代码行数，请稍候...')
            self.summary_title.setText('正在统计代码')
        else:
            self.status_label.setText('正在生成文档，请稍候...')
            self.summary_title.setText('正在生成文档')
        self.worker = GenerateWorker(config, mode)
        self.worker.finished.connect(self.handle_finished)
        self.worker.failed.connect(self.handle_failed)
        self.worker.start()

    def handle_finished(self, result):
        self.set_buttons_enabled(True)
        mode = result.get('mode')
        
        if mode == 'scan':
            self.last_scan_count = result.get('file_count', 0)
            self.available_files = result.get('files', [])
            if not self.selected_files:
                self.selected_files = self.available_files[:]
            if not self.file_order:
                self.file_order = self.available_files[:]
            self._update_summary()
            self.status_label.setText('扫描完成，共找到 {} 个文件'.format(self.last_scan_count))
            self.summary_title.setText('扫描完成')
            self._notify('success', '扫描完成', '共找到 {} 个文件'.format(self.last_scan_count))
            return
        
        if mode == 'count':
            stats = result.get('stats', {})
            self.last_stats = stats
            self._update_summary()
            content_text = '代码：{} 行，注释：{} 行，空行：{} 行'.format(
                stats.get('code_lines', 0), stats.get('comment_lines', 0), stats.get('blank_lines', 0)
            )
            self.status_label.setText('统计完成，' + content_text)
            self.summary_title.setText('统计完成')
            self._notify('success', '统计完成', content_text)
            return
        
        # generate 模式
        total_lines_info = result.get('total_lines', {})
        if isinstance(total_lines_info, dict):
            self.last_stats = total_lines_info
            total_lines = total_lines_info.get('total_lines', 0)
        else:
            total_lines = total_lines_info
        
        self.status_label.setText('生成完成，共写入 {} 个文件，{} 行代码'.format(
            result.get('file_count', 0), total_lines))
        self.summary_title.setText('生成完成')
        
        # 检查 PDF 导出结果
        pdf_path = result.get('pdf_path')
        pdf_warning = result.get('pdf_warning')
        if pdf_path:
            self._notify('success', '生成完成', '输出文件：{}\nPDF：{}'.format(
                result.get('outfile', ''), pdf_path))
        elif pdf_warning:
            self._notify('warning', '文档已生成，PDF导出失败', pdf_warning)
        else:
            self._notify('success', '生成完成', '输出文件：{}'.format(result.get('outfile', '')))
        self._update_open_output_enabled()

    def handle_failed(self, message):
        self.set_buttons_enabled(True)
        self.status_label.setText('失败：{}'.format(message))
        self.summary_title.setText('任务失败')
        self._notify('error', '生成失败', message)
        # 显示错误弹窗
        w = MessageBox('操作失败', f'任务执行失败:\n\n{message}', self)
        w.exec_()
        w.hide()
        w.close()

    def set_buttons_enabled(self, enabled):
        self.scan_btn.setEnabled(enabled)
        self.generate_btn.setEnabled(enabled)
        self.select_files_btn.setEnabled(enabled)
        self.count_lines_btn.setEnabled(enabled)
        self.open_output_btn.setEnabled(enabled)

    def reset_style_defaults(self):
        self.font_name_edit.setCurrentText('Consolas')
        self.font_size_spin.setValue(10.5)
        self.header_font_combo.setCurrentText('宋体')
        self.header_font_size_spin.setValue(10.5)
        self.space_before_spin.setValue(0.0)
        self.space_after_spin.setValue(2.3)
        self.line_spacing_spin.setValue(10.5)
        self.top_margin_spin.setValue(2.5)
        self.bottom_margin_spin.setValue(2.5)
        self.left_margin_spin.setValue(2.5)
        self.right_margin_spin.setValue(2.5)
        self.header_distance_spin.setValue(1.5)
        self.footer_distance_spin.setValue(1.75)
        self.page_number_combo.setCurrentText('{page}')
        self.title_align_combo.setCurrentText('左对齐')
        self.page_loc_combo.setCurrentText('页眉')
        self.page_align_combo.setCurrentText('右对齐')

    def _notify(self, level, title, content):
        """统一的通知方法"""
        duration = 5000 if level == 'error' else 3500
        notify_func = {
            'success': InfoBar.success,
            'error': InfoBar.error,
            'warning': InfoBar.warning
        }.get(level, InfoBar.info)
        notify_func(
            title=title, 
            content=content, 
            duration=duration,
            isClosable=True,
            position=InfoBarPosition.TOP, 
            parent=self
        )

    def _validate_inputs(self, config, mode):
        if not config['indirs']:
            return False, '请至少选择一个源码目录'
        for indir in config['indirs']:
            if not os.path.isdir(indir):
                return False, '无效源码目录：{}'.format(indir)
        for exclude in config['excludes']:
            if not os.path.exists(exclude):
                return False, '排除路径不存在：{}'.format(exclude)
        if mode == 'generate':
            outfile = config['outfile']
            output_dir = os.path.dirname(outfile) if outfile else ''
            if output_dir and not os.path.isdir(output_dir):
                return False, '输出目录不存在：{}'.format(output_dir)
            if config['template_path'] and not os.path.isfile(config['template_path']):
                return False, '模板文件不存在：{}'.format(config['template_path'])
            if config.get('page_number_position') == '页眉' and config.get('header_title_align') == config.get('page_number_align'):
                return False, '页眉标题和页码不能在同一位置（目前同为【{}】）'.format(config.get('header_title_align'))
        return True, ''

    def _update_open_output_enabled(self):
        path = self.outfile_edit.text().strip()
        if not path:
            self.open_output_btn.setEnabled(False)
            return
        directory = os.path.dirname(path)
        if not directory:
            directory = os.getcwd()
        self.open_output_btn.setEnabled(os.path.isdir(directory))

    def apply_theme_stylesheet(self):
        """根据主题动态更新主窗口背景样式"""
        is_dark = isDarkTheme()
        bg_color = "rgb(32, 32, 32)" if is_dark else "rgb(243, 243, 243)"
        # 同时作用于主窗口和中央包裹容器，确保背景色深度一致
        self.setStyleSheet(f"GeneratorWindow, #centralWrapper {{ background-color: {bg_color}; }}")
        
        # 动态设置操作系统窗口标题栏样式（暗黑模式 / 浅色模式）
        set_window_dark_title_bar(self, is_dark)
        
        # 确保滚动区域及其内部容器背景透明且能够重绘，防止 QSS 重置导致左侧卡片区保留白色底色
        if hasattr(self, 'scroll_area') and self.scroll_area:
            self.scroll_area.enableTransparentBackground()
            if self.scroll_area.widget():
                self.scroll_area.widget().setStyleSheet("#scrollContainer { background-color: transparent; }")

    def _on_theme_changed(self, text):
        """当手动更改主题模式时触发"""
        if text == '浅色模式':
            setTheme(Theme.LIGHT)
        elif text == '深色模式':
            setTheme(Theme.DARK)
        else:
            setTheme(Theme.AUTO)
        self.apply_theme_stylesheet()


def launch_gui():
    """启动GUI应用"""
    app = QApplication(sys.argv)
    
    # 设置全局字体以保持默认与浅/深色模式下的字体表现一致
    app.setFont(QFont("Microsoft YaHei", 9))
    
    window = GeneratorWindow()
    window.show()
    sys.exit(app.exec_())
