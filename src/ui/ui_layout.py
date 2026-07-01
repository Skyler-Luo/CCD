# -*- coding: utf-8 -*-
"""
UI 布局装配模块
负责主窗口静态组件的创建和排版布局，与逻辑解耦
"""
import os
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout
)
from qfluentwidgets import (
    PrimaryPushButton, PushButton, LineEdit, TextEdit,
    DoubleSpinBox, ComboBox, CheckBox, HorizontalSeparator,
    TitleLabel, BodyLabel, CaptionLabel, StrongBodyLabel, CardWidget,
    ScrollArea
)

class GeneratorWindowUI:
    """主窗口 UI 布局装配类"""
    
    def setup_ui(self, window):
        """装配静态界面组件"""
        # 1. 声明中央包裹容器
        wrapper = QWidget(window)
        wrapper.setObjectName("centralWrapper")
        wrapper_layout = QHBoxLayout(wrapper)
        wrapper_layout.setContentsMargins(16, 16, 16, 16)
        wrapper_layout.setSpacing(18)
        window.setCentralWidget(wrapper)

        # 2. 声明滚动卡片区域 (左侧)
        scroll = ScrollArea(window)
        scroll.setWidgetResizable(True)
        scroll.enableTransparentBackground()
        window.scroll_area = scroll
        
        container = QWidget()
        container.setObjectName("scrollContainer")
        scroll.setWidget(container)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(16)

        # 3. 头部标题行
        title = TitleLabel('软件著作权源代码文档生成器')
        subtitle = CaptionLabel('支持多目录扫描、注释过滤、模板与排版参数配置')
        title.setFont(QFont('Microsoft YaHei', 18, QFont.Bold))
        subtitle.setStyleSheet('font-size: 12px;')
        subtitle.setWordWrap(True)
        
        window.theme_label = BodyLabel('主题模式')
        window.theme_combo = ComboBox()
        window.theme_combo.addItems(['系统自适应', '浅色模式', '深色模式'])
        window.theme_combo.blockSignals(True)
        window.theme_combo.setCurrentText('系统自适应')
        window.theme_combo.blockSignals(False)

        header_row = QWidget()
        header_layout = QHBoxLayout(header_row)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.addWidget(title)
        header_layout.addStretch(1)
        header_layout.addWidget(window.theme_label)
        header_layout.addWidget(window.theme_combo)
        layout.addWidget(header_row)
        layout.addWidget(subtitle)

        # 4. 声明右侧控制栏
        right_column = QWidget()
        right_column.setMinimumWidth(340)
        right_layout = QVBoxLayout(right_column)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(16)

        # 5. 源码设置卡片
        source_group, source_layout = self._create_group('源码设置')
        source_grid = QGridLayout()
        source_grid.setHorizontalSpacing(12)
        source_grid.setVerticalSpacing(12)
        source_grid.setColumnStretch(1, 1)
        source_grid.setColumnMinimumWidth(0, 110)
        source_layout.addLayout(source_grid)

        # 源码目录输入区
        window.indirs_edit = TextEdit()
        window.indirs_edit.setPlaceholderText('每行一个目录')
        window.indirs_edit.setText('')
        window.indirs_edit.setMinimumHeight(90)
        
        indir_buttons = QHBoxLayout()
        indir_buttons.setSpacing(8)
        window.add_indir_btn = PushButton('添加目录')
        window.clear_indir_btn = PushButton('清空')
        window.add_indir_btn.setMinimumHeight(32)
        window.clear_indir_btn.setMinimumHeight(32)
        indir_buttons.addWidget(window.add_indir_btn)
        indir_buttons.addWidget(window.clear_indir_btn)
        indir_buttons.addStretch(1)
        
        indir_container = QWidget()
        indir_container_layout = QVBoxLayout(indir_container)
        indir_container_layout.setContentsMargins(0, 0, 0, 0)
        indir_container_layout.addWidget(window.indirs_edit)
        indir_container_layout.addLayout(indir_buttons)
        source_grid.addWidget(BodyLabel('源码目录'), 0, 0)
        source_grid.addWidget(indir_container, 0, 1)

        # 后缀选择
        window.exts_edit = LineEdit()
        window.exts_edit.setPlaceholderText('如 py, java, cpp')
        window.exts_edit.setText('')
        window.exts_edit.setMinimumHeight(32)
        window.exts_select_btn = PushButton('选择')
        window.exts_select_btn.setMinimumHeight(32)
        exts_row = QWidget()
        exts_layout = QHBoxLayout(exts_row)
        exts_layout.setContentsMargins(0, 0, 0, 0)
        exts_layout.setSpacing(8)
        exts_layout.addWidget(window.exts_edit, 1)
        exts_layout.addWidget(window.exts_select_btn)
        source_grid.addWidget(BodyLabel('文件后缀'), 1, 0)
        source_grid.addWidget(exts_row, 1, 1)

        # 注释前缀选择
        window.comment_chars_edit = LineEdit()
        window.comment_chars_edit.setPlaceholderText('如 #, //')
        window.comment_chars_edit.setText('')
        window.comment_chars_edit.setMinimumHeight(32)
        window.comment_select_btn = PushButton('选择')
        window.comment_select_btn.setMinimumHeight(32)
        comment_row = QWidget()
        comment_layout = QHBoxLayout(comment_row)
        comment_layout.setContentsMargins(0, 0, 0, 0)
        comment_layout.setSpacing(8)
        comment_layout.addWidget(window.comment_chars_edit, 1)
        comment_layout.addWidget(window.comment_select_btn)
        source_grid.addWidget(BodyLabel('注释前缀'), 2, 0)
        source_grid.addWidget(comment_row, 2, 1)

        # 排除路径
        window.excludes_edit = TextEdit()
        window.excludes_edit.setPlaceholderText('每行一个排除路径')
        window.excludes_edit.setMinimumHeight(90)
        
        exclude_buttons = QHBoxLayout()
        exclude_buttons.setSpacing(8)
        window.add_exclude_dir_btn = PushButton('添加排除目录')
        window.add_exclude_file_btn = PushButton('添加排除文件')
        window.read_gitignore_btn = PushButton('读取.gitignore')
        window.clear_exclude_btn = PushButton('清空')
        window.add_exclude_dir_btn.setMinimumHeight(32)
        window.add_exclude_file_btn.setMinimumHeight(32)
        window.read_gitignore_btn.setMinimumHeight(32)
        window.clear_exclude_btn.setMinimumHeight(32)
        exclude_buttons.addWidget(window.add_exclude_dir_btn)
        exclude_buttons.addWidget(window.add_exclude_file_btn)
        exclude_buttons.addWidget(window.read_gitignore_btn)
        exclude_buttons.addWidget(window.clear_exclude_btn)
        exclude_buttons.addStretch(1)
        
        exclude_container = QWidget()
        exclude_container_layout = QVBoxLayout(exclude_container)
        exclude_container_layout.setContentsMargins(0, 0, 0, 0)
        exclude_container_layout.addWidget(window.excludes_edit)
        exclude_container_layout.addLayout(exclude_buttons)
        source_grid.addWidget(BodyLabel('排除路径'), 3, 0)
        source_grid.addWidget(exclude_container, 3, 1)

        # 文件编码
        window.encoding_combo = ComboBox()
        window.encoding_combo.addItems(['自动', 'utf-8', 'utf-8-sig', 'gbk', 'gb18030'])
        window.encoding_combo.setMinimumHeight(32)
        source_grid.addWidget(BodyLabel('文件编码'), 4, 0)
        source_grid.addWidget(window.encoding_combo, 4, 1)

        # 过滤规则
        window.skip_blank_check = CheckBox('过滤空行')
        window.skip_blank_check.setChecked(True)
        window.skip_comment_check = CheckBox('过滤注释')
        window.skip_comment_check.setChecked(True)
        check_row = QWidget()
        check_layout = QHBoxLayout(check_row)
        check_layout.setContentsMargins(0, 0, 0, 0)
        check_layout.addWidget(window.skip_blank_check)
        check_layout.addWidget(window.skip_comment_check)
        check_layout.addStretch(1)
        source_grid.addWidget(BodyLabel('过滤规则'), 5, 0)
        source_grid.addWidget(check_row, 5, 1)

        layout.addWidget(source_group)

        # 6. 文档设置卡片
        output_group, output_layout = self._create_group('文档设置')
        output_grid = QGridLayout()
        output_grid.setHorizontalSpacing(12)
        output_grid.setVerticalSpacing(12)
        output_grid.setColumnStretch(1, 1)
        output_grid.setColumnMinimumWidth(0, 110)
        output_layout.addLayout(output_grid)

        # 页眉标题
        window.title_edit = LineEdit()
        window.title_edit.setText('软件著作权程序鉴别材料生成器V1.0')
        window.title_edit.setMinimumHeight(32)
        
        window.title_align_combo = ComboBox()
        window.title_align_combo.addItems(['左对齐', '居中', '右对齐'])
        window.title_align_combo.setCurrentText('左对齐')
        window.title_align_combo.setMinimumHeight(32)
        window.title_align_combo.setMinimumWidth(100)
        
        title_row = QWidget()
        title_layout = QHBoxLayout(title_row)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(8)
        title_layout.addWidget(window.title_edit, 1)
        title_layout.addWidget(window.title_align_combo)
        
        output_grid.addWidget(BodyLabel('页眉标题'), 0, 0)
        output_grid.addWidget(title_row, 0, 1)

        # 输出路径
        window.outfile_edit = LineEdit()
        window.outfile_edit.setText(os.path.abspath('code.docx'))
        window.outfile_edit.setPlaceholderText('如 D:\\输出\\code.docx')
        window.outfile_edit.setMinimumHeight(32)
        window.outfile_btn = PushButton('选择输出文件')
        window.outfile_btn.setMinimumHeight(32)
        outfile_row = QWidget()
        outfile_layout = QHBoxLayout(outfile_row)
        outfile_layout.setContentsMargins(0, 0, 0, 0)
        outfile_layout.setSpacing(8)
        outfile_layout.addWidget(window.outfile_edit)
        outfile_layout.addWidget(window.outfile_btn)
        output_grid.addWidget(BodyLabel('输出路径'), 1, 0)
        output_grid.addWidget(outfile_row, 1, 1)

        # 模板文件
        window.template_edit = LineEdit()
        window.template_edit.setPlaceholderText('可选，留空使用默认模板')
        window.template_edit.setMinimumHeight(32)
        window.template_btn = PushButton('选择模板')
        window.template_clear_btn = PushButton('清空')
        window.template_btn.setMinimumHeight(32)
        window.template_clear_btn.setMinimumHeight(32)
        template_row = QWidget()
        template_layout = QHBoxLayout(template_row)
        template_layout.setContentsMargins(0, 0, 0, 0)
        template_layout.setSpacing(8)
        template_layout.addWidget(window.template_edit)
        template_layout.addWidget(window.template_btn)
        template_layout.addWidget(window.template_clear_btn)
        output_grid.addWidget(BodyLabel('模板文件'), 2, 0)
        output_grid.addWidget(template_row, 2, 1)

        layout.addWidget(output_group)

        # 7. 排版设置卡片
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

        window.font_name_edit = ComboBox()
        window.font_name_edit.addItems(['Arial', 'Consolas', 'Courier New', 'Times New Roman', '仿宋', '黑体', '楷体', '宋体', '微软雅黑'])
        window.font_name_edit.setCurrentText('Consolas')
        window.font_name_edit.setMinimumHeight(32)
        font_grid.addWidget(BodyLabel('正文字体'), 0, 0)
        font_grid.addWidget(window.font_name_edit, 0, 1)

        window.font_size_spin = DoubleSpinBox()
        window.font_size_spin.setRange(1.0, 72.0)
        window.font_size_spin.setValue(10.5)
        window.font_size_spin.setMinimumHeight(32)
        font_grid.addWidget(BodyLabel('正文字号'), 1, 0)
        font_grid.addWidget(window.font_size_spin, 1, 1)

        window.header_font_combo = ComboBox()
        window.header_font_combo.addItems(['Arial', 'Times New Roman', '仿宋', '黑体', '楷体', '宋体', '微软雅黑'])
        window.header_font_combo.setCurrentText('宋体')
        window.header_font_combo.setMinimumHeight(32)
        font_grid.addWidget(BodyLabel('页眉页脚字体'), 2, 0)
        font_grid.addWidget(window.header_font_combo, 2, 1)

        window.header_font_size_spin = DoubleSpinBox()
        window.header_font_size_spin.setRange(1.0, 72.0)
        window.header_font_size_spin.setValue(10.5)
        window.header_font_size_spin.setMinimumHeight(32)
        font_grid.addWidget(BodyLabel('页眉页脚字号'), 3, 0)
        font_grid.addWidget(window.header_font_size_spin, 3, 1)

        # 分隔线
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

        window.space_before_spin = DoubleSpinBox()
        window.space_before_spin.setRange(0.0, 72.0)
        window.space_before_spin.setValue(0.0)
        window.space_before_spin.setMinimumHeight(32)
        para_grid.addWidget(BodyLabel('段前间距(pt)'), 0, 0)
        para_grid.addWidget(window.space_before_spin, 0, 1)

        window.space_after_spin = DoubleSpinBox()
        window.space_after_spin.setRange(0.0, 72.0)
        window.space_after_spin.setValue(2.3)
        window.space_after_spin.setMinimumHeight(32)
        para_grid.addWidget(BodyLabel('段后间距(pt)'), 1, 0)
        para_grid.addWidget(window.space_after_spin, 1, 1)

        window.line_spacing_spin = DoubleSpinBox()
        window.line_spacing_spin.setRange(0.0, 72.0)
        window.line_spacing_spin.setValue(10.5)
        window.line_spacing_spin.setMinimumHeight(32)
        para_grid.addWidget(BodyLabel('行距(pt)'), 2, 0)
        para_grid.addWidget(window.line_spacing_spin, 2, 1)

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

        window.top_margin_spin = DoubleSpinBox()
        window.top_margin_spin.setRange(0.0, 10.0)
        window.top_margin_spin.setValue(2.5)
        window.top_margin_spin.setSingleStep(0.1)
        window.top_margin_spin.setDecimals(2)
        window.top_margin_spin.setSuffix(' cm')
        window.top_margin_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('上边距'), 0, 0)
        margin_grid.addWidget(window.top_margin_spin, 0, 1)

        window.bottom_margin_spin = DoubleSpinBox()
        window.bottom_margin_spin.setRange(0.0, 10.0)
        window.bottom_margin_spin.setValue(2.5)
        window.bottom_margin_spin.setSingleStep(0.1)
        window.bottom_margin_spin.setDecimals(2)
        window.bottom_margin_spin.setSuffix(' cm')
        window.bottom_margin_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('下边距'), 0, 2)
        margin_grid.addWidget(window.bottom_margin_spin, 0, 3)

        window.left_margin_spin = DoubleSpinBox()
        window.left_margin_spin.setRange(0.0, 10.0)
        window.left_margin_spin.setValue(2.5)
        window.left_margin_spin.setSingleStep(0.1)
        window.left_margin_spin.setDecimals(2)
        window.left_margin_spin.setSuffix(' cm')
        window.left_margin_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('左边距'), 1, 0)
        margin_grid.addWidget(window.left_margin_spin, 1, 1)

        window.right_margin_spin = DoubleSpinBox()
        window.right_margin_spin.setRange(0.0, 10.0)
        window.right_margin_spin.setValue(2.5)
        window.right_margin_spin.setSingleStep(0.1)
        window.right_margin_spin.setDecimals(2)
        window.right_margin_spin.setSuffix(' cm')
        window.right_margin_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('右边距'), 1, 2)
        margin_grid.addWidget(window.right_margin_spin, 1, 3)

        window.header_distance_spin = DoubleSpinBox()
        window.header_distance_spin.setRange(0.0, 10.0)
        window.header_distance_spin.setValue(1.5)
        window.header_distance_spin.setSingleStep(0.1)
        window.header_distance_spin.setDecimals(2)
        window.header_distance_spin.setSuffix(' cm')
        window.header_distance_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('页眉距离'), 2, 0)
        margin_grid.addWidget(window.header_distance_spin, 2, 1)

        window.footer_distance_spin = DoubleSpinBox()
        window.footer_distance_spin.setRange(0.0, 10.0)
        window.footer_distance_spin.setValue(1.75)
        window.footer_distance_spin.setSingleStep(0.1)
        window.footer_distance_spin.setDecimals(2)
        window.footer_distance_spin.setSuffix(' cm')
        window.footer_distance_spin.setMinimumHeight(32)
        margin_grid.addWidget(BodyLabel('页脚距离'), 2, 2)
        margin_grid.addWidget(window.footer_distance_spin, 2, 3)

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

        window.page_number_combo = ComboBox()
        window.page_number_combo.addItems([
            '{page}',
            '第 {page} 页 共 {total} 页',
            '{page} / {total}',
            'Page {page} of {total}',
            '自定义...'
        ])
        window.page_number_combo.setMinimumHeight(32)

        window.custom_page_number_edit = LineEdit()
        window.custom_page_number_edit.setPlaceholderText('使用 {page} 和 {total} 作为占位符')
        window.custom_page_number_edit.setMinimumHeight(32)
        window.custom_page_number_edit.setVisible(False)
        page_grid.addWidget(BodyLabel('格式'), 0, 0)
        page_grid.addWidget(window.page_number_combo, 0, 1)
        page_grid.addWidget(window.custom_page_number_edit, 1, 1)

        window.page_loc_combo = ComboBox()
        window.page_loc_combo.addItems(['页眉', '页脚'])
        window.page_loc_combo.setCurrentText('页眉')
        window.page_loc_combo.setMinimumHeight(32)
        page_grid.addWidget(BodyLabel('页码位置'), 2, 0)
        page_grid.addWidget(window.page_loc_combo, 2, 1)

        window.page_align_combo = ComboBox()
        window.page_align_combo.addItems(['左对齐', '居中', '右对齐'])
        window.page_align_combo.setCurrentText('右对齐')
        window.page_align_combo.setMinimumHeight(32)
        page_grid.addWidget(BodyLabel('页码对齐'), 3, 0)
        page_grid.addWidget(window.page_align_combo, 3, 1)

        # 恢复默认按钮
        style_layout.addSpacing(8)
        style_btn_row = QWidget()
        style_btn_layout = QHBoxLayout(style_btn_row)
        style_btn_layout.setContentsMargins(0, 0, 0, 0)
        style_btn_layout.setSpacing(8)
        window.reset_style_btn = PushButton('恢复默认')
        window.reset_style_btn.setMinimumHeight(32)
        style_btn_layout.addWidget(window.reset_style_btn)
        style_btn_layout.addStretch(1)
        style_layout.addWidget(style_btn_row)

        layout.addWidget(style_group)

        # 8. 运行概览卡片
        summary_group, summary_layout = self._create_group('运行概览')
        window.summary_title = StrongBodyLabel('尚未执行任务')
        window.summary_title.setStyleSheet('font-size: 12px;')
        window.summary_body = CaptionLabel('')
        window.summary_body.setWordWrap(True)
        window.summary_body.setStyleSheet('font-size: 12px;')
        summary_layout.addWidget(window.summary_title)
        summary_layout.addWidget(window.summary_body)
        right_layout.addWidget(summary_group)

        # 9. 操作区卡片
        action_group, action_layout = self._create_group('操作区')
        action_row = QWidget()
        action_row_layout = QVBoxLayout(action_row)
        action_row_layout.setContentsMargins(0, 0, 0, 0)
        action_row_layout.setSpacing(8)
        
        window.pages_60_check = CheckBox('启用60页模式（前30+后30）')
        window.pages_60_check.setChecked(False)
        action_row_layout.addWidget(window.pages_60_check)
        
        window.export_pdf_check = CheckBox('同时导出 PDF')
        window.export_pdf_check.setChecked(False)
        action_row_layout.addWidget(window.export_pdf_check)
        
        window.generate_btn = PrimaryPushButton('生成文档')
        window.scan_btn = PushButton('扫描文件数')
        window.select_files_btn = PushButton('选择文件')
        window.count_lines_btn = PushButton('统计代码行数')
        window.open_output_btn = PushButton('打开输出目录')
        window.view_log_btn = PushButton('查看日志')
        
        window.generate_btn.setMinimumHeight(36)
        window.scan_btn.setMinimumHeight(34)
        window.select_files_btn.setMinimumHeight(34)
        window.count_lines_btn.setMinimumHeight(34)
        window.open_output_btn.setMinimumHeight(34)
        window.view_log_btn.setMinimumHeight(34)
        
        action_row_layout.addWidget(window.generate_btn)
        action_row_layout.addWidget(window.scan_btn)
        action_row_layout.addWidget(window.select_files_btn)
        action_row_layout.addWidget(window.count_lines_btn)
        action_row_layout.addWidget(window.open_output_btn)
        action_row_layout.addWidget(window.view_log_btn)
        action_layout.addWidget(action_row)
        
        window.status_label = CaptionLabel('')
        window.status_label.setWordWrap(True)
        window.status_label.setStyleSheet('font-size: 12px;')
        action_layout.addWidget(window.status_label)
        right_layout.addWidget(action_group)
        right_layout.addStretch(1)

        layout.addStretch(1)
        wrapper_layout.addWidget(scroll, 1)
        wrapper_layout.addWidget(right_column)

        # 10. 信号与事件槽绑定 (绑定回 window 的方法)
        window.add_indir_btn.clicked.connect(window.add_indir)
        window.clear_indir_btn.clicked.connect(lambda: window.indirs_edit.setText(''))
        window.add_exclude_dir_btn.clicked.connect(window.add_exclude_dir)
        window.add_exclude_file_btn.clicked.connect(window.add_exclude_file)
        window.read_gitignore_btn.clicked.connect(window.load_gitignore_excludes)
        window.clear_exclude_btn.clicked.connect(lambda: window.excludes_edit.setText(''))
        window.outfile_btn.clicked.connect(window.choose_outfile)
        window.template_btn.clicked.connect(window.choose_template)
        window.template_clear_btn.clicked.connect(lambda: window.template_edit.setText(''))
        window.reset_style_btn.clicked.connect(window.reset_style_defaults)
        
        window.generate_btn.clicked.connect(lambda: window.start_worker('generate'))
        window.scan_btn.clicked.connect(lambda: window.start_worker('scan'))
        window.select_files_btn.clicked.connect(window.open_file_select_dialog)
        window.count_lines_btn.clicked.connect(lambda: window.start_worker('count'))
        window.open_output_btn.clicked.connect(window.open_output_dir)
        window.view_log_btn.clicked.connect(window.view_logs)
        
        window.outfile_edit.textChanged.connect(window._update_open_output_enabled)
        window.outfile_edit.textChanged.connect(window._update_summary)
        window.indirs_edit.textChanged.connect(window._update_summary)
        window.indirs_edit.textChanged.connect(window.schedule_extension_scan)
        window.exts_edit.textChanged.connect(window._update_summary)
        window.comment_chars_edit.textChanged.connect(window._update_summary)
        window.excludes_edit.textChanged.connect(window._update_summary)
        window.encoding_combo.currentTextChanged.connect(window._update_summary)
        window.skip_blank_check.toggled.connect(window._update_summary)
        window.skip_comment_check.toggled.connect(window._update_summary)
        window.pages_60_check.toggled.connect(window._update_summary)
        
        window.title_align_combo.currentTextChanged.connect(lambda: window._validate_header_footer_positions('title_align'))
        window.title_align_combo.currentTextChanged.connect(window._update_summary)
        window.page_loc_combo.currentTextChanged.connect(lambda: window._validate_header_footer_positions('page_loc'))
        window.page_loc_combo.currentTextChanged.connect(window._update_summary)
        window.page_align_combo.currentTextChanged.connect(lambda: window._validate_header_footer_positions('page_align'))
        window.page_align_combo.currentTextChanged.connect(window._update_summary)
        
        window.page_number_combo.currentTextChanged.connect(window._on_page_number_format_changed)
        window.exts_select_btn.clicked.connect(window.open_extension_dialog)
        window.comment_select_btn.clicked.connect(window.open_comment_prefix_dialog)
        window.theme_combo.currentTextChanged.connect(window._on_theme_changed)

        # 11. 安装事件过滤器
        for widget in (right_column, summary_group, action_group, window.status_label, window.indirs_edit, window.excludes_edit):
            widget.installEventFilter(window)

    def _create_group(self, title):
        group = CardWidget()
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        header = StrongBodyLabel(title)
        header.setStyleSheet('font-size: 13px;')
        layout.addWidget(header)
        return group, layout
