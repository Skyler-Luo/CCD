# -*- coding: utf-8 -*-
import os
import sys

from PyQt5.QtCore import QThread, QUrl, pyqtSignal
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QFileDialog, QScrollArea, QMainWindow
)
from qfluentwidgets import (
    PrimaryPushButton, PushButton, LineEdit, TextEdit,
    DoubleSpinBox, ComboBox, CheckBox, InfoBar, InfoBarPosition,
    TitleLabel, BodyLabel, CardWidget, setTheme, Theme
)

from core import (
    collect_code_files, generate_code_doc, normalize_items, normalize_exts
)


class GenerateWorker(QThread):
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, config, mode):
        super().__init__()
        self.config = config
        self.mode = mode

    def run(self):
        try:
            if self.mode == 'scan':
                files = collect_code_files(
                    self.config['indirs'],
                    self.config['exts'],
                    self.config['excludes']
                )
                self.finished.emit({
                    'mode': 'scan',
                    'file_count': len(files)
                })
            else:
                result = generate_code_doc(**self.config)
                result['mode'] = 'generate'
                self.finished.emit(result)
        except Exception as exc:
            self.failed.emit(str(exc))


class GeneratorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        setTheme(Theme.AUTO)
        self.setWindowTitle('软著源代码文档生成器')
        self.resize(1000, 720)
        self.worker = None
        self._build_ui()

    def _build_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        scroll.setWidget(container)
        self.setCentralWidget(scroll)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(18)

        title = TitleLabel('软件著作权源代码文档生成器')
        subtitle = BodyLabel('支持多目录扫描、注释过滤、模板与排版参数配置')
        title.setStyleSheet('font-size: 22px;')
        subtitle.setStyleSheet('font-size: 12px;')
        subtitle.setWordWrap(True)
        header_row = QWidget()
        header_layout = QHBoxLayout(header_row)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.addWidget(title)
        header_layout.addStretch(1)
        layout.addWidget(header_row)
        layout.addWidget(subtitle)

        source_group, source_layout = self._create_group('源码设置')
        source_grid = QGridLayout()
        source_grid.setHorizontalSpacing(12)
        source_grid.setVerticalSpacing(12)
        source_grid.setColumnStretch(1, 1)
        source_grid.setColumnMinimumWidth(0, 110)
        source_layout.addLayout(source_grid)

        self.indirs_edit = TextEdit()
        self.indirs_edit.setPlaceholderText('每行一个目录')
        self.indirs_edit.setText(os.getcwd())
        self.indirs_edit.setMinimumHeight(96)
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
        self.exts_edit.setText('py')
        self.exts_edit.setMinimumHeight(32)
        source_grid.addWidget(BodyLabel('文件后缀'), 1, 0)
        source_grid.addWidget(self.exts_edit, 1, 1)

        self.comment_chars_edit = LineEdit()
        self.comment_chars_edit.setPlaceholderText('如 #, //')
        self.comment_chars_edit.setText('#, //')
        self.comment_chars_edit.setMinimumHeight(32)
        source_grid.addWidget(BodyLabel('注释前缀'), 2, 0)
        source_grid.addWidget(self.comment_chars_edit, 2, 1)

        self.excludes_edit = TextEdit()
        self.excludes_edit.setPlaceholderText('每行一个排除路径')
        self.excludes_edit.setMinimumHeight(96)
        exclude_buttons = QHBoxLayout()
        exclude_buttons.setSpacing(8)
        self.add_exclude_dir_btn = PushButton('添加排除目录')
        self.add_exclude_file_btn = PushButton('添加排除文件')
        self.clear_exclude_btn = PushButton('清空')
        self.add_exclude_dir_btn.setMinimumHeight(32)
        self.add_exclude_file_btn.setMinimumHeight(32)
        self.clear_exclude_btn.setMinimumHeight(32)
        exclude_buttons.addWidget(self.add_exclude_dir_btn)
        exclude_buttons.addWidget(self.add_exclude_file_btn)
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
        self.encoding_combo.addItems(['utf-8', 'utf-8-sig', 'gbk', 'gb18030'])
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
        output_grid.addWidget(BodyLabel('页眉标题'), 0, 0)
        output_grid.addWidget(self.title_edit, 0, 1)

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

        style_group, style_layout = self._create_group('排版设置')
        style_grid = QGridLayout()
        style_grid.setHorizontalSpacing(12)
        style_grid.setVerticalSpacing(12)
        style_grid.setColumnStretch(1, 1)
        style_grid.setColumnMinimumWidth(0, 110)
        style_layout.addLayout(style_grid)

        self.font_name_edit = LineEdit()
        self.font_name_edit.setText('宋体')
        self.font_name_edit.setMinimumHeight(32)
        style_grid.addWidget(BodyLabel('字体'), 0, 0)
        style_grid.addWidget(self.font_name_edit, 0, 1)

        self.font_size_spin = DoubleSpinBox()
        self.font_size_spin.setRange(1.0, 72.0)
        self.font_size_spin.setValue(10.5)
        self.font_size_spin.setMinimumHeight(32)
        style_grid.addWidget(BodyLabel('字号'), 1, 0)
        style_grid.addWidget(self.font_size_spin, 1, 1)

        self.space_before_spin = DoubleSpinBox()
        self.space_before_spin.setRange(0.0, 72.0)
        self.space_before_spin.setValue(0.0)
        self.space_before_spin.setMinimumHeight(32)
        style_grid.addWidget(BodyLabel('段前间距'), 2, 0)
        style_grid.addWidget(self.space_before_spin, 2, 1)

        self.space_after_spin = DoubleSpinBox()
        self.space_after_spin.setRange(0.0, 72.0)
        self.space_after_spin.setValue(2.3)
        self.space_after_spin.setMinimumHeight(32)
        style_grid.addWidget(BodyLabel('段后间距'), 3, 0)
        style_grid.addWidget(self.space_after_spin, 3, 1)

        self.line_spacing_spin = DoubleSpinBox()
        self.line_spacing_spin.setRange(0.0, 72.0)
        self.line_spacing_spin.setValue(10.5)
        self.line_spacing_spin.setMinimumHeight(32)
        style_grid.addWidget(BodyLabel('行距'), 4, 0)
        style_grid.addWidget(self.line_spacing_spin, 4, 1)

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

        action_row = QWidget()
        action_layout = QHBoxLayout(action_row)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(8)
        self.scan_btn = PushButton('扫描文件数')
        self.generate_btn = PrimaryPushButton('生成文档')
        self.open_output_btn = PushButton('打开输出目录')
        self.scan_btn.setMinimumHeight(34)
        self.generate_btn.setMinimumHeight(34)
        self.open_output_btn.setMinimumHeight(34)
        action_layout.addWidget(self.scan_btn)
        action_layout.addWidget(self.generate_btn)
        action_layout.addWidget(self.open_output_btn)
        action_layout.addStretch(1)

        action_card = CardWidget()
        action_card_layout = QVBoxLayout(action_card)
        action_card_layout.setContentsMargins(16, 16, 16, 16)
        action_card_layout.setSpacing(10)
        action_card_layout.addWidget(action_row)
        self.status_label = BodyLabel('')
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet('font-size: 12px;')
        action_card_layout.addWidget(self.status_label)
        layout.addWidget(action_card)
        layout.addStretch(1)

        self.add_indir_btn.clicked.connect(self.add_indir)
        self.clear_indir_btn.clicked.connect(lambda: self.indirs_edit.setText(''))
        self.add_exclude_dir_btn.clicked.connect(self.add_exclude_dir)
        self.add_exclude_file_btn.clicked.connect(self.add_exclude_file)
        self.clear_exclude_btn.clicked.connect(lambda: self.excludes_edit.setText(''))
        self.outfile_btn.clicked.connect(self.choose_outfile)
        self.template_btn.clicked.connect(self.choose_template)
        self.template_clear_btn.clicked.connect(lambda: self.template_edit.setText(''))
        self.reset_style_btn.clicked.connect(self.reset_style_defaults)
        self.generate_btn.clicked.connect(lambda: self.start_worker('generate'))
        self.scan_btn.clicked.connect(lambda: self.start_worker('scan'))
        self.open_output_btn.clicked.connect(self.open_output_dir)
        self.outfile_edit.textChanged.connect(self._update_open_output_enabled)
        self._update_open_output_enabled()

    def _create_group(self, title):
        group = CardWidget()
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        header = BodyLabel(title)
        header.setStyleSheet('font-size: 14px; font-weight: 600;')
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

    def choose_outfile(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, '选择输出文件', self.outfile_edit.text() or os.getcwd(), 'Word 文档 (*.docx)'
        )
        if file_path:
            if not file_path.lower().endswith('.docx'):
                file_path += '.docx'
            self.outfile_edit.setText(file_path)

    def choose_template(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择模板文件', os.getcwd(), 'Word 文档 (*.docx)'
        )
        if file_path:
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
        QDesktopServices.openUrl(QUrl.fromLocalFile(directory))

    def _append_line(self, widget, text):
        current = widget.toPlainText().strip()
        if current:
            widget.setText(current + '\n' + text)
        else:
            widget.setText(text)

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
        return {
            'title': title,
            'indirs': indirs,
            'exts': exts,
            'comment_chars': comment_chars,
            'font_name': self.font_name_edit.text().strip() or '宋体',
            'font_size': self.font_size_spin.value(),
            'space_before': self.space_before_spin.value(),
            'space_after': self.space_after_spin.value(),
            'line_spacing': self.line_spacing_spin.value(),
            'excludes': excludes,
            'outfile': outfile,
            'template_path': template_path,
            'skip_blank_lines': self.skip_blank_check.isChecked(),
            'skip_comment_lines': self.skip_comment_check.isChecked(),
            'encoding': encoding
        }

    def start_worker(self, mode):
        if self.worker and self.worker.isRunning():
            self._notify('warning', '任务进行中', '请等待当前任务完成后再操作')
            return
        config = self.build_config()
        valid, message = self._validate_inputs(config, mode)
        if not valid:
            self._notify('warning', '请检查输入', message)
            return
        self.set_buttons_enabled(False)
        if mode == 'scan':
            self.status_label.setText('正在扫描文件，请稍候...')
        else:
            self.status_label.setText('正在生成文档，请稍候...')
        self.worker = GenerateWorker(config, mode)
        self.worker.finished.connect(self.handle_finished)
        self.worker.failed.connect(self.handle_failed)
        self.worker.start()

    def handle_finished(self, result):
        self.set_buttons_enabled(True)
        if result.get('mode') == 'scan':
            self.status_label.setText('扫描完成，共找到 {} 个文件'.format(result.get('file_count', 0)))
            InfoBar.success(
                title='扫描完成',
                content='共找到 {} 个文件'.format(result.get('file_count', 0)),
                duration=3000,
                position=InfoBarPosition.TOP,
                parent=self
            )
            return
        self.status_label.setText('生成完成，共写入 {} 个文件'.format(result.get('file_count', 0)))
        InfoBar.success(
            title='生成完成',
            content='输出文件：{}'.format(result.get('outfile', '')),
            duration=4000,
            position=InfoBarPosition.TOP,
            parent=self
        )
        self._update_open_output_enabled()

    def handle_failed(self, message):
        self.set_buttons_enabled(True)
        self.status_label.setText('失败：{}'.format(message))
        InfoBar.error(
            title='生成失败',
            content=message,
            duration=5000,
            position=InfoBarPosition.TOP,
            parent=self
        )

    def set_buttons_enabled(self, enabled):
        self.scan_btn.setEnabled(enabled)
        self.generate_btn.setEnabled(enabled)
        self.open_output_btn.setEnabled(enabled)

    def reset_style_defaults(self):
        self.font_name_edit.setText('宋体')
        self.font_size_spin.setValue(10.5)
        self.space_before_spin.setValue(0.0)
        self.space_after_spin.setValue(2.3)
        self.line_spacing_spin.setValue(10.5)

    def _notify(self, level, title, content):
        if level == 'success':
            InfoBar.success(title=title, content=content, duration=3500, position=InfoBarPosition.TOP, parent=self)
            return
        if level == 'error':
            InfoBar.error(title=title, content=content, duration=4500, position=InfoBarPosition.TOP, parent=self)
            return
        InfoBar.warning(title=title, content=content, duration=3500, position=InfoBarPosition.TOP, parent=self)

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


def launch_gui():
    app = QApplication(sys.argv)
    window = GeneratorWindow()
    window.show()
    sys.exit(app.exec_())
