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
from src.core.word_com_helper import check_word_running, kill_word_process
from src.ui.ui_tasks import GenerateWorker, ExtensionScanWorker
from src.ui.ui_layout import GeneratorWindowUI
from src.ui.dialogs import (
    ExtensionSelectDialog, CommentPrefixDialog, FileSelectDialog, LogViewDialog,
    COMMENT_PREFIX_BY_LANG
)
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
        
        # 装配 UI 布局
        self.ui = GeneratorWindowUI()
        self.ui.setup_ui(self)
        
        setTheme(Theme.AUTO)
        self.apply_theme_stylesheet()

    
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
        config_dict = {
            'skip_blank_lines': skip_blank_lines,
            'skip_comment_lines': skip_comment_lines,
            'comment_chars': comment_chars,
            'encoding': encoding
        }
        dialog = FileSelectDialog(
            self.available_files, 
            self.selected_files,
            config_dict,
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

        # 仅检查 winword.exe 是否运行
        word_running = check_word_running()

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
                killed_ok = kill_word_process()
                
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
