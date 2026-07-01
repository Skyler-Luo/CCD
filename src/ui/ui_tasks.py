# -*- coding: utf-8 -*-
"""
UI后台任务模块
负责在后台线程中执行耗时操作
"""
import os
from PyQt5.QtCore import QThread, pyqtSignal

from src.core.code_processor import normalize_paths, DEFAULT_SKIP_DIRS, DEFAULT_SKIP_FILES
from src.core.code_scanner import collect_code_files, collect_all_file_extensions, count_total_lines
from src.core.doc_generator import generate_code_doc
from src.utils.logger import get_logger


class GenerateWorker(QThread):
    """文档生成/扫描/统计后台任务"""
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, config, mode):
        super().__init__()
        self.config = config
        self.mode = mode
        self.logger = get_logger('Worker')

    def run(self):
        try:
            self.logger.info(f'Worker启动: mode={self.mode}')
            
            if self.mode == 'scan':
                indirs = [os.path.abspath(indir) for indir in self.config['indirs']]
                excludes = normalize_paths(self.config['excludes'])
                outfile = self.config.get('outfile')
                if outfile:
                    excludes = normalize_paths(excludes + [outfile])
                files = collect_code_files(
                    indirs,
                    self.config['exts'],
                    excludes,
                    DEFAULT_SKIP_DIRS,
                    DEFAULT_SKIP_FILES
                )
                self.logger.info(f'扫描完成: 找到 {len(files)} 个文件')
                self.finished.emit({
                    'mode': 'scan',
                    'file_count': len(files),
                    'files': files
                })
            elif self.mode == 'count':
                # 统计代码行数
                indirs = [os.path.abspath(indir) for indir in self.config['indirs']]
                excludes = normalize_paths(self.config['excludes'])
                outfile = self.config.get('outfile')
                if outfile:
                    excludes = normalize_paths(excludes + [outfile])
                files = collect_code_files(
                    indirs,
                    self.config['exts'],
                    excludes,
                    DEFAULT_SKIP_DIRS,
                    DEFAULT_SKIP_FILES
                )
                # 如果指定了选中的文件，只统计选中的
                if self.config.get('selected_files'):
                    files = [f for f in files if f in self.config['selected_files']]
                
                stats = count_total_lines(
                    files,
                    self.config['skip_blank_lines'],
                    self.config['skip_comment_lines'],
                    self.config['comment_chars'],
                    self.config['encoding']
                )
                self.logger.info(f'统计完成: {len(files)} 个文件, {stats["total_lines"]} 行')
                self.finished.emit({
                    'mode': 'count',
                    'file_count': len(files),
                    'stats': stats
                })
            else:
                self.logger.info(f'开始生成文档: {self.config["outfile"]}')
                result = generate_code_doc(**self.config)
                result['mode'] = 'generate'
                self.logger.info(f'文档生成成功: {result["file_count"]} 个文件, {result["total_lines"]["total_lines"]} 行')
                self.finished.emit(result)
        except Exception as exc:
            error_msg = str(exc)
            self.logger.exception(f'Worker执行失败: {error_msg}')
            self.failed.emit(error_msg)


class ExtensionScanWorker(QThread):
    """文件后缀扫描后台任务"""
    finished = pyqtSignal(list)
    failed = pyqtSignal(str)

    def __init__(self, indirs, excludes):
        super().__init__()
        self.indirs = indirs
        self.excludes = excludes
        self.logger = get_logger('ExtensionWorker')

    def run(self):
        try:
            self.logger.debug(f'开始扫描文件后缀: {len(self.indirs)} 个目录')
            exts = collect_all_file_extensions(
                self.indirs, self.excludes,
                DEFAULT_SKIP_DIRS, DEFAULT_SKIP_FILES
            )
            self.logger.debug(f'扫描完成: 找到 {len(exts)} 种后缀')
            self.finished.emit(exts)
        except Exception as exc:
            error_msg = str(exc)
            self.logger.error(f'后缀扫描失败: {error_msg}', exc_info=True)
            self.failed.emit(error_msg)


class FileInfoWorker(QThread):
    """文件详细信息异步统计任务（用于文件选择对话框）"""
    file_processed = pyqtSignal(str, dict)
    finished = pyqtSignal(dict)

    def __init__(self, files, skip_blank_lines, skip_comment_lines, comment_chars, encoding):
        super().__init__()
        self.files = files
        self.skip_blank_lines = skip_blank_lines
        self.skip_comment_lines = skip_comment_lines
        self.comment_chars = comment_chars
        self.encoding = encoding
        self.logger = get_logger('FileInfoWorker')

    def run(self):
        self.logger.info(f'开始异步统计 {len(self.files)} 个文件信息')
        cache = {}
        for file_path in self.files:
            if self.isInterruptionRequested():
                self.logger.info('收到中断请求，停止统计')
                break
            try:
                from src.core.code_processor import decode_content, get_language_by_extension, filter_lines
                info = {
                    'path': file_path,
                    'name': os.path.basename(file_path),
                    'size': os.path.getsize(file_path) if os.path.exists(file_path) else 0,
                    'lines': 0,
                    'filtered_lines': 0
                }
                if os.path.exists(file_path):
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
            except Exception as e:
                self.logger.error(f'处理文件出错 {file_path}: {e}')
                info = {
                    'path': file_path,
                    'name': os.path.basename(file_path),
                    'size': 0,
                    'lines': 0,
                    'filtered_lines': 0
                }
            cache[file_path] = info
            self.file_processed.emit(file_path, info)
        self.finished.emit(cache)

