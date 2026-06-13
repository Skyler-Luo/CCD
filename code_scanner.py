# -*- coding: utf-8 -*-
"""
代码文件扫描模块
负责扫描目录并查找符合条件的代码文件
"""
import logging
import os
from os.path import abspath
try:
    from os import scandir
except ImportError:
    from scandir import scandir

from code_processor import (
    is_binary_file, decode_content, get_language_by_extension,
    strip_comments, filter_lines
)

logger = logging.getLogger(__name__)


class CodeFinder:
    """
    递归扫描目录，收集指定后缀的代码文件
    
    会跳过：
        - 隐藏文件/目录（以 . 开头）
        - 指定的目录名/文件名
        - excludes 命中的文件或目录
        - 疑似二进制文件
    """
    def __init__(self, exts=None, skip_dir_names=None, skip_file_names=None):
        """初始化代码查找器"""
        self.exts = exts or ['py']
        self.skip_dir_names = skip_dir_names or []
        self.skip_file_names = skip_file_names or []

    def is_code(self, file):
        """判断文件是否为代码文件"""
        lowered = file.lower()
        return any(lowered.endswith('.' + ext.lower().lstrip('.')) for ext in self.exts)

    @staticmethod
    def is_hidden_file(file):
        """是否是隐藏文件（以 '.' 开头）"""
        return file.startswith('.')

    @staticmethod
    def should_be_excluded(file, excludes=None):
        """判断路径是否在排除列表中（包含子路径）"""
        if not excludes:
            return False
        excludes = excludes if isinstance(excludes, list) else [excludes]
        for exclude in excludes:
            if file == exclude or file.startswith(exclude + os.sep):
                return True
        return False

    def find(self, indir, excludes=None):
        """
        查找目录下所有符合后缀的代码文件
        
        Args:
            indir: 需要扫描的目录
            excludes: 排除文件或目录（绝对路径）
            
        Returns:
            代码文件列表（绝对路径）
        """
        files = []
        for entry in scandir(indir):
            entry_name = entry.name
            entry_path = abspath(entry.path)
            if self.is_hidden_file(entry_name):
                continue
            if entry.is_dir() and entry_name in self.skip_dir_names:
                continue
            if entry.is_file() and entry_name in self.skip_file_names:
                continue
            if self.should_be_excluded(entry_path, excludes):
                continue
            if entry.is_file():
                if is_binary_file(entry_path):
                    continue
                if self.is_code(entry_name):
                    files.append(entry_path)
                continue
            for file in self.find(entry_path, excludes=excludes):
                files.append(file)
        logger.debug('在%s目录下找到%d个代码文件.', indir, len(files))
        return files


def collect_code_files(indirs, exts, excludes, skip_dir_names=None, skip_file_names=None):
    """
    收集所有代码文件路径
    
    Args:
        indirs: 源码目录列表
        exts: 后缀列表
        excludes: 排除路径列表（绝对路径）
        skip_dir_names: 跳过目录名列表
        skip_file_names: 跳过文件名列表
        
    Returns:
        文件路径列表（绝对路径）
    """
    finder = CodeFinder(exts, skip_dir_names=skip_dir_names, skip_file_names=skip_file_names)
    files = []
    for indir in indirs:
        files.extend(finder.find(indir, excludes=excludes))
    return files


def collect_all_file_extensions(indirs, excludes, skip_dir_names=None, skip_file_names=None):
    """
    扫描目录并收集出现过的文件后缀（用于 GUI 的"自动识别后缀"）
    
    Args:
        indirs: 源码目录列表
        excludes: 排除路径列表（绝对路径）
        skip_dir_names: 跳过目录名列表
        skip_file_names: 跳过文件名列表
        
    Returns:
        排序后的后缀列表
    """
    extensions = set()
    skip_dir_names = set(skip_dir_names or [])
    skip_file_names = set(skip_file_names or [])
    for indir in indirs:
        for root, dirs, files in os.walk(indir):
            root_path = abspath(root)
            if CodeFinder.should_be_excluded(root_path, excludes):
                dirs[:] = []
                continue
            dirs[:] = [d for d in dirs if not d.startswith('.') and d not in skip_dir_names]
            for name in files:
                if name.startswith('.'):
                    continue
                if name in skip_file_names:
                    continue
                file_path = abspath(os.path.join(root, name))
                if CodeFinder.should_be_excluded(file_path, excludes):
                    continue
                if is_binary_file(file_path):
                    continue
                ext = os.path.splitext(name)[1].lower().lstrip('.')
                if ext:
                    extensions.add(ext)
    return sorted(extensions)


def count_total_lines(files, skip_blank_lines, skip_comment_lines, comment_chars, encoding):
    """
    统计文件的总行数，并分别统计代码行数和注释行数
    
    Args:
        files: 文件路径列表
        skip_blank_lines: 是否过滤空行
        skip_comment_lines: 是否过滤注释
        comment_chars: 注释前缀列表
        encoding: 文件编码
    
    Returns:
        dict: 包含以下键值：
            - total_lines: 总行数（过滤后）
            - code_lines: 代码行数（非注释、非空行）
            - comment_lines: 注释行数
            - blank_lines: 空行数
            - raw_lines: 原始总行数（未过滤）
    """
    total_filtered = 0
    code_lines = 0
    comment_lines = 0
    blank_lines = 0
    raw_lines = 0
    
    for file in files:
        content = decode_content(file, encoding)
        language = get_language_by_extension(file)
        
        # 统计原始行数
        raw_file_lines = content.splitlines()
        raw_lines += len(raw_file_lines)
        
        # 统计注释行（去除注释后的行数差）
        if language:
            stripped = strip_comments(content, language)
            stripped_lines = stripped.splitlines()
            
            # 统计空行
            for line in raw_file_lines:
                if not line.strip():
                    blank_lines += 1
            
            # 统计注释行（原始行数 - 去除注释后的行数 - 空行）
            # 这是一个近似值，因为注释可能和代码在同一行
            lines_after_comment_removal = len([l for l in stripped_lines if l.strip()])
            file_blank = sum(1 for line in raw_file_lines if not line.strip())
            file_comment = len(raw_file_lines) - lines_after_comment_removal - file_blank
            comment_lines += file_comment
            
            # 统计代码行
            code_lines += lines_after_comment_removal
        else:
            # 无法识别语言时，使用简单的前缀匹配
            for line in raw_file_lines:
                stripped = line.strip()
                if not stripped:
                    blank_lines += 1
                elif any(stripped.startswith(ch) for ch in comment_chars):
                    comment_lines += 1
                else:
                    code_lines += 1
        
        # 统计过滤后的行数
        lines = filter_lines(
            content, language,
            skip_blank_lines,
            skip_comment_lines,
            comment_chars
        )
        total_filtered += len(lines)
    
    return {
        'total_lines': total_filtered,
        'code_lines': code_lines,
        'comment_lines': comment_lines,
        'blank_lines': blank_lines,
        'raw_lines': raw_lines
    }
