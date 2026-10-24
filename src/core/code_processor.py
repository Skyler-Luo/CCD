# -*- coding: utf-8 -*-
"""
代码处理工具模块
包含常量定义和通用工具函数
"""
import ast
import codecs
import io
import logging
import os
import re
import tokenize
from os.path import abspath

logger = logging.getLogger(__name__)

# ==================== 常量定义 ====================

DEFAULT_INDIRS = ['.']
DEFAULT_EXTS = ['py']
DEFAULT_COMMENT_CHARS = ('#', '//')
DEFAULT_SKIP_DIRS = [
    'node_modules', 'uni_modules', '__pycache__', '.git', '.venv', 'venv'
]
DEFAULT_SKIP_FILES = [
    'package.json', 'package-lock.json', 'pnpm-lock.yaml', 'yarn.lock'
]

LANGUAGE_BY_EXT = {
    'py': 'python',
    'js': 'javascript',
    'jsx': 'javascript',
    'ts': 'typescript',
    'tsx': 'typescript',
    'go': 'go',
    'php': 'php',
    'cs': 'csharp',
    'kt': 'kotlin',
    'kts': 'kotlin',
    'swift': 'swift',
    'rs': 'rust',
    'dart': 'dart',
    'scala': 'scala',
    'sql': 'sql',
    'r': 'r',
    'lua': 'lua',
    'ps1': 'powershell',
    'psm1': 'powershell',
    'json': 'json',
    'yaml': 'yaml',
    'yml': 'yaml',
    'java': 'java',
    'c': 'c',
    'h': 'c',
    'cpp': 'cpp',
    'hpp': 'cpp',
    'cc': 'cpp',
    'html': 'html',
    'htm': 'html',
    'xml': 'xml',
    'css': 'css',
    'sh': 'shellscript',
    'bash': 'shellscript',
    'rb': 'ruby',
    'pl': 'perl'
}

COMMENT_REGEX_BY_LANG = {
    'javascript': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'typescript': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'go': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'php': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|//.*|#.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'csharp': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'kotlin': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'swift': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'rust': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'dart': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'scala': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'sql': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|--.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'r': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|#.*', re.MULTILINE),
    'lua': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|--\\[\\[[\\s\\S]*?\\]\\]|--.*', re.MULTILINE),
    'powershell': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|#.*|<#[\\s\\S]*?#>', re.MULTILINE),
    'yaml': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\')|#.*', re.MULTILINE),
    'java': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'c': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'cpp': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|//.*|/\*[\s\S]*?\*/', re.MULTILINE),
    'html': re.compile(r'<!--[\s\S]*?-->', re.MULTILINE),
    'xml': re.compile(r'<!--[\s\S]*?-->', re.MULTILINE),
    'css': re.compile(r'/\*[\s\S]*?\*/', re.MULTILINE),
    'shellscript': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|#.*|=begin[\s\S]*?=end', re.MULTILINE),
    'ruby': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|#.*|=begin[\s\S]*?=end', re.MULTILINE),
    'perl': re.compile(r'("(?:(?:\\.|[^"\\])*)"|\'(?:(?:\\.|[^\'\\])*)\'|`(?:(?:\\.|[^`\\])*)`)|#.*|=begin[\s\S]*?=end', re.MULTILINE)
}

# ==================== 路径处理工具 ====================

def normalize_paths(paths):
    """将路径转为绝对路径并去重（大小写不敏感）"""
    if not paths:
        return []
    normalized = [abspath(item).rstrip(os.sep) for item in paths if item]
    seen = set()
    unique = []
    for item in normalized:
        key = os.path.normcase(item)
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


def read_gitignore_excludes(indirs):
    """读取.gitignore并提取静态路径排除项（忽略通配符和否定规则）"""
    excludes = []
    for indir in indirs:
        root = abspath(indir)
        if not os.path.isdir(root):
            continue
        for current_root, dirs, files in os.walk(root):
            # 过滤掉不需要遍历的目录以优化性能，例如 .git, node_modules 等
            dirs[:] = [d for d in dirs if d not in DEFAULT_SKIP_DIRS and not d.startswith('.')]
            if '.gitignore' not in files:
                continue
            gitignore_path = os.path.join(current_root, '.gitignore')
            try:
                with codecs.open(gitignore_path, encoding='utf-8', errors='ignore') as fp:
                    for raw in fp:
                        line = raw.strip().replace('\\', '/')
                        if not line or line.startswith(('#', '!')):
                            continue
                        if any(ch in line for ch in ('*', '?', '[')):
                            continue
                        line = line.strip('/')
                        if line:
                            path = abspath(os.path.join(current_root, line))
                            if os.path.exists(path):
                                excludes.append(path)
            except OSError:
                continue
    return normalize_paths(excludes)

# ==================== 文件处理工具 ====================

def is_binary_file(file_path):
    """粗略判断文件是否为二进制文件"""
    try:
        with open(file_path, 'rb') as fp:
            chunk = fp.read(2048)
        return b'\x00' in chunk
    except OSError:
        return True


def decode_content(file_path, encoding):
    """读取源码文件并解码（支持auto自动检测）"""
    if encoding and encoding != 'auto':
        try:
            with codecs.open(file_path, encoding=encoding) as fp:
                return fp.read()
        except UnicodeDecodeError as exc:
            raise UnicodeError(f'文件 {file_path} 无法按 {encoding} 解码，请检查编码设置') from exc
    try:
        with open(file_path, 'rb') as fp:
            data = fp.read()
        for name in ('utf-8-sig', 'utf-8', 'gb18030', 'gbk'):
            try:
                return data.decode(name)
            except UnicodeDecodeError as exc:
                last_error = exc
                continue
        raise UnicodeError(f'文件 {file_path} 无法自动识别编码，请手动指定编码') from last_error
    except OSError:
        raise

# ==================== 语言和注释处理 ====================

def get_language_by_extension(file_path):
    """根据文件后缀推断语言标识"""
    ext = os.path.splitext(file_path)[1].lower().lstrip('.')
    return LANGUAGE_BY_EXT.get(ext, '')


def strip_comments(content, language):
    """按语言规则移除注释内容（尽量保留字符串字面量）"""
    if language == 'python':
        try:
            tokens = list(tokenize.generate_tokens(io.StringIO(content).readline))
        except (SyntaxError, tokenize.TokenError):
            return content

        spans = [
            (token.start, token.end)
            for token in tokens
            if token.type == tokenize.COMMENT
        ]
        try:
            tree = ast.parse(content)
        except SyntaxError:
            tree = None
        if tree is not None:
            source_lines = content.splitlines()
            nodes = [tree]
            nodes.extend(node for node in ast.walk(tree)
                         if isinstance(node, (ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)))
            for node in nodes:
                if not node.body:
                    continue
                first = node.body[0]
                if (isinstance(first, ast.Expr)
                        and isinstance(first.value, ast.Constant)
                        and isinstance(first.value.value, str)):
                    start_line = source_lines[first.lineno - 1]
                    end_line = source_lines[first.end_lineno - 1]
                    start_col = len(start_line.encode('utf-8')[:first.col_offset].decode('utf-8'))
                    end_col = len(end_line.encode('utf-8')[:first.end_col_offset].decode('utf-8'))
                    spans.append(((first.lineno, start_col), (first.end_lineno, end_col)))

        lines = content.splitlines(keepends=True)
        for (start_line, start_col), (end_line, end_col) in sorted(spans, reverse=True):
            start_index = start_line - 1
            end_index = end_line - 1
            if start_index == end_index:
                lines[start_index] = lines[start_index][:start_col] + lines[start_index][end_col:]
            else:
                first_line = lines[start_index]
                ending = '\r\n' if first_line.endswith('\r\n') else first_line[-1:] if first_line.endswith(('\n', '\r')) else ''
                lines[start_index] = first_line[:start_col] + ending
                for index in range(start_index + 1, end_index):
                    line = lines[index]
                    lines[index] = '\r\n' if line.endswith('\r\n') else line[-1:] if line.endswith(('\n', '\r')) else ''
                lines[end_index] = lines[end_index][end_col:]
        return ''.join(lines)
    
    pattern = COMMENT_REGEX_BY_LANG.get(language)
    if not pattern:
        return content
    if language in ('html', 'xml', 'css'):
        return pattern.sub('', content)
    return pattern.sub(lambda m: m.group(1) if m.group(1) else '', content)


def filter_lines(content, language, skip_blank_lines, skip_comment_lines, comment_chars):
    """按规则过滤源码内容为需要写入文档的行列表"""
    text = content
    if skip_comment_lines:
        if language:
            text = strip_comments(text, language)
        else:
            lines = [line for line in text.splitlines() 
                    if not any(line.lstrip().startswith(ch) for ch in comment_chars)]
            text = '\n'.join(lines)
    
    lines = text.splitlines()
    if skip_blank_lines:
        lines = [line for line in lines if line.strip()]
    return lines

# ==================== 文本处理工具 ====================

def normalize_items(text):
    """将多行/逗号/分号分隔的文本整理为字符串列表"""
    items = []
    for line in text.splitlines():
        items.extend(part.strip() for part in line.replace(';', ',').split(',') if part.strip())
    return items


def normalize_exts(exts):
    """规范化后缀列表（去空白、去掉前导 '.'）"""
    return [ext.strip().lstrip('.') for ext in exts if ext.strip()]
