# -*- coding: utf-8 -*-
import codecs
import logging
from os.path import abspath
try:
    from os import scandir
except ImportError:
    from scandir import scandir

logger = logging.getLogger(__name__)

DEFAULT_INDIRS = ['.']
DEFAULT_EXTS = ['py']
DEFAULT_COMMENT_CHARS = (
    '#', '//'
)


def del_slash(dirs):
    """
    删除文件夹最后一位的/

    Args:
        dirs: 文件夹列表
    Returns:
        删除之后的文件夹
    """
    no_slash_dirs = []
    for dir_ in dirs:
        if dir_[-1] == '/':
            no_slash_dirs.append(dir_[: -1])
        else:
            no_slash_dirs.append(dir_)
    return no_slash_dirs


class CodeFinder(object):
    """
    给定一个目录，和若干个后缀名，
    递归地遍历该目录，找到该目录下
    所有以这些后缀结束的文件
    """
    def __init__(self, exts=None):
        """
        Args:
            exts: 后缀名，默认为以py结尾
        """
        self.exts = exts if exts else ['py']

    def is_code(self, file):
        for ext in self.exts:
            if file.endswith(ext):
                return True
        return False

    @staticmethod
    def is_hidden_file(file):
        """
        是否是隐藏文件
        """
        return file[0] == '.'

    @staticmethod
    def should_be_excluded(file, excludes=None):
        """
        是否需要略过此文件
        """
        if not excludes:
            return False
        if not isinstance(excludes, list):
            excludes = [excludes]
        should_be_excluded = False
        for exclude in excludes:
            if file.startswith(exclude):
                should_be_excluded = True
                break
        return should_be_excluded

    def find(self, indir, excludes=None):
        """
        给定一个文件夹查找这个文件夹下所有的代码

        Args:
            indir: 需要查到代码的目录
            excludes: 排除文件或目录
        Returns:
            代码文件列表
        """
        files = []
        for entry in scandir(indir):
            entry_name = entry.name
            entry_path = abspath(entry.path)
            if self.is_hidden_file(entry_name):
                continue
            if self.should_be_excluded(entry_path, excludes):
                continue
            if entry.is_file():
                if self.is_code(entry_name):
                    files.append(entry_path)
                continue
            for file in self.find(entry_path, excludes=excludes):
                files.append(file)
        logger.debug('在%s目录下找到%d个代码文件.', indir, len(files))
        return files


class CodeWriter(object):
    def __init__(
            self, font_name='宋体',
            font_size=10.5, space_before=0.0,
            space_after=2.3, line_spacing=10.5,
            command_chars=None, document=None,
            template_path=None, skip_blank_lines=True,
            skip_comment_lines=True, encoding='utf-8'
    ):
        Document, Pt, WD_PARAGRAPH_ALIGNMENT = load_docx_dependencies()
        self.font_name = font_name
        self.font_size = font_size
        self.space_before = space_before
        self.space_after = space_after
        self.line_spacing = line_spacing
        self.command_chars = command_chars if command_chars else DEFAULT_COMMENT_CHARS
        self.skip_blank_lines = skip_blank_lines
        self.skip_comment_lines = skip_comment_lines
        self.encoding = encoding
        self._Pt = Pt
        self._WD_PARAGRAPH_ALIGNMENT = WD_PARAGRAPH_ALIGNMENT
        self.document = document if document else create_document(template_path, Document)

    @staticmethod
    def is_blank_line(line):
        """
        判断是否是空行
        """
        return not bool(line)

    def is_comment_line(self, line):
        line = line.lstrip()
        is_comment = False
        for comment_char in self.command_chars:
            if line.startswith(comment_char):
                is_comment = True
                break
        return is_comment

    def write_header(self, title):
        """
        写入页眉
        """
        paragraph = self.document.sections[0].header.paragraphs[0]
        paragraph.alignment = self._WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.add_run(title)
        run.font.name = self.font_name
        run.font.size = self._Pt(self.font_size)
        return self

    def write_file(self, file):
        """
        把单个文件添加到程序文档里面
        """
        with codecs.open(file, encoding=self.encoding, errors='ignore') as fp:
            for line in fp:
                line = line.rstrip()
                if self.skip_blank_lines and self.is_blank_line(line):
                    continue
                if self.skip_comment_lines and self.is_comment_line(line):
                    continue
                paragraph = self.document.add_paragraph()
                paragraph.paragraph_format.space_before = self._Pt(self.space_before)
                paragraph.paragraph_format.space_after = self._Pt(self.space_after)
                paragraph.paragraph_format.line_spacing = self._Pt(self.line_spacing)
                run = paragraph.add_run(line)
                run.font.name = self.font_name
                run.font.size = self._Pt(self.font_size)
        return self

    def save(self, file):
        self.document.save(file)


def load_docx_dependencies():
    try:
        from docx import Document
        from docx.shared import Pt
        from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
        return Document, Pt, WD_PARAGRAPH_ALIGNMENT
    except Exception as exc:
        raise RuntimeError('未检测到可用的python-docx，请安装python-docx并卸载docx包') from exc


def create_document(template_path, Document):
    if template_path:
        return Document(template_path)
    return Document()


def normalize_items(text):
    items = []
    for line in text.splitlines():
        for part in line.replace(';', ',').split(','):
            item = part.strip()
            if item:
                items.append(item)
    return items


def normalize_exts(exts):
    items = []
    for ext in exts:
        item = ext.strip()
        if not item:
            continue
        if item.startswith('.'):
            item = item[1:]
        items.append(item)
    return items


def collect_code_files(indirs, exts, excludes):
    finder = CodeFinder(exts)
    files = []
    for indir in indirs:
        files.extend(finder.find(indir, excludes=excludes))
    return files


def generate_code_doc(
        title, indirs, exts, comment_chars,
        font_name, font_size, space_before,
        space_after, line_spacing, excludes,
        outfile, template_path=None,
        skip_blank_lines=True, skip_comment_lines=True,
        encoding='utf-8'
):
    if not indirs:
        indirs = DEFAULT_INDIRS
    if not exts:
        exts = DEFAULT_EXTS
    if not comment_chars:
        comment_chars = DEFAULT_COMMENT_CHARS
    indirs = [abspath(indir) for indir in indirs]
    excludes = del_slash(
        [abspath(exclude) for exclude in excludes] if excludes else []
    )
    files = collect_code_files(indirs, exts, excludes)
    writer = CodeWriter(
        command_chars=comment_chars,
        font_name=font_name,
        font_size=font_size,
        space_before=space_before,
        space_after=space_after,
        line_spacing=line_spacing,
        template_path=template_path,
        skip_blank_lines=skip_blank_lines,
        skip_comment_lines=skip_comment_lines,
        encoding=encoding
    )
    writer.write_header(title)
    for file in files:
        writer.write_file(file)
    writer.save(outfile)
    return {'file_count': len(files), 'outfile': outfile}
