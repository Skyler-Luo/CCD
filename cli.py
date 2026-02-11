# -*- coding: utf-8 -*-
import logging

import click

from core import (
    DEFAULT_COMMENT_CHARS, DEFAULT_EXTS, DEFAULT_INDIRS, generate_code_doc
)
from gui import launch_gui


@click.command(name='ccd')
@click.option(
    '-t', '--title', default='软件著作权程序鉴别材料生成器V1.0',
    help='软件名称+版本号，默认为软件著作权程序鉴别材料生成器V1.0，此名称用于生成页眉'
)
@click.option(
    '-i', '--indir', 'indirs',
    multiple=True, type=click.Path(exists=True),
    help='源码所在文件夹，可以指定多个，默认为当前目录'
)
@click.option(
    '-e', '--ext', 'exts',
    multiple=True, help='源代码后缀，可以指定多个，默认为Python源代码'
)
@click.option(
    '-c', '--comment-char', 'comment_chars',
    multiple=True, help='注释字符串，可以指定多个，默认为#、//'
)
@click.option(
    '--font-name', default='宋体',
    help='字体，默认为宋体'
)
@click.option(
    '--font-size', default=10.5,
    type=click.FloatRange(min=1.0),
    help='字号，默认为五号，即10.5号'
)
@click.option(
    '--space-before', default=0.0,
    type=click.FloatRange(min=0.0),
    help='段前间距，默认为0'
)
@click.option(
    '--space-after', default=2.3,
    type=click.FloatRange(min=0.0),
    help='段后间距，默认为2.3'
)
@click.option(
    '--line-spacing', default=10.5,
    type=click.FloatRange(min=0.0),
    help='行距，默认为固定值10.5'
)
@click.option(
    '--exclude', 'excludes',
    multiple=True, type=click.Path(exists=True),
    help='需要排除的文件或路径，可以指定多个'
)
@click.option(
    '-o', '--outfile', default='code.docx',
    type=click.Path(exists=False),
    help='输出文件（docx格式），默认为当前目录的code.docx'
)
@click.option(
    '--template', 'template_path', default=None,
    type=click.Path(exists=True),
    help='docx模板文件路径'
)
@click.option(
    '--encoding', default='utf-8',
    help='源码文件编码，默认为utf-8'
)
@click.option(
    '--keep-blank-lines', is_flag=True,
    help='保留空行'
)
@click.option(
    '--keep-comment-lines', is_flag=True,
    help='保留注释行'
)
@click.option(
    '--gui', is_flag=True,
    help='启动图形界面'
)
@click.option('-v', '--verbose', is_flag=True, help='打印调试信息')
def main(
        title, indirs, exts,
        comment_chars, font_name,
        font_size, space_before,
        space_after, line_spacing,
        excludes, outfile, template_path,
        encoding, keep_blank_lines,
        keep_comment_lines, gui, verbose
):
    if gui:
        launch_gui()
        return 0
    if not indirs:
        indirs = DEFAULT_INDIRS
    if not exts:
        exts = DEFAULT_EXTS
    if not comment_chars:
        comment_chars = DEFAULT_COMMENT_CHARS
    if verbose:
        logging.basicConfig(level=logging.DEBUG)
    generate_code_doc(
        title=title,
        indirs=indirs,
        exts=exts,
        comment_chars=comment_chars,
        font_name=font_name,
        font_size=font_size,
        space_before=space_before,
        space_after=space_after,
        line_spacing=line_spacing,
        excludes=excludes,
        outfile=outfile,
        template_path=template_path,
        skip_blank_lines=not keep_blank_lines,
        skip_comment_lines=not keep_comment_lines,
        encoding=encoding
    )
    return 0


if __name__ == '__main__':
    main()
