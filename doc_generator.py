# -*- coding: utf-8 -*-
"""
文档生成主流程模块
负责协调整个文档生成过程
"""
from os.path import abspath

from code_processor import (
    DEFAULT_INDIRS, DEFAULT_EXTS, DEFAULT_COMMENT_CHARS,
    DEFAULT_SKIP_DIRS, DEFAULT_SKIP_FILES,
    normalize_paths, normalize_exts,
    decode_content, get_language_by_extension, filter_lines
)
from code_scanner import collect_code_files, count_total_lines
from doc_writer import CodeWriter


def generate_code_doc(
        title, indirs, exts, comment_chars,
        font_name, font_size, space_before,
        space_after, line_spacing, excludes,
        outfile, template_path=None,
        skip_blank_lines=True, skip_comment_lines=True,
        encoding='utf-8', skip_dir_names=None, skip_file_names=None,
        header_font='宋体', header_font_size=10.5,
        page_number_format='第 {page} 页 共 {total} 页',
        selected_files=None, pages_60_mode=False
):
    """
    生成 docx 源代码文档
    
    Args:
        title: 页眉标题
        indirs: 源码目录列表
        exts: 后缀列表
        comment_chars: 注释前缀列表（language 未识别时使用）
        font_name/font_size/space_before/space_after/line_spacing: 排版参数
        excludes: 排除路径列表
        outfile: 输出 docx 路径
        template_path: 模板 docx 路径
        skip_blank_lines: 是否过滤空行
        skip_comment_lines: 是否过滤注释
        encoding: 源码文件编码；支持 'auto'
        skip_dir_names/skip_file_names: 跳过目录名/文件名列表
        header_font: 页眉字体
        header_font_size: 页眉字号
        page_number_format: 页码格式
        selected_files: 用户选择的文件列表（如果为 None 则使用所有文件）
        pages_60_mode: 是否启用60页模式（前30页+后30页）
        
    Returns:
        dict：包含 file_count、outfile、total_lines
    """
    if not indirs:
        indirs = DEFAULT_INDIRS
    if not exts:
        exts = DEFAULT_EXTS
    else:
        exts = normalize_exts(exts)
    if not comment_chars:
        comment_chars = DEFAULT_COMMENT_CHARS
    indirs = [abspath(indir) for indir in indirs]
    excludes = normalize_paths(excludes)
    if outfile:
        excludes = normalize_paths(excludes + [outfile])
    if not skip_dir_names:
        skip_dir_names = DEFAULT_SKIP_DIRS
    if not skip_file_names:
        skip_file_names = DEFAULT_SKIP_FILES
    
    files = collect_code_files(indirs, exts, excludes, skip_dir_names, skip_file_names)
    
    # 如果指定了 selected_files，则只使用选中的文件，并保持其顺序
    if selected_files:
        # 创建文件集合用于快速查找
        available_files_set = set(files)
        # 按照 selected_files 的顺序，只保留实际存在的文件
        files = [f for f in selected_files if f in available_files_set]
    
    # 统计总行数
    total_lines_info = count_total_lines(files, skip_blank_lines, skip_comment_lines, comment_chars, encoding)
    total_lines = total_lines_info['total_lines']
    
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
        encoding=encoding,
        header_font=header_font,
        header_font_size=header_font_size,
        page_number_format=page_number_format
    )
    writer.write_header(title)
    writer.write_footer()
    
    # 60页模式：只写入前30页和后30页（假设每页50行）
    if pages_60_mode and total_lines >= 3000:
        lines_per_page = 50
        first_30_lines = lines_per_page * 30  # 1500行
        last_30_lines = lines_per_page * 30   # 1500行
        
        current_line = 0
        for file in files:
            content = decode_content(file, encoding)
            language = get_language_by_extension(file)
            lines = filter_lines(
                content, language,
                skip_blank_lines,
                skip_comment_lines,
                comment_chars
            )
            
            for line in lines:
                if current_line < first_30_lines or current_line >= (total_lines_info['total_lines'] - last_30_lines):
                    paragraph = writer.document.add_paragraph()
                    paragraph.paragraph_format.space_before = writer._Pt(space_before)
                    paragraph.paragraph_format.space_after = writer._Pt(space_after)
                    paragraph.paragraph_format.line_spacing = writer._Pt(line_spacing)
                    run = paragraph.add_run(line.rstrip())
                    run.font.name = font_name
                    run.font.size = writer._Pt(font_size)
                current_line += 1
    else:
        # 正常模式：写入所有文件
        for file in files:
            writer.write_file(file)
    
    writer.save(outfile)
    return {
        'file_count': len(files),
        'outfile': outfile,
        'total_lines': total_lines_info
    }
