# -*- coding: utf-8 -*-
"""
文档生成主流程模块
负责协调整个文档生成过程
"""
import os
import tempfile
from os.path import abspath

from code_processor import (
    DEFAULT_INDIRS, DEFAULT_EXTS, DEFAULT_COMMENT_CHARS,
    DEFAULT_SKIP_DIRS, DEFAULT_SKIP_FILES,
    normalize_paths, normalize_exts,
    decode_content, get_language_by_extension, filter_lines
)
from code_scanner import collect_code_files, count_total_lines
from doc_writer import CodeWriter
from logger import get_logger

logger = get_logger(__name__)


def generate_code_doc(
        title, indirs, exts, comment_chars,
        font_name, font_size, space_before,
        space_after, line_spacing, excludes,
        outfile, template_path=None,
        skip_blank_lines=True, skip_comment_lines=True,
        encoding='utf-8', skip_dir_names=None, skip_file_names=None,
        header_font='宋体', header_font_size=10.5,
        top_margin=2.5, bottom_margin=2.5,
        left_margin=2.5, right_margin=2.5,
        header_distance=1.5, footer_distance=1.75,
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
        header_font: 页眉和页脚字体
        header_font_size: 页眉和页脚字号
        top_margin: 上边距(cm)
        bottom_margin: 下边距(cm)
        left_margin: 左边距(cm)
        right_margin: 右边距(cm)
        header_distance: 页眉距离页面顶端(cm)
        footer_distance: 页脚距离页面底端(cm)
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
    total_lines = total_lines_info['total_lines']  # 过滤后的总行数
    
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
        top_margin=top_margin,
        bottom_margin=bottom_margin,
        left_margin=left_margin,
        right_margin=right_margin,
        header_distance=header_distance,
        footer_distance=footer_distance,
        page_number_format=page_number_format
    )
    writer.write_header(title)
    writer.write_footer()
    
    # 60页模式：使用Word COM精确提取前30页和后30页
    if pages_60_mode:
        # 先生成完整文档到临时文件
        logger.info("60页模式已启用，将先生成完整文档，然后使用Word COM提取前30页和后30页")
        
        # 创建临时文件
        temp_fd, temp_path = tempfile.mkstemp(suffix='.docx', prefix='ccd_full_')
        os.close(temp_fd)
        
        try:
            # 生成完整文档
            logger.info(f"正在生成完整文档到临时文件: {temp_path}")
            for file in files:
                writer.write_file(file)
            writer.save(temp_path)
            
            # 使用Word COM提取60页
            logger.info("正在使用Word COM提取前30页和后30页...")
            from word_com_helper import extract_60_pages, check_word_available
            
            # 检查Word是否可用
            available, error_msg = check_word_available()
            if not available:
                logger.warning(f"Word COM不可用: {error_msg}，将使用完整文档")
                logger.warning("如需使用60页模式，请确保:")
                logger.warning("1. 在Windows系统上运行")
                logger.warning("2. 已安装Microsoft Word")
                logger.warning("3. 已安装pywin32: pip install pywin32")
                
                # 将完整文档保存为输出文件
                import shutil
                shutil.move(temp_path, outfile)
                temp_path = None
                
                return {
                    'file_count': len(files),
                    'outfile': outfile,
                    'total_lines': total_lines_info,
                    'pages_60_mode': False,
                    'warning': f'60页模式失败: {error_msg}'
                }
            
            # 执行提取
            result = extract_60_pages(temp_path, outfile)
            
            if result['success']:
                logger.info(f"成功生成60页文档！原文档{result['total_pages']}页，新文档{result['output_pages']}页")
                return {
                    'file_count': len(files),
                    'outfile': outfile,
                    'total_lines': total_lines_info,
                    'pages_60_mode': True,
                    'original_pages': result['total_pages'],
                    'output_pages': result['output_pages']
                }
            else:
                logger.error(f"60页模式失败: {result['message']}")
                # 回退：保存完整文档
                import shutil
                shutil.move(temp_path, outfile)
                temp_path = None
                
                return {
                    'file_count': len(files),
                    'outfile': outfile,
                    'total_lines': total_lines_info,
                    'pages_60_mode': False,
                    'error': result['message']
                }
                
        finally:
            # 清理临时文件
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                    logger.info(f"已清理临时文件: {temp_path}")
                except:
                    pass
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
