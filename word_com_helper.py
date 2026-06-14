# -*- coding: utf-8 -*-
"""
Word COM 自动化辅助模块
用于精确控制Word文档的页面操作
仅支持Windows系统
"""
import os
import sys
import tempfile
from logger import get_logger

logger = get_logger(__name__)


def is_windows():
    """检查是否为Windows系统"""
    return sys.platform.startswith('win')


def check_word_available():
    """
    检查Word是否可用
    
    Returns:
        tuple: (是否可用, 错误信息)
    """
    if not is_windows():
        return False, "Word COM自动化仅支持Windows系统"
    
    try:
        import win32com.client
        return True, None
    except ImportError:
        return False, "未安装pywin32，请运行: pip install pywin32"
    except Exception as e:
        return False, f"检查Word时出错: {str(e)}"


def extract_60_pages(input_docx, output_docx):
    """
    从Word文档中提取前30页和后30页，生成新的60页文档
    
    Args:
        input_docx: 输入的完整文档路径
        output_docx: 输出的60页文档路径
        
    Returns:
        dict: 包含操作结果的字典
            - success: 是否成功
            - total_pages: 原文档总页数
            - output_pages: 输出文档页数
            - message: 操作信息
            
    Raises:
        RuntimeError: 当Word不可用时
    """
    available, error_msg = check_word_available()
    if not available:
        raise RuntimeError(error_msg)
    
    import win32com.client
    
    # Word常量定义（手动定义避免依赖问题）
    wdStatisticPages = 2
    wdGoToPage = 1
    wdGoToAbsolute = 1
    wdStory = 6
    wdHeaderFooterPrimary = 1
    wdLineSpaceSingle = 0
    
    word = None
    doc = None
    temp_doc = None
    
    try:
        logger.info(f"启动Word应用程序...")
        # 创建Word应用程序对象
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 不显示Word窗口
        word.DisplayAlerts = False  # 不显示警告
        
        # 打开文档
        logger.info(f"打开文档: {input_docx}")
        input_docx = os.path.abspath(input_docx)
        doc = word.Documents.Open(input_docx)
        
        # 强制重新计算分页
        doc.Repaginate()
        
        # 获取总页数
        total_pages = doc.ComputeStatistics(wdStatisticPages)
        logger.info(f"文档总页数: {total_pages}")
        
        # 如果少于60页，直接复制
        if total_pages <= 60:
            logger.info(f"文档只有{total_pages}页，无需提取，直接复制")
            doc.Close(False)
            doc = None
            
            # 复制文件
            import shutil
            shutil.copy2(input_docx, output_docx)
            
            return {
                'success': True,
                'total_pages': total_pages,
                'output_pages': total_pages,
                'message': f'文档只有{total_pages}页，已直接复制'
            }
        
        # 创建新文档
        logger.info("创建新文档用于存放前30页和后30页...")
        temp_doc = word.Documents.Add()
        
        # 复制页眉页脚设置 - 简单复制，稍后统一处理
        logger.info("复制页眉页脚...")
        for i in range(1, min(doc.Sections.Count, temp_doc.Sections.Count) + 1):
            # 复制页眉
            source_header = doc.Sections(i).Headers(wdHeaderFooterPrimary)
            target_header = temp_doc.Sections(i).Headers(wdHeaderFooterPrimary)
            
            target_header.Range.Delete()
            if source_header.Range.Text.strip():
                source_header.Range.Copy()
                target_header.Range.Paste()
            
            # 复制页脚
            source_footer = doc.Sections(i).Footers(wdHeaderFooterPrimary)
            target_footer = temp_doc.Sections(i).Footers(wdHeaderFooterPrimary)
            
            target_footer.Range.Delete()
            if source_footer.Range.Text.strip():
                source_footer.Range.Copy()
                target_footer.Range.Paste()
        
        # 第一步：复制前30页 - 一次性复制，但使用精确的页面范围
        logger.info("复制前30页...")
        
        if total_pages > 30:
            # 获取第1页起始位置（通常是0）
            doc.Range(0, 0).Select()
            word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=1)
            page1_start = word.Selection.Start
            
            # 获取第31页起始位置
            word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=31)
            page31_start = word.Selection.Start
            
            # 选择第1页到第30页的所有内容（不包括第31页）
            # 使用 page31_start 作为结束位置，让Word自动处理边界
            first_30_range = doc.Range(page1_start, page31_start)
            first_30_range.Select()
            logger.info(f"选择前30页范围: {page1_start} 到 {page31_start}")
            
            # 复制并粘贴
            word.Selection.Copy()
            temp_doc.Range(0, 0).Select()
            word.Selection.Paste()
        else:
            # 如果不超过30页，选择全部
            doc.Range(0, 0).Select()
            word.Selection.EndKey(Unit=wdStory, Extend=1)
            word.Selection.Copy()
            temp_doc.Range(0, 0).Select()
            word.Selection.Paste()
        
        # 检查粘贴后的页数并修正
        temp_doc.Repaginate()
        pages_after_first = temp_doc.ComputeStatistics(wdStatisticPages)
        logger.info(f"复制前30页后，新文档页数: {pages_after_first}")
        
        # 如果超过30页，删除多余的页
        if pages_after_first > 30:
            logger.info(f"检测到多余的 {pages_after_first - 30} 页，正在删除...")
            # 定位到第31页
            temp_doc.Range().Select()
            word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=31)
            page31_in_temp = word.Selection.Start
            
            # 删除从第31页到文档末尾的所有内容
            temp_doc.Range().Select()
            word.Selection.EndKey(Unit=wdStory)
            doc_end_in_temp = word.Selection.End
            
            if page31_in_temp < doc_end_in_temp:
                temp_doc.Range(page31_in_temp, doc_end_in_temp).Delete()
                temp_doc.Repaginate()
                pages_after_first = temp_doc.ComputeStatistics(wdStatisticPages)
                logger.info(f"删除后页数: {pages_after_first}")
        
        # 第二步：复制后30页 - 但要多复制几页以确保最后一页有足够内容
        logger.info("复制后30页...")
        
        # 计算后30页的起始页码，但实际多复制3-5页
        # 这样即使重新分页损失几页，也能保证有足够内容
        extra_pages = 5  # 额外多复制5页作为缓冲
        start_page = total_pages - 29 - extra_pages  # 往前多取5页
        logger.info(f"为确保内容充足，实际复制第 {start_page} 页到第 {total_pages} 页（共 {total_pages - start_page + 1} 页）")
        
        # 获取后30+页起始位置
        doc.Range(0, 0).Select()
        word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=start_page)
        last_pages_start = word.Selection.Start
        
        # 获取文档末尾位置
        doc.Range(0, 0).Select()
        word.Selection.EndKey(Unit=wdStory)
        doc_end = word.Selection.End
        
        # 选择并复制
        doc.Range(last_pages_start, doc_end).Select()
        word.Selection.Copy()
        
        # 粘贴到新文档末尾
        temp_doc.Range().Select()
        word.Selection.EndKey(Unit=wdStory)
        word.Selection.Paste()
        
        # 检查粘贴后的总页数
        temp_doc.Repaginate()
        pages_after_paste = temp_doc.ComputeStatistics(wdStatisticPages)
        logger.info(f"粘贴后新文档总页数: {pages_after_paste}")
        
        # 保存新文档
        logger.info(f"保存新文档: {output_docx}")
        output_docx = os.path.abspath(output_docx)
        temp_doc.SaveAs2(output_docx)
        
        # 先计算保存后的页数
        temp_doc.Repaginate()
        output_pages_before_clean = temp_doc.ComputeStatistics(wdStatisticPages)
        logger.info(f"清理前文档页数: {output_pages_before_clean}")
        
        # 清理页眉页脚中的多余段落
        logger.info("清理页眉页脚中的多余段落...")
        cleaned_count = 0
        for i in range(1, temp_doc.Sections.Count + 1):
            section = temp_doc.Sections(i)
            
            # 清理页眉
            header = section.Headers(wdHeaderFooterPrimary)
            if header.Range.Paragraphs.Count > 1:
                para_count = header.Range.Paragraphs.Count
                for j in range(para_count, 1, -1):  # 从后往前删
                    para = header.Range.Paragraphs(j)
                    if len(para.Range.Text.strip()) == 0:
                        para.Range.Delete()
                        cleaned_count += 1
            
            # 清理页脚
            footer = section.Footers(wdHeaderFooterPrimary)
            if footer.Range.Paragraphs.Count > 1:
                para_count = footer.Range.Paragraphs.Count
                for j in range(para_count, 1, -1):  # 从后往前删
                    para = footer.Range.Paragraphs(j)
                    if len(para.Range.Text.strip()) == 0:
                        para.Range.Delete()
                        cleaned_count += 1
        
        logger.info(f"清理了 {cleaned_count} 个空段落")
        
        # 清理文档末尾的空白页
        logger.info("检查并清理文档末尾空白页...")
        doc_content = temp_doc.Content
        doc_content.Select()
        word.Selection.EndKey(Unit=wdStory)
        
        # 删除末尾的空段落和分页符
        removed_empty = 0
        while True:
            # 移到最后一个字符
            word.Selection.EndKey(Unit=wdStory)
            
            # 向前选择一个字符
            word.Selection.MoveStart(Unit=5, Count=-1)  # 5 = wdCharacter
            
            selected_text = word.Selection.Text
            # 如果是空白字符、段落标记或分页符，删除
            if selected_text in ['\r', '\n', '\r\n', '\f', '\x0c', ' ', '\t'] or len(selected_text.strip()) == 0:
                word.Selection.Delete()
                removed_empty += 1
                if removed_empty > 100:  # 防止无限循环
                    break
            else:
                break
        
        if removed_empty > 0:
            logger.info(f"清理了 {removed_empty} 个末尾空白字符")
        
        # 重新保存清理后的文档
        logger.info("保存清理后的文档...")
        temp_doc.Save()
        
        # 关闭文档
        temp_doc.Close(False)
        temp_doc = None
        doc.Close(False)
        doc = None
        
        # 重新打开文档验证实际页数
        logger.info("重新打开文档验证实际页数...")
        verification_doc = word.Documents.Open(output_docx)
        verification_doc.Repaginate()
        actual_pages = verification_doc.ComputeStatistics(wdStatisticPages)
        logger.info(f"验证后实际页数: {actual_pages}")
        
        # 调整到精确60页
        if actual_pages > 60:
            # 超过60页，从前面删除多余的页（保留最后的代码结束部分）
            pages_to_remove = actual_pages - 60
            logger.info(f"文档有 {actual_pages} 页，需要删除前面多余的 {pages_to_remove} 页")
            
            # 删除第31页开始的多余页面（在前30页之后删除）
            delete_start_page = 31
            delete_end_page = 30 + pages_to_remove
            
            logger.info(f"删除第 {delete_start_page} 页到第 {delete_end_page} 页")
            
            # 定位到要删除的起始页
            verification_doc.Range().Select()
            word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=delete_start_page)
            delete_start_pos = word.Selection.Start
            
            # 定位到要删除的结束页的下一页
            verification_doc.Range().Select()
            word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=delete_end_page + 1)
            delete_end_pos = word.Selection.Start
            
            # 删除范围
            if delete_start_pos < delete_end_pos:
                verification_doc.Range(delete_start_pos, delete_end_pos).Delete()
                verification_doc.Save()
                verification_doc.Repaginate()
                actual_pages = verification_doc.ComputeStatistics(wdStatisticPages)
                logger.info(f"删除后页数: {actual_pages}")
        
        elif actual_pages < 60:
            # 少于60页，在第31页位置插入中间内容
            pages_needed = 60 - actual_pages
            logger.info(f"文档只有 {actual_pages} 页，需要补充 {pages_needed} 页")
            
            # 重新打开原文档
            doc = word.Documents.Open(input_docx)
            
            # 从中间区域选择内容（第31页开始）
            middle_start = 31
            fill_end_page = middle_start + pages_needed - 1
            
            logger.info(f"从原文档第 {middle_start} 页到第 {fill_end_page} 页补充内容")
            
            # 获取补充内容范围
            doc.Range(0, 0).Select()
            word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=middle_start)
            fill_start_pos = word.Selection.Start
            
            doc.Range(0, 0).Select()
            word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=fill_end_page + 1)
            fill_end_pos = word.Selection.Start
            
            # 复制内容
            doc.Range(fill_start_pos, fill_end_pos).Select()
            word.Selection.Copy()
            
            # 插入到新文档第31页位置（前30页之后）
            verification_doc.Range().Select()
            word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=31)
            word.Selection.Paste()
            
            doc.Close(False)
            doc = None
            
            # 保存并验证
            verification_doc.Save()
            verification_doc.Repaginate()
            actual_pages = verification_doc.ComputeStatistics(wdStatisticPages)
            logger.info(f"补充后页数: {actual_pages}")
            
            # 如果补充后超过60页，删除多余部分
            if actual_pages > 60:
                logger.info(f"补充后超过60页，删除第61页及之后的 {actual_pages - 60} 页")
                verification_doc.Range().Select()
                word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=61)
                page61_pos = word.Selection.Start
                
                verification_doc.Range().Select()
                word.Selection.EndKey(Unit=wdStory)
                end_pos = word.Selection.End
                
                if page61_pos < end_pos:
                    verification_doc.Range(page61_pos, end_pos).Delete()
                    verification_doc.Save()
                    verification_doc.Repaginate()
                    actual_pages = verification_doc.ComputeStatistics(wdStatisticPages)
                    logger.info(f"最终页数: {actual_pages}")
        
        verification_doc.Close(False)
        verification_doc = None
        
        output_pages = actual_pages
        
        logger.info(f"成功生成60页文档，实际页数: {output_pages}")
        
        if output_pages != 60:
            logger.warning(f"警告：生成的文档有 {output_pages} 页，不是预期的60页（差异：{60 - output_pages}页）")
        
        return {
            'success': True,
            'total_pages': total_pages,
            'output_pages': output_pages,
            'message': f'成功从{total_pages}页文档中提取前30页和后30页'
        }
        
    except Exception as e:
        logger.error(f"Word COM操作失败: {str(e)}", exc_info=True)
        return {
            'success': False,
            'total_pages': 0,
            'output_pages': 0,
            'message': f'操作失败: {str(e)}'
        }
        
    finally:
        # 清理资源
        try:
            if temp_doc:
                temp_doc.Close(False)
        except:
            pass
        
        try:
            if doc:
                doc.Close(False)
        except:
            pass
        
        try:
            if word:
                word.Quit()
        except:
            pass


def test_word_com():
    """
    测试Word COM是否正常工作
    
    Returns:
        tuple: (是否成功, 消息)
    """
    available, error_msg = check_word_available()
    if not available:
        return False, error_msg
    
    import win32com.client
    
    try:
        # 尝试创建Word应用程序
        word = win32com.client.Dispatch("Word.Application")
        version = word.Version
        word.Quit()
        return True, f"Word {version} 可用"
    except Exception as e:
        return False, f"无法启动Word: {str(e)}"


def diagnose_document_pages(docx_path):
    """
    诊断文档页数问题
    
    Args:
        docx_path: 文档路径
        
    Returns:
        dict: 诊断信息
    """
    available, error_msg = check_word_available()
    if not available:
        return {'error': error_msg}
    
    import win32com.client
    
    # Word常量
    wdStatisticPages = 2
    wdStatisticParagraphs = 4
    wdStatisticWords = 0
    wdHeaderFooterPrimary = 1
    
    word = None
    doc = None
    
    try:
        logger.info("启动Word进行诊断...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        logger.info(f"打开文档: {docx_path}")
        doc = word.Documents.Open(os.path.abspath(docx_path))
        
        # 强制重新分页
        doc.Repaginate()
        
        # 收集统计信息
        pages = doc.ComputeStatistics(wdStatisticPages)
        paragraphs = doc.ComputeStatistics(wdStatisticParagraphs)
        words = doc.ComputeStatistics(wdStatisticWords)
        
        # 检查节信息
        sections_count = doc.Sections.Count
        
        # 检查每一节的页眉页脚
        sections_info = []
        for i in range(1, sections_count + 1):
            section = doc.Sections(i)
            header = section.Headers(wdHeaderFooterPrimary)
            footer = section.Footers(wdHeaderFooterPrimary)
            
            sections_info.append({
                'section': i,
                'header_paragraphs': header.Range.Paragraphs.Count,
                'footer_paragraphs': footer.Range.Paragraphs.Count,
                'header_text_length': len(header.Range.Text),
                'footer_text_length': len(footer.Range.Text)
            })
        
        # 检查文档末尾内容
        doc_content = doc.Content
        last_100_chars = doc_content.Text[-100:] if len(doc_content.Text) > 100 else doc_content.Text
        
        result = {
            'pages': pages,
            'paragraphs': paragraphs,
            'words': words,
            'sections': sections_count,
            'sections_info': sections_info,
            'document_length': len(doc_content.Text),
            'last_100_chars': repr(last_100_chars),
            'has_trailing_breaks': '\f' in last_100_chars or '\x0c' in last_100_chars
        }
        
        doc.Close(False)
        doc = None
        
        return result
        
    except Exception as e:
        logger.error(f"诊断失败: {str(e)}", exc_info=True)
        return {'error': str(e)}
        
    finally:
        try:
            if doc:
                doc.Close(False)
        except:
            pass
        
        try:
            if word:
                word.Quit()
        except:
            pass

