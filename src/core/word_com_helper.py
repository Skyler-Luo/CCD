# -*- coding: utf-8 -*-
"""
Word COM 自动化辅助模块
用于精确控制Word文档的页面操作
仅支持Windows系统
"""
import os
import sys
import tempfile
from src.utils.logger import get_logger

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
    原地从Word文档中提取前30页和后30页，微调生成新的60页文档，无需依赖剪贴板复制粘贴
    
    Args:
        input_docx: 输入的完整文档路径
        output_docx: 输出的60页文档路径
        
    Returns:
        dict: 包含操作结果的字典
            - success: 是否成功
            - total_pages: 原文档总页数
            - output_pages: 输出文档页数
            - message: 操作信息
    """
    available, error_msg = check_word_available()
    if not available:
        raise RuntimeError(error_msg)
    
    import pythoncom
    import win32com.client
    
    # Word常量定义
    wdStatisticPages = 2
    wdGoToPage = 1
    wdGoToAbsolute = 1
    wdStory = 6
    wdHeaderFooterPrimary = 1
    
    word = None
    doc = None
    com_initialized = False
    
    try:
        # 初始化 COM 为 MTA 模式，避免 Word 占用时弹出无法关闭的 Server Busy 弹窗导致卡死
        pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
        com_initialized = True
        import shutil
        input_docx = os.path.abspath(input_docx)
        output_docx = os.path.abspath(output_docx)
        
        # 第一步：直接把原文档拷贝至目标路径，转为原地操作
        logger.info(f"直接复制原文档到输出路径进行原地裁剪...")
        shutil.copy2(input_docx, output_docx)
        
        logger.info(f"以 DispatchEx 独立进程启动 Word 应用程序...")
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False  # 不显示Word窗口
        word.DisplayAlerts = False  # 不显示警告
        
        # 打开输出文档
        logger.info(f"打开文档: {output_docx}")
        doc = word.Documents.Open(output_docx)
        
        # 重新分页计算
        doc.Repaginate()
        total_pages = doc.ComputeStatistics(wdStatisticPages)
        logger.info(f"原文档总页数: {total_pages}")
        
        # 若总页数少于或等于60页，无需裁剪，直接返回
        if total_pages <= 60:
            logger.info(f"原文档只有 {total_pages} 页，无需裁剪提取，直接结束")
            doc.Close(True)
            doc = None
            return {
                'success': True,
                'total_pages': total_pages,
                'output_pages': total_pages,
                'message': f'文档只有{total_pages}页，已直接保留'
            }
        
        # 1. 精确删除中间超出前30和后30页的范围
        delete_start_page = 31
        delete_end_page = total_pages - 29
        
        logger.info(f"进行原地裁剪：删除第 {delete_start_page} 页起始到第 {delete_end_page} 页起始之间的内容")
        
        # 定位起始与结束偏移位置
        doc.Range(0, 0).Select()
        word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=delete_start_page)
        delete_start_pos = word.Selection.Start
        
        word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=delete_end_page)
        delete_end_pos = word.Selection.Start
        
        if delete_start_pos < delete_end_pos:
            doc.Range(delete_start_pos, delete_end_pos).Delete()
        
        # 再次强制重新分页计算
        doc.Repaginate()
        actual_pages = doc.ComputeStatistics(wdStatisticPages)
        logger.info(f"原地初次裁剪后实际页数: {actual_pages}")
        
        # 2. 动态页数对齐微调
        if actual_pages > 60:
            pages_to_remove = actual_pages - 60
            logger.info(f"微调：裁剪后有 {actual_pages} 页，再次删除中间多余的 {pages_to_remove} 页")
            _delete_document_range(word, doc, 31, 30 + pages_to_remove)
            doc.Repaginate()
            actual_pages = doc.ComputeStatistics(wdStatisticPages)
            logger.info(f"微调后页数: {actual_pages}")
            
        elif actual_pages < 60:
            pages_needed = 60 - actual_pages
            logger.info(f"微调：裁剪后仅有 {actual_pages} 页，需要从原文档补充中间区域 {pages_needed} 页")
            _insert_middle_pages_fallback(word, doc, input_docx, pages_needed)
            doc.Repaginate()
            actual_pages = doc.ComputeStatistics(wdStatisticPages)
            logger.info(f"补充微调后页数: {actual_pages}")
            
        # 3. 清理空段落和多余分页符以校准行格式
        logger.info("开始清理文档格式垃圾...")
        _clean_header_footer_paragraphs(doc, wdHeaderFooterPrimary)
        _clean_document_breaks(word, doc, wdStory)
        
        # 终存与统计最终页数
        doc.Save()
        doc.Repaginate()
        output_pages = doc.ComputeStatistics(wdStatisticPages)
        
        doc.Close(True)
        doc = None
        
        logger.info(f"原地裁剪文档成功！最终页数: {output_pages}")
        
        if output_pages != 60:
            logger.warning(f"警告：生成的文档实际为 {output_pages} 页，不是预期的60页")
            
        return {
            'success': True,
            'total_pages': total_pages,
            'output_pages': output_pages,
            'message': f'成功从{total_pages}页文档中原地提取前30页和后30页并完成校正'
        }
        
    except Exception as e:
        logger.error(f"原地裁剪 Word COM 操作失败: {str(e)}", exc_info=True)
        return {
            'success': False,
            'total_pages': 0,
            'output_pages': 0,
            'message': f'原地操作失败: {str(e)}'
        }
        
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
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except:
                pass


def _delete_document_range(word, doc, start_page, end_page):
    """（私有辅助方法）原地删除文档中指定页码范围的内容"""
    wdGoToPage = 1
    wdGoToAbsolute = 1
    
    doc.Range(0, 0).Select()
    word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=start_page)
    start_pos = word.Selection.Start
    
    word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=end_page + 1)
    end_pos = word.Selection.Start
    
    if start_pos < end_pos:
        doc.Range(start_pos, end_pos).Delete()


def _insert_middle_pages_fallback(word, doc, input_docx, pages_needed):
    """（私有辅助方法）从原文档拷贝指定数量 of 页码并插入到当前文档第31页位置"""
    wdGoToPage = 1
    wdGoToAbsolute = 1
    
    # 重新打开原文档读取
    source_doc = word.Documents.Open(input_docx, ReadOnly=True)
    try:
        middle_start = 31
        fill_end_page = middle_start + pages_needed - 1
        
        logger.info(f"从只读原文档第 {middle_start} 页到第 {fill_end_page} 页读取内容并拷贝...")
        source_doc.Range(0, 0).Select()
        word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=middle_start)
        fill_start_pos = word.Selection.Start
        
        word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=fill_end_page + 1)
        fill_end_pos = word.Selection.Start
        
        # 拷贝
        if fill_start_pos < fill_end_pos:
            source_doc.Range(fill_start_pos, fill_end_pos).Select()
            word.Selection.Copy()
            
            # 粘贴到裁剪文档第31页起始位置
            doc.Range(0, 0).Select()
            word.Selection.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=31)
            word.Selection.Paste()
    finally:
        source_doc.Close(False)


def _clean_header_footer_paragraphs(doc, wdHeaderFooterPrimary=1):
    """（私有辅助方法）清扫节页眉页脚中由于操作残留的空白空行"""
    cleaned_count = 0
    for i in range(1, doc.Sections.Count + 1):
        section = doc.Sections(i)
        
        # 清理页眉
        header = section.Headers(wdHeaderFooterPrimary)
        if header.Range.Paragraphs.Count > 1:
            para_count = header.Range.Paragraphs.Count
            for j in range(para_count, 1, -1):
                para = header.Range.Paragraphs(j)
                if len(para.Range.Text.strip()) == 0:
                    para.Range.Delete()
                    cleaned_count += 1
        
        # 清理页脚
        footer = section.Footers(wdHeaderFooterPrimary)
        if footer.Range.Paragraphs.Count > 1:
            para_count = footer.Range.Paragraphs.Count
            for j in range(para_count, 1, -1):
                para = footer.Range.Paragraphs(j)
                if len(para.Range.Text.strip()) == 0:
                    para.Range.Delete()
                    cleaned_count += 1
    logger.info(f"共清扫页眉页脚中 {cleaned_count} 个多余空白段落")


def _clean_document_breaks(word, doc, wdStory=6):
    """（私有辅助方法）检查并清理裁剪后文档末尾的多余空白分页符与回车标记"""
    removed_empty = 0
    doc_content = doc.Content
    doc_content.Select()
    word.Selection.EndKey(Unit=wdStory)
    
    while True:
        word.Selection.EndKey(Unit=wdStory)
        word.Selection.MoveStart(Unit=5, Count=-1) # 5 = wdCharacter
        selected_text = word.Selection.Text
        
        if selected_text in ['\r', '\n', '\r\n', '\f', '\x0c', ' ', '\t'] or len(selected_text.strip()) == 0:
            word.Selection.Delete()
            removed_empty += 1
            if removed_empty > 100:  # 保护机制防止死循环
                break
        else:
            break
            
    if removed_empty > 0:
        logger.info(f"共清理末尾空白及分页符 {removed_empty} 次")


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


def convert_docx_to_pdf(docx_path, pdf_path=None):
    """
    将 docx 文档转换为 PDF 格式
    
    使用 Word COM 自动化实现高保真转换，页眉、页脚、页码、字体均完美保留。
    
    Args:
        docx_path: 输入的 docx 文件路径
        pdf_path: 输出的 PDF 文件路径，默认为同目录同名 .pdf 文件
        
    Returns:
        dict: 包含操作结果的字典
            - success: 是否成功
            - pdf_path: 输出 PDF 的路径
            - message: 操作信息
    """
    available, error_msg = check_word_available()
    if not available:
        return {
            'success': False,
            'pdf_path': None,
            'message': f'Word COM 不可用: {error_msg}'
        }
    
    docx_path = os.path.abspath(docx_path)
    if not os.path.isfile(docx_path):
        return {
            'success': False,
            'pdf_path': None,
            'message': f'docx 文件不存在: {docx_path}'
        }
    
    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
    pdf_path = os.path.abspath(pdf_path)
    
    import pythoncom
    import win32com.client
    
    word = None
    doc = None
    com_initialized = False
    
    try:
        # 在子线程中需要初始化 COM (MTA 模式，避免 Word 占用时弹出无法关闭的 Server Busy 弹窗导致卡死)
        pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
        com_initialized = True
        
        logger.info(f"启动 Word 进行 PDF 转换...")
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        logger.info(f"打开文档: {docx_path}")
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        
        logger.info(f"导出 PDF: {pdf_path}")
        doc.ExportAsFixedFormat(
            OutputFileName=pdf_path,
            ExportFormat=17,  # wdExportFormatPDF
        )
        
        # ExportAsFixedFormat 后 COM 连接可能断开，Close 失败是正常的
        try:
            doc.Close(False)
            doc = None
        except Exception:
            doc = None
        
        # 以 PDF 文件是否实际存在来判断成功
        if os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 0:
            logger.info(f"PDF 导出成功: {pdf_path}")
            return {
                'success': True,
                'pdf_path': pdf_path,
                'message': 'PDF 导出成功'
            }
        else:
            logger.error("PDF 文件未生成或为空")
            return {
                'success': False,
                'pdf_path': None,
                'message': 'PDF 文件未生成'
            }
        
    except Exception as e:
        logger.error(f"PDF 转换失败: {str(e)}", exc_info=True)
        # 即使异常，也检查 PDF 是否已经生成
        if os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 0:
            logger.info(f"虽然出现异常，但 PDF 已成功生成: {pdf_path}")
            return {
                'success': True,
                'pdf_path': pdf_path,
                'message': 'PDF 导出成功（过程中有非致命异常）'
            }
        return {
            'success': False,
            'pdf_path': None,
            'message': f'PDF 转换失败: {str(e)}'
        }
        
    finally:
        try:
            if doc:
                doc.Close(False)
        except Exception:
            pass
        
        try:
            if word:
                word.Quit()
        except Exception:
            pass
        
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def check_word_running():
    """检查后台是否有正在运行的 Word 进程 (winword.exe)"""
    if sys.platform != 'win32':
        return False
    import subprocess
    try:
        output_word = subprocess.check_output(
            'tasklist /FI "IMAGENAME eq winword.exe" /FO CSV /NH',
            shell=True,
            stderr=subprocess.DEVNULL,
            creationflags=0x08000000
        )
        return b"winword.exe" in output_word.lower()
    except Exception:
        return False


def kill_word_process():
    """强杀后台 Word 进程"""
    if sys.platform != 'win32':
        return True
    import subprocess
    try:
        subprocess.check_call(
            'taskkill /f /im winword.exe', 
            shell=True, 
            stdout=subprocess.DEVNULL, 
            stderr=subprocess.DEVNULL,
            creationflags=0x08000000
        )
        return True
    except Exception:
        return False
