# -*- coding: utf-8 -*-
"""
主题样式辅助工具
提供调用 Windows DWM API 动态改变窗口标题栏颜色的功能
"""
import sys
import ctypes

def set_window_dark_title_bar(window, is_dark: bool, custom_bg_color=None):
    """
    通过 Windows DWM API 动态设置窗口标题栏为深色或浅色模式
    并且统一标题栏颜色与背景颜色一致，实现无边框沉浸感
    
    custom_bg_color: 可选，指定的标题栏背景色 (格式为 0x00BBGGRR 的 COLORREF 整数值)
    """
    if sys.platform != 'win32':
        return
        
    try:
        # 获取窗口 HWND 句柄
        hwnd = int(window.winId())
        dwmapi = ctypes.windll.dwmapi
        
        # 1. 基础暗黑/浅色模式切换
        # DWMWA_USE_IMMERSIVE_DARK_MODE 属性在 Windows 中的定义值：
        # Windows 11 及更新版本中使用 20
        # Windows 10 (Build 17763 至 19044) 中使用 19
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19
        
        # 将布尔值转换为 ctypes 的 32 位整型 (1 表示开启暗黑标题栏，0 表示浅色标题栏)
        value = ctypes.c_int(1 if is_dark else 0)
        
        # 首先尝试应用 Windows 11 的属性
        res = dwmapi.DwmSetWindowAttribute(
            hwnd, 
            DWMWA_USE_IMMERSIVE_DARK_MODE, 
            ctypes.byref(value), 
            ctypes.sizeof(value)
        )
        # 如果失败，则尝试 Windows 10 的属性
        if res != 0:
            dwmapi.DwmSetWindowAttribute(
                hwnd, 
                DWMWA_USE_IMMERSIVE_DARK_MODE_OLD, 
                ctypes.byref(value), 
                ctypes.sizeof(value)
            )
            
        # 2. 统一标题栏背景色 (Windows 11 Build 22000+)
        # DWMWA_CAPTION_COLOR = 35, 颜色格式为 0x00BBGGRR
        # 32 = 0x20, 243 = 0xF3
        DWMWA_CAPTION_COLOR = 35
        if custom_bg_color is not None:
            caption_color = custom_bg_color
        else:
            caption_color = 0x00202020 if is_dark else 0x00F3F3F3
            
        color_val = ctypes.c_int(caption_color)
        dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_CAPTION_COLOR,
            ctypes.byref(color_val),
            ctypes.sizeof(color_val)
        )
        
        # 3. 统一标题栏文字颜色 (Windows 11 Build 22000+)
        # DWMWA_TEXT_COLOR = 36, 颜色格式为 0x00BBGGRR
        # 深色模式下文字为白色 (0x00FFFFFF)，浅色模式下为黑色 (0x00000000)
        DWMWA_TEXT_COLOR = 36
        text_color = 0x00FFFFFF if is_dark else 0x00000000
        text_val = ctypes.c_int(text_color)
        dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_TEXT_COLOR,
            ctypes.byref(text_val),
            ctypes.sizeof(text_val)
        )
            
        # 刷新窗口的非客户区（标题栏）
        window.update()
    except Exception as e:
        # 如果调用失败则静默失败，不影响主程序运行
        pass
