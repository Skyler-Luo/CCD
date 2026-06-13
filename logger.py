# -*- coding: utf-8 -*-
"""
日志管理模块
提供统一的日志记录功能，自动记录错误和运行信息
"""
import logging
import os
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler


class Logger:
    """日志管理器"""
    
    _instance = None
    _initialized = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        
        self._initialized = True
        self.log_dir = self._get_log_dir()
        self.log_file = os.path.join(self.log_dir, 'ccd.log')
        self.error_log_file = os.path.join(self.log_dir, 'ccd_error.log')
        
        # 确保日志目录存在
        os.makedirs(self.log_dir, exist_ok=True)
        
        # 设置根日志记录器
        self.logger = logging.getLogger('CCD')
        self.logger.setLevel(logging.DEBUG)
        
        # 避免重复添加处理器
        if not self.logger.handlers:
            self._setup_handlers()
    
    def _get_log_dir(self):
        """获取日志目录"""
        # 在项目目录下创建日志文件夹
        # 获取当前文件所在目录（即项目根目录）
        project_dir = os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(project_dir, 'logs')
        return log_dir
    
    def _setup_handlers(self):
        """设置日志处理器"""
        # 日志格式
        formatter = logging.Formatter(
            '[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 控制台处理器 (只显示WARNING及以上)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.WARNING)
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)
        
        # 文件处理器 (记录所有级别，最大10MB，保留5个备份)
        try:
            file_handler = RotatingFileHandler(
                self.log_file,
                maxBytes=10*1024*1024,  # 10MB
                backupCount=5,
                encoding='utf-8'
            )
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)
        except Exception as e:
            print(f'无法创建日志文件处理器: {e}')
        
        # 错误日志处理器 (只记录ERROR及以上)
        try:
            error_handler = RotatingFileHandler(
                self.error_log_file,
                maxBytes=5*1024*1024,  # 5MB
                backupCount=3,
                encoding='utf-8'
            )
            error_handler.setLevel(logging.ERROR)
            error_handler.setFormatter(formatter)
            self.logger.addHandler(error_handler)
        except Exception as e:
            print(f'无法创建错误日志文件处理器: {e}')
    
    def get_logger(self, name=None):
        """获取日志记录器"""
        if name:
            return logging.getLogger(f'CCD.{name}')
        return self.logger
    
    def debug(self, message):
        """记录调试信息"""
        self.logger.debug(message)
    
    def info(self, message):
        """记录一般信息"""
        self.logger.info(message)
    
    def warning(self, message):
        """记录警告信息"""
        self.logger.warning(message)
    
    def error(self, message, exc_info=False):
        """记录错误信息"""
        self.logger.error(message, exc_info=exc_info)
    
    def critical(self, message, exc_info=False):
        """记录严重错误信息"""
        self.logger.critical(message, exc_info=exc_info)
    
    def exception(self, message):
        """记录异常信息（自动包含堆栈跟踪）"""
        self.logger.exception(message)
    
    def get_log_file_path(self):
        """获取日志文件路径"""
        return self.log_file
    
    def get_error_log_file_path(self):
        """获取错误日志文件路径"""
        return self.error_log_file
    
    def get_log_dir(self):
        """获取日志目录路径"""
        return self.log_dir
    
    def clear_old_logs(self, days=30):
        """清理指定天数之前的日志文件"""
        try:
            import time
            cutoff_time = time.time() - (days * 24 * 3600)
            
            for filename in os.listdir(self.log_dir):
                filepath = os.path.join(self.log_dir, filename)
                if os.path.isfile(filepath):
                    file_time = os.path.getmtime(filepath)
                    if file_time < cutoff_time:
                        os.remove(filepath)
                        self.info(f'清理旧日志: {filename}')
        except Exception as e:
            self.error(f'清理旧日志失败: {e}')


# 全局日志实例
_logger_instance = Logger()


def get_logger(name=None):
    """获取日志记录器的便捷函数"""
    return _logger_instance.get_logger(name)


def log_operation(operation_name):
    """装饰器：记录操作日志"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            logger = get_logger()
            logger.info(f'开始执行: {operation_name}')
            try:
                result = func(*args, **kwargs)
                logger.info(f'完成执行: {operation_name}')
                return result
            except Exception as e:
                logger.exception(f'执行失败: {operation_name} - {str(e)}')
                raise
        return wrapper
    return decorator
