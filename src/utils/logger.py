# -*- coding: utf-8 -*-
"""
日志管理模块
提供统一的日志记录功能，自动记录错误和运行信息
"""
import logging
import os
import sys
import time
from logging.handlers import RotatingFileHandler


class Logger:
    """日志管理器（单例模式）"""
    
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialize()
        return cls._instance
    
    def _initialize(self):
        """初始化日志系统"""
        utils_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.dirname(utils_dir)
        project_dir = os.path.dirname(src_dir)
        self.log_dir = os.path.join(project_dir, 'logs')
        os.makedirs(self.log_dir, exist_ok=True)
        
        self.log_file = os.path.join(self.log_dir, 'ccd.log')
        self.error_log_file = os.path.join(self.log_dir, 'ccd_error.log')
        
        self.logger = logging.getLogger('CCD')
        self.logger.setLevel(logging.DEBUG)
        
        if not self.logger.handlers:
            self._setup_handlers()
    
    def _setup_handlers(self):
        """设置日志处理器"""
        formatter = logging.Formatter(
            '[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 控制台处理器 (只显示WARNING及以上)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.WARNING)
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)
        
        # 添加文件处理器
        self._add_file_handler(self.log_file, logging.DEBUG, 10*1024*1024, 5, formatter)
        self._add_file_handler(self.error_log_file, logging.ERROR, 5*1024*1024, 3, formatter)
    
    def _add_file_handler(self, log_file, level, max_bytes, backup_count, formatter):
        """添加文件处理器"""
        try:
            handler = RotatingFileHandler(
                log_file, maxBytes=max_bytes, backupCount=backup_count, encoding='utf-8'
            )
            handler.setLevel(level)
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
        except Exception as e:
            print(f'无法创建日志文件处理器 {log_file}: {e}')
    
    def get_logger(self, name=None):
        """获取日志记录器"""
        return logging.getLogger(f'CCD.{name}') if name else self.logger
    
    def clear_old_logs(self, days=30):
        """清理指定天数之前的日志文件"""
        try:
            cutoff_time = time.time() - (days * 24 * 3600)
            for filename in os.listdir(self.log_dir):
                filepath = os.path.join(self.log_dir, filename)
                if os.path.isfile(filepath) and os.path.getmtime(filepath) < cutoff_time:
                    os.remove(filepath)
                    self.logger.info(f'清理旧日志: {filename}')
        except Exception as e:
            self.logger.error(f'清理旧日志失败: {e}')


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
