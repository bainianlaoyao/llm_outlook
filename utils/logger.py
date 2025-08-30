import logging
import logging.handlers
from pathlib import Path
from typing import Optional

from config.defaults import LOG_CONFIG


class Logger:
    """简单的日志记录器"""

    def __init__(self, name: str = 'auto_mail', level: Optional[str] = None,
                 format_str: Optional[str] = None, file_path: Optional[str] = None):
        self.name = name
        self.level = level or LOG_CONFIG['level']
        self.format_str = format_str or LOG_CONFIG['format']
        self.file_path = file_path or LOG_CONFIG['file_path']

        # 确保日志目录存在
        log_dir = Path(self.file_path).parent
        log_dir.mkdir(parents=True, exist_ok=True)

        # 创建logger
        self.logger = logging.getLogger(name)
        self.logger.setLevel(getattr(logging, self.level.upper()))

        # 创建格式器
        formatter = logging.Formatter(self.format_str)

        # 创建控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)

        # 创建文件处理器
        file_handler = logging.handlers.RotatingFileHandler(
            self.file_path,
            maxBytes=LOG_CONFIG['max_file_size'],
            backupCount=LOG_CONFIG['backup_count'],
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)

    def debug(self, message: str):
        """记录debug级别日志"""
        self.logger.debug(message)

    def info(self, message: str):
        """记录info级别日志"""
        self.logger.info(message)

    def warning(self, message: str):
        """记录warning级别日志"""
        self.logger.warning(message)

    def error(self, message: str):
        """记录error级别日志"""
        self.logger.error(message)

    def critical(self, message: str):
        """记录critical级别日志"""
        self.logger.critical(message)

    def log(self, level: str, message: str):
        """记录指定级别的日志"""
        level_method = getattr(self.logger, level.lower())
        level_method(message)

    def set_level(self, level: str):
        """设置日志级别"""
        self.logger.setLevel(getattr(logging, level.upper()))
        self.level = level


# 全局logger实例
_default_logger = None

def get_logger(name: str = 'auto_mail') -> Logger:
    """获取logger实例"""
    global _default_logger
    if _default_logger is None:
        _default_logger = Logger(name)
    return _default_logger

def log_email_fetch(email: str, count: int):
    """记录邮件拉取日志"""
    logger = get_logger()
    logger.info(f"从邮箱 {email} 拉取了 {count} 封邮件")

def log_email_parse(email_id: str, success: bool):
    """记录邮件解析日志"""
    logger = get_logger()
    status = "成功" if success else "失败"
    logger.info(f"邮件 {email_id} 解析{status}")

def log_push_message(channel: str, success: bool):
    """记录消息推送日志"""
    logger = get_logger()
    status = "成功" if success else "失败"
    logger.info(f"消息推送至 {channel} {status}")

def log_error(component: str, error: Exception):
    """记录错误日志"""
    logger = get_logger()
    logger.error(f"{component} 出现错误: {str(error)}")

def log_info(message: str):
    """记录信息日志"""
    logger = get_logger()
    logger.info(message)

def log_warning(message: str):
    """记录警告日志"""
    logger = get_logger()
    logger.warning(message)