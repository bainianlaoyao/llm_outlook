"""
纯Outlook本地邮件处理系统主控制器
零网络依赖，企业级安全，Linus风格架构

作者：Linus Torvalds风格架构师
"""

import json
import os
from typing import Optional, Dict, Any, List

from core.outlook_email_fetcher import OutlookEmailFetcher
from core.email_parser import EmailParser
from core.message_pusher import MultiChannelPusher
from models.email import EmailData
from utils.logger import get_logger, log_email_fetch, log_push_message, log_info


class PureOutlookController:
    """
    纯Outlook邮件处理控制器
    简洁架构：拉取→解析→推送
    """

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or {}
        self.logger = get_logger('pure_outlook_controller')

        # 状态管理
        self.running = False

        # 初始化核心组件
        self._setup_components()

        self.logger.info("✓ 纯Outlook控制器初始化完成")

    def _setup_components(self):
        """设置核心组件"""
        # Outlook获取器
        outlook_profile = self.config.get('outlook', {}).get('profile', 'Outlook')
        self.fetcher = OutlookEmailFetcher(profile=outlook_profile,email_keyword="bnly")

        # 解析器
        self.parser = EmailParser(
            api_key=self.config.get('parser', {}).get('api_key'),
            logger=self.logger
        )

        # 推送器
        self.pusher = MultiChannelPusher(
            channels=self.config.get('push', {}).get('channels'),
            serverchan_sendkey=self.config.get('push', {}).get('sendkey')
        )

    

    def start(self) -> bool:
        """启动系统 - 一次性执行模式"""
        if not self.fetcher.connect():
            self.logger.error("✗ Outlook连接失败")
            return False

        self.logger.info("🚀 执行一次性邮件处理")
        log_info("执行一次性邮件处理")

        try:
            # 执行一次处理循环
            self._process_cycle()
            self.logger.info("✅ 邮件处理完成")
            log_info("邮件处理已完成")
            return True
        except Exception as e:
            self.logger.error(f"邮件处理异常: {e}")
            return False
        finally:
            self._shutdown()

    def stop(self):
        """停止系统"""
        self.logger.info("正在停止系统...")
        self.running = False

    def _process_cycle(self):
        """处理循环：拉取→批量解析→推送"""
        try:
            # 1. 拉取新邮件 - 取消数量限制，只按时间过滤
            emails, _ = self.fetcher.fetch_emails(7)  # 过去7天的所有邮件

            if not emails:
                self.logger.debug("没有新邮件")
                return

            log_email_fetch("Outlook", len(emails))
            self.logger.info(f"📧 获取 {len(emails)} 封新邮件")

            # 2. 批量处理所有邮件
            if self._handle_emails_batch(emails):
                self.logger.info(f"✅ 批量处理完成 {len(emails)} 封邮件")
            else:
                self.logger.warning(f"❌ 批量处理失败")

        except Exception as e:
            self.logger.error(f"处理循环异常: {e}")

    def _handle_emails_batch(self, emails: List[EmailData]) -> bool:
        """批量处理邮件：批量解析→推送"""

        # 1. 批量解析邮件
        batch_result = self.parser.parse_emails_batch(emails)
        # 2. 推送批量结果
        push_result = self.pusher.push(batch_result.split('\n')[0], batch_result)

        if push_result.success:
            log_push_message("批量推送", True)
            self.logger.info(f"✅ 批量推送成功: {batch_result.split('\n')[0]}")
            return True
        else:
            log_push_message("批量推送", False)
            self.logger.error(f"❌ 批量推送失败: {push_result.message}")
            return False

    def _shutdown(self):
        """系统关闭处理"""
        self.logger.info("系统正在关闭...")

        # 断开连接
        if self.fetcher:
            self.fetcher.disconnect()

        log_info("纯Outlook邮件处理系统已停止")
        self.logger.info("✓ 系统已安全关闭")


def create_outlook_controller_from_config(config_file: Optional[str] = None) -> PureOutlookController:
    """
    从配置文件创建控制器

    参数:
        config_file: 配置文件路径

    返回:
        配置好的控制器实例
    """
    config = {}

    if config_file and os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            print(f"✅ 已加载配置文件: {config_file}")
        except Exception as e:
            print(f"❌ 加载配置文件失败: {e}")
            print("使用默认配置继续...")

    return PureOutlookController(config)