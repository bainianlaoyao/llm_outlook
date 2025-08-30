#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
纯Outlook本地邮件获取器
零网络依赖，完全本地化解决方案

核心优势：
- 零配置复杂度：无需IMAP服务器、端口、密码配置
- 企业级安全性：使用Windows认证，无需存储敏感凭证
- 最高性能：直接COM接口访问，本地数据获取速度极快
- 实现简单：单个获取器，消除复杂的分支逻辑
"""

from typing import List, Tuple

from models.email import EmailData
from core.outlook_email_fetcher import OutlookEmailFetcher
from utils.logger import get_logger

logger = get_logger(__name__)


def fetch_emails_auto(last_uid: str = '', max_emails: int = 50, profile: str = "Outlook") -> Tuple[List[EmailData], str]:
    """
    纯Outlook邮件获取函数
    直接通过本地COM接口获取mail，零网络依赖

    参数:
        last_uid: 上次处理的EntryID（空字符串表示初次运行）
        max_emails: 最大拉取数量
        profile: Outlook配置文件名称

    返回:
        (邮件列表, 最新EntryID)
    """
    return fetch_outlook_emails(last_uid, max_emails, profile)


def fetch_outlook_emails(last_uid: str, max_emails: int, profile: str = "Outlook") -> Tuple[List[EmailData], str]:
    """
    获取Outlook邮件
    直接使用核心Outlook获取器
    """
    fetcher = OutlookEmailFetcher(profile)
    emails, latest_uid = fetcher.fetch_emails(max_emails)
    fetcher.disconnect()
    return emails, latest_uid


class OutlookAutoFetcher:
    """纯Outlook自动邮件获取器类"""

    def __init__(self, profile: str = "Outlook"):
        """
        初始化Outlook获取器

        参数:
            profile: Outlook配置文件名称
        """
        self.profile = profile
        self.fetcher = OutlookEmailFetcher(profile)

    def connect(self) -> bool:
        """
        连接到本地Outlook

        返回:
            bool: 连接是否成功
        """
        return self.fetcher.connect()

    def fetch_emails(self, last_uid: str = '', max_emails: int = 50) -> Tuple[List[EmailData], str]:
        """
        拉取邮件

        参数:
            last_uid: 上次处理的EntryID
            max_emails: 最大拉取数量

        返回:
            (邮件列表, 最新EntryID)
        """
        return self.fetcher.fetch_emails(max_emails)

    def disconnect(self):
        """断开连接"""
        self.fetcher.disconnect()


def test_outlook_fetcher():
    """测试纯Outlook邮件获取器"""
    print("=== 纯Outlook邮件获取器测试 ===")
    print("使用本地COM接口，无需网络连接")
    print()

    fetcher = OutlookAutoFetcher()

    if fetcher.connect():
        print("✅ Outlook连接成功")
        print("准备获取邮件...")

        emails, latest_entry_id = fetcher.fetch_emails(max_emails=3)
        print(f"成功获取 {len(emails)} 封邮件")

        if emails:
            print("\n最新邮件:")
            for i, email in enumerate(emails[:2], 1):
                print(f"{i}. {email.subject[:50]}...")

    else:
        print("❌ Outlook连接失败")
        print("请确保Outlook已正确安装并且有邮件数据")

    fetcher.disconnect()
    print("\n测试完成")


# 向后兼容的便捷函数（简化参数）
def fetch_emails(last_uid: str = '', max_emails: int = 50, profile: str = "Outlook") -> Tuple[List[EmailData], str]:
    """
    向后兼容的邮件获取函数
    直接使用纯Outlook方案
    """
    return fetch_emails_auto(last_uid, max_emails, profile)


if __name__ == "__main__":
    test_outlook_fetcher()