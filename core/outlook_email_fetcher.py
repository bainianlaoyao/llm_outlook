#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
纯Outlook本地邮件获取器
零网络依赖，直接通过COM接口访问本地Outlook邮件
"""

import win32com.client
from typing import List, Optional, Tuple
import pythoncom
from datetime import datetime, timedelta

from models.email import EmailData
from utils.logger import get_logger

logger = get_logger(__name__)


class OutlookEmailFetcher:
    """纯Outlook本地邮件获取器"""

    def __init__(self, profile: str = "Outlook", email_keyword: Optional[str] = None):
        self.profile = profile
        self.email_keyword = email_keyword
        self.outlook_app = None
        self.namespace = None
        self.inbox = None
        self.processed_ids = set()
        logger.info(f"初始化Outlook获取器: profile={profile}, email_keyword={email_keyword}")

    def connect(self) -> bool:
        """连接到本地Outlook"""
        try:
            pythoncom.CoInitialize()
            self.outlook_app = win32com.client.Dispatch('Outlook.Application')
            self.namespace = self.outlook_app.GetNamespace('MAPI')

            if self.email_keyword:
                account = self._find_account_by_keyword(self.email_keyword)
                if account:
                    self.inbox = account.DeliveryStore.GetDefaultFolder(6) # olFolderInbox
                    logger.info(f"✓ 使用账户: {account.DisplayName}")
                else:
                    logger.warning(f"未找到账户 '{self.email_keyword}'，使用默认收件箱")
                    self.inbox = self.namespace.GetDefaultFolder(6)
            else:
                self.inbox = self.namespace.GetDefaultFolder(6)

            logger.info("✓ Outlook连接成功")
            return True
        except Exception as e:
            logger.error(f"✗ Outlook连接失败: {e}")
            self._cleanup()
            return False

    def fetch_emails(self, last_days: int = 7) -> Tuple[List[EmailData], str]:
        """获取指定天数内的新邮件"""
        if not self._ensure_connection():
            return [], ""

        try:
            items = self.inbox.Items
            items.Sort("[ReceivedTime]", True)

            email_list = []
            latest_entry_id = ""
            cutoff_time = datetime.now() - timedelta(days=last_days)

            for item in items:
                try:
                    received_time = item.ReceivedTime.replace(tzinfo=None)
                    if received_time < cutoff_time:
                        break

                    entry_id = item.EntryID
                    if entry_id in self.processed_ids:
                        continue

                    email_data = self._convert_to_email_data(item, entry_id)
                    if email_data:
                        email_list.append(email_data)
                        if not latest_entry_id:
                            latest_entry_id = entry_id
                        self.processed_ids.add(entry_id)
                except Exception:
                    continue

            logger.info(f"获取 {len(email_list)} 封过去 {last_days} 天的新邮件")
            return email_list, latest_entry_id
        except Exception as e:
            logger.error(f"拉取邮件异常: {e}")
            return [], ""

    def disconnect(self):
        """断开连接"""
        self._cleanup()
        logger.info("Outlook连接已断开")

    def _ensure_connection(self) -> bool:
        if not self.outlook_app:
            return self.connect()
        return True

    def _convert_to_email_data(self, mail_item, entry_id: str) -> Optional[EmailData]:
        try:
            return EmailData(
                uid=entry_id,
                subject=mail_item.Subject or "",
                sender=getattr(mail_item.Sender, 'Name', mail_item.SenderEmailAddress or "Unknown"),
                recipients=[rec.Address for rec in mail_item.Recipients],
                content=mail_item.Body or "",
                raw_content=mail_item.Body or "",
                timestamp=mail_item.ReceivedTime,
                message_id=mail_item.EntryID
            )
        except Exception as e:
            logger.error(f"转换邮件数据失败: {e}")
            return None
    
    def _find_account_by_keyword(self, keyword: str):
        """通过关键字查找账户"""
        for account in self.namespace.Accounts:
            if keyword.lower() in account.SmtpAddress.lower():
                return account
        return None

    def _cleanup(self):
        """清理COM对象"""
        self.inbox = None
        self.namespace = None
        self.outlook_app = None
        pythoncom.CoUninitialize()