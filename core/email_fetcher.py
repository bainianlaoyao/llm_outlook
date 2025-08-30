import win32com.client
from typing import List, Optional
from datetime import datetime
import pythoncom

from models.email import EmailData


class OutlookEmailFetcher:
    """基于Outlook的邮件获取器"""

    def __init__(self, profile: str = "Outlook"):
        self.profile = profile  # Outlook配置文件名称
        self.outlook_app = None
        self.namespace = None
        self.inbox = None

    def connect(self) -> bool:
        """连接到Outlook"""
        try:
            pythoncom.CoInitialize()
            self.outlook_app = win32com.client.Dispatch('Outlook.Application')
            self.namespace = self.outlook_app.GetNamespace('MAPI')

            # 获取默认收件箱
            self.inbox = self.namespace.GetDefaultFolder(6)  # olFolderInbox = 6
            return True
        except Exception as e:
            print(f"Outlook连接失败: {str(e)}")
            self._cleanup()
            return False

    def disconnect(self):
        """断开连接"""
        self._cleanup()

    def _cleanup(self):
        """清理资源"""
        if self.inbox:
            self.inbox = None
        if self.namespace:
            self.namespace = None
        if self.outlook_app:
            self.outlook_app = None
        pythoncom.CoUninitialize()

    def fetch_emails(self, last_entry_id: str = '', max_emails: int = 50) -> tuple[List[EmailData], str]:
        """拉取新邮件

        参数:
            last_entry_id: 上次处理的最后EntryID
            max_emails: 最大拉取数量

        返回:
            (邮件列表, 最后EntryID)
        """
        if not self._ensure_connection():
            return [], last_entry_id

        try:
            # 获取邮件项，按接收时间降序排序
            items = self.inbox.Items
            if hasattr(items, 'Sort'):
                items.Sort("[ReceivedTime]", False)  # False表示降序

            email_list = []
            latest_entry_id = last_entry_id

            # 遍历邮件
            count = getattr(items, 'Count', 0)
            for i in range(min(max_emails, count)):
                try:
                    mail_item = items[i]
                    entry_id = str(getattr(mail_item, 'EntryID', ''))

                    # 跳过已处理的邮件
                    if last_entry_id and entry_id == last_entry_id:
                        break

                    email_data = self._mailitem_to_emaildata(mail_item, entry_id)
                    if email_data:
                        email_list.append(email_data)
                        latest_entry_id = entry_id

                except Exception as e:
                    print(f"处理邮件失败: {str(e)}")
                    continue

            return email_list, latest_entry_id

        except Exception as e:
            print(f"拉取邮件失败: {str(e)}")
            return [], last_entry_id

    def _ensure_connection(self) -> bool:
        """确保连接正常"""
        if not self.outlook_app:
            return self.connect()
        return True

    def _mailitem_to_emaildata(self, mail_item, entry_id: str) -> Optional[EmailData]:
        """将Outlook MailItem转换为EmailData"""
        try:
            # 提取基本信息
            subject = self._get_subject(mail_item)
            sender = self._get_sender(mail_item)
            recipients = self._get_recipients(mail_item)
            content = self._get_content(mail_item)
            raw_content = mail_item.Body  # 使用Body作为原始内容
            timestamp = mail_item.ReceivedTime
            message_id = mail_item.EntryID  # 使用EntryID作为MessageID

            return EmailData(
                uid=entry_id,
                subject=subject,
                sender=sender,
                recipients=recipients,
                content=content,
                raw_content=raw_content,
                timestamp=timestamp,
                message_id=message_id
            )

        except Exception as e:
            print(f"转换邮件数据失败: {str(e)}")
            return None

    def _get_subject(self, mail_item) -> str:
        """获取邮件主题"""
        try:
            return str(mail_item.Subject) if mail_item.Subject else ""
        except:
            return ""

    def _get_sender(self, mail_item) -> str:
        """获取发件人"""
        try:
            if mail_item.SenderName:
                return str(mail_item.SenderName)
            elif mail_item.SenderEmailAddress:
                return str(mail_item.SenderEmailAddress)
            else:
                return "Unknown"
        except:
            return "Unknown"

    def _get_recipients(self, mail_item) -> list:
        """获取收件人列表"""
        try:
            recipients = []
            if hasattr(mail_item, 'Recipients') and mail_item.Recipients:
                for recipient in mail_item.Recipients:
                    if recipient.Address:
                        recipients.append(str(recipient.Address))
            return recipients
        except:
            return []

    def _get_content(self, mail_item) -> str:
        """获取邮件内容"""
        try:
            # 优先使用纯文本内容
            if hasattr(mail_item, 'Body') and mail_item.Body:
                return str(mail_item.Body)
            # 如果没有纯文本，使用HTML内容的文本版本
            elif hasattr(mail_item, 'HTMLBody'):
                # 将HTML转换为纯文本的简单实现
                html_content = str(mail_item.HTMLBody)
                # 这里可以添加HTML到文本的转换逻辑
                return self._html_to_text(html_content)
            else:
                return ""
        except:
            return ""

    def _html_to_text(self, html: str) -> str:
        """简单的HTML到文本转换"""
        try:
            # 移除HTML标签的基本实现
            import re
            # 移除HTML标签
            text = re.sub(r'<[^>]+>', '', html)
            # 移除多余的空白字符
            text = re.sub(r'\s+', ' ', text).strip()
            return text
        except:
            return html

    def get_attachments_info(self, mail_item) -> List[str]:
        """获取附件信息（不下载）"""
        attachments_info = []
        try:
            if hasattr(mail_item, 'Attachments') and mail_item.Attachments:
                for attachment in mail_item.Attachments:
                    info = f"附件: {attachment.FileName} ({attachment.Size} bytes)"
                    attachments_info.append(info)
        except Exception as e:
            print(f"获取附件信息失败: {str(e)}")
        return attachments_info


# 全局连接池用于多账号支持
_outlook_connection_pool = {}

def get_or_create_outlook_fetcher(profile: str = "Outlook") -> OutlookEmailFetcher:
    """获取或创建Outlook邮件获取器实例"""
    if profile not in _outlook_connection_pool:
        _outlook_connection_pool[profile] = OutlookEmailFetcher(profile)
    return _outlook_connection_pool[profile]


def fetch_emails_outlook(last_entry_id: str = '', max_emails: int = 50,
                         profile: str = "Outlook") -> tuple[List[EmailData], str]:
    """使用Outlook拉取邮件的便捷函数

    参数:
        last_entry_id: 上次处理的最后EntryID
        max_emails: 最大拉取数量
        profile: Outlook配置文件名称

    返回:
        (邮件列表, 最后EntryID)
    """
    fetcher = get_or_create_outlook_fetcher(profile)
    emails, latest_entry_id = fetcher.fetch_emails(last_entry_id, max_emails)
    fetcher.disconnect()  # 每次调用后自动断开，释放资源
    return emails, latest_entry_id


# IMAP支持已被移除，专注于Outlook本地解决方案
# 如需IMAP支持，请使用专门的IMAP客户端