#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
纯Outlook本地邮件获取器
零网络依赖，直接通过COM接口访问本地Outlook邮件

核心特性：
- 100%本地无网络连接
- 使用Windows账户认证，无需密码
- 本地EntryID缓存，精确追踪已处理邮件
- 极快的COM接口访问速度
- 企业级安全性，无敏感信息存储

作者：Linus风格架构师
原则：简洁、实用、最少的概念
"""

import win32com.client
from typing import List, Optional, Tuple, Dict
import pythoncom
import os
from datetime import datetime, timedelta
import time

from models.email import EmailData
from utils.logger import get_logger

logger = get_logger(__name__)


class OutlookEmailFetcher:
    """纯Outlook本地邮件获取器"""

    def __init__(self, profile: str = "Outlook", email_keyword: Optional[str] = None):
        """
        初始化Outlook获取器

        参数:
            profile: Outlook配置文件名称
            email_keyword: 邮箱地址关键字，用于选择特定账户
        """
        # Outlook基础配置
        self.profile = profile
        self.email_keyword = email_keyword

        # COM对象
        self.outlook_app = None
        self.namespace = None
        self.inbox = None

        # 处理状态 - 在内存中维护已处理的EntryID集合
        self.processed_ids = set()

        logger.info(f"初始化Outlook获取器: profile={profile}, email_keyword={email_keyword}")

    def connect(self) -> bool:
        """
        连接到本地Outlook，支持通过关键字选择账户

        返回:
            bool: 连接是否成功
        """
        try:
            # 初始化COM环境
            pythoncom.CoInitialize()

            # 获取Outlook应用
            self.outlook_app = win32com.client.Dispatch('Outlook.Application')
            
            # 如果指定了邮箱关键字，尝试选择匹配的账户
            if self.email_keyword:
                selected_account = self._select_account_by_email_keyword(self.email_keyword)
                if selected_account:
                    # 使用指定账户的namespace并尝试切换
                    self.namespace = self.outlook_app.GetNamespace('MAPI')
                    # 尝试切换到指定账户
                    if self._switch_to_account(selected_account):
                        logger.info(f"✓ 选择账户: {selected_account}")
                    else:
                        logger.warning(f"切换到账户 {selected_account} 失败，使用默认账户")
                else:
                    # 使用默认namespace
                    self.namespace = self.outlook_app.GetNamespace('MAPI')
                    logger.warning(f"未找到匹配关键字 '{self.email_keyword}' 的账户，使用默认账户")
            else:
                # 使用默认namespace
                self.namespace = self.outlook_app.GetNamespace('MAPI')
                logger.info("✓ 使用默认账户")

            # 获取收件箱
            self.inbox = self.namespace.GetDefaultFolder(6)  # olFolderInbox = 6

            logger.info("✓ Outlook连接成功")
            return True

        except Exception as e:
            logger.error(f"✗ Outlook连接失败: {e}")
            self._cleanup()
            return False

    def fetch_emails(self, last_days: int = 7) -> Tuple[List[EmailData], str]:
        """
        获取指定天数内的新邮件，移除数量限制

        参数:
            last_days: 获取过去几天的邮件（默认7天）

        返回:
            (邮件列表, 最新EntryID)
        """
        if not self._ensure_connection():
            logger.error("连接不可用，无法获取邮件")
            return [], ""

        try:
            # 获取邮件项目并按接收时间排序
            items = self.inbox.Items  # type: ignore
            items.Sort("[ReceivedTime]", True)  # 降序（最新在前）

            email_list = []
            latest_entry_id = ""

            # 按时间过滤：获取过去last_days天的所有邮件
            cutoff_time = datetime.now() - timedelta(days=last_days)
            current_time = datetime.now()

            logger.info(f"=== 邮件获取调试信息 ===")
            logger.info(f"当前时间: {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
            logger.info(f"截止时间: {cutoff_time.strftime('%Y-%m-%d %H:%M:%S')}")
            logger.info(f"时间范围: 过去 {last_days} 天")
            logger.info(f"Outlook邮箱总邮件数: {getattr(items, 'Count', 0)}")
            logger.info(f"已处理邮件数: {len(self.processed_ids)}")
            
            # 显示最近3封邮件的时间信息
            logger.info(f"=== 最近邮件时间信息 ===")
            for i in range(min(3, getattr(items, 'Count', 0))):
                try:
                    mail_item = items[i]
                    received_time = getattr(mail_item, 'ReceivedTime', None)
                    subject = getattr(mail_item, 'Subject', '无主题')
                    entry_id = str(getattr(mail_item, 'EntryID', ''))[:10]
                    
                    if received_time:
                        if received_time.tzinfo is not None:
                            received_time = received_time.astimezone().replace(tzinfo=None)
                        logger.info(f"邮件[{i}] - 时间:{received_time}, 主题:{subject[:30]}, EntryID:{entry_id}...")
                    else:
                        logger.info(f"邮件[{i}] - 无时间信息, 主题:{subject[:30]}, EntryID:{entry_id}...")
                except Exception as e:
                    logger.warning(f"读取邮件[{i}]信息失败: {e}")

            # 遍历所有邮件，检查时间
            for i in range(getattr(items, 'Count', 0)):
                try:
                    mail_item = items[i]
                    received_time = getattr(mail_item, 'ReceivedTime', None)
                    entry_id = str(getattr(mail_item, 'EntryID', ''))

                    # 时间过滤：只处理指定时间范围内的邮件
                    if received_time:
                        # 确保两个时间都是naive datetime（无时区信息）进行比较
                        if received_time.tzinfo is not None:
                            # 如果邮件时间是aware的，转换为本地时间
                            received_time = received_time.astimezone().replace(tzinfo=None)

                        logger.debug(f"邮件分析[{i}] - 接收时间: {received_time}, 截止时间: {cutoff_time}, EntryID: {entry_id[:10]}...")

                        if received_time >= cutoff_time:
                            # 跳过已处理邮件
                            if entry_id in self.processed_ids:
                                logger.debug(f"邮件[{i}] 已被处理，跳过")
                                continue

                            # 转换邮件数据
                            email_data = self._convert_to_email_data(mail_item, entry_id)
                            if email_data:
                                email_list.append(email_data)
                                latest_entry_id = entry_id
                                self.processed_ids.add(entry_id)
                                logger.debug(f"邮件[{i}] 符合时间要求，已添加到列表")
                        else:
                            logger.debug(f"邮件[{i}] 早于时间范围，停止遍历")
                            # 由于邮件按时间排序，一旦遇到早于指定时间的邮件就可以停止
                            break
                    else:
                        logger.debug(f"邮件[{i}] 无时间信息，跳过")
                        # 如果没有时间信息，跳过
                        continue

                except Exception as e:
                    logger.warning(f"处理邮件失败 #{i}: {e}")
                    continue

            logger.info(f"获取 {len(email_list)} 封过去 {last_days} 天的新邮件")

            # 保存处理状态 - 已移除缓存功能
            # self._save_state()

            return email_list, latest_entry_id

        except Exception as e:
            logger.error(f"拉取邮件异常: {e}")
            return [], ""

    def fetch_all_emails(self) -> Tuple[List[EmailData], str]:
        """
        获取所有邮件（不进行时间过滤），用于调试

        返回:
            (邮件列表, 最新EntryID)
        """
        if not self._ensure_connection():
            logger.error("连接不可用，无法获取邮件")
            return [], ""

        try:
            # 获取邮件项目并按接收时间排序
            items = self.inbox.Items  # type: ignore
            items.Sort("[ReceivedTime]", False)  # 降序（最新在前）

            email_list = []
            latest_entry_id = ""

            logger.info(f"=== 获取所有邮件调试信息 ===")
            logger.info(f"Outlook邮箱总邮件数: {getattr(items, 'Count', 0)}")
            logger.info(f"已处理邮件数: {len(self.processed_ids)}")

            # 遍历所有邮件
            for i in range(getattr(items, 'Count', 0)):
                try:
                    mail_item = items[i]
                    received_time = getattr(mail_item, 'ReceivedTime', None)
                    entry_id = str(getattr(mail_item, 'EntryID', ''))

                    # 跳过已处理邮件
                    if entry_id in self.processed_ids:
                        logger.debug(f"邮件[{i}] 已被处理，跳过")
                        continue

                    # 转换邮件数据
                    email_data = self._convert_to_email_data(mail_item, entry_id)
                    if email_data:
                        email_list.append(email_data)
                        latest_entry_id = entry_id
                        self.processed_ids.add(entry_id)
                        logger.debug(f"邮件[{i}] 已添加到列表")
                        
                        # 只显示前10封邮件的详细信息
                        if i < 10:
                            logger.info(f"邮件[{i}] - 主题:{email_data.subject[:30]}, 时间:{email_data.timestamp}")

                except Exception as e:
                    logger.warning(f"处理邮件失败 #{i}: {e}")
                    continue

            logger.info(f"获取 {len(email_list)} 封未处理邮件")

            # 保存处理状态 - 已移除缓存功能
            # self._save_state()

            return email_list, latest_entry_id

        except Exception as e:
            logger.error(f"拉取所有邮件异常: {e}")
            return [], ""

    def fetch_weekly_emails(self) -> List[EmailData]:
        """
        获取上一周的所有邮件

        返回:
            邮件列表
        """
        if not self._ensure_connection():
            logger.error("连接不可用，无法获取邮件")
            return []

        try:
            # 计算一周前的时间
            one_week_ago = datetime.now() - timedelta(days=7)
            
            # 获取邮件项目并按接收时间排序
            items = self.inbox.Items  # type: ignore
            items.Sort("[ReceivedTime]", False)  # 降序（最新在前）

            email_list = []

            # 遍历邮件，筛选上一周的邮件
            for i in range(getattr(items, 'Count', 0)):
                try:
                    mail_item = items[i]
                    received_time = getattr(mail_item, 'ReceivedTime', None)
                    
                    # 检查邮件时间是否在上一周范围内
                    if received_time and received_time >= one_week_ago:
                        entry_id = str(getattr(mail_item, 'EntryID', ''))
                        
                        # 转换邮件数据
                        email_data = self._convert_to_email_data(mail_item, entry_id)
                        if email_data:
                            email_list.append(email_data)
                    else:
                        # 由于邮件按时间排序，一旦遇到早于一周的邮件就可以停止
                        break

                except Exception as e:
                    logger.warning(f"处理邮件失败 #{i}: {e}")
                    continue

            logger.info(f"获取 {len(email_list)} 封上周邮件")
            return email_list

        except Exception as e:
            logger.error(f"拉取周邮件异常: {e}")
            return []

    def disconnect(self):
        """断开连接"""
        self._cleanup()
        logger.info("Outlook连接已断开")

    def _ensure_connection(self) -> bool:
        """确保连接正常"""
        if not self.outlook_app:
            return self.connect()
        return True

    def _convert_to_email_data(self, mail_item, entry_id: str) -> Optional[EmailData]:
        """
        将Outlook MailItem转换为EmailData

        参数:
            mail_item: Outlook邮件对象
            entry_id: Outlook EntryID

        返回:
            EmailData对象或None
        """
        try:
            # 提取基本信息
            subject = str(getattr(mail_item, 'Subject', ''))
            sender = self._extract_sender(mail_item)
            recipients = self._extract_recipients(mail_item)
            content = self._extract_content(mail_item)
            timestamp = getattr(mail_item, 'ReceivedTime', datetime.now())

            # 返回EmailData对象
            return EmailData(
                uid=entry_id,
                subject=subject,
                sender=sender,
                recipients=recipients,
                content=content,
                raw_content=str(getattr(mail_item, 'Body', '')),
                timestamp=timestamp,
                message_id=str(getattr(mail_item, 'EntryID', entry_id))
            )

        except Exception as e:
            logger.error(f"转换邮件数据失败: {e}")
            return None

    def _extract_sender(self, mail_item) -> str:
        """提取发件人"""
        try:
            # 优先使用SenderName
            if hasattr(mail_item, 'SenderName') and mail_item.SenderName:
                return str(mail_item.SenderName)

            # 回退到SenderEmailAddress
            elif hasattr(mail_item, 'SenderEmailAddress') and mail_item.SenderEmailAddress:
                return str(mail_item.SenderEmailAddress)

            else:
                return "Unknown Sender"

        except Exception as e:
            logger.debug(f"提取发件人失败: {e}")
            return "Unknown Sender"

    def _extract_recipients(self, mail_item) -> List[str]:
        """提取收件人列表"""
        try:
            recipients = []
            if hasattr(mail_item, 'Recipients') and mail_item.Recipients:
                for recipient in mail_item.Recipients:
                    if hasattr(recipient, 'Address') and recipient.Address:
                        recipients.append(str(recipient.Address))

            return recipients

        except Exception as e:
            logger.debug(f"提取收件人失败: {e}")
            return []

    def _extract_content(self, mail_item) -> str:
        """提取邮件内容"""
        try:
            # 优先使用纯文本内容
            if hasattr(mail_item, 'Body') and mail_item.Body:
                return str(mail_item.Body)

            # 尝试HTML转文本
            elif hasattr(mail_item, 'HTMLBody'):
                return self._html_to_text(str(mail_item.HTMLBody))

            else:
                return ""

        except Exception as e:
            logger.debug(f"提取邮件内容失败: {e}")
            return ""

    def _html_to_text(self, html: str) -> str:
        """简单的HTML到文本转换"""
        try:
            import re
            # 移除HTML标签
            text = re.sub(r'<[^>]+>', '', html)
            # 清理空白字符
            text = re.sub(r'\s+', ' ', text).strip()
            return text
        except:
            return html

    def _cleanup(self):
        """清理COM对象"""
        if self.outlook_app:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

        self.inbox = None
        self.namespace = None
        self.outlook_app = None

    def _select_account_by_email_keyword(self, keyword: str) -> Optional[str]:
        """
        通过邮箱地址关键字选择账户

        参数:
            keyword: 邮箱地址关键字

        返回:
            匹配的账户名称，如果没有匹配则返回None
        """
        try:
            # 获取所有可用账户
            accounts = self._get_available_accounts()
            
            # 查找匹配关键字的账户
            for account_info in accounts:
                account_name, email_address = account_info
                
                # 检查邮箱地址是否包含关键字
                if keyword.lower() in email_address.lower():
                    logger.info(f"找到匹配账户: {account_name} ({email_address})")
                    return account_name
            
            logger.warning(f"未找到匹配关键字 '{keyword}' 的账户")
            return None
            
        except Exception as e:
            logger.error(f"选择账户时发生错误: {e}")
            return None

    def _switch_to_account(self, account_name: str) -> bool:
        """
        切换到指定账户

        参数:
            account_name: 账户名称

        返回:
            bool: 切换是否成功
        """
        try:
            # 确保已经连接
            if not self.outlook_app or not self.namespace:
                logger.warning("Outlook未连接，无法切换账户")
                return False
            
            # 方法1：尝试通过不同的存储切换账户
            try:
                # 获取所有存储（stores），每个存储对应一个账户
                stores = self.namespace.Stores
                
                # 查找匹配的账户存储
                for store in stores:
                    store_name = store.DisplayName
                    if account_name.lower() in store_name.lower():
                        # 获取该存储的收件箱
                        inbox_folder = store.GetDefaultFolder(6)  # olFolderInbox
                        if inbox_folder:
                            self.inbox = inbox_folder
                            logger.info(f"✓ 成功切换到账户收件箱: {store_name}")
                            return True
                
                logger.debug(f"未找到匹配账户 {account_name} 的存储")
                
            except Exception as e:
                logger.debug(f"通过存储切换账户失败: {e}")
            
            # 方法2：尝试通过不同的配置文件切换
            try:
                # 创建新的namespace并指定配置文件
                new_namespace = self.outlook_app.GetNamespace('MAPI')
                # 尝试使用不同的配置文件
                if hasattr(new_namespace, 'Logon'):
                    new_namespace.Logon(account_name)
                    self.namespace = new_namespace
                    logger.info(f"✓ 成功切换到账户: {account_name}")
                    return True
                else:
                    logger.debug("Logon方法不可用")
                    
            except Exception as e:
                logger.debug(f"使用配置文件 {account_name} 切换失败: {e}")
            
            logger.warning(f"无法切换到账户: {account_name}")
            return False
            
        except Exception as e:
            logger.error(f"切换账户时发生错误: {e}")
            return False

    def _get_available_accounts(self) -> List[Tuple[str, str]]:
        """
        获取所有可用的Outlook账户及其邮箱地址

        返回:
            账户名称和邮箱地址的元组列表
        """
        try:
            accounts = []
            
            # 确保已经连接
            if not self.outlook_app or not self.namespace:
                logger.warning("Outlook未连接，无法获取账户列表")
                return [(self.profile, self.profile)]
            
            # 获取所有MAPI配置文件
            try:
                profiles = self.namespace.Session.Profiles
                for i in range(1, profiles.Count + 1):
                    profile_name = profiles[i].Name
                    
                    # 尝试获取该配置文件的默认邮箱地址
                    try:
                        # 获取默认邮箱地址
                        if hasattr(self.namespace, 'CurrentUserAddress'):
                            email_address = self.namespace.CurrentUserAddress
                        else:
                            # 如果无法获取邮箱地址，使用profile名称作为邮箱
                            email_address = profile_name
                        
                        accounts.append((profile_name, email_address))
                        
                    except Exception as e:
                        logger.debug(f"获取配置文件 {profile_name} 的邮箱地址失败: {e}")
                        # 即使无法获取邮箱地址，也添加账户信息
                        accounts.append((profile_name, profile_name))
                
            except Exception as e:
                logger.debug(f"获取配置文件列表失败: {e}")
                # 如果无法获取配置文件列表，使用当前配置文件
                accounts.append((self.profile, self.profile))
            
            logger.info(f"找到 {len(accounts)} 个可用账户")
            return accounts
            
        except Exception as e:
            logger.error(f"获取可用账户失败: {e}")
            # 回退方案：返回当前配置文件
            return [(self.profile, self.profile)]


def fetch_outlook_emails(profile: str = "Outlook", last_days: int = 7) -> Tuple[List[EmailData], str]:
    """
    便捷函数：获取Outlook邮件

    参数:
        profile: Outlook配置文件
        last_days: 获取过去几天的邮件（默认7天）

    返回:
        (邮件列表, 最新EntryID)
    """
    fetcher = OutlookEmailFetcher(profile)
    emails, latest_id = fetcher.fetch_emails(last_days)
    fetcher.disconnect()
    return emails, latest_id


def fetch_weekly_outlook_emails(profile: str = "Outlook") -> List[EmailData]:
    """
    便捷函数：获取上一周的Outlook邮件

    参数:
        profile: Outlook配置文件

    返回:
        邮件列表
    """
    fetcher = OutlookEmailFetcher(profile)
    emails = fetcher.fetch_weekly_emails()
    fetcher.disconnect()
    return emails