from dataclasses import dataclass
from typing import List, Optional, Dict


@dataclass
class ParseRequest:
    """邮件解析请求数据结构"""
    email: str
    language: Optional[str] = None

    def __init__(self, email: str, language: Optional[str] = None):
        self.email = email
        self.language = language

    def to_dict(self) -> dict:
        """转换为字典格式"""
        return {
            'email': self.email,
            'language': self.language
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'ParseRequest':
        """从字典创建ParseRequest对象"""
        return cls(
            email=data['email'],
            language=data.get('language')
        )


@dataclass
class ParseResult:
    """邮件解析结果数据结构"""
    summary: str
    key_points: List[str]
    content: str
    attachments: List[str]
    category: str

    def __init__(self, summary: str = "", key_points: Optional[List[str]] = None,
                 content: str = "", attachments: Optional[List[str]] = None, category: str = "general"):
        self.summary = summary
        self.key_points = key_points or []
        self.content = content
        self.attachments = attachments or []
        self.category = category

    def to_dict(self) -> dict:
        """转换为字典格式"""
        return {
            'summary': self.summary,
            'key_points': self.key_points,
            'content': self.content,
            'attachments': self.attachments,
            'category': self.category
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'ParseResult':
        """从字典创建ParseResult对象"""
        return cls(
            summary=data.get('summary', ''),
            key_points=data.get('key_points', []),
            content=data.get('content', ''),
            attachments=data.get('attachments', []),
            category=data.get('category', 'general')
        )


@dataclass
class EmailSummaryResult:
    """周邮件解析结果数据结构"""
    # 整体分析
    has_important_emails: bool  # 是否有可能重要的邮件
    overall_summary: str        # 所有邮件的内容总结
    
    # 单邮件总结 (邮件索引 -> 80字简短总结)
    email_summaries: Dict[int, str]
    
    # 推送标题
    push_title: str
    
    # 推送内容
    push_content: str

    def __init__(self, has_important_emails: bool = False, overall_summary: str = "",
                 email_summaries: Optional[Dict[int, str]] = None, push_title: str = "",
                 push_content: str = ""):
        self.has_important_emails = has_important_emails
        self.overall_summary = overall_summary
        self.email_summaries = email_summaries or {}
        self.push_title = push_title
        self.push_content = push_content

    def to_dict(self) -> dict:
        """转换为字典格式"""
        return {
            'has_important_emails': self.has_important_emails,
            'overall_summary': self.overall_summary,
            'email_summaries': self.email_summaries,
            'push_title': self.push_title,
            'push_content': self.push_content
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'EmailSummaryResult':
        """从字典创建WeeklySummaryResult对象"""
        return cls(
            has_important_emails=data.get('has_important_emails', False),
            overall_summary=data.get('overall_summary', ''),
            email_summaries=data.get('email_summaries', {}),
            push_title=data.get('push_title', ''),
            push_content=data.get('push_content', '')
        )