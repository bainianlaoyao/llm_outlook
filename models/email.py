from dataclasses import dataclass
from typing import List, Optional
from datetime import datetime


@dataclass
class EmailData:
    """邮件数据结构"""
    uid: str
    subject: str
    sender: str
    recipients: List[str]
    content: str
    raw_content: str
    timestamp: datetime
    message_id: str


    def to_dict(self) -> dict:
        """转换为字典格式"""
        return {
            'uid': self.uid,
            'subject': self.subject,
            'sender': self.sender,
            'recipients': self.recipients,
            'content': self.content,
            'raw_content': self.raw_content,
            'timestamp': self.timestamp.isoformat(),
            'message_id': self.message_id
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'EmailData':
        """从字典创建EmailData对象"""
        return cls(
            uid=data['uid'],
            subject=data.get('subject', ''),
            sender=data.get('sender', ''),
            recipients=data.get('recipients', []),
            content=data.get('content', ''),
            raw_content=data.get('raw_content', ''),
            timestamp=datetime.fromisoformat(data['timestamp']) if 'timestamp' in data else datetime.now(),
            message_id=data.get('message_id', '')
        )