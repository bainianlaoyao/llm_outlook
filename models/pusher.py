from dataclasses import dataclass


@dataclass
class PushResult:
    """消息推送结果数据结构"""
    success: bool
    message: str
    push_id: str

    def __init__(self, success: bool = False, message: str = "", push_id: str = ""):
        self.success = success
        self.message = message
        self.push_id = push_id

    def to_dict(self) -> dict:
        """转换为字典格式"""
        return {
            'success': self.success,
            'message': self.message,
            'push_id': self.push_id
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'PushResult':
        """从字典创建PushResult对象"""
        return cls(
            success=data.get('success', False),
            message=data.get('message', ''),
            push_id=data.get('push_id', '')
        )