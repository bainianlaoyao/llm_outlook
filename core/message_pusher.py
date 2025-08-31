"""
消息推送核心模块
实现多通道推送功能，重点支持Server酱消息推送
"""

import json
import os
import time
import traceback
from typing import Optional
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
from urllib.parse import urlencode

from models.pusher import PushResult
from utils.logger import get_logger, log_push_message


class ServerChanPusher:
    """Server酱推送器"""

    def __init__(self, sendkey: Optional[str] = None, logger=None):
        self.sendkey = sendkey or os.getenv('SERVERCHAN_SENDKEY', '')
        self.url = 'https://sctapi.ftqq.com/'
        self.max_retry = 3
        self.retry_delay = 10
        self.logger = logger or get_logger('serverchan_pusher')

    def push(self, title: str, content: str = "") -> PushResult:
        """推送消息到Server酱"""
        if not self.sendkey:
            return PushResult(False, "未配置Server酱SendKey")

        data = {'title': title, 'desp': content}
        encoded_data = urlencode(data).encode('utf-8')
        push_url = f"{self.url}{self.sendkey}.send"
        req = Request(push_url, data=encoded_data, method='POST')
        req.add_header('Content-Type', 'application/x-www-form-urlencoded')

        for attempt in range(self.max_retry):
            try:
                with urlopen(req, timeout=30) as response:
                    result = json.loads(response.read().decode('utf-8'))
                if result.get('code') == 0:
                    log_push_message('Server酱', True)
                    return PushResult(True, "推送成功", result.get('data', {}).get('pushid', ''))
                else:
                    return PushResult(False, f"推送失败: {result.get('message', '未知错误')}")
            except (HTTPError, URLError) as e:
                self.logger.warning(f"推送失败 (尝试 {attempt + 1}): {e}")
                if attempt < self.max_retry - 1:
                    time.sleep(self.retry_delay)
            except Exception as e:
                self.logger.error(f"推送时出现未预期异常: {e}\n{traceback.format_exc()}")

        log_push_message('Server酱', False)
        return PushResult(False, f"推送失败，已重试{self.max_retry}次")


class MultiChannelPusher:
    """多通道推送管理器"""

    def __init__(self, channels: Optional[list] = None,
                 serverchan_sendkey: Optional[str] = None,
                 logger=None):
        self.channels = channels or ['serverchan']
        self.logger = logger or get_logger('multi_channel_pusher')
        self.pushers = {}
        if 'serverchan' in self.channels:
            self.pushers['serverchan'] = ServerChanPusher(
                sendkey=serverchan_sendkey,
                logger=self.logger
            )

    def push(self, title: str, content: str = "") -> PushResult:
        """推送消息到所有配置的渠道"""
        if not self.pushers:
            return PushResult(False, "未配置任何推送渠道")

        for channel, pusher in self.pushers.items():
            result = pusher.push(title, content)
            if result.success:
                return result
        
        return PushResult(False, "所有推送渠道均失败")
