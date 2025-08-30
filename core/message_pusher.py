"""
消息推送核心模块
实现多通道推送功能，重点支持Server酱消息推送
"""

import json
import os
import time
from typing import Dict, Any, Optional
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
from urllib.parse import urlencode

from models.pusher import PushResult
from utils.logger import get_logger, log_push_message


class ServerChanPusher:
    """Server酱推送器"""

    def __init__(self, sendkey: Optional[str] = None, url: Optional[str] = None, logger=None):
        # 从参数或环境变量获取sendkey
        self.sendkey = 'SCT154542THtyS9cBPpp5PuiHwzGaIXeR8'
        self.url = url or 'https://sctapi.ftqq.com/'
        self.max_retry = 3
        self.retry_delay = 10
        self.logger = logger or get_logger('serverchan_pusher')

    def push(self, title: str, content: str = "") -> PushResult:
        """推送消息到Server酱"""
        if not self.sendkey:
            return PushResult(False, "未配置Server酱SendKey")

        # 构建推送数据
        data = {
            'title': title,
            'desp': content
        }

        # 发送推送请求
        for attempt in range(self.max_retry):
            try:
                push_url = f"{self.url}{self.sendkey}.send"
                encoded_data = urlencode(data).encode('utf-8')

                req = Request(push_url, data=encoded_data, method='POST')
                req.add_header('Content-Type', 'application/x-www-form-urlencoded')

                self.logger.info(f"发送推送请求，尝试 {attempt + 1}/{self.max_retry}")

                with urlopen(req, timeout=30) as response:
                    result = json.loads(response.read().decode('utf-8'))

                # 检查推送结果
                if result.get('code') == 0:
                    log_push_message('Server酱', True)
                    return PushResult(True, "推送成功", result.get('data', {}).get('pushid', ''))
                else:
                    error_msg = result.get('message', '未知错误')
                    return PushResult(False, f"推送失败: {error_msg}")

            except HTTPError as e:
                error_msg = f"HTTP错误: {e.code} - {e.reason}"
                self.logger.error(error_msg)
                if e.code == 400:
                    return PushResult(False, "推送参数错误")

            except URLError as e:
                error_msg = f"网络错误: {e.reason}"
                self.logger.error(error_msg)

            except Exception as e:
                error_msg = f"未知错误: {e}"
                self.logger.error(error_msg)

            if attempt < self.max_retry - 1:
                time.sleep(self.retry_delay)

        log_push_message('Server酱', False)
        return PushResult(False, f"推送失败，已重试{self.max_retry}次")


class MultiChannelPusher:
    """多通道推送管理器"""

    def __init__(self, channels: Optional[list] = None,
                 serverchan_sendkey: Optional[str] = None,
                 logger=None):
        # 硬编码配置：只保留serverchan渠道
        self.channels = channels or ['serverchan']
        self.logger = logger or get_logger('multi_channel_pusher')

        # 初始化推送器
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

        results = []
        for channel, pusher in self.pushers.items():
            try:
                result = pusher.push(title, content)
                if result.success:
                    return result  # 任一渠道推送成功即返回成功
                results.append(result)
            except Exception as e:
                self.logger.error(f"{channel}推送异常: {e}")
                results.append(PushResult(False, f"{channel}异常: {e}"))

        # 所有渠道都失败了
        if results:
            message = "; ".join([r.message for r in results])
            return PushResult(False, f"所有推送渠道失败: {message}")

        return PushResult(False, "推送失败")


def push_message(summary: str,
                 channels: Optional[list] = None,
                 serverchan_sendkey: Optional[str] = None) -> PushResult:
    """
    推送消息的主要入口函数

    Args:
        summary: 消息摘要内容
        channels: 推送渠道列表，默认为配置中的渠道
        serverchan_sendkey: Server酱推送密钥

    Returns:
        PushResult: 推送结果对象
    """
    # 提取标题和内容
    lines = summary.strip().split('\n', 1)
    title = lines[0][:64]  # 限制标题长度
    content = lines[1] if len(lines) > 1 else ""

    # 创建推送器并发送
    pusher = MultiChannelPusher(channels=channels,
                              serverchan_sendkey=serverchan_sendkey)
    return pusher.push(title, content)


# 便捷函数 - Server酱专用
def push_to_serverchan(title: str, content: str = "", sendkey: Optional[str] = None) -> PushResult:
    """
    推送消息到Server酱的便捷函数

    Args:
        title: 消息标题
        content: 消息内容
        sendkey: Server酱SendKey

    Returns:
        PushResult: 推送结果对象
    """
    pusher = ServerChanPusher(sendkey=sendkey)
    return pusher.push(title, content)