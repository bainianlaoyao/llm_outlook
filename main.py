#!/usr/bin/env python3
"""
自动邮件处理系统 - 主入口点
"""

import argparse
import json
import signal
import sys
import os
from pathlib import Path
from typing import Optional

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from core.main_controller import PureOutlookController
from config.config import get_config
from utils.logger import get_logger, log_info


class AutoMailSystem:
    """自动邮件处理系统主类"""

    def __init__(self):
        self.controller = None
        self.logger = get_logger('auto_mail_system')
        signal.signal(signal.SIGINT, self._signal_handler)
        signal.signal(signal.SIGTERM, self._signal_handler)

    def run(self, config_file: Optional[str] = None):
        """运行系统 - 一次性执行模式"""

        config = get_config()

        try:
            self.controller = PureOutlookController(config)
            if self.controller.start():
                self.logger.info("系统正常退出")
                return 0
            else:
                self.logger.error("系统启动失败")
                return 1
        except Exception as e:
            self.logger.error(f"系统运行异常: {e}")
            return 1
        finally:
            if self.controller:
                self.controller.stop()
            log_info("自动邮件处理系统已停止")

    def _signal_handler(self, signum, frame):
        """信号处理器"""
        self.logger.info(f"接收到信号 {signal.Signals(signum).name}，正在停止系统...")
        if self.controller:
            self.controller.stop()



def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='自动邮件处理系统 - 一次性执行模式')
    parser.add_argument(
        '--config', '-c',
        type=str,
        help='指定配置文件路径（已弃用，配置来自 config.config）'
    )
    args = parser.parse_args()

    system = AutoMailSystem()
    exit_code = system.run(config_file=args.config)
    sys.exit(exit_code)


if __name__ == '__main__':
    main()