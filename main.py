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

from core.main_controller import create_outlook_controller_from_config
from config.defaults import CONFIG_FILE_PATH, DEFAULT_CONFIG_TEMPLATE
from utils.logger import get_logger, log_info


class AutoMailSystem:
    """自动邮件处理系统主类"""

    def __init__(self):
        self.controller = None
        self.logger = get_logger('auto_mail_system')
        signal.signal(signal.SIGINT, self._signal_handler)
        signal.signal(signal.SIGTERM, self._signal_handler)

    def run(self, config_file: Optional[str] = None, create_config: bool = False):
        """运行系统 - 一次性执行模式"""
        if create_config:
            self._create_sample_config()
            return 0

        config_path = Path(config_file or CONFIG_FILE_PATH)
        if not config_path.exists():
            self.logger.error(f"配置文件不存在: {config_path}")
            self.logger.info("请使用 --create-config 创建示例配置文件，或指定 --config 参数")
            return 1

        try:
            self.logger.info(f"使用配置文件: {config_path}")
            self.controller = create_outlook_controller_from_config(str(config_path))
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

    def _create_sample_config(self):
        """生成示例配置文件"""
        config_path = Path(CONFIG_FILE_PATH)
        try:
            config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(DEFAULT_CONFIG_TEMPLATE, f, indent=2, ensure_ascii=False)
            self.logger.info(f"示例配置文件已创建: {config_path}")
        except Exception as e:
            self.logger.error(f"创建示例配置文件失败: {e}")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='自动邮件处理系统 - 一次性执行模式')
    parser.add_argument(
        '--config', '-c',
        type=str,
        help=f'指定配置文件路径 (默认: {CONFIG_FILE_PATH})'
    )
    parser.add_argument(
        '--create-config',
        action='store_true',
        help='创建示例配置文件并退出'
    )
    args = parser.parse_args()

    system = AutoMailSystem()
    exit_code = system.run(
        config_file=args.config,
        create_config=args.create_config
    )
    sys.exit(exit_code)


if __name__ == '__main__':
    main()