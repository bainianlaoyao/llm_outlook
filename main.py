#!/usr/bin/env python3
"""
自动邮件处理系统 - 主入口点
整合IMAP邮件拉取、z.ai邮件解析、消息推送的完整系统

使用方法:
  python main.py                    # 使用默认配置文件
  python main.py --config config.json   # 指定配置文件
  python main.py --help            # 显示帮助信息
"""

import argparse
import json
import signal
import sys
import os
from pathlib import Path
from typing import Optional

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from core.main_controller import PureOutlookController, create_outlook_controller_from_config
from config.defaults import CONFIG_FILE_PATH, DEFAULT_CONFIG_TEMPLATE
from utils.logger import get_logger, log_info


class AutoMailSystem:
    """自动邮件处理系统主类"""

    def __init__(self):
        self.controller = None
        self.logger = get_logger('auto_mail_system')
        self.running = False

        # 设置信号处理器
        signal.signal(signal.SIGINT, self._signal_handler)
        signal.signal(signal.SIGTERM, self._signal_handler)

    def run(self, config_file: Optional[str] = None, create_config: bool = False):
        """
        运行系统 - 一次性执行模式

        Args:
            config_file: 配置文件路径
            create_config: 是否创建示例配置文件
        """
        try:
            # 1. 解析配置
            if create_config:
                self._create_sample_config()
                return

            # 2. 加载配置并创建控制器
            self.logger.info("正在初始化系统...")
            if config_file:
                config_path = Path(config_file)
            else:
                config_path = Path(CONFIG_FILE_PATH)

            if not config_path.exists():
                self.logger.error(f"配置文件不存在: {config_path}")
                self.logger.info("请使用 --create-config 创建示例配置文件，或指定 --config 参数")
                return

            self.logger.info(f"使用配置文件: {config_path}")
            self.controller = create_outlook_controller_from_config(str(config_path))

            # 3. 启动系统 - 一次性执行
            self.running = True
            log_info("执行一次性邮件处理")
            success = self.controller.start()

            if success:
                self.logger.info("系统正常退出")
            else:
                self.logger.error("系统启动失败")
                return 1

        except KeyboardInterrupt:
            self.logger.info("接收到中断信号，正在停止...")
        except Exception as e:
            self.logger.error(f"系统运行异常: {e}")
            return 1
        finally:
            self._cleanup()

        return 0

    def _signal_handler(self, signum, frame):
        """信号处理器"""
        signal_name = "SIGINT" if signum == signal.SIGINT else "SIGTERM"
        self.logger.info(f"接收到信号 {signal_name}，正在停止系统...")

        if self.controller:
            self.controller.stop()

        self.running = False

    def _cleanup(self):
        """清理资源"""
        if self.controller:
            self.controller.stop()

        log_info("自动邮件处理系统已停止")
        self.logger.info("系统清理完成")

    def _create_sample_config(self):
        """生成示例配置文件"""
        config_path = Path(CONFIG_FILE_PATH)

        try:
            # 确保配置目录存在
            config_path.parent.mkdir(parents=True, exist_ok=True)

            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(DEFAULT_CONFIG_TEMPLATE, f, indent=2, ensure_ascii=False)

            self.logger.info(f"示例配置文件已创建: {config_path}")
            print(f"\n示例配置文件已创建: {config_path}")
            print("\n请编辑配置文件，填入正确的邮箱和API信息，然后运行:")
            print(f"python {__file__}")

        except Exception as e:
            self.logger.error(f"创建示例配置文件失败: {e}")
            print(f"创建配置文件失败: {e}")


def create_sample_config():
    """创建示例配置文件"""
    system = AutoMailSystem()
    system._create_sample_config()


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='自动邮件处理系统 - 一次性执行模式，拉取→解析→推送后退出',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python main.py                           # 使用默认配置执行一次
  python main.py --config my_config.json   # 指定配置文件执行
  python main.py --create-config           # 创建示例配置文件

配置文件格式:
{
  "email": {
    "address": "your-email@example.com",
    "password": "your-app-password"
  },
  "imap": {
    "server": "imap.example.com",
    "port": 993
  },
  "parser": {
    "api_key": "your-zai-api-key"
  },
  "push": {
    "channel": "serverchan",
    "sendkey": "your-serverchan-sendkey"
  }
}
        """
    )

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

    parser.add_argument(
        '--version', '-v',
        action='version',
        version='自动邮件处理系统 v1.0.0'
    )

    args = parser.parse_args()

    # 创建系统实例
    system = AutoMailSystem()

    # 设置日志级别（可以扩展为命令行参数）
    os.environ.setdefault('LOG_LEVEL', 'INFO')

    # 运行系统
    exit_code = system.run(
        config_file=args.config,
        create_config=args.create_config
    )

    sys.exit(exit_code)


if __name__ == '__main__':
    main()
