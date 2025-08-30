import os

# 纯Outlook本地邮件系统配置
# 零网络依赖，企业级安全解决方案

# 配置文件路径
CONFIG_FILE_PATH = 'config/config.json'

# 日志配置
LOG_CONFIG = {
    'level': 'INFO',
    'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    'file_path': 'logs/auto_mail.log',
    'max_file_size': 10485760,  # 10MB
    'backup_count': 5
}

# 示例配置文件模板
DEFAULT_CONFIG_TEMPLATE = {
    "outlook": {
        "profile": "Outlook"  # Outlook配置文件名
    },
    "parser": {
        "language": "auto",
        "api_key": "${ZAI_API_KEY}"
    },
    "push": {
        "channel": "serverchan",
        "sendkey": "${SERVERCHAN_SENDKEY}"
    },
    "log": {
        "level": "INFO",
        "file_path": "logs/auto_mail.log"
    }
}
