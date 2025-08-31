"""配置模块

提供应用所需的静态配置项。

说明:
- 本文件将所有配置集中到一个可导入的结构中，便于在应用中引用。
- API 密钥按任务要求以硬编码方式保留在此文件中。
"""
from typing import Dict, Any

CONFIG: Dict[str, Any] = {
    "outlook": {
        "profile": "Outlook",
    },
    "parser": {
        "language": "auto",
        "api_key": "", #zai api key
    },
    "push": {
        "channel": "serverchan",
        "sendkey": "", # serverchan api key
    },
    "log": {
        "level": "INFO",
        "file_path": "logs/auto_mail.log",
    },
}

def get_config() -> Dict[str, Any]:
    """返回配置的浅拷贝，避免外部直接修改全局配置对象。"""
    return CONFIG.copy()

# 方便直接从模块导入常用项
OUTLOOK_PROFILE: str = CONFIG["outlook"]["profile"]
PARSER_LANGUAGE: str = CONFIG["parser"]["language"]
PARSER_API_KEY: str = CONFIG["parser"]["api_key"]
PUSH_CHANNEL: str = CONFIG["push"]["channel"]
PUSH_SENDKEY: str = CONFIG["push"]["sendkey"]
LOG_LEVEL: str = CONFIG["log"]["level"]
LOG_FILE_PATH: str = CONFIG["log"]["file_path"]
ANALYZE_DAYS: int = 1