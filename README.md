# auto_mail — Outlook 本地邮件解析与推送

简短描述：基于本地 Outlook COM 接口，拉取邮件、调用 z.ai 解析并通过 Server酱 等通道推送摘要。

## 特性
- 使用本地 Outlook COM（pywin32）直接读取邮件，零网络依赖用于企业环境
- 断点续拉与去重（基于 EntryID）
- 批量 AI 解析（使用 z.ai，见 [`core/email_parser.py`](core/email_parser.py:1)）
- 多通道推送支持（Server酱 实现见 [`core/message_pusher.py`](core/message_pusher.py:1)）
- 简洁控制器与一次性执行流程（见 [`core/main_controller.py`](core/main_controller.py:1)）

## 系统要求
- Windows（需安装 Microsoft Outlook 并启用 COM 自动化）
- Python 3.8+

## 安装
1. 克隆仓库并进入项目根目录

```bash
git clone <repo-url>
cd auto_mail
```

2. 安装依赖

```bash
pip install pywin32 zai
```

（可选）如果使用 virtualenv：

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 zai
```

## 配置
默认配置位于 [`config/config_template.py`](config/config_template.py:1)（运行时由 `config.config.get_config()` 加载），常用项说明：

```python
# 示例：config/config_template.py 中的关键配置结构
CONFIG = {
    "outlook": {
        "profile": "Outlook",
    },
    "parser": {
        "api_key": "", # zai API key，或使用环境变量 ZAI_API_KEY
    },
    "push": {
        "channel": "serverchan",
        "sendkey": "", # Server酱 SendKey 或通过环境变量 SERVERCHAN_SENDKEY
    },
    "log": {
        "level": "INFO",
        "file_path": "logs/auto_mail.log",
    },
}
```

说明：
- 修改 [`config/config_template.py`](config/config_template.py:1) 或通过环境变量（例如 `ZAI_API_KEY`、`SERVERCHAN_SENDKEY`）提供敏感凭据。
- 日志配置可参考 [`config/defaults.py`](config/defaults.py:1)。

## 使用说明
- 运行一次性任务（主入口）：[`main.py`](main.py:1)

```bash
python main.py
```

- 代码示例（在 Python 中调用控制器）：

```python
from core.main_controller import PureOutlookController
from config.config import get_config

cfg = get_config()
controller = PureOutlookController(cfg)
controller.start()
```

模块说明：
- [`core/outlook_email_fetcher.py`](core/outlook_email_fetcher.py:1)：通过 COM 获取邮件并转换为 `EmailData`
- [`core/email_parser.py`](core/email_parser.py:1)：调用 z.ai 生成结构化摘要
- [`core/message_pusher.py`](core/message_pusher.py:1)：Server酱推送与多通道管理
- [`core/main_controller.py`](core/main_controller.py:1)：拉取→解析→推送的一次性执行流

## 日志与数据
- 日志文件默认：`logs/auto_mail.log`（可在配置中修改）
- 已处理邮件 ID 存储于 `data/processed_entry_ids.txt`、`data/last_entry_id.txt`（用于断点续拉）

## 贡献
感谢贡献。请遵循以下流程：
1. fork 仓库并新建分支 feature/xxx
2. 编写单元测试并确保本地通过
3. 提交 PR，描述变更并关联 issue（如有）

## 许可证
本项目采用 MIT 许可证。详见 LICENSE 文件（若无，请考虑补充）。

---
参见旧说明：[`OUTLOOK_MAIL_README.md`](OUTLOOK_MAIL_README.md:1) 以获取更多实现细节。
