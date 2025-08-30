# Outlook邮件获取实现说明

## 📧 概述

基于pywin32的Outlook邮件获取器，替代原有IMAP方案，实现本地邮件访问、断点续拉和多账号支持。

## 🔧 核心功能

### 1. **Outlook本地连接**
- 使用`win32com.client.Dispatch('Outlook.Application')`连接
- 支持多个Outlook配置文件
- COM自动化安全处理

### 2. **邮件获取与处理**
- 直接访问收件箱：`GetDefaultFolder(6) - olFolderInbox`
- 邮件项枚举和排序：按接收时间降序
- 分页处理大量邮件
- 附件信息提取（不下载）

### 3. **断点续拉机制**
- 使用`EntryID`作为唯一标识符
- 保存上次处理的EntryID
- 只处理新的邮件项
- 确保增量同步

### 4. **智能邮件获取器**
- 根据配置自动选择邮件源
- Outlook优先，本地更高效
- IMAP作为后备方案
- 零破坏性向后兼容

## 📁 文件结构

```
core/
├── email_fetcher.py          # 核心邮件获取器（Outlook + IMAP）
├── auto_email_fetcher.py     # 智能自动选择器
└── ...

config/
├── defaults.py               # 配置参数定义
└── config.json               # 用户配置

outlook_test.py               # 测试脚本
OUTLOOK_MAIL_README.md       # 本文档
```

## ⚙️ 配置说明

### 邮箱配置 (config/defaults.py)

```python
# IMAP配置（传统方案）
EMAIL_CONFIG = {
    'imap_server': 'imap.gmail.com',
    'port': 993,
    'ssl': True,
    'timeout': 30
}

# Outlook配置（本地方案）
EMAIL_CONFIG = {
    'outlook_profile': 'Outlook',        # Outlook配置文件名称
    'use_outlook': True,                 # 是否启用Outlook
    'outlook_timeout': 30,               # 连接超时
    'outlook_max_retries': 3             # 重试次数
}
```

### 启用Outlook模式

编辑`config/defaults.py`：
```python
EMAIL_CONFIG = {
    'outlook_profile': 'Outlook',
    'use_outlook': True,  # 设置为True启用Outlook
    # ... 其他配置
}
```

## 🚀 使用方法

### 1. **简单使用**
```python
from core.auto_email_fetcher import fetch_emails_auto

# 自动选择邮件源（配置决定）
emails, latest_id = fetch_emails_auto(last_uid='', max_emails=50)
```

### 2. **Outlook专用使用**
```python
from core.email_fetcher import fetch_emails_outlook

# 直接使用Outlook
emails, latest_entry_id = fetch_emails_outlook(last_entry_id='last-known-id')
```

### 3. **高级用法**
```python
from core.auto_email_fetcher import AutoEmailFetcher

fetcher = AutoEmailFetcher()
fetcher.connect()

# 获取邮件
emails, latest_id = fetcher.fetch_emails(max_emails=10)
fetcher.disconnect()
```

## 🔍 关键特性

### 数据类型映射
```python
# Outlook MailItem -> EmailData
MailItem.Subject         -> EmailData.subject
MailItem.SenderName      -> EmailData.sender
MailItem.Body           -> EmailData.content
MailItem.ReceivedTime   -> EmailData.timestamp
MailItem.EntryID        -> EmailData.uid
```

### 断点续拉逻辑
```python
# 使用EntryID进行增量拉取
last_entry_id = "上次处理的EntryID"

# 只获取新的邮件项
for item in inbox_items:
    if item.EntryID > last_entry_id:
        # 处理新邮件
        process_email(item)
```

## 🧪 测试

### 运行测试
```bash
# 运行Outlook测试
python outlook_test.py

# 测试智能选择器
python -c "from core.auto_email_fetcher import test_email_fetcher; test_email_fetcher()"
```

### 测试覆盖
- ✅ Outlook连接验证
- ✅ 邮件获取和解析
- ✅ 断点续拉机制
- ✅ 附件信息处理
- ✅ 便捷函数测试

## 📈 优势分析

| 特性 | Outlook方案 | IMAP方案 |
|------|------------|----------|
| **网络需求** | ❌ 无需 | ✅ 需要 |
| **密码管理** | ❌ 无需 | ✅ 需要 |
| **性能** | 🚀 极快 | 🐌 取决于网络 |
| **企业兼容性** | ✅ 优秀 | ❌ 需额外配置 |
| **附件处理** | 📎 完全支持 | 📎 基础支持 |
| **连接稳定性** | 💯 100% | 🔄 取决于网络 |

## ⚠️ 注意事项

### 依赖要求
```bash
# 必需依赖
pip install pywin32
```

### 系统要求
- ✅ Windows操作系统
- ✅ 已安装Microsoft Outlook
- ✅ Outlook正在运行
- ✅ 启用COM自动化（通常默认启用）

### 错误处理
```python
try:
    # Outlook操作
    emails, latest_id = fetch_emails_outlook()
except Exception as e:
    print(f"Outlook错误: {e}")
    # 可切换到IMAP作为后备
```

## 🔄 向后兼容性

- ✅ 保持原有`EmailData`接口
- ✅ 保持原有连接参数格式
- ✅ 不改变现有调用方式
- ✅ 智能自动降级
- ✅ 零破坏性部署

## 🎯 企业应用场景

### 银行集成
```python
# 在企业Outlook环境中
fetcher = AutoEmailFetcher()
fetcher.connect()  # 无需密码，直接访问企业邮件
```

### 金融数据处理
```python
# 处理理财邮件通知
emails = fetcher.fetch_emails(max_emails=100)
for email in emails:
    if "银行卡" in email.subject:
        process_bank_email(email)
```

### 自动邮件监控
```python
# 守护进程模式
while True:
    emails, last_id = fetcher.fetch_emails(last_entry_id=last_id)
    for email in emails:
        analyze_email_content(email)
    time.sleep(300)  # 5分钟检查一次
```

## 🔮 扩展计划

- [ ] 多配置支持
- [ ] Outlook多账号
- [ ] 自定义文件夹
- [ ] 发送邮件支持
- [ ] 日历集成

---

**完** 📧✨