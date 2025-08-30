"""
邮件解析核心模块
使用z.ai API解析邮件内容并生成结构化摘要
"""

from typing import Optional, List, Dict
from zai import ZhipuAiClient
from utils.logger import get_logger

from models.email import EmailData
from models.parser import ParseResult, EmailSummaryResult

logger = get_logger(__name__)


class EmailParser:
    """邮件解析器类 - Linus式精简版"""

    def __init__(self, api_key: Optional[str] = None, logger=None):
        self.api_key = api_key or ""
        self.model = "glm-4.5"
        self.timeout = 30
        self.max_retries = 3
        self.retry_delay = 5
        self.max_tokens = 2000
        self.temperature = 0.6
        self.client = ZhipuAiClient(api_key="6a74bb8623194ed881d47eb5f8d2ba3e.xUi3KT2NQJcUn0EA")

    def parse_email(self, email_data: EmailData, language: Optional[str] = None) -> ParseResult:
        """
        解析邮件内容

        Args:
            email_data: 邮件数据对象
            language: 解析语言('auto'表示自动检测)

        Returns:
            ParseResult: 解析结果对象
        """
        # 直接调用API，最简单的数据流
        response = self._call_api(email_data, language or 'auto')
        return self._parse_response(email_data, response)

    def _call_api(self, email_data: EmailData, language: str):
        """调用z.ai API - Linus式精简版"""
        
        # 直接构建prompt，消除中间转换
        prompt = self._build_prompt(email_data, language)
        
        messages = [
            {"role": "system", "content": "你是一个专业的邮件解析助手。请按指定格式解析邮件内容。"},
            {"role": "user", "content": prompt}
        ]
        
        # 直接SDK调用，零多余操作
        return self.client.chat.completions.create(
            model="glm-4.5-air",
            messages=messages,
            temperature=0.6,
            max_tokens=2000,
            stream=False
        )

    def _parse_response(self, email_data: EmailData, response) -> ParseResult:
        """解析API响应并组装结果"""
        try:
            # 兼容真实API响应和模拟数据
            if hasattr(response, 'choices'):
                # 真实API响应
                ai_response = response.choices[0].message.content
            else:
                # 模拟数据（字典格式）
                ai_response = response["choices"][0]["message"]["content"]
            
            # 简化解析逻辑
            parsed_data = self._parse_structured_response(ai_response)
            
            return ParseResult(
                summary=parsed_data.get('summary', '解析失败'),
                key_points=parsed_data.get('key_points', []),
                content=parsed_data.get('content', email_data.content),
                attachments=parsed_data.get('attachments', []),
                category=parsed_data.get('category', 'general')
            )
            
        except Exception as e:
            # 简化异常处理，直接返回失败结果
            return ParseResult('解析失败', [], email_data.content)

    def _parse_structured_response(self, response: str):
        """解析结构化的AI响应 - 精简版"""
        parsed_data = {
            'summary': '',
            'key_points': [],
            'category': 'general',
            'content': '',
            'attachments': []
        }

        for line in response.split('\n'):
            if line.startswith(('摘要:', 'Summary:')):
                parsed_data['summary'] = line.split(':', 1)[1].strip()
            elif line.startswith(('关键点:', 'Key Points:')):
                points = line.split(':', 1)[1].strip()
                parsed_data['key_points'] = [p.strip() for p in points.split('|') if p.strip()]
            elif line.startswith(('类别:', 'Category:')):
                parsed_data['category'] = line.split(':', 1)[1].strip()
            elif line.startswith(('附件:', 'Attachments:')):
                attachments = line.split(':', 1)[1].strip()
                parsed_data['attachments'] = [a.strip() for a in attachments.split('|') if a.strip()]
            elif line.startswith(('内容:', 'Content:')):
                parsed_data['content'] = line.split(':', 1)[1].strip()

        return parsed_data

    def _build_prompt(self, email_data: EmailData, language: str) -> str:
        """构建解析提示语 - 精简版"""
        language_instruction = "" if language == 'auto' else f"请用{language}语言回复。"
        
        return f"""请解析以下邮件，提取关键信息并生成结构化摘要:

{language_instruction}

**邮件内容:**
发件人: {email_data.sender}
收件人: {', '.join(email_data.recipients)}
主题: {email_data.subject}
时间: {email_data.timestamp}

内容:
{email_data.content}
{f'\n原始内容:\n{email_data.raw_content}' if email_data.raw_content else ''}

**解析要求:**
1. 生成简洁的邮件摘要(不超过100字)
2. 提取3-5个关键信息点
3. 确定邮件类别(工作、通知、广告、个人等)
4. 保留重要内容片段
5. 如果有附件，列出附件名称

**输出格式:**
摘要: [摘要内容]
关键点: [点1] | [点2] | [点3]
类别: [邮件类别]
内容: [重要内容片段]
附件: [附件1] | [附件2]"""


    def _extract_api_response_content(self, response) -> str:
        """提取API响应内容 - 与现有单邮件解析逻辑保持一致"""
        # 兼容真实API响应和模拟数据
        if hasattr(response, 'choices'):
            # 真实API响应
            return response.choices[0].message.content
        else:
            # 模拟数据（字典格式）
            return response["choices"][0]["message"]["content"]

    def parse_emails_batch(self, email_list: List[EmailData], language: Optional[str] = None) -> EmailSummaryResult:
        """
        批量解析邮件内容，生成周总结

        Args:
            email_list: 邮件数据列表
            language: 解析语言

        Returns:
            WeeklySummaryResult: 周解析结果对象
        """
        if not email_list:
            return EmailSummaryResult(
                has_important_emails=False,
                overall_summary="本周无新邮件",
                email_summaries={},
                push_title="本周邮件总结",
                push_content="本周无新邮件"
            )

        # 构建批量prompt
        prompt = self._build_batch_prompt(email_list, language or 'auto')
        
        messages = [
            {"role": "system", "content": "你是一个专业的邮件分析助手。请分析一批邮件并生成结构化总结。"},
            {"role": "user", "content": prompt}
        ]
        
        try:
            # 调用API
            response = self.client.chat.completions.create(
                model="glm-4.5-air",
                messages=messages,
                temperature=0.6,
                max_tokens=4000,  # 增加token限制以处理批量邮件
                stream=False
            )
            
            # 解析响应 - 使用与单邮件解析相同的逻辑
            ai_response = self._extract_api_response_content(response)
            
            return self._parse_batch_response(ai_response, email_list)
            
        except Exception as e:
            logger.error(f"批量解析邮件异常: {e}")
            # 返回失败结果
            return EmailSummaryResult(
                has_important_emails=False,
                overall_summary=f"解析失败: {str(e)}",
                email_summaries={i: f"解析失败: {str(e)}" for i in range(len(email_list))},
                push_title="邮件解析失败",
                push_content=f"无法解析本周邮件: {str(e)}"
            )

    def _build_batch_prompt(self, email_list: List[EmailData], language: str) -> str:
        """构建批量解析提示语 - 精简版"""
        language_instruction = "" if language == 'auto' else f"请用{language}语言回复。"
        
        # 构建邮件列表
        email_details = []
        for i, email in enumerate(email_list):
            email_details.append(f"""
**邮件 {i+1}:**
发件人: {email.sender}
主题: {email.subject}
时间: {email.timestamp}
内容: {email.content[:200]}{'...' if len(email.content) > 200 else ''}
""")
        
        return f"""请分析以下{len(email_list)}封上周邮件，生成结构化总结：

{language_instruction}

{chr(10).join(email_details)}

**分析要求:**
1. 判断是否有重要邮件（工作相关、紧急事务、重要通知等）
2. 生成所有邮件的总体总结（200字以内）
3. 为每封邮件生成80字以内的简短总结
4. 确保总结简洁明了，突出关键信息

**输出格式:**
重要邮件: [是/否]
总体总结: [总结内容， 首先单独给出重要邮件的概括， 然后给出剩余邮件的概括]
邮件分别总结（无论如何， 为每一封邮件的内容都生成简短总结， 不要只使用标题）:
邮件1: [总结1]
邮件2: [总结2]
...
邮件N: [总结N]"""

    def _parse_batch_response(self, response: str, email_list: List[EmailData]) -> EmailSummaryResult:
        """解析批量AI响应并组装结果 - 精简版"""
        try:
            lines = response.strip().split('\n')
            
            has_important = False
            overall_summary = ""
            email_summaries = {}
            
            # 解析响应
            current_section = None
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                    
                if line.startswith(('重要邮件:', 'Important Emails:')):
                    current_section = 'important'
                    has_important = line.split(':', 1)[1].strip().lower() in ['是', 'yes', 'true']
                elif line.startswith(('总体总结:', 'Overall Summary:')):
                    current_section = 'summary'
                    overall_summary = line.split(':', 1)[1].strip()
                elif line.startswith(('邮件总结:', 'Email Summaries:')):
                    current_section = 'summaries'
                elif current_section == 'summaries' and line.startswith(('邮件', 'Email')) and ':' in line:
                    # 解析邮件总结，如 "邮件1: 总结内容"
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        try:
                            email_num = int(parts[0].replace('邮件', '').replace('Email', '').strip())
                            if 1 <= email_num <= len(email_list):
                                email_summaries[email_num - 1] = parts[1].strip()[:80]  # 限制80字
                        except ValueError:
                            pass
            
            # 如果某些邮件没有总结，生成默认总结
            for i in range(len(email_list)):
                if i not in email_summaries:
                    email_summaries[i] = f"{email_list[i].subject[:50]}{'...' if len(email_list[i].subject) > 50 else ''}"
            
            # 生成推送内容
            push_title = "📧 重要邮件" if has_important else "📧 本周邮件总结"
            push_content = f"重要邮件: {'是' if has_important else '否'}\n\n{overall_summary}\n\n"
            push_content += "邮件总结:\n" + "\n".join([f"• {summary}" for summary in email_summaries.values()])
            
            return EmailSummaryResult(
                has_important_emails=has_important,
                overall_summary=overall_summary or "暂无总结",
                email_summaries=email_summaries,
                push_title=push_title,
                push_content=push_content
            )
            
        except Exception as e:
            logger.error(f"解析批量响应失败: {e}")
            # 返回默认结果
            return EmailSummaryResult(
                has_important_emails=False,
                overall_summary=f"解析响应失败: {str(e)}",
                email_summaries={i: email_list[i].subject[:80] for i in range(len(email_list))},
                push_title="邮件解析异常",
                push_content=f"解析响应时出错: {str(e)}"
            )


def parse_email(email_data: EmailData, language: Optional[str] = None, api_key: Optional[str] = None) -> ParseResult:
    """
    解析邮件内容的便捷函数 - 保持完全兼容

    Args:
        email_data: 邮件数据对象
        language: 解析语言
        api_key: API密钥

    Returns:
        ParseResult: 解析结果对象
    """
    parser = EmailParser(api_key=api_key)
    return parser.parse_email(email_data, language)


def parse_emails_batch(email_list: List[EmailData], language: Optional[str] = None, api_key: Optional[str] = None) -> EmailSummaryResult:
    """
    批量解析邮件内容的便捷函数 - 新增功能

    Args:
        email_list: 邮件数据列表
        language: 解析语言
        api_key: API密钥

    Returns:
        WeeklySummaryResult: 周解析结果对象
    """
    parser = EmailParser(api_key=api_key)
    return parser.parse_emails_batch(email_list, language)