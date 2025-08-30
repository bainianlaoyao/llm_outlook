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

    def parse_emails_batch(self, email_list: List[EmailData], language: Optional[str] = None) -> str:
        """
        批量解析邮件内容，生成周总结

        Args:
            email_list: 邮件数据列表
            language: 解析语言

        Returns:
            WeeklySummaryResult: 周解析结果对象
        """
        if not email_list:
            return "无邮件"

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
                stream=False
            )
            
            # 解析响应 - 使用与单邮件解析相同的逻辑
            ai_response = self._extract_api_response_content(response)
            
            return self._parse_batch_response(ai_response, email_list)
            
        except Exception as e:
            logger.error(f"批量解析邮件异常: {e}")
            # 返回失败结果
            return "解析失败"
    def _build_batch_prompt(self, email_list: List[EmailData], language: str) -> str:
        """构建批量解析提示语 - 精简版"""
        
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
        
        return f"""请分析以上{len(email_list)}封上周邮件，生成结构化总结：

{chr(10).join(email_details)}

**分析要求:**
一步步思考
1. 判断是否有重要邮件（工作相关、紧急事务、重要通知， 学校事务等， 判断遵循不可漏过原则等）
2. 对所有重要邮件的进行总结（200字以内）
3. 为剩余邮件生成80字以内的简短总结
4. 确保总结简洁明了，突出关键信息


**输出格式(自行使用优雅的markdown排版, 不要遗漏$$号, 这是重要的程序化处理符号)**
一步步的思考分析过程

$
总体评价(例: 本次分析中有3封邮件值得终点关注, 其中标题为"xxx"的邮件尤其值得关注, 其次, 其余邮件中标题为"xxx"的邮件可能也值得一看)

重要邮件
---
邮件1: 
原标题
内容总结
邮件2:
原标题
内容总结
---

其余邮件
邮件1: 
原标题
内容总结
邮件2:
原标题
内容总结
...
邮件N: [总结N]
$
"""


    def _parse_batch_response(self, response: str, email_list: List[EmailData]) -> str:
        """解析批量AI响应并组装结果 - 精简版"""
        import re
        # 使用正则表达式提取$xx$之间的内容
        pattern = r'\$(.*?)\$'
        matches = re.findall(pattern, response, re.DOTALL)
        if matches:
            # 返回第一个匹配内容（如有多个可根据需求调整）
            return matches[0].strip()
        else:
            return "未找到$...$格式内容"
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


def parse_emails_batch(email_list: List[EmailData], language: Optional[str] = None, api_key: Optional[str] = None) -> str:
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