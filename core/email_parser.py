"""
邮件解析核心模块
使用z.ai API解析邮件内容并生成结构化摘要
"""

from typing import Optional, List
from zai import ZhipuAiClient
import os
import re
from utils.logger import get_logger

from models.email import EmailData

logger = get_logger(__name__)


class EmailParser:
    """邮件解析器类"""

    def __init__(self, api_key: Optional[str] = None, logger=None):
        self.api_key = api_key or os.getenv('ZAI_API_KEY')
        self.client = ZhipuAiClient(api_key=self.api_key)

    def parse_emails_batch(self, email_list: List[EmailData], language: Optional[str] = None) -> str:
        """
        批量解析邮件内容，生成周总结
        """
        if not email_list:
            return "无邮件"

        prompt = self._build_batch_prompt(email_list, language or 'auto')
        
        messages = [
            {"role": "system", "content": "你是一个专业的邮件分析助手。请分析一批邮件并生成结构化总结。"},
            {"role": "user", "content": prompt}
        ]
        
        try:
            response = self.client.chat.completions.create(
                model="glm-4.5-air",
                messages=messages,
                temperature=0.6,
                stream=False
            )
            ai_response = response.choices[0].message.content
            return self._parse_batch_response(ai_response)
        except Exception as e:
            logger.error(f"批量解析邮件异常: {e}")
            return "解析失败"

    def _build_batch_prompt(self, email_list: List[EmailData], language: str) -> str:
        """构建批量解析提示语"""
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

{''.join(email_details)}

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

    def _parse_batch_response(self, response: str) -> str:
        """解析批量AI响应并组装结果"""
        pattern = r'\$(.*?)\$'
        matches = re.findall(pattern, response, re.DOTALL)
        if matches:
            return matches[0].strip()
        else:
            return "未找到$...$格式内容"

def parse_emails_batch(email_list: List[EmailData], language: Optional[str] = None, api_key: Optional[str] = None) -> str:
    """
    批量解析邮件内容的便捷函数
    """
    parser = EmailParser(api_key=api_key)
    return parser.parse_emails_batch(email_list, language)
