import os
from openai import OpenAI
from dotenv import load_dotenv
import json
from typing import Optional, Dict, Any

class DocumentAI:
    def __init__(self, model: str = "gpt-3.5-turbo"):
        """
        初始化DocumentAI
        Args:
            model: 使用的OpenAI模型名称
        """
        load_dotenv()
        self.api_key = os.getenv('OPENAI_API_KEY')
        if not self.api_key:
            raise ValueError("未找到OPENAI_API_KEY环境变量")
            
        self.client = OpenAI(api_key=self.api_key)
        self.model = model
        
    def analyze_document(self, text: str) -> Optional[Dict[str, Any]]:
        """
        分析文档内容，识别各个部分
        Args:
            text: 文档内容
        Returns:
            解析后的文档结构，如果解析失败返回None
        """
        prompt = """
        请分析以下学术论文内容，识别并返回以下部分：
        1. 标题 (title)
        2. 摘要 (abstract)
        3. 关键词 (keywords)
        4. 各章节标题及其层级 (sections)
        5. 参考文献 (references)

        请以JSON格式返回结果，格式如下：
        {
            "title": "论文标题",
            "abstract": "摘要内容",
            "keywords": ["关键词1", "关键词2"],
            "sections": [
                {
                    "title": "章节标题",
                    "level": 1,
                    "content": "章节内容"
                }
            ],
            "references": [
                "参考文献1",
                "参考文献2"
            ]
        }

        文档内容：
        {text}
        """
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一个专业的学术论文结构分析专家。"},
                    {"role": "user", "content": prompt.format(text=text)}
                ],
                temperature=0.3,
                max_tokens=2000
            )
            
            result = json.loads(response.choices[0].message.content)
            return result
        except Exception as e:
            print(f"AI分析出错: {str(e)}")
            return None
    
    def suggest_formatting(self, section_type: str, content: str) -> Optional[Dict[str, Any]]:
        """
        为特定类型的内容提供格式建议
        Args:
            section_type: 部分类型（如title, abstract等）
            content: 需要格式化的内容
        Returns:
            格式建议，如果分析失败返回None
        """
        format_requirements = {
            "title": "标题格式要求：字体、大小、对齐方式等",
            "abstract": "摘要格式要求：段落缩进、行距等",
            "keywords": "关键词格式要求：字体、间距等",
            "heading": "标题格式要求：字号、对齐等",
            "body": "正文格式要求：字体、段落等",
            "references": "参考文献格式要求：缩进、间距等"
        }
        
        prompt = f"""
        请为以下{section_type}内容提供详细的格式建议。

        需要考虑的格式要素：
        1. 字体大小（磅值）
        2. 字体类型
        3. 段落间距
        4. 行距
        5. 对齐方式
        6. 缩进要求
        7. 其他特殊要求

        请以JSON格式返回建议，包含以下字段：
        - font_size: 字号（磅）
        - font_name: 字体名称
        - line_spacing: 行距
        - paragraph_spacing: 段落间距
        - alignment: 对齐方式
        - indent: 缩进值
        - special_requirements: 其他特殊要求

        参考格式要求：
        {format_requirements.get(section_type, "通用格式要求")}

        内容：
        {content}
        """
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一个专业的学术论文格式专家。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3
            )
            
            result = json.loads(response.choices[0].message.content)
            return result
        except Exception as e:
            print(f"格式建议生成失败: {str(e)}")
            return None

    def validate_format(self, content: str, format_spec: Dict[str, Any]) -> bool:
        """
        验证内容是否符合指定的格式要求
        Args:
            content: 要验证的内容
            format_spec: 格式规范
        Returns:
            是否符合格式要求
        """
        prompt = f"""
        请验证以下内容是否符合格式要求。

        格式要求：
        {json.dumps(format_spec, indent=2)}

        内容：
        {content}

        请返回 true 或 false，并说明原因。
        """
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一个专业的格式验证专家。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3
            )
            
            result = response.choices[0].message.content.lower()
            return 'true' in result
        except Exception as e:
            print(f"格式验证失败: {str(e)}")
            return False 