import os
from openai import OpenAI
from dotenv import load_dotenv

class DocumentAI:
    def __init__(self):
        load_dotenv()
        self.client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
        
    def analyze_document(self, text):
        """
        使用AI分析文档内容，识别各个部分
        """
        prompt = """
        请分析以下学术论文内容，识别并返回以下部分：
        1. 标题
        2. 摘要
        3. 关键词
        4. 各章节标题及其层级
        5. 参考文献

        请以JSON格式返回结果。
        
        文档内容：
        {text}
        """
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "你是一个专业的学术论文结构分析助手。"},
                    {"role": "user", "content": prompt.format(text=text)}
                ],
                temperature=0.3
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"AI分析出错: {str(e)}")
            return None
    
    def suggest_formatting(self, section_type, content):
        """
        为特定类型的内容提供格式建议
        """
        prompt = f"""
        请为以下{section_type}内容提供格式建议，包括：
        1. 建议的字体大小
        2. 段落间距
        3. 缩进要求
        4. 其他相关的格式要求
        
        内容：
        {content}
        """
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "你是一个专业的学术论文格式专家。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"AI建议出错: {str(e)}")
            return None 