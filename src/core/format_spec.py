from dataclasses import dataclass
from typing import Dict, Optional, List
import json
import yaml
from pathlib import Path

@dataclass
class SectionFormat:
    font_size: float
    font_name: str = "Times New Roman"
    bold: bool = False
    italic: bool = False
    alignment: str = "LEFT"
    first_line_indent: float = 0
    line_spacing: float = 1.0
    space_before: float = 0
    space_after: float = 0

@dataclass
class DocumentFormat:
    title: SectionFormat
    abstract: SectionFormat
    keywords: SectionFormat
    heading1: SectionFormat
    heading2: SectionFormat
    body: SectionFormat
    references: SectionFormat
    page_margin: Dict[str, float]
    
class FormatSpecParser:
    def __init__(self):
        self.preset_formats = self._load_preset_formats()
    
    def _load_preset_formats(self) -> Dict[str, DocumentFormat]:
        """
        加载预设的格式模板
        """
        preset_path = Path(__file__).parent / "presets"
        formats = {}
        
        if preset_path.exists():
            for format_file in preset_path.glob("*.yaml"):
                with open(format_file, 'r', encoding='utf-8') as f:
                    format_data = yaml.safe_load(f)
                    formats[format_file.stem] = self._parse_format_data(format_data)
        
        return formats
    
    def _parse_format_data(self, data: dict) -> DocumentFormat:
        """
        解析格式数据为DocumentFormat对象
        """
        return DocumentFormat(
            title=SectionFormat(**data['title']),
            abstract=SectionFormat(**data['abstract']),
            keywords=SectionFormat(**data['keywords']),
            heading1=SectionFormat(**data['heading1']),
            heading2=SectionFormat(**data['heading2']),
            body=SectionFormat(**data['body']),
            references=SectionFormat(**data['references']),
            page_margin=data['page_margin']
        )
    
    def parse_user_requirements(self, requirements: str) -> DocumentFormat:
        """
        解析用户提供的格式要求
        使用AI助手理解并标准化用户的格式要求
        """
        from .ai_assistant import DocumentAI
        ai = DocumentAI()
        
        prompt = f"""
        请将以下论文格式要求转换为标准的JSON格式，包含以下字段：
        - title: 标题格式
        - abstract: 摘要格式
        - keywords: 关键词格式
        - heading1: 一级标题格式
        - heading2: 二级标题格式
        - body: 正文格式
        - references: 参考文献格式
        - page_margin: 页边距设置

        每个部分都应包含以下属性：
        - font_size: 字号（磅）
        - font_name: 字体名称
        - bold: 是否加粗（true/false）
        - italic: 是否斜体（true/false）
        - alignment: 对齐方式（LEFT/CENTER/RIGHT/JUSTIFY）
        - first_line_indent: 首行缩进（磅）
        - line_spacing: 行距
        - space_before: 段前距（磅）
        - space_after: 段后距（磅）

        格式要求：
        {requirements}
        """
        
        try:
            response = ai.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "你是一个专业的论文格式规范分析专家。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3
            )
            
            format_data = json.loads(response.choices[0].message.content)
            return self._parse_format_data(format_data)
        except Exception as e:
            print(f"解析格式要求时出错: {str(e)}")
            # 返回默认格式
            return self.get_default_format()
    
    def get_default_format(self) -> DocumentFormat:
        """
        获取默认格式
        """
        return self.preset_formats.get('default', self._get_fallback_format())
    
    def _get_fallback_format(self) -> DocumentFormat:
        """
        获取后备的默认格式
        """
        return DocumentFormat(
            title=SectionFormat(font_size=16, bold=True, alignment="CENTER"),
            abstract=SectionFormat(font_size=12, first_line_indent=24),
            keywords=SectionFormat(font_size=12),
            heading1=SectionFormat(font_size=14, bold=True),
            heading2=SectionFormat(font_size=13, bold=True),
            body=SectionFormat(font_size=12, first_line_indent=24, line_spacing=1.5),
            references=SectionFormat(font_size=10.5, first_line_indent=-24),
            page_margin={"top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25}
        ) 