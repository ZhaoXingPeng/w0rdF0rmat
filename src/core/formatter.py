from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from .format_spec import FormatSpecParser, DocumentFormat

class WordFormatter:
    def __init__(self, document, format_spec: DocumentFormat = None):
        self.document = document
        self.format_parser = FormatSpecParser()
        self.format_spec = format_spec or self.format_parser.get_default_format()
    
    def apply_format_spec(self, format_spec: DocumentFormat):
        """
        应用新的格式规范
        """
        self.format_spec = format_spec
        self.format()
    
    def apply_user_requirements(self, requirements: str):
        """
        应用用户提供的格式要求
        """
        format_spec = self.format_parser.parse_user_requirements(requirements)
        self.apply_format_spec(format_spec)
    
    def _apply_section_format(self, paragraph, section_format):
        """
        应用段落格式
        """
        # 应用字体格式
        for run in paragraph.runs:
            run.font.size = Pt(section_format.font_size)
            run.font.name = section_format.font_name
            run.font.bold = section_format.bold
            run.font.italic = section_format.italic
        
        # 应用段落格式
        paragraph.paragraph_format.first_line_indent = Pt(section_format.first_line_indent)
        paragraph.paragraph_format.line_spacing = section_format.line_spacing
        paragraph.paragraph_format.space_before = Pt(section_format.space_before)
        paragraph.paragraph_format.space_after = Pt(section_format.space_after)
        
        # 设置对齐方式
        alignment_map = {
            "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "JUSTIFY": WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        }
        paragraph.alignment = alignment_map.get(section_format.alignment, WD_PARAGRAPH_ALIGNMENT.LEFT)
    
    def format(self):
        """
        实现文档格式化的主要逻辑
        """
        self.format_title()
        self.format_abstract()
        self.format_keywords()
        self.format_sections()
        self.format_references()
    
    def format_title(self):
        """
        格式化标题
        """
        title = self.document.get_title()
        if title:
            self._apply_section_format(title, self.format_spec.title)
    
    def format_abstract(self):
        """
        格式化摘要
        """
        abstract = self.document.get_abstract()
        if abstract:
            self._apply_section_format(abstract, self.format_spec.abstract)
    
    def format_keywords(self):
        """
        格式化关键词
        """
        keywords = self.document.get_keywords()
        if keywords:
            self._apply_section_format(keywords, self.format_spec.keywords)
    
    def format_sections(self):
        """
        格式化正文章节
        """
        sections = self.document.get_all_sections()
        for section_name, paragraphs in sections.items():
            # 格式化章节标题
            if section_name in self.document.sections:
                section_para = next(p for p in self.document.doc.paragraphs 
                                 if p.text.strip() == section_name)
                self._apply_section_format(section_para, self.format_spec.heading1)
            
            # 格式化章节内容
            for para in paragraphs:
                self._apply_section_format(para, self.format_spec.body)
    
    def format_references(self):
        """
        格式化参考文献
        """
        references = self.document.get_references()
        for ref in references:
            self._apply_section_format(ref, self.format_spec.references)
    
    def format_paragraphs(self):
        """
        段落格式化
        """
        pass
    
    def format_tables(self):
        """
        表格格式化
        """
        pass 