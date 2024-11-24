from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

class WordFormatter:
    def __init__(self, document):
        self.document = document
    
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
        - 居中对齐
        - 字体大小：16pt
        - 加粗
        """
        title = self.document.get_title()
        if title:
            title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = title.runs[0] if title.runs else title.add_run(title.text)
            run.font.size = Pt(16)
            run.font.bold = True
    
    def format_abstract(self):
        """
        格式化摘要
        - 字体大小：12pt
        - 首行缩进：2字符
        """
        abstract = self.document.get_abstract()
        if abstract:
            for run in abstract.runs:
                run.font.size = Pt(12)
            abstract.paragraph_format.first_line_indent = Pt(24)
    
    def format_sections(self):
        """
        格式化正文章节
        - 字体大小：12pt
        - 首行缩进：2字符
        - 1.5倍行距
        """
        sections = self.document.get_all_sections()
        for section_name, paragraphs in sections.items():
            # 格式化章节标题
            if section_name in self.document.sections:
                section_para = next(p for p in self.document.doc.paragraphs 
                                 if p.text.strip() == section_name)
                section_para.runs[0].font.size = Pt(14)
                section_para.runs[0].font.bold = True
            
            # 格式化章节内容
            for para in paragraphs:
                for run in para.runs:
                    run.font.size = Pt(12)
                para.paragraph_format.first_line_indent = Pt(24)
                para.paragraph_format.line_spacing = 1.5
    
    def format_references(self):
        """
        格式化参考文献
        - 悬挂缩进
        - 字体大小：10.5pt
        """
        references = self.document.get_references()
        for ref in references:
            for run in ref.runs:
                run.font.size = Pt(10.5)
            ref.paragraph_format.first_line_indent = Pt(-24)
            ref.paragraph_format.left_indent = Pt(24)
    
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