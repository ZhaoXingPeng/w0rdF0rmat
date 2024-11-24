from docx import Document as DocxDocument
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

class Document:
    def __init__(self, path):
        self.path = path
        self.doc = DocxDocument(path)
        # 存储论文各部分内容
        self.title = None
        self.abstract = None
        self.keywords = None
        self.sections = {}
        self._parse_document()
    
    def _parse_document(self):
        """
        解析文档，识别论文各个部分
        """
        current_section = None
        
        for para in self.doc.paragraphs:
            text = para.text.strip()
            
            # 识别标题（通常是文档第一个非空段落）
            if not self.title and text:
                self.title = para
                continue
                
            # 识别摘要部分
            if text.lower().startswith('abstract'):
                self.abstract = para
                continue
                
            # 识别关键词
            if text.lower().startswith('keywords'):
                self.keywords = para
                continue
                
            # 识别章节标题（假设使用数字编号如：1. 2. 等开头）
            if text and any(text.startswith(f"{i}.") for i in range(1, 10)):
                current_section = text
                self.sections[current_section] = []
            elif current_section and text:
                self.sections[current_section].append(para)
    
    def get_title(self):
        """获取论文标题"""
        return self.title
    
    def get_abstract(self):
        """获取摘要部分"""
        return self.abstract
    
    def get_keywords(self):
        """获取关键词部分"""
        return self.keywords
    
    def get_section(self, section_name):
        """获取指定章节的内容"""
        return self.sections.get(section_name, [])
    
    def get_all_sections(self):
        """获取所有章节"""
        return self.sections
    
    def get_references(self):
        """获取参考文献部分"""
        for section_name, paragraphs in self.sections.items():
            if '参考文献' in section_name or 'references' in section_name.lower():
                return paragraphs
        return []
    
    def save(self, output_path):
        """保存文档"""
        self.doc.save(output_path)
    
    def get_paragraphs(self):
        """
        获取所有段落
        """
        return self.doc.paragraphs
    
    def get_tables(self):
        """
        获取所有表格
        """
        return self.doc.tables 