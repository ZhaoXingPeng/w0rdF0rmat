from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from .format_spec import FormatSpecParser, DocumentFormat

class WordFormatter:
    def __init__(self, document, config_manager):
        self.document = document
        self.config_manager = config_manager
        self.format_parser = FormatSpecParser()
        
        # 获取格式规范
        template_path = self.config_manager.get_template_path()
        self.format_spec = self.format_parser.parse_format_file(template_path)
        
        if not self.format_spec:
            # 如果无法加载模板，使用默认格式
            self.format_spec = self.format_parser.get_default_format()
    
    def apply_format_spec(self, format_spec: DocumentFormat):
        """
        应用新的格式规范
        """
        self.format_spec = format_spec
        self.format()
    
    def apply_user_requirements(self, requirements: str):
        """
        应用用户提供的格式要求
        如果没有提供要求，则使用本地样式进行格式化
        """
        if requirements.strip():
            format_spec = self.format_parser.parse_user_requirements(
                requirements, 
                self.config_manager
            )
            self.apply_format_spec(format_spec)
        else:
            # 如果没有用户要求，直接使用本地样式
            doc_format = self.format_parser.parse_document_styles(self.document)
            if doc_format:
                self.apply_format_spec(doc_format)
            else:
                # 使用默认格式
                self.apply_format_spec(self.format_parser.get_default_format())
    
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
        try:
            self.format_title()
            self.format_abstract()
            self.format_keywords()
            self.format_sections()
            self.format_references()
            self.format_tables()  # 添加对表格的格式化支持
        except Exception as e:
            print(f"格式化过程中出错: {str(e)}")
    
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
    
    def format_tables(self):
        """
        格式化表格
        """
        tables = self.document.get_tables()
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._apply_section_format(paragraph, self.format_spec.body)  # 假设表格内容使用正文格式
    