from docx.shared import Pt, Inches, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.table import _Cell, _Row, _Column
from docx.shared import RGBColor
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENT

# 导入格式相关的类
from .format_spec import (
    DocumentFormat,
    SectionFormat,
    TableFormat,
    TableCellFormat,
    ImageFormat,
    CaptionFormat,
    FormatSpecParser
)

# 定义制表符相关的常量
class WD_TAB_ALIGNMENT:
    LEFT = 0
    CENTER = 1
    RIGHT = 2

class WD_TAB_LEADER:
    SPACES = 0
    DOTS = 1
    DASHES = 2
    LINES = 3
    HEAVY = 4
    MIDDLE_DOT = 5

class WordFormatter:
    def __init__(self, document, config_manager):
        self.document = document
        self.config_manager = config_manager
        self.format_spec = {}  # 初始化格式规范
    
    def set_format_spec(self, format_spec):
        """设置格式规范"""
        self.format_spec = format_spec
    
    def format(self):
        """应用格式到文档"""
        try:
            doc = self.document.doc
            
            # 应用段落格式
            if 'paragraph' in self.format_spec:
                para_spec = self.format_spec['paragraph']
                for paragraph in doc.paragraphs:
                    # 设置段落间距
                    if 'before_spacing' in para_spec:
                        paragraph.paragraph_format.space_before = Pt(para_spec['before_spacing'])
                    if 'after_spacing' in para_spec:
                        paragraph.paragraph_format.space_after = Pt(para_spec['after_spacing'])
                    if 'line_spacing' in para_spec:
                        paragraph.paragraph_format.line_spacing = para_spec['line_spacing']
                    
                    # 设置对齐方式
                    if 'alignment' in para_spec:
                        alignment_map = {
                            "左对齐": WD_PARAGRAPH_ALIGNMENT.LEFT,
                            "居中": WD_PARAGRAPH_ALIGNMENT.CENTER,
                            "右对齐": WD_PARAGRAPH_ALIGNMENT.RIGHT,
                            "两端对齐": WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                        }
                        paragraph.alignment = alignment_map.get(para_spec['alignment'], WD_PARAGRAPH_ALIGNMENT.LEFT)
            
            # 应用字体格式
            if 'font' in self.format_spec:
                font_spec = self.format_spec['font']
                for paragraph in doc.paragraphs:
                    for run in paragraph.runs:
                        if 'name' in font_spec:
                            run.font.name = font_spec['name']
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_spec['name'])
                        if 'size' in font_spec:
                            run.font.size = Pt(font_spec['size'])
            
        except Exception as e:
            print(f"格式化失败: {str(e)}")
            raise
    