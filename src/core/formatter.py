from docx.shared import Pt, Inches, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.table import _Cell, _Row, _Column
from docx.shared import RGBColor
from docx.oxml import parse_xml
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
        if not paragraph:
            return
        
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
            self.format_tables()
            self.format_images()
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
                section_para = next((p for p in self.document.doc.paragraphs 
                                 if p.text.strip() == section_name), None)
                if section_para:
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
        """格式化文档中的所有表格"""
        tables = self.document.get_tables()
        for table in tables:
            table_style = self._detect_table_style(table)
            if table_style == "THREE_LINE":
                self._apply_three_line_style(table)
            elif table_style == "GRID":
                self._apply_grid_style(table)
            else:
                self._apply_default_table_style(table)

    def _detect_table_style(self, table) -> str:
        """
        检测表格应该使用的样式
        根据表格的特征（行数、列数、内容等）来判断
        """
        # 检测是否为三线表（通常用于变量说明、数据分析等）
        if (len(table.rows) > 1 and 
            any("变量" in cell.text or "指标" in cell.text 
                for cell in table.rows[0].cells)):
            return "THREE_LINE"
        
        # 检测是否为网格表（通常用于复杂数据展示）
        if len(table.rows) > 3 and len(table.columns) > 3:
            return "GRID"
        
        return "DEFAULT"

    def _apply_three_line_style(self, table):
        """应用三线表样式"""
        try:
            # 设置表格整体格式
            table.style = 'Table Grid'
            table.autofit = True
            
            # 清除所有边框
            for row in table.rows:
                for cell in row.cells:
                    # 使用python-docx提供的方法设置边框
                    cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'))
            
            # 添加三条主要横线
            # 顶线
            for cell in table.rows[0].cells:
                cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:top w:val="single" w:sz="12"/></w:tcBorders>'))
            
            # 表头下横线
            for cell in table.rows[0].cells:
                cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bottom w:val="single" w:sz="12"/></w:tcBorders>'))
            
            # 底线
            for cell in table.rows[-1].cells:
                cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bottom w:val="single" w:sz="12"/></w:tcBorders>'))
            
            # 设置单元格格式
            for i, row in enumerate(table.rows):
                for cell in row.cells:
                    # 保存原始文本
                    text = cell.text.strip()
                    
                    # ���除现有内容
                    for paragraph in cell.paragraphs:
                        paragraph.clear()
                    
                    # 创建新的段落和运行
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run(text)
                    
                    # 设置字体格式
                    font = run.font
                    font.size = Pt(self.format_spec.tables.font_size)
                    font.name = self.format_spec.tables.font_name
                    
                    # 表头加粗
                    if i == 0:
                        font.bold = True
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    else:
                        font.bold = False
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    
        except Exception as e:
            print(f"应用三线表样式时出错: {str(e)}")

    def _apply_grid_style(self, table):
        """应用网格表样式"""
        # 设置表格整体格式
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.allow_autofit = True
        
        # 设置所有边框
        self._set_all_borders(table)
        
        # 设置单元格格式
        for i, row in enumerate(table.rows):
            for cell in row.cells:
                self._format_table_cell(cell,
                    bold=(i == 0),  # 表头加粗
                    font_size=self.format_spec.tables.font_size,
                    font_name=self.format_spec.tables.font_name,
                    alignment="CENTER"
                )

    def _apply_default_table_style(self, table):
        """应用默认表格样式"""
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.allow_autofit = True
        
        # 设置适当的边框
        self._set_outer_borders(table)
        
        # 设置单元格格式
        for i, row in enumerate(table.rows):
            for cell in row.cells:
                self._format_table_cell(cell,
                    bold=(i == 0),
                    font_size=self.format_spec.tables.font_size,
                    font_name=self.format_spec.tables.font_name
                )

    def _format_table_cell(self, cell, bold=False, font_size=10.5, 
                          font_name="Times New Roman", alignment="LEFT"):
        """设置单元格格式"""
        # 清除现有内容并重新创建
        paragraph = cell.paragraphs[0]
        text = paragraph.text
        paragraph.clear()
        run = paragraph.add_run(text)
        
        # 设置字体格式
        run.font.size = Pt(font_size)
        run.font.name = font_name
        run.font.bold = bold
        
        # 设置对齐方式
        if alignment in WD_PARAGRAPH_ALIGNMENT.__members__:
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.__members__[alignment]

    def _clear_table_borders(self, table):
        """清除表格所有边框"""
        for row in table.rows:
            for cell in row.cells:
                for border in cell._tc.get_or_add_tcPr().xpath('./w:tcBorders/*'):
                    border.getparent().remove(border)

    def _set_row_border(self, row, border_position, enabled):
        """设置行的边框"""
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = tcPr.get_or_add_tcBorders()
            border = getattr(tcBorders, f"get_or_add_{border_position}")()
            border.val = 'single' if enabled else 'nil'

    def _set_outer_borders(self, table):
        """设置表格外边框"""
        for row in table.rows:
            self._set_row_border(row, "left", True)
            self._set_row_border(row, "right", True)
        self._set_row_border(table.rows[0], "top", True)
        self._set_row_border(table.rows[-1], "bottom", True)

    def _set_all_borders(self, table):
        """设置表格所有边框"""
        for row in table.rows:
            for cell in row.cells:
                for border_position in ['top', 'bottom', 'left', 'right']:
                    self._set_cell_border(cell, border_position, True)

    def _set_cell_border(self, cell, border_position, enabled):
        """设置单元格边框"""
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = tcPr.get_or_add_tcBorders()
        border = getattr(tcBorders, f"get_or_add_{border_position}")()
        border.val = 'single' if enabled else 'nil'
    
    def format_images(self):
        """格式化文档中的所有图片"""
        for paragraph in self.document.doc.paragraphs:
            for run in paragraph.runs:
                if run._r.xml.find('.//w:drawing') != -1:  # 检查是否包含图片
                    self._format_image_paragraph(paragraph)
                    # 检查下一段是否为图注
                    next_para = self._get_next_paragraph(paragraph)
                    if next_para and self._is_image_caption(next_para):
                        self._format_image_caption(next_para)

    def _format_image_paragraph(self, paragraph):
        """格式化包含图片的段落"""
        # 设置图片段落的对齐方式
        alignment_map = {
            "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT
        }
        paragraph.alignment = alignment_map.get(
            self.format_spec.images.alignment, 
            WD_PARAGRAPH_ALIGNMENT.CENTER
        )
        
        # 设置段落间距
        paragraph.paragraph_format.space_before = Pt(self.format_spec.images.space_before)
        paragraph.paragraph_format.space_after = Pt(self.format_spec.images.space_after)
        
        # 调整图片大小（如果指定了大小）
        for run in paragraph.runs:
            if run._r.xml.find('.//w:drawing') != -1:
                drawing = run._r.xpath('.//w:drawing')[0]
                if self.format_spec.images.width:
                    # 设置图片宽度
                    extent = drawing.xpath('.//wp:extent')[0]
                    extent.set('cx', str(int(self.format_spec.images.width * 360000)))
                if self.format_spec.images.height:
                    # 设置图片高度
                    extent = drawing.xpath('.//wp:extent')[0]
                    extent.set('cy', str(int(self.format_spec.images.height * 360000)))

    def _format_image_caption(self, paragraph):
        """格式化图注"""
        # 设置图注格式
        alignment_map = {
            "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT
        }
        paragraph.alignment = alignment_map.get(
            self.format_spec.images.caption_alignment,
            WD_PARAGRAPH_ALIGNMENT.CENTER
        )
        
        for run in paragraph.runs:
            run.font.size = Pt(self.format_spec.images.caption_font_size)
            run.font.name = self.format_spec.images.caption_font_name

    def _get_next_paragraph(self, paragraph):
        """获取下一个段落"""
        paragraphs = list(self.document.doc.paragraphs)
        try:
            current_index = paragraphs.index(paragraph)
            if current_index < len(paragraphs) - 1:
                return paragraphs[current_index + 1]
        except ValueError:
            pass
        return None

    def _is_image_caption(self, paragraph):
        """判断段落是否为图注"""
        text = paragraph.text.strip().lower()
        return text.startswith('图') or text.startswith('fig')
    