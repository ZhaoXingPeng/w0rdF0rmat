from docx.shared import Pt, Inches, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.table import _Cell, _Row, _Column
from docx.shared import RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENT
from .format_spec import (
    FormatSpecParser, 
    DocumentFormat, 
    TableCellFormat, 
    TableFormat,
    CaptionFormat,
    ImageFormat,
    PageSetupFormat
)
from docx.oxml import register_element_cls
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement

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
            # 首先设置页面格式
            self.format_page_setup()
            
            # 其他格式化操作
            self.format_title()
            self.format_abstract()
            self.format_keywords()
            self.format_sections()
            self.format_references()
            self.format_tables()
            self.format_images()
            self.format_captions()
            self.format_toc()
        except Exception as e:
            print(f"格式化过程中出错: {str(e)}")
    
    def format_page_setup(self):
        """设置页面格式"""
        try:
            # 获取文档的节
            section = self.document.doc.sections[0]
            
            # 设置页面大小
            section.page_width = Pt(self.format_spec.page_setup.page_width)
            section.page_height = Pt(self.format_spec.page_setup.page_height)
            
            # 设置页边距
            section.top_margin = Pt(self.format_spec.page_setup.margin_top)
            section.bottom_margin = Pt(self.format_spec.page_setup.margin_bottom)
            section.left_margin = Pt(self.format_spec.page_setup.margin_left)
            section.right_margin = Pt(self.format_spec.page_setup.margin_right)
            
            # 设置页眉页脚距离
            section.header_distance = Pt(self.format_spec.page_setup.header_distance)
            section.footer_distance = Pt(self.format_spec.page_setup.footer_distance)
            
            # 设置首页不同
            section.different_first_page_header_footer = self.format_spec.page_setup.different_first_page
            
            # 设置页码
            self._setup_page_numbers(section)
            
            # 设置分栏
            if self.format_spec.page_setup.columns > 1:
                self._setup_columns(section)
            
            # 设置纸张方向
            if self.format_spec.page_setup.orientation == "LANDSCAPE":
                section.orientation = WD_ORIENT.LANDSCAPE
            else:
                section.orientation = WD_ORIENT.PORTRAIT
                
        except Exception as e:
            print(f"设置页面格式时出错: {str(e)}")
    
    def _setup_page_numbers(self, section):
        """设置页码"""
        # 获取页脚段落
        footer = section.footer
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        
        # 清除现有内容
        paragraph.clear()
        
        # 如果不需要在首页显示页码且是首页不同
        if not self.format_spec.page_setup.page_number_show_first and \
           self.format_spec.page_setup.different_first_page:
            return
        
        # 设置页码格式
        page_number_format = {
            "ARABIC": "decimal",
            "ROMAN": "upperRoman",
            "LETTER": "upperLetter"
        }.get(self.format_spec.page_setup.page_number_format, "decimal")
        
        # 添加页码字段
        run = paragraph.add_run()
        fldChar = parse_xml(r'<w:fldChar w:fldCharType="begin"/>')
        run._r.append(fldChar)
        
        instr = parse_xml(f'<w:instrText>PAGE \* {page_number_format}</w:instrText>')
        run._r.append(instr)
        
        fldChar = parse_xml(r'<w:fldChar w:fldCharType="end"/>')
        run._r.append(fldChar)
        
        # 设置页码位置
        position_map = {
            "TOP_LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "TOP_CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "TOP_RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "BOTTOM_LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "BOTTOM_CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "BOTTOM_RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT
        }
        paragraph.alignment = position_map.get(
            self.format_spec.page_setup.page_number_position,
            WD_PARAGRAPH_ALIGNMENT.CENTER
        )
    
    def _setup_columns(self, section):
        """设置分栏"""
        # 获取节属性
        sectPr = section._sectPr
        
        # 清除现有的分栏设置
        for cols in sectPr.xpath('./w:cols'):
            cols.getparent().remove(cols)
        
        # 添加新的分栏设置
        cols = parse_xml(f'''
            <w:cols xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                   w:num="{self.format_spec.page_setup.columns}"
                   w:space="{int(self.format_spec.page_setup.column_spacing * 20)}"/>
        ''')
        sectPr.append(cols)
    
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
            # 首先清除表格的所有样式
            table.style = None
            
            # 清除所有边框
            for row in table.rows:
                for cell in row.cells:
                    for border_position in ['top', 'bottom', 'left', 'right']:
                        self._set_cell_border(cell, border_position, False)
            
            # 添加三条主要横线
            # 顶线
            for cell in table.rows[0].cells:
                self._set_cell_border(cell, "top", True)
            
            # 表头下横线
            for cell in table.rows[0].cells:
                self._set_cell_border(cell, "bottom", True)
            
            # 底线
            for cell in table.rows[-1].cells:
                self._set_cell_border(cell, "bottom", True)
            
            # 设置单元格格式
            for i, row in enumerate(table.rows):
                for cell in row.cells:
                    # 保存原始文本
                    text = cell.text.strip()
                    
                    # 清除现有内容
                    cell.text = ""
                    paragraph = cell.paragraphs[0]
                    
                    # 添加新的文本
                    run = paragraph.add_run(text)
                    
                    # 设置字体格式
                    font = run.font
                    if i == 0:
                        # 使用表头格式
                        font.size = Pt(self.format_spec.tables.header_format.font_size)
                        font.name = self.format_spec.tables.header_format.font_name
                        font.bold = self.format_spec.tables.header_format.bold
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    else:
                        # 使用数据格式
                        font.size = Pt(self.format_spec.tables.data_format.font_size)
                        font.name = self.format_spec.tables.data_format.font_name
                        font.bold = self.format_spec.tables.data_format.bold
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                
        except Exception as e:
            print(f"应用三线表样式时出错: {str(e)}")

    def _apply_grid_style(self, table):
        """应用网格表样式"""
        try:
            # 首先清除表格的所有样式
            table.style = None
            
            # 设置所有单元格的边框
            for row in table.rows:
                for cell in row.cells:
                    for border_position in ['top', 'bottom', 'left', 'right']:
                        self._set_cell_border(cell, border_position, True)
                    
                    # 保存原始文本
                    text = cell.text.strip()
                    
                    # 清除现有内容
                    cell.text = ""
                    paragraph = cell.paragraphs[0]
                    
                    # 添加新的文本
                    run = paragraph.add_run(text)
                    
                    # 设置字体格式
                    font = run.font
                    if row == table.rows[0]:  # 表头行
                        font.size = Pt(self.format_spec.tables.header_format.font_size)
                        font.name = self.format_spec.tables.header_format.font_name
                        font.bold = self.format_spec.tables.header_format.bold
                    else:  # 数据行
                        font.size = Pt(self.format_spec.tables.data_format.font_size)
                        font.name = self.format_spec.tables.data_format.font_name
                        font.bold = self.format_spec.tables.data_format.bold
                    
                    # 设置对齐方式
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    
        except Exception as e:
            print(f"应用网格表样式时出错: {str(e)}")

    def _apply_default_table_style(self, table):
        """应用默认表格样式"""
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.allow_autofit = True
        
        # 设置适当的边框
        self._set_outer_borders(table)
        
        # 设置单元格格式
        for i, row in enumerate(table.rows):
            for cell in row.cells:
                # 保存原始文本
                text = cell.text.strip()
                
                # 清除现有内容
                cell.text = ""
                paragraph = cell.paragraphs[0]
                
                # 添加新的文本
                run = paragraph.add_run(text)
                
                # 设置字体格式
                font = run.font
                if i == 0:  # 表头行
                    font.size = Pt(self.format_spec.tables.header_format.font_size)
                    font.name = self.format_spec.tables.header_format.font_name
                    font.bold = self.format_spec.tables.header_format.bold
                else:  # 数据行
                    font.size = Pt(self.format_spec.tables.data_format.font_size)
                    font.name = self.format_spec.tables.data_format.font_name
                    font.bold = self.format_spec.tables.data_format.bold
                
                # 设置对齐方式
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    def _format_table_cell(self, cell, format_spec: TableCellFormat):
        """
        设置单元格格式
        Args:
            cell: 要格式化的单元格
            format_spec: 单元格格式规范
        """
        # 保存原始文本
        text = cell.text.strip()
        
        # 清除现有内容
        paragraph = cell.paragraphs[0]
        paragraph.clear()
        
        # 添加新的文本
        run = paragraph.add_run(text)
        
        # 设置字体格式
        font = run.font
        font.size = Pt(format_spec.font_size)
        font.name = format_spec.font_name
        font.bold = format_spec.bold
        font.italic = format_spec.italic
        
        # 设置颜色
        if format_spec.text_color:
            font.color.rgb = RGBColor.from_string(format_spec.text_color)
        
        # 设置背景色
        if format_spec.background_color:
            cell._tc.get_or_add_tcPr().append(parse_xml(
                f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                f'w:fill="{format_spec.background_color}"/>'
            ))
        
        # 设置对齐方式
        alignment_map = {
            "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT
        }
        paragraph.alignment = alignment_map.get(format_spec.alignment, WD_PARAGRAPH_ALIGNMENT.LEFT)
        
        # 设置垂直对齐
        v_alignment_map = {
            "TOP": WD_CELL_VERTICAL_ALIGNMENT.TOP,
            "CENTER": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
            "BOTTOM": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        }
        cell.vertical_alignment = v_alignment_map.get(
            format_spec.vertical_alignment, 
            WD_CELL_VERTICAL_ALIGNMENT.CENTER
        )
        
        # 设置行距
        paragraph.paragraph_format.line_spacing = format_spec.line_spacing

    def _apply_table_format(self, table, format_spec: TableFormat):
        """应用表格格式"""
        try:
            # 设置表格整体属性
            table.alignment = WD_TABLE_ALIGNMENT.__members__.get(
                format_spec.alignment, 
                WD_TABLE_ALIGNMENT.CENTER
            )
            table.allow_autofit = format_spec.auto_fit
            
            # 设置表格宽度
            if format_spec.width:
                table.width = Pt(format_spec.width)
            
            # 设置行高和列宽
            for row in table.rows:
                row.height = Pt(format_spec.row_height)
            for column in table.columns:
                column.width = Pt(format_spec.col_width)
            
            # 格式化单元格
            for i, row in enumerate(table.rows):
                for cell in row.cells:
                    if i == 0:  # 表头行
                        self._format_table_cell(cell, format_spec.header_format)
                    else:  # 数据行
                        self._format_table_cell(cell, format_spec.data_format)
            
            # 设置表格间距
            table._element.get_or_add_tblPr().append(parse_xml(
                f'<w:tblspcBefore xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                f'w:w="{int(format_spec.spacing_before * 20)}" w:type="dxa"/>'
            ))
            table._element.get_or_add_tblPr().append(parse_xml(
                f'<w:tblspcAfter xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                f'w:w="{int(format_spec.spacing_after * 20)}" w:type="dxa"/>'
            ))
            
        except Exception as e:
            print(f"应用表格格式时出错: {str(e)}")

    def _clear_table_borders(self, table):
        """清除表格所有边框"""
        for row in table.rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcBorders = tcPr.xpath('./w:tcBorders')
                if tcBorders:
                    for border in tcBorders[0].getchildren():
                        tcBorders[0].remove(border)

    def _set_row_border(self, row, border_position, enabled):
        """设置行的边框"""
        for cell in row.cells:
            self._set_cell_border(cell, border_position, enabled)

    def _set_outer_borders(self, table):
        """设置表格外边框"""
        # 设置第一行和最后一行的上下边框
        for cell in table.rows[0].cells:
            self._set_cell_border(cell, "top", True)
        for cell in table.rows[-1].cells:
            self._set_cell_border(cell, "bottom", True)
        
        # 设置左右边框
        for row in table.rows:
            self._set_cell_border(row.cells[0], "left", True)
            self._set_cell_border(row.cells[-1], "right", True)

    def _set_all_borders(self, table):
        """设置表格所有边框"""
        for row in table.rows:
            for cell in row.cells:
                for border_position in ['top', 'bottom', 'left', 'right']:
                    self._set_cell_border(cell, border_position, True)

    def _set_cell_border(self, cell, border_position, enabled):
        """设置单元格边框"""
        try:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            
            # 获取现有的tcBorders元素，如果不存在则创建
            tcBorders = tcPr.xpath('./w:tcBorders')
            if not tcBorders:
                tcBorders = OxmlElement('w:tcBorders')
                tcBorders.set(nsdecls('w'), '')  # 正确设置命名空间
                tcPr.append(tcBorders)
            else:
                tcBorders = tcBorders[0]
            
            # 移除现有的特定边框设置
            existing = tcBorders.xpath(f'./w:{border_position}')
            if existing:
                for e in existing:
                    tcBorders.remove(e)
            
            # 创建新的边框元素
            border = OxmlElement(f'w:{border_position}')
            border.set(qn('w:val'), 'single' if enabled else 'nil')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), 'auto')
            
            # 添加边框元素
            tcBorders.append(border)
            
        except Exception as e:
            print(f"设置边框失败: {str(e)}")
    
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
    
    def format_captions(self):
        """格式化图表标题"""
        for paragraph in self.document.doc.paragraphs:
            if self._is_figure_caption(paragraph):
                self._format_figure_caption(paragraph)
            elif self._is_table_caption(paragraph):
                self._format_table_caption(paragraph)

    def _is_figure_caption(self, paragraph) -> bool:
        """判断是否为图片标题"""
        text = paragraph.text.strip()
        return text.startswith('图') or text.lower().startswith('fig')

    def _is_table_caption(self, paragraph) -> bool:
        """判断是否为表格标题"""
        text = paragraph.text.strip()
        return text.startswith('表') or text.lower().startswith('table')

    def _format_figure_caption(self, paragraph):
        """格式化图片标题"""
        try:
            # 获取标题文本
            text = paragraph.text.strip()
            
            # 清除现有格式
            paragraph.clear()
            
            # 应用新格式
            run = paragraph.add_run(text)
            font = run.font
            font.size = Pt(self.format_spec.figure_caption.font_size)
            font.name = self.format_spec.figure_caption.font_name
            font.bold = self.format_spec.figure_caption.bold
            
            # 设置对齐方式
            alignment_map = {
                "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
                "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
                "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT
            }
            paragraph.alignment = alignment_map.get(
                self.format_spec.figure_caption.alignment,
                WD_PARAGRAPH_ALIGNMENT.CENTER
            )
            
            # 设置段落间距
            paragraph.paragraph_format.space_before = Pt(self.format_spec.figure_caption.space_before)
            paragraph.paragraph_format.space_after = Pt(self.format_spec.figure_caption.space_after)
            
        except Exception as e:
            print(f"格式化图片标题时出错: {str(e)}")

    def _format_table_caption(self, paragraph):
        """格式化表格标题"""
        try:
            # 获取标题文本
            text = paragraph.text.strip()
            
            # 清除现有格式
            paragraph.clear()
            
            # 应用新格式
            run = paragraph.add_run(text)
            font = run.font
            font.size = Pt(self.format_spec.table_caption.font_size)
            font.name = self.format_spec.table_caption.font_name
            font.bold = self.format_spec.table_caption.bold
            
            # 设置对齐方式
            alignment_map = {
                "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
                "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
                "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT
            }
            paragraph.alignment = alignment_map.get(
                self.format_spec.table_caption.alignment,
                WD_PARAGRAPH_ALIGNMENT.CENTER
            )
            
            # 设置段落间距
            paragraph.paragraph_format.space_before = Pt(self.format_spec.table_caption.space_before)
            paragraph.paragraph_format.space_after = Pt(self.format_spec.table_caption.space_after)
            
        except Exception as e:
            print(f"格式化表格标题时出错: {str(e)}")
    
    def format_toc(self):
        """生成并格式化目录"""
        try:
            if self.format_spec.toc.start_on_new_page:
                self.document.doc.add_page_break()
            
            # 添加目录标题
            title_paragraph = self.document.doc.add_paragraph(self.format_spec.toc.title)
            self._format_toc_title(title_paragraph)
            
            # 收集标题信息
            headings = self._collect_headings()
            
            # 生成目录项
            for level, text, page_number in headings:
                if level <= self.format_spec.toc.include_heading_levels:
                    self._add_toc_entry(level, text, page_number)
        
        except Exception as e:
            print(f"生成目录时出错: {str(e)}")

    def _format_toc_title(self, paragraph):
        """格式化目录标题"""
        # 清除现有格式
        paragraph.clear()
        
        # 添加标题文本
        run = paragraph.add_run(self.format_spec.toc.title)
        font = run.font
        font.size = Pt(self.format_spec.toc.title_font_size)
        font.name = self.format_spec.toc.title_font_name
        font.bold = self.format_spec.toc.title_bold
        
        # 设置对齐方式
        alignment_map = {
            "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT
        }
        paragraph.alignment = alignment_map.get(
            self.format_spec.toc.title_alignment,
            WD_PARAGRAPH_ALIGNMENT.CENTER
        )

    def _collect_headings(self):
        """收集文档中的标题信息"""
        headings = []
        current_page = 1
        
        for paragraph in self.document.doc.paragraphs:
            # 检查是否遇到分页符
            if paragraph._p.xpath('.//w:br[@w:type="page"]'):
                current_page += 1
            
            # 检查段落是否为标题
            if paragraph.style.name.startswith('Heading'):
                level = int(paragraph.style.name[-1])
                headings.append((level, paragraph.text, current_page))
        
        return headings

    def _add_toc_entry(self, level, text, page_number):
        """添加目录项"""
        paragraph = self.document.doc.add_paragraph()
        
        # 设置缩进
        indent = self.format_spec.toc.level1_indent if level == 1 else self.format_spec.toc.level2_indent
        paragraph.paragraph_format.left_indent = Pt(indent)
        
        # 添加标题文本
        run = paragraph.add_run(text)
        
        # 设置格式
        if level == 1:
            font_size = self.format_spec.toc.level1_font_size
            font_name = self.format_spec.toc.level1_font_name
            bold = self.format_spec.toc.level1_bold
            tab_space = self.format_spec.toc.level1_tab_space
        else:
            font_size = self.format_spec.toc.level2_font_size
            font_name = self.format_spec.toc.level2_font_name
            bold = self.format_spec.toc.level2_bold
            tab_space = self.format_spec.toc.level2_tab_space
        
        run.font.size = Pt(font_size)
        run.font.name = font_name
        run.font.bold = bold
        
        # 添加制表符和页码
        if self.format_spec.toc.show_page_numbers:
            paragraph.add_run('\t')
            page_number_run = paragraph.add_run(str(page_number))
            page_number_run.font.size = Pt(font_size)
            page_number_run.font.name = font_name
            
            # 设置制表符
            paragraph.paragraph_format.tab_stops.add_tab_stop(
                Pt(tab_space),
                WD_TAB_ALIGNMENT.RIGHT,
                WD_TAB_LEADER.DOTS
            )
        
        # 设置行距
        paragraph.paragraph_format.line_spacing = self.format_spec.toc.line_spacing
        paragraph.paragraph_format.space_before = Pt(self.format_spec.toc.space_before)
        paragraph.paragraph_format.space_after = Pt(self.format_spec.toc.space_after)
    