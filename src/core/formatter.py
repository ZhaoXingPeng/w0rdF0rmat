from docx.shared import Pt, Inches, Cm, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn

class WordFormatter:
    def __init__(self, document, config_manager):
        self.document = document
        self.config_manager = config_manager
        self.format_spec = {}

    def set_format_spec(self, format_spec):
        """设置格式规范"""
        self.format_spec = format_spec
        print("格式规范已更新:", format_spec)

    def format(self):
        """应用格式到文档"""
        try:
            doc = self.document.doc
            print("开始应用格式...")

            # 遍历所有段落
            for paragraph in doc.paragraphs:
                # 根据段落内容或样式应用不同的格式
                if "摘要" in paragraph.text:
                    self._apply_abstract_format(paragraph)
                elif paragraph.style.name.startswith('Heading'):
                    self._apply_heading_format(paragraph)
                else:
                    self._apply_body_format(paragraph)

            print("格式应用完成")
            return True

        except Exception as e:
            print(f"格式化失败: {str(e)}")
            raise

    def _apply_abstract_format(self, paragraph):
        """应用摘要格式"""
        spec = self.format_spec.get('abstract', {})
        title_spec = spec.get('title', {})
        content_spec = spec.get('content', {})
        margin_spec = spec.get('margin', {})
        
        if "摘要" in paragraph.text:
            # 应用标题格式
            self._apply_font_format(paragraph, title_spec)
            # 应用对齐方式
            align_map = {
                "居中": WD_PARAGRAPH_ALIGNMENT.CENTER,
                "左对齐": WD_PARAGRAPH_ALIGNMENT.LEFT,
                "右对齐": WD_PARAGRAPH_ALIGNMENT.RIGHT
            }
            if 'align' in title_spec:
                paragraph.alignment = align_map.get(title_spec['align'], WD_PARAGRAPH_ALIGNMENT.CENTER)
        else:
            # 应用正文格式
            self._apply_font_format(paragraph, content_spec)
            # 应用行间距
            if 'line_spacing' in content_spec:
                paragraph.paragraph_format.line_spacing = content_spec['line_spacing']
            # 应用段落间距
            if 'para_spacing' in content_spec:
                paragraph.paragraph_format.space_after = Pt(content_spec['para_spacing'])
            # 应用首行缩进
            if 'first_line_indent' in content_spec:
                paragraph.paragraph_format.first_line_indent = Pt(content_spec['first_line_indent'] * content_spec.get('size', 12))
            # 应用对齐方式
            align_map = {
                "两端对齐": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
                "左对齐": WD_PARAGRAPH_ALIGNMENT.LEFT,
                "右对齐": WD_PARAGRAPH_ALIGNMENT.RIGHT,
                "居中": WD_PARAGRAPH_ALIGNMENT.CENTER
            }
            if 'align' in content_spec:
                paragraph.alignment = align_map.get(content_spec['align'], WD_PARAGRAPH_ALIGNMENT.JUSTIFY)
        
        # 应用页边距
        section = paragraph.part.document.sections[0]
        if margin_spec:
            if 'top' in margin_spec:
                section.top_margin = Mm(margin_spec['top'])
            if 'bottom' in margin_spec:
                section.bottom_margin = Mm(margin_spec['bottom'])
            if 'left' in margin_spec:
                section.left_margin = Mm(margin_spec['left'])
            if 'right' in margin_spec:
                section.right_margin = Mm(margin_spec['right'])

    def _apply_heading_format(self, paragraph):
        """应用标题格式"""
        spec = self.format_spec.get('main_text', {}).get('chapter', {})
        self._apply_font_format(paragraph, spec)

    def _apply_body_format(self, paragraph):
        """应用正文格式"""
        spec = self.format_spec.get('main_text', {}).get('body', {})
        self._apply_font_format(paragraph, spec)
        if 'line_spacing' in spec:
            paragraph.paragraph_format.line_spacing = spec['line_spacing']

    def _apply_font_format(self, paragraph, spec):
        """应用字体格式"""
        if not spec:
            return
            
        for run in paragraph.runs:
            if 'font' in spec:
                run.font.name = spec['font']
                # 设置中文字体
                run._element.rPr.rFonts.set(qn('w:eastAsia'), spec['font'])
            if 'size' in spec:
                run.font.size = Pt(spec['size'])
    