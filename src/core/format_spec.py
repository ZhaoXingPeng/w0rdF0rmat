from dataclasses import dataclass, field
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
class TableCellFormat:
    """表格单元格格式定义"""
    font_size: float = 10.5
    font_name: str = "Times New Roman"
    bold: bool = False
    italic: bool = False
    alignment: str = "CENTER"
    vertical_alignment: str = "CENTER"
    text_color: str = "000000"
    background_color: Optional[str] = None
    line_spacing: float = 1.0

@dataclass
class TableFormat:
    """表格格式定义"""
    style: str = "DEFAULT"
    alignment: str = "CENTER"
    width: Optional[float] = None
    header_format: TableCellFormat = field(default_factory=lambda: TableCellFormat(
        font_size=10.5,
        font_name="Times New Roman",
        bold=True,
        alignment="CENTER"
    ))
    data_format: TableCellFormat = field(default_factory=lambda: TableCellFormat(
        font_size=10.5,
        font_name="Times New Roman",
        bold=False,
        alignment="LEFT"
    ))
    row_height: float = 12
    col_width: float = 100
    auto_fit: bool = True
    border_size: float = 1.0
    border_color: str = "000000"
    cell_padding: float = 2
    spacing_before: float = 6
    spacing_after: float = 6

@dataclass
class ImageFormat:
    """图片格式定义"""
    width: float = None  # 宽度（磅值），None表示保持原始大小
    height: float = None  # 高度（磅值），None表示保持原始大小
    alignment: str = "CENTER"  # 对齐方式
    caption_font_size: float = 10.5  # 图注字号
    caption_font_name: str = "Times New Roman"  # 图注字体
    caption_alignment: str = "CENTER"  # 图注对齐方式
    space_before: float = 12  # 图片前间距
    space_after: float = 12  # 图片后间距

@dataclass
class CaptionFormat:
    """图表标题格式定义"""
    prefix: str = ""  # 前缀（如"图"或"表"）
    font_size: float = 10.5
    font_name: str = "Times New Roman"
    bold: bool = False
    alignment: str = "CENTER"
    space_before: float = 6
    space_after: float = 6
    numbering_style: str = "ARABIC"  # ARABIC (1,2,3) or CHINESE (一,二,三)
    separator: str = " "  # 编号与标题文本之间的分隔符
    end_mark: str = ""  # 标题末尾的标记（如句号）

@dataclass
class PageSetupFormat:
    """页面设置格式定义"""
    # 页面大小
    page_width: float = 595.0  # A4纸宽度（磅值）
    page_height: float = 842.0  # A4纸高度（磅值）
    
    # 页边距（磅值）
    margin_top: float = 72.0    # 1英寸 = 72磅
    margin_bottom: float = 72.0
    margin_left: float = 90.0   # 约1.25英寸
    margin_right: float = 90.0
    
    # 页眉页脚
    header_distance: float = 36.0  # 页眉距离页面顶部的距离
    footer_distance: float = 36.0  # 页脚距离页面底部的距离
    different_first_page: bool = True  # 首页是否不同
    
    # 页码设置
    page_number_format: str = "ARABIC"  # ARABIC (1,2,3), ROMAN (I,II,III), LETTER (A,B,C)
    page_number_start: int = 1
    page_number_position: str = "BOTTOM_CENTER"  # TOP_CENTER, BOTTOM_RIGHT 等
    page_number_show_first: bool = False  # 首页是否显示页码
    
    # 分栏设置
    columns: int = 1
    column_spacing: float = 36.0  # 栏间距（磅值）
    
    # 纸张方向
    orientation: str = "PORTRAIT"  # PORTRAIT 或 LANDSCAPE

@dataclass
class TOCFormat:
    """目录格式定义"""
    title: str = "目录"  # 目录标题
    title_font_size: float = 14
    title_font_name: str = "Times New Roman"
    title_bold: bool = True
    title_alignment: str = "CENTER"
    
    # 一级目录项格式
    level1_font_size: float = 12
    level1_font_name: str = "Times New Roman"
    level1_bold: bool = False
    level1_indent: float = 0
    level1_tab_space: float = 24  # 标题与页码之间的间距
    
    # 二级目录项格式
    level2_font_size: float = 12
    level2_font_name: str = "Times New Roman"
    level2_bold: bool = False
    level2_indent: float = 24
    level2_tab_space: float = 24
    
    # 目录整体格式
    line_spacing: float = 1.5
    space_before: float = 0
    space_after: float = 12
    show_page_numbers: bool = True
    right_align_page_numbers: bool = True
    include_heading_levels: int = 2  # 包含的标题级别数
    start_on_new_page: bool = True

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
    tables: TableFormat = None
    images: ImageFormat = None
    figure_caption: CaptionFormat = None
    table_caption: CaptionFormat = None
    page_setup: PageSetupFormat = None
    toc: TOCFormat = None

    def __post_init__(self):
        if self.tables is None:
            self.tables = TableFormat()
        if self.images is None:
            self.images = ImageFormat()
        if self.figure_caption is None:
            self.figure_caption = CaptionFormat(prefix="图")
        if self.table_caption is None:
            self.table_caption = CaptionFormat(prefix="表")
        if self.page_setup is None:
            self.page_setup = PageSetupFormat()
        if self.toc is None:
            self.toc = TOCFormat()

class FormatSpecParser:
    def __init__(self):
        self.preset_formats = {}
        self._load_preset_formats()
    
    def _load_preset_formats(self) -> None:
        """
        加载预设的格式模板
        """
        preset_path = Path(__file__).parent / "presets"
        
        if preset_path.exists():
            for format_file in preset_path.glob("*.yaml"):
                try:
                    with open(format_file, 'r', encoding='utf-8') as f:
                        format_data = yaml.safe_load(f)
                        self.preset_formats[format_file.stem] = self._parse_format_data(format_data)
                except Exception as e:
                    print(f"加载预设格式 {format_file.name} 失败: {str(e)}")
                    continue
        
        # 如果没有成功加载任何预设格式，使用后备格式
        if not self.preset_formats:
            self.preset_formats['default'] = self._get_fallback_format()
    
    def parse_format_file(self, file_path: str) -> Optional[DocumentFormat]:
        """
        解析格式文件（支持YAML和JSON）
        """
        try:
            path = Path(file_path)
            with open(path, 'r', encoding='utf-8') as f:
                if path.suffix.lower() == '.yaml':
                    format_data = yaml.safe_load(f)
                else:
                    format_data = json.load(f)
                return self._parse_format_data(format_data)
        except Exception as e:
            print(f"解析格式文件失败: {str(e)}")
            return None
    
    def _parse_format_data(self, data: dict) -> DocumentFormat:
        """解析格式数据为DocumentFormat对象"""
        try:
            # 解析表格格式
            table_data = data.get('tables', {})
            header_format = TableCellFormat(**table_data.get('header_format', {}))
            data_format = TableCellFormat(**table_data.get('data_format', {}))
            table_format = TableFormat(
                header_format=header_format,
                data_format=data_format,
                **{k: v for k, v in table_data.items() if k not in ['header_format', 'data_format']}
            )
            
            return DocumentFormat(
                title=SectionFormat(**data.get('title', {})),
                abstract=SectionFormat(**data.get('abstract', {})),
                keywords=SectionFormat(**data.get('keywords', {})),
                heading1=SectionFormat(**data.get('heading1', {})),
                heading2=SectionFormat(**data.get('heading2', {})),
                body=SectionFormat(**data.get('body', {})),
                references=SectionFormat(**data.get('references', {})),
                page_margin=data.get('page_margin', {
                    "top": 1.0,
                    "bottom": 1.0,
                    "left": 1.25,
                    "right": 1.25
                }),
                tables=table_format
            )
        except Exception as e:
            print(f"解析格式数据失败: {str(e)}")
            return self._get_fallback_format()
    
    def parse_user_requirements(self, requirements: str, config_manager=None) -> DocumentFormat:
        """
        解析用户提供的格式要求
        Args:
            requirements: 户提供的格式要求文本
            config_manager: 配置管理器实例
        Returns:
            解析后的DocumentFormat对象
        """
        if config_manager and config_manager.is_ai_enabled():
            from .ai_assistant import DocumentAI
            ai = DocumentAI(config_manager)
            
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
                result = ai.suggest_formatting("document", requirements)
                if result:
                    return self._parse_format_data(result)
            except Exception as e:
                print(f"AI解析格式要求失败: {str(e)}")
        
        # 如果AI解析失败或未启用，尝试使用简单的规则解析
        try:
            # 这里可以添加简单的规则解析逻辑
            # 暂时返回默认格式
            return self.get_default_format()
        except Exception as e:
            print(f"解析格式要求失败: {str(e)}")
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
            page_margin={"top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25},
            tables=TableFormat(),
            images=ImageFormat(),
            figure_caption=CaptionFormat(prefix="图"),
            table_caption=CaptionFormat(prefix="表"),
            page_setup=PageSetupFormat(),
            toc=TOCFormat()
        ) 
    
    def parse_document_styles(self, document) -> Optional[DocumentFormat]:
        """
        尝试从文档现有样式创建格式规范
        """
        try:
            # 获取文档中使用的样式
            styles = {}
            for para in document.doc.paragraphs:
                if para.style and para.text.strip():
                    style = para.style
                    styles[style.name] = {
                        'font_size': style.font.size.pt if style.font.size else 12,
                        'font_name': style.font.name if style.font.name else "Times New Roman",
                        'bold': style.font.bold if style.font.bold else False,
                        'italic': style.font.italic if style.font.italic else False,
                        'alignment': self._get_alignment_name(para.alignment),
                        'first_line_indent': para.paragraph_format.first_line_indent.pt if para.paragraph_format.first_line_indent else 0,
                        'line_spacing': para.paragraph_format.line_spacing if para.paragraph_format.line_spacing else 1.0,
                        'space_before': para.paragraph_format.space_before.pt if para.paragraph_format.space_before else 0,
                        'space_after': para.paragraph_format.space_after.pt if para.paragraph_format.space_after else 0
                    }
            
            if styles:
                return self._create_format_from_styles(styles)
            return None
            
        except Exception as e:
            print(f"解析文档样式时出错: {str(e)}")
            return None
    
    def _create_format_from_styles(self, styles: Dict) -> DocumentFormat:
        """
        从样式字典创建格式规范
        """
        # 映射样式到文档部分
        title_style = next((s for name, s in styles.items() if 'title' in name.lower()), None)
        abstract_style = next((s for name, s in styles.items() if 'abstract' in name.lower()), None)
        # ... 其他部分类似
        
        return DocumentFormat(
            title=SectionFormat(**(title_style or self._get_fallback_format().title.__dict__)),
            abstract=SectionFormat(**(abstract_style or self._get_fallback_format().abstract.__dict__)),
            # ... 其他部分类似
        )
    
    def _get_alignment_name(self, alignment) -> str:
        """
        将对齐方式转换为字符串
        """
        alignment_map = {
            0: "LEFT",
            1: "CENTER",
            2: "RIGHT",
            3: "JUSTIFY"
        }
        return alignment_map.get(alignment, "LEFT")
    
    # ... 其他方法保持不变 ... 