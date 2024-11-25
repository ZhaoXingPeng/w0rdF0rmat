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
class TableFormat:
    """表格格式定义"""
    style: str = "DEFAULT"  # 表格样式：DEFAULT, THREE_LINE, GRID
    font_size: float = 10.5
    font_name: str = "Times New Roman"
    alignment: str = "CENTER"
    header_bold: bool = True
    row_height: float = 12  # 行高（磅值）
    col_width: float = 100  # 默认列宽（磅值）
    cell_padding: float = 2  # 单元格内边距（磅值）
    line_spacing: float = 1.0  # 单元格内行距
    borders: Dict[str, bool] = None  # 边框设置

    def __post_init__(self):
        if self.borders is None:
            self.borders = {
                "top": True,
                "bottom": True,
                "left": True,
                "right": True,
                "inside_h": True,
                "inside_v": True
            }

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

    def __post_init__(self):
        if self.tables is None:
            self.tables = TableFormat()

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
        """
        解析格式数据为DocumentFormat对象
        """
        try:
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
                tables=TableFormat(**data.get('tables', {}))
            )
        except Exception as e:
            print(f"解析格式数据失败: {str(e)}")
            return self._get_fallback_format()
    
    def parse_user_requirements(self, requirements: str, config_manager=None) -> DocumentFormat:
        """
        解析用户提供的格式要求
        Args:
            requirements: 用户提供的格式要求文本
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
            tables=TableFormat()
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