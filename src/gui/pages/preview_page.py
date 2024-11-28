from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QTextBrowser, QSplitter,
    QLabel
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QTextDocument, QFont
from docx.shared import Pt

class PreviewPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        
        # 添加标签
        labels_layout = QHBoxLayout()
        labels_layout.addWidget(QLabel("原始文档"))
        labels_layout.addWidget(QLabel("格式化预览"))
        layout.addLayout(labels_layout)
        
        # 创建分割视图
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 原始文档视图
        self.original_view = QTextBrowser()
        self.original_view.setOpenExternalLinks(False)
        self.original_view.setFont(QFont("Times New Roman", 12))
        
        # 格式化后的视图
        self.formatted_view = QTextBrowser()
        self.formatted_view.setOpenExternalLinks(False)
        self.formatted_view.setFont(QFont("Times New Roman", 12))
        
        splitter.addWidget(self.original_view)
        splitter.addWidget(self.formatted_view)
        splitter.setSizes([600, 600])  # 设置初始宽度
        
        layout.addWidget(splitter)
        
        # 添加按钮区域
        button_layout = QHBoxLayout()
        
        self.save_btn = QPushButton('保存文档')
        self.save_btn.clicked.connect(self.save_document)
        button_layout.addWidget(self.save_btn)
        
        layout.addLayout(button_layout)
    
    def update_preview(self):
        """更新预览内容"""
        if not self.main_window.document:
            return
        
        # 显示原始内容
        original_html = self._convert_to_html(self.main_window.document.doc, formatted=False)
        self.original_view.setHtml(original_html)
        
        # 显示格式化后的内容
        formatted_html = self._convert_to_html(self.main_window.document.doc, formatted=True)
        self.formatted_view.setHtml(formatted_html)
    
    def _convert_to_html(self, doc, formatted=False) -> str:
        """将文档转换为HTML"""
        html_content = ['<html><body>']
        
        for para in doc.paragraphs:
            # 获取段落样式
            style = para.style
            style_attr = []
            
            if formatted:
                # 应用格式化后的样式
                if style.name.startswith('Heading'):
                    font_size = 14 if style.name == 'Heading 1' else 13
                    style_attr.extend([
                        f'font-size: {font_size}pt',
                        'font-weight: bold',
                        'margin-top: 12pt',
                        'margin-bottom: 12pt'
                    ])
                else:
                    # 正文样式
                    style_attr.extend([
                        'font-size: 12pt',
                        'text-indent: 24pt',
                        'line-height: 1.5',
                        'text-align: justify'
                    ])
            else:
                # 保持原始样式
                if para.paragraph_format.first_line_indent:
                    indent = para.paragraph_format.first_line_indent.pt
                    style_attr.append(f'text-indent: {indent}pt')
                if para.paragraph_format.line_spacing:
                    style_attr.append(f'line-height: {para.paragraph_format.line_spacing}')
            
            # 创建样式字符串
            style_str = f' style="{"; ".join(style_attr)}"' if style_attr else ''
            
            # 添加段落
            if style.name.startswith('Heading'):
                level = style.name[-1]
                html_content.append(f'<h{level}{style_str}>{para.text}</h{level}>')
            else:
                html_content.append(f'<p{style_str}>{para.text}</p>')
        
        html_content.append('</body></html>')
        return '\n'.join(html_content)
    
    def save_document(self):
        """保存文档"""
        if not self.main_window.document:
            return
            
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存文档",
            "",
            "Word文档 (*.docx)"
        )
        
        if file_path:
            try:
                self.main_window.document.save(file_path)
                self.main_window.show_message(f"文档已保存至: {file_path}")
            except Exception as e:
                self.main_window.show_message(f"保存失败: {str(e)}", error=True) 