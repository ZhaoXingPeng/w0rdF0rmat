from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QTextBrowser, QSplitter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QTextDocument

class PreviewPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        
        # 创建分割视图
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 原始文档视图
        self.original_view = QTextBrowser()
        self.original_view.setOpenExternalLinks(False)
        
        # 格式化后的视图
        self.formatted_view = QTextBrowser()
        self.formatted_view.setOpenExternalLinks(False)
        
        splitter.addWidget(self.original_view)
        splitter.addWidget(self.formatted_view)
        
        layout.addWidget(splitter)
        
        # 添加保存按钮
        self.save_btn = QPushButton('保存文档')
        self.save_btn.clicked.connect(self.save_document)
        layout.addWidget(self.save_btn)
    
    def update_preview(self):
        """更新预览内容"""
        if not self.main_window.document:
            return
        
        # 显示原始内容
        original_html = self.convert_to_html(self.main_window.document.doc)
        self.original_view.setHtml(original_html)
        
        # 显示格式化后的内容
        formatted_html = self.convert_to_html(self.main_window.document.doc)
        self.formatted_view.setHtml(formatted_html)
    
    def convert_to_html(self, doc) -> str:
        """将文档转换为HTML"""
        html_content = []
        for para in doc.paragraphs:
            style = para.style
            if style.name.startswith('Heading'):
                html_content.append(f'<h{style.name[-1]}>{para.text}</h{style.name[-1]}>')
            else:
                # 保持段落格式
                style_attr = []
                if para.paragraph_format.first_line_indent:
                    indent = para.paragraph_format.first_line_indent.pt
                    style_attr.append(f'text-indent: {indent}pt')
                
                style_str = f' style="{"; ".join(style_attr)}"' if style_attr else ''
                html_content.append(f'<p{style_str}>{para.text}</p>')
        
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
                self.main_window.show_message(f"文档已保存")
            except Exception as e:
                self.main_window.show_message(f"保存失败: {str(e)}", error=True) 