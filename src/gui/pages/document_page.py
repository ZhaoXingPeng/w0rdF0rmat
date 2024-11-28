# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QTextBrowser
)
from PyQt6.QtCore import Qt
from src.core.document import Document
from src.core.formatter import WordFormatter
from src.config.config_manager import ConfigManager

class DocumentPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.config_manager = ConfigManager()
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        
        # 文件操作按钮
        self.open_btn = QPushButton('打开文档')
        self.open_btn.clicked.connect(self.open_document)
        layout.addWidget(self.open_btn)
        
        # 文档预览（使用QTextBrowser支持富文本显示）
        self.preview = QTextBrowser()
        self.preview.setOpenExternalLinks(False)
        layout.addWidget(self.preview)
    
    def open_document(self):
        """打开文档"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文档",
            "",
            "Word文档 (*.docx)"
        )
        
        if file_path:
            try:
                self.main_window.document = Document(file_path, self.config_manager)
                self.main_window.formatter = WordFormatter(
                    self.main_window.document, 
                    self.config_manager
                )
                
                # 显示文档内容（转换为HTML以保持格式）
                self.show_document_content()
                
                self.main_window.show_message(f"已加载文档")
            except Exception as e:
                self.main_window.show_message(f"加载文档失败: {str(e)}", error=True)
    
    def show_document_content(self):
        """显示文档内容"""
        if not self.main_window.document:
            return
            
        # 将文档内容转换为HTML
        html_content = []
        for para in self.main_window.document.doc.paragraphs:
            # 保持段落格式
            style = para.style
            if style.name.startswith('Heading'):
                html_content.append(f'<h{style.name[-1]}>{para.text}</h{style.name[-1]}>')
            else:
                html_content.append(f'<p>{para.text}</p>')
        
        self.preview.setHtml('\n'.join(html_content)) 