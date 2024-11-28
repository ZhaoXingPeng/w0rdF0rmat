# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QScrollArea, QLabel,
    QHBoxLayout, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QImage
from pathlib import Path
import tempfile
import os
import win32com.client
import pythoncom
from PIL import Image
import fitz  # PyMuPDF
from src.core.document import Document
from src.core.formatter import WordFormatter
from src.config.config_manager import ConfigManager

class DocumentPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.config_manager = ConfigManager()
        self.temp_dir = tempfile.mkdtemp()  # 创建临时目录
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)  # 移除边距
        
        # 顶部工具栏
        toolbar = QFrame()
        toolbar.setStyleSheet("background-color: #f0f0f0;")
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(10, 5, 10, 5)
        
        self.open_btn = QPushButton('打开文档')
        self.open_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
        """)
        self.open_btn.clicked.connect(self.open_document)
        toolbar_layout.addWidget(self.open_btn)
        toolbar_layout.addStretch()
        
        layout.addWidget(toolbar)
        
        # 文档预览区域
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # 预览容器
        self.preview_container = QWidget()
        self.preview_layout = QVBoxLayout(self.preview_container)
        self.preview_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.scroll_area.setWidget(self.preview_container)
        
        layout.addWidget(self.scroll_area)
    
    def open_document(self):
        """选择文件"""
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
                
                # 显示文档内容
                self.show_document_content(file_path)
                
                # 更新工具栏状态
                self.main_window.update_toolbar_state()
                
                self.main_window.show_message(f"已加载文档: {Path(file_path).name}")
            except Exception as e:
                self.main_window.show_message(f"加载文档失败: {str(e)}", error=True)
    
    def show_document_content(self, docx_path):
        """显示文档内容"""
        try:
            # 清除现有内容
            for i in reversed(range(self.preview_layout.count())): 
                self.preview_layout.itemAt(i).widget().setParent(None)
            
            # 将Word转换为PDF
            pdf_path = os.path.join(self.temp_dir, "temp.pdf")
            self.convert_word_to_pdf(docx_path, pdf_path)
            
            # 使用PyMuPDF渲染PDF页面
            doc = fitz.open(pdf_path)
            for page_num in range(len(doc)):
                page = doc[page_num]
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x缩放以获得更好的质量
                
                # 将页面转换为QImage
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                
                # 创建标签显示页面
                page_label = QLabel()
                page_label.setPixmap(pixmap)
                page_label.setStyleSheet("""
                    QLabel {
                        background-color: white;
                        border: 1px solid #ddd;
                        margin: 10px;
                        padding: 10px;
                    }
                """)
                
                self.preview_layout.addWidget(page_label)
            
            doc.close()
            
        except Exception as e:
            self.main_window.show_message(f"预览失败: {str(e)}", error=True)
    
    def convert_word_to_pdf(self, docx_path, pdf_path):
        """将Word文档转换为PDF"""
        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            doc = word.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 represents PDF format
            doc.Close()
            word.Quit()
            
        finally:
            pythoncom.CoUninitialize()
    
    def cleanup(self):
        """清理临时文件"""
        try:
            for file in os.listdir(self.temp_dir):
                os.remove(os.path.join(self.temp_dir, file))
            os.rmdir(self.temp_dir)
        except Exception as e:
            print(f"清理临时文件失败: {str(e)}")
    
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