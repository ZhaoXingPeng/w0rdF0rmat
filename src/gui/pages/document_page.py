# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QScrollArea, QLabel,
    QHBoxLayout, QFrame, QSizePolicy
)
from PyQt6.QtCore import Qt, QSize
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
        
        # 创建预览区域容器
        self.preview_container = QWidget()
        self.preview_container.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
            }
        """)
        
        # 创建预览区域布局
        self.preview_layout = QVBoxLayout(self.preview_container)
        self.preview_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.preview_layout.setSpacing(20)  # 设置页面之间的间距
        
        # 创建滚动区域
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidget(self.preview_container)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: #f8f9fa;
            }
        """)
        
        # 添加预览标签
        self.preview_label = QLabel("请打开一个Word文档")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setStyleSheet("""
            QLabel {
                color: #666;
                font-size: 14px;
                padding: 20px;
            }
        """)
        self.preview_layout.addWidget(self.preview_label)
        
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
                self.show_document_content()
                
                # 更新工具栏状态
                self.main_window.update_toolbar_state()
                
                self.main_window.show_message(f"已加载文档: {Path(file_path).name}")
            except Exception as e:
                self.main_window.show_message(f"加载文档失败: {str(e)}", error=True)
    
    def show_document_content(self):
        """显示文档内容"""
        if not self.main_window.document:
            return
            
        try:
            # 清除现有内容
            self.preview_label.hide()  # 隐藏提示标签
            for i in reversed(range(self.preview_layout.count())): 
                widget = self.preview_layout.itemAt(i).widget()
                if widget is not None:
                    widget.setParent(None)
            
            # 将文档保存为临时文件
            temp_docx = os.path.join(self.temp_dir, "temp.docx")
            self.main_window.document.doc.save(temp_docx)
            
            # 将Word转换为PDF
            pdf_path = os.path.join(self.temp_dir, "temp.pdf")
            self.convert_word_to_pdf(temp_docx, pdf_path)
            
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
                        padding: 20px;
                        border-radius: 5px;
                        margin: 10px;
                        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
                    }
                """)
                page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                
                self.preview_layout.addWidget(page_label)
            
            doc.close()
            
        except Exception as e:
            self.main_window.show_message(f"预览失败: {str(e)}", error=True)
            self.preview_label.show()  # 显示错误时显示提示标签
    
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