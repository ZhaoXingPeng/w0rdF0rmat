from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QScrollArea, QLabel,
    QHBoxLayout, QFrame, QSplitter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QImage
from pathlib import Path
import tempfile
import os
import win32com.client
import pythoncom
import fitz  # PyMuPDF

class PreviewPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.temp_dir = tempfile.mkdtemp()
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 添加标题标签
        title_layout = QHBoxLayout()
        original_label = QLabel("原始文档")
        formatted_label = QLabel("格式化预览")
        original_label.setStyleSheet("font-size: 14px; font-weight: bold; padding: 10px;")
        formatted_label.setStyleSheet("font-size: 14px; font-weight: bold; padding: 10px;")
        title_layout.addWidget(original_label)
        title_layout.addWidget(formatted_label)
        layout.addLayout(title_layout)
        
        # 创建分割视图
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 原始文档视图
        self.original_scroll = QScrollArea()
        self.original_container = QWidget()
        self.original_layout = QVBoxLayout(self.original_container)
        self.original_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.original_scroll.setWidget(self.original_container)
        self.original_scroll.setWidgetResizable(True)
        self.original_scroll.setStyleSheet("""
            QScrollArea {
                border: 1px solid #ddd;
                background-color: #f8f9fa;
            }
        """)
        
        # 格式化后的视图
        self.formatted_scroll = QScrollArea()
        self.formatted_container = QWidget()
        self.formatted_layout = QVBoxLayout(self.formatted_container)
        self.formatted_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.formatted_scroll.setWidget(self.formatted_container)
        self.formatted_scroll.setWidgetResizable(True)
        self.formatted_scroll.setStyleSheet("""
            QScrollArea {
                border: 1px solid #ddd;
                background-color: #f8f9fa;
            }
        """)
        
        splitter.addWidget(self.original_scroll)
        splitter.addWidget(self.formatted_scroll)
        splitter.setSizes([600, 600])  # 设置初始宽度
        
        layout.addWidget(splitter)
        
        # 添加底部按钮
        button_layout = QHBoxLayout()
        self.save_btn = QPushButton('保存文档')
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
        """)
        self.save_btn.clicked.connect(self.save_document)
        button_layout.addStretch()
        button_layout.addWidget(self.save_btn)
        layout.addLayout(button_layout)
    
    def update_preview(self):
        """更新预览内容"""
        if not self.main_window.document:
            return
            
        try:
            # 保存原始文档
            original_docx = os.path.join(self.temp_dir, "original.docx")
            self.main_window.document.doc.save(original_docx)
            
            # 转换原始文档为PDF
            original_pdf = os.path.join(self.temp_dir, "original.pdf")
            self.convert_word_to_pdf(original_docx, original_pdf)
            
            # 显示原始文档
            self.show_pdf_preview(original_pdf, self.original_layout)
            
            # 应用格式化
            formatted_docx = os.path.join(self.temp_dir, "formatted.docx")
            self.main_window.formatter.format()
            self.main_window.document.save(formatted_docx)
            
            # 转换格式化后的文档为PDF
            formatted_pdf = os.path.join(self.temp_dir, "formatted.pdf")
            self.convert_word_to_pdf(formatted_docx, formatted_pdf)
            
            # 显示格式化后的文档
            self.show_pdf_preview(formatted_pdf, self.formatted_layout)
            
        except Exception as e:
            self.main_window.show_message(f"更新预览失败: {str(e)}", error=True)
    
    def show_pdf_preview(self, pdf_path, target_layout):
        """显示PDF预览"""
        # 清除现有内容
        for i in reversed(range(target_layout.count())): 
            widget = target_layout.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)
        
        # 使用PyMuPDF渲染PDF页面
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
            
            # 将页面转换为QImage
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(img)
            
            # 创建页面容器
            page_container = QFrame()
            page_container.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border: 1px solid #ddd;
                    border-radius: 5px;
                    margin: 10px;
                }
            """)
            page_layout = QVBoxLayout(page_container)
            page_layout.setContentsMargins(20, 20, 20, 20)
            
            # 创建标签显示页面
            page_label = QLabel()
            page_label.setPixmap(pixmap)
            page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            page_layout.addWidget(page_label)
            
            # 添加页码
            page_number = QLabel(f"第 {page_num + 1} 页")
            page_number.setAlignment(Qt.AlignmentFlag.AlignCenter)
            page_number.setStyleSheet("color: #666; padding: 5px;")
            page_layout.addWidget(page_number)
            
            target_layout.addWidget(page_container)
        
        doc.close()
    
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
    
    def cleanup(self):
        """清理临时文件"""
        try:
            for file in os.listdir(self.temp_dir):
                os.remove(os.path.join(self.temp_dir, file))
            os.rmdir(self.temp_dir)
        except Exception as e:
            print(f"清理临时文件失败: {str(e)}") 