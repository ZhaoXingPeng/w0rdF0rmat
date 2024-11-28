from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QSplitter, QLabel,
    QFileDialog, QScrollArea
)
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtGui import QImage, QPixmap
import tempfile
import os
from pathlib import Path
from docx2pdf import convert  # 用于转换docx到pdf
import pythoncom  # 用于COM组件初始化
import fitz  # 用于PDF处理

class PreviewPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.temp_dir = tempfile.mkdtemp()  # 创建临时目录
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        
        # 添加标题标签
        title_layout = QHBoxLayout()
        title_layout.addWidget(QLabel("原始文档"))
        title_layout.addWidget(QLabel("格式化预览"))
        layout.addLayout(title_layout)
        
        # 创建分割视图
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 原始文档视图
        self.original_container = QWidget()
        self.original_layout = QVBoxLayout(self.original_container)
        original_scroll = QScrollArea()
        original_scroll.setWidget(self.original_container)
        original_scroll.setWidgetResizable(True)
        splitter.addWidget(original_scroll)
        
        # 格式化后的视图
        self.formatted_container = QWidget()
        self.formatted_layout = QVBoxLayout(self.formatted_container)
        formatted_scroll = QScrollArea()
        formatted_scroll.setWidget(self.formatted_container)
        formatted_scroll.setWidgetResizable(True)
        splitter.addWidget(formatted_scroll)
        
        layout.addWidget(splitter)
        
        # 添加保存按钮
        self.save_btn = QPushButton('保存文档')
        self.save_btn.clicked.connect(self.save_document)
        layout.addWidget(self.save_btn)
    
    def update_preview(self):
        """更新预览内容"""
        if not self.main_window.document:
            return
            
        try:
            # 显示原始文档
            self.show_original_document()
            
            # 显示格式化后的文档
            self.show_formatted_document()
            
        except Exception as e:
            self.main_window.show_message(f"更新预览失败: {str(e)}", error=True)
    
    def show_original_document(self):
        """显示原始文档"""
        # 清除现有内容
        self.clear_layout(self.original_layout)
        
        # 保存并转换原始文档
        original_path = os.path.join(self.temp_dir, "original.docx")
        self.main_window.document.doc.save(original_path)
        self.convert_and_display(original_path, self.original_layout)
    
    def show_formatted_document(self):
        """显示格式化后的文档"""
        # 清除现有内容
        self.clear_layout(self.formatted_layout)
        
        # 应用格式化并保存
        formatted_path = os.path.join(self.temp_dir, "formatted.docx")
        self.main_window.formatter.format()
        self.main_window.document.save(formatted_path)
        self.convert_and_display(formatted_path, self.formatted_layout)
    
    def convert_and_display(self, docx_path, target_layout):
        """转换并显示文档"""
        try:
            # 转换为PDF
            pdf_path = docx_path.replace('.docx', '.pdf')
            self.convert_word_to_pdf(docx_path, pdf_path)
            
            # 显示页面
            doc = fitz.open(pdf_path)
            for page in doc:
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                
                label = QLabel()
                label.setPixmap(pixmap)
                label.setStyleSheet("""
                    QLabel {
                        background-color: white;
                        border: 1px solid #ddd;
                        padding: 20px;
                        border-radius: 5px;
                        margin: 10px;
                    }
                """)
                target_layout.addWidget(label)
            
            doc.close()
            
        except Exception as e:
            raise Exception(f"转换文档失败: {str(e)}")
    
    def clear_layout(self, layout):
        """清除布局中的所有部件"""
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
    
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