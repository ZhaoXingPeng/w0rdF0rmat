from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QSplitter, QLabel,
    QFileDialog
)
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtWebEngineWidgets import QWebEngineView
import tempfile
import os
from pathlib import Path
from docx2pdf import convert  # 用于转换docx到pdf
import pythoncom  # 用于COM组件初始化

class PreviewPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.temp_dir = tempfile.mkdtemp()  # 创建临时目录
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
        self.original_view = QWebEngineView()
        
        # 格式化后的视图
        self.formatted_view = QWebEngineView()
        
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
            
        try:
            # 初始化COM组件（用于Word转PDF）
            pythoncom.CoInitialize()
            
            # 保存原始文档到临时文件
            original_docx = os.path.join(self.temp_dir, "original.docx")
            self.main_window.document.doc.save(original_docx)
            
            # 转换为PDF以便预览
            original_pdf = os.path.join(self.temp_dir, "original.pdf")
            convert(original_docx, original_pdf)
            
            # 显示原始文档
            self.original_view.setUrl(QUrl.fromLocalFile(original_pdf))
            
            # 保存格式化后的文档到临时文件
            formatted_docx = os.path.join(self.temp_dir, "formatted.docx")
            self.main_window.document.save(formatted_docx)
            
            # 转换为PDF以便预览
            formatted_pdf = os.path.join(self.temp_dir, "formatted.pdf")
            convert(formatted_docx, formatted_pdf)
            
            # 显示格式化后的文档
            self.formatted_view.setUrl(QUrl.fromLocalFile(formatted_pdf))
            
        except Exception as e:
            self.main_window.show_message(f"预览更新失败: {str(e)}", error=True)
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