# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QScrollArea, QLabel,
    QHBoxLayout, QFrame, QApplication
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QImage
from pathlib import Path
import tempfile
import os
import win32com.client
import pythoncom
import fitz  # PyMuPDF
from src.core.document import Document
from src.core.formatter import WordFormatter
from src.config.config_manager import ConfigManager
from src.gui.components.loading_indicator import LoadingIndicator

class DocumentPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.config_manager = ConfigManager()
        self.temp_dir = tempfile.mkdtemp()
        self.last_directory = self.config_manager.get('last_directory', str(Path.home()))  # 获取上次目录
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 创建可点击的文档上传区域
        self.upload_area = QFrame()
        self.upload_area.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 2px dashed #0078d4;
                border-radius: 12px;
                min-height: 300px;
            }
            QFrame:hover {
                background-color: #f0f9ff;
                border-color: #106ebe;
                cursor: pointer;
            }
        """)
        
        # 创建上传区域的布局
        upload_layout = QVBoxLayout(self.upload_area)
        upload_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        upload_layout.setSpacing(20)
        
        # 添加图标（可选）
        icon_label = QLabel()
        icon_label.setPixmap(QPixmap(":/icons/upload.png").scaled(64, 64, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        upload_layout.addWidget(icon_label, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # 添加文字提示
        text_label = QLabel("点击此处打开Word文档")
        text_label.setStyleSheet("""
            QLabel {
                color: #0078d4;
                font-size: 16px;
                font-weight: bold;
            }
        """)
        upload_layout.addWidget(text_label, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # 添加子标题
        sub_text = QLabel("或将文件拖放到此处")
        sub_text.setStyleSheet("""
            QLabel {
                color: #666666;
                font-size: 14px;
            }
        """)
        upload_layout.addWidget(sub_text, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # 创建加载指示器
        self.loading_indicator = LoadingIndicator(self)
        self.loading_indicator.hide()
        upload_layout.addWidget(self.loading_indicator, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # 将上传区域添加到主布局
        layout.addWidget(self.upload_area)
        
        # 为上传区域添加点击事件
        self.upload_area.mousePressEvent = self.open_document
        
        # 设置拖放支持
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            self.process_document(files[0])

    def open_document(self, event=None):
        """选择文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文档",
            self.last_directory,
            "Word文档 (*.docx)"
        )
        
        if file_path:
            self.process_document(file_path)

    def process_document(self, file_path):
        """处理文档"""
        try:
            # 显示加载动画
            self.loading_indicator.show()
            self.loading_indicator.start()
            QApplication.processEvents()
            
            # 保存当前目录
            self.last_directory = str(Path(file_path).parent)
            self.config_manager.set('last_directory', self.last_directory)
            
            # 加载文档
            self.main_window.document = Document(file_path, self.config_manager)
            self.main_window.formatter = WordFormatter(
                self.main_window.document, 
                self.config_manager
            )
            
            # 更新状态
            self.main_window.set_document_uploaded(True)
            self.main_window.update_toolbar_state()
            self.main_window.show_message(f"已加载文档: {Path(file_path).name}")
            
            # 自动切换到格式设置页面
            self.main_window.show_format_page()
            
        except Exception as e:
            self.main_window.show_message(f"加载文档失败: {str(e)}", error=True)
        finally:
            self.loading_indicator.stop()
            self.loading_indicator.hide()
    
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
    
    def handle_document_upload(self, file_path):
        """处理文档上传"""
        try:
            # 处理文档上传的代码...
            
            # 设置文档上传状态
            self.main_window.set_document_uploaded(True)
            
            # 其他代码...
            
        except Exception as e:
            self.main_window.show_message(f"文档上传失败：{str(e)}", error=True) 