# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QComboBox,
    QTabWidget, QFormLayout, QSpinBox,
    QLineEdit, QCheckBox, QFileDialog,
    QDoubleSpinBox, QMessageBox
)
from PyQt6.QtCore import Qt

class FormatPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # 添加说明标签
        help_label = QLabel("请先上传要处理的文档，然后进行格式设置")
        help_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(help_label)
        
        # 上传文档按钮
        self.upload_btn = QPushButton("上传文档")
        self.upload_btn.clicked.connect(self.upload_document)
        layout.addWidget(self.upload_btn)
        
        # 文档状态标签
        self.status_label = QLabel("未上传文档")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        
        # 格式设置部分
        # ... 添加您的格式设置控件 ...
        
        # 确认格式按钮
        self.confirm_btn = QPushButton("确认格式设置")
        self.confirm_btn.clicked.connect(self.confirm_format)
        self.confirm_btn.setEnabled(False)  # 初始禁用
        layout.addWidget(self.confirm_btn)
        
        # 预览按钮
        self.preview_btn = QPushButton("预览文档")
        self.preview_btn.clicked.connect(self.show_preview)
        self.preview_btn.setEnabled(False)  # 初始禁用
        layout.addWidget(self.preview_btn)
        
        self.setLayout(layout)
    
    def upload_document(self):
        """上传文档"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择文档",
            "",
            "Word Documents (*.docx);;All Files (*.*)"
        )
        
        if file_path:
            # 处理文档上传
            try:
                # 这里添加文档处理逻辑
                self.status_label.setText(f"已上传: {file_path}")
                self.main_window.set_document_uploaded(True)
                self.confirm_btn.setEnabled(True)
                QMessageBox.information(self, "成功", "文档上传成功！\n现在您可以进行格式设置了。")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"文档上传失败：{str(e)}")
    
    def confirm_format(self):
        """确认格式设置"""
        # 这里添加格式验证逻辑
        try:
            # 验证格式设置是否完整
            # ...
            
            self.main_window.set_format_configured(True)
            self.preview_btn.setEnabled(True)
            QMessageBox.information(self, "成功", "格式设置已保存！\n您现在可以预览文档了。")
        except Exception as e:
            QMessageBox.warning(self, "警告", f"格式设置有误：{str(e)}")
    
    def show_preview(self):
        """显示预览页面"""
        self.main_window.switch_to_preview_page() 