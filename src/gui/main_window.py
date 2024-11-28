# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QFileDialog, QComboBox,
    QProgressBar, QMessageBox, QTextEdit
)
from PyQt6.QtCore import Qt
from pathlib import Path
from src.core.document import Document
from src.core.formatter import WordFormatter
from src.config.config_manager import ConfigManager
from src.core.format_validator import FormatValidator

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle('Word文档格式化工具')
        self.setMinimumSize(800, 600)
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        layout = QVBoxLayout(central_widget)
        
        # 文件选择区域
        file_layout = QHBoxLayout()
        self.file_label = QLabel('未选择文件')
        self.select_file_btn = QPushButton('选择文件')
        self.select_file_btn.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.select_file_btn)
        layout.addLayout(file_layout)
        
        # 格式选择区域
        format_layout = QHBoxLayout()
        self.format_combo = QComboBox()
        self.format_combo.addItems(['默认格式', '自定义格式'])
        self.load_format_btn = QPushButton('加载格式文件')
        self.load_format_btn.clicked.connect(self.load_format_file)
        format_layout.addWidget(QLabel('格式选择:'))
        format_layout.addWidget(self.format_combo)
        format_layout.addWidget(self.load_format_btn)
        layout.addLayout(format_layout)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        self.format_btn = QPushButton('格式化')
        self.format_btn.clicked.connect(self.format_document)
        self.validate_btn = QPushButton('验证格式')
        self.validate_btn.clicked.connect(self.validate_document)
        button_layout.addWidget(self.format_btn)
        button_layout.addWidget(self.validate_btn)
        layout.addLayout(button_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # 日志区域
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        
        # 初始化状态
        self.document = None
        self.formatter = None
        self.update_ui_state()
    
    def select_file(self):
        """选择文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文档",
            "",
            "Word文档 (*.docx)"
        )
        
        if file_path:
            self.file_label.setText(file_path)
            try:
                self.document = Document(file_path, self.config_manager)
                self.formatter = WordFormatter(self.document, self.config_manager)
                self.log_message(f"成功加载文档: {Path(file_path).name}")
            except Exception as e:
                self.log_message(f"加载文档失败: {str(e)}", error=True)
            
            self.update_ui_state()
    
    def load_format_file(self):
        """加载格式文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择格式文件",
            "",
            "YAML文件 (*.yaml);;JSON文件 (*.json)"
        )
        
        if file_path and self.formatter:
            try:
                format_spec = self.formatter.format_parser.parse_format_file(file_path)
                if format_spec:
                    self.formatter.format_spec = format_spec
                    self.log_message(f"成功加载格式文件: {Path(file_path).name}")
                else:
                    self.log_message("格式文件解析失败", error=True)
            except Exception as e:
                self.log_message(f"加载格式文件失败: {str(e)}", error=True)
    
    def format_document(self):
        """格式化文档"""
        if not self.document or not self.formatter:
            return
        
        try:
            self.progress_bar.setValue(0)
            self.formatter.format()
            self.progress_bar.setValue(50)
            
            # 保存文档
            output_path = Path(self.document.path).parent / f"formatted_{Path(self.document.path).name}"
            self.document.save(str(output_path))
            self.progress_bar.setValue(100)
            
            self.log_message(f"格式化完成，已保存至: {output_path}")
            QMessageBox.information(self, "成功", "文档格式化完成！")
        except Exception as e:
            self.log_message(f"格式化失败: {str(e)}", error=True)
            QMessageBox.critical(self, "错误", f"格式化失败: {str(e)}")
    
    def validate_document(self):
        """验证文档格式"""
        if not self.document or not self.formatter:
            return
        
        try:
            validator = FormatValidator(self.document, self.formatter.format_spec)
            results = validator.validate_all()
            
            # 显示验证结果
            self.log_message("\n=== 格式验证结果 ===")
            for result in results:
                if not result['is_valid']:
                    self.log_message(
                        f"[{result['section']}] {result['element']}: {result['message']}", 
                        error=True
                    )
            
            # 如果没有错误，显示成功消息
            if all(r['is_valid'] for r in results):
                self.log_message("文档格式符合要求！")
                QMessageBox.information(self, "验证结果", "文档格式符合要求！")
            else:
                QMessageBox.warning(self, "验证结果", "发现格式问题，请查看日志。")
        except Exception as e:
            self.log_message(f"验证失败: {str(e)}", error=True)
            QMessageBox.critical(self, "错误", f"验证失败: {str(e)}")
    
    def log_message(self, message: str, error: bool = False):
        """添加日志消息"""
        color = "red" if error else "black"
        self.log_text.append(f'<p style="color: {color};">{message}</p>')
    
    def update_ui_state(self):
        """更新UI状态"""
        has_document = bool(self.document and self.formatter)
        self.format_btn.setEnabled(has_document)
        self.validate_btn.setEnabled(has_document)
        self.load_format_btn.setEnabled(has_document)
        self.format_combo.setEnabled(has_document) 