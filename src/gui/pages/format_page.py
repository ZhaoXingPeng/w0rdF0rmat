# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QComboBox,
    QTabWidget, QFormLayout, QSpinBox,
    QLineEdit, QCheckBox, QFileDialog
)
from PyQt6.QtCore import Qt

class FormatPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        
        # 格式选择区域
        format_layout = QHBoxLayout()
        self.format_combo = QComboBox()
        self.format_combo.addItems(['默认格式', '自定义格式', '从文件加载'])
        self.format_combo.currentTextChanged.connect(self.on_format_changed)
        format_layout.addWidget(QLabel('选择格式:'))
        format_layout.addWidget(self.format_combo)
        layout.addLayout(format_layout)
        
        # 创建设置选项卡
        self.tab_widget = QTabWidget()
        
        # 添加各个部分的设置选项卡
        self.add_title_tab()
        self.add_body_tab()
        self.add_table_tab()
        self.add_image_tab()
        
        layout.addWidget(self.tab_widget)
        
        # 添加应用按钮
        self.apply_btn = QPushButton('应用格式')
        self.apply_btn.clicked.connect(self.apply_format)
        layout.addWidget(self.apply_btn)
    
    def add_title_tab(self):
        """添加标题设置选项卡"""
        title_widget = QWidget()
        layout = QFormLayout(title_widget)
        
        # 添加各种设置选项
        self.title_size = QSpinBox()
        self.title_size.setRange(8, 72)
        self.title_size.setValue(16)
        layout.addRow('字号:', self.title_size)
        
        self.title_font = QLineEdit('Times New Roman')
        layout.addRow('字体:', self.title_font)
        
        self.title_bold = QCheckBox('加粗')
        self.title_bold.setChecked(True)
        layout.addRow('样式:', self.title_bold)
        
        self.tab_widget.addTab(title_widget, '标题')
    
    # ... 添加其他选项卡的方法 ...
    
    def on_format_changed(self, text):
        """处理格式选择变化"""
        if text == '从文件加载':
            self.load_format_file()
        self.tab_widget.setEnabled(text == '自定义格式')
    
    def load_format_file(self):
        """加载格式文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择格式文件",
            "",
            "YAML文件 (*.yaml);;JSON文件 (*.json)"
        )
        
        if file_path and self.main_window.formatter:
            try:
                format_spec = self.main_window.formatter.format_parser.parse_format_file(file_path)
                if format_spec:
                    self.main_window.formatter.format_spec = format_spec
                    self.main_window.show_message(f"已加载格式文件")
                else:
                    self.main_window.show_message("格式文件解析失败", error=True)
            except Exception as e:
                self.main_window.show_message(f"加载格式文件失败: {str(e)}", error=True)
    
    def apply_format(self):
        """应用格式设置"""
        if not self.main_window.formatter:
            self.main_window.show_message("请先加载文档", error=True)
            return
        
        try:
            # 应用格式
            self.main_window.formatter.format()
            self.main_window.show_message("格式已应用")
            
            # 切换到预览页面
            self.main_window.stacked_widget.setCurrentWidget(
                self.main_window.preview_page
            )
        except Exception as e:
            self.main_window.show_message(f"应用格式失败: {str(e)}", error=True) 