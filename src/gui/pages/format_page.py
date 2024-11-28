# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QComboBox,
    QTabWidget, QFormLayout, QSpinBox,
    QLineEdit, QCheckBox, QFileDialog,
    QDoubleSpinBox
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
        self.add_page_tab()
        
        layout.addWidget(self.tab_widget)
        
        # 添加应用按钮
        self.apply_btn = QPushButton('应用格式')
        self.apply_btn.clicked.connect(self.apply_format)
        layout.addWidget(self.apply_btn)
    
    def add_title_tab(self):
        """添加标题设置选项卡"""
        title_widget = QWidget()
        layout = QFormLayout(title_widget)
        
        # 字号设置
        self.title_size = QSpinBox()
        self.title_size.setRange(8, 72)
        self.title_size.setValue(16)
        layout.addRow('字号:', self.title_size)
        
        # 字体设置
        self.title_font = QLineEdit('Times New Roman')
        layout.addRow('字体:', self.title_font)
        
        # 样式设置
        self.title_bold = QCheckBox('加粗')
        self.title_bold.setChecked(True)
        self.title_italic = QCheckBox('斜体')
        style_layout = QHBoxLayout()
        style_layout.addWidget(self.title_bold)
        style_layout.addWidget(self.title_italic)
        layout.addRow('样式:', style_layout)
        
        # 对齐方式
        self.title_alignment = QComboBox()
        self.title_alignment.addItems(['左对齐', '居中', '右对齐'])
        self.title_alignment.setCurrentText('居中')
        layout.addRow('对齐:', self.title_alignment)
        
        self.tab_widget.addTab(title_widget, '标题')
    
    def add_body_tab(self):
        """添加正文设置选项卡"""
        body_widget = QWidget()
        layout = QFormLayout(body_widget)
        
        # 字号设置
        self.body_size = QSpinBox()
        self.body_size.setRange(8, 72)
        self.body_size.setValue(12)
        layout.addRow('字号:', self.body_size)
        
        # 字体设置
        self.body_font = QLineEdit('Times New Roman')
        layout.addRow('字体:', self.body_font)
        
        # 行距设置
        self.body_line_spacing = QDoubleSpinBox()
        self.body_line_spacing.setRange(1.0, 3.0)
        self.body_line_spacing.setValue(1.5)
        self.body_line_spacing.setSingleStep(0.1)
        layout.addRow('行距:', self.body_line_spacing)
        
        # 首行缩进
        self.body_indent = QSpinBox()
        self.body_indent.setRange(0, 72)
        self.body_indent.setValue(24)
        layout.addRow('首行缩进:', self.body_indent)
        
        self.tab_widget.addTab(body_widget, '正文')
    
    def add_table_tab(self):
        """添加表格设置选项卡"""
        table_widget = QWidget()
        layout = QFormLayout(table_widget)
        
        # 表格样式
        self.table_style = QComboBox()
        self.table_style.addItems(['默认样式', '三线表', '网格表'])
        layout.addRow('表格样式:', self.table_style)
        
        # 字号设置
        self.table_size = QSpinBox()
        self.table_size.setRange(8, 72)
        self.table_size.setValue(10.5)
        layout.addRow('字号:', self.table_size)
        
        # 表头设置
        self.table_header_bold = QCheckBox('表头加粗')
        self.table_header_bold.setChecked(True)
        layout.addRow('表头样式:', self.table_header_bold)
        
        self.tab_widget.addTab(table_widget, '表格')
    
    def add_image_tab(self):
        """添加图片设置选项卡"""
        image_widget = QWidget()
        layout = QFormLayout(image_widget)
        
        # 图片对齐方式
        self.image_alignment = QComboBox()
        self.image_alignment.addItems(['左对齐', '居中', '右对齐'])
        self.image_alignment.setCurrentText('居中')
        layout.addRow('对齐:', self.image_alignment)
        
        # 图注设置
        self.caption_size = QSpinBox()
        self.caption_size.setRange(8, 72)
        self.caption_size.setValue(10.5)
        layout.addRow('图注字号:', self.caption_size)
        
        self.tab_widget.addTab(image_widget, '图片')
    
    def add_page_tab(self):
        """添加页面设置选项卡"""
        page_widget = QWidget()
        layout = QFormLayout(page_widget)
        
        # 页边距设置
        margin_layout = QHBoxLayout()
        self.margin_top = QSpinBox()
        self.margin_top.setRange(0, 100)
        self.margin_top.setValue(25)
        margin_layout.addWidget(QLabel('上:'))
        margin_layout.addWidget(self.margin_top)
        
        self.margin_bottom = QSpinBox()
        self.margin_bottom.setRange(0, 100)
        self.margin_bottom.setValue(25)
        margin_layout.addWidget(QLabel('下:'))
        margin_layout.addWidget(self.margin_bottom)
        
        self.margin_left = QSpinBox()
        self.margin_left.setRange(0, 100)
        self.margin_left.setValue(30)
        margin_layout.addWidget(QLabel('左:'))
        margin_layout.addWidget(self.margin_left)
        
        self.margin_right = QSpinBox()
        self.margin_right.setRange(0, 100)
        self.margin_right.setValue(30)
        margin_layout.addWidget(QLabel('右:'))
        margin_layout.addWidget(self.margin_right)
        
        layout.addRow('页边距:', margin_layout)
        
        # 纸张方向
        self.page_orientation = QComboBox()
        self.page_orientation.addItems(['纵向', '横向'])
        layout.addRow('纸张方向:', self.page_orientation)
        
        self.tab_widget.addTab(page_widget, '页面')
    
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
            print("开始应用格式...")
            # 应用格式
            self.main_window.formatter.format()
            print("格式应用完成")
            
            # 切换到预览页面并更新预览
            print("切换到预览页面...")
            self.main_window.show_preview_page()
            
            self.main_window.show_message("格式已应用，请查看预览")
            
        except Exception as e:
            error_msg = f"应用格式失败: {str(e)}"
            print(error_msg)
            self.main_window.show_message(error_msg, error=True) 