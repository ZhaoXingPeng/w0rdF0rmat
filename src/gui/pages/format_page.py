# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QComboBox,
    QTabWidget, QFormLayout, QSpinBox,
    QLineEdit, QCheckBox, QMessageBox,
    QGroupBox, QDoubleSpinBox
)
from PyQt6.QtCore import Qt

class FormatPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)
        
        # 设置整体样式
        self.setStyleSheet("""
            QWidget {
                background-color: #2c2c2c;
                color: #ffffff;
            }
            QLabel {
                color: #ffffff;
            }
            QGroupBox {
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 4px;
                margin-top: 8px;
                padding-top: 8px;
            }
            QGroupBox::title {
                color: #ffffff;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
            QSpinBox, QDoubleSpinBox, QComboBox {
                background-color: #3d3d3d;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 4px;
            }
            QSpinBox::up-button, QDoubleSpinBox::up-button,
            QSpinBox::down-button, QDoubleSpinBox::down-button {
                border: none;
                background-color: #4d4d4d;
            }
            QComboBox::drop-down {
                border: none;
                background-color: #4d4d4d;
            }
            QTabWidget::pane {
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 10px;
            }
            QTabBar::tab {
                background-color: #3d3d3d;
                color: #ffffff;
                padding: 8px 16px;
                margin-right: 2px;
                border: 1px solid #555555;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #0078d4;
            }
            QPushButton {
                background-color: #0078d4;
                color: #ffffff;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-size: 14px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
        """)
        
        # 添加标题
        title = QLabel("文档格式设置")
        title.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #ffffff;
                padding-bottom: 10px;
            }
        """)
        layout.addWidget(title)
        
        # 创建标签页
        tab_widget = QTabWidget()
        tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #555555;
                background-color: #2c2c2c;
                border-radius: 4px;
                padding: 10px;
            }
            QTabBar::tab {
                background-color: #3d3d3d;
                color: #ffffff;
                padding: 8px 16px;
                margin-right: 2px;
                border: 1px solid #555555;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #0078d4;
            }
            QTabBar::tab:hover {
                background-color: #4d4d4d;
            }
            
            /* 表单样式 */
            QFormLayout {
                background-color: #2c2c2c;
                color: #ffffff;
            }
            QLabel {
                color: #ffffff;
            }
            QSpinBox, QDoubleSpinBox {
                background-color: #3d3d3d;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 4px;
            }
            QComboBox {
                background-color: #3d3d3d;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 4px;
            }
            QComboBox::drop-down {
                border: none;
                background-color: #4d4d4d;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #ffffff;
                margin-right: 5px;
            }
        """)
        
        # 添加段落格式设置
        paragraph_tab = QWidget()
        paragraph_layout = QVBoxLayout(paragraph_tab)
        
        # 段落间距组
        spacing_group = QGroupBox("段落间距")
        spacing_group.setStyleSheet("""
            QGroupBox {
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 4px;
                margin-top: 8px;
                padding-top: 15px;
                background-color: #2c2c2c;
            }
            QGroupBox::title {
                color: #ffffff;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """)
        spacing_layout = QFormLayout()
        
        self.before_spacing = QSpinBox()
        self.before_spacing.setRange(0, 100)
        self.before_spacing.setValue(12)
        spacing_layout.addRow("段前间距(磅):", self.before_spacing)
        
        self.after_spacing = QSpinBox()
        self.after_spacing.setRange(0, 100)
        self.after_spacing.setValue(12)
        spacing_layout.addRow("段后间距(磅):", self.after_spacing)
        
        self.line_spacing = QDoubleSpinBox()
        self.line_spacing.setRange(1, 5)
        self.line_spacing.setValue(1.5)
        self.line_spacing.setSingleStep(0.5)
        spacing_layout.addRow("行间距:", self.line_spacing)
        
        spacing_group.setLayout(spacing_layout)
        paragraph_layout.addWidget(spacing_group)
        
        # 对齐方式
        alignment_group = QGroupBox("对齐方式")
        alignment_group.setStyleSheet("""
            QGroupBox {
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 4px;
                margin-top: 8px;
                padding-top: 15px;
                background-color: #2c2c2c;
            }
            QGroupBox::title {
                color: #ffffff;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """)
        alignment_layout = QFormLayout()
        self.alignment = QComboBox()
        self.alignment.addItems(["左对齐", "居中", "右对齐", "两端对齐"])
        alignment_layout.addRow("段落对齐:", self.alignment)
        alignment_group.setLayout(alignment_layout)
        paragraph_layout.addWidget(alignment_group)
        
        tab_widget.addTab(paragraph_tab, "段落格式")
        
        # 添加字体格式设置
        font_tab = QWidget()
        font_layout = QVBoxLayout(font_tab)
        
        # 字体设置组
        font_group = QGroupBox("字体设置")
        font_group.setStyleSheet("""
            QGroupBox {
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 4px;
                margin-top: 8px;
                padding-top: 15px;
                background-color: #2c2c2c;
            }
            QGroupBox::title {
                color: #ffffff;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """)
        font_form = QFormLayout()
        
        self.font_name = QComboBox()
        self.font_name.addItems(["宋体", "黑体", "楷体", "微软雅黑", "Times New Roman"])
        font_form.addRow("字体:", self.font_name)
        
        self.font_size = QSpinBox()
        self.font_size.setRange(8, 72)
        self.font_size.setValue(12)
        font_form.addRow("字号:", self.font_size)
        
        font_group.setLayout(font_form)
        font_layout.addWidget(font_group)
        
        tab_widget.addTab(font_tab, "字体格式")
        
        layout.addWidget(tab_widget)
        
        # 添加按钮区域
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.apply_btn = QPushButton("应用格式")
        self.apply_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: #ffffff;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-size: 14px;
                min-width: 100px;
                margin: 10px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #666666;
                color: #999999;
            }
        """)
        self.apply_btn.clicked.connect(self.apply_format)
        button_layout.addWidget(self.apply_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def apply_format(self):
        """应用格式设置"""
        try:
            # 获取所有格式设置
            format_settings = {
                'paragraph': {
                    'before_spacing': self.before_spacing.value(),
                    'after_spacing': self.after_spacing.value(),
                    'line_spacing': self.line_spacing.value(),
                    'alignment': self.alignment.currentText()
                },
                'font': {
                    'name': self.font_name.currentText(),
                    'size': self.font_size.value()
                }
            }
            
            # 更新格式设置
            if hasattr(self.main_window.formatter, 'set_format_spec'):
                self.main_window.formatter.set_format_spec(format_settings)
            else:
                # 如果没有 set_format_spec 方法，尝试直接设置属性
                self.main_window.formatter.format_spec = format_settings
            
            # 设置格式已配置状态
            self.main_window.set_format_configured(True)
            
            # 显示成功消息
            QMessageBox.information(self, "成功", "格式设置已应用！\n您现在可以预览文档了。")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"应用格式失败：{str(e)}")
            import traceback
            traceback.print_exc()  # 打印详细错误信息
    
    def show_preview(self):
        """显示预览页面"""
        self.main_window.show_preview_page() 