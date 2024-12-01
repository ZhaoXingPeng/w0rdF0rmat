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
        
        # 添加标题
        title = QLabel("论文格式设置")
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
        
        # 1. 封面页格式
        cover_tab = self.create_cover_tab()
        tab_widget.addTab(cover_tab, "封面格式")
        
        # 2. 摘要格式
        abstract_tab = self.create_abstract_tab()
        tab_widget.addTab(abstract_tab, "摘要格式")
        
        # 3. 目录格式
        contents_tab = self.create_contents_tab()
        tab_widget.addTab(contents_tab, "目录格式")
        
        # 4. 正文格式
        main_tab = self.create_main_text_tab()
        tab_widget.addTab(main_tab, "正文格式")
        
        # 5. 参考文献格式
        references_tab = self.create_references_tab()
        tab_widget.addTab(references_tab, "参考文献格式")
        
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
            QPushButton:disabled {
                background-color: #666666;
                color: #999999;
            }
        """)
        self.apply_btn.clicked.connect(self.apply_format)
        button_layout.addWidget(self.apply_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)

    def create_cover_tab(self):
        """创建封面格式标签页"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 标题格式组
        title_group = QGroupBox("论文标题")
        title_layout = QFormLayout()
        
        self.title_font = QComboBox()
        self.title_font.addItems(["黑体", "宋体", "楷体", "微软雅黑"])
        title_layout.addRow("字体:", self.title_font)
        
        self.title_size = QSpinBox()
        self.title_size.setRange(12, 72)
        self.title_size.setValue(22)
        title_layout.addRow("字号:", self.title_size)
        
        title_group.setLayout(title_layout)
        layout.addWidget(title_group)
        
        # 其他封面元素格式组
        other_group = QGroupBox("其他封面元素")
        other_layout = QFormLayout()
        
        self.school_font = QComboBox()
        self.school_font.addItems(["宋体", "黑体", "楷体", "微软雅黑"])
        other_layout.addRow("学校名称字体:", self.school_font)
        
        self.school_size = QSpinBox()
        self.school_size.setRange(12, 48)
        self.school_size.setValue(16)
        other_layout.addRow("学校名称字号:", self.school_size)
        
        other_group.setLayout(other_layout)
        layout.addWidget(other_group)
        
        layout.addStretch()
        return tab

    def create_abstract_tab(self):
        """创建摘要格式标签页"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 摘要标题格式
        title_group = QGroupBox("摘要标题")
        title_layout = QFormLayout()
        
        self.abstract_title_font = QComboBox()
        self.abstract_title_font.addItems(["黑体", "宋体", "楷体"])
        title_layout.addRow("字体:", self.abstract_title_font)
        
        self.abstract_title_size = QSpinBox()
        self.abstract_title_size.setRange(12, 24)
        self.abstract_title_size.setValue(16)
        title_layout.addRow("字号:", self.abstract_title_size)
        
        self.abstract_title_align = QComboBox()
        self.abstract_title_align.addItems(["居中", "左对齐", "右对齐"])
        title_layout.addRow("对齐方式:", self.abstract_title_align)
        
        title_group.setLayout(title_layout)
        layout.addWidget(title_group)
        
        # 摘要正文格式
        content_group = QGroupBox("摘要正文")
        content_layout = QFormLayout()
        
        self.abstract_font = QComboBox()
        self.abstract_font.addItems(["宋体", "楷体", "微软雅黑"])
        content_layout.addRow("字体:", self.abstract_font)
        
        self.abstract_size = QSpinBox()
        self.abstract_size.setRange(10, 16)
        self.abstract_size.setValue(12)
        content_layout.addRow("字号:", self.abstract_size)
        
        self.abstract_line_spacing = QDoubleSpinBox()
        self.abstract_line_spacing.setRange(1.0, 3.0)
        self.abstract_line_spacing.setValue(1.5)
        self.abstract_line_spacing.setSingleStep(0.25)
        content_layout.addRow("行间距:", self.abstract_line_spacing)
        
        self.abstract_para_spacing = QSpinBox()
        self.abstract_para_spacing.setRange(0, 30)
        self.abstract_para_spacing.setValue(10)
        content_layout.addRow("段落间距:", self.abstract_para_spacing)
        
        self.abstract_first_line_indent = QSpinBox()
        self.abstract_first_line_indent.setRange(0, 4)
        self.abstract_first_line_indent.setValue(2)
        content_layout.addRow("首行缩进(字符):", self.abstract_first_line_indent)
        
        self.abstract_align = QComboBox()
        self.abstract_align.addItems(["两端对齐", "左对齐", "右对齐", "居中"])
        content_layout.addRow("对齐方式:", self.abstract_align)
        
        # 页边距设置
        margin_group = QGroupBox("页边距")
        margin_layout = QFormLayout()
        
        self.abstract_margin_top = QSpinBox()
        self.abstract_margin_top.setRange(10, 50)
        self.abstract_margin_top.setValue(25)
        margin_layout.addRow("上边距(毫米):", self.abstract_margin_top)
        
        self.abstract_margin_bottom = QSpinBox()
        self.abstract_margin_bottom.setRange(10, 50)
        self.abstract_margin_bottom.setValue(25)
        margin_layout.addRow("下边距(毫米):", self.abstract_margin_bottom)
        
        self.abstract_margin_left = QSpinBox()
        self.abstract_margin_left.setRange(10, 50)
        self.abstract_margin_left.setValue(30)
        margin_layout.addRow("左边距(毫米):", self.abstract_margin_left)
        
        self.abstract_margin_right = QSpinBox()
        self.abstract_margin_right.setRange(10, 50)
        self.abstract_margin_right.setValue(30)
        margin_layout.addRow("右边距(毫米):", self.abstract_margin_right)
        
        margin_group.setLayout(margin_layout)
        layout.addWidget(margin_group)
        
        content_group.setLayout(content_layout)
        layout.addWidget(content_group)
        
        layout.addStretch()
        return tab

    def create_main_text_tab(self):
        """创建正文格式标签页"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 章节标题格式
        chapter_group = QGroupBox("章节标题")
        chapter_layout = QFormLayout()
        
        self.chapter_font = QComboBox()
        self.chapter_font.addItems(["黑体", "宋体", "微软雅黑"])
        chapter_layout.addRow("字体:", self.chapter_font)
        
        self.chapter_size = QSpinBox()
        self.chapter_size.setRange(12, 24)
        self.chapter_size.setValue(16)
        chapter_layout.addRow("字号:", self.chapter_size)
        
        self.chapter_align = QComboBox()
        self.chapter_align.addItems(["左对齐", "居中", "右对齐"])
        chapter_layout.addRow("对齐方式:", self.chapter_align)
        
        self.chapter_spacing = QSpinBox()
        self.chapter_spacing.setRange(0, 50)
        self.chapter_spacing.setValue(24)
        chapter_layout.addRow("段后间距:", self.chapter_spacing)
        
        chapter_group.setLayout(chapter_layout)
        layout.addWidget(chapter_group)
        
        # 正��格式
        body_group = QGroupBox("正文格式")
        body_layout = QFormLayout()
        
        self.body_font = QComboBox()
        self.body_font.addItems(["宋体", "楷体", "微软雅黑"])
        body_layout.addRow("字体:", self.body_font)
        
        self.body_size = QSpinBox()
        self.body_size.setRange(10, 16)
        self.body_size.setValue(12)
        body_layout.addRow("字号:", self.body_size)
        
        self.line_spacing = QDoubleSpinBox()
        self.line_spacing.setRange(1.0, 3.0)
        self.line_spacing.setValue(1.5)
        self.line_spacing.setSingleStep(0.25)
        body_layout.addRow("行间距:", self.line_spacing)
        
        self.para_spacing = QSpinBox()
        self.para_spacing.setRange(0, 30)
        self.para_spacing.setValue(10)
        body_layout.addRow("段落间距:", self.para_spacing)
        
        self.first_line_indent = QSpinBox()
        self.first_line_indent.setRange(0, 4)
        self.first_line_indent.setValue(2)
        body_layout.addRow("首行缩进(字符):", self.first_line_indent)
        
        self.body_align = QComboBox()
        self.body_align.addItems(["两端对齐", "左对齐", "右对齐", "居中"])
        body_layout.addRow("对齐方式:", self.body_align)
        
        body_group.setLayout(body_layout)
        layout.addWidget(body_group)
        
        # 页边距设置
        margin_group = QGroupBox("页边距")
        margin_layout = QFormLayout()
        
        self.margin_top = QSpinBox()
        self.margin_top.setRange(10, 50)
        self.margin_top.setValue(25)
        margin_layout.addRow("上边距(毫米):", self.margin_top)
        
        self.margin_bottom = QSpinBox()
        self.margin_bottom.setRange(10, 50)
        self.margin_bottom.setValue(25)
        margin_layout.addRow("下边距(毫米):", self.margin_bottom)
        
        self.margin_left = QSpinBox()
        self.margin_left.setRange(10, 50)
        self.margin_left.setValue(30)
        margin_layout.addRow("左边距(毫米):", self.margin_left)
        
        self.margin_right = QSpinBox()
        self.margin_right.setRange(10, 50)
        self.margin_right.setValue(30)
        margin_layout.addRow("右边距(毫米):", self.margin_right)
        
        margin_group.setLayout(margin_layout)
        layout.addWidget(margin_group)
        
        layout.addStretch()
        return tab

    def create_contents_tab(self):
        """创建目录格式标签页"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 目录标题格式
        title_group = QGroupBox("目录标题")
        title_layout = QFormLayout()
        
        self.contents_title_font = QComboBox()
        self.contents_title_font.addItems(["黑体", "宋体", "楷体"])
        title_layout.addRow("字体:", self.contents_title_font)
        
        self.contents_title_size = QSpinBox()
        self.contents_title_size.setRange(12, 24)
        self.contents_title_size.setValue(16)
        title_layout.addRow("字号:", self.contents_title_size)
        
        self.contents_title_align = QComboBox()
        self.contents_title_align.addItems(["居中", "左对齐", "右对齐"])
        title_layout.addRow("对齐方式:", self.contents_title_align)
        
        self.contents_title_spacing = QSpinBox()
        self.contents_title_spacing.setRange(0, 50)
        self.contents_title_spacing.setValue(24)
        title_layout.addRow("段后间距:", self.contents_title_spacing)
        
        title_group.setLayout(title_layout)
        layout.addWidget(title_group)
        
        # 目录项格式
        items_group = QGroupBox("目录项格式")
        items_layout = QFormLayout()
        
        self.contents_font = QComboBox()
        self.contents_font.addItems(["宋体", "楷体", "微软雅黑"])
        items_layout.addRow("字体:", self.contents_font)
        
        self.contents_size = QSpinBox()
        self.contents_size.setRange(10, 16)
        self.contents_size.setValue(12)
        items_layout.addRow("字号:", self.contents_size)
        
        self.contents_line_spacing = QDoubleSpinBox()
        self.contents_line_spacing.setRange(1.0, 2.0)
        self.contents_line_spacing.setValue(1.15)
        self.contents_line_spacing.setSingleStep(0.05)
        items_layout.addRow("行间距:", self.contents_line_spacing)
        
        self.contents_level_indent = QSpinBox()
        self.contents_level_indent.setRange(0, 4)
        self.contents_level_indent.setValue(2)
        items_layout.addRow("层级缩进(字符):", self.contents_level_indent)
        
        self.contents_align = QComboBox()
        self.contents_align.addItems(["左对齐", "两端对齐"])
        items_layout.addRow("对齐方式:", self.contents_align)
        
        items_group.setLayout(items_layout)
        layout.addWidget(items_group)
        
        # 页码格式
        page_num_group = QGroupBox("页码格式")
        page_num_layout = QFormLayout()
        
        self.page_num_font = QComboBox()
        self.page_num_font.addItems(["Times New Roman", "宋体", "Arial"])
        page_num_layout.addRow("字体:", self.page_num_font)
        
        self.page_num_size = QSpinBox()
        self.page_num_size.setRange(8, 14)
        self.page_num_size.setValue(10)
        page_num_layout.addRow("字号:", self.page_num_size)
        
        page_num_group.setLayout(page_num_layout)
        layout.addWidget(page_num_group)
        
        layout.addStretch()
        return tab
    
    def create_references_tab(self):
        """创建参考文献格式标签页"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 参考文献标题格式
        title_group = QGroupBox("参考文献标题")
        title_layout = QFormLayout()
        
        self.ref_title_font = QComboBox()
        self.ref_title_font.addItems(["黑体", "宋体", "楷体"])
        title_layout.addRow("字体:", self.ref_title_font)
        
        self.ref_title_size = QSpinBox()
        self.ref_title_size.setRange(12, 24)
        self.ref_title_size.setValue(16)
        title_layout.addRow("字号:", self.ref_title_size)
        
        self.ref_title_align = QComboBox()
        self.ref_title_align.addItems(["居中", "左对齐", "右对齐"])
        title_layout.addRow("对齐方式:", self.ref_title_align)
        
        self.ref_title_spacing = QSpinBox()
        self.ref_title_spacing.setRange(0, 50)
        self.ref_title_spacing.setValue(24)
        title_layout.addRow("段后间距:", self.ref_title_spacing)
        
        title_group.setLayout(title_layout)
        layout.addWidget(title_group)
        
        # 参考文献条目格式
        items_group = QGroupBox("参考文献条目")
        items_layout = QFormLayout()
        
        self.ref_font = QComboBox()
        self.ref_font.addItems(["宋体", "楷体", "微软雅黑"])
        items_layout.addRow("字体:", self.ref_font)
        
        self.ref_size = QSpinBox()
        self.ref_size.setRange(10, 16)
        self.ref_size.setValue(12)
        items_layout.addRow("字号:", self.ref_size)
        
        self.ref_line_spacing = QDoubleSpinBox()
        self.ref_line_spacing.setRange(1.0, 2.0)
        self.ref_line_spacing.setValue(1.15)
        self.ref_line_spacing.setSingleStep(0.05)
        items_layout.addRow("行间距:", self.ref_line_spacing)
        
        self.ref_para_spacing = QSpinBox()
        self.ref_para_spacing.setRange(0, 20)
        self.ref_para_spacing.setValue(6)
        items_layout.addRow("条目间距:", self.ref_para_spacing)
        
        self.ref_hanging_indent = QSpinBox()
        self.ref_hanging_indent.setRange(0, 4)
        self.ref_hanging_indent.setValue(2)
        items_layout.addRow("悬挂缩进(字符):", self.ref_hanging_indent)
        
        self.ref_align = QComboBox()
        self.ref_align.addItems(["两端对齐", "左对齐"])
        items_layout.addRow("对齐方式:", self.ref_align)
        
        items_group.setLayout(items_layout)
        layout.addWidget(items_group)
        
        layout.addStretch()
        return tab

    def apply_format(self):
        """应用格式设置"""
        try:
            format_settings = {
                'cover': {
                    'title': {
                        'font': self.title_font.currentText(),
                        'size': self.title_size.value()
                    },
                    'school': {
                        'font': self.school_font.currentText(),
                        'size': self.school_size.value()
                    }
                },
                'abstract': {
                    'title': {
                        'font': self.abstract_title_font.currentText(),
                        'size': self.abstract_title_size.value(),
                        'align': self.abstract_title_align.currentText()
                    },
                    'content': {
                        'font': self.abstract_font.currentText(),
                        'size': self.abstract_size.value(),
                        'line_spacing': self.abstract_line_spacing.value(),
                        'para_spacing': self.abstract_para_spacing.value(),
                        'first_line_indent': self.abstract_first_line_indent.value(),
                        'align': self.abstract_align.currentText()
                    },
                    'margin': {
                        'top': self.abstract_margin_top.value(),
                        'bottom': self.abstract_margin_bottom.value(),
                        'left': self.abstract_margin_left.value(),
                        'right': self.abstract_margin_right.value()
                    }
                },
                'contents': {
                    'title': {
                        'font': self.contents_title_font.currentText(),
                        'size': self.contents_title_size.value(),
                        'align': self.contents_title_align.currentText(),
                        'spacing': self.contents_title_spacing.value()
                    },
                    'items': {
                        'font': self.contents_font.currentText(),
                        'size': self.contents_size.value(),
                        'line_spacing': self.contents_line_spacing.value(),
                        'level_indent': self.contents_level_indent.value(),
                        'align': self.contents_align.currentText()
                    }
                },
                'main_text': {
                    'chapter': {
                        'font': self.chapter_font.currentText(),
                        'size': self.chapter_size.value(),
                        'align': self.chapter_align.currentText(),
                        'spacing': self.chapter_spacing.value()
                    },
                    'body': {
                        'font': self.body_font.currentText(),
                        'size': self.body_size.value(),
                        'line_spacing': self.line_spacing.value(),
                        'para_spacing': self.para_spacing.value(),
                        'first_line_indent': self.first_line_indent.value(),
                        'align': self.body_align.currentText()
                    },
                    'margin': {
                        'top': self.margin_top.value(),
                        'bottom': self.margin_bottom.value(),
                        'left': self.margin_left.value(),
                        'right': self.margin_right.value()
                    }
                },
                'references': {
                    'title': {
                        'font': self.ref_title_font.currentText(),
                        'size': self.ref_title_size.value(),
                        'align': self.ref_title_align.currentText(),
                        'spacing': self.ref_title_spacing.value()
                    },
                    'items': {
                        'font': self.ref_font.currentText(),
                        'size': self.ref_size.value(),
                        'line_spacing': self.ref_line_spacing.value(),
                        'para_spacing': self.ref_para_spacing.value(),
                        'hanging_indent': self.ref_hanging_indent.value(),
                        'align': self.ref_align.currentText()
                    }
                }
            }
            
            # 更新格式设置
            self.main_window.formatter.set_format_spec(format_settings)
            
            # 设置格式已配置状态
            self.main_window.set_format_configured(True)
            
            # 直接跳转到预览页面
            self.main_window.show_preview_page()
            
        except Exception as e:
            error_msg = f"应用格式失败：{str(e)}"
            print(error_msg)  # 保留调试信息
            self.main_window.show_message(error_msg, error=True)
            import traceback
            traceback.print_exc()
    
    def show_preview(self):
        """显示预览页面"""
        self.main_window.show_preview_page() 