# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QStackedWidget, 
    QToolBar, QStatusBar, QMessageBox
)
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtCore import Qt
from .pages.document_page import DocumentPage
from .pages.format_page import FormatPage
from .pages.preview_page import PreviewPage

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle('Word文档格式化工具')
        self.setMinimumSize(1200, 800)
        
        # 创建工具栏
        self.create_toolbar()
        
        # 创建状态栏
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        # 创建堆叠部件用于管理页面
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)
        
        # 创建各个页面
        self.document_page = DocumentPage(self)
        self.format_page = FormatPage(self)
        self.preview_page = PreviewPage(self)
        
        # 添加页面到堆叠部件
        self.stacked_widget.addWidget(self.document_page)
        self.stacked_widget.addWidget(self.format_page)
        self.stacked_widget.addWidget(self.preview_page)
        
        # 初始化数据
        self.document = None
        self.formatter = None
    
    def create_toolbar(self):
        """创建工具栏"""
        toolbar = QToolBar()
        self.addToolBar(toolbar)
        
        # 文档管理动作
        doc_action = QAction('文档', self)
        doc_action.triggered.connect(lambda: self.stacked_widget.setCurrentWidget(self.document_page))
        toolbar.addAction(doc_action)
        
        # 格式设置动作
        format_action = QAction('格式', self)
        format_action.triggered.connect(lambda: self.stacked_widget.setCurrentWidget(self.format_page))
        toolbar.addAction(format_action)
        
        # 预览动作
        preview_action = QAction('预览', self)
        preview_action.triggered.connect(lambda: self.stacked_widget.setCurrentWidget(self.preview_page))
        toolbar.addAction(preview_action)
    
    def show_message(self, message: str, error: bool = False):
        """显示消息"""
        if error:
            QMessageBox.critical(self, "错误", message)
        else:
            self.statusBar.showMessage(message) 