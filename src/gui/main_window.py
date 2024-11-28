# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QStackedWidget, 
    QToolBar, QStatusBar, QMessageBox, QToolButton
)
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtCore import Qt
from src.gui.pages.document_page import DocumentPage
from src.gui.pages.format_page import FormatPage
from src.gui.pages.preview_page import PreviewPage

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle('Word文档格式化工具')
        self.setMinimumSize(1200, 800)
        
        # 创建工具栏
        toolbar = QToolBar()
        toolbar.setMovable(False)  # 禁止移动工具栏
        toolbar.setStyleSheet("""
            QToolBar {
                background-color: #2c2c2c;
                border: none;
                padding: 5px;
                spacing: 5px;
            }
            QToolButton {
                color: #b8b8b8;
                background-color: transparent;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                margin: 0 2px;
            }
            QToolButton:hover {
                background-color: #3d3d3d;
            }
            QToolButton[selected="true"] {
                color: white;
                background-color: #0078d4;
            }
            QToolButton:disabled {
                color: #666666;
            }
        """)
        self.addToolBar(toolbar)
        
        # 文档管理动作
        self.doc_action = QAction('1. 打开文档', self)
        self.doc_action.triggered.connect(self.show_document_page)
        self.doc_action.setProperty("selected", True)
        toolbar.addAction(self.doc_action)
        
        # 格式设置动作
        self.format_action = QAction('2. 设置格式', self)
        self.format_action.triggered.connect(self.show_format_page)
        self.format_action.setEnabled(False)
        self.format_action.setProperty("selected", False)
        toolbar.addAction(self.format_action)
        
        # 预览动作
        self.preview_action = QAction('3. 预览结果', self)
        self.preview_action.triggered.connect(self.show_preview_page)
        self.preview_action.setEnabled(False)
        self.preview_action.setProperty("selected", False)
        toolbar.addAction(self.preview_action)
        
        # 保存这些动作的引用
        self.toolbar_actions = [self.doc_action, self.format_action, self.preview_action]
        
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
        
        # 创建状态栏
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
    
    def show_message(self, message: str, error: bool = False):
        """显示消息
        Args:
            message: 消息内容
            error: 是否为错误消息
        """
        if error:
            QMessageBox.critical(self, "错误", message)
        else:
            self.statusBar.showMessage(message, 5000)  # 显示5秒
    
    def create_toolbar(self):
        """创建工具栏"""
        toolbar = QToolBar()
        self.addToolBar(toolbar)
        
        # 文档管理动作
        self.doc_action = QAction('1. 打开文档', self)
        self.doc_action.triggered.connect(self.show_document_page)
        toolbar.addAction(self.doc_action)
        
        # 格式设置动作
        self.format_action = QAction('2. 设置格式', self)
        self.format_action.triggered.connect(self.show_format_page)
        toolbar.addAction(self.format_action)
        
        # 预览动作
        self.preview_action = QAction('3. 预览结果', self)
        self.preview_action.triggered.connect(self.show_preview_page)
        toolbar.addAction(self.preview_action)
        
        # 保存这些动作的引用
        self.toolbar_actions = [self.doc_action, self.format_action, self.preview_action]
    
    def show_document_page(self):
        """显示文档页面"""
        self.stacked_widget.setCurrentWidget(self.document_page)
        self.show_message("第一步：请选择要格式化的Word文档")
    
    def show_format_page(self):
        """显示格式页面"""
        if not self.document:
            self.show_message("请先打开文档！", error=True)
            return
        self.stacked_widget.setCurrentWidget(self.format_page)
        self.show_message("第二步：选择或自定义格式设置")
    
    def show_preview_page(self):
        """显示预览页面"""
        if not self.document:
            self.show_message("请先打开文档！", error=True)
            return
        
        try:
            print("切换到预览页面...")
            self.stacked_widget.setCurrentWidget(self.preview_page)
            
            # 强制更新预览
            print("开始更新预览...")
            self.preview_page.update_preview()
            print("预览更新完成")
            
            self.show_message("第三步：预览格式化结果并保存")
            
        except Exception as e:
            error_msg = f"预览失败: {str(e)}"
            print(error_msg)
            self.show_message(error_msg, error=True)
    
    def update_toolbar_state(self):
        """更新工具栏状态"""
        # 更新按钮状态
        has_document = bool(self.document)
        self.format_action.setEnabled(has_document)
        self.preview_action.setEnabled(has_document)
        
        # 更新选中状态
        current_widget = self.stacked_widget.currentWidget()
        
        # 重置所有按钮状态
        for action in self.toolbar_actions:
            action.setProperty("selected", False)
            toolbar_button = self.findChild(QToolButton, action.text())
            if toolbar_button:
                toolbar_button.setProperty("selected", False)
                toolbar_button.style().unpolish(toolbar_button)
                toolbar_button.style().polish(toolbar_button)
        
        # 设置当前页面对应的按钮状态
        if current_widget == self.document_page:
            self.doc_action.setProperty("selected", True)
            toolbar_button = self.findChild(QToolButton, self.doc_action.text())
        elif current_widget == self.format_page:
            self.format_action.setProperty("selected", True)
            toolbar_button = self.findChild(QToolButton, self.format_action.text())
        elif current_widget == self.preview_page:
            self.preview_action.setProperty("selected", True)
            toolbar_button = self.findChild(QToolButton, self.preview_action.text())
        
        if toolbar_button:
            toolbar_button.setProperty("selected", True)
            toolbar_button.style().unpolish(toolbar_button)
            toolbar_button.style().polish(toolbar_button)
        
        # 强制更新工具栏样式
        self.style().polish(self.findChild(QToolBar))
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
    def unpolish(self, widget):
        """强制取消样式"""
        widget.style().unpolish(widget)
    
    def findChild(self, child_type, name):
        """查找子部件
        Args:
            child_type: 子部件类型
            name: 子部件名称
        Returns:
            找到的子部件
        """
        for child in self.findChildren(child_type):
            if child.objectName() == name:
                return child
        return None
    
    def style(self):
        """获取样式"""
        return self.style()
    
    def polish(self, widget):
        """强制更新样式"""
        widget.style().polish(widget)
    
        # 强制更新样式
        self.style().polish(self) 