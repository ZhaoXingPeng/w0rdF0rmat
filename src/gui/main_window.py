# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QStackedWidget, 
    QToolBar, QStatusBar, QMessageBox, QToolButton, QPushButton
)
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtCore import Qt
from pathlib import Path
from src.gui.pages.document_page import DocumentPage
from src.gui.pages.format_page import FormatPage
from src.gui.pages.preview_page import PreviewPage
from src.config.config_manager import ConfigManager

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # 初始化配置管理器
        self.config_manager = ConfigManager()
        
        # 初始化数据
        self.document = None
        self.formatter = None
        
        # 添加状态标志
        self.document_uploaded = False
        self.format_configured = False
        
        # 设置应用图标
        icon_path = Path(__file__).parent.parent / "resources" / "icons" / "app_icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        
        # 初始化界面
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle('w0rdF0rmat')
        self.setMinimumSize(1200, 800)
        
        # 创建工具栏
        toolbar = QToolBar()
        toolbar.setMovable(False)
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
        
        # 文档管理动作
        doc_button = QToolButton()
        doc_button.setText('1. 打开文档')
        doc_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        doc_button.clicked.connect(self.show_document_page)
        doc_button.setProperty("selected", True)
        toolbar.addWidget(doc_button)
        
        # 格式设置动作
        format_button = QToolButton()
        format_button.setText('2. 设置格式')
        format_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        format_button.clicked.connect(self.show_format_page)
        format_button.setEnabled(False)
        format_button.setProperty("selected", False)
        toolbar.addWidget(format_button)
        
        # 预览动作
        preview_button = QToolButton()
        preview_button.setText('3. 预览结果')
        preview_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        preview_button.clicked.connect(self.show_preview_page)
        preview_button.setEnabled(False)
        preview_button.setProperty("selected", False)
        toolbar.addWidget(preview_button)
        
        # 保存这些按钮的引用
        self.toolbar_buttons = [doc_button, format_button, preview_button]
        
        # 初始状态只启用文档页面
        self.update_toolbar_state()
    
    def show_message(self, message: str, error: bool = False):
        """显示消息"""
        if error:
            msg = QMessageBox(self)
            msg.setWindowTitle("错误")
            msg.setText(message)
            msg.setIcon(QMessageBox.Icon.Critical)
            
            # 完全禁用声音和系统样式
            msg.setWindowFlags(
                Qt.WindowType.Dialog |
                Qt.WindowType.FramelessWindowHint |
                Qt.WindowType.WindowSystemMenuHint |
                Qt.WindowType.NoDropShadowWindowHint |
                Qt.WindowType.WindowStaysOnTopHint |
                Qt.WindowType.MSWindowsFixedSizeDialogHint
            )
            
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: #2c2c2c;
                    color: #ffffff;
                    border: 1px solid #555555;
                    border-radius: 8px;
                }
                QLabel {
                    color: #ffffff;
                    font-size: 14px;
                    padding: 10px;
                }
            """)
            
            # 自定义按钮
            ok_button = QPushButton("确定")
            ok_button.setStyleSheet("""
                QPushButton {
                    background-color: #0078d4;
                    color: white;
                    border: none;
                    padding: 8px 20px;
                    border-radius: 4px;
                    font-size: 14px;
                    min-width: 80px;
                }
                QPushButton:hover {
                    background-color: #106ebe;
                }
                QPushButton:pressed {
                    background-color: #005a9e;
                }
            """)
            msg.addButton(ok_button, QMessageBox.ButtonRole.AcceptRole)
            msg.setDefaultButton(ok_button)
            
            # 禁用默认声音
            msg.setWindowFlags(msg.windowFlags() | Qt.WindowType.WindowDoesNotAcceptFocus)
            
            msg.exec()
        else:
            self.statusBar.showMessage(message, 5000)
    
    def update_toolbar_state(self):
        """更新工具栏状态"""
        try:
            # 更新按钮状态
            self.toolbar_buttons[1].setEnabled(self.document_uploaded)  # 格式按钮
            self.toolbar_buttons[2].setEnabled(self.document_uploaded and self.format_configured)  # 预览按钮
            
            # 更新选中状态
            current_widget = self.stacked_widget.currentWidget()
            
            # 重置所有按钮状态
            for button in self.toolbar_buttons:
                button.setProperty("selected", False)
                button.style().unpolish(button)
                button.style().polish(button)
            
            # 设置当前页面对应的按钮状态
            if current_widget == self.document_page:
                self.toolbar_buttons[0].setProperty("selected", True)
            elif current_widget == self.format_page:
                self.toolbar_buttons[1].setProperty("selected", True)
            elif current_widget == self.preview_page:
                self.toolbar_buttons[2].setProperty("selected", True)
            
            # 强制更新选中按钮的样式
            for button in self.toolbar_buttons:
                if button.property("selected"):
                    button.style().unpolish(button)
                    button.style().polish(button)
            
        except Exception as e:
            print(f"更新工具栏状态失败: {str(e)}")
    
    def show_document_page(self):
        """显示文档页面"""
        self.stacked_widget.setCurrentWidget(self.document_page)
        self.show_message("第一步：请选择要格式化的Word文档")
        self.update_toolbar_state()
    
    def show_format_page(self):
        """显示格式页面"""
        if not self.document_uploaded:
            self.show_message("请先上传文档！", error=True)
            return
        self.stacked_widget.setCurrentWidget(self.format_page)
        self.show_message("第二步：选择或自定义格式设置")
        self.update_toolbar_state()
    
    def show_preview_page(self):
        """显示预览页面"""
        if not self.document_uploaded:
            self.show_message("请先上传文档！", error=True)
            return
        if not self.format_configured:
            self.show_message("请先完成格式设置！", error=True)
            return
        
        try:
            self.stacked_widget.setCurrentWidget(self.preview_page)
            # 只有在需要时才更新预览
            if self.preview_page._needs_reload:
                self.preview_page.update_preview()
            self.show_message("第三步：预览格式化结果并保存")
            self.update_toolbar_state()
        except Exception as e:
            error_msg = f"预览失败: {str(e)}"
            print(error_msg)
            self.show_message(error_msg, error=True)
    
    def set_document_uploaded(self, status: bool):
        """设置文档上传状态"""
        self.document_uploaded = status
        self.update_toolbar_state()
        if status:
            self.show_message("文档已上传，请设置格式")
    
    def set_format_configured(self, status: bool):
        """设置格式配置状态"""
        self.format_configured = status
        self.update_toolbar_state()
        if status:
            self.show_message("格式已设置，可以预览文档了")
    
    def switch_to_preview(self):
        # 切换到预览页面时，使用当前的预览页面状态
        current_page = self.stacked_widget.currentWidget()
        preview_page = self.findChild(PreviewPage)
        if preview_page:
            preview_page.show_preview(self.get_current_text(), self.get_format_settings())