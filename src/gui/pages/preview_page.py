# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QScrollArea, QLabel,
    QHBoxLayout, QFrame, QSplitter,
    QDialog, QApplication, QVBoxLayout, QMenu
)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QUrl, 
    QSize, QPoint, QMimeData  # 从 QtCore 导入 QMimeData
)
from PyQt6.QtGui import (
    QPixmap, QImage, QIcon,
    QDrag, QKeySequence, QAction
)
from pathlib import Path
import tempfile
import os
import win32com.client
import pythoncom
import fitz  # PyMuPDF
from src.gui.components.loading_indicator import LoadingIndicator
from src.utils.temp_manager import TempManager
import threading
import queue

class PreviewWorker(QThread):
    """异步预览工作线程"""
    progress = pyqtSignal(int)  # 进度信号
    finished = pyqtSignal(dict)  # 完成信号
    error = pyqtSignal(str)  # 错误信号

    def __init__(self, doc_path, temp_manager):
        super().__init__()
        self.doc_path = doc_path
        self.temp_manager = temp_manager
        self.cache = {}  # 缓存预览图像

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

    def run(self):
        try:
            print("开始生成预览...")
            
            # 转换原始文档为PDF
            original_pdf = self.temp_manager.get_temp_path("original.pdf")
            self.convert_word_to_pdf(self.doc_path, original_pdf)
            
            # 转换格式化文档为PDF
            formatted_docx = self.temp_manager.get_temp_path("formatted.docx")
            formatted_pdf = self.temp_manager.get_temp_path("formatted.pdf")
            self.convert_word_to_pdf(formatted_docx, formatted_pdf)
            
            # 使用PyMuPDF渲染页面
            doc = fitz.open(original_pdf)  # 打开原始文档
            formatted_doc = fitz.open(formatted_pdf)  # 打开格式化文档
            
            page_images = {}
            total_pages = max(len(doc), len(formatted_doc))
            
            for page_num in range(total_pages):
                self.progress.emit(int((page_num + 1) * 100 / total_pages))
                
                # 渲染原始页面
                if page_num < len(doc):
                    page = doc[page_num]
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                    page_images[f"original_{page_num}"] = QPixmap.fromImage(img)
                
                # 渲染格式化页面
                if page_num < len(formatted_doc):
                    page = formatted_doc[page_num]
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                    page_images[f"formatted_{page_num}"] = QPixmap.fromImage(img)
            
            doc.close()
            formatted_doc.close()
            self.finished.emit(page_images)
            
        except Exception as e:
            self.error.emit(str(e))

class PreviewPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.temp_manager = TempManager()
        self.preview_worker = None
        self.last_format_hash = None  # 添加格式哈希缓存
        self.init_ui()
        
        # 添加快捷键
        self.save_action = QAction("保存文档", self)
        self.save_action.setShortcut("Ctrl+S")
        self.save_action.triggered.connect(self.save_document)
        self.addAction(self.save_action)
    
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # 添加标题区域
        title_container = QFrame()
        title_container.setFixedHeight(50)
        title_container.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border-bottom: 1px solid #e0e0e0;
            }
        """)
        title_layout = QHBoxLayout(title_container)
        title_layout.setContentsMargins(30, 0, 30, 0)
        
        original_label = QLabel("原始文档")
        formatted_label = QLabel("格式化预览")
        label_style = """
            QLabel {
                color: #333333;
                font-size: 15px;
                font-weight: bold;
                padding: 5px 15px;
                border-radius: 4px;
                background-color: #f8f9fa;
            }
        """
        original_label.setStyleSheet(label_style)
        formatted_label.setStyleSheet(label_style)
        title_layout.addWidget(original_label)
        title_layout.addStretch()
        title_layout.addWidget(formatted_label)
        
        layout.addWidget(title_container)
        
        # 创建中央部件
        central_widget = QWidget()
        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(20, 20, 20, 20)
        central_layout.setSpacing(0)
        
        # 创建外层容器
        container_frame = QFrame()
        container_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 12px;
            }
        """)
        container_layout = QVBoxLayout(container_frame)
        container_layout.setContentsMargins(1, 1, 1, 1)
        container_layout.setSpacing(0)
        
        # 创建分割视图
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(1)
        splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #e0e0e0;
            }
        """)
        
        # 创建滚动区域和容器
        # 原始文档视图
        self.original_scroll = QScrollArea()
        self.original_container = QWidget()
        self.original_container.setObjectName("scrollContainer")
        self.original_layout = QVBoxLayout(self.original_container)
        self.original_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.original_layout.setContentsMargins(15, 15, 15, 15)
        self.original_scroll.setWidget(self.original_container)
        self.original_scroll.setWidgetResizable(True)
        
        # 格式化后的视图
        self.formatted_scroll = QScrollArea()
        self.formatted_container = QWidget()
        self.formatted_container.setObjectName("scrollContainer")
        self.formatted_layout = QVBoxLayout(self.formatted_container)
        self.formatted_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.formatted_layout.setContentsMargins(15, 15, 15, 15)
        self.formatted_scroll.setWidget(self.formatted_container)
        self.formatted_scroll.setWidgetResizable(True)
        
        # 设置滚动区域样式
        scroll_style = """
            QScrollArea {
                background-color: #f8f9fa;
                border: none;
                border-radius: 12px;
                margin: 10px;
            }
            QWidget#scrollContainer {
                background-color: #f8f9fa;
                border-radius: 12px;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 8px;
                margin: 10px 2px;
            }
            QScrollBar::handle:vertical {
                background: #c1c1c1;
                min-height: 30px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a8a8a8;
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar::add-page:vertical,
            QScrollBar::sub-page:vertical {
                background: none;
            }
            QScrollBar:horizontal {
                height: 0px;
            }
        """
        
        self.original_scroll.setStyleSheet(scroll_style)
        self.formatted_scroll.setStyleSheet(scroll_style)
        
        # 添加到分割视图
        splitter.addWidget(self.original_scroll)
        splitter.addWidget(self.formatted_scroll)
        
        # 添加分割视图到容器
        container_layout.addWidget(splitter)
        
        # 添加容器到中央部件
        central_layout.addWidget(container_frame)
        layout.addWidget(central_widget)
        
        # 添加保存按钮到右上角
        save_button = QPushButton("保存文档", self)
        save_button.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
        """)
        save_button.clicked.connect(self.save_document)
        
        # 添加按钮到标题布局
        title_layout.addWidget(save_button)
        
        # 添加快捷键
        self.save_action = QAction("保存文档", self)
        self.save_action.setShortcut("Ctrl+S")
        self.save_action.triggered.connect(self.save_document)
        self.addAction(self.save_action)
        
        # 添加右键菜单
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        
        # 添加窗口大小变化事件处理
        def resizeEvent(self, event):
            super().resizeEvent(event)
            # 更新悬浮按钮位置
            self.save_fab.move(
                self.width() - self.save_fab.width() - 20,
                self.height() - self.save_fab.height() - 20
            )
        
        # 启用拖放
        self.setAcceptDrops(True)
        
        # 添加拖放提示区域
        self.drag_hint = QLabel("将预览拖放到文件夹以保存", self)
        self.drag_hint.setStyleSheet("""
            QLabel {
                color: #666666;
                background-color: #f8f9fa;
                padding: 8px 16px;
                border-radius: 20px;
                font-size: 13px;
            }
        """)
        self.drag_hint.hide()
    
    def update_preview(self):
        """更新预览内容"""
        if not self.main_window.document:
            return
            
        try:
            # 显示加载指示器
            self.show_loading_indicators()
            
            # 保存原始文档
            original_docx = self.temp_manager.get_temp_path("original.docx")
            formatted_docx = self.temp_manager.get_temp_path("formatted.docx")
            
            # 复制原始文档
            import shutil
            shutil.copy2(self.main_window.document.path, original_docx)
            shutil.copy2(self.main_window.document.path, formatted_docx)
            
            # 应用格式到复制的文档
            from docx import Document
            formatted_doc = Document(formatted_docx)
            
            # 创建新的格式化器并应用格式
            from src.core.formatter import WordFormatter
            formatter = WordFormatter(
                type('TempDoc', (), {'doc': formatted_doc}),
                self.main_window.config_manager
            )
            formatter.format_spec = self.main_window.formatter.format_spec
            formatter.format()
            
            # 保存格式化后的文档
            formatted_doc.save(formatted_docx)
            
            # 创建并启动预览工作线程
            self.preview_worker = PreviewWorker(
                original_docx,  # 使用临时文档路径
                self.temp_manager
            )
            self.preview_worker.progress.connect(self.update_progress)
            self.preview_worker.finished.connect(self.show_preview_images)
            self.preview_worker.error.connect(self.handle_preview_error)
            self.preview_worker.start()
            
        except Exception as e:
            self.main_window.show_message(f"预览失败: {str(e)}", error=True)
    
    def show_loading_indicators(self):
        """显示加载指示器"""
        self.clear_layout(self.original_layout)
        self.clear_layout(self.formatted_layout)
        
        # 创建预览区域占位符并保存加载指示器的引用
        self.loading_indicators = []
        
        for layout in [self.original_layout, self.formatted_layout]:
            placeholder = QFrame()
            placeholder.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border: 1px solid #e0e0e0;
                    border-radius: 12px;
                    margin: 15px;
                    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
                }
            """)
            placeholder.setMinimumHeight(800)
            
            placeholder_layout = QVBoxLayout(placeholder)
            placeholder_layout.setContentsMargins(0, 0, 0, 0)
            
            loading = LoadingIndicator(self)
            self.loading_indicators.append(loading)
            placeholder_layout.addWidget(loading, 0, Qt.AlignmentFlag.AlignCenter)
            loading.start()
            
            layout.addWidget(placeholder)
    
    def update_progress(self, value):
        """更新进度"""
        # 只在控制台显示进度
        print(f"加载进度: {value}%")
        # 不在状态栏显示"正在加载"文字
        # self.main_window.statusBar.showMessage(f"正在加载预览... {value}%")
    
    def show_preview_images(self, page_images):
        """显示预览图像"""
        try:
            # 移除加载指示器
            for loading in self.loading_indicators:
                loading.stop()
                loading.deleteLater()
            self.loading_indicators.clear()
            
            # 清除现有内容
            self.clear_layout(self.original_layout)
            self.clear_layout(self.formatted_layout)
            
            # 修改页面显示宽度计算
            scroll_width = self.original_scroll.width()
            page_width = int(scroll_width * 0.92)  # 减小宽度比例，留出滚动条空间
            
            # 分别处理原始文档和格式化文档的页面
            original_pages = sorted([k for k in page_images.keys() if k.startswith('original_')])
            formatted_pages = sorted([k for k in page_images.keys() if k.startswith('formatted_')])
            
            # 显示原始文档页面
            for page_key in original_pages:
                page_num = int(page_key.split('_')[1]) + 1  # 提取页码加1
                pixmap = page_images[page_key]
                
                # 缩放图片以适应宽度
                scaled_pixmap = pixmap.scaledToWidth(
                    page_width, 
                    Qt.TransformationMode.SmoothTransformation
                )
                
                # 创建页面容器
                container = self.create_page_container(scaled_pixmap, page_num)
                self.original_layout.addWidget(container)
            
            # 显示格式化文档页面
            for page_key in formatted_pages:
                page_num = int(page_key.split('_')[1]) + 1  # 提取页码并加1
                pixmap = page_images[page_key]
                
                # 缩放图片以适应宽度
                scaled_pixmap = pixmap.scaledToWidth(
                    page_width, 
                    Qt.TransformationMode.SmoothTransformation
                )
                
                # 创建页面容器
                container = self.create_page_container(scaled_pixmap, page_num)
                self.formatted_layout.addWidget(container)
            
            # 添加底部空白
            original_spacer = QWidget()
            original_spacer.setMinimumHeight(20)
            self.original_layout.addWidget(original_spacer)
            
            formatted_spacer = QWidget()
            formatted_spacer.setMinimumHeight(20)
            self.formatted_layout.addWidget(formatted_spacer)
            
            self.main_window.statusBar.showMessage("预览加载完成", 3000)
            print("预览加载完成")
            
        except Exception as e:
            error_msg = f"显示预览失败: {str(e)}"
            print(error_msg)
            self.main_window.show_message(error_msg, error=True)
            import traceback
            traceback.print_exc()  # 添加详细错误信息
    
    def handle_preview_error(self, error_msg):
        """处理预览错误"""
        self.main_window.show_message(f"预览失败: {error_msg}", error=True)
    
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
    
    def save_document(self):
        """保存格式化后的文档"""
        if not self.main_window.document:
            self.main_window.show_message("没有可保存的文档", error=True)
            return
        
        try:
            # 获取保存路径
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存文档",
                "",
                "Word文档 (*.docx)"
            )
            
            if not file_path:  # 用户取消了保存
                return
            
            # 确保文件扩展名正确
            if not file_path.lower().endswith('.docx'):
                file_path += '.docx'
            
            # 获取格式化后的临文档路径
            formatted_docx = self.temp_manager.get_temp_path("formatted.docx")
            
            if os.path.exists(formatted_docx):
                # 复制格式化后的文档到目标位置
                import shutil
                shutil.copy2(formatted_docx, file_path)
                
                self.main_window.show_message(f"文档已保存至: {file_path}")
            else:
                # 如果找不到格式化后的文档，尝试保存原始文档
                self.main_window.document.save(file_path)
                self.main_window.show_message(f"文档已保存至: {file_path}")
                
        except Exception as e:
            error_msg = f"保存文档失败: {str(e)}"
            print(error_msg)
            self.main_window.show_message(error_msg, error=True)
    
    def cleanup(self):
        """清理临时文件"""
        try:
            for file in os.listdir(self.temp_dir):
                os.remove(os.path.join(self.temp_dir, file))
            os.rmdir(self.temp_dir)
        except Exception as e:
            print(f"清理临时文件失败: {str(e)}")
    
    def resizeEvent(self, event):
        """处理窗口大小变化事件"""
        super().resizeEvent(event)
        
        # 重新计算文档宽度
        window_width = self.width()
        available_width = window_width - 60  # 减小边距
        doc_width = available_width // 2
        
        # 更新滚动区域宽度
        self.original_scroll.setMinimumWidth(doc_width)
        self.formatted_scroll.setMinimumWidth(doc_width)
        
        # 如果有预览内容，重新加载以适应新的大小
        if hasattr(self, 'preview_worker') and self.preview_worker:
            self.update_preview()
    
    def clear_layout(self, layout):
        """清除布局中的所有部件"""
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
    
    def _show_error_preview(self, message):
        """显示错误提示"""
        error_dialog = QDialog(self)
        error_dialog.setWindowTitle("错误提示")
        error_dialog.setStyleSheet("""
            QDialog {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                padding: 20px;
                border-radius: 5px;
            }
        """)
        
        error_label = QLabel(message)
        error_label.setStyleSheet("""
            QLabel {
                color: #dc3545;
                font-size: 14px;
                font-weight: bold;
            }
        """)
        
        layout = QVBoxLayout()
        error_dialog.setLayout(layout)
        layout.addWidget(error_label)
        layout.addStretch()
        
        error_dialog.exec()
    
    def _calculate_format_hash(self):
        """计算当前格式的哈希值"""
        if not self.main_window.formatter:
            return None
        import hashlib
        import json
        
        def format_to_dict(obj):
            """将格式对象转换为字典"""
            if hasattr(obj, '__dict__'):
                return {k: format_to_dict(v) for k, v in obj.__dict__.items()
                       if not k.startswith('_')}
            elif isinstance(obj, (list, tuple)):
                return [format_to_dict(x) for x in obj]
            elif isinstance(obj, dict):
                return {k: format_to_dict(v) for k, v in obj.items()}
            else:
                return obj
        
        try:
            # 将格式规范转换为字典
            format_dict = format_to_dict(self.main_window.formatter.format_spec)
            
            # 转换为JSON字符串并排序键值
            format_str = json.dumps(format_dict, sort_keys=True)
            
            # 计算哈希值
            return hashlib.md5(format_str.encode()).hexdigest()
            
        except Exception as e:
            print(f"计算格式哈希失败: {str(e)}")
            return None
    
    def _preview_content_exists(self):
        """检查预览内容是否已存在"""
        original_pdf = self.temp_manager.get_temp_path("original.pdf")
        formatted_pdf = self.temp_manager.get_temp_path("formatted.pdf")
        return os.path.exists(original_pdf) and os.path.exists(formatted_pdf)
    
    def create_page_container(self, pixmap, page_num):
        """创建页面容器"""
        page_container = QFrame()
        page_container.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                margin: 5px;
                padding: 0px;
                box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
            }
            QFrame:hover {
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
            }
        """)
        
        page_layout = QVBoxLayout(page_container)
        page_layout.setContentsMargins(10, 10, 10, 10)
        page_layout.setSpacing(8)
        
        # 创建图片标签
        page_label = QLabel()
        page_label.setPixmap(pixmap)
        page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        page_label.setStyleSheet("""
            QLabel {
                background-color: white;
                padding: 0px;
            }
        """)
        page_layout.addWidget(page_label)
        
        # 添加页码
        page_number = QLabel(f"第 {str(page_num)} 页")
        page_number.setAlignment(Qt.AlignmentFlag.AlignCenter)
        page_number.setStyleSheet("""
            QLabel {
                color: #666666;
                font-size: 12px;
                padding: 4px 8px;
                background-color: #f8f9fa;
                border-radius: 4px;
                margin-top: 5px;
            }
        """)
        page_layout.addWidget(page_number)
        
        return page_container
    
    def show_context_menu(self, position):
        """显示右键菜单"""
        context_menu = QMenu(self)
        context_menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 5px;
            }
            QMenu::item {
                padding: 8px 20px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #f0f9ff;
                color: #0078d4;
            }
        """)
        
        # 添加保存选项
        save_action = context_menu.addAction("保存文档")
        save_action.setIcon(QIcon(":/icons/save.png"))
        save_action.triggered.connect(self.save_document)
        
        # 显示菜单
        context_menu.exec(self.mapToGlobal(position))
    
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_start_position = event.pos()

    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.MouseButton.LeftButton):
            return
            
        if (event.pos() - self.drag_start_position).manhattanLength() < QApplication.startDragDistance():
            return

        drag = QDrag(self)
        mimedata = QMimeData()
        
        # 创建临时文件
        temp_file = self.temp_manager.get_temp_path("temp.docx")
        self.main_window.document.save(temp_file)
        
        # 设置拖放数据
        mimedata.setUrls([QUrl.fromLocalFile(temp_file)])
        drag.setMimeData(mimedata)
        
        # 显示拖放提示
        self.drag_hint.show()
        
        # 执行拖放
        result = drag.exec(Qt.DropAction.CopyAction)
        
        # 隐藏提示
        self.drag_hint.hide()
    