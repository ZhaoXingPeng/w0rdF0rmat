# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QScrollArea, QLabel,
    QHBoxLayout, QFrame, QSplitter,
    QDialog, QApplication, QVBoxLayout
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QImage
from pathlib import Path
import tempfile
import os
import win32com.client
import pythoncom
import fitz  # PyMuPDF
from src.gui.components.loading_indicator import LoadingIndicator

class PreviewPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.temp_dir = tempfile.mkdtemp()
        self.last_format_hash = None  # 添加格式哈希值记录
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)  # 减少组件间距
        
        # 添加标题区域
        title_container = QFrame()
        title_container.setFixedHeight(36)  # 减小标题高度
        title_container.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border-bottom: 1px solid #dee2e6;
            }
        """)
        title_layout = QHBoxLayout(title_container)
        title_layout.setContentsMargins(20, 0, 20, 0)
        
        original_label = QLabel("原始文档")
        formatted_label = QLabel("格式化预览")
        label_style = """
            QLabel {
                color: #495057;
                font-size: 13px;
                font-weight: bold;
            }
        """
        original_label.setStyleSheet(label_style)
        formatted_label.setStyleSheet(label_style)
        title_layout.addWidget(original_label)
        title_layout.addStretch()
        title_layout.addWidget(formatted_label)
        
        layout.addWidget(title_container)
        
        # 创建中央部件来容纳分割视图和文档内容
        central_widget = QWidget()
        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(0, 0, 0, 0)
        central_layout.setSpacing(0)
        
        # 创建分割视图
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        
        # 计算初始宽度
        window_width = self.window().width()
        available_width = window_width - 40  # 减去边距
        doc_width = (available_width - splitter.handleWidth()) // 2  # 考虑分割条宽度
        
        # 原始文档视图
        self.original_scroll = QScrollArea()
        self.original_container = QWidget()
        self.original_layout = QVBoxLayout(self.original_container)
        self.original_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.original_scroll.setWidget(self.original_container)
        self.original_scroll.setWidgetResizable(True)
        self.original_scroll.setFixedWidth(doc_width)
        
        # 格式化后的视图
        self.formatted_scroll = QScrollArea()
        self.formatted_container = QWidget()
        self.formatted_layout = QVBoxLayout(self.formatted_container)
        self.formatted_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.formatted_scroll.setWidget(self.formatted_container)
        self.formatted_scroll.setWidgetResizable(True)
        self.formatted_scroll.setFixedWidth(doc_width)
        
        # 设置滚动区域样式
        scroll_style = """
            QScrollArea {
                background-color: #ffffff;
                border: none;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 8px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #c1c1c1;
                min-height: 30px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a8a8a8;
            }
            QScrollBar:horizontal {
                height: 0px;  /* 隐藏水平滚动条 */
            }
        """
        self.original_scroll.setStyleSheet(scroll_style)
        self.formatted_scroll.setStyleSheet(scroll_style)
        
        # 设置最小和最大高度
        min_height = self.window().height() - 100  # 减去标题和按钮区域的高度
        self.original_scroll.setMinimumHeight(min_height)
        self.formatted_scroll.setMinimumHeight(min_height)
        
        splitter.addWidget(self.original_scroll)
        splitter.addWidget(self.formatted_scroll)
        
        # 设置分割比例
        splitter.setSizes([doc_width, doc_width])
        
        central_layout.addWidget(splitter)
        layout.addWidget(central_widget, 1)  # 1表示伸展因子
        
        # 添加底部按钮
        button_container = QFrame()
        button_container.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border-top: 1px solid #dee2e6;
                padding: 10px;
            }
        """)
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(20, 10, 20, 10)
        
        self.save_btn = QPushButton('保存文档')
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-size: 14px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
        """)
        self.save_btn.clicked.connect(self.save_document)
        
        button_layout.addStretch()
        button_layout.addWidget(self.save_btn)
        
        layout.addWidget(button_container)
    
    def update_preview(self):
        """更新预览内容"""
        if not self.main_window.document:
            return
            
        # 计算当前格式的哈希值
        current_format_hash = self._calculate_format_hash()
        
        # 如果格式没有变化且预览内容已存在，则不需要重新加载
        if (current_format_hash == self.last_format_hash and 
            self._preview_content_exists()):
            print("格式未变化，跳过预览更新")
            return
            
        try:
            # 创建并显示加载指示器
            loading_left = LoadingIndicator(self)
            loading_right = LoadingIndicator(self)
            
            # 清除原有内容并添加加载指示器
            self.clear_layout(self.original_layout)
            self.clear_layout(self.formatted_layout)
            
            # 将加载指示器添加到布局中央
            self.original_layout.addWidget(loading_left, 0, Qt.AlignmentFlag.AlignCenter)
            self.formatted_layout.addWidget(loading_right, 0, Qt.AlignmentFlag.AlignCenter)
            
            # 启动加载动画
            loading_left.start()
            loading_right.start()
            
            # 强制更新界面
            QApplication.processEvents()
            
            # 保存原始文档
            original_docx = os.path.join(self.temp_dir, "original.docx")
            print(f"保存原始文档到: {original_docx}")
            try:
                # 创建原始文档的副本
                import shutil
                shutil.copy2(self.main_window.document.path, original_docx)
                print("原始文档保存成功")
            except Exception as e:
                print(f"保存原始文档失败: {str(e)}")
                raise
            
            # 转换原始文档为PDF
            original_pdf = os.path.join(self.temp_dir, "original.pdf")
            print(f"转换原始文档为PDF: {original_pdf}")
            try:
                self.convert_word_to_pdf(original_docx, original_pdf)
                print("原始文档转换为PDF成功")
            except Exception as e:
                print(f"转换原始文档为PDF失败: {str(e)}")
                raise
            
            # 显示原始文档
            print("显示原始文档预览")
            try:
                # 移除左侧加载指示器
                loading_left.stop()
                loading_left.deleteLater()
                
                self.show_pdf_preview(original_pdf, self.original_layout)
                print("原始文档预览显示成功")
            except Exception as e:
                print(f"显示原始文档预览失败: {str(e)}")
                raise
            
            # 创建并格式化新文档
            formatted_docx = os.path.join(self.temp_dir, "formatted.docx")
            print(f"创建格式化文档: {formatted_docx}")
            try:
                # 复制原始文档
                shutil.copy2(original_docx, formatted_docx)
                print("文档副本创建成功")
                
                # 创建新的Document对象并应用格式
                from docx import Document
                formatted_doc = Document(formatted_docx)
                
                # 保存当前文档对象
                original_doc = self.main_window.document.doc
                
                # 临时替换文档对象进行格式化
                self.main_window.document.doc = formatted_doc
                print("开始应用格式...")
                self.main_window.formatter.format()
                print("格式应用完成")
                formatted_doc.save(formatted_docx)
                print("格式化文档保存成功")
                
                # 恢复原始文档对象
                self.main_window.document.doc = original_doc
                
            except Exception as e:
                print(f"格式化文档失败: {str(e)}")
                raise
            
            # 转换格式化后的文档为PDF
            formatted_pdf = os.path.join(self.temp_dir, "formatted.pdf")
            print(f"转换格式化文档为PDF: {formatted_pdf}")
            try:
                self.convert_word_to_pdf(formatted_docx, formatted_pdf)
                print("格式化文档转换为PDF成功")
            except Exception as e:
                print(f"转换格式化文档为PDF失败: {str(e)}")
                raise
            
            # 显示格式化后的文档
            print("显示格式化文档预览")
            try:
                # 移除右侧加载指示器
                loading_right.stop()
                loading_right.deleteLater()
                
                self.show_pdf_preview(formatted_pdf, self.formatted_layout)
                print("格式化文档预览显示成功")
            except Exception as e:
                print(f"显示格式化文档预览失败: {str(e)}")
                raise
            
            # 更新格式哈希值
            self.last_format_hash = current_format_hash
            
            print("=== 预览更新完成 ===\n")
            
        except Exception as e:
            error_msg = f"更新预览失败: {str(e)}"
            print(error_msg)
            self.main_window.show_message(error_msg, error=True)
            
            # 移除加载指示器
            if 'loading_left' in locals():
                loading_left.stop()
                loading_left.deleteLater()
            if 'loading_right' in locals():
                loading_right.stop()
                loading_right.deleteLater()
            
            # 显示错误提示
            self._show_error_preview("预览加载失败")
    
    def show_pdf_preview(self, pdf_path, target_layout):
        """显示PDF预览"""
        try:
            # 清除现有内容
            for i in reversed(range(target_layout.count())): 
                widget = target_layout.itemAt(i).widget()
                if widget is not None:
                    widget.setParent(None)
            
            # 显示加载指示器
            loading = LoadingIndicator(self)
            target_layout.addWidget(loading, alignment=Qt.AlignmentFlag.AlignCenter)
            loading.start()
            
            # 计算可用宽度（考虑边距和居中显示）
            scroll_width = self.original_scroll.width()
            available_width = int(scroll_width * 0.8)  # 使用滚动区域宽度的80%
            
            # 使用PyMuPDF渲染PDF页面
            doc = fitz.open(pdf_path)
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # 计算缩放比例以适应宽度
                zoom = available_width / page.rect.width
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                
                # 将页面转换为QImage
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                
                # 创建页面容器
                page_container = QFrame()
                page_container.setStyleSheet("""
                    QFrame {
                        background-color: white;
                        border: 1px solid #dee2e6;
                        border-radius: 5px;
                        margin: 10px;
                    }
                """)
                
                # 创建页面布局
                page_layout = QVBoxLayout(page_container)
                page_layout.setContentsMargins(20, 20, 20, 20)
                
                # 创建内容容器（用于居中显示）
                content_container = QWidget()
                content_layout = QVBoxLayout(content_container)
                content_layout.setContentsMargins(0, 0, 0, 0)
                
                # 创建标签显示页面
                page_label = QLabel()
                page_label.setPixmap(pixmap)
                page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                content_layout.addWidget(page_label)
                
                # 添加页码
                page_number = QLabel(f"第 {page_num + 1} 页")
                page_number.setAlignment(Qt.AlignmentFlag.AlignCenter)
                page_number.setStyleSheet("""
                    QLabel {
                        color: #6c757d;
                        font-size: 12px;
                        padding: 5px;
                    }
                """)
                content_layout.addWidget(page_number)
                
                # 将内容容器添加到页面布局
                page_layout.addWidget(content_container)
                
                # 创建外部容器用于水平居中
                outer_container = QWidget()
                outer_layout = QHBoxLayout(outer_container)
                outer_layout.setContentsMargins(
                    (scroll_width - available_width - 60) // 2,  # 左边距
                    0,  # 上边距
                    (scroll_width - available_width - 60) // 2,  # 右边距
                    0   # 下边距
                )
                outer_layout.addWidget(page_container)
                
                target_layout.addWidget(outer_container)
            
            # 移除加载指示器
            loading.stop()
            loading.setParent(None)
            
            doc.close()
            
        except Exception as e:
            print(f"显示PDF预��失败: {str(e)}")
            raise
    
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
            
            # 获取格式化后的临时文档路径
            formatted_docx = os.path.join(self.temp_dir, "formatted.docx")
            
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
        available_width = window_width - 40
        doc_width = (available_width - 2) // 2  # 2是分割条的默认宽度
        
        # 更新滚动区域宽度
        self.original_scroll.setFixedWidth(doc_width)
        self.formatted_scroll.setFixedWidth(doc_width)
    
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
            print(f"计算格式哈希值失败: {str(e)}")
            return None
    
    def _preview_content_exists(self):
        """检查预览内容是否已存在"""
        original_pdf = os.path.join(self.temp_dir, "original.pdf")
        formatted_pdf = os.path.join(self.temp_dir, "formatted.pdf")
        return os.path.exists(original_pdf) and os.path.exists(formatted_pdf)
    