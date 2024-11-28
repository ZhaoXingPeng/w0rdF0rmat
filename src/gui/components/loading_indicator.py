from PyQt6.QtWidgets import QWidget, QLabel, QVBoxLayout, QFrame
from PyQt6.QtCore import Qt, QTimer, QSize
from PyQt6.QtGui import QPainter, QColor, QPen, QPainterPath

class LoadingIndicator(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.angle = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.rotate)
        
        # 设置固定大小
        self.setFixedSize(120, 120)
        
        # 设置样式
        self.setStyleSheet("""
            QFrame {
                background-color: rgba(0, 0, 0, 0.7);
                border-radius: 10px;
            }
            QLabel {
                color: white;
                font-size: 14px;
                margin-top: 10px;
            }
        """)
        
        # 创建布局
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # 添加文本标签
        self.label = QLabel("正在加载...")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)
        
        # 设置窗口透明
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
    
    def start(self):
        """开始动画"""
        self.show()
        self.timer.start(50)  # 每50毫秒更新一次
    
    def stop(self):
        """停止动画"""
        self.timer.stop()
        self.hide()
    
    def rotate(self):
        """旋转动画"""
        self.angle = (self.angle + 10) % 360
        self.update()
    
    def paintEvent(self, event):
        """绘制加载动画"""
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 绘制背景
        painter.fillRect(self.rect(), QColor(0, 0, 0, 180))
        
        # 设置画笔
        pen = QPen(QColor("#ffffff"))
        pen.setWidth(4)
        painter.setPen(pen)
        
        # 计算中心点和半径
        center = self.rect().center()
        radius = min(center.x(), center.y()) - 20
        
        # 绘制加载动画
        painter.translate(center)
        painter.rotate(self.angle)
        
        for i in range(8):
            painter.rotate(45)
            painter.setOpacity(0.125 * (i + 1))
            painter.drawLine(0, -radius + 10, 0, -radius + 20)
    
    def sizeHint(self):
        return QSize(120, 120)
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 设置画笔
        pen = QPen(QColor("#0078d4"))
        pen.setWidth(3)
        painter.setPen(pen)
        
        # 绘制圆弧
        center = self.rect().center()
        painter.translate(center)
        painter.rotate(self.angle)
        
        for i in range(8):
            painter.rotate(45)
            painter.setOpacity(0.125 * (i + 1))
            painter.drawLine(0, 10, 0, 15) 