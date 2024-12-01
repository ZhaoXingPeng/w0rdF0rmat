from PyQt6.QtWidgets import QWidget, QLabel, QVBoxLayout, QFrame
from PyQt6.QtCore import Qt, QTimer, QSize, QRect
from PyQt6.QtGui import QPainter, QColor, QPen

class LoadingIndicator(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.angle = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.rotate)
        self.setFixedSize(40, 40)  # 只显示旋转动画，不显示文字
        
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 绘制旋转的圆弧
        pen = QPen()
        pen.setWidth(3)
        pen.setColor(QColor("#0078d4"))
        painter.setPen(pen)
        
        rect = QRect(5, 5, 30, 30)  # 调整大小以适应新的尺寸
        painter.drawArc(rect, self.angle * 16, 300 * 16)
    
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
        self.update()  # 触发重绘
    
    def sizeHint(self):
        return QSize(40, 40) 