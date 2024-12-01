from pathlib import Path
from PIL import Image

def create_default_icon():
    """创建默认的应用图标"""
    icon_size = (256, 256)
    icon = Image.new('RGBA', icon_size, (0, 0, 0, 0))
    
    # 这里可以自定义图标的绘制逻辑
    # 例如：使用 ImageDraw 绘制简单图形或添加文字
    
    icon_path = Path(__file__).parent / "app_icon.ico"
    icon.save(str(icon_path), format="ICO")
    return icon_path

if __name__ == "__main__":
    create_default_icon() 