from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

def create_default_icon():
    """创建默认的应用图标"""
    # 创建一个256x256的图像
    icon_size = (256, 256)
    icon = Image.new('RGBA', icon_size, (255, 255, 255, 0))  # 透明背景
    draw = ImageDraw.Draw(icon)
    
    # 绘制一个蓝色圆形背景
    padding = 20
    draw.ellipse(
        [padding, padding, icon_size[0]-padding, icon_size[1]-padding],
        fill='#2196F3'  # 使用Material Design蓝色
    )
    
    # 添加文字
    try:
        # 尝试使用系统字体
        font_size = 120
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", font_size)
    except:
        # 如果找不到系统字体，使用默认字体
        font = ImageFont.load_default()
    
    # 添加文字 "AI"
    text = "AI"
    # 获取文本边界框
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    
    # 计算文字居中位置
    x = (icon_size[0] - text_width) // 2
    y = (icon_size[1] - text_height) // 2
    
    # 绘制白色文字
    draw.text((x, y), text, fill='white', font=font)
    
    # 保存图标
    icon_path = Path(__file__).parent / "app_icon.ico"
    # 保存为多尺寸的ICO文件
    sizes = [(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)]
    icons = []
    for size in sizes:
        icons.append(icon.resize(size, Image.Resampling.LANCZOS))
    
    icons[0].save(
        str(icon_path),
        format='ICO',
        sizes=sizes,
        append_images=icons[1:]
    )
    return icon_path

if __name__ == "__main__":
    create_default_icon() 