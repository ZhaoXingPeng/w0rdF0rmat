from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    # 创建一个 256x256 的图像
    size = 256
    image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    
    # 绘制圆形背景
    margin = 10
    draw.ellipse([margin, margin, size-margin, size-margin], 
                 fill='#0078d4')  # 使用蓝色背景
    
    # 添加文字
    try:
        font = ImageFont.truetype("arial.ttf", 100)
    except:
        font = ImageFont.load_default()
    
    text = "w0"
    # 获取文本大小
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    
    # 计算文本位置使其居中
    x = (size - text_width) // 2
    y = (size - text_height) // 2
    
    # 绘制文字
    draw.text((x, y), text, fill='white', font=font)
    
    # 保存为ICO文件
    icon_path = os.path.join(os.path.dirname(__file__), 'app_icon.ico')
    image.save(icon_path, format='ICO')

if __name__ == '__main__':
    create_icon() 