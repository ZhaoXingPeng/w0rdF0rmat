from src.core.formatter import WordFormatter
from src.core.document import Document

def main():
    # 示例用法
    doc = Document("./src/test/test.docx")
    formatter = WordFormatter(doc)
    
    # 使用预设格式
    formatter.format()
    
    # 或者使用用户提供的格式要求
    user_requirements = """
    标题要求：
    - 三号字（16pt）
    - 黑体
    - 居中
    
    正文要求：
    - 小四号字（12pt）
    - 首行缩进2字符
    - 1.5倍行距
    
    参考文献：
    - 五号字（10.5pt）
    - 悬挂缩进
    """
    formatter.apply_user_requirements(user_requirements)
    
    doc.save("output.docx")

if __name__ == "__main__":
    main() 