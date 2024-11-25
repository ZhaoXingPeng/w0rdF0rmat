from src.config.config_manager import ConfigManager
from src.core.formatter import WordFormatter
from src.core.document import Document

def main():
    try:
        # 加载配置
        config_manager = ConfigManager()
        
        # 读取测试文档
        doc = Document("./src/test/test.docx")
        
        # 创建格式化器并使用默认格式
        formatter = WordFormatter(doc, config_manager)
        
        # 应用默认格式
        formatter.format()
        
        # 保存格式化后的文档
        doc.save("./src/test/output.docx")
        print("文档格式化完成，已保存为 output.docx")
        
    except FileNotFoundError:
        print("错误：找不到测试文档 test.docx")
    except Exception as e:
        print(f"错误：格式化过程中出现异常 - {str(e)}")

if __name__ == "__main__":
    main() 