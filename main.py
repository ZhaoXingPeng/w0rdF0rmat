from src.core.formatter import WordFormatter
from src.core.document import Document

def main():
    # 示例用法
    doc = Document("input.docx")
    formatter = WordFormatter(doc)
    formatter.format()
    doc.save("output.docx")

if __name__ == "__main__":
    main() 