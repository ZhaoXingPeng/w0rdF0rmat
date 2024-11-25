# Word Document Formatting Tool  
# Word文档格式化工具

A Python-based automation tool designed to streamline formatting academic papers, with intelligent document structure recognition and customizable formatting options.  
一个用于自动化处理学术论文格式的Python工具，支持智能识别文档结构并应用指定格式要求。

---

## Features  
## 功能特点

- 🔍 **Intelligent Document Structure Recognition**  
  Identifies titles, abstracts, keywords, sections, and more.  
  **智能识别文档结构**（标题、摘要、关键词、章节等）

- 📝 **Supports Multiple Parsing Methods**  
  Includes style parsing, traditional parsing, and optional AI assistance.  
  **支持多种解析方式**（样式解析、传统解析、AI辅助）

- 🎨 **Flexible Formatting Configuration**  
  Define your own formatting rules with ease.  
  **灵活的格式规范配置**

- 🤖 **AI-Assisted Analysis (Optional)**  
  Provides formatting suggestions using AI (e.g., GPT-3.5-Turbo).  
  **可选的AI辅助分析**，提供智能格式建议（如GPT-3.5-Turbo）

- 📋 **Preset Formatting Templates**  
  Includes default templates for common styles.  
  **预设格式模板支持**

- 🔧 **Customizable Format Requirements**  
  Adjust settings to fit specific style guidelines.  
  **支持自定义格式要求**

---

## Project Structure  
## 项目结构

```
src/
├── __init__.py
├── config/
│   ├── __init__.py
│   ├── config.yaml               # Global configuration file 全局配置文件
│   └── config_manager.py         # Configuration manager 配置管理器
├── core/
│   ├── __init__.py
│   ├── ai_assistant.py           # AI assistance functions AI辅助功能
│   ├── document.py               # Core document processing 文档处理核心
│   ├── format_spec.py            # Formatting specifications 格式规范定义
│   ├── formatter.py              # Formatting implementations 格式化实现
│   └── presets/
│       └── default.yaml          # Default formatting template 默认格式模板
└── test/
    └── test.docx                 # Sample document for testing 测试用文档
```

---

## Installation and Setup  
## 安装与设置

1. **Install Dependencies**  
   安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the Application**  
   启动程序：
   ```bash
   python main.py
   ```

---

## Usage Instructions  
## 使用说明

### Import Key Modules  
### 导入核心模块
```python
from src.config.config_manager import ConfigManager
from src.core.formatter import WordFormatter
from src.core.document import Document
```

### Initialize Components  
### 初始化组件
```python
# Load configuration 加载配置
config_manager = ConfigManager()

# Load the document 加载文档
doc = Document("path/to/your/document.docx")

# Initialize the formatter 初始化格式化工具
formatter = WordFormatter(doc, config_manager)
```

### Apply Formatting  
### 应用格式
```python
formatter.format()
```

### Save the Document  
### 保存文档
```python
doc.save("output.docx")
```

---

## Configuration Guide  
## 配置指南

The primary configuration file is located at `src/config/config.yaml`. Below is an example configuration:  
主要配置文件位于 `src/config/config.yaml`。以下是示例配置：

```yaml
ai_assistant:
  enabled: false                  # Enable or disable AI assistance 是否启用AI辅助
  model: "gpt-3.5-turbo"          # Specify the AI model 使用的AI模型

formatting:
  use_default_template: true      # Use the default formatting template 是否使用默认模板
  template_path: "src/core/presets/default.yaml"  # Path to the template 模板路径
```

You can modify these options to tailor the tool to your requirements.  
可以根据需求修改这些选项。

---

## Contribution Guidelines  
## 贡献指南

1. Fork the repository and create your feature branch (`git checkout -b feature/AmazingFeature`).  
   Fork项目并创建功能分支（`git checkout -b feature/AmazingFeature`）。
2. Commit your changes (`git commit -m 'Add some AmazingFeature'`).  
   提交更改（`git commit -m 'Add some AmazingFeature'`）。
3. Push to the branch (`git push origin feature/AmazingFeature`).  
   推送分支（`git push origin feature/AmazingFeature`）。
4. Open a pull request.  
   发起Pull Request。

---

## Notes  
## 注意事项

- Ensure Python 3.8+ is installed.  
  确保已安装Python 3.8或更高版本。
- When using AI-assisted features, ensure your API credentials for OpenAI are correctly configured.  
  使用AI辅助功能时，请确保正确配置了OpenAI的API凭据。
- Contributions and bug reports are welcome!  
  欢迎贡献代码和提交问题反馈！