# LocalAItable

<div align="right">
  <a href="#chinese">中文</a> | <a href="#english">English</a>
</div>

---

<a name="chinese"></a>
# LocalAItable - 本地AI表格处理工具

LocalAItable是一个强大的本地化AI表格处理工具，允许您通过本地大模型或云端API批量处理Excel/CSV表格数据，实现类似"多维表格"的智能化数据处理能力。

![项目标识](https://via.placeholder.com/800x400?text=LocalAItable)

## 🌟 项目特点

- **双模式AI支持**：同时支持OpenAI API和本地部署的Ollama模型
- **表格数据处理**：轻松导入/导出Excel和CSV文件，自动检测文件编码
- **批量AI生成**：为表格中的数据批量生成AI内容，支持多线程并行处理
- **模板系统**：强大的提示词模板管理，支持变量替换和条件逻辑
- **友好界面**：直观的图形用户界面，无需编程经验即可操作
- **完全本地化**：使用本地模型时，所有数据处理均在本地完成，保护数据隐私

## 🚀 应用场景

- **文本摘要生成**：批量将长文本内容转化为简洁摘要
- **数据提取与解析**：从非结构化文本中提取结构化数据(如血压、日期等)
- **内容翻译**：批量翻译表格中的文本内容
- **情感分析与分类**：分析文本情感倾向或进行内容分类
- **关键词提取**：从大量文本中提取关键词和核心概念
- **医疗数据处理**：提取和整理医疗记录中的关键数据

## 📋 系统要求

- Python 3.8或更高版本
- 本地运行Ollama模型推荐8GB以上内存
- 支持Windows、macOS和Linux系统

## 🔧 安装指南

1. 克隆仓库到本地
```bash
git clone https://github.com/yourusername/LocalAItable.git
cd LocalAItable
```

2. 安装依赖包
```bash
pip install -r requirements.txt
```

3. (可选) 设置OpenAI API密钥
   - 在程序界面中设置
   - 或设置环境变量 `OPENAI_API_KEY`

4. (可选) 安装并配置Ollama
   - 从[Ollama官网](https://ollama.ai/)下载并安装
   - 下载所需模型，如 `ollama pull deepseek-r1:14b`

## 📊 使用方法

1. 运行应用程序
```bash
python ai_column_generator.py
```

2. 导入数据
   - 点击"选择Excel/CSV文件"按钮
   - 如遇编码问题，可使用"手动指定编码打开"功能

3. 配置AI
   - 选择API类型(OpenAI或Ollama)
   - 配置相应API密钥或URL地址
   - 选择合适的AI模型

4. 选择处理列
   - 指定要处理的表格列(引用列)
   - 指定AI生成内容的保存列(目标列)

5. 设置提示词
   - 使用内置模板或创建自定义模板
   - 支持变量替换和条件逻辑

6. 生成内容
   - 点击"预览"按钮测试效果
   - 点击"生成并更新"按钮批量处理
   - 处理完成后，可导出更新后的表格文件

## 📝 模板示例

基础模板示例：
```
请根据以下内容生成一段简洁的摘要：

{引用内容}
```

条件逻辑模板：
```
请分析以下内容，{如果:关键词:重点关注这些关键词: {关键词}
}

{引用内容}
```

## 📜 许可证

本项目基于MIT许可证开源 - 详见 [LICENSE](LICENSE) 文件

## 🤝 贡献

欢迎提交问题和功能建议！如果您想贡献代码，请先fork仓库并创建拉取请求。

## 📞 联系方式

如有问题或建议，请通过GitHub Issues与我们联系。

---

<a name="english"></a>
# LocalAItable - Local AI Spreadsheet Processor

LocalAItable is a powerful local AI spreadsheet processing tool that allows you to batch process Excel/CSV spreadsheet data through local large language models or cloud APIs, achieving intelligent data processing capabilities similar to "multi-dimensional tables".

![Project Logo](https://via.placeholder.com/800x400?text=LocalAItable)

## 🌟 Features

- **Dual AI Support**: Supports both OpenAI API and locally deployed Ollama models
- **Spreadsheet Processing**: Easily import/export Excel and CSV files with automatic encoding detection
- **Batch AI Generation**: Generate AI content for spreadsheet data in batch with multi-threading support
- **Template System**: Powerful prompt template management with variable substitution and conditional logic
- **User-Friendly Interface**: Intuitive graphical user interface requiring no programming experience
- **Fully Localized**: When using local models, all data processing is done locally to protect data privacy

## 🚀 Use Cases

- **Text Summarization**: Batch convert long text content into concise summaries
- **Data Extraction & Parsing**: Extract structured data from unstructured text (e.g., blood pressure, dates)
- **Content Translation**: Batch translate text content in spreadsheets
- **Sentiment Analysis & Classification**: Analyze text sentiment or classify content
- **Keyword Extraction**: Extract keywords and core concepts from large volumes of text
- **Medical Data Processing**: Extract and organize key data from medical records

## 📋 System Requirements

- Python 3.8 or higher
- 8GB+ RAM recommended for running Ollama models locally
- Supports Windows, macOS, and Linux systems

## 🔧 Installation Guide

1. Clone the repository
```bash
git clone https://github.com/yourusername/LocalAItable.git
cd LocalAItable
```

2. Install dependencies
```bash
pip install -r requirements.txt
```

3. (Optional) Set up OpenAI API key
   - Configure in the program interface
   - Or set the environment variable `OPENAI_API_KEY`

4. (Optional) Install and configure Ollama
   - Download and install from [Ollama website](https://ollama.ai/)
   - Download required models, e.g., `ollama pull deepseek-r1:14b`

## 📊 How to Use

1. Run the application
```bash
python ai_column_generator.py
```

2. Import data
   - Click the "Select Excel/CSV File" button
   - For encoding issues, use the "Open with Manual Encoding" feature

3. Configure AI
   - Select API type (OpenAI or Ollama)
   - Configure corresponding API key or URL
   - Choose an appropriate AI model

4. Select processing columns
   - Specify the spreadsheet columns to process (reference columns)
   - Specify the column to save AI-generated content (target column)

5. Set up prompts
   - Use built-in templates or create custom templates
   - Support variable substitution and conditional logic

6. Generate content
   - Click the "Preview" button to test the effect
   - Click "Generate and Update" for batch processing
   - After processing, export the updated spreadsheet file

## 📝 Template Examples

Basic template example:
```
Please generate a concise summary based on the following content:

{引用内容}
```

Conditional logic template:
```
Please analyze the following content, {如果:关键词:with special attention to these keywords: {关键词}
}

{引用内容}
```

## 📜 License

This project is open-sourced under the MIT License - see the [LICENSE](LICENSE) file for details

## 🤝 Contribution

Issues and feature suggestions are welcome! If you'd like to contribute code, please fork the repository and create a pull request.

## 📞 Contact

For questions or suggestions, please contact us through GitHub Issues. 
