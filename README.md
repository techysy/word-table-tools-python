# 📝 Word 表格批量工具 | Word Table Batch Tool

批量替换和格式化 Word 文档中的表格内容，让繁琐的文档处理变得轻松高效！

Batch replace and format table content in Word documents, making tedious document processing easy and efficient!

---

## ✨ 功能 | Features

- 🔄 **内容替换** - 批量替换表格中的指定文字（支持精确匹配和包含匹配）
- 🎨 **格式刷** - 从模板文件复制表格格式到其他文件
- 📋 **日志记录** - 详细记录每次操作，方便排查问题

- 🔄 **Content Replacement** - Batch replace specified text in tables (supports exact and contains matching)
- 🎨 **Format Brush** - Copy table format from template files to other files
- 📋 **Logging** - Detailed operation logs for easy troubleshooting

---

## 🚀 使用方法 | Usage

### 1. 🔄 内容替换工具 | Content Replacement Tool

双击 `启动替换工具.bat` 运行

Double-click `启动替换工具.bat` to run

1. 在"查找"和"替换"输入框填写规则 | Enter find and replace rules
2. 拖拽 Word 文件到窗口，或点击添加文件 | Drag Word files to window or click to add
3. 点击"执行替换" | Click "Execute Replacement"

### 2. 🎨 格式刷工具 | Format Brush Tool

双击 `启动格式刷.bat` 运行

Double-click `启动格式刷.bat` to run

1. 选择一个格式正确的模板文件 | Select a template file with correct format
2. 添加需要修改格式的目标文件 | Add target files that need format changes
3. 点击"执行格式复制" | Click "Execute Format Copy"

---

## 🎯 匹配规则 | Matching Rules

- 🔍 **包含匹配** - 单元格内容包含查找文字即替换
- 🛡️ **自动排除表头** - 长文本（>50字符）自动跳过，避免误替换
- ⏭️ **跳过关键字** - 输入标题关键字（逗号分隔）可跳过这些标题不替换

- 🔍 **Contains Matching** - Replace if cell content contains the search text
- 🛡️ **Auto Exclude Headers** - Long text (>50 chars) auto-skip to avoid wrong replacements
- ⏭️ **Skip Keywords** - Enter header keywords (comma-separated) to skip these headers from replacement

---

## 📁 文件说明 | File Description

| 文件 | 说明 | Description |
|------|------|-------------|
| `word_replacer_tool.py` | 🔄 内容替换工具 | Content Replacement Tool |
| `format_brush.py` | 🎨 格式刷工具 | Format Brush Tool |
| `logs/` | 📋 操作日志 | Operation Logs |

---

## ⚙️ 环境要求 | Requirements

- 🐍 Python 3.8+
- 📦 依赖库 | Dependencies: `python-docx`, `ttkbootstrap`, `lxml`, `windnd`

```bash
pip install python-docx ttkbootstrap lxml windnd
```

---

## 📜 许可证 | License

MIT License
