# Word文档智能模板生成系统

本项目用于自动化地将Word文档中的内容（段落、表格、图片等）提取、语义匹配并替换为可复用的模板格式。系统集成了文档结构解析、内容与关键字描述的智能匹配（支持本地/远程大语言模型LLM）、以及模板文档的自动生成，适用于报告自动化、批量文档模板化等场景。

---

## 主要功能

1. **文档内容提取**  
   - 自动提取Word文档中的段落、表格、图片等元素，分别保存为独立文件。
2. **智能语义匹配**  
   - 利用大语言模型（本地或远程API）对提取的内容与关键字描述文件进行语义匹配，自动识别内容对应的业务字段。
3. **模板生成**  
   - 根据匹配结果，将原文档中的实际内容替换为占位符，生成可复用的Word模板。

---

## 目录结构

```
word_to_template_2/
│
├── main.py                      # 主程序，整合提取、匹配、替换流程
├── requirements.txt             # 依赖包列表
├── extractors/                  # 文档内容提取模块
│   ├── extractor.py             # 提取调度器
│   ├── paragraph_extractor.py   # 段落提取
│   ├── table_extractor.py       # 表格提取
│   └── image_extractor.py       # 图片提取
├── matchers/                    # 智能匹配模块
│   ├── matcher.py               # 匹配调度器
│   ├── table_matcher.py         # 表格匹配
│   ├── image_matcher.py         # 图片匹配
│   ├── table_system_prompt.txt  # 表格匹配LLM提示词
│   └── image_system_prompt.md   # 图片匹配LLM提示词
├── replacers/                   # 模板生成与内容替换模块
│   ├── replacer.py              # 替换调度器
│   ├── paragraph_replacer.py    # 段落替换
│   └── table_replacer.py        # 表格替换
├── models/                      # LLM模型管理与调用
│   ├── model_manager.py         # 本地/远程模型统一接口
│   └── Qwen3-0.6B-Q8_0.gguf     # 示例本地模型文件
├── document/                    # 示例文档及中间结果
│   ├── document.docx            # 原始Word文档
│   ├── document_parts/          # 拆分后的文档元素
│   ├── match_results/           # 匹配结果
│   └── key_descriptions/        # 关键字描述文件
└── ...
```

---

## 依赖环境

- Python 3.8+
- 主要依赖包见 `requirements.txt`，如：
  - `python-docx`：Word文档解析
  - `llama-cpp-python`：本地LLM推理（可选）
  - `openai`：远程API调用（可选）

---

## 使用方法

1. **准备文档与关键字描述文件**  
   - 将待处理的Word文档放入 `document/document.docx`
   - 在 `document/key_descriptions/` 下准备表格、图片等关键字描述文件

2. **运行主程序**  
   ```bash
   python main.py
   ```
   - 程序将自动完成文档内容提取、智能匹配、模板生成等步骤
   - 生成的模板文档保存在 `document/template.docx`

3. **自定义与扩展**  
   - 可根据实际业务需求，扩展 `extractors/`、`matchers/`、`replacers/` 下的功能模块
   - 支持自定义LLM模型或API

---

## 典型流程

1. **提取**：将Word文档拆分为段落、表格、图片等独立文件
2. **匹配**：用LLM对每个元素与关键字描述进行语义匹配，输出匹配结果
3. **替换**：根据匹配结果，将文档内容替换为占位符，生成标准化模板

---

## 适用场景

- 智能报告模板生成
- 批量文档结构化处理
- Word文档内容自动归档与字段提取

---

## 联系与反馈

如有问题或建议，请在 `document/问题记录.txt` 中记录，或通过Issue反馈。

---

如需更详细的开发文档或二次开发指导，请查阅各模块源码及注释。
