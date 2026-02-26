# word-ast

Word ↔ AST round-trip prototype implementation.

## 项目简介

`word-ast` 是一个用于 **DOCX 与 AST（抽象语法树）相互转换** 的原型工具，支持：

- 将 `.docx` 解析为结构化 AST。
- 将 AST 渲染回 `.docx`。
- 基础段落、文本样式与表格（含合并单元格）往返。

---

## 环境要求

- Python `>= 3.10`
- 建议使用虚拟环境（`venv`）

---

## 安装与准备

### 1) 创建并激活虚拟环境

```bash
python -m venv .venv
source .venv/bin/activate
```

### 2) 安装依赖

```bash
pip install -r requirements.txt
```

如果你希望以包形式安装当前项目（可选）：

```bash
pip install -e .
```

---

## 快速使用指南

本项目提供命令行脚本：`scripts/convert.py`。

### A. DOCX -> AST

```bash
python scripts/convert.py parse <输入docx路径> --output-dir <输出目录>
```

示例：

```bash
python scripts/convert.py parse ./examples/input.docx --output-dir ./out
```

### B. AST -> DOCX

```bash
python scripts/convert.py render <输入AST文件路径> --output <输出docx路径>
```

示例：

```bash
python scripts/convert.py render ./out/ast.json --output ./out/rebuilt.docx
```

---

## Python API 使用

你也可以在代码中直接调用：

```python
from word_ast import parse_docx, render_ast

# 解析 docx
ast = parse_docx("input.docx")

# 渲染回 docx
render_ast(ast, "rebuilt.docx")
```

---

## 运行测试

```bash
pytest
```

---

## 当前能力边界（原型阶段）

- 已覆盖：
  - 文本段落与基础 run 样式（如加粗）
  - 表格结构与部分合并单元格跨度
- 尚未声明完整覆盖：
  - 复杂分页、页眉页脚、批注、修订记录等高级 Word 特性

建议在接入生产场景前，使用你的目标文档样本进行回归验证。
