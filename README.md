# word-ast

一个"沙漏型"中间层，让 AI 通过语义化 AST JSON 操作 Word 文档，而不需要理解 OOXML。

## 项目简介

`word-ast` 是一个 **DOCX ↔ AST（抽象语法树）双向转换**工具，带有 AI 编辑层。支持：

- 将 `.docx` 解析为结构化 AST（JSON 格式），保留完整格式信息
- 将 AST 渲染回 `.docx`（round-trip）
- 段落、标题、表格（含合并单元格）、图片内联、TOC、样式继承
- AI 编辑层：生成干净的 AI 视图，将 AI 修改合并回保真 AST

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

---

## Python API 快速使用

```python
from word_ast import parse_docx, render_ast, to_ai_view, merge_ai_edits
```

### `parse_docx(docx_path, output_dir=None) -> dict`

将 `.docx` 文件解析为 AST 字典。

```python
ast = parse_docx("input.docx")
```

`output_dir` 可选，若指定则同时将 AST 和提取的图片资源保存到该目录。

### `render_ast(ast, output_path)`

将 AST 渲染回 `.docx` 文件。

```python
render_ast(ast, "rebuilt.docx")
```

### `to_ai_view(ast) -> dict`

将完整 AST（含内部 `_raw_*` XML 字段）转换为适合 AI 查看和修改的精简视图（去掉所有 `_raw_*` 字段），返回深拷贝。

```python
ai_view = to_ai_view(ast)
# 将 ai_view 发给 AI，AI 返回修改后的 JSON
```

### `merge_ai_edits(original_ast, ai_ast) -> dict`

将 AI 修改后的视图合并回含 `_raw_*` 的完整 AST，同步 AI 修改到底层 XML，以保留格式保真度。

```python
merged_ast = merge_ai_edits(ast, ai_modified)
render_ast(merged_ast, "output.docx")
```

支持合并的字段：
- 段落格式：`alignment`、`indent_left`、`indent_right`、`indent_first_line`、`space_before`、`space_after`
- run 格式：`font_ascii`、`font_east_asia`、`size`、`bold`、`italic`、`color`、`text`
- 目前仅 Paragraph 块支持合并，Table 块暂不支持

---

## AI 编辑 Pipeline

`scripts/ai_edit.py` 提供完整的 5 步 AI 编辑 CLI：

```
1. Parse:     docx → full AST（含 _raw_* XML）
2. AI view:   full AST → to_ai_view() → 干净 JSON（无 _raw_*）
3. AI edits:  将视图发给 AI，获取修改后的 JSON
4. Merge:     merge_ai_edits(full AST, modified JSON) → merged AST
5. Render:    merged AST → docx
```

### 导出 AI 视图

```bash
python scripts/ai_edit.py --input doc.docx --ai-view-output view.json
```

### 应用 AI 修改

```bash
python scripts/ai_edit.py --input doc.docx --ai-edit-input modified.json --output out.docx
```

### 直接 round-trip（不经 AI 修改）

```bash
python scripts/ai_edit.py --input doc.docx --output out.docx
```

---

## 运行测试

```bash
pytest
```
