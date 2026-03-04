# word-ast

## 项目简介

通过将 docx 转为 AI 友好的 JSON（AI 视图），让用户可以用任意 LLM（ChatGPT、Claude 等）创建或修改 Word 文档，**无需 AI 能运行代码**。

---

## 环境要求与安装

Python >= 3.10，步骤：

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## 场景 A：修改已有文档

### 第一步：导出 AI 视图

```bash
python scripts/ai_edit.py export -I report.docx -O ./out/
```

产出：
- `out/report.ai_view.json` — 发给 LLM
- `out/report.full_ast.json` — 本地留存，**不要发给 LLM**

### 第二步：让 LLM 修改

1. 将 `docs/AI_PROMPT.md` 的内容作为 System Prompt 发给 LLM
2. 将 `out/report.ai_view.json` 的内容 + 修改需求发给 LLM
3. 将 LLM 返回的 JSON 保存为 `modified.json`

### 第三步：渲染输出

```bash
python scripts/ai_edit.py render -V modified.json -S out/report.full_ast.json -O output.docx
```

---

## 场景 B：从零创建文档

### 第一步：让 LLM 创建

1. 将 `docs/AI_PROMPT.md` 的内容作为 System Prompt 发给 LLM
2. 将你的文档需求发给 LLM（纯文本描述即可）
3. 将 LLM 返回的 JSON 保存为 `new_doc.json`

### 第二步：渲染输出

```bash
python scripts/ai_edit.py render -V new_doc.json -O output.docx
```

---

## 命令参考

### export 子命令

| 参数 | 简写 | 说明 |
|------|------|------|
| `--input` | `-I` | 输入 .docx 文件路径 |
| `--outdir` | `-O` | 输出目录（自动生成两个 JSON 文件）|

### render 子命令

| 参数 | 简写 | 说明 |
|------|------|------|
| `--view` | `-V` | AI 视图 JSON 文件路径（必需）|
| `--schema` | `-S` | 保真数据 full_ast JSON（可选，不传=从零创建模式）|
| `--output` | `-O` | 输出 .docx 文件路径 |

---

## 注意事项

- `full_ast.json` 含原始 XML 数据，文件较大，**不要发给 LLM**
- 修改文档时，LLM 返回的 JSON 中 block 的 `id` 必须与原始一致
- 如果 LLM 返回的内容包含说明文字，需手动删除，只保留 JSON 部分

---

## Python API（面向开发者）

```python
from word_ast import parse_docx, render_ast, to_ai_view, merge_ai_edits
```

- `parse_docx(path)` — docx → 完整 AST（含 _raw_*）
- `to_ai_view(ast)` — 完整 AST → AI 视图（去掉 _raw_*）
- `merge_ai_edits(full_ast, ai_view)` — 将 AI 修改合并回完整 AST
- `render_ast(ast, output_path)` — AST → docx

---

## Web 前端（最小可运行页面）

`web/` 目录包含一个无需构建工具、可直接在浏览器中运行的最小前端页面，提供：

- **API 配置持久化**：在模态框中填写 API Key / Provider / Base URL / Model 后点击「保存」，配置会写入 `localStorage`（键前缀 `webaiword:v1:`），刷新页面后自动填充。
- **文档编辑状态持久化**：在编辑区修改 `ai_view` JSON 后，内容会在 800ms 内自动保存到 `localStorage`，刷新页面后出现「恢复上次编辑」提示条，一键恢复。
- **清除本地缓存**：工具栏提供「🗑️ 清除本地缓存」按钮，一键删除所有 `webaiword:v1:` 前缀的键；也可在恢复提示条中点击「忽略并清除」仅清除文档缓存。
- **安全提示**：API Key 配置模态框内注明"仅保存在本机 localStorage，不会上传"。

### 启动方式

```bash
cd web
python -m http.server 8080
# 然后在浏览器打开 http://localhost:8080
```

> **注**：由于使用了 ES 模块（`type="module"`），需要通过 HTTP 服务器访问，不能直接以 `file://` 协议打开。

### 工作流程

1. 用 Python 命令行导出 `ai_view.json`（见「场景 A」）。
2. 在浏览器中点击「📂 导入 ai_view」，选择生成的 JSON 文件——内容会自动保存到 `localStorage`。
3. 在编辑区修改 JSON，改动在 800ms 内自动持久化。
4. 刷新页面后，页面顶部出现提示条，点击「恢复上次编辑」即可继续。

---

## 运行测试

```bash
pytest
```
