# Word ↔ AST Round-Trip 实现规格文档

## 1. 项目背景与目标

### 1.1 为什么要做这个

现有 AI 操作 Word 文档的方式是通过 MCP 工具或 Skills 间接调用 python-docx/win32com 等库，存在两个根本性问题：

1. **语义鸿沟**：AI 看到的是 OOXML 的 XML 结构，或者 python-docx 的对象 API，这两者都不是为 AI 理解设计的，字段命名晦涩、结构嵌套复杂、格式继承关系隐式。
2. **操作原子性差**：AI 每次操作都要调用多个工具，没有事务保障，出错难以回滚。

**解决思路**：设计一个"沙漏型"中间层。

```
原始 .docx
    │
    ▼  [解析层]
含 _raw_* 的完整 AST（JSON 格式）
    │
    ▼  [AI 视图层]
干净的 AI 视图（去掉 _raw_*）
    │
    ▼  [AI 修改]
AI 修改后的视图
    │
    ▼  [Merge 层]
合并回完整 AST
    │
    ▼  [渲染层]
新的 .docx
```

AI 只需要理解和操作 AST，不需要了解 OOXML 细节。

### 1.2 已实现成果

**阶段一（已完成）：DOCX ↔ AST round-trip**

`原始 .docx → AST → 重建 .docx`，重建后的文档在普通用户可感知的维度上与原始文档一致，支持：
- 文本段落与 run 级字符格式（粗体、斜体、颜色、字号、字体）
- 标题（通过 style 字段区分层级）
- 表格（含横向/纵向合并单元格）
- 图片内联（base64 编码保存在 AST 中）
- 目录（TOC）
- 样式库与样式继承

**阶段二（已完成）：AI 编辑层**

- `to_ai_view()` — 将完整 AST 转换为去掉 `_raw_*` 字段的干净 AI 视图
- `merge_ai_edits()` — 将 AI 修改的视图合并回完整 AST，同步更新底层 XML

### 1.3 成功标准

1. 给定一批测试 docx 文件，转换后的文档在以下维度与原文件一致：
   - 文字内容完整无丢失
   - 段落结构（标题层级）正确
   - 基础字符格式（粗体、斜体、颜色、字号、字体）正确
   - 基础段落格式（对齐、缩进、间距）正确
   - 表格结构（行列数、合并单元格）正确
   - 图片位置和尺寸正确
2. 不支持的特性通过 passthrough 机制保留，不导致文档损坏
3. AST 是合法 JSON，可被人类阅读和理解

---

## 2. 技术选型

| 组件 | 选型 | 理由 |
|------|------|------|
| 主解析库 | `python-docx` | 直接操作 OOXML，信息最完整 |
| 辅助解析 | `lxml` | 处理 python-docx 未封装的底层 XML |
| AST 格式 | JSON | AI 友好，人类可读 |
| 图片提取 | base64 内嵌 | 直接嵌入 AST，无需额外文件 |
| 测试框架 | `pytest` | 标准选择 |
| 语言 | Python 3.10+ | 类型注解支持好 |

**不使用 Pandoc**：Pandoc 在样式、格式信息上丢失过多，本项目直接用 python-docx 从 OOXML 层面解析，自己控制信息保留的完整性。

---

## 3. 仓库结构

```
AIWord/
├── README.md
├── requirements.txt
├── pytest.ini
│
├── word_ast/                    # 核心库
│   ├── __init__.py              # 公开 API：parse_docx, render_ast, to_ai_view, merge_ai_edits
│   ├── schema.py                # AST 数据结构定义
│   ├── ai_view.py               # AI 视图层：to_ai_view()
│   ├── ai_merge.py              # AI Merge 层：merge_ai_edits()
│   ├── parser/
│   │   ├── __init__.py
│   │   ├── document_parser.py   # 顶层解析入口
│   │   ├── paragraph_parser.py  # 段落/Text run 解析
│   │   ├── table_parser.py      # 表格解析
│   │   └── style_parser.py      # 样式库解析
│   ├── renderer/
│   │   ├── __init__.py
│   │   ├── document_renderer.py # 顶层渲染入口
│   │   ├── paragraph_renderer.py
│   │   ├── table_renderer.py
│   │   ├── toc_renderer.py      # TOC 渲染
│   │   └── style_renderer.py
│   └── utils/
│       └── units.py             # 单位转换工具
│
├── tests/
│   ├── word/                    # 测试用 docx 文件
│   ├── test_parser.py
│   ├── test_roundtrip.py
│   └── test_ai_view_merge.py    # AI 层测试
│
└── scripts/
    ├── ai_edit.py               # AI 编辑 5 步 pipeline CLI
    └── convert.py               # 简单 docx↔AST 转换 CLI
```

---

## 4. AST Schema 完整定义

### 4.1 单位约定

**所有长度单位统一使用 twip（缇）**。

```
1 英寸 = 1440 twip
1 厘米 = 567 twip
1 磅(pt) = 20 twip
12pt 字号 = 240 twip（但字号字段用"半点"，12pt = 24）
```

字号单独使用**半点（half-point）**单位，原因是 OOXML 原生如此存储，避免转换精度损失。字段名为 `size`。

颜色统一使用 `#RRGGBB` 十六进制字符串，或 `"auto"` 表示自动颜色。

### 4.2 顶层结构

```json
{
  "schema_version": "1.0",
  "document": {
    "meta": { },
    "styles": { },
    "body": [ ],
    "passthrough": { }
  }
}
```

### 4.3 meta 块

```json
"meta": {
  "page": {
    "size": "A4",
    "orientation": "portrait",
    "margin": {
      "top": 1440,
      "bottom": 1440,
      "left": 1800,
      "right": 1800
    }
  },
  "default_style": "Normal",
  "language": "zh-CN"
}
```

字段说明：
- `size`: `"A4"` | `"Letter"` | `"A3"` | `"custom"`，当为 `custom` 时需要额外字段 `width` 和 `height`（twip）
- `orientation`: `"portrait"` | `"landscape"`
- `margin`: 四边页边距，单位 twip
- `default_style`: 文档默认段落样式名，通常是 `"Normal"`
- `language`: BCP 47 语言标签

### 4.4 styles 块

样式库，Key 为 `style_id`（与 Word 内部 XML 属性 `w:styleId` 一致）。

```json
"styles": {
  "Normal": {
    "style_id": "Normal",
    "name": "Normal",
    "type": "paragraph",
    "based_on": null,
    "paragraph_format": {
      "alignment": "left",
      "indent_left": 0,
      "indent_right": 0,
      "indent_first_line": 0,
      "space_before": 0,
      "space_after": 160,
      "line_spacing": {
        "rule": "auto",
        "value": 240
      },
      "keep_with_next": false,
      "keep_lines_together": false,
      "page_break_before": false
    },
    "character_format": {
      "font_ascii": "Calibri",
      "font_east_asia": "宋体",
      "size": 24,
      "bold": false,
      "italic": false,
      "underline": "none",
      "strike": false,
      "color": "auto",
      "highlight": null,
      "vertical_align": "baseline"
    }
  },
  "Heading1": {
    "style_id": "Heading1",
    "name": "heading 1",
    "type": "paragraph",
    "based_on": "Normal",
    "paragraph_format": {
      "space_before": 240,
      "space_after": 120,
      "keep_with_next": true
    },
    "character_format": {
      "size": 32,
      "bold": true,
      "color": "#2E74B5"
    }
  },
  "DefaultParagraphFont": {
    "style_id": "DefaultParagraphFont",
    "name": "Default Paragraph Font",
    "type": "character",
    "based_on": null,
    "character_format": {}
  }
}
```

字段说明：
- `type`: `"paragraph"` | `"character"` | `"table"` | `"numbering"`
- `based_on`: 父样式的 `style_id`，`null` 表示无继承
- `paragraph_format` 和 `character_format` 只存储与父样式**不同的字段**（差量存储）
- `alignment`: `"left"` | `"right"` | `"center"` | `"justify"`
- `line_spacing.rule`: `"auto"` | `"exact"` | `"atLeast"`；`auto` 时 value 为行距倍数×240（单倍=240，1.5 倍=360，双倍=480）；`exact`/`atLeast` 时 value 为 twip
- `size`: 字号，单位半点（half-point），12pt = 24
- `underline`: `"none"` | `"single"` | `"double"` | `"dotted"` | `"dashed"` | `"wavy"`
- `vertical_align`: `"baseline"` | `"superscript"` | `"subscript"`
- `highlight`: `null` 或 Word 高亮颜色名

### 4.5 body 块

`body` 是 Block 节点的数组，顺序与文档内容顺序一致。

#### 节点通用字段

所有 Block 节点都有 `id` 和 `type` 字段：

```json
{
  "id": "p0",
  "type": "Paragraph"
}
```

`id` 规则：
- 段落：`p0`, `p1`, `p2` ... 按顺序编号
- 表格：`t0`，表格单元格内段落：`t0.r0c0.p0`（点号分隔）

#### 4.5.1 Paragraph 节点

```json
{
  "id": "p0",
  "type": "Paragraph",
  "style": "Heading1",
  "paragraph_format": {
    "alignment": "center",
    "indent_left": 0,
    "indent_first_line": 420,
    "space_before": 240,
    "space_after": 120
  },
  "content": [
    {
      "type": "Text",
      "text": "示例文字",
      "overrides": {
        "bold": true,
        "size": 32,
        "color": "#000000",
        "font_ascii": "Times New Roman",
        "font_east_asia": "宋体"
      }
    }
  ]
}
```

- `style`: 引用 styles 中的 `style_id`
- `paragraph_format`: 顶层字段，存储对继承样式的段落格式覆盖，只写有覆盖的字段
- `content`: Text 节点和行内节点的数组

**Heading 不单独设类型**，用 `style` 字段区分，例如 `"style": "Heading1"`。

**列表无独立节点类型**，列表项以普通 `Paragraph` 处理，通过 `style`（如 `"ListParagraph"`）区分。

#### 4.5.2 Text 节点（行内节点）

```json
{
  "type": "Text",
  "text": "示例文字",
  "overrides": {
    "bold": true,
    "italic": false,
    "size": 28,
    "color": "#FF0000",
    "underline": "single",
    "font_ascii": "Arial",
    "font_east_asia": "黑体"
  }
}
```

- Text 节点没有 `id` 字段
- `overrides` 只写有直接格式化的字段，结构与 styles 中 `character_format` 相同
- 空 `overrides`（`{}`）时可省略该字段
- 字号字段名为 `size`，单位半点

#### 4.5.3 Hyperlink 节点（行内）

```json
{
  "type": "Hyperlink",
  "url": "https://example.com",
  "content": [
    { "type": "Text", "text": "点击访问" }
  ]
}
```

#### 4.5.4 InlineImage 节点（行内）

```json
{
  "type": "InlineImage",
  "data": "<base64编码的图片数据>",
  "content_type": "image/png",
  "alt": "",
  "width": 1440,
  "height": 1080
}
```

- `data`: 图片二进制数据的 base64 编码字符串（内嵌在 AST 中，无需外部文件）
- `content_type`: MIME 类型，如 `"image/png"`、`"image/jpeg"`

#### 4.5.5 Table 节点

```json
{
  "id": "t0",
  "type": "Table",
  "style": "TableGrid",
  "rows": [
    {
      "cells": [
        {
          "id": "t0.r0c0",
          "col_span": 1,
          "row_span": 1,
          "content": [
            {
              "id": "t0.r0c0.p0",
              "type": "Paragraph",
              "content": [{"type": "Text", "text": "姓名"}]
            }
          ]
        },
        {
          "id": "t0.r0c1",
          "col_span": 2,
          "row_span": 1,
          "content": [
            {
              "id": "t0.r0c1.p0",
              "type": "Paragraph",
              "content": [{"type": "Text", "text": "横向合并示例"}]
            }
          ]
        }
      ]
    }
  ]
}
```

**合并单元格的处理**：

- 横向合并（gridSpan）：第一个单元格 `col_span` = 合并列数，被合并的后续单元格**在 cells 数组中不出现**
- 纵向合并（vMerge）：第一个单元格 `row_span` = 合并行数，后续行中对应位置的单元格**在 cells 数组中不出现**
- `col_span` 和 `row_span` 直接从底层 XML（`w:tcPr/w:gridSpan`、`w:vMerge`）读取

`content` 是 Block 节点数组（单元格内可以有多个段落）。

#### 4.5.6 分隔节点

```json
{ "id": "p10", "type": "PageBreak" }
```

### 4.6 passthrough 块

存储本系统不解析、但必须在 round-trip 中保留的内容。

```json
"passthrough": {
  "header_xml": "<w:hdr>...</w:hdr>",
  "footer_xml": "<w:ftr>...</w:ftr>",
  "numbering_xml": "<w:numbering>...</w:numbering>",
  "theme_xml": "<a:theme>...</a:theme>"
}
```

passthrough 内容原样写回，不做任何修改。

---

## 5. 解析层

解析层实现位于 `word_ast/parser/` 目录：

- `document_parser.py` — 顶层入口 `parse_docx()`，遍历文档 body，协调各子解析器
- `paragraph_parser.py` — 段落与 Text run 解析，处理字符格式、行内图片
- `table_parser.py` — 表格解析，直接读取 `w:tcPr/w:gridSpan` 和 `w:vMerge` 计算合并跨度
- `style_parser.py` — 样式库解析，`_normalize_style_type()` 将 `WD_STYLE_TYPE` 枚举转为字符串

**`parse_docx()` 签名：**

```python
def parse_docx(docx_path: str | Path, output_dir: str | Path | None = None) -> dict
```

`output_dir` 可选，若指定则将 AST JSON 和提取的资源保存到该目录。

---

## 6. 渲染层

渲染层实现位于 `word_ast/renderer/` 目录：

- `document_renderer.py` — 顶层入口 `render_ast()`
- `paragraph_renderer.py` — 段落与 Text run 渲染，优先使用 `_raw_pPr`/`_raw_rPr` XML（保真路径），回退到结构化字段
- `table_renderer.py` — 表格渲染，从 `col_span`/`row_span` 还原为 OOXML 的 vMerge/gridSpan
- `toc_renderer.py` — TOC 渲染
- `style_renderer.py` — 样式库渲染，按拓扑顺序创建样式（父样式先于子样式）

---

## 7. AI 层

### 7.1 `_raw_*` 字段（保真层）

解析时，段落格式 XML（`<w:pPr>`）和 run 格式 XML（`<w:rPr>`）原样以字符串存储在 AST 中，字段名以 `_raw_` 开头：

- `paragraph_format._raw_pPr` — 段落属性 XML 字符串
- `run.overrides._raw_rPr` — run 属性 XML 字符串

渲染时优先使用 `_raw_*` XML 直接写回，确保原始格式细节（如复杂的 spacing、shading 等）不丢失。

### 7.2 `to_ai_view(ast) -> dict`

实现位于 `word_ast/ai_view.py`。

递归删除 AST 中所有以 `_raw_` 开头的字段，返回深拷贝，生成干净的 AI 视图。

```python
from word_ast import to_ai_view

ai_view = to_ai_view(ast)
# ai_view 不含任何 _raw_* 字段，适合发给 AI 查看和修改
```

### 7.3 `merge_ai_edits(original_ast, ai_ast) -> dict`

实现位于 `word_ast/ai_merge.py`。

将 AI 修改的视图合并回含 `_raw_*` 的完整 AST：

- **block 匹配**：按 `id` 字段匹配，未在 ai_ast 中出现的 block 保持不变
- **run 匹配**：在每个 paragraph 的 `content` 列表中按位置匹配
- **AI 未修改的字段**：保留 original_ast 中的值（含 `_raw_*`）
- **AI 修改了的字段**：更新结构化字段，并同步写入 `_raw_*` 对应的 XML 元素
- **XML 解析失败时**：删除对应 `_raw_*`，让 Renderer 走结构化路径（降级）
- **目前仅支持 Paragraph 块合并**，Table 块暂不支持

```python
from word_ast import merge_ai_edits, render_ast

merged_ast = merge_ai_edits(original_ast, ai_modified_view)
render_ast(merged_ast, "output.docx")
```

### 7.4 AI 编辑 Pipeline 流程

```
┌─────────────┐
│  input.docx │
└──────┬──────┘
       │ parse_docx()
       ▼
┌─────────────────────────────┐
│  完整 AST（含 _raw_* XML）  │
└──────┬──────────────────────┘
       │ to_ai_view()
       ▼
┌─────────────────────────────┐
│  干净 AI 视图（无 _raw_*）  │  ──► 发给 AI
└──────────────────────────────┘
                                      │ AI 修改
                                      ▼
                              ┌───────────────┐
                              │  AI 修改后视图│
                              └───────┬───────┘
       ┌──────────────────────────────┘
       │ merge_ai_edits(完整 AST, AI 修改后视图)
       ▼
┌─────────────────────────────┐
│  合并后完整 AST             │
└──────┬──────────────────────┘
       │ render_ast()
       ▼
┌─────────────┐
│ output.docx │
└─────────────┘
```

---

## 8. 已实现功能

- DOCX → AST 解析
  - 文本段落与 run 级字符格式
  - 标题（通过 style 区分层级）
  - 列表（以带 style 的普通 Paragraph 保留）
  - 表格（含横向/纵向合并单元格）
  - 图片内联（base64 内嵌）
  - TOC（目录）
  - 样式库与样式继承
  - passthrough 机制保留不支持的内容
- AST → DOCX 渲染（round-trip）
- AI 编辑层（`to_ai_view` + `merge_ai_edits`）

尚未完整支持：复杂分页、页眉页脚、批注、修订记录等高级 Word 特性。

---

## 9. 已知难点与注意事项

### 9.1 python-docx 的遍历顺序问题

`doc.paragraphs` 和 `doc.tables` 是平铺的列表，无法直接反映文档中段落和表格的交替顺序。正确的遍历方式是直接遍历 `doc.element.body` 的子元素，根据 tag 判断类型。

### 9.2 Word 双字体体系

Word 对中文文档使用双字体：西文字体（`w:ascii`/`w:hAnsi`）和东亚字体（`w:eastAsia`）。python-docx 的 `font.name` 只对应西文字体，东亚字体必须通过 XML 直接读写。

### 9.3 合并单元格的 span 计算

通过直接读取底层 XML 计算 span：
- `col_span`: 读取 `w:tcPr/w:gridSpan[@w:val]`，默认 1
- `row_span`: 检查后续行同列位置的 `w:vMerge`（无 `restart` 属性则为续接行），累计计数

### 9.4 样式名本地化

Word 的内置样式有本地化名称（英文 Word 中叫 `"Heading 1"`，中文 Word 中叫 `"标题 1"`），但 `style_id`（XML 属性）通常保持为英文（`"Heading1"`）。AST 内部统一用 `style_id`。

### 9.5 样式 type 序列化

python-docx 的 `style.type` 是枚举值对象（`WD_STYLE_TYPE`），不能直接写入 JSON。解析时通过 `_normalize_style_type()` 统一转为字符串（`"paragraph"` / `"character"` / `"table"` / `"numbering"`）。
