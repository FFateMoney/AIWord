# Word ↔ AST Round-Trip 实现规格文档

## 1. 项目背景与目标

### 1.1 为什么要做这个

现有AI操作Word文档的方式是通过MCP工具或Skills间接调用python-docx/win32com等库，存在两个根本性问题：

1. **语义鸿沟**：AI看到的是OOXML的XML结构，或者python-docx的对象API，这两者都不是为AI理解设计的，字段命名晦涩、结构嵌套复杂、格式继承关系隐式。
2. **操作原子性差**：AI每次操作都要调用多个工具，没有事务保障，出错难以回滚。

**解决思路**：设计一个"沙漏型"中间层。

```
原始 .docx
    │
    ▼  [解析层]
AI友好的 AST（JSON格式）
    │
    ▼  [原语层]（后续阶段实现）
AI操作AST
    │
    ▼  [渲染层]
新的 .docx
```

AI只需要理解和操作AST，不需要了解OOXML细节。

### 1.2 本阶段目标

**本阶段只做一件事：验证round-trip的正确性。**

即：`原始.docx → AST → 重建.docx`，重建后的文档在普通用户可感知的维度上与原始文档一致。

本阶段**不涉及**：
- AI操作原语
- 前端界面
- 高级Word特性（修订记录、宏、复杂域代码、内容控件）

### 1.3 成功标准

1. 给定一批测试docx文件，转换后的文档在以下维度与原文件一致：
   - 文字内容完整无丢失
   - 段落结构（标题层级、列表嵌套）正确
   - 基础字符格式（粗体、斜体、颜色、字号、字体）正确
   - 基础段落格式（对齐、缩进、间距）正确
   - 表格结构（行列数、合并单元格）正确
   - 图片位置和尺寸正确
2. 不支持的特性通过passthrough机制保留，不导致文档损坏
3. AST是合法JSON，可被人类阅读和理解

---

## 2. 技术选型

| 组件 | 选型 | 理由 |
|------|------|------|
| 主解析库 | `python-docx` | 直接操作OOXML，信息最完整 |
| 辅助解析 | `lxml` | 处理python-docx未封装的底层XML |
| AST格式 | JSON | AI友好，人类可读 |
| 图片提取 | `python-docx`内置 | 直接从docx包中提取 |
| 测试框架 | `pytest` | 标准选择 |
| 语言 | Python 3.10+ | 类型注解支持好 |

**不使用Pandoc**：Pandoc在样式、格式信息上丢失过多，本项目直接用python-docx从OOXML层面解析，自己控制信息保留的完整性。

---

## 3. 仓库结构

```
word-ast/
├── README.md
├── requirements.txt
├── pyproject.toml
│
├── word_ast/                    # 核心库
│   ├── __init__.py
│   ├── schema.py                # AST数据结构定义（dataclass或TypedDict）
│   ├── parser/
│   │   ├── __init__.py
│   │   ├── document_parser.py   # 顶层解析入口
│   │   ├── paragraph_parser.py  # 段落/Run解析
│   │   ├── table_parser.py      # 表格解析
│   │   ├── list_parser.py       # 列表解析
│   │   ├── image_parser.py      # 图片解析
│   │   └── style_parser.py      # 样式库解析
│   ├── renderer/
│   │   ├── __init__.py
│   │   ├── document_renderer.py # 顶层渲染入口
│   │   ├── paragraph_renderer.py
│   │   ├── table_renderer.py
│   │   ├── list_renderer.py
│   │   ├── image_renderer.py
│   │   └── style_renderer.py
│   └── utils/
│       ├── __init__.py
│       ├── units.py             # 单位转换工具
│       └── xml_helpers.py       # lxml辅助函数
│
├── tests/
│   ├── fixtures/                # 测试用docx文件
│   │   ├── basic_text.docx
│   │   ├── headings.docx
│   │   ├── lists.docx
│   │   ├── table_simple.docx
│   │   ├── table_merged.docx
│   │   ├── images.docx
│   │   └── mixed.docx           # 综合测试文件
│   ├── test_parser.py
│   ├── test_renderer.py
│   └── test_roundtrip.py        # 最重要的测试
│
├── scripts/
│   ├── convert.py               # CLI工具：docx→ast 或 ast→docx
│   └── generate_fixtures.py     # 生成测试文件的脚本
│
└── docs/
    └── ast_schema.md            # AST格式文档（本文档的精简版）
```

---

## 4. AST Schema 完整定义

### 4.1 单位约定

**所有长度单位统一使用 twip（缇）**。

```
1 英寸 = 1440 twip
1 厘米 = 567 twip
1 磅(pt) = 20 twip
12pt字号 = 240 twip（但字号字段用"半点"，12pt = 24）
```

字号单独使用**半点（half-point）**单位，原因是OOXML原生如此存储，避免转换精度损失。

颜色统一使用`#RRGGBB`十六进制字符串，或`"auto"`表示自动颜色。

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
- `size`: `"A4"` | `"Letter"` | `"A3"` | `"custom"`，当为`custom`时需要额外字段`width`和`height`（twip）
- `orientation`: `"portrait"` | `"landscape"`
- `margin`: 四边页边距，单位twip
- `default_style`: 文档默认段落样式名，通常是`"Normal"`
- `language`: BCP 47语言标签

### 4.4 styles 块

样式库，Key为样式名（与Word内部样式名保持一致，中文文档可能是"正文"、"标题1"等）。

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
      "font_size": 24,
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
      "font_size": 32,
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
- `based_on`: 父样式名，`null`表示无继承
- `paragraph_format`和`character_format`只存储与父样式**不同的字段**（差量存储）
- `alignment`: `"left"` | `"right"` | `"center"` | `"justify"`
- `line_spacing.rule`: `"auto"` | `"exact"` | `"atLeast"`；`auto`时value为行距倍数×240（单倍=240，1.5倍=360，双倍=480）；`exact`/`atLeast`时value为twip
- `underline`: `"none"` | `"single"` | `"double"` | `"dotted"` | `"dashed"` | `"wavy"`
- `vertical_align`: `"baseline"` | `"superscript"` | `"subscript"`
- `highlight`: `null` 或 Word高亮颜色名 `"yellow"` | `"cyan"` | `"magenta"` | `"brightGreen"` | `"blue"` | `"red"` | `"darkBlue"` | `"darkCyan"` | `"darkMagenta"` | `"darkGreen"` | `"darkRed"` | `"darkYellow"` | `"darkGray"` | `"lightGray"` | `"black"` | `"white"`

### 4.5 body 块

`body`是Block节点的数组，顺序与文档内容顺序一致。

#### 节点通用字段

所有Block节点都有：

```json
{
  "id": "b001",
  "type": "Paragraph"
}
```

`id`规则：
- 段落：`b001`, `b002` ... 按顺序编号
- 表格：`t001`，表格行：`t001_r0`，单元格：`t001_r0_c0`
- 列表项：`l001_i0`，嵌套列表项：`l001_i0_c0`（c表示children）

#### 4.5.1 Paragraph 节点

```json
{
  "id": "b001",
  "type": "Paragraph",
  "style": "Normal",
  "overrides": {
    "paragraph_format": {
      "alignment": "justify",
      "indent_first_line": 480
    },
    "character_format": {}
  },
  "content": [ ]
}
```

- `style`: 引用styles中的样式名
- `overrides`: 对style的局部覆盖，结构与styles中的format字段相同，只写有覆盖的字段
- `content`: Run节点和行内节点的数组

**Heading不单独设类型**，用`style`字段区分，例如`"style": "Heading1"`。

#### 4.5.2 Run 节点（行内节点）

```json
{
  "type": "Run",
  "text": "示例文字",
  "overrides": {
    "bold": true,
    "italic": false,
    "font_size": 28,
    "color": "#FF0000",
    "underline": "single",
    "font_ascii": "Arial",
    "font_east_asia": "黑体"
  }
}
```

- Run没有`id`字段（粒度太细，不需要AI直接寻址）
- `overrides`只写有直接格式化的字段，结构与styles中`character_format`相同
- 空`overrides`（`{}`）时可以省略该字段

#### 4.5.3 Hyperlink 节点（行内）

```json
{
  "type": "Hyperlink",
  "url": "https://example.com",
  "content": [
    { "type": "Run", "text": "点击访问" }
  ]
}
```

#### 4.5.4 InlineImage 节点（行内）

```json
{
  "type": "InlineImage",
  "src": "media/image1.png",
  "alt": "",
  "width": 1440,
  "height": 1080
}
```

`src`是相对于AST文件的媒体资源路径。解析时需要将docx中的图片解包到`media/`目录。

#### 4.5.5 List 节点

```json
{
  "id": "l001",
  "type": "List",
  "list_type": "bullet",
  "ordered_format": null,
  "style": "ListParagraph",
  "items": [
    {
      "id": "l001_i0",
      "content": [
        { "type": "Run", "text": "第一项" }
      ],
      "children": []
    },
    {
      "id": "l001_i1",
      "content": [
        { "type": "Run", "text": "第二项" }
      ],
      "children": [
        {
          "id": "l001_i1_c0",
          "content": [
            { "type": "Run", "text": "子项" }
          ],
          "children": []
        }
      ]
    }
  ]
}
```

- `list_type`: `"bullet"` | `"ordered"`
- `ordered_format`: `null`（bullet时）或 `"decimal"` | `"lowerLetter"` | `"upperLetter"` | `"lowerRoman"` | `"upperRoman"` | `"chineseCounting"`
- 列表项的`content`是Run节点数组（行内内容）
- 嵌套列表用`children`数组递归表达

**解析难点**：Word中列表在OOXML层面是带有`numPr`属性的普通段落，需要通过`numId`和`ilvl`（缩进级别）识别同一列表和层级，然后在AST中重建为嵌套结构。

#### 4.5.6 Table 节点

```json
{
  "id": "t001",
  "type": "Table",
  "style": "TableGrid",
  "overrides": {
    "width": 9360,
    "alignment": "center",
    "border": {
      "top": { "style": "single", "size": 4, "color": "auto" },
      "bottom": { "style": "single", "size": 4, "color": "auto" },
      "left": { "style": "single", "size": 4, "color": "auto" },
      "right": { "style": "single", "size": 4, "color": "auto" },
      "insideH": { "style": "single", "size": 4, "color": "auto" },
      "insideV": { "style": "single", "size": 4, "color": "auto" }
    }
  },
  "col_widths": [2000, 3000, 4360],
  "rows": [
    {
      "id": "t001_r0",
      "is_header": true,
      "height": null,
      "cells": [
        {
          "id": "t001_r0_c0",
          "col_span": 1,
          "row_span": 1,
          "overrides": {
            "background_color": "#D9E1F2",
            "vertical_alignment": "center"
          },
          "content": [
            {
              "type": "Paragraph",
              "style": "Normal",
              "content": [{ "type": "Run", "text": "姓名" }]
            }
          ]
        }
      ]
    }
  ]
}
```

**合并单元格的关键处理**：

OOXML用`<w:vMerge>`和`<w:hMerge>`表示合并，逻辑晦涩。本AST改用`col_span`/`row_span`，但需要在解析时做转换：

- 横向合并（hMerge）：第一个单元格`col_span`=合并列数，被合并的后续单元格**在cells数组中不出现**
- 纵向合并（vMerge）：第一个单元格`row_span`=合并行数，后续行中对应位置的单元格**在cells数组中不出现**
- 渲染时需要反向推算哪些位置被跳过，还原为OOXML的vMerge/hMerge

`content`是Block节点数组（单元格内可以有多个段落）。

#### 4.5.7 Image 节点（块级）

```json
{
  "id": "b030",
  "type": "Image",
  "src": "media/image2.png",
  "alt": "流程图",
  "width": 5760,
  "height": 3240,
  "alignment": "center",
  "caption": {
    "text": "图1-1 系统架构图",
    "style": "Caption"
  }
}
```

- `alignment`: 图片所在段落的对齐方式
- `caption`: 可选，题注信息；`null`表示无题注

#### 4.5.8 分隔节点

```json
{ "id": "b040", "type": "PageBreak" }
{ "id": "b041", "type": "SectionBreak", "break_type": "nextPage" }
{ "id": "b042", "type": "HorizontalRule" }
```

`break_type`: `"nextPage"` | `"continuous"` | `"evenPage"` | `"oddPage"`

### 4.6 passthrough 块

存储本系统不解析、但必须在round-trip中保留的内容。

```json
"passthrough": {
  "header_xml": "<w:hdr>...</w:hdr>",
  "footer_xml": "<w:ftr>...</w:ftr>",
  "numbering_xml": "<w:numbering>...</w:numbering>",
  "theme_xml": "<a:theme>...</a:theme>",
  "unknown_parts": [
    {
      "position_after_block_id": "b020",
      "raw_xml": "<w:sdt>...</w:sdt>"
    }
  ]
}
```

passthrough内容原样写回，不做任何修改。

---

## 5. 解析层实现指南

### 5.1 入口逻辑（document_parser.py）

```python
from docx import Document
from docx.oxml.ns import qn
import json
import zipfile
import os

def parse_docx(docx_path: str, output_dir: str) -> dict:
    """
    主解析入口
    
    Args:
        docx_path: 输入的.docx文件路径
        output_dir: 输出目录，AST的JSON和media/目录都放这里
    
    Returns:
        AST字典
    """
    doc = Document(docx_path)
    
    # 1. 解包图片资源
    extract_media(docx_path, os.path.join(output_dir, "media"))
    
    # 2. 解析meta
    meta = parse_meta(doc)
    
    # 3. 解析样式库
    styles = parse_styles(doc)
    
    # 4. 解析body
    body = parse_body(doc)
    
    # 5. 提取passthrough
    passthrough = extract_passthrough(docx_path)
    
    return {
        "schema_version": "1.0",
        "document": {
            "meta": meta,
            "styles": styles,
            "body": body,
            "passthrough": passthrough
        }
    }

def extract_media(docx_path: str, media_dir: str):
    """从docx包中提取所有图片到media目录"""
    os.makedirs(media_dir, exist_ok=True)
    with zipfile.ZipFile(docx_path, 'r') as z:
        for name in z.namelist():
            if name.startswith('word/media/'):
                filename = os.path.basename(name)
                z.extract(name, media_dir)
                # 整理路径：把word/media/xxx.png移到media/xxx.png
```

### 5.2 Meta解析（document_parser.py）

```python
from docx.oxml.ns import qn
from docx.shared import Inches

def parse_meta(doc) -> dict:
    section = doc.sections[0]
    
    # 纸张尺寸识别
    width_twip = section.page_width.twips if section.page_width else 11906
    height_twip = section.page_height.twips if section.page_height else 16838
    size = detect_page_size(width_twip, height_twip)
    
    return {
        "page": {
            "size": size,
            "orientation": "landscape" if width_twip > height_twip else "portrait",
            "margin": {
                "top": int(section.top_margin.twips) if section.top_margin else 1440,
                "bottom": int(section.bottom_margin.twips) if section.bottom_margin else 1440,
                "left": int(section.left_margin.twips) if section.left_margin else 1800,
                "right": int(section.right_margin.twips) if section.right_margin else 1800,
            }
        },
        "default_style": "Normal",
        "language": detect_language(doc)
    }

def detect_page_size(width_twip: int, height_twip: int) -> str:
    """根据尺寸判断纸张规格，允许±50 twip误差"""
    sizes = {
        "A4": (11906, 16838),
        "Letter": (12240, 15840),
        "A3": (16838, 23811),
    }
    for name, (w, h) in sizes.items():
        if abs(width_twip - w) < 50 and abs(height_twip - h) < 50:
            return name
    return "custom"
```

### 5.3 样式解析（style_parser.py）

样式解析需要直接操作OOXML，因为python-docx的高级API对样式的封装不完整。

```python
from docx.oxml.ns import qn
from lxml import etree

def parse_styles(doc) -> dict:
    styles = {}
    
    for style in doc.styles:
        if style.type.name not in ('PARAGRAPH', 'CHARACTER', 'TABLE'):
            continue
        
        style_dict = {
            "style_id": style.style_id,
            "name": style.name,
            "type": style.type.name.lower(),
            "based_on": style.base_style.name if style.base_style else None,
        }
        
        # 解析paragraph_format
        if style.type.name == 'PARAGRAPH' and style.paragraph_format:
            style_dict["paragraph_format"] = parse_paragraph_format(style.paragraph_format)
        
        # 解析character_format
        if style.font:
            style_dict["character_format"] = parse_character_format(style.font)
        
        styles[style.name] = style_dict
    
    return styles

def parse_paragraph_format(pf) -> dict:
    """解析段落格式，None值字段跳过（表示继承自父样式）"""
    result = {}
    
    if pf.alignment is not None:
        result["alignment"] = pf.alignment.name.lower()
    
    for field in ['left_indent', 'right_indent', 'first_line_indent']:
        val = getattr(pf, field)
        if val is not None:
            key_map = {
                'left_indent': 'indent_left',
                'right_indent': 'indent_right',
                'first_line_indent': 'indent_first_line'
            }
            result[key_map[field]] = int(val.twips)
    
    if pf.space_before is not None:
        result["space_before"] = int(pf.space_before.twips)
    if pf.space_after is not None:
        result["space_after"] = int(pf.space_after.twips)
    
    if pf.line_spacing is not None:
        rule = pf.line_spacing_rule
        result["line_spacing"] = {
            "rule": rule.name.lower() if rule else "auto",
            "value": int(pf.line_spacing.twips) if hasattr(pf.line_spacing, 'twips') else int(pf.line_spacing * 240)
        }
    
    for bool_field in ['keep_with_next', 'keep_together', 'page_break_before']:
        val = getattr(pf, bool_field)
        if val is not None:
            key_map = {'keep_together': 'keep_lines_together'}
            ast_key = key_map.get(bool_field, bool_field)
            result[ast_key] = val
    
    return result

def parse_character_format(font) -> dict:
    result = {}
    
    if font.name is not None:
        result["font_ascii"] = font.name
    
    # 中文字体需要从XML直接读
    if font._element is not None:
        rFonts = font._element.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia'))
            if ea:
                result["font_east_asia"] = ea
    
    if font.size is not None:
        result["font_size"] = int(font.size.pt * 2)  # 转换为半点
    
    for bool_field in ['bold', 'italic', 'strike']:
        val = getattr(font, bool_field)
        if val is not None:
            result[bool_field] = val
    
    if font.underline is not None:
        if font.underline is True:
            result["underline"] = "single"
        elif font.underline is False:
            result["underline"] = "none"
        else:
            result["underline"] = str(font.underline)
    
    if font.color and font.color.type is not None:
        if str(font.color.type) == 'RGB (1)':
            result["color"] = f"#{font.color.rgb}"
        else:
            result["color"] = "auto"
    
    if font.highlight_color is not None:
        result["highlight"] = font.highlight_color.name.lower()
    
    return result
```

### 5.4 Body解析（document_parser.py）

Body解析的核心难点是**列表识别**：OOXML中列表只是带有numPr属性的普通段落，需要在遍历时识别并重建为嵌套List节点。

```python
def parse_body(doc) -> list:
    body = []
    block_counter = {"n": 0, "t": 0, "l": 0}
    
    # 遍历文档body中的顶层元素
    i = 0
    elements = list(doc.element.body)
    
    while i < len(elements):
        elem = elements[i]
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        
        if tag == 'p':
            para = doc.paragraphs[get_para_index(doc, elem)]
            
            if is_list_paragraph(para):
                # 收集连续的列表段落，整体处理
                list_paras = [para]
                while i + 1 < len(elements):
                    next_elem = elements[i + 1]
                    next_tag = next_elem.tag.split('}')[-1]
                    if next_tag == 'p':
                        next_para = doc.paragraphs[get_para_index(doc, next_elem)]
                        if is_list_paragraph(next_para) and same_list(para, next_para):
                            list_paras.append(next_para)
                            i += 1
                        else:
                            break
                    else:
                        break
                
                block_counter["l"] += 1
                list_id = f"l{block_counter['l']:03d}"
                body.append(parse_list_block(list_paras, list_id))
            
            elif is_image_paragraph(para):
                block_counter["n"] += 1
                body.append(parse_image_block(para, f"b{block_counter['n']:03d}"))
            
            elif is_page_break(para):
                block_counter["n"] += 1
                body.append({"id": f"b{block_counter['n']:03d}", "type": "PageBreak"})
            
            else:
                block_counter["n"] += 1
                body.append(parse_paragraph_block(para, f"b{block_counter['n']:03d}"))
        
        elif tag == 'tbl':
            block_counter["t"] += 1
            table_id = f"t{block_counter['t']:03d}"
            # 找到对应的Table对象
            table = find_table_by_element(doc, elem)
            body.append(parse_table_block(table, table_id))
        
        i += 1
    
    return body

def is_list_paragraph(para) -> bool:
    """判断段落是否为列表段落"""
    return para._element.find(qn('w:pPr/w:numPr')) is not None
```

### 5.5 列表解析（list_parser.py）

列表解析是最复杂的部分，需要理解Word的numId/ilvl体系。

```python
def parse_list_block(paras: list, list_id: str) -> dict:
    """
    将一组连续的列表段落转换为嵌套List节点
    
    Word列表的OOXML结构：
    <w:numPr>
      <w:ilvl w:val="0"/>  <!-- 缩进级别，0=顶层，1=第一层嵌套 -->
      <w:numId w:val="1"/> <!-- 列表实例ID -->
    </w:numPr>
    """
    # 获取列表格式信息（项目符号 vs 有序）
    first_para = paras[0]
    num_id = get_num_id(first_para)
    list_type, ordered_format = get_list_type(first_para)
    
    # 递归构建嵌套结构
    items, _ = build_list_items(paras, 0, 0, list_id)
    
    return {
        "id": list_id,
        "type": "List",
        "list_type": list_type,
        "ordered_format": ordered_format,
        "style": first_para.style.name if first_para.style else "ListParagraph",
        "items": items
    }

def build_list_items(paras: list, start_idx: int, level: int, list_id: str) -> tuple:
    """递归构建列表项，返回(items, 处理到的索引)"""
    items = []
    item_counter = [0]
    i = start_idx
    
    while i < len(paras):
        para = paras[i]
        para_level = get_ilvl(para)
        
        if para_level < level:
            # 回到上层
            break
        elif para_level == level:
            item_id = f"{list_id}_i{len(items)}"
            children = []
            
            # 往前看，收集子级项
            if i + 1 < len(paras) and get_ilvl(paras[i + 1]) > level:
                children, i = build_list_items(paras, i + 1, level + 1, item_id + "_c")
            
            items.append({
                "id": item_id,
                "content": parse_inline_content(para),
                "children": children
            })
        
        i += 1
    
    return items, i

def get_ilvl(para) -> int:
    """获取段落的列表缩进级别"""
    numPr = para._element.find(qn('w:pPr/w:numPr'))
    if numPr is None:
        return -1
    ilvl = numPr.find(qn('w:ilvl'))
    return int(ilvl.get(qn('w:val'))) if ilvl is not None else 0
```

### 5.6 表格解析（table_parser.py）

```python
def parse_table_block(table, table_id: str) -> dict:
    # 解析列宽
    col_widths = parse_col_widths(table)
    
    # 解析表格样式和属性
    tbl_style = table.style.name if table.style else "TableGrid"
    
    rows = []
    for r_idx, row in enumerate(table.rows):
        row_id = f"{table_id}_r{r_idx}"
        cells = parse_row_cells(row, row_id, table, r_idx)
        rows.append({
            "id": row_id,
            "is_header": r_idx == 0,  # 简单规则，也可以读表格属性
            "height": get_row_height(row),
            "cells": cells
        })
    
    return {
        "id": table_id,
        "type": "Table",
        "style": tbl_style,
        "overrides": parse_table_overrides(table),
        "col_widths": col_widths,
        "rows": rows
    }

def parse_row_cells(row, row_id: str, table, r_idx: int) -> list:
    """
    解析行中的单元格，处理合并单元格
    
    OOXML合并逻辑：
    - 横向合并(gridSpan)：<w:gridSpan w:val="2"/> 表示跨2列
    - 纵向合并(vMerge)：
        第一个单元格：<w:vMerge w:val="restart"/>
        后续单元格：<w:vMerge/>（无val属性）
    
    本AST策略：被合并的单元格直接从cells数组中移除，
    用col_span/row_span表达合并数量
    """
    cells = []
    
    for c_idx, cell in enumerate(row.cells):
        # 检查是否是被纵向合并的续接单元格（应跳过）
        vMerge = cell._tc.find(qn('w:tcPr/w:vMerge'))
        if vMerge is not None and vMerge.get(qn('w:val')) != 'restart':
            # 这是续接单元格，跳过
            continue
        
        # 检查是否是被横向合并的续接单元格
        # 注意：python-docx的row.cells会返回合并后的重复对象，需要去重
        if c_idx > 0 and cell._tc == row.cells[c_idx - 1]._tc:
            continue
        
        col_span = get_col_span(cell)
        row_span = get_row_span(table, r_idx, c_idx)
        
        cell_id = f"{row_id}_c{c_idx}"
        cells.append({
            "id": cell_id,
            "col_span": col_span,
            "row_span": row_span,
            "overrides": parse_cell_overrides(cell),
            "content": parse_cell_content(cell, cell_id)
        })
    
    return cells

def get_col_span(cell) -> int:
    gridSpan = cell._tc.find(qn('w:tcPr/w:gridSpan'))
    if gridSpan is not None:
        return int(gridSpan.get(qn('w:val')))
    return 1

def get_row_span(table, r_idx: int, c_idx: int) -> int:
    """计算从当前单元格开始的纵向合并数"""
    span = 1
    for r in range(r_idx + 1, len(table.rows)):
        try:
            cell = table.rows[r].cells[c_idx]
            vMerge = cell._tc.find(qn('w:tcPr/w:vMerge'))
            if vMerge is not None and vMerge.get(qn('w:val')) != 'restart':
                span += 1
            else:
                break
        except IndexError:
            break
    return span
```

### 5.7 行内内容解析（paragraph_parser.py）

```python
def parse_inline_content(para) -> list:
    """解析段落的行内内容，返回Run/Hyperlink/InlineImage节点数组"""
    content = []
    
    for child in para._element:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        
        if tag == 'r':  # Run
            run_content = parse_run_element(child)
            if run_content:
                content.append(run_content)
        
        elif tag == 'hyperlink':  # 超链接
            hyperlink = parse_hyperlink_element(child, para)
            if hyperlink:
                content.append(hyperlink)
        
        elif tag == 'bookmarkStart' or tag == 'bookmarkEnd':
            pass  # 书签暂时忽略（可后续加入passthrough）
    
    return content

def parse_run_element(r_elem) -> dict | None:
    """解析单个Run元素"""
    # 获取文本
    text_parts = []
    for t in r_elem.findall('.//' + qn('w:t')):
        text_parts.append(t.text or '')
    
    # 处理特殊Run（图片、分页符等）
    if r_elem.find('.//' + qn('w:drawing')) is not None:
        return parse_inline_drawing(r_elem)
    
    if r_elem.find(qn('w:br')) is not None:
        br = r_elem.find(qn('w:br'))
        br_type = br.get(qn('w:type'))
        if br_type == 'page':
            return {"type": "PageBreak"}
        # 其他break类型暂时忽略
        return None
    
    text = ''.join(text_parts)
    if not text:
        return None
    
    # 解析字符格式
    rPr = r_elem.find(qn('w:rPr'))
    overrides = parse_run_properties(rPr) if rPr is not None else {}
    
    result = {"type": "Run", "text": text}
    if overrides:
        result["overrides"] = overrides
    
    return result

def parse_run_properties(rPr) -> dict:
    """解析w:rPr元素，返回character_format覆盖"""
    overrides = {}
    
    # 粗体
    b = rPr.find(qn('w:b'))
    if b is not None:
        overrides['bold'] = b.get(qn('w:val'), 'true').lower() not in ('0', 'false')
    
    # 斜体
    i = rPr.find(qn('w:i'))
    if i is not None:
        overrides['italic'] = i.get(qn('w:val'), 'true').lower() not in ('0', 'false')
    
    # 字号（w:sz单位是半点，直接使用）
    sz = rPr.find(qn('w:sz'))
    if sz is not None:
        overrides['font_size'] = int(sz.get(qn('w:val')))
    
    # 颜色
    color = rPr.find(qn('w:color'))
    if color is not None:
        val = color.get(qn('w:val'))
        overrides['color'] = f"#{val}" if val and val != 'auto' else 'auto'
    
    # 字体（需要同时处理w:rFonts的多个属性）
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is not None:
        ascii_font = rFonts.get(qn('w:ascii')) or rFonts.get(qn('w:hAnsi'))
        ea_font = rFonts.get(qn('w:eastAsia'))
        if ascii_font:
            overrides['font_ascii'] = ascii_font
        if ea_font:
            overrides['font_east_asia'] = ea_font
    
    # 高亮
    highlight = rPr.find(qn('w:highlight'))
    if highlight is not None:
        overrides['highlight'] = highlight.get(qn('w:val'))
    
    # 删除线
    strike = rPr.find(qn('w:strike'))
    if strike is not None:
        overrides['strike'] = strike.get(qn('w:val'), 'true').lower() not in ('0', 'false')
    
    # 下划线
    u = rPr.find(qn('w:u'))
    if u is not None:
        val = u.get(qn('w:val'), 'single')
        overrides['underline'] = 'none' if val == 'none' else val
    
    return overrides
```

---

## 6. 渲染层实现指南

渲染层是解析层的逆过程，核心难点：
1. 样式的继承链要正确重建
2. 列表的嵌套结构要转回OOXML的numPr体系
3. 合并单元格要从col_span/row_span还原为vMerge/gridSpan

### 6.1 入口逻辑（document_renderer.py）

```python
from docx import Document
from docx.shared import Pt, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH

def render_docx(ast: dict, output_path: str, media_dir: str = None):
    """
    将AST渲染为docx文件
    
    Args:
        ast: AST字典
        output_path: 输出的.docx文件路径
        media_dir: 图片资源目录，默认为output_path同目录下的media/
    """
    doc = Document()
    
    # 清空默认内容
    for para in doc.paragraphs:
        para._element.getparent().remove(para._element)
    
    document_data = ast['document']
    
    # 1. 应用页面设置
    render_meta(doc, document_data['meta'])
    
    # 2. 应用样式库
    render_styles(doc, document_data['styles'])
    
    # 3. 渲染body
    context = {
        'doc': doc,
        'media_dir': media_dir or os.path.dirname(output_path),
        'numbering_engine': NumberingEngine(doc)  # 管理列表编号
    }
    render_body(document_data['body'], context)
    
    # 4. 写回passthrough
    apply_passthrough(doc, document_data.get('passthrough', {}))
    
    doc.save(output_path)
```

### 6.2 样式渲染（style_renderer.py）

**关键**：必须先按继承链顺序创建样式（父样式先于子样式），否则`based_on`引用会失败。

```python
def render_styles(doc, styles: dict):
    """渲染样式库到文档"""
    # 拓扑排序：保证父样式先创建
    ordered = topological_sort_styles(styles)
    
    for style_name in ordered:
        style_data = styles[style_name]
        apply_or_update_style(doc, style_name, style_data)

def apply_or_update_style(doc, style_name: str, style_data: dict):
    """创建或更新样式"""
    try:
        style = doc.styles[style_name]
    except KeyError:
        style_type_map = {
            'paragraph': WD_STYLE_TYPE.PARAGRAPH,
            'character': WD_STYLE_TYPE.CHARACTER,
            'table': WD_STYLE_TYPE.TABLE,
        }
        style = doc.styles.add_style(style_name, style_type_map[style_data['type']])
    
    if style_data.get('based_on'):
        try:
            style.base_style = doc.styles[style_data['based_on']]
        except KeyError:
            pass  # 父样式不存在时忽略
    
    if 'paragraph_format' in style_data:
        apply_paragraph_format(style.paragraph_format, style_data['paragraph_format'])
    
    if 'character_format' in style_data:
        apply_character_format(style.font, style_data['character_format'])
```

### 6.3 段落渲染（paragraph_renderer.py）

```python
ALIGNMENT_MAP = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
}

def render_paragraph(doc, block: dict) -> 'Paragraph':
    style_name = block.get('style', 'Normal')
    
    try:
        para = doc.add_paragraph(style=style_name)
    except KeyError:
        para = doc.add_paragraph()
    
    # 应用paragraph_format覆盖
    overrides = block.get('overrides', {})
    if 'paragraph_format' in overrides:
        apply_paragraph_format(para.paragraph_format, overrides['paragraph_format'])
    
    # 渲染行内内容
    for inline_node in block.get('content', []):
        render_inline_node(para, inline_node)
    
    return para

def render_inline_node(para, node: dict):
    node_type = node['type']
    
    if node_type == 'Run':
        run = para.add_run(node['text'])
        if 'overrides' in node:
            apply_character_format(run.font, node['overrides'])
    
    elif node_type == 'Hyperlink':
        # 添加超链接需要操作OOXML
        add_hyperlink(para, node['url'], node['content'])
    
    elif node_type == 'InlineImage':
        run = para.add_run()
        # 行内图片
        img_path = os.path.join(media_dir, node['src'])
        if os.path.exists(img_path):
            run.add_picture(img_path, width=Twips(node['width']))

def apply_character_format(font, fmt: dict):
    if 'bold' in fmt:
        font.bold = fmt['bold']
    if 'italic' in fmt:
        font.italic = fmt['italic']
    if 'font_size' in fmt:
        font.size = Pt(fmt['font_size'] / 2)  # 半点转pt
    if 'color' in fmt and fmt['color'] != 'auto':
        from docx.dml.color import ColorFormat
        from docx.shared import RGBColor
        rgb = fmt['color'].lstrip('#')
        font.color.rgb = RGBColor(int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16))
    if 'font_ascii' in fmt:
        font.name = fmt['font_ascii']
    if 'font_east_asia' in fmt:
        # 需要直接操作XML设置东亚字体
        set_east_asia_font(font, fmt['font_east_asia'])
    if 'underline' in fmt:
        font.underline = fmt['underline'] != 'none'
    if 'strike' in fmt:
        font.strike = fmt['strike']
```

### 6.4 表格渲染（table_renderer.py）

```python
def render_table(doc, block: dict) -> 'Table':
    rows_data = block['rows']
    cols = len(block.get('col_widths', []))
    
    # 计算实际行列数（考虑合并）
    num_rows = len(rows_data)
    # 列数从col_widths推断
    
    table = doc.add_table(rows=num_rows, cols=cols)
    
    try:
        table.style = block.get('style', 'TableGrid')
    except KeyError:
        table.style = 'TableGrid'
    
    # 设置列宽
    for i, width in enumerate(block.get('col_widths', [])):
        for row in table.rows:
            row.cells[i].width = Twips(width)
    
    # 渲染每一行
    # 需要维护一个"占用矩阵"来处理行列合并
    occupied = [[False] * cols for _ in range(num_rows)]
    
    for r_idx, row_data in enumerate(rows_data):
        col_cursor = 0
        
        for cell_data in row_data['cells']:
            # 跳过被占用的位置
            while col_cursor < cols and occupied[r_idx][col_cursor]:
                col_cursor += 1
            
            if col_cursor >= cols:
                break
            
            col_span = cell_data.get('col_span', 1)
            row_span = cell_data.get('row_span', 1)
            
            cell = table.cell(r_idx, col_cursor)
            
            # 处理横向合并
            if col_span > 1:
                end_cell = table.cell(r_idx, col_cursor + col_span - 1)
                cell.merge(end_cell)
            
            # 处理纵向合并
            if row_span > 1:
                end_cell = table.cell(r_idx + row_span - 1, col_cursor)
                cell.merge(end_cell)
            
            # 标记被占用的格子
            for dr in range(row_span):
                for dc in range(col_span):
                    if r_idx + dr < num_rows and col_cursor + dc < cols:
                        occupied[r_idx + dr][col_cursor + dc] = True
            
            # 渲染单元格内容
            # 清空默认空段落
            for p in cell.paragraphs:
                p._element.getparent().remove(p._element)
            
            for content_block in cell_data.get('content', []):
                render_block_in_cell(cell, content_block)
            
            # 应用单元格格式
            cell_overrides = cell_data.get('overrides', {})
            if 'background_color' in cell_overrides:
                set_cell_background(cell, cell_overrides['background_color'])
            if 'vertical_alignment' in cell_overrides:
                cell.vertical_alignment = VA_MAP[cell_overrides['vertical_alignment']]
            
            col_cursor += col_span
    
    return table
```

### 6.5 列表渲染（list_renderer.py）

Word的列表渲染需要创建numbering定义。

```python
class NumberingEngine:
    """管理文档中的列表编号定义"""
    
    def __init__(self, doc):
        self.doc = doc
        self._num_id_counter = 1
    
    def create_numbering(self, list_type: str, ordered_format: str = None) -> int:
        """创建新的numbering定义，返回numId"""
        # 操作numbering.xml，创建abstractNum和num定义
        # 这需要直接操作OOXML
        num_id = self._num_id_counter
        self._num_id_counter += 1
        # ... 具体实现见下方
        return num_id

def render_list(doc, block: dict, numbering_engine):
    """将List节点渲染为Word的numPr段落"""
    num_id = numbering_engine.create_numbering(
        block['list_type'],
        block.get('ordered_format')
    )
    
    def render_items(items, level=0):
        for item in items:
            para = doc.add_paragraph()
            para.style = block.get('style', 'ListParagraph')
            
            # 设置numPr
            set_list_numbering(para, num_id, level)
            
            # 渲染内容
            for inline_node in item.get('content', []):
                render_inline_node(para, inline_node)
            
            # 递归渲染子列表
            if item.get('children'):
                render_items(item['children'], level + 1)
    
    render_items(block['items'])

def set_list_numbering(para, num_id: int, ilvl: int):
    """在段落上设置numPr"""
    from docx.oxml import OxmlElement
    
    pPr = para._element.get_or_add_pPr()
    numPr = OxmlElement('w:numPr')
    
    ilvl_elem = OxmlElement('w:ilvl')
    ilvl_elem.set(qn('w:val'), str(ilvl))
    numPr.append(ilvl_elem)
    
    numId_elem = OxmlElement('w:numId')
    numId_elem.set(qn('w:val'), str(num_id))
    numPr.append(numId_elem)
    
    pPr.append(numPr)
```

---

## 7. 测试策略

### 7.1 测试文件准备（scripts/generate_fixtures.py）

编写脚本程序性地生成覆盖各场景的测试docx文件，而不是手工制作，便于精确控制预期内容。

需要生成的测试文件：

| 文件名 | 测试内容 |
|--------|----------|
| `basic_text.docx` | 纯文字段落，各种字符格式（粗/斜/下划线/颜色/字号） |
| `headings.docx` | H1-H6标题，验证样式继承 |
| `lists_bullet.docx` | 无序列表，2层嵌套 |
| `lists_ordered.docx` | 有序列表，decimal/roman格式 |
| `table_simple.docx` | 3×3简单表格，无合并 |
| `table_merged.docx` | 含横向+纵向合并的表格 |
| `images.docx` | 块级图片，不同对齐方式 |
| `mixed.docx` | 以上所有内容混合 |
| `chinese.docx` | 中文内容，中文字体，首行缩进 |

### 7.2 Round-trip测试（tests/test_roundtrip.py）

```python
import pytest
import os
import json
from word_ast.parser.document_parser import parse_docx
from word_ast.renderer.document_renderer import render_docx

FIXTURES_DIR = os.path.join(os.path.dirname(__file__), 'fixtures')

@pytest.mark.parametrize("fixture_name", [
    "basic_text",
    "headings",
    "lists_bullet",
    "lists_ordered",
    "table_simple",
    "table_merged",
    "images",
    "mixed",
    "chinese",
])
def test_roundtrip(fixture_name, tmp_path):
    """测试round-trip：docx → AST → docx，验证关键属性一致性"""
    
    input_path = os.path.join(FIXTURES_DIR, f"{fixture_name}.docx")
    ast_path = tmp_path / f"{fixture_name}.ast.json"
    output_path = tmp_path / f"{fixture_name}_output.docx"
    media_dir = tmp_path / "media"
    
    # Step 1: 解析
    ast = parse_docx(str(input_path), str(tmp_path))
    with open(ast_path, 'w', encoding='utf-8') as f:
        json.dump(ast, f, ensure_ascii=False, indent=2)
    
    # Step 2: 验证AST结构合法性
    validate_ast_schema(ast)
    
    # Step 3: 渲染
    render_docx(ast, str(output_path), str(media_dir))
    
    # Step 4: 重新解析输出文件，比较关键属性
    ast2 = parse_docx(str(output_path), str(tmp_path / "output_media"))
    
    compare_ast(ast, ast2, fixture_name)

def compare_ast(ast1: dict, ast2: dict, fixture_name: str):
    """比较两个AST的关键属性"""
    body1 = ast1['document']['body']
    body2 = ast2['document']['body']
    
    assert len(body1) == len(body2), \
        f"{fixture_name}: body block count mismatch: {len(body1)} vs {len(body2)}"
    
    for i, (b1, b2) in enumerate(zip(body1, body2)):
        assert b1['type'] == b2['type'], \
            f"{fixture_name}: block {i} type mismatch: {b1['type']} vs {b2['type']}"
        
        if b1['type'] == 'Paragraph':
            compare_paragraph(b1, b2, f"{fixture_name}.b{i}")
        elif b1['type'] == 'Table':
            compare_table(b1, b2, f"{fixture_name}.t{i}")
        elif b1['type'] == 'List':
            compare_list(b1, b2, f"{fixture_name}.l{i}")

def compare_paragraph(p1: dict, p2: dict, ctx: str):
    # 文字内容必须完全一致
    text1 = extract_text(p1)
    text2 = extract_text(p2)
    assert text1 == text2, f"{ctx}: text mismatch:\n  expected: {text1!r}\n  got: {text2!r}"
    
    # 样式名一致
    assert p1.get('style') == p2.get('style'), \
        f"{ctx}: style mismatch: {p1.get('style')} vs {p2.get('style')}"

def validate_ast_schema(ast: dict):
    """基础schema验证"""
    assert 'schema_version' in ast
    assert 'document' in ast
    doc = ast['document']
    assert 'meta' in doc
    assert 'styles' in doc
    assert 'body' in doc
    assert isinstance(doc['body'], list)
    
    for block in doc['body']:
        assert 'id' in block, f"Block missing id: {block}"
        assert 'type' in block, f"Block missing type: {block}"
        assert block['type'] in ('Paragraph', 'Table', 'List', 'Image', 'PageBreak', 'SectionBreak', 'HorizontalRule'), \
            f"Unknown block type: {block['type']}"
```

### 7.3 解析单元测试（tests/test_parser.py）

```python
def test_parse_character_format_bold():
    """验证粗体解析正确"""
    doc = create_test_doc_with_bold_run()
    ast = parse_docx_from_doc(doc)
    
    runs = ast['document']['body'][0]['content']
    bold_run = next(r for r in runs if r.get('overrides', {}).get('bold'))
    assert bold_run['overrides']['bold'] == True

def test_parse_table_with_merged_cells():
    """验证合并单元格解析为正确的col_span/row_span"""
    # 构造一个2×2表格，左上角横向合并2列
    ...
    ast = parse_docx_from_doc(doc)
    table = ast['document']['body'][0]
    
    assert table['type'] == 'Table'
    first_row_cells = table['rows'][0]['cells']
    assert len(first_row_cells) == 1  # 合并后只有1个单元格
    assert first_row_cells[0]['col_span'] == 2

def test_parse_nested_list():
    """验证两层嵌套列表解析为正确的children结构"""
    ...
    list_block = ast['document']['body'][0]
    assert list_block['type'] == 'List'
    assert len(list_block['items'][1]['children']) > 0  # 第二项有子项
```

---

## 8. 开发顺序建议

按以下顺序实现，每步都能独立验证：

**第一步：基础框架 + Meta解析/渲染**
- 搭建项目结构，安装依赖
- 实现`parse_meta`和`render_meta`
- 验证：解析basic_text.docx，AST中meta字段正确；渲染后页边距一致

**第二步：样式库解析/渲染**
- 实现`parse_styles`和`render_styles`
- 验证：所有内置样式都能正确提取和重建

**第三步：段落 + Run 解析/渲染**
- 实现`parse_paragraph_block`、`parse_inline_content`、`render_paragraph`
- 验证：basic_text.docx完整round-trip，文字内容、粗斜体、颜色、字号一致

**第四步：标题**
- 只需确保Heading样式被正确映射，复用段落的逻辑
- 验证：headings.docx round-trip，标题层级正确

**第五步：表格（先无合并，再合并）**
- 实现`parse_table_block`和`render_table`
- 先支持简单表格，再加合并单元格
- 这是整个项目最容易出bug的地方，要认真测试

**第六步：列表**
- 实现`parse_list_block`和`render_list`
- 先支持单层，再加嵌套
- NumberingEngine需要仔细处理

**第七步：图片**
- 实现图片提取和渲染
- 验证图片尺寸、位置不变

**第八步：Passthrough**
- 实现passthrough的提取和写回
- 验证不支持的内容不丢失

**第九步：综合测试**
- 运行mixed.docx的round-trip测试
- 修复发现的问题

---

## 9. 已知难点与注意事项

### 9.1 python-docx的遍历顺序问题

`doc.paragraphs`和`doc.tables`是平铺的列表，无法直接反映文档中段落和表格的交替顺序。正确的遍历方式是直接遍历`doc.element.body`的子元素，根据tag判断类型。

```python
for child in doc.element.body:
    tag = child.tag.split('}')[-1]
    if tag == 'p':
        # 段落
    elif tag == 'tbl':
        # 表格
```

### 9.2 Word双字体体系

Word对中文文档使用双字体：西文字体（`w:ascii`/`w:hAnsi`）和东亚字体（`w:eastAsia`）。python-docx的`font.name`只对应西文字体，东亚字体必须通过XML直接读写。

```python
# 读
rFonts = rPr.find(qn('w:rFonts'))
ea_font = rFonts.get(qn('w:eastAsia'))

# 写
rFonts = OxmlElement('w:rFonts')
rFonts.set(qn('w:eastAsia'), '宋体')
rFonts.set(qn('w:eastAsiaTheme'), '')  # 清除主题字体引用，防止被覆盖
```

### 9.3 合并单元格的去重

python-docx在`row.cells`中会对合并的单元格返回同一个对象的多次引用，迭代时必须去重：

```python
seen = set()
for cell in row.cells:
    if id(cell._tc) in seen:
        continue
    seen.add(id(cell._tc))
    # 处理cell
```

### 9.4 列表的numId复用

同一文档中可能有多个列表共用同一个`abstractNumId`但有不同的`numId`，也可能两个外观不同的列表有不同的`abstractNumId`。解析时要通过`numId`区分不同列表实例，不能只看`abstractNumId`。

### 9.5 样式名本地化

Word的内置样式有本地化名称。英文Word中叫`"Heading 1"`，中文Word中叫`"标题 1"`，但`style_id`（XML属性）通常保持为英文（`"Heading1"`）。解析和渲染时要注意区分`style.name`（本地化）和`style.style_id`（英文ID），建议内部统一用`style_id`。

---

## 10. requirements.txt

```
python-docx>=1.1.0
lxml>=5.0.0
pytest>=8.0.0
pytest-cov>=4.0.0
```

---

## 11. 文件输出格式

解析后输出一个目录，包含：

```
output/
├── document.ast.json    # AST文件
└── media/
    ├── image1.png
    ├── image2.jpeg
    └── ...
```

CLI工具接口：

```bash
# docx → AST
python scripts/convert.py parse input.docx --output-dir ./output/

# AST → docx
python scripts/convert.py render ./output/document.ast.json --output output.docx
```
