# WORD_AST_SPEC 勘误记录（已修复项）

## 1) `parse_styles` 示例中的 `style.type` 序列化不稳定
- 问题：`python-docx` 的 `style.type` 是枚举对象，直接写 JSON 可能失败。
- 修复：在实现中新增 `_normalize_style_type`，将样式类型规范化为可序列化字符串（`paragraph/character/table/numbering`）。
- 代码位置：`word_ast/parser/style_parser.py`。

## 2) 表格合并单元格仅去重不足以恢复 `row_span`/`col_span`
- 问题：只按 `id(cell._tc)` 去重无法得到真实合并跨度。
- 修复：解析阶段读取底层 XML 的 `w:gridSpan` 与 `w:vMerge`，并计算 `col_span` / `row_span`；写回阶段按 `grid_col + span` 进行单元格合并渲染。
- 代码位置：`word_ast/parser/table_parser.py` 与 `word_ast/renderer/table_renderer.py`。
