# WORD_AST_SPEC 勘误记录

## 1) `parse_styles` 示例中的 `style.type` 序列化不稳定
- 文档中示例倾向直接把 `style.type` 放进 AST（或仅描述为字符串）。
- 在 `python-docx` 中，`style.type` 是枚举值对象，直接写入 JSON 会报 `TypeError: Object of type WD_STYLE_TYPE is not JSON serializable`。
- 修正建议：统一转成字符串（例如 `paragraph` / `character` / `table` / `numbering`）。

## 2) 表格合并单元格“仅靠 `row.cells` 去重”不足以恢复 `row_span`/`col_span`
- 规格中给出用 `id(cell._tc)` 去重的建议，这能避免重复处理同一 XML 节点。
- 但该方法本身不能直接计算纵向合并(`vMerge`)与横向合并(`gridSpan`)跨度，若按去重结果直接写 `row_span=1,col_span=1` 会丢失合并信息。
- 修正建议：需读取单元格底层 XML（`w:tcPr/w:gridSpan`、`w:vMerge`）来计算真实跨度。
