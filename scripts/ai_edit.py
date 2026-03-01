#!/usr/bin/env python3
"""完整 AI 编辑工作流示例 / Complete AI-assisted docx editing workflow.

Usage examples:

  # Parse and export the AI view (clean JSON without _raw_* fields):
  python scripts/ai_edit.py --input doc.docx --ai-view-output view.json

  # Apply an AI-modified JSON back to the original and render output:
  python scripts/ai_edit.py --input doc.docx --ai-edit-input modified.json --output out.docx

  # Round-trip without AI changes (identity check):
  python scripts/ai_edit.py --input doc.docx --output out.docx

Pipeline
--------
  1. 解析 / Parse:           docx → full AST (with _raw_*)
  2. AI 视图 / AI view:      full AST → to_ai_view() → clean JSON (no _raw_*)
  3. AI 修改 / AI edits:     send view to AI, receive modified JSON
  4. Merge:                  merge_ai_edits(full AST, modified JSON) → merged AST
  5. 渲染 / Render:          merged AST → docx
"""
import argparse
import json
import sys
from pathlib import Path

# Ensure local package imports work when running as a script
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from word_ast import parse_docx, render_ast
from word_ast.ai_view import to_ai_view
from word_ast.ai_merge import merge_ai_edits


def main():
    parser = argparse.ArgumentParser(
        description="AI-assisted docx editing workflow",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--input", required=True, help="Input .docx file path")
    parser.add_argument("--output", help="Output .docx file path (required for rendering)")
    parser.add_argument(
        "--ai-view-output",
        help="Save the AI view (clean JSON without _raw_* fields) to this path",
    )
    parser.add_argument(
        "--ai-edit-input",
        help="Load AI-modified JSON from this path and merge it back before rendering",
    )
    args = parser.parse_args()

    # Step 1: Parse
    ast = parse_docx(args.input)
    print(f"Parsed: {args.input}")

    # Step 2: Generate AI view
    ai_view = to_ai_view(ast)
    if args.ai_view_output:
        Path(args.ai_view_output).write_text(
            json.dumps(ai_view, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"AI view saved to: {args.ai_view_output}")

    if args.output is None:
        return

    # Step 3: Load AI-modified JSON (or pass through unmodified)
    if args.ai_edit_input:
        ai_modified = json.loads(
            Path(args.ai_edit_input).read_text(encoding="utf-8")
        )
        print(f"AI edits loaded from: {args.ai_edit_input}")
    else:
        ai_modified = ai_view  # identity: no AI changes

    # Step 4: Merge AI edits back into the full AST
    merged_ast = merge_ai_edits(ast, ai_modified)

    # Step 5: Render
    render_ast(merged_ast, args.output)
    print(f"Output written to: {args.output}")


if __name__ == "__main__":
    main()
