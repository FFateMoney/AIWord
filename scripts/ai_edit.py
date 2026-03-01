#!/usr/bin/env python3
"""AI-assisted Word document workflow.

两个子命令 / Two subcommands:

  export  —— docx → AI 视图 + 保真数据（两个 JSON 文件）
  render  —— AI 视图 [+ 保真数据] → docx

导出 / Export:
  python scripts/ai_edit.py export -I report.docx -O ./out/
  产出:
    ./out/report.ai_view.json    # 给 LLM 操作的干净视图
    ./out/report.full_ast.json   # 保真数据，本地留存

渲染（修改已有文档）/ Render (edit existing doc):
  python scripts/ai_edit.py render -V ./out/modified.ai_view.json \\
                                    -S ./out/report.full_ast.json \\
                                    -O output.docx

渲染（从零创建）/ Render (create from scratch):
  python scripts/ai_edit.py render -V new_doc.json -O output.docx
"""
import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from word_ast import parse_docx, render_ast
from word_ast.ai_view import to_ai_view
from word_ast.ai_merge import merge_ai_edits


def cmd_export(args):
    input_path = Path(args.input)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    stem = input_path.stem
    ai_view_path = outdir / f"{stem}.ai_view.json"
    full_ast_path = outdir / f"{stem}.full_ast.json"

    # Step 1: Parse → full AST (含 _raw_*)
    full_ast = parse_docx(input_path)
    print(f"Parsed: {input_path}")

    # Step 2: Save full AST（保真数据，用户本地留存）
    full_ast_path.write_text(
        json.dumps(full_ast, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Full AST saved : {full_ast_path}")

    # Step 3: Generate AI view（去掉 _raw_*，给 LLM）
    ai_view = to_ai_view(full_ast)
    ai_view_path.write_text(
        json.dumps(ai_view, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"AI view saved  : {ai_view_path}")


def cmd_render(args):
    output_path = Path(args.output)

    # 读取 AI 视图（必需）
    ai_view = json.loads(Path(args.view).read_text(encoding="utf-8"))
    print(f"AI view loaded : {args.view}")

    if args.schema:
        # 场景 A：修改已有文档 — merge AI 视图回保真 AST
        full_ast = json.loads(Path(args.schema).read_text(encoding="utf-8"))
        print(f"Full AST loaded: {args.schema}")
        ast_to_render = merge_ai_edits(full_ast, ai_view)
        print("Merged AI edits into full AST.")
    else:
        # 场景 B：从零创建 — ai_view 本身就是完整 AST（无 _raw_*）
        ast_to_render = ai_view
        print("No schema provided — rendering AI view directly (create mode).")

    render_ast(ast_to_render, output_path)
    print(f"Output written : {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="AI-assisted Word document workflow",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    # ── export ──────────────────────────────────────────────────────────────
    p_export = sub.add_parser(
        "export",
        help="docx → AI view + full AST（两个 JSON 文件）",
    )
    p_export.add_argument("-I", "--input", required=True, metavar="DOCX",
                          help="输入 .docx 文件路径")
    p_export.add_argument("-O", "--outdir", required=True, metavar="DIR",
                          help="输出目录（自动生成 <stem>.ai_view.json 和 <stem>.full_ast.json）")

    # ── render ──────────────────────────────────────────────────────────────
    p_render = sub.add_parser(
        "render",
        help="AI view [+ full AST] → docx",
    )
    p_render.add_argument("-V", "--view", required=True, metavar="JSON",
                          help="AI 修改后的 ai_view JSON 文件路径")
    p_render.add_argument("-S", "--schema", default=None, metavar="JSON",
                          help="保真数据 full_ast JSON（可选；不传则为从零创建模式）")
    p_render.add_argument("-O", "--output", required=True, metavar="DOCX",
                          help="输出 .docx 文件路径")

    args = parser.parse_args()

    if args.cmd == "export":
        cmd_export(args)
    elif args.cmd == "render":
        cmd_render(args)


if __name__ == "__main__":
    main()
