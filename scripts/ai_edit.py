#!/usr/bin/env python3
"""AI-assisted Word document workflow.

Subcommands:

  export
      Export both the editable AI view and the full AST.

  export-content
      Export a content-only JSON view for AI reading tasks. Formatting fields
      are removed, but inline image data is retained.

  render
      Render an AI view back to ``.docx``. When a full AST schema is provided,
      AI edits are merged back into the fidelity-preserving AST first.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from word_ast import merge_ai_edits, parse_docx, render_ast, to_content_view
from word_ast.ai_view import to_ai_view


def cmd_export(args):
    input_path = Path(args.input)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    stem = input_path.stem
    ai_view_path = outdir / f"{stem}.ai_view.json"
    full_ast_path = outdir / f"{stem}.full_ast.json"

    full_ast = parse_docx(input_path)
    print(f"Parsed: {input_path}")

    full_ast_path.write_text(
        json.dumps(full_ast, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Full AST saved : {full_ast_path}")

    ai_view = to_ai_view(full_ast)
    ai_view_path.write_text(
        json.dumps(ai_view, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"AI view saved  : {ai_view_path}")


def cmd_export_content(args):
    input_path = Path(args.input)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    stem = input_path.stem
    content_view_path = outdir / f"{stem}.content_view.json"

    full_ast = parse_docx(input_path)
    print(f"Parsed: {input_path}")

    content_view = to_content_view(full_ast)
    content_view_path.write_text(
        json.dumps(content_view, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Content view saved: {content_view_path}")


def cmd_render(args):
    output_path = Path(args.output)

    ai_view = json.loads(Path(args.view).read_text(encoding="utf-8"))
    print(f"AI view loaded : {args.view}")

    if args.schema:
        full_ast = json.loads(Path(args.schema).read_text(encoding="utf-8"))
        print(f"Full AST loaded: {args.schema}")
        ast_to_render = merge_ai_edits(full_ast, ai_view)
        print("Merged AI edits into full AST.")
    else:
        ast_to_render = ai_view
        print("No schema provided; rendering AI view directly.")

    render_ast(ast_to_render, output_path)
    print(f"Output written : {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="AI-assisted Word document workflow",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_export = sub.add_parser(
        "export",
        help="Export AI view and full AST",
    )
    p_export.add_argument("-I", "--input", required=True, metavar="DOCX", help="Input .docx path")
    p_export.add_argument("-O", "--outdir", required=True, metavar="DIR", help="Output directory")

    p_export_content = sub.add_parser(
        "export-content",
        help="Export content-only JSON view",
    )
    p_export_content.add_argument(
        "-I", "--input", required=True, metavar="DOCX", help="Input .docx path"
    )
    p_export_content.add_argument(
        "-O", "--outdir", required=True, metavar="DIR", help="Output directory"
    )

    p_render = sub.add_parser(
        "render",
        help="Render AI view back to .docx",
    )
    p_render.add_argument("-V", "--view", required=True, metavar="JSON", help="AI view JSON path")
    p_render.add_argument(
        "-S",
        "--schema",
        default=None,
        metavar="JSON",
        help="Optional full AST JSON path for merge mode",
    )
    p_render.add_argument("-O", "--output", required=True, metavar="DOCX", help="Output .docx path")

    args = parser.parse_args()

    if args.cmd == "export":
        cmd_export(args)
    elif args.cmd == "export-content":
        cmd_export_content(args)
    elif args.cmd == "render":
        cmd_render(args)


if __name__ == "__main__":
    main()
