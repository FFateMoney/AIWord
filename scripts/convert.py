#!/usr/bin/env python3
import argparse
import sys
from pathlib import Path


# Ensure local package imports work when running as a script:
#   python scripts/convert.py ...
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from word_ast import parse_docx, render_ast


def main():
    parser = argparse.ArgumentParser(description="docx <-> AST converter")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_parse = sub.add_parser("parse")
    p_parse.add_argument("input")
    p_parse.add_argument("--output-dir", required=True)

    p_render = sub.add_parser("render")
    p_render.add_argument("input")
    p_render.add_argument("--output", required=True)

    args = parser.parse_args()

    if args.cmd == "parse":
        parse_docx(args.input, args.output_dir)
    elif args.cmd == "render":
        render_ast(args.input, args.output)


if __name__ == "__main__":
    main()
