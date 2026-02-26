from typing import Any


def parse_docx(*args: Any, **kwargs: Any):
    from .parser.document_parser import parse_docx as _parse_docx

    return _parse_docx(*args, **kwargs)


def render_ast(*args: Any, **kwargs: Any):
    from .renderer.document_renderer import render_ast as _render_ast

    return _render_ast(*args, **kwargs)


__all__ = ["parse_docx", "render_ast"]
