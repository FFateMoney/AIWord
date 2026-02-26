from typing import Any


def parse_docx(*args: Any, **kwargs: Any):
    from .document_parser import parse_docx as _parse_docx

    return _parse_docx(*args, **kwargs)


__all__ = ["parse_docx"]
