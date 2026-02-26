from typing import Any


def render_ast(*args: Any, **kwargs: Any):
    from .document_renderer import render_ast as _render_ast

    return _render_ast(*args, **kwargs)


__all__ = ["render_ast"]
