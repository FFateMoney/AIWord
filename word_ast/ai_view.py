"""AI 视图层：将完整 AST 转换为适合 AI 查看和修改的精简 JSON。

AI view layer: converts the full AST (containing _raw_* XML fields) into a
clean JSON view suitable for AI consumption and modification.

The core round-trip pipeline is:
  docx → Parser → full AST (with _raw_*) → to_ai_view() → AI
  AI returns modified view → merge_ai_edits() → full AST → Renderer → docx

AI only sees and modifies semantic fields; _raw_* XML is managed internally
so that round-trip fidelity is preserved even when AI has not touched a field.
"""
import copy


def to_ai_view(ast: dict) -> dict:
    """返回适合给 AI 看的精简 AST，去掉所有 _raw_* 字段。

    Returns a deep copy of *ast* with all keys starting with ``_raw_``
    removed recursively.  AI only needs to see and modify semantic fields;
    it does not need to understand the underlying XML representation.
    """
    return _strip_raw(copy.deepcopy(ast))


def _strip_raw(obj):  # type: ignore[return]  # returns same type as input
    """递归删除 dict/list 中所有以 '_raw_' 开头的 key。

    Recursively removes all keys starting with ``_raw_`` from every dict
    nested inside *obj* (which may be a dict, list, or any scalar value).
    Modifies *obj* in place and returns it for convenience.
    """
    if isinstance(obj, dict):
        keys_to_del = [k for k in obj if k.startswith("_raw_")]
        for k in keys_to_del:
            del obj[k]
        for v in obj.values():
            _strip_raw(v)
    elif isinstance(obj, list):
        for item in obj:
            _strip_raw(item)
    return obj
