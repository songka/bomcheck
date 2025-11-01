from __future__ import annotations

from functools import lru_cache

try:
    from opencc import OpenCC
except ImportError:  # pragma: no cover
    OpenCC = None  # type: ignore


@lru_cache(maxsize=None)
def _get_converter() -> OpenCC | None:
    if OpenCC is None:
        return None
    try:
        return OpenCC("t2s")
    except Exception:  # pragma: no cover
        return None


def normalize_text(value: str) -> str:
    value = value.strip().lower()
    converter = _get_converter()
    if converter:
        try:
            value = converter.convert(value)
        except Exception:  # pragma: no cover
            pass
    return value
